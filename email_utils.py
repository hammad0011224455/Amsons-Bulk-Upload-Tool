# email_utils.py
import os
import ssl
import io
import smtplib
from typing import Tuple, Optional
from email.message import EmailMessage
from contextlib import redirect_stdout
from pathlib import Path

DEFAULTS = {
    "enabled": True,
    "admin_to": "arif@amsons.co.uk",   # change in config.json if needed
    "smtp_host": "smtp.office365.com",
    "smtp_port": 587,                    # STARTTLS
    "smtp_user": "",                     # e.g. no-reply@yourdomain.com
    "smtp_pass": "",                     # leave blank if you prefer env var
    "use_env_password": False,           # read password from AMS_SMTP_PASS if True
    "from_name": "Amsons PM Bot",
    "from_email": "",                    # if blank, falls back to smtp_user
    "use_ssl": False                     # Office365 = False (use STARTTLS)
}

def _merge_email_cfg(cfg: Optional[dict]) -> dict:
    """
    Merge user config with defaults, preserving the DEFAULTS booleans.
    """
    email_cfg = dict(DEFAULTS)
    try:
        email_cfg.update((cfg or {}).get("email", {}))
    except Exception:
        pass

    # normalize types
    try:
        email_cfg["smtp_port"] = int(email_cfg.get("smtp_port", DEFAULTS["smtp_port"]))
    except Exception:
        email_cfg["smtp_port"] = DEFAULTS["smtp_port"]

    email_cfg["use_ssl"] = bool(email_cfg.get("use_ssl", DEFAULTS["use_ssl"]))
    email_cfg["use_env_password"] = bool(email_cfg.get("use_env_password", DEFAULTS["use_env_password"]))

    # If someone set smtp_pass to False in JSON, treat as empty string (no password)
    if not isinstance(email_cfg.get("smtp_pass", ""), str):
        email_cfg["smtp_pass"] = ""

    return email_cfg

def _send_via_outlook(admin_to: str, subject: str, body: str) -> Tuple[bool, str]:
    """
    Passwordless fallback: use local Outlook profile via COM.
    Requires: pip install pywin32 and Outlook desktop configured on this machine.
    """
    try:
        import win32com.client  # type: ignore
    except Exception as e:
        return False, f"Outlook/pywin32 not available: {e}"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = admin_to
        mail.Subject = subject
        mail.Body = body
        mail.Send()  # uses default Outlook profile; no password in code
        return True, "Sent via Outlook profile."
    except Exception as e:
        return False, f"Outlook send failed: {e}"

def send_admin_email(cfg: dict, subject: str, body: str, *, log_fn=None) -> Tuple[bool, str]:
    """
    Try SMTP first if fully configured; otherwise fall back to Outlook (no password).
    Returns (ok, transcript_or_error).
    """
    try:
        e = _merge_email_cfg(cfg)
        if not e["enabled"]:
            return False, "Email disabled in config."

        to_addr = (e.get("admin_to") or "").strip()
        if not to_addr:
            return False, "Missing admin_to in config."

        # Determine if SMTP is usable
        user = (e.get("smtp_user") or "").strip()
        pw = os.getenv("AMS_SMTP_PASS", "") if e["use_env_password"] else (e.get("smtp_pass") or "")
        from_addr = (e.get("from_email") or user).strip()
        host, port = e["smtp_host"], e["smtp_port"]

        # If SMTP not configured (no user or password), fall back to Outlook
        if (not user) or (not pw) or (not from_addr):
            ok, msg = _send_via_outlook(to_addr, subject, body)
            return (ok, msg if ok else f"No SMTP creds and Outlook failed: {msg}")

        # --- SMTP path ---
        msg = EmailMessage()
        from_name = e.get("from_name") or "Amsons PM Bot"
        msg["Subject"] = subject
        msg["From"] = f"{from_name} <{from_addr}>"
        msg["To"] = to_addr
        msg.set_content(body)

        transcript = io.StringIO()
        ok = False
        with redirect_stdout(transcript):
            if e["use_ssl"] and port == 465:
                # SSL (rare for O365)
                with smtplib.SMTP_SSL(host, port, context=ssl.create_default_context(), timeout=25) as s:
                    s.set_debuglevel(1)
                    s.login(user, pw)
                    s.send_message(msg)
                    ok = True
            else:
                # STARTTLS (Office365 recommended)
                with smtplib.SMTP(host, port, timeout=25) as s:
                    s.set_debuglevel(1)
                    s.ehlo()
                    s.starttls(context=ssl.create_default_context())
                    s.ehlo()
                    s.login(user, pw)
                    s.send_message(msg)
                    ok = True

        out = transcript.getvalue()
        if log_fn:
            log_fn("\n[SMTP transcript]\n" + out + "\n")
        return ok, out if ok else "Unknown failure."

    except Exception as ex:
        err = f"Email send failed: {ex}"
        if log_fn:
            log_fn(err + "\n")
        return False, err
