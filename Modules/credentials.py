import getpass
import os
import logging
import sys
import win32cred
from .config import CRED_TARGET

target = CRED_TARGET
log = logging.getLogger(__name__)

def get_secret_with_fallback() -> tuple[str, str]:
    """
    Try environment variables, then Windows Credential Manager, then interactive prompt.
    Environment variables checked (in order):
      - SWITCH_USER / SWITCH_PASS
      - CREDENTIAL_USER / CREDENTIAL_PASS
    For win32 cred the target name can be set via CREDENTIAL_TARGET (default 'MyApp/ADM').
    """
    # 1) Environment overrides
    for user_env, pass_env in (("SWITCH_USER", "SWITCH_PASS"), ("CREDENTIAL_USER", "CREDENTIAL_PASS")):
        user = os.environ.get(user_env)
        pwd = os.environ.get(pass_env)
        if user and pwd:
            return user.strip(), pwd

    # 2) Windows Credential Manager (optional)
    target = os.environ.get("CREDENTIAL_TARGET", "MyApp/ADM")
    try:
        import win32cred  # type: ignore
        cred = win32cred.CredRead(target, win32cred.CRED_TYPE_GENERIC)  # type: ignore
        user = cred.get('UserName')
        blob = cred.get('CredentialBlob')
        pwd = None
        if blob:
            # CredentialBlob is typically UTF-16LE for Windows generic creds
            try:
                pwd = blob.decode('utf-16le')
            except Exception:
                try:
                    pwd = blob.decode('utf-8', errors='ignore')
                except Exception:
                    pwd = None
        if user and pwd:
            log.critical("Found stored Primary user: %s (target: %s)", user, target)
            override = input("Press Enter to accept, or type a different username: ").strip()
            if override:
                primary_user = override
                primary_pass = getpass.getpass("Enter switch/jump password (Primary): ")
                if _prompt_yes_no(f"Save these Primary creds to Credential Manager as '{target}'?", default_no=True):
                    _write_win_cred(target, primary_user, primary_pass)
                else:
                    primary_user, primary_pass = user, pwd
                return primary_user, primary_pass    
            return user, pwd
        
    except Exception:
        log.debug("Win32 credential read failed or not available for target %s", target, exc_info=True)

    user = input("Enter switch/jump username: ").strip()
    pwd = getpass.getpass("Enter switch/jump password: ")
    if _prompt_yes_no(f"Save these Primary creds to Credential Manager as '{target}'?", default_no=True):
                    _write_win_cred(target, user, pwd)
    if not user or not pwd:
        raise RuntimeError("Credentials not found in Windows Credential Manager, env vars, or provided interactively.")
    return user, pwd

def _write_win_cred(target: str, username: str, password: str, persist: int = 2) -> bool:
        """
        Write or update a generic credential in Windows Credential Manager.

        Args:
            target: Credential target name (e.g., 'MyApp/ADM').
            username: Username to store.
            password: Password to store.
            persist: Persistence (2 = local machine).

        Returns:
            True if the write succeeded, False otherwise.
        """
        try:
            if not sys.platform.startswith("win"):
                log.warning("Not a Windows platform; cannot store credentials in Credential Manager.")
                return False

            # Prefer bytes; fallback to str if the installed pywin32 expects unicode.
            blob_bytes = password.encode("utf-16le")
            credential = {
                "Type": win32cred.CRED_TYPE_GENERIC,
                "TargetName": target,
                "UserName": username,
                "CredentialBlob": blob_bytes,
                "Comment": "Created by CDP Network Audit tool",
                "Persist": persist,
            }
            try:
                win32cred.CredWrite(credential, 0)
            except TypeError as te:
                log.debug("CredWrite rejected bytes for CredentialBlob (%s). Retrying with unicode string.", te)
                credential["CredentialBlob"] = password
                win32cred.CredWrite(credential, 0)
            log.info("Stored/updated credentials in Windows Credential Manager: %s", target)
            return True
        except Exception:
            log.exception("Failed to write credentials for '%s'", target)
            return False

def get_enable_secret() -> str | None:
    use_enable = os.environ.get("USE_ENABLE", "false").lower() in {"1", "true", "yes", "y"}
    if not use_enable:
        return None
    return os.environ.get("ENABLE_SECRET")

def _prompt_yes_no(msg: str, default_no: bool = True) -> bool:
    """Simple interactive [y/N] or [Y/n] prompt."""
    suffix = " [y/N] " if default_no else " [Y/n] "
    ans = input(msg + suffix).strip().lower()
    if ans == "":
        return not default_no
    return ans in ("y", "yes")