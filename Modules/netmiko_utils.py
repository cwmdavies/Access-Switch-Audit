import logging
from netmiko import ConnectHandler
from netmiko.base_connection import BaseConnection
from typing import Any, Optional
from .jump_manager import JumpManager

log = logging.getLogger(__name__)

def connect_to_device(
    ip: str,
    username: str,
    password: str,
    device_type: str = "cisco_ios",
    jump: Optional[JumpManager] = None,
    port: int = 22,
    allow_agent: bool = False,
    look_for_keys: bool = False,
    **extras: Any,
) -> BaseConnection:
    """
    Lightweight wrapper around Netmiko.ConnectHandler.
    Accepts optional jump (JumpManager) which must provide a direct-tcpip channel via .open_channel().
    """
    kwargs: dict[str, Any] = {
        "device_type": device_type,
        "host": ip,
        "username": username,
        "password": password,
        "port": port,
        "allow_agent": allow_agent,
        "look_for_keys": look_for_keys,
        "auth_timeout": 20,
        "banner_timeout": 30,
        "conn_timeout": 25,
        "fast_cli": False,
    }
    kwargs.update(extras)
    if jump:
        try:
            sock = jump.open_channel(ip, port)
            kwargs["sock"] = sock
            log.debug("Opened jump channel to %s:%s", ip, port)
        except Exception:
            log.exception("Failed to open jump channel to %s:%s", ip, port)
            raise

    # Some netmiko/BaseConnection variants don't accept paramiko-specific
    # flags (look_for_keys / allow_agent) in their __init__; remove them
    # so they aren't forwarded to BaseConnection.__init__.
    for _k in ("look_for_keys", "allow_agent"):
        if _k in kwargs:
            log.debug("Removing unsupported kwarg %s before ConnectHandler()", _k)
            kwargs.pop(_k)

    log.debug("Connecting to device %s (%s)", ip, device_type)
    try:
        return ConnectHandler(**kwargs)
    except Exception:
        log.exception("ConnectHandler failed for %s", ip)
        raise