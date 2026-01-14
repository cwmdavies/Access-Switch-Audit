import logging
from typing import Optional, Tuple
import paramiko

log = logging.getLogger(__name__)

class JumpManager:
    """
    Maintains a single SSH connection to the jump server and opens per-device
    'direct-tcpip' channels for Netmiko via the 'sock' parameter.
    Use as context manager to ensure the jump connection is closed.
    """
    def __init__(
        self,
        jump_host: str,
        username: str,
        password: str,
        port: int = 22,
        look_for_keys: bool = False,
        allow_agent: bool = False,
        host_key_policy: Optional[paramiko.MissingHostKeyPolicy] = None,
    ):
        self.jump_host = jump_host
        self.username = username
        self.password = password
        self.port = port
        self.look_for_keys = look_for_keys
        self.allow_agent = allow_agent
        self.host_key_policy = host_key_policy or paramiko.AutoAddPolicy()
        self.client: Optional[paramiko.SSHClient] = None

    def connect(self) -> None:
        """Establish the persistent SSH connection to the jump host."""
        if self.client:
            return
        cli = paramiko.SSHClient()
        cli.set_missing_host_key_policy(self.host_key_policy)
        try:
            cli.connect(
                hostname=self.jump_host,
                port=self.port,
                username=self.username,
                password=self.password,
                look_for_keys=self.look_for_keys,
                allow_agent=self.allow_agent,
                timeout=20,
            )
            self.client = cli
            log.debug("Connected to jump host %s:%s", self.jump_host, self.port)
        except Exception:
            log.exception("Failed to connect to jump host %s:%s", self.jump_host, self.port)
            raise

    def open_channel(self, target_host: str, target_port: int = 22) -> paramiko.Channel:
        """
        Open a direct-tcpip channel from jump -> target_host:target_port.
        Returns a paramiko.Channel suitable for Netmiko's 'sock' kwarg.
        """
        if not self.client:
            self.connect()
        transport = self.client.get_transport()
        if transport is None or not transport.is_active():
            # try reconnect
            log.debug("Transport inactive, reconnecting to jump host")
            self.connect()
            transport = self.client.get_transport()
        try:
            # remote_addr is (target_host, target_port), local_addr is a tuple for the origin
            chan = transport.open_channel("direct-tcpip", (target_host, target_port), ("127.0.0.1", 0))
            log.debug("Opened channel to %s:%s via jump host", target_host, target_port)
            return chan
        except Exception:
            log.exception("Failed to open channel to %s:%s via jump host", target_host, target_port)
            raise

    def close(self) -> None:
        """Close the persistent jump connection."""
        if self.client:
            try:
                self.client.close()
                log.debug("Closed jump host connection %s", self.jump_host)
            finally:
                self.client = None

    def __enter__(self) -> "JumpManager":
        self.connect()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()