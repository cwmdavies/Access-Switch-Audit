"""
Network Automation Utilities Package
"""
__version__ = "1.0.1"

from .credentials import get_secret_with_fallback
from .jump_manager import JumpManager
from .config import JUMP_HOST
from .netmiko_utils import connect_to_device

__all__ = [
    "get_secret_with_fallback",
    "JumpManager",
    "JUMP_HOST",
    "connect_to_device",
]

# package logger
import logging
logger = logging.getLogger(__name__)