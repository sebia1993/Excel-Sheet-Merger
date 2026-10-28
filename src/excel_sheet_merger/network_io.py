"""Recognize explicit OS network/share failures without guessing from messages.

A lost SMB response is not proof that a published file was never created.
Callers must not blindly retry a write with an uncertain server outcome.
"""
from __future__ import annotations

import errno

_NETWORK_ERRNO_NAMES = (
    'ENETDOWN', 'ENETUNREACH', 'ENETRESET', 'ENOTCONN',
    'ECONNRESET', 'ECONNABORTED', 'EHOSTUNREACH', 'EHOSTDOWN',
    'ETIMEDOUT', 'ESTALE', 'EREMOTEIO',
)
_NETWORK_ERRNOS = frozenset(getattr(errno, name) for name in _NETWORK_ERRNO_NAMES if hasattr(errno, name))
# Windows errors when a remote share / session disappears or stops responding.
_NETWORK_WINERRORS = frozenset((53, 59, 64, 67, 121, 1231, 1232))


def is_network_io_error(error: BaseException) -> bool:
    """Whether *this* OS exception specifically denotes network/share failure.

    Bare EIO/PermissionError/ENOSPC are intentionally excluded. Use the
    exception cause chain at the diagnostic call site for wrapped errors.
    """
    return isinstance(error, OSError) and (
        error.errno in _NETWORK_ERRNOS
        or getattr(error, 'winerror', None) in _NETWORK_WINERRORS
    )
