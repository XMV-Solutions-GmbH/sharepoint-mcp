# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Token persistence backends.

Two implementations behind a `TokenStore` Protocol:

- `KeyringTokenStore`: uses python-keyring, which delegates to the
  active OS keyring (Secret Service / Keychain / Credential Locker).
  Available when the OS provides one of those.
- `EncryptedFileTokenStore`: cryptography.fernet ciphertext on disk
  with a Scrypt-derived key from a passphrase env var. Used as the
  fallback for headless servers and CI environments.

Auto-detection picks keyring when the active backend is real (not
`fail.Keyring` and not a known-plaintext class), file otherwise.
`SP_TOKEN_STORE=keyring|file` overrides the auto-pick.

Rationale: see docs/spikes/2026-05-06-keyring-vs-encrypted-file.md.
"""

from __future__ import annotations

import os
import secrets
from base64 import urlsafe_b64encode
from pathlib import Path
from typing import Protocol

import keyring
import keyring.backend
import keyring.backends.fail
import keyring.errors
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives.kdf.scrypt import Scrypt

KEYRING_SERVICE = "sharepoint-mcp"
DEFAULT_CACHE_DIR = Path.home() / ".cache" / "sharepoint-mcp"
PASSPHRASE_ENV = "SP_TOKEN_PASSPHRASE"
STORE_OVERRIDE_ENV = "SP_TOKEN_STORE"

# Scrypt parameters chosen for ~50ms KDF on a typical laptop —
# memory-hard enough to defeat GPU brute force on a leaked file
# while keeping CLI startup snappy.
_SCRYPT_N = 2**14
_SCRYPT_R = 8
_SCRYPT_P = 1
_SALT_BYTES = 16


class NoUsableTokenStoreError(RuntimeError):
    """Raised when no token-store backend can be activated.

    Either keyring has no real backend AND `SP_TOKEN_PASSPHRASE` is
    unset, or the user gave an invalid `SP_TOKEN_STORE` override.
    """


class TokenStore(Protocol):
    """Persistence interface for OAuth tokens, namespaced by profile.

    Implementations must be safe to instantiate multiple times against
    the same underlying storage; concurrent access is not required for
    v0.1 (one MCP process per profile).
    """

    def get(self, profile: str) -> bytes | None:
        """Return stored bytes for `profile`, or None if not stored."""
        ...

    def set(self, profile: str, value: bytes) -> None:
        """Store `value` under `profile`. Overwrites if it exists."""
        ...

    def delete(self, profile: str) -> None:
        """Remove stored value for `profile`. No-op if not present."""
        ...


class KeyringTokenStore:
    """python-keyring-backed token store.

    Stores tokens as the password value under a `(service, profile)`
    key pair in the active OS keyring. Requires a real backend; calls
    raise `keyring.errors.NoKeyringError` on `fail.Keyring`.
    """

    def __init__(self, service: str = KEYRING_SERVICE) -> None:
        self._service = service

    def get(self, profile: str) -> bytes | None:
        value = keyring.get_password(self._service, profile)
        return value.encode() if value is not None else None

    def set(self, profile: str, value: bytes) -> None:
        keyring.set_password(self._service, profile, value.decode())

    def delete(self, profile: str) -> None:
        try:
            keyring.delete_password(self._service, profile)
        except keyring.errors.PasswordDeleteError:
            # Already absent — match Protocol's "no-op if not present" contract.
            pass


class EncryptedFileTokenStore:
    """Fernet-encrypted-file token store.

    Layout per profile under `base_dir`:

        <base_dir>/<profile>/token.enc    Fernet ciphertext
        <base_dir>/<profile>/token.salt   16 random bytes (Scrypt salt)

    Both files are mode 0o600 (owner-only) on POSIX.

    Passphrase is read from the env var named by `passphrase_env` at
    every call — it is never cached in memory beyond the duration of
    the call. A wrong passphrase produces `cryptography.fernet.InvalidToken`
    on `get()`.
    """

    def __init__(
        self,
        base_dir: Path | None = None,
        passphrase_env: str = PASSPHRASE_ENV,
    ) -> None:
        self._base_dir = base_dir if base_dir is not None else DEFAULT_CACHE_DIR
        self._passphrase_env = passphrase_env

    def _passphrase(self) -> bytes:
        pp = os.environ.get(self._passphrase_env, "").encode()
        if not pp:
            raise NoUsableTokenStoreError(
                f"Encrypted-file token store requires the {self._passphrase_env} "
                "environment variable to be set and non-empty.",
            )
        return pp

    def _profile_dir(self, profile: str) -> Path:
        d = self._base_dir / profile
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _derive_key(self, salt: bytes) -> bytes:
        kdf = Scrypt(salt=salt, length=32, n=_SCRYPT_N, r=_SCRYPT_R, p=_SCRYPT_P)
        return urlsafe_b64encode(kdf.derive(self._passphrase()))

    def get(self, profile: str) -> bytes | None:
        d = self._profile_dir(profile)
        token_file = d / "token.enc"
        salt_file = d / "token.salt"
        if not token_file.exists() or not salt_file.exists():
            return None
        salt = salt_file.read_bytes()
        ciphertext = token_file.read_bytes()
        return Fernet(self._derive_key(salt)).decrypt(ciphertext)

    def set(self, profile: str, value: bytes) -> None:
        d = self._profile_dir(profile)
        salt_file = d / "token.salt"
        if salt_file.exists():
            salt = salt_file.read_bytes()
        else:
            salt = secrets.token_bytes(_SALT_BYTES)
            salt_file.write_bytes(salt)
            salt_file.chmod(0o600)
        ciphertext = Fernet(self._derive_key(salt)).encrypt(value)
        token_file = d / "token.enc"
        token_file.write_bytes(ciphertext)
        token_file.chmod(0o600)

    def delete(self, profile: str) -> None:
        d = self._profile_dir(profile)
        for name in ("token.enc", "token.salt"):
            try:
                (d / name).unlink()
            except FileNotFoundError:
                pass


def _is_real_keyring_backend(kr: keyring.backend.KeyringBackend) -> bool:
    """Return True if `kr` is a real OS keychain integration.

    Excludes `fail.Keyring` (placeholder when no backend is available)
    and any backend whose class name suggests plaintext storage
    (`keyrings.alt.file.PlaintextKeyring` and friends).
    """
    if isinstance(kr, keyring.backends.fail.Keyring):
        return False
    if "Plaintext" in type(kr).__name__:
        return False
    return True


def get_token_store() -> TokenStore:
    """Pick a token-store backend for the current environment.

    Resolution order:

    1. `SP_TOKEN_STORE=keyring` or `=file` — explicit, no auto-detect.
    2. Auto: keyring if a real OS backend is detected.
    3. Auto: encrypted-file if `SP_TOKEN_PASSPHRASE` is set.
    4. Otherwise raise `NoUsableTokenStoreError` with a message that
       names both options the user could enable.

    Note: this returns a backend; the backend may still raise at
    `set()` / `get()` time if e.g. a forced keyring choice has no
    real backend available. We don't probe at construction.
    """
    forced = os.environ.get(STORE_OVERRIDE_ENV, "").strip().lower()
    if forced == "keyring":
        return KeyringTokenStore()
    if forced == "file":
        return EncryptedFileTokenStore()
    if forced:
        raise NoUsableTokenStoreError(
            f"{STORE_OVERRIDE_ENV} must be 'keyring' or 'file', got {forced!r}",
        )

    if _is_real_keyring_backend(keyring.get_keyring()):
        return KeyringTokenStore()

    if os.environ.get(PASSPHRASE_ENV):
        return EncryptedFileTokenStore()

    raise NoUsableTokenStoreError(
        "No usable token store. Either: (a) install a real OS keyring "
        "backend (Linux: gnome-keyring-daemon or KWallet running with a "
        "session bus; macOS: Keychain Access; Windows: Credential Locker), "
        f"or (b) set {PASSPHRASE_ENV} to enable the encrypted-file backend. "
        "See docs/spikes/2026-05-06-keyring-vs-encrypted-file.md.",
    )
