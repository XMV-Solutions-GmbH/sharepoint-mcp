<!--
SPDX-License-Identifier: MIT OR Apache-2.0
SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
SPDX-FileContributor: David Koller <david.koller@xmv.de>
-->

# Spike: keyring vs encrypted-file fallback for headless servers

**Date**: 2026-05-06
**Issue**: [#9](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/9)
**Decision**: **dual backend** — `keyring` when it has a real backend, `cryptography.fernet`-encrypted file otherwise; auto-detect at first use with `SP_TOKEN_STORE` as override.

---

## Question

Where do we cache the OAuth access + refresh tokens?

- `keyring` (cross-platform, secure, but requires DBus/Secret-Service on Linux — fails on truly headless servers).
- An encrypted file with a passphrase from env var.

The concept (§ Authentication) flagged this as an open question; pick one or both.

## What I verified

### keyring on truly-headless Linux is honest

Tested on this box (Ubuntu 24.04, `XDG_SESSION_TYPE=tty`, no `gnome-keyring-daemon` running, DBus user-session present but no Secret Service):

```text
Active backend: keyring.backends.fail.Keyring
Priority: 0
```

Calls to `set_password` / `get_password` raise `keyring.errors.NoKeyringError` immediately. **No silent fallback to plaintext.** This is the load-bearing property — we can rely on keyring to refuse to store secrets when it can't actually secure them, instead of writing a plaintext file pretending to be safe.

### `cryptography.fernet` roundtrip works as expected

Verified:

- Scrypt (n=2¹⁴, r=8, p=1) → 32-byte key → `Fernet`. Roundtrip succeeds with the right passphrase.
- Wrong passphrase → `InvalidToken` raised cleanly. No partial-decrypt accept.
- Salt is per-token-blob (16 random bytes), stored alongside the ciphertext.

### Backend-detection probing

`keyring.get_keyring()` returns the active backend without doing I/O. We can detect "no real backend" by checking `isinstance(current, fail.Keyring)` or simply by attempting a no-op get-password against a known-impossible key and catching `NoKeyringError`. Either works.

## Decision

**Two `TokenStore` implementations behind a `Protocol`. Auto-pick. Env override.**

### Backend selection logic (at first token-store access)

```python
forced = os.environ.get("SP_TOKEN_STORE")  # "keyring" | "file" | unset

if forced == "keyring":
    return KeyringTokenStore()  # raises if no backend

if forced == "file":
    return EncryptedFileTokenStore()  # requires SP_TOKEN_PASSPHRASE

# Auto-detect:
keyring_store = KeyringTokenStore()
if keyring_store.has_real_backend():
    return keyring_store

# Keyring not usable. Try encrypted file.
if os.environ.get("SP_TOKEN_PASSPHRASE"):
    return EncryptedFileTokenStore()

raise NoUsableTokenStoreError(
    "No keyring backend available on this system, and SP_TOKEN_PASSPHRASE "
    "is not set. Either install gnome-keyring-daemon (Linux) / unlock "
    "the system keychain (macOS) / sign in to a session (Windows), or "
    "set SP_TOKEN_PASSPHRASE to enable the encrypted-file fallback. "
    "See docs/testconcept.md for the headless-server setup."
)
```

### Why both, not just one

- **Keyring-only** would lock out CI containers, ssh-into-a-VM developer flows, and any Linux server without a desktop session. That's where this MCP gets the most use (headless dev boxes for AI workflows). Hard pass.
- **File-only** would forfeit the OS-keychain integration on every developer's laptop, which is the more secure and less surprising option when it's available. Also forfeit zero-config install on macOS / Windows / Linux-with-DBus.

The dual-backend cost is one `TokenStore` Protocol + two ~50-line implementations + auto-detection. Manageable.

### File layout for the encrypted-file backend

```text
~/.cache/sharepoint-mcp/<profile>/
    token.enc          # Fernet ciphertext, ~280 bytes for a typical refresh+access pair
    token.salt         # 16 random bytes (Scrypt salt)
    working/           # already in concept § Architecture
```

Per-profile salt prevents cross-profile passphrase reuse from leaking. The salt is **not** secret — it's safe to store alongside the ciphertext.

### Passphrase source

`SP_TOKEN_PASSPHRASE` env var, no default. If unset when the file backend is selected, raise an error at the first store-access (not at startup; lazy) with the same clear message.

The passphrase IS a secret, but it protects only at-rest data on the same host the agent runs on. If an attacker has both disk access AND the passphrase, they have everything anyway. The encryption defends against:

- Backups / disk imaging picking up plaintext tokens.
- A different process on the same machine reading the file (assuming the env var is process-private).
- Accidental check-in or `scp` of the token file.

It does **not** defend against a compromised host running our process. That's the same threat model as keyring — fundamentally, this is a "trust the local OS user account" tool.

### KDF choice: Scrypt over PBKDF2

Scrypt (n=2¹⁴, r=8, p=1) is memory-hard and resists offline GPU brute force on a leaked file better than PBKDF2 with similar wall-clock cost. The `cryptography` package ships both; Scrypt is the modern default.

## CI implications

CI containers don't have keyring. The harness layer needs the encrypted-file backend by definition.

```yaml
# .github/workflows/ci.yml — harness job
env:
  SP_TOKEN_STORE: file
  SP_TOKEN_PASSPHRASE: ${{ secrets.SHAREPOINT_HARNESS_PASSPHRASE }}
steps:
  - name: Restore harness token from secret
    run: |
      mkdir -p ~/.cache/sharepoint-mcp/harness
      echo "${{ secrets.SHAREPOINT_HARNESS_TOKEN_ENC_B64 }}" | base64 -d \
          > ~/.cache/sharepoint-mcp/harness/token.enc
      echo "${{ secrets.SHAREPOINT_HARNESS_TOKEN_SALT_B64 }}" | base64 -d \
          > ~/.cache/sharepoint-mcp/harness/token.salt
  - name: Run harness tests
    run: ./tests/run_tests.sh harness
```

Three secrets per CI environment: passphrase, ciphertext, salt. Generated once on David's machine after the initial Device Code login (#28), encoded with the chosen passphrase, stored in repo secrets. Rotate when the refresh token expires (every 60–90 days per Microsoft).

## What this rules out

- A "plaintext-fallback because no other option" mode. Either keyring works, or the user provides a passphrase. There is no third path. We will not write tokens to disk in plaintext under any circumstance.
- Trying to autostart `gnome-keyring-daemon` from our process. That's an OS-level concern, not a tool concern.
- A custom `keyring` backend implementation (e.g., registering an encrypted-file backend with the keyring library itself). Cleaner-looking but adds magic and another extension surface to audit. Direct two-store ABC is more honest.

## Follow-ups landed by this spike

- `cryptography>=42` will be added to `pyproject.toml` runtime deps when #10 lands.
- `keyring>=25` likewise.
- `docs/app-concept.md` open-question #2 closed.
- `EncryptedFileTokenStore` and `KeyringTokenStore` implementations come with #10 (auth module).
- CI harness-job stub above is a sketch; the working version lands as part of #28 (harness gate) once the test refresh token exists to encrypt.
