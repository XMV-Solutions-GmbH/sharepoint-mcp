<!--
SPDX-License-Identifier: MIT OR Apache-2.0
SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
SPDX-FileContributor: David Koller <david.koller@xmv.de>
-->

# Spike: keyring vs encrypted-file fallback for headless servers

**Date**: 2026-05-06
**Issue**: [#9](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/9)
**Revised**: 2026-05-07 — original decision required `SP_TOKEN_PASSPHRASE` for headless installs, which created unnecessary friction (no other CLI tool does this). Adjusted to add `PlainFileTokenStore` (mode 0600) as the universal fallback; encrypted-file becomes opt-in via passphrase. See "Revision" section at the bottom.

**Decision (current)**: **three backends** — `keyring` (preferred when real OS backend exists), `PlainFileTokenStore` (default fallback, mode 0600 JSON, same convention as `gh auth` / `aws configure` / `npm login`), `EncryptedFileTokenStore` (opt-in via `SP_TOKEN_PASSPHRASE`, useful for CI). Auto-detect at first use with `SP_TOKEN_STORE=keyring|file|encrypted-file` as override. **No env vars required for the typical install.**

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

- Trying to autostart `gnome-keyring-daemon` from our process. That's an OS-level concern, not a tool concern.
- A custom `keyring` backend implementation (e.g., registering an encrypted-file backend with the keyring library itself). Cleaner-looking but adds magic and another extension surface to audit. Direct ABC is more honest.
- **Silent** plaintext-fallback while pretending to be a keyring. The `keyrings.alt.file.PlaintextKeyring` backend is explicitly rejected by `_is_real_keyring_backend`. (See revision below for what changed about the explicit, documented plain-file path.)

## Revision (2026-05-07)

The original decision required `SP_TOKEN_PASSPHRASE` to be set whenever no keyring was available. First contact with the actual user UX revealed this is too much friction:

- No comparable CLI tool requires this. `gh auth login`, `aws configure`, `npm login`, `git credential` all write to a mode-0600 JSON file in the user's home and trust the local user account. The "set an env var before logging in" step doesn't exist in any of those.
- The original "no plaintext-on-disk, ever" rule was a reaction to `python-keyring`'s sneaky `keyrings.alt.PlaintextKeyring` fallback (silently downgrading from a secure-looking keyring API to plaintext). That risk is real and we still defend against it. But an **explicit, documented** plain-file backend is a different thing — it doesn't pretend to be encrypted, and it's the standard convention for this class of tool.

### Adjusted backend roster

| Tier | Backend | When | Setup |
|---|---|---|---|
| 1 | `KeyringTokenStore` | macOS / Windows / Linux desktop with Secret Service | none |
| 2 | `PlainFileTokenStore` (NEW) | headless Linux, CI without secrets, default fallback | none |
| 3 | `EncryptedFileTokenStore` | opt-in via `SP_TOKEN_PASSPHRASE`; useful for CI where passphrase + ciphertext are separate secrets | env var |

`SP_TOKEN_STORE=keyring|file|encrypted-file` forces one specifically.

### Plain-file layout

`<base_dir>/<profile>/token.json`, mode `0o600`, JSON of the `CachedToken` dict. Directory created on demand. No salt file (no encryption). Same security as `~/.ssh/id_rsa`.

### What stays the same

- `_is_real_keyring_backend` still rejects `fail.Keyring` and any class whose name contains `Plaintext`. We do not silently downgrade through `python-keyring`.
- Encrypted-file backend stays as an opt-in path for users who want belt-and-suspenders, and for CI where the passphrase is in one secret and the ciphertext+salt are in two more.
- The CI workflow sketch above (passphrase + base64 ciphertext + base64 salt as three secrets) still works for users who want encryption in CI. Users who don't can simpler-store the plaintext JSON as a single secret and write it directly to `~/.cache/sharepoint-mcp/<profile>/token.json` in CI.

### Why this is still safe

`gh`, `npm`, `aws`, `kubectl`, `ssh` — all of these put credentials in mode-0600 files in `$HOME`. They've shipped to billions of installs over decades. The threat model is well-understood: if your local user account is compromised, you have bigger problems than a credential file. Our previous "passphrase required" rule was over-corrective relative to the industry norm and the actual threat model.

## Follow-ups landed by this spike

- `cryptography>=42` was added to `pyproject.toml` runtime deps with #10. (Still required because `EncryptedFileTokenStore` is opt-in but live.)
- `keyring>=25` likewise.
- `docs/app-concept.md` open-question #2 closed.
- `KeyringTokenStore`, `EncryptedFileTokenStore`, and `PlainFileTokenStore` implementations come with #10 (auth module).
- CI harness-job sketch (encryption variant) above stays as one option; users who prefer plain-secret-as-JSON write directly to `token.json` and skip the encryption.
