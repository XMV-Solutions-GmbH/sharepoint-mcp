<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->

# Releasing mcp-server-sharepoint

Steps to cut a new version. Follows [Semantic Versioning](https://semver.org/).

---

## Prerequisites (one-time)

### PyPI Trusted Publisher

This project uses [PyPI Trusted Publishers](https://docs.pypi.org/trusted-publishers/) for OIDC-based publishing — no long-lived API tokens stored anywhere. Setup steps:

1. **Reserve the project name on PyPI** (one-off):
   - Log in at <https://pypi.org/>.
   - Go to "Your projects" → "Add a pending publisher".
   - PyPI Project Name: `mcp-server-sharepoint`
   - Owner: `XMV-Solutions-GmbH`
   - Repository name: `mcp-server-sharepoint`
   - Workflow name: `release.yml`
   - Environment name: `pypi`
2. **Add a `pypi` deployment environment** in this repo (Settings → Environments → New environment "pypi"). Optional: require manual approval for production releases.
3. After the first successful publish, the pending-publisher record converts to a regular Trusted Publisher; subsequent releases just work.

### Alternative: API token

If Trusted Publisher isn't available (e.g., for private mirrors), generate a PyPI API token at <https://pypi.org/manage/account/token/> and store it as repo secret `PYPI_API_TOKEN`. Then change `release.yml`'s publish step to:

```yaml
- name: Publish to PyPI (API token)
  env:
    UV_PUBLISH_TOKEN: ${{ secrets.PYPI_API_TOKEN }}
  run: uv publish
```

Trusted Publisher is preferred — fewer secrets to rotate.

---

## Cutting a release

For a normal release (version bumps follow SemVer; v0.x increments freely while pre-1.0):

```bash
# 1. Update CHANGELOG.md — move [Unreleased] entries under a new versioned section
$EDITOR CHANGELOG.md

# 2. Bump version in pyproject.toml
$EDITOR pyproject.toml   # change version = "x.y.z"

# 3. Commit
git add CHANGELOG.md pyproject.toml
git commit -m "chore(release): vX.Y.Z"
git push origin main

# 4. Wait for CI green on main

# 5. Tag + push
git tag -a vX.Y.Z -m "Release vX.Y.Z"
git push origin vX.Y.Z
```

The `release.yml` workflow triggers on `v*` tag push:

1. Re-runs unit + integration tests as a gate.
2. Builds wheel + sdist via `uv build`.
3. Publishes to PyPI via OIDC.
4. Creates a GitHub Release with auto-generated notes.

Verify the published artifact:

```bash
uvx mcp-server-sharepoint@X.Y.Z --version
```

---

## Hotfix flow

For a bugfix on top of a release that's now diverged from `main`:

```bash
git checkout vX.Y.0
git checkout -b hotfix/X.Y.1
# ... fix ...
git tag -a vX.Y.1 -m "Hotfix vX.Y.1"
git push origin vX.Y.1
```

Then port the fix forward to `main` via PR.

---

## When to bump major / minor / patch

While pre-1.0:

- **Patch** (`v0.x.Y → v0.x.(Y+1)`): bug fixes, doc updates, internal refactors. No behaviour changes.
- **Minor** (`v0.X.* → v0.(X+1).0`): new tools, new features, behaviour changes that are additive (existing tools keep working).
- **Major** (`v0.* → v1.0.0`): the first stable release. Document the SemVer commitments in CHANGELOG. After v1.0, breaking changes require a major bump and a deprecation cycle.

For v0.1 specifically: the entire v0.1 line is "alpha — APIs may change". Don't depend on tool signatures from v0.x in third-party code yet.
