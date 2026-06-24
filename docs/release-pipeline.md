# Release pipeline (office-ai → GitHub Releases → optional host sync)

> Auto-publishes every push to `main` as a new patch release: builds a
> self-contained `@officeai/agent` bundle via `pnpm deploy`, attaches it
> to a GitHub Release, then optionally commits the new version pin into
> a configured downstream host repository. The artifact is a tarball on
> a GitHub Release instead of a package registry publish. **No npm
> registry, no Trusted Publishers, no GHCR** — pure GitHub.

## Flow

```
push to main
   ↓
verify (pnpm verify, includes bundle:dry-run)
   ↓
release: bump every publishable package via scripts/bump-version.mjs
   ↓ commit "chore: bump to X.Y.Z [skip ci]"
   ↓ tag vX.Y.Z, push
   ↓ pnpm install (refresh lockfile after bump)
   ↓ pnpm build
   ↓ pnpm --filter @officeai/agent          --prod deploy → self-contained dir
   ↓ pnpm --filter @officeai/react-editors  --prod deploy → self-contained dir
   ↓ tar czf officeai-agent-X.Y.Z.tgz …
   ↓ tar czf officeai-react-editors-X.Y.Z.tgz …
   ↓ gh release create vX.Y.Z {agent,react-editors}-X.Y.Z.tgz
   ↓
notify-downstream-host: rewrite configured officeai lockfile
   ↓
downstream sandbox build  curls officeai-agent-X.Y.Z.tgz       (CLI for agents)
downstream web app        postinstall pulls
                          officeai-react-editors-X.Y.Z.tgz     (browser editors)
```

To skip a release for a doc-only change, append `[skip ci]` or
`[skip release]` to the commit message.

## Why GitHub Releases instead of npm

`@officeai/agent` is a Node CLI with 9 internal `workspace:*` deps. The
`pnpm --filter @officeai/agent --prod deploy <out>` step resolves every
workspace dep into a fully populated `node_modules/` inside `<out>`, so
the resulting tarball is **self-contained** — the consumer sandbox
Dockerfile just downloads it and extracts; it never needs to reach a
registry.

Trade-off vs. publishing each `@officeai/*` package separately to npm:

- ✅ Zero external account claims (no npm org, no Trusted Publishers).
- ✅ One artefact per release instead of 12.
- ✅ Stays valid even if office-ai is private (just pass an
  `OFFICEAI_DOWNLOAD_TOKEN` to docker build in the downstream deployer).
- ⚠️ Tarball is larger than 12 individual npm packages
  (~150 MB self-contained vs. summing-to-similar split tarballs).
  Only matters during the docker build's `RUN curl …` step, which
  caches at the Docker layer level and only re-runs when the version
  pin moves.

## Publishable packages

`@officeai/{core, text-formatting, docx, xlsx, pptx, pdf, pdf-edit,
pdf-forms, pdf-ocr, pdf-annotations, pdf-engine, agent}` are all
version-bumped in lockstep so the deployed bundle's resolved
dependency versions match the release tag. The list lives in
[`scripts/bump-version.mjs`](../scripts/bump-version.mjs).

Two packages are actually packaged for release — `@officeai/agent`
(headless CLI) and `@officeai/react-editors` (browser editor surfaces

- blank-file builders). The other publishable packages travel inside
  their `node_modules/` after `pnpm deploy`. Anything not in the list
  (`apps/web`, `ui`, `design-tokens`, `comments`, `realtime`,
  `realtime-server`, `tests`) stays `"private": true`.

### Lockfile schema

The optional `notify-downstream-host` job writes both pins into the
configured lockfile path, defaulting to `infra/officeai.lock.json`, so
a single bump informs the backend CLI and web editor integration at
once:

```json
{
  "version": "X.Y.Z",
  "agent_version": "X.Y.Z",
  "agent_tarball": "https://github.com/jhoetter/office-ai/releases/download/vX.Y.Z/officeai-agent-X.Y.Z.tgz",
  "react_editors_version": "X.Y.Z",
  "react_editors_tarball": "https://github.com/jhoetter/office-ai/releases/download/vX.Y.Z/officeai-react-editors-X.Y.Z.tgz",
  "published_at": "…",
  "source_repo": "jhoetter/office-ai",
  "source_sha": "…"
}
```

`version` stays for backwards compatibility with the bootstrap
lockfile shape.

## One-time external setup

Two GitHub-side actions, no third-party accounts:

### 1. Make office-ai public

`office-ai` repo → **Settings → General → Danger Zone → Change
visibility → Public**. With the repo public, the GitHub Releases
download URLs are anonymous — downstream sandbox Dockerfiles can `curl`
them with no auth header.

(If you must keep office-ai private: skip this step and add
`OFFICEAI_DOWNLOAD_TOKEN` as a secret on the downstream host; its
Dockerfile will need a one-line tweak to forward it as
`Authorization: Bearer …`. Ask for a separate plan if you want this
branch wired up.)

### 2. Optional cross-repo PAT for downstream host bump

Generate a fine-grained PAT (`Settings → Developer settings →
Personal access tokens → Fine-grained tokens`) named
`OFFICEAI_SYNC_TOKEN`:

- Resource owner: downstream repository owner
- Repository access: only the downstream host repository
- Permissions: `Contents: Read and write`, `Metadata: Read-only`
- Expiration: 1 year (renew via reminder)

Add it as a repository secret on `office-ai`:
**Settings → Secrets and variables → Actions → New repository
secret** named `OFFICEAI_SYNC_TOKEN`.

Add these repository variables on `office-ai` when you want the sync:

- `OFFICEAI_DOWNSTREAM_REPOSITORY=owner/repo`
- `OFFICEAI_DOWNSTREAM_LOCKFILE=infra/officeai.lock.json` (optional;
  this is the default)

The `notify-downstream-host` job uses this token and variables to push
the configured lockfile bump. If either the repository variable or token
is missing, the job exits cleanly and the GitHub Release remains the
complete standalone distribution artifact.

That's the entire external setup — no PyPI, no npm, no GHCR, no
GitHub Environment, no Trusted Publishers.

### Optional: seed the version

The bumper defaults to `0.1.0` when no `vX.Y.Z` tag exists, so the
very first run on a freshly enabled workflow will tag `v0.1.0` and
publish `officeai-agent-0.1.0.tgz`. If you want to start at a
different version, manually tag the bootstrap commit before merging
the workflow:

```sh
git tag v0.1.0
git push origin v0.1.0
```

The next push to `main` will then tag `v0.1.1`. Drop the tag and
re-tag if you want a different starting point.

## Local dry-run

```sh
pnpm verify             # full quality gate, ends with bundle:dry-run
pnpm bundle:dry-run     # only the bundle smoke-test
```

`bundle:dry-run` runs the same `pnpm --filter <pkg> --prod deploy`
as CI for **both** release artefacts (`@officeai/agent` and
`@officeai/react-editors`) into a temp directory, sanity-checks that
the expected entry points exist and `node_modules/` got populated
with the inlined `@officeai/*` workspace deps, prints the bundle
size, then deletes the temp dir. Use it before merging anything that
touches a publishable `package.json`.
