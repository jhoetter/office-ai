# Release pipeline (office-ai → GitHub Releases → hof-os)

> Auto-publishes every push to `main` as a new patch release: builds a
> self-contained `@officeai/agent` bundle via `pnpm deploy`, attaches it
> to a GitHub Release, then commits the new version pin into hof-os.
> Mirrors the [hof-engine ↔ hof-os
> pattern](https://github.com/jhoetter/hof-engine/blob/main/.github/workflows/auto-release.yml)
> — but adapted for a Node monorepo: the artefact is a tarball on a
> GitHub Release instead of a wheel on PyPI / a tarball on npm. **No npm
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
   ↓ pnpm --filter @officeai/agent --prod deploy → self-contained dir
   ↓ tar czf officeai-agent-X.Y.Z.tgz …
   ↓ gh release create vX.Y.Z officeai-agent-X.Y.Z.tgz
   ↓
notify-hof-os: rewrite infra/officeai.lock.json → push to hof-os/main
   ↓
hof-os deploy curls
   https://github.com/jhoetter/office-ai/releases/download/vX.Y.Z/officeai-agent-X.Y.Z.tgz
into the sandbox image
```

To skip a release for a doc-only change, append `[skip ci]` or
`[skip release]` to the commit message.

## Why GitHub Releases instead of npm

`@officeai/agent` is a Node CLI with 9 internal `workspace:*` deps. The
`pnpm --filter @officeai/agent --prod deploy <out>` step resolves every
workspace dep into a fully populated `node_modules/` inside `<out>`, so
the resulting tarball is **self-contained** — the consumer (the hof-os
sandbox Dockerfile) just downloads it and extracts; it never needs to
reach a registry.

Trade-off vs. publishing each `@officeai/*` package separately to npm:

- ✅ Zero external account claims (no npm org, no Trusted Publishers).
- ✅ One artefact per release instead of 12.
- ✅ Stays valid even if office-ai is private (just pass an
  `OFFICEAI_DOWNLOAD_TOKEN` to docker build — see hof-os deployer).
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

Only `@officeai/agent` is actually packaged for release — the other
packages travel inside its `node_modules/` after `pnpm deploy`.
Anything not in the list (`apps/web`, `ui`, `design-tokens`,
`comments`, `realtime`, `realtime-server`, `tests`) stays
`"private": true`.

## One-time external setup

Two GitHub-side actions, no third-party accounts:

### 1. Make office-ai public

`office-ai` repo → **Settings → General → Danger Zone → Change
visibility → Public**. With the repo public, the GitHub Releases
download URLs are anonymous — hof-os's sandbox Dockerfile can `curl`
them with no auth header.

(If you must keep office-ai private: skip this step and add
`OFFICEAI_DOWNLOAD_TOKEN` as a secret on hof-os; the `Dockerfile.officeai-sandbox`
will need a one-line tweak to forward it as `Authorization: Bearer …`.
Asks for a separate plan if you want this branch wired up.)

### 2. Cross-repo PAT for hof-os bump

Generate a fine-grained PAT (`Settings → Developer settings →
Personal access tokens → Fine-grained tokens`) named
`OFFICEAI_SYNC_TOKEN`:

- Resource owner: `jhoetter`
- Repository access: only `jhoetter/hof-os`
- Permissions: `Contents: Read and write`, `Metadata: Read-only`
- Expiration: 1 year (renew via reminder)

Add it as a repository secret on `office-ai`:
**Settings → Secrets and variables → Actions → New repository
secret** named `OFFICEAI_SYNC_TOKEN`.

The `notify-hof-os` job uses this token to push the
`infra/officeai.lock.json` bump.

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

`bundle:dry-run` runs the same `pnpm --filter @officeai/agent --prod
deploy` as CI into a temp directory, sanity-checks that
`dist/cli.js` exists with a node shebang and `node_modules/` got
populated, prints the bundle size, then deletes the temp dir. Use it
before merging anything that touches a publishable `package.json`.
