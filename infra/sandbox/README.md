# Agent terminal sandbox image

The canonical Dockerfile that layers `@officeai/agent` (Node CLI) on
top of [hof-engine's `hof-skill-base`](https://github.com/jhoetter/hof-engine/blob/main/hof/agent/sandbox/Dockerfile.skill-base)
lives at [`packages/agent/sandbox/Dockerfile`](../../packages/agent/sandbox/Dockerfile)
so it ships inside the npm tarball (`@officeai/agent` `files` array).

Consumers can extract it without a git clone:

```sh
npm pack @officeai/agent
tar -xf officeai-agent-*.tgz package/sandbox/Dockerfile
docker build \
  --build-arg HOF_SKILL_BASE=hof-skill-base:latest \
  --build-arg OFFICEAI_VERSION=$(npm view @officeai/agent version) \
  -t hof-skill-base-officeai:latest \
  -f package/sandbox/Dockerfile .
```

The current production consumer is hof-os, which mirrors the
Dockerfile inline in
[`backend/app/services/deployer/compose.py`](https://github.com/jhoetter/hof-os/blob/main/backend/app/services/deployer/compose.py)
and builds it on each Hetzner host during data-app deploys (no
container registry required). When the Dockerfile here changes, update
the inline copy in compose.py too.
