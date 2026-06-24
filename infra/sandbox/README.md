# Agent terminal sandbox image

The canonical Dockerfile that layers `@officeai/agent` (Node CLI) on
top of a host-provided agent base image
lives at [`packages/agent/sandbox/Dockerfile`](../../packages/agent/sandbox/Dockerfile)
so it ships inside the npm tarball (`@officeai/agent` `files` array).

Consumers can extract it without a git clone:

```sh
npm pack @officeai/agent
tar -xf officeai-agent-*.tgz package/sandbox/Dockerfile
docker build \
  --build-arg OFFICEAI_AGENT_BASE=officeai-agent-base:latest \
  --build-arg OFFICEAI_VERSION=$(npm view @officeai/agent version) \
  -t officeai-agent-sandbox:latest \
  -f package/sandbox/Dockerfile .
```

Downstream hosts may either build this Dockerfile directly or mirror it
into their own deploy pipeline. When mirroring it, keep the inline copy
in sync with this canonical file.
