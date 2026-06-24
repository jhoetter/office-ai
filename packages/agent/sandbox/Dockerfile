# Reference Dockerfile for the office-ai agent terminal sandbox.
# Layers Node 20 + the published `@officeai/agent` CLI on top of a
# host-provided base image so the LLM agent can call
# `office-agent docx|xlsx|pptx|pdf …` from the terminal alongside the
# existing Python skill set (openpyxl, python-pptx, pdfplumber, …).
#
# Downstream hosts may mirror this Dockerfile inline when they need a
# no-build-context deploy path. This file is the canonical spec — keep
# any mirrored copy in sync.
#
# Build args:
#   OFFICEAI_AGENT_BASE   base image to extend (default: officeai-agent-base:latest)
#   OFFICEAI_VERSION      exact npm version of @officeai/agent to install
#
# Local build:
#   docker build \
#     --build-arg OFFICEAI_AGENT_BASE=officeai-agent-base:latest \
#     --build-arg OFFICEAI_VERSION=0.1.0 \
#     -t officeai-agent-sandbox:0.1.0 .
ARG OFFICEAI_AGENT_BASE=officeai-agent-base:latest
FROM ${OFFICEAI_AGENT_BASE}

USER root
RUN apt-get update \
 && apt-get install -y --no-install-recommends curl ca-certificates gnupg \
 && curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
 && apt-get install -y --no-install-recommends nodejs \
 && rm -rf /var/lib/apt/lists/*

# Pinned npm install — placed last so a version bump only re-runs this
# layer; the heavy Node + apt layer is cached forever per OFFICEAI_AGENT_BASE.
ARG OFFICEAI_VERSION
RUN test -n "${OFFICEAI_VERSION}" \
 && npm install -g @officeai/agent@${OFFICEAI_VERSION} \
 && office-agent --version

USER sandbox
WORKDIR /workspace
CMD ["sleep", "infinity"]
