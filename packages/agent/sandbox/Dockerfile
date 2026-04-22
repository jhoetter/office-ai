# Reference Dockerfile for the office-ai flavour of hof-engine's agent
# terminal sandbox. Layers Node 20 + the published `@officeai/agent`
# CLI on top of `hof-skill-base` so the LLM agent can call
# `office-agent docx|xlsx|pptx|pdf …` from the terminal alongside the
# existing Python skill set (openpyxl, python-pptx, pdfplumber, …).
#
# hof-os mirrors this Dockerfile inline (no build context needed) in
# backend/app/services/deployer/compose.py so each Hetzner host can
# build the image locally during deploy. This file is the canonical
# spec — keep the two in sync.
#
# Build args:
#   HOF_SKILL_BASE     base image to extend (default: hof-skill-base:latest)
#   OFFICEAI_VERSION   exact npm version of @officeai/agent to install
#
# Local build:
#   docker build \
#     --build-arg HOF_SKILL_BASE=hof-skill-base:latest \
#     --build-arg OFFICEAI_VERSION=0.1.0 \
#     -t hof-skill-base-officeai:0.1.0 .
ARG HOF_SKILL_BASE=hof-skill-base:latest
FROM ${HOF_SKILL_BASE}

USER root
RUN apt-get update \
 && apt-get install -y --no-install-recommends curl ca-certificates gnupg \
 && curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
 && apt-get install -y --no-install-recommends nodejs \
 && rm -rf /var/lib/apt/lists/*

# Pinned npm install — placed last so a version bump only re-runs this
# layer; the heavy Node + apt layer is cached forever per HOF_SKILL_BASE.
ARG OFFICEAI_VERSION
RUN test -n "${OFFICEAI_VERSION}" \
 && npm install -g @officeai/agent@${OFFICEAI_VERSION} \
 && office-agent --version

USER sandbox
WORKDIR /workspace
CMD ["sleep", "infinity"]
