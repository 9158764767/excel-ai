#!/usr/bin/env bash

# Simple script to run the FastAPI app locally with hot reload. Assumes the
# dependencies from requirements.txt are installed in the current environment.
set -euo pipefail

if [[ -z "${OPENAI_API_KEY:-}" ]]; then
  echo "OPENAI_API_KEY is not set. Please set it before running." >&2
  exit 1
fi

# Default model; allow override via environment.
export OPENAI_MODEL="${OPENAI_MODEL:-gpt-4o}"
export REQUIRE_APP_KEY="${REQUIRE_APP_KEY:-0}"
export APP_KEY="${APP_KEY:-change-this-key}"

exec uvicorn app.main:app --reload --port 8000
