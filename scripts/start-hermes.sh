#!/usr/bin/env bash
set -euo pipefail

mkdir -p /app/.hermes

cat > /app/.hermes/.env <<ENVEOF
API_SERVER_ENABLED=true
API_SERVER_KEY=${API_SERVER_KEY}
OPENAI_API_KEY=${OPENAI_API_KEY}
ENVEOF

cat > /app/.hermes/config.yaml <<CFGEOF
model:
  default: "${MODEL_DEFAULT:-gpt-5.4}"
CFGEOF

if [ -n "${GOOGLE_CLIENT_SECRET_JSON:-}" ]; then
  printf '%s' "$GOOGLE_CLIENT_SECRET_JSON" > /app/.hermes/google_client_secret.json
fi

if [ -n "${GOOGLE_TOKEN_JSON:-}" ]; then
  printf '%s' "$GOOGLE_TOKEN_JSON" > /app/.hermes/google_token.json
fi

export HOME=/app
export HERMES_HOME=/app/.hermes
export PATH="/opt/hermes-agent/venv/bin:$PATH"

exec /opt/hermes-agent/venv/bin/hermes gateway
