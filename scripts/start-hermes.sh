#!/usr/bin/env bash
set -euo pipefail

mkdir -p /app/.hermes

cat > /app/.hermes/.env <<ENVEOF
API_SERVER_ENABLED=true
API_SERVER_KEY=${API_SERVER_KEY}
API_SERVER_HOST=0.0.0.0
OPENAI_API_KEY=${OPENAI_API_KEY}
OLLAMA_API_KEY=${OLLAMA_API_KEY}
OLLAMA_BASE_URL=${OLLAMA_BASE_URL:-https://ollama.com/v1}
HERMES_MICROSOFT_CLIENT_ID=${HERMES_MICROSOFT_CLIENT_ID:-}
HERMES_MICROSOFT_TENANT_ID=${HERMES_MICROSOFT_TENANT_ID:-}
ENVEOF

cat > /app/.hermes/config.yaml <<CFGEOF
model:
  default: "${MODEL_DEFAULT:-gpt-oss:120b}"
  provider: "${MODEL_PROVIDER:-ollama-cloud}"
  base_url: "${MODEL_BASE_URL:-https://ollama.com/v1}"

stt:
  enabled: true
  provider: local
  wake_word: "hey hermes"
  activation_phrases:
    - "hey hermes"
    - "ok hermes"
    - "hermes start"
  trigger_map:
    "read this": "/tts speak"

tts:
  provider: edge

web:
  backend: tavily

platform_toolsets:
  api_server:
    - web
    - memory
    - todo
    - microsoft365
CFGEOF

cat > /app/.hermes/SOUL.md <<'SOULEOF'
Behavior rules:

1. For Gmail, Google Calendar, Google Drive, Google Docs, and Google Sheets, always use the google-workspace integration.
2. Never use himalaya for Gmail or Outlook when a provider-native integration exists.
3. Do not ask for IMAP, SMTP, app passwords, or generic mailbox credentials for Gmail or Outlook if native OAuth/API integrations are configured.
4. If a Gmail request is made, prefer google-workspace automatically.
5. If an Outlook request is made, prefer the Microsoft-native integration automatically when it is available.
6. If no provider-native integration exists, state that clearly instead of falling back to himalaya.

Configuration invariants:
7. Never remove, overwrite, or change the model/provider/base_url configuration unless explicitly instructed by Sean.
8. Preserve these model settings whenever modifying YAML or config files:
   - model.default = gpt-oss:120b
   - model.provider = ollama-cloud
   - model.base_url = https://ollama.com/v1
9. If STT, TTS, or other YAML settings are changed, merge them without deleting or replacing the model block.
10. Treat the model/provider/base_url block as protected configuration.
SOULEOF

if [ -n "${GOOGLE_CLIENT_SECRET_JSON:-}" ]; then
  printf '%s' "$GOOGLE_CLIENT_SECRET_JSON" > /app/.hermes/google_client_secret.json
fi

if [ -n "${GOOGLE_TOKEN_JSON:-}" ]; then
  printf '%s' "$GOOGLE_TOKEN_JSON" > /app/.hermes/google_token.json
fi

mkdir -p /app/.hermes/skills/productivity
if [ -d /opt/microsoft-365-skill ]; then
  rm -rf /app/.hermes/skills/productivity/microsoft-365
  cp -R /opt/microsoft-365-skill /app/.hermes/skills/productivity/microsoft-365
fi

mkdir -p /app/.hermes/auth
if [ -n "${MICROSOFT_TOKEN_JSON:-}" ]; then
  printf '%s' "$MICROSOFT_TOKEN_JSON" > /app/.hermes/auth/microsoft_oauth.json
  chmod 600 /app/.hermes/auth/microsoft_oauth.json
fi

export HOME=/app
export HERMES_HOME=/app/.hermes
export PATH="/opt/hermes-agent/venv/bin:$PATH"

exec /opt/hermes-agent/venv/bin/hermes gateway
