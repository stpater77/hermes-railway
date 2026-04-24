"""Microsoft OAuth PKCE flow for Hermes Office 365 / Microsoft Graph access."""

from __future__ import annotations

import base64
import hashlib
import http.server
import json
import os
import secrets
import stat
import threading
import time
import urllib.parse
import urllib.request
import webbrowser
from pathlib import Path
from typing import Any


# === CONFIG ===

AUTH_ENDPOINT_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/authorize"
TOKEN_ENDPOINT_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

REDIRECT_HOST = "127.0.0.1"
DEFAULT_REDIRECT_PORT = 8086
CALLBACK_PATH = "/oauth2callback"

SCOPES = "offline_access User.Read https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/Calendars.ReadWrite https://graph.microsoft.com/Files.ReadWrite https://graph.microsoft.com/Tasks.ReadWrite"

TOKEN_SKEW_SECONDS = 60


# === PATHS ===

def _hermes_home() -> Path:
    return Path(os.path.expanduser("~/.hermes"))


def _credentials_path() -> Path:
    return _hermes_home() / "auth" / "microsoft_oauth.json"


# === ENV CONFIG ===

def _load_env_value(key: str) -> str:
    env_path = os.path.expanduser("~/.hermes/.env")

    if key in os.environ:
        return os.environ[key].strip()

    if not os.path.exists(env_path):
        return ""

    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            if k.strip() == key:
                return v.strip().strip('"').strip("'")

    return ""


def _get_client_id() -> str:
    return _load_env_value("HERMES_MICROSOFT_CLIENT_ID")


def _get_tenant_id() -> str:
    tenant_id = _load_env_value("HERMES_MICROSOFT_TENANT_ID")
    return tenant_id or "common"


# === PKCE HELPERS ===

def _generate_pkce_pair() -> tuple[str, str]:
    verifier = base64.urlsafe_b64encode(os.urandom(64)).decode("utf-8").rstrip("=")
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode("utf-8")).digest()
    ).decode("utf-8").rstrip("=")
    return verifier, challenge


def build_auth_url(client_id: str, redirect_uri: str, state: str, challenge: str) -> str:
    tenant_id = _get_tenant_id()
    auth_endpoint = AUTH_ENDPOINT_TEMPLATE.format(tenant_id=tenant_id)

    params = {
        "client_id": client_id,
        "response_type": "code",
        "redirect_uri": redirect_uri,
        "response_mode": "query",
        "scope": SCOPES,
        "state": state,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
    }

    return auth_endpoint + "?" + urllib.parse.urlencode(params)


# === TOKEN STORAGE ===

def save_token_payload(payload: dict[str, Any]) -> None:
    path = _credentials_path()
    path.parent.mkdir(parents=True, exist_ok=True)

    now = int(time.time())
    expires_in = int(payload.get("expires_in", 3600))
    payload["expires_at"] = now + expires_in

    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    os.chmod(path, stat.S_IRUSR | stat.S_IWUSR)


def load_token_payload() -> dict[str, Any] | None:
    path = _credentials_path()
    if not path.exists():
        return None

    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def has_valid_access_token() -> bool:
    payload = load_token_payload()
    if not payload:
        return False

    access_token = payload.get("access_token")
    expires_at = int(payload.get("expires_at", 0))

    return bool(access_token and expires_at > int(time.time()) + TOKEN_SKEW_SECONDS)


# === TOKEN EXCHANGE + REFRESH ===

def exchange_code_for_token(code: str, redirect_uri: str, verifier: str) -> dict[str, Any]:
    client_id = _get_client_id()
    tenant_id = _get_tenant_id()

    token_endpoint = TOKEN_ENDPOINT_TEMPLATE.format(tenant_id=tenant_id)

    data = {
        "client_id": client_id,
        "scope": SCOPES,
        "code": code,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
        "code_verifier": verifier,
    }

    encoded = urllib.parse.urlencode(data).encode("utf-8")

    req = urllib.request.Request(
        token_endpoint,
        data=encoded,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )

    with urllib.request.urlopen(req, timeout=30) as resp:
        payload = json.loads(resp.read().decode("utf-8"))

    save_token_payload(payload)
    return payload


def refresh_access_token() -> dict[str, Any]:
    existing = load_token_payload()
    if not existing or not existing.get("refresh_token"):
        raise RuntimeError("No Microsoft refresh token found. Run login() first.")

    client_id = _get_client_id()
    tenant_id = _get_tenant_id()
    token_endpoint = TOKEN_ENDPOINT_TEMPLATE.format(tenant_id=tenant_id)

    data = {
        "client_id": client_id,
        "scope": SCOPES,
        "refresh_token": existing["refresh_token"],
        "grant_type": "refresh_token",
    }

    encoded = urllib.parse.urlencode(data).encode("utf-8")

    req = urllib.request.Request(
        token_endpoint,
        data=encoded,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )

    with urllib.request.urlopen(req, timeout=30) as resp:
        payload = json.loads(resp.read().decode("utf-8"))

    if not payload.get("refresh_token"):
        payload["refresh_token"] = existing["refresh_token"]

    save_token_payload(payload)
    return payload


def get_access_token() -> str:
    payload = load_token_payload()

    if payload and has_valid_access_token():
        return str(payload["access_token"])

    refreshed = refresh_access_token()
    return str(refreshed["access_token"])


# === CALLBACK HANDLER ===

class _OAuthCallbackHandler(http.server.BaseHTTPRequestHandler):
    expected_state: str = ""
    captured_code: str | None = None
    captured_error: str | None = None
    ready: threading.Event | None = None

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path != CALLBACK_PATH:
            self.send_response(404)
            self.end_headers()
            return

        params = urllib.parse.parse_qs(parsed.query)
        state = (params.get("state") or [""])[0]
        code = (params.get("code") or [""])[0]
        error = (params.get("error") or [""])[0]

        if state != type(self).expected_state:
            type(self).captured_error = "state_mismatch"
        elif error:
            type(self).captured_error = error
        elif code:
            type(self).captured_code = code
        else:
            type(self).captured_error = "no_code"

        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"You can close this window and return to Hermes.")

        if type(self).ready:
            type(self).ready.set()

    def log_message(self, format, *args):
        return


def _bind_callback_server(preferred_port: int = DEFAULT_REDIRECT_PORT):
    try:
        server = http.server.HTTPServer(
            (REDIRECT_HOST, preferred_port),
            _OAuthCallbackHandler,
        )
        return server, preferred_port
    except OSError:
        server = http.server.HTTPServer(
            (REDIRECT_HOST, 0),
            _OAuthCallbackHandler,
        )
        return server, server.server_address[1]


# === LOGIN FLOW ===

def login(open_browser: bool = True) -> dict[str, Any]:
    client_id = _get_client_id()
    if not client_id:
        raise RuntimeError("Missing HERMES_MICROSOFT_CLIENT_ID in ~/.hermes/.env")

    verifier, challenge = _generate_pkce_pair()
    state = secrets.token_urlsafe(16)

    server, port = _bind_callback_server(DEFAULT_REDIRECT_PORT)
    redirect_uri = f"http://{REDIRECT_HOST}:{port}{CALLBACK_PATH}"

    _OAuthCallbackHandler.expected_state = state
    _OAuthCallbackHandler.captured_code = None
    _OAuthCallbackHandler.captured_error = None
    ready = threading.Event()
    _OAuthCallbackHandler.ready = ready

    auth_url = build_auth_url(client_id, redirect_uri, state, challenge)

    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    print("\nOpen this Microsoft login URL:\n")
    print(auth_url)
    print("\nWaiting for Microsoft callback...\n")

    if open_browser:
        webbrowser.open(auth_url)

    ready.wait(300)
    server.shutdown()
    server.server_close()

    if _OAuthCallbackHandler.captured_error:
        raise RuntimeError(f"Microsoft OAuth failed: {_OAuthCallbackHandler.captured_error}")

    if not _OAuthCallbackHandler.captured_code:
        raise RuntimeError("No authorization code captured.")

    payload = exchange_code_for_token(
        _OAuthCallbackHandler.captured_code,
        redirect_uri,
        verifier,
    )

    print("Microsoft OAuth token saved to ~/.hermes/auth/microsoft_oauth.json")
    return payload


def start_login_test(open_browser: bool = True) -> str:
    payload = login(open_browser=open_browser)
    return payload.get("access_token", "")
