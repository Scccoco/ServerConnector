import json
import secrets
import hashlib
import ipaddress
import base64
import binascii
import hmac
import time
import re
import sqlite3
import subprocess
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import parse_qs, quote
from urllib.error import HTTPError, URLError
from urllib.request import Request as UrlRequest, urlopen

from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from pydantic import BaseModel


BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
DB_PATH = BASE_DIR / "connector.db"
FW_SCRIPT = BASE_DIR / "firewall_manager.ps1"
SMB_SCRIPT = BASE_DIR / "smb_user_manager.ps1"
ADMIN_UI_PATH = BASE_DIR / "admin_ui.html"
ADMIN_LOGIN_UI_PATH = BASE_DIR / "admin_login.html"
OPS_UI_PATH = BASE_DIR / "ops_ui.html"
UPDATES_DIR = BASE_DIR / "updates"
UPDATE_MANIFEST_PATH = UPDATES_DIR / "latest.json"
DEFAULT_TOKEN_EXPORT_DIR = Path(r"\\62.113.36.107\BIM_Models\Tokens")
DEFAULT_SMB_SERVER_HOST = "62.113.36.107"
DEFAULT_SMB_SHARE_NAME = "BIM_Models"
DEFAULT_SMB_SHARE_PATH = r"D:\BIM_Models"
DEFAULT_SMB_USER_PREFIX = "bim_"
ADMIN_SESSION_COOKIE = "connector_admin_session"
ADMIN_SESSION_TTL_SECONDS = 60 * 60 * 12
GITHUB_UPDATES_CACHE_TTL_SECONDS_DEFAULT = 300

_github_manifest_cache_key = ""
_github_manifest_cache_until = 0.0
_github_manifest_cache_value: dict | None = None


class Heartbeat(BaseModel):
    device_id: str
    public_ip: str | None = None
    hostname: str | None = None
    agent_version: str | None = None


class CreateTokenRequest(BaseModel):
    device_id: str
    issued_to: str | None = None


class RotateTokenRequest(BaseModel):
    issued_to: str | None = None


class BootstrapRequest(BaseModel):
    hostname: str | None = None
    agent_version: str | None = None
    public_ip: str | None = None


class NetworkRuleRequest(BaseModel):
    ip: str
    ports: list[int] | None = None


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise RuntimeError("Missing config.json (copy config.example.json)")
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))


def init_db() -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS devices (
                device_id TEXT PRIMARY KEY,
                public_ip TEXT NOT NULL,
                hostname TEXT,
                agent_version TEXT,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS device_tokens (
                device_id TEXT PRIMARY KEY,
                token_hash TEXT NOT NULL,
                issued_to TEXT,
                created_at TEXT NOT NULL,
                last_used_at TEXT,
                revoked_at TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS audit_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_type TEXT NOT NULL,
                device_id TEXT,
                actor TEXT,
                details TEXT,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS device_access (
                device_id TEXT PRIMARY KEY,
                smb_login TEXT NOT NULL,
                smb_username TEXT NOT NULL,
                smb_password TEXT NOT NULL,
                smb_share_unc TEXT NOT NULL,
                smb_share_path TEXT,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS device_sessions (
                device_id TEXT PRIMARY KEY,
                session_id TEXT NOT NULL,
                hostname TEXT,
                public_ip TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def hash_token(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def add_audit(event_type: str, device_id: str | None, actor: str | None, details: str | None) -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO audit_events(event_type, device_id, actor, details, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (event_type, device_id, actor, details, utc_now()),
        )


def check_token(device_id: str, token: str | None, cfg: dict) -> None:
    if not token:
        raise HTTPException(status_code=401, detail="Missing token")

    token_hash = hash_token(token)
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            """
            SELECT device_id FROM device_tokens
            WHERE device_id = ? AND token_hash = ? AND revoked_at IS NULL
            """,
            (device_id, token_hash),
        ).fetchone()
        if row:
            conn.execute(
                "UPDATE device_tokens SET last_used_at = ? WHERE device_id = ?",
                (utc_now(), device_id),
            )
            return

    # Backward-compatible fallback for config-based tokens
    expected = cfg.get("devices", {}).get(device_id)
    if not expected:
        raise HTTPException(status_code=403, detail="Unknown device")
    if token != expected:
        raise HTTPException(status_code=401, detail="Invalid token")


def upsert_heartbeat(payload: Heartbeat) -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO devices(device_id, public_ip, hostname, agent_version, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(device_id) DO UPDATE SET
                public_ip=excluded.public_ip,
                hostname=excluded.hostname,
                agent_version=excluded.agent_version,
                updated_at=excluded.updated_at
            """,
            (
                payload.device_id,
                payload.public_ip,
                payload.hostname,
                payload.agent_version,
                utc_now(),
            ),
        )


def parse_basic_auth_header(auth_header: str | None) -> tuple[str, str] | None:
    if not auth_header:
        return None

    prefix = "Basic "
    if not auth_header.startswith(prefix):
        return None

    encoded = auth_header[len(prefix):].strip()
    if not encoded:
        return None

    try:
        decoded = base64.b64decode(encoded).decode("utf-8")
    except (binascii.Error, UnicodeDecodeError):
        return None

    if ":" not in decoded:
        return None

    username, password = decoded.split(":", 1)
    return username, password


def expected_admin_credentials(cfg: dict) -> tuple[str, str]:
    expected_user = str(cfg.get("admin_username", "admin")).strip() or "admin"
    expected_password = str(cfg.get("admin_password", "")).strip()
    if not expected_password:
        expected_password = str(cfg.get("admin_api_key", "")).strip()
    return expected_user, expected_password


def session_signing_secret(cfg: dict) -> str:
    custom = str(cfg.get("admin_session_secret", "")).strip()
    if custom:
        return custom

    _, password = expected_admin_credentials(cfg)
    if password:
        return password

    return str(cfg.get("admin_api_key", "")).strip()


def should_use_secure_admin_cookie(request: Request, cfg: dict) -> bool:
    mode_raw = cfg.get("admin_session_cookie_secure", "auto")
    if isinstance(mode_raw, bool):
        return mode_raw

    mode = str(mode_raw).strip().lower()
    if mode in {"1", "true", "yes", "on"}:
        return True
    if mode in {"0", "false", "no", "off"}:
        return False

    forwarded_proto = (request.headers.get("x-forwarded-proto") or "").split(",", 1)[0].strip().lower()
    if forwarded_proto:
        return forwarded_proto == "https"

    return request.url.scheme == "https"


def create_admin_session(username: str, cfg: dict) -> str:
    payload = {
        "u": username,
        "exp": int(time.time()) + ADMIN_SESSION_TTL_SECONDS,
    }
    payload_json = json.dumps(payload, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    payload_b64 = base64.urlsafe_b64encode(payload_json).decode("ascii").rstrip("=")

    secret = session_signing_secret(cfg).encode("utf-8")
    signature = hmac.new(secret, payload_b64.encode("utf-8"), hashlib.sha256).digest()
    sig_b64 = base64.urlsafe_b64encode(signature).decode("ascii").rstrip("=")
    return f"{payload_b64}.{sig_b64}"


def verify_admin_session(token: str | None, cfg: dict) -> str | None:
    if not token or "." not in token:
        return None

    payload_b64, sig_b64 = token.split(".", 1)
    if not payload_b64 or not sig_b64:
        return None

    secret = session_signing_secret(cfg).encode("utf-8")
    expected_sig = hmac.new(secret, payload_b64.encode("utf-8"), hashlib.sha256).digest()
    expected_sig_b64 = base64.urlsafe_b64encode(expected_sig).decode("ascii").rstrip("=")
    if not secrets.compare_digest(sig_b64, expected_sig_b64):
        return None

    padding = "=" * (-len(payload_b64) % 4)
    try:
        payload_raw = base64.urlsafe_b64decode(payload_b64 + padding).decode("utf-8")
        payload = json.loads(payload_raw)
    except (ValueError, binascii.Error):
        return None

    username = str(payload.get("u", "")).strip()
    expires_at = int(payload.get("exp", 0))
    if not username or expires_at <= int(time.time()):
        return None

    return username


def authenticated_admin_user(request: Request, cfg: dict, x_admin_key: str | None = None) -> str | None:
    expected_key = str(cfg.get("admin_api_key", "")).strip()
    if expected_key and x_admin_key and secrets.compare_digest(x_admin_key, expected_key):
        return str(cfg.get("admin_username", "admin")).strip() or "admin"

    session_user = verify_admin_session(request.cookies.get(ADMIN_SESSION_COOKIE), cfg)
    if session_user:
        return session_user

    creds = parse_basic_auth_header(request.headers.get("authorization"))
    expected_user, expected_password = expected_admin_credentials(cfg)
    if expected_password and creds:
        username, password = creds
        if secrets.compare_digest(username, expected_user) and secrets.compare_digest(password, expected_password):
            return username

    return None


def require_admin_access(
    request: Request,
    cfg: dict,
    x_admin_key: str | None = None,
    include_www_auth: bool = False,
) -> str:
    user = authenticated_admin_user(request, cfg, x_admin_key)
    if user:
        return user

    _, expected_password = expected_admin_credentials(cfg)
    if not expected_password:
        raise HTTPException(status_code=500, detail="Server admin credentials not configured")

    headers = {"WWW-Authenticate": "Basic realm=\"Structura Connector Admin\""} if include_www_auth else None
    raise HTTPException(
        status_code=401,
        detail="Invalid admin credentials",
        headers=headers,
    )


def safe_admin_next_url(next_url: str | None) -> str:
    raw = (next_url or "").strip()
    if not raw:
        return "/admin/ui"
    if not raw.startswith("/"):
        return "/admin/ui"
    if raw.startswith("//"):
        return "/admin/ui"
    return raw


def render_admin_login_page(next_url: str, error: str | None = None) -> str:
    if not ADMIN_LOGIN_UI_PATH.exists():
        raise HTTPException(status_code=404, detail="Admin login UI not found")

    html = ADMIN_LOGIN_UI_PATH.read_text(encoding="utf-8")
    html = html.replace("__NEXT_URL__", quote(next_url, safe="/%?=&"))
    html = html.replace("__ERROR__", error or "")
    return html


def admin_actor_name(x_admin_actor: str | None) -> str:
    return (x_admin_actor or "admin").strip() or "admin"


def resolve_token_export_dir(cfg: dict) -> Path:
    configured_path = str(cfg.get("token_export_dir", "")).strip()
    if configured_path:
        return Path(configured_path)
    return DEFAULT_TOKEN_EXPORT_DIR


def sanitize_filename_part(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    cleaned = cleaned.strip("._-")
    return cleaned or "value"


def generate_smb_password(length: int = 20) -> str:
    lower = "abcdefghjkmnpqrstuvwxyz"
    upper = "ABCDEFGHJKLMNPQRSTUVWXYZ"
    digits = "23456789"
    specials = "!@#$%*_-"
    all_chars = lower + upper + digits + specials

    base_chars = [
        secrets.choice(lower),
        secrets.choice(upper),
        secrets.choice(digits),
        secrets.choice(specials),
    ]
    base_chars.extend(secrets.choice(all_chars) for _ in range(max(0, length - len(base_chars))))
    secrets.SystemRandom().shuffle(base_chars)
    return "".join(base_chars)


def build_smb_username(device_id: str, cfg: dict) -> str:
    prefix = str(cfg.get("smb_user_prefix", DEFAULT_SMB_USER_PREFIX)).strip() or DEFAULT_SMB_USER_PREFIX
    safe = sanitize_filename_part(device_id).lower().replace(".", "_").replace("-", "_")
    safe = re.sub(r"[^a-z0-9_]", "_", safe)
    username = (prefix + safe)[:20].strip("_")
    return username or "bim_user"


def provision_smb_access(device_id: str, cfg: dict) -> dict:
    server_host = str(cfg.get("smb_server_host", DEFAULT_SMB_SERVER_HOST)).strip() or DEFAULT_SMB_SERVER_HOST
    share_name = str(cfg.get("smb_share_name", DEFAULT_SMB_SHARE_NAME)).strip() or DEFAULT_SMB_SHARE_NAME
    share_path = str(cfg.get("smb_share_path", DEFAULT_SMB_SHARE_PATH)).strip() or DEFAULT_SMB_SHARE_PATH
    username = build_smb_username(device_id, cfg)
    password = generate_smb_password()
    login = f"{server_host}\\{username}"

    cmd = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(SMB_SCRIPT),
        "-UserName",
        username,
        "-Password",
        password,
        "-ShareName",
        share_name,
        "-SharePath",
        share_path,
    ]
    run = subprocess.run(cmd, capture_output=True, text=True)
    if run.returncode != 0:
        details = (run.stderr or run.stdout).strip()
        raise RuntimeError(f"SMB user provisioning failed: {details}")

    return {
        "login": login,
        "username": username,
        "password": password,
        "share_unc": f"\\\\{server_host}\\{share_name}",
        "share_path": share_path,
    }


def save_device_access(device_id: str, smb_access: dict) -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO device_access(device_id, smb_login, smb_username, smb_password, smb_share_unc, smb_share_path, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(device_id) DO UPDATE SET
                smb_login=excluded.smb_login,
                smb_username=excluded.smb_username,
                smb_password=excluded.smb_password,
                smb_share_unc=excluded.smb_share_unc,
                smb_share_path=excluded.smb_share_path,
                updated_at=excluded.updated_at
            """,
            (
                device_id,
                str(smb_access.get("login", "")),
                str(smb_access.get("username", "")),
                str(smb_access.get("password", "")),
                str(smb_access.get("share_unc", "")),
                str(smb_access.get("share_path", "")),
                utc_now(),
            ),
        )


def get_device_access(device_id: str) -> dict | None:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            """
            SELECT smb_login, smb_username, smb_password, smb_share_unc, smb_share_path
            FROM device_access
            WHERE device_id = ?
            """,
            (device_id,),
        ).fetchone()

    if not row:
        return None

    return {
        "login": row[0],
        "username": row[1],
        "password": row[2],
        "share_unc": row[3],
        "share_path": row[4],
    }


def get_or_create_device_access(device_id: str, cfg: dict, force_rotate: bool = False) -> dict:
    if not force_rotate:
        existing = get_device_access(device_id)
        if existing:
            return existing

    smb_access = provision_smb_access(device_id, cfg)
    save_device_access(device_id, smb_access)
    return smb_access


def create_device_session(device_id: str, hostname: str | None, public_ip: str) -> str:
    session_id = secrets.token_urlsafe(24)
    now = utc_now()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO device_sessions(device_id, session_id, hostname, public_ip, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(device_id) DO UPDATE SET
                session_id=excluded.session_id,
                hostname=excluded.hostname,
                public_ip=excluded.public_ip,
                updated_at=excluded.updated_at
            """,
            (device_id, session_id, hostname, public_ip, now, now),
        )
    return session_id


def get_device_session(device_id: str) -> dict | None:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            """
            SELECT session_id, hostname, public_ip, created_at, updated_at
            FROM device_sessions
            WHERE device_id = ?
            """,
            (device_id,),
        ).fetchone()

    if not row:
        return None

    return {
        "session_id": row[0],
        "hostname": row[1],
        "public_ip": row[2],
        "created_at": row[3],
        "updated_at": row[4],
    }


def touch_device_session(device_id: str, hostname: str | None, public_ip: str) -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            UPDATE device_sessions
            SET hostname = ?, public_ip = ?, updated_at = ?
            WHERE device_id = ?
            """,
            (hostname, public_ip, utc_now(), device_id),
        )


def delete_device_session(device_id: str) -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("DELETE FROM device_sessions WHERE device_id = ?", (device_id,))


def ensure_current_device_session(device_id: str, session_id: str | None) -> None:
    current = get_device_session(device_id)
    if not current:
        return

    provided = (session_id or "").strip()
    if not provided:
        raise HTTPException(status_code=409, detail="Session required. Reconnect by token.")

    expected = str(current.get("session_id", "")).strip()
    if not expected or not secrets.compare_digest(provided, expected):
        raise HTTPException(status_code=409, detail="Session superseded by newer login")


def resolve_device_by_token(token: str, cfg: dict) -> tuple[str, str | None, bool]:
    token_hash = hash_token(token)
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            """
            SELECT device_id, issued_to FROM device_tokens
            WHERE token_hash = ? AND revoked_at IS NULL
            """,
            (token_hash,),
        ).fetchone()
        if row:
            conn.execute("UPDATE device_tokens SET last_used_at = ? WHERE device_id = ?", (utc_now(), row[0]))
            return row[0], row[1], True

    for device_id, configured_token in (cfg.get("devices", {}) or {}).items():
        if configured_token == token:
            return str(device_id), None, False

    raise HTTPException(status_code=401, detail="Invalid token")


def normalize_ip(value: str) -> str:
    ip_obj = ipaddress.ip_address(value.strip())
    return str(ip_obj)


def parse_first_forwarded_ip(value: str | None) -> str | None:
    if not value:
        return None
    first = value.split(",", 1)[0].strip()
    if not first:
        return None
    try:
        return normalize_ip(first)
    except ValueError:
        return None


def resolve_client_ip(request: Request, payload_ip: str | None = None) -> str:
    candidates = [
        request.headers.get("cf-connecting-ip"),
        parse_first_forwarded_ip(request.headers.get("x-forwarded-for")),
        request.headers.get("x-real-ip"),
        payload_ip,
        request.client.host if request.client else None,
    ]

    for candidate in candidates:
        if not candidate:
            continue
        try:
            return normalize_ip(str(candidate))
        except ValueError:
            continue

    raise HTTPException(status_code=400, detail="Cannot resolve client IP")


def run_powershell(command: str) -> str:
    run = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
        capture_output=True,
        text=True,
    )
    if run.returncode != 0:
        raise RuntimeError((run.stderr or run.stdout).strip() or "PowerShell command failed")
    return run.stdout.strip()


def upsert_static_ip_firewall_rule(ip: str, port: int) -> None:
    safe_ip = normalize_ip(ip)
    safe_port = int(port)
    if safe_port < 1 or safe_port > 65535:
        raise ValueError("Port must be in 1..65535")

    rule_name = f"Connector allow {safe_ip} {safe_port}"
    command = (
        f"$name='{rule_name}'; "
        f"$ip='{safe_ip}'; "
        f"$port={safe_port}; "
        "if (Get-NetFirewallRule -DisplayName $name -ErrorAction SilentlyContinue) { "
        "Set-NetFirewallRule -DisplayName $name -Enabled True -Direction Inbound -Action Allow | Out-Null; "
        "Set-NetFirewallRule -DisplayName $name -RemoteAddress $ip | Out-Null "
        "} else { "
        "New-NetFirewallRule -DisplayName $name -Direction Inbound -Action Allow -Protocol TCP -LocalPort $port -RemoteAddress $ip | Out-Null "
        "}"
    )
    run_powershell(command)


def remove_static_ip_firewall_rule(ip: str, port: int) -> None:
    safe_ip = normalize_ip(ip)
    safe_port = int(port)
    rule_name = f"Connector allow {safe_ip} {safe_port}"
    command = f"Remove-NetFirewallRule -DisplayName '{rule_name}' -ErrorAction SilentlyContinue | Out-Null"
    run_powershell(command)


def list_static_ip_firewall_rules() -> list[dict]:
    command = (
        "Get-NetFirewallRule -DisplayName 'Connector allow *' -ErrorAction SilentlyContinue "
        "| ForEach-Object { "
        "$addr=(Get-NetFirewallAddressFilter -AssociatedNetFirewallRule $_ | Select-Object -First 1 -ExpandProperty RemoteAddress); "
        "$port=(Get-NetFirewallPortFilter -AssociatedNetFirewallRule $_ | Select-Object -First 1 -ExpandProperty LocalPort); "
        "[pscustomobject]@{ display_name=$_.DisplayName; enabled=$_.Enabled; remote_address=$addr; local_port=$port } "
        "} | ConvertTo-Json -Compress"
    )
    output = run_powershell(command)
    if not output:
        return []
    parsed = json.loads(output)
    if isinstance(parsed, dict):
        return [parsed]
    return list(parsed)


def write_token_export_file(
    cfg: dict,
    event_type: str,
    device_id: str,
    issued_to: str | None,
    token: str,
    actor: str,
    smb_access: dict | None = None,
) -> dict:
    export_dir = resolve_token_export_dir(cfg)
    exported_at = utc_now()
    safe_device_id = sanitize_filename_part(device_id)
    safe_event = sanitize_filename_part(event_type)
    file_name = f"{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S_%f')}_{safe_device_id}_{safe_event}.txt"
    file_path = export_dir / file_name

    content = "\n".join(
        [
            f"event_type: {event_type}",
            f"created_at_utc: {exported_at}",
            f"device_id: {device_id}",
            f"issued_to: {issued_to or ''}",
            f"actor: {actor}",
            f"token: {token}",
            f"smb_login: {(smb_access or {}).get('login', '')}",
            f"smb_username: {(smb_access or {}).get('username', '')}",
            f"smb_password: {(smb_access or {}).get('password', '')}",
            f"smb_share_unc: {(smb_access or {}).get('share_unc', '')}",
            "",
        ]
    )

    try:
        export_dir.mkdir(parents=True, exist_ok=True)
        file_path.write_text(content, encoding="utf-8")
        return {
            "saved": True,
            "path": str(file_path),
            "error": None,
        }
    except OSError as exc:
        return {
            "saved": False,
            "path": str(file_path),
            "error": str(exc),
        }


def load_local_update_manifest() -> dict:
    if not UPDATE_MANIFEST_PATH.exists():
        raise HTTPException(status_code=404, detail="Update manifest not found")

    try:
        manifest = json.loads(UPDATE_MANIFEST_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=500, detail="Invalid update manifest JSON") from exc

    version = str(manifest.get("version", "")).strip()
    msi_url = str(manifest.get("msiUrl", "")).strip()
    if not version or not msi_url:
        raise HTTPException(status_code=500, detail="Update manifest must contain version and msiUrl")

    return {
        "version": version,
        "msiUrl": msi_url,
        "notes": str(manifest.get("notes", "")).strip(),
    }


def normalize_release_version(tag_name: str) -> str:
    version = tag_name.strip()
    if version.lower().startswith("v"):
        version = version[1:]
    return version


def parse_bool(value: object, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    normalized = str(value).strip().lower()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    return default


def parse_int(value: object, default: int) -> int:
    try:
        parsed = int(str(value).strip())
        if parsed < 0:
            return default
        return parsed
    except (TypeError, ValueError):
        return default


def github_manifest_cache_key(cfg: dict) -> str:
    repo = str(cfg.get("github_updates_repo", "")).strip().lower()
    asset = str(cfg.get("github_updates_asset_name", "Connector.Desktop.Setup.msi")).strip().lower()
    return f"{repo}|{asset}"


def get_cached_github_manifest(key: str) -> dict | None:
    global _github_manifest_cache_key, _github_manifest_cache_until, _github_manifest_cache_value
    if _github_manifest_cache_key != key:
        return None
    if _github_manifest_cache_until <= time.time():
        return None
    return _github_manifest_cache_value


def save_cached_github_manifest(key: str, manifest: dict, ttl_seconds: int) -> None:
    global _github_manifest_cache_key, _github_manifest_cache_until, _github_manifest_cache_value
    _github_manifest_cache_key = key
    _github_manifest_cache_value = manifest
    _github_manifest_cache_until = time.time() + max(0, ttl_seconds)


def load_github_update_manifest(cfg: dict) -> dict:
    repo = str(cfg.get("github_updates_repo", "")).strip()
    if not repo:
        raise HTTPException(status_code=500, detail="github_updates_repo is not configured")
    if "/" not in repo:
        raise HTTPException(status_code=500, detail="github_updates_repo must look like owner/repo")

    api_url = f"https://api.github.com/repos/{repo}/releases/latest"
    req = UrlRequest(api_url)
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("User-Agent", "StructuraConnectorServer")

    token = str(cfg.get("github_updates_token", "")).strip()
    if token:
        req.add_header("Authorization", f"Bearer {token}")

    try:
        with urlopen(req, timeout=20) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        message = f"GitHub API HTTP {exc.code}"
        if detail:
            message = f"{message}: {detail[:300]}"
        raise HTTPException(status_code=502, detail=message) from exc
    except URLError as exc:
        raise HTTPException(status_code=502, detail=f"GitHub API unavailable: {exc.reason}") from exc
    except ValueError as exc:
        raise HTTPException(status_code=502, detail="Invalid JSON from GitHub release API") from exc

    raw_tag = str(payload.get("tag_name", "")).strip()
    version = normalize_release_version(raw_tag)
    if not version:
        raise HTTPException(status_code=502, detail="GitHub release does not contain tag_name")

    configured_asset_name = str(cfg.get("github_updates_asset_name", "Connector.Desktop.Setup.msi")).strip()
    assets = payload.get("assets") or []
    selected_asset = None

    if configured_asset_name:
        for asset in assets:
            if str(asset.get("name", "")).strip().lower() == configured_asset_name.lower():
                selected_asset = asset
                break

    if selected_asset is None:
        for asset in assets:
            if str(asset.get("name", "")).strip().lower().endswith(".msi"):
                selected_asset = asset
                break

    if selected_asset is None:
        raise HTTPException(status_code=502, detail="GitHub release does not contain an MSI asset")

    msi_url = str(selected_asset.get("browser_download_url", "")).strip()
    if not msi_url:
        raise HTTPException(status_code=502, detail="GitHub release asset has no browser_download_url")

    return {
        "version": version,
        "msiUrl": msi_url,
        "notes": str(payload.get("body", "")).strip(),
    }


def load_update_manifest(cfg: dict | None = None) -> dict:
    active_cfg = cfg or load_config()
    repo = str(active_cfg.get("github_updates_repo", "")).strip()
    if not repo:
        return load_local_update_manifest()

    cache_ttl = parse_int(
        active_cfg.get("github_updates_cache_seconds", GITHUB_UPDATES_CACHE_TTL_SECONDS_DEFAULT),
        GITHUB_UPDATES_CACHE_TTL_SECONDS_DEFAULT,
    )
    cache_key = github_manifest_cache_key(active_cfg)
    cached = get_cached_github_manifest(cache_key)
    if cached:
        return cached

    fallback_to_local = parse_bool(active_cfg.get("github_updates_fallback_local_manifest", True), True)
    try:
        manifest = load_github_update_manifest(active_cfg)
        save_cached_github_manifest(cache_key, manifest, cache_ttl)
        return manifest
    except HTTPException:
        if fallback_to_local:
            return load_local_update_manifest()
        raise


def create_device_token(device_id: str, issued_to: str | None) -> str:
    raw_token = secrets.token_urlsafe(32)
    token_hash = hash_token(raw_token)
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO device_tokens(device_id, token_hash, issued_to, created_at, last_used_at, revoked_at)
            VALUES (?, ?, ?, ?, NULL, NULL)
            ON CONFLICT(device_id) DO UPDATE SET
                token_hash=excluded.token_hash,
                issued_to=excluded.issued_to,
                created_at=excluded.created_at,
                last_used_at=NULL,
                revoked_at=NULL
            """,
            (device_id, token_hash, issued_to, utc_now()),
        )
        conn.execute("DELETE FROM device_sessions WHERE device_id = ?", (device_id,))
    return raw_token


def revoke_device_token(device_id: str) -> bool:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT device_id FROM device_tokens WHERE device_id = ? AND revoked_at IS NULL",
            (device_id,),
        ).fetchone()
        if not row:
            return False
        conn.execute("UPDATE device_tokens SET revoked_at = ? WHERE device_id = ?", (utc_now(), device_id))
        conn.execute("DELETE FROM device_sessions WHERE device_id = ?", (device_id,))
        return True


def delete_device_token(device_id: str) -> bool:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute("SELECT device_id FROM device_tokens WHERE device_id = ?", (device_id,)).fetchone()
        if not row:
            return False
        conn.execute("DELETE FROM device_tokens WHERE device_id = ?", (device_id,))
        conn.execute("DELETE FROM device_access WHERE device_id = ?", (device_id,))
        conn.execute("DELETE FROM device_sessions WHERE device_id = ?", (device_id,))
        return True


def apply_firewall(device_id: str, ip: str, ports: list[int]) -> None:
    ports_csv = ",".join(str(p) for p in ports)
    cmd = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(FW_SCRIPT),
        "-DeviceId",
        device_id,
        "-RemoteIp",
        ip,
        "-PortsCsv",
        ports_csv,
    ]
    run = subprocess.run(cmd, capture_output=True, text=True)
    if run.returncode != 0:
        raise RuntimeError(f"Firewall update failed: {run.stderr.strip()}")


app = FastAPI(title="Connector API")


@app.on_event("startup")
def startup() -> None:
    init_db()
    load_config()


@app.get("/health")
def health() -> dict:
    return {"ok": True}


@app.get("/updates/latest.json")
def updates_latest_manifest() -> dict:
    return load_update_manifest()


@app.get("/updates/files/{file_name}")
def updates_file(file_name: str) -> FileResponse:
    safe_name = Path(file_name).name
    file_path = UPDATES_DIR / safe_name
    if not file_path.exists() or not file_path.is_file():
        raise HTTPException(status_code=404, detail="Update file not found")
    return FileResponse(path=str(file_path), filename=safe_name, media_type="application/octet-stream")


@app.post("/connect/bootstrap")
def connect_bootstrap(
    payload: BootstrapRequest,
    request: Request,
    x_device_token: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    token = (x_device_token or "").strip()
    if not token:
        raise HTTPException(status_code=401, detail="Missing token")

    device_id, issued_to, db_token = resolve_device_by_token(token, cfg)
    client_ip = resolve_client_ip(request, payload.public_ip)
    previous_session = get_device_session(device_id)
    session_id = create_device_session(device_id, payload.hostname, client_ip)
    heartbeat_payload = Heartbeat(
        device_id=device_id,
        public_ip=client_ip,
        hostname=payload.hostname,
        agent_version=payload.agent_version,
    )

    upsert_heartbeat(heartbeat_payload)
    apply_firewall(device_id, client_ip, cfg.get("managed_ports", [3389, 1238, 445]))

    try:
        smb_access = get_or_create_device_access(device_id, cfg, force_rotate=False)
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    add_audit(
        event_type="client_bootstrap",
        device_id=device_id,
        actor=device_id,
        details=(
            f"ip={client_ip}; token_source={'db' if db_token else 'config'}; "
            f"session_replaced={'yes' if previous_session else 'no'}"
        ),
    )

    return {
        "ok": True,
        "session_id": session_id,
        "device_id": device_id,
        "issued_to": issued_to,
        "public_ip": client_ip,
        "heartbeat_seconds": int(cfg.get("default_heartbeat_seconds", 60)),
        "update_manifest_url": str(cfg.get("update_manifest_url", "")).strip(),
        "smb_access": smb_access,
    }


@app.post("/heartbeat")
def heartbeat(
    payload: Heartbeat,
    request: Request,
    x_device_token: str | None = Header(default=None),
    x_device_session: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    check_token(payload.device_id, x_device_token, cfg)
    ensure_current_device_session(payload.device_id, x_device_session)
    resolved_ip = resolve_client_ip(request, payload.public_ip)
    payload.public_ip = resolved_ip
    upsert_heartbeat(payload)
    touch_device_session(payload.device_id, payload.hostname, resolved_ip)
    apply_firewall(payload.device_id, resolved_ip, cfg.get("managed_ports", [3389, 1238, 445]))
    return {"ok": True}


@app.post("/admin/tokens")
def admin_create_token(
    request: Request,
    payload: CreateTokenRequest,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_admin_access(request, cfg, x_admin_key)
    actor = admin_actor_name(x_admin_actor)
    if actor == "admin":
        actor = auth_user

    try:
        smb_access = get_or_create_device_access(payload.device_id, cfg, force_rotate=True)
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    token = create_device_token(payload.device_id, payload.issued_to)

    export_result = write_token_export_file(
        cfg=cfg,
        event_type="token_created",
        device_id=payload.device_id,
        issued_to=payload.issued_to,
        token=token,
        actor=actor,
        smb_access=smb_access,
    )
    add_audit(
        event_type="token_created",
        device_id=payload.device_id,
        actor=actor,
        details=(
            f"issued_to={payload.issued_to or ''}; "
            f"smb_user={smb_access['username']}; "
            f"token_export_saved={export_result['saved']}; "
            f"token_export_path={export_result['path']}"
        ),
    )
    return {
        "ok": True,
        "device_id": payload.device_id,
        "issued_to": payload.issued_to,
        "token": token,
        "smb_access": smb_access,
        "note": "Token is shown once. Store it securely.",
        "token_export": export_result,
    }


@app.post("/admin/tokens/{device_id}/rotate")
def admin_rotate_token(
    request: Request,
    device_id: str,
    payload: RotateTokenRequest,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_admin_access(request, cfg, x_admin_key)
    actor = admin_actor_name(x_admin_actor)
    if actor == "admin":
        actor = auth_user

    try:
        smb_access = get_or_create_device_access(device_id, cfg, force_rotate=True)
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    token = create_device_token(device_id, payload.issued_to)

    export_result = write_token_export_file(
        cfg=cfg,
        event_type="token_rotated",
        device_id=device_id,
        issued_to=payload.issued_to,
        token=token,
        actor=actor,
        smb_access=smb_access,
    )
    add_audit(
        event_type="token_rotated",
        device_id=device_id,
        actor=actor,
        details=(
            f"issued_to={payload.issued_to or ''}; "
            f"smb_user={smb_access['username']}; "
            f"token_export_saved={export_result['saved']}; "
            f"token_export_path={export_result['path']}"
        ),
    )
    return {
        "ok": True,
        "device_id": device_id,
        "issued_to": payload.issued_to,
        "token": token,
        "smb_access": smb_access,
        "note": "Token is shown once. Store it securely.",
        "token_export": export_result,
    }


@app.post("/admin/tokens/{device_id}/revoke")
def admin_revoke_token(
    request: Request,
    device_id: str,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_admin_access(request, cfg, x_admin_key)
    revoked = revoke_device_token(device_id)
    if revoked:
        add_audit(
            event_type="token_revoked",
            device_id=device_id,
            actor=admin_actor_name(x_admin_actor) if x_admin_actor else auth_user,
            details=None,
        )
    return {"ok": True, "device_id": device_id, "revoked": revoked}


@app.delete("/admin/tokens/{device_id}")
def admin_delete_token(
    request: Request,
    device_id: str,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_admin_access(request, cfg, x_admin_key)
    deleted = delete_device_token(device_id)
    if deleted:
        add_audit(
            event_type="token_deleted",
            device_id=device_id,
            actor=admin_actor_name(x_admin_actor) if x_admin_actor else auth_user,
            details=None,
        )
    return {"ok": True, "device_id": device_id, "deleted": deleted}


@app.get("/admin/tokens")
def admin_list_tokens(request: Request, x_admin_key: str | None = Header(default=None)) -> dict:
    cfg = load_config()
    require_admin_access(request, cfg, x_admin_key)
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            """
            SELECT device_id, issued_to, created_at, last_used_at, revoked_at
            FROM device_tokens
            ORDER BY created_at DESC
            """
        ).fetchall()
    return {
        "items": [
            {
                "device_id": r[0],
                "issued_to": r[1],
                "created_at": r[2],
                "last_used_at": r[3],
                "revoked_at": r[4],
            }
            for r in rows
        ]
    }


@app.get("/admin/devices")
def admin_devices(request: Request, x_admin_key: str | None = Header(default=None)) -> dict:
    cfg = load_config()
    require_admin_access(request, cfg, x_admin_key)
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            """
            SELECT d.device_id, d.public_ip, d.hostname, d.agent_version, d.updated_at,
                   t.issued_to, t.created_at, t.last_used_at, t.revoked_at
            FROM devices d
            LEFT JOIN device_tokens t ON t.device_id = d.device_id
            ORDER BY d.updated_at DESC
            """
        ).fetchall()
    return {
        "items": [
            {
                "device_id": r[0],
                "public_ip": r[1],
                "hostname": r[2],
                "agent_version": r[3],
                "updated_at": r[4],
                "issued_to": r[5],
                "token_created_at": r[6],
                "token_last_used_at": r[7],
                "token_revoked_at": r[8],
            }
            for r in rows
        ]
    }


@app.get("/admin/audit")
def admin_audit(request: Request, x_admin_key: str | None = Header(default=None), limit: int = 100) -> dict:
    cfg = load_config()
    require_admin_access(request, cfg, x_admin_key)
    safe_limit = max(1, min(limit, 500))
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            """
            SELECT id, event_type, device_id, actor, details, created_at
            FROM audit_events
            ORDER BY id DESC
            LIMIT ?
            """,
            (safe_limit,),
        ).fetchall()
    return {
        "items": [
            {
                "id": r[0],
                "event_type": r[1],
                "device_id": r[2],
                "actor": r[3],
                "details": r[4],
                "created_at": r[5],
            }
            for r in rows
        ]
    }


@app.get("/admin/network/rules")
def admin_network_rules(request: Request, x_admin_key: str | None = Header(default=None)) -> dict:
    cfg = load_config()
    require_admin_access(request, cfg, x_admin_key)
    try:
        items = list_static_ip_firewall_rules()
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    return {"items": items}


@app.post("/admin/network/allow-ip")
def admin_network_allow_ip(
    request: Request,
    payload: NetworkRuleRequest,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_admin_access(request, cfg, x_admin_key)
    ip = normalize_ip(payload.ip)
    ports = payload.ports or [8080, 445, 3389]
    unique_ports = sorted(set(int(p) for p in ports))

    try:
        for port in unique_ports:
            upsert_static_ip_firewall_rule(ip, port)
    except (RuntimeError, ValueError) as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    add_audit(
        event_type="network_allow_ip",
        device_id=None,
        actor=admin_actor_name(x_admin_actor) if x_admin_actor else auth_user,
        details=f"ip={ip}; ports={','.join(str(p) for p in unique_ports)}",
    )
    return {"ok": True, "ip": ip, "ports": unique_ports}


@app.post("/admin/network/revoke-ip")
def admin_network_revoke_ip(
    request: Request,
    payload: NetworkRuleRequest,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_admin_access(request, cfg, x_admin_key)
    ip = normalize_ip(payload.ip)
    ports = payload.ports or [8080, 445, 3389]
    unique_ports = sorted(set(int(p) for p in ports))

    try:
        for port in unique_ports:
            remove_static_ip_firewall_rule(ip, port)
    except (RuntimeError, ValueError) as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    add_audit(
        event_type="network_revoke_ip",
        device_id=None,
        actor=admin_actor_name(x_admin_actor) if x_admin_actor else auth_user,
        details=f"ip={ip}; ports={','.join(str(p) for p in unique_ports)}",
    )
    return {"ok": True, "ip": ip, "ports": unique_ports}


@app.get("/devices")
def devices() -> dict:
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT device_id, public_ip, hostname, agent_version, updated_at FROM devices ORDER BY updated_at DESC"
        ).fetchall()
    return {
        "items": [
            {
                "device_id": r[0],
                "public_ip": r[1],
                "hostname": r[2],
                "agent_version": r[3],
                "updated_at": r[4],
            }
            for r in rows
        ]
    }


@app.get("/admin/login", response_class=HTMLResponse)
def admin_login_page(request: Request, next: str | None = None) -> str:
    cfg = load_config()
    if authenticated_admin_user(request, cfg):
        return RedirectResponse(url=safe_admin_next_url(next), status_code=303)
    return render_admin_login_page(safe_admin_next_url(next))


@app.post("/admin/login")
async def admin_login_submit(request: Request):
    cfg = load_config()
    expected_user, expected_password = expected_admin_credentials(cfg)

    content_type = (request.headers.get("content-type") or "").lower()
    username = ""
    password = ""
    next_url = "/admin/ui"

    if "application/json" in content_type:
        try:
            data = await request.json()
        except ValueError:
            data = {}
        username = str(data.get("username", "")).strip()
        password = str(data.get("password", ""))
        next_url = str(data.get("next", "/admin/ui"))
    else:
        raw = (await request.body()).decode("utf-8", errors="ignore")
        parsed = parse_qs(raw)
        username = (parsed.get("username") or [""])[0].strip()
        password = (parsed.get("password") or [""])[0]
        next_url = (parsed.get("next") or ["/admin/ui"])[0]

    safe_next = safe_admin_next_url(next_url)

    if not expected_password:
        raise HTTPException(status_code=500, detail="Server admin credentials not configured")

    if not (
        secrets.compare_digest(username, expected_user)
        and secrets.compare_digest(password, expected_password)
    ):
        return HTMLResponse(
            content=render_admin_login_page(safe_next, "Неверный логин или пароль."),
            status_code=401,
        )

    response = RedirectResponse(url=safe_next, status_code=303)
    use_secure_cookie = should_use_secure_admin_cookie(request, cfg)
    response.set_cookie(
        key=ADMIN_SESSION_COOKIE,
        value=create_admin_session(expected_user, cfg),
        max_age=ADMIN_SESSION_TTL_SECONDS,
        httponly=True,
        samesite="lax",
        secure=use_secure_cookie,
        path="/",
    )
    return response


@app.post("/admin/logout")
def admin_logout(request: Request) -> RedirectResponse:
    cfg = load_config()
    response = RedirectResponse(url="/admin/login", status_code=303)
    response.delete_cookie(
        key=ADMIN_SESSION_COOKIE,
        path="/",
        secure=should_use_secure_admin_cookie(request, cfg),
        httponly=True,
        samesite="lax",
    )
    return response


@app.get("/admin/ui", response_class=HTMLResponse)
def admin_ui(request: Request, x_admin_key: str | None = Header(default=None)) -> str:
    cfg = load_config()
    if not ADMIN_UI_PATH.exists():
        raise HTTPException(status_code=404, detail="Admin UI not found")

    user = authenticated_admin_user(request, cfg, x_admin_key)
    if not user:
        next_url = quote("/admin/ui", safe="/")
        return RedirectResponse(url=f"/admin/login?next={next_url}", status_code=303)

    if not request.cookies.get(ADMIN_SESSION_COOKIE):
        response = HTMLResponse(content=ADMIN_UI_PATH.read_text(encoding="utf-8"))
        use_secure_cookie = should_use_secure_admin_cookie(request, cfg)
        response.set_cookie(
            key=ADMIN_SESSION_COOKIE,
            value=create_admin_session(user, cfg),
            max_age=ADMIN_SESSION_TTL_SECONDS,
            httponly=True,
            samesite="lax",
            secure=use_secure_cookie,
            path="/",
        )
        return response

    return ADMIN_UI_PATH.read_text(encoding="utf-8")


@app.get("/ops/ui", response_class=HTMLResponse)
def ops_ui() -> str:
    if not OPS_UI_PATH.exists():
        raise HTTPException(status_code=404, detail="Ops UI not found")
    return OPS_UI_PATH.read_text(encoding="utf-8")
