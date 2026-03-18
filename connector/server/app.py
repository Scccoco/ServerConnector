import json
import secrets
import hashlib
import ipaddress
import base64
import binascii
import hmac
import time
import re
import shutil
import sqlite3
import subprocess
import threading
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote
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
TEKLA_FIRM_MANIFEST_PATH = UPDATES_DIR / "tekla_firm_latest.json"
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
_tekla_publish_lock = threading.Lock()


class Heartbeat(BaseModel):
    device_id: str
    public_ip: str | None = None
    hostname: str | None = None
    agent_version: str | None = None
    tekla_installed_version: str | None = None
    tekla_target_version: str | None = None
    tekla_installed_revision: str | None = None
    tekla_target_revision: str | None = None
    tekla_pending_after_close: bool | None = None
    tekla_running: bool | None = None
    tekla_last_check_utc: str | None = None
    tekla_last_success_utc: str | None = None
    tekla_last_error: str | None = None


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


class TeklaManifestUpdateRequest(BaseModel):
    version: str
    revision: str
    published_at: str
    target_path: str
    minimum_connector_version: str
    repo_url: str
    repo_ref: str
    notes: str | None = None


class TeklaManifestPublishRequest(BaseModel):
    source_path: str
    comment: str


class AdminRoleUpdateRequest(BaseModel):
    username: str


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
            CREATE TABLE IF NOT EXISTS tekla_client_state (
                device_id TEXT PRIMARY KEY,
                installed_version TEXT,
                target_version TEXT,
                installed_revision TEXT,
                target_revision TEXT,
                pending_after_close INTEGER,
                tekla_running INTEGER,
                last_check_utc TEXT,
                last_success_utc TEXT,
                last_error TEXT,
                updated_at TEXT NOT NULL
            )
            """
        )

        columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(tekla_client_state)").fetchall()
        }
        if "installed_version" not in columns:
            conn.execute("ALTER TABLE tekla_client_state ADD COLUMN installed_version TEXT")
        if "target_version" not in columns:
            conn.execute("ALTER TABLE tekla_client_state ADD COLUMN target_version TEXT")
        if "last_error" not in columns:
            conn.execute("ALTER TABLE tekla_client_state ADD COLUMN last_error TEXT")

        token_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(device_tokens)").fetchall()
        }
        if "token_value" not in token_columns:
            conn.execute("ALTER TABLE device_tokens ADD COLUMN token_value TEXT")
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
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS admin_user_roles (
                username TEXT PRIMARY KEY,
                is_system_admin INTEGER NOT NULL DEFAULT 0,
                is_firm_admin INTEGER NOT NULL DEFAULT 0,
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


FIRM_AUDIT_EVENT_TYPES = {
    "firm_admin_granted",
    "firm_admin_revoked",
    "tekla_manifest_updated",
    "tekla_client_state",
}


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


def upsert_tekla_client_state(payload: Heartbeat) -> None:
    has_any = any(
        [
            payload.tekla_installed_version is not None,
            payload.tekla_target_version is not None,
            payload.tekla_installed_revision is not None,
            payload.tekla_target_revision is not None,
            payload.tekla_pending_after_close is not None,
            payload.tekla_running is not None,
            payload.tekla_last_check_utc is not None,
            payload.tekla_last_success_utc is not None,
            payload.tekla_last_error is not None,
        ]
    )
    if not has_any:
        return

    installed_version = (payload.tekla_installed_version or "").strip()
    target_version = (payload.tekla_target_version or "").strip()
    installed_revision = (payload.tekla_installed_revision or "").strip()
    target_revision = (payload.tekla_target_revision or "").strip()
    pending_after_close = 1 if payload.tekla_pending_after_close else 0
    tekla_running = 1 if payload.tekla_running else 0
    last_check_utc = (payload.tekla_last_check_utc or "").strip()
    last_success_utc = (payload.tekla_last_success_utc or "").strip()
    last_error = (payload.tekla_last_error or "").strip()

    with sqlite3.connect(DB_PATH) as conn:
        prev = conn.execute(
            """
            SELECT installed_version, target_version, installed_revision, target_revision, pending_after_close, tekla_running,
                   last_check_utc, last_success_utc, last_error
            FROM tekla_client_state
            WHERE device_id = ?
            """,
            (payload.device_id,),
        ).fetchone()

        conn.execute(
            """
            INSERT INTO tekla_client_state(
                device_id,
                installed_version,
                target_version,
                installed_revision,
                target_revision,
                pending_after_close,
                tekla_running,
                last_check_utc,
                last_success_utc,
                last_error,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(device_id) DO UPDATE SET
                installed_version=excluded.installed_version,
                target_version=excluded.target_version,
                installed_revision=excluded.installed_revision,
                target_revision=excluded.target_revision,
                pending_after_close=excluded.pending_after_close,
                tekla_running=excluded.tekla_running,
                last_check_utc=excluded.last_check_utc,
                last_success_utc=excluded.last_success_utc,
                last_error=excluded.last_error,
                updated_at=excluded.updated_at
            """,
            (
                payload.device_id,
                installed_version,
                target_version,
                installed_revision,
                target_revision,
                pending_after_close,
                tekla_running,
                last_check_utc,
                last_success_utc,
                last_error,
                utc_now(),
            ),
        )

    current = (
        installed_version,
        target_version,
        installed_revision,
        target_revision,
        pending_after_close,
        tekla_running,
        last_check_utc,
        last_success_utc,
        last_error,
    )
    if prev != current:
        add_audit(
            event_type="tekla_client_state",
            device_id=payload.device_id,
            actor=payload.device_id,
            details=(
                f"installed_version={installed_version}; target_version={target_version}; "
                f"installed={installed_revision}; target={target_revision}; "
                f"pending={pending_after_close}; running={tekla_running}; "
                f"last_error={last_error}"
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
    raw = (x_admin_actor or "admin").strip() or "admin"
    try:
        return unquote(raw).strip() or "admin"
    except Exception:
        return raw


def normalize_admin_username(value: str) -> str:
    normalized = value.strip().lower()
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized


def get_admin_roles(username: str, cfg: dict) -> dict:
    user = normalize_admin_username(username)
    if not user:
        return {"is_system_admin": False, "is_firm_admin": False}

    default_admin = normalize_admin_username(str(cfg.get("admin_username", "admin")))
    system_admin = user == default_admin
    firm_admin = user == default_admin

    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            """
            SELECT is_system_admin, is_firm_admin
            FROM admin_user_roles
            WHERE username = ?
            """,
            (user,),
        ).fetchone()

    if row:
        system_admin = bool(row[0])
        firm_admin = bool(row[1])

    return {"is_system_admin": system_admin, "is_firm_admin": firm_admin}


def set_firm_admin_role(username: str, is_firm_admin: bool, cfg: dict) -> dict:
    user = normalize_admin_username(username)
    if not user:
        raise HTTPException(status_code=400, detail="username is required")

    default_admin = normalize_admin_username(str(cfg.get("admin_username", "admin")))
    default_system = user == default_admin

    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT is_system_admin FROM admin_user_roles WHERE username = ?",
            (user,),
        ).fetchone()

        current_system = bool(row[0]) if row else default_system
        now = utc_now()
        conn.execute(
            """
            INSERT INTO admin_user_roles(username, is_system_admin, is_firm_admin, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(username) DO UPDATE SET
                is_system_admin=excluded.is_system_admin,
                is_firm_admin=excluded.is_firm_admin,
                updated_at=excluded.updated_at
            """,
            (user, 1 if current_system else 0, 1 if is_firm_admin else 0, now, now),
        )

    roles = get_admin_roles(user, cfg)
    return {"username": user, **roles}


def require_system_admin_access(
    request: Request,
    cfg: dict,
    x_admin_key: str | None = None,
    include_www_auth: bool = False,
) -> str:
    user = require_admin_access(request, cfg, x_admin_key, include_www_auth)
    roles = get_admin_roles(user, cfg)
    if not roles["is_system_admin"]:
        raise HTTPException(status_code=403, detail="System admin role required")
    return user


def require_firm_admin_access(
    request: Request,
    cfg: dict,
    x_admin_key: str | None = None,
    include_www_auth: bool = False,
) -> str:
    user = require_admin_access(request, cfg, x_admin_key, include_www_auth)
    roles = get_admin_roles(user, cfg)
    if not roles["is_firm_admin"]:
        raise HTTPException(status_code=403, detail="Firm admin role required")
    return user


def ensure_device_firm_admin(device_id: str, cfg: dict) -> None:
    roles = get_admin_roles(device_id, cfg)
    if not roles["is_firm_admin"]:
        raise HTTPException(status_code=403, detail="Firm admin role required for this device")


def ensure_device_system_admin(device_id: str, cfg: dict) -> None:
    roles = get_admin_roles(device_id, cfg)
    if not roles["is_system_admin"]:
        raise HTTPException(status_code=403, detail="System admin role required for this device")


def restart_tekla_service_internal() -> dict:
    command = (
        "$svc = Get-Service | "
        "Where-Object { $_.DisplayName -match 'Tekla Structures Multiuser Server|Tekla.*Multiuser|Multi-user' -or $_.Name -match 'tekla|multiuser|multi' } | "
        "Sort-Object Name | Select-Object -First 1; "
        "if (-not $svc) { throw 'Tekla multiuser service not found' }; "
        "Restart-Service -Name $svc.Name -Force; "
        "Start-Sleep -Seconds 2; "
        "$after = Get-Service -Name $svc.Name; "
        "[pscustomobject]@{ service_name=$after.Name; display_name=$after.DisplayName; status=[string]$after.Status; start_type=[string]$after.StartType } | ConvertTo-Json -Compress"
    )
    output = run_powershell(command)
    return json.loads(output) if output else {}


def restart_revit_service_internal() -> dict:
    command = (
        "$was = Get-Service -Name 'WAS' -ErrorAction SilentlyContinue; "
        "$w3svc = Get-Service -Name 'W3SVC' -ErrorAction SilentlyContinue; "
        "if ($was -and $was.Status -ne 'Running') { Start-Service -Name 'WAS' }; "
        "if ($w3svc -and $w3svc.Status -ne 'Running') { Start-Service -Name 'W3SVC' }; "
        "$revitSvcs = Get-Service | Where-Object { $_.DisplayName -match 'Revit Server AutoSync' -or $_.Name -match '^Revit.*AutoSync' }; "
        "foreach ($svc in $revitSvcs) { Restart-Service -Name $svc.Name -Force }; "
        "Import-Module WebAdministration -ErrorAction SilentlyContinue | Out-Null; "
        "$pools = Get-ChildItem IIS:/AppPools -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^RevitServerAppPool' } | Select-Object -ExpandProperty Name; "
        "foreach ($pool in $pools) { Restart-WebAppPool -Name $pool }; "
        "Start-Sleep -Seconds 2; "
        "$svcOut = @(); foreach ($s in $revitSvcs) { $cur = Get-Service -Name $s.Name; $svcOut += [pscustomobject]@{ service_name=$cur.Name; display_name=$cur.DisplayName; status=[string]$cur.Status; start_type=[string]$cur.StartType } }; "
        "$poolOut = @(); foreach ($p in $pools) { $state = (& 'C:/Windows/System32/inetsrv/appcmd.exe' list apppool $p /text:state); $poolOut += [pscustomobject]@{ app_pool=$p; state=$state } }; "
        "[pscustomobject]@{ services=$svcOut; app_pools=$poolOut; was_status=([string](Get-Service -Name 'WAS' -ErrorAction SilentlyContinue).Status); w3svc_status=([string](Get-Service -Name 'W3SVC' -ErrorAction SilentlyContinue).Status) } | ConvertTo-Json -Compress -Depth 5"
    )
    output = run_powershell(command)
    return json.loads(output) if output else {}


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
    wrapped_command = (
        "[Console]::InputEncoding = [System.Text.Encoding]::UTF8; "
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
        "$OutputEncoding = [System.Text.Encoding]::UTF8; "
        + command
    )
    run = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", wrapped_command],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
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


def load_local_tekla_firm_manifest(cfg: dict | None = None) -> dict:
    active_cfg = cfg or load_config()
    if not TEKLA_FIRM_MANIFEST_PATH.exists():
        raise HTTPException(status_code=404, detail="Tekla firm manifest not found")

    try:
        manifest = json.loads(TEKLA_FIRM_MANIFEST_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=500, detail="Invalid Tekla firm manifest JSON") from exc

    version = str(manifest.get("version", "")).strip()
    revision = str(manifest.get("revision", "")).strip()
    published_at = str(manifest.get("published_at", "")).strip()
    target_path = str(manifest.get("target_path", "")).strip() or str(
        active_cfg.get("tekla_firm_target_path", "")
    ).strip()
    minimum_connector_version = str(manifest.get("minimum_connector_version", "")).strip()
    repo_url = str(manifest.get("repo_url", "")).strip()
    repo_ref = str(manifest.get("repo_ref", "")).strip()

    if not version:
        raise HTTPException(status_code=500, detail="Tekla firm manifest must contain version")
    if not revision:
        raise HTTPException(status_code=500, detail="Tekla firm manifest must contain revision")
    if not published_at:
        raise HTTPException(status_code=500, detail="Tekla firm manifest must contain published_at")
    if not target_path:
        raise HTTPException(
            status_code=500,
            detail="Tekla firm manifest must contain target_path or tekla_firm_target_path must be configured",
        )
    if not minimum_connector_version:
        raise HTTPException(status_code=500, detail="Tekla firm manifest must contain minimum_connector_version")
    if not repo_url:
        raise HTTPException(status_code=500, detail="Tekla firm manifest must contain repo_url")
    if not repo_ref:
        raise HTTPException(status_code=500, detail="Tekla firm manifest must contain repo_ref")

    return {
        "version": version,
        "revision": revision,
        "published_at": published_at,
        "target_path": target_path,
        "minimum_connector_version": minimum_connector_version,
        "repo_url": repo_url,
        "repo_ref": repo_ref,
        "notes": str(manifest.get("notes", "")).strip(),
    }


def save_local_tekla_firm_manifest(payload: TeklaManifestUpdateRequest) -> dict:
    manifest = {
        "version": payload.version.strip(),
        "revision": payload.revision.strip(),
        "published_at": payload.published_at.strip(),
        "target_path": payload.target_path.strip(),
        "minimum_connector_version": payload.minimum_connector_version.strip(),
        "repo_url": payload.repo_url.strip(),
        "repo_ref": payload.repo_ref.strip(),
        "notes": (payload.notes or "").strip(),
    }

    if not manifest["version"]:
        raise HTTPException(status_code=400, detail="version is required")
    if not manifest["revision"]:
        raise HTTPException(status_code=400, detail="revision is required")
    if not manifest["published_at"]:
        raise HTTPException(status_code=400, detail="published_at is required")
    if not manifest["target_path"]:
        raise HTTPException(status_code=400, detail="target_path is required")
    if not manifest["minimum_connector_version"]:
        raise HTTPException(status_code=400, detail="minimum_connector_version is required")
    if not manifest["repo_url"]:
        raise HTTPException(status_code=400, detail="repo_url is required")
    if not manifest["repo_ref"]:
        raise HTTPException(status_code=400, detail="repo_ref is required")

    UPDATES_DIR.mkdir(parents=True, exist_ok=True)
    TEKLA_FIRM_MANIFEST_PATH.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return manifest


def parse_tekla_next_version(current_version: str) -> str:
    raw = current_version.strip()
    if not raw:
        now = datetime.now(timezone.utc)
        return f"{now.year}.{now.month:02d}.1"

    match = re.match(r"^(.*?)(\d+)$", raw)
    if not match:
        return raw + ".1"

    prefix = match.group(1)
    counter = int(match.group(2)) + 1
    return f"{prefix}{counter}"


def run_git_command(
    git_executable: str,
    repo_worktree: Path,
    args: list[str],
    timeout_seconds: int = 120,
) -> str:
    cmd = [git_executable, *args]
    try:
        run = subprocess.run(
            cmd,
            cwd=str(repo_worktree),
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            f"{' '.join(args)}: command timed out after {timeout_seconds} seconds"
        ) from exc
    output = (run.stdout or "").strip()
    error_output = (run.stderr or "").strip()
    if run.returncode != 0:
        details = error_output or output or "git command failed"
        raise RuntimeError(f"{' '.join(args)}: {details}")
    return output


def replace_directory_contents(source_dir: Path, destination_dir: Path) -> None:
    destination_dir.mkdir(parents=True, exist_ok=True)

    for child in destination_dir.iterdir():
        if child.name == ".git":
            continue
        if child.is_dir():
            shutil.rmtree(child)
        else:
            child.unlink(missing_ok=True)

    for child in source_dir.iterdir():
        target = destination_dir / child.name
        if child.is_dir():
            shutil.copytree(child, target, dirs_exist_ok=True)
        else:
            shutil.copy2(child, target)


def resolve_tekla_publish_settings(cfg: dict, current_manifest: dict) -> dict:
    repo_worktree = Path(str(cfg.get("tekla_firm_repo_worktree", "")).strip())
    if not str(repo_worktree).strip():
        raise HTTPException(status_code=500, detail="tekla_firm_repo_worktree is not configured")
    if not repo_worktree.exists() or not repo_worktree.is_dir():
        raise HTTPException(status_code=500, detail="tekla_firm_repo_worktree does not exist")
    if not (repo_worktree / ".git").exists():
        raise HTTPException(status_code=500, detail="tekla_firm_repo_worktree is not a git repository")

    repo_subdir = str(cfg.get("tekla_firm_repo_subdir", "")).strip()
    if not repo_subdir:
        raise HTTPException(status_code=500, detail="tekla_firm_repo_subdir is not configured")

    git_executable = str(cfg.get("tekla_firm_git_executable", "git")).strip() or "git"
    repo_ref = str(cfg.get("tekla_firm_repo_ref", "")).strip() or str(current_manifest.get("repo_ref", "")).strip() or "main"
    repo_url = str(cfg.get("tekla_firm_repo_url", "")).strip() or str(current_manifest.get("repo_url", "")).strip()
    target_path = str(cfg.get("tekla_firm_target_path", "")).strip() or str(current_manifest.get("target_path", "")).strip()
    minimum_connector_version = str(cfg.get("tekla_firm_minimum_connector_version", "")).strip() or str(
        current_manifest.get("minimum_connector_version", "")
    ).strip() or "1.0.0"

    if not repo_url:
        raise HTTPException(status_code=500, detail="tekla_firm_repo_url is not configured")
    if not target_path:
        raise HTTPException(status_code=500, detail="tekla_firm_target_path is not configured")

    return {
        "repo_worktree": repo_worktree,
        "repo_subdir": repo_subdir,
        "git_executable": git_executable,
        "repo_ref": repo_ref,
        "repo_url": repo_url,
        "target_path": target_path,
        "minimum_connector_version": minimum_connector_version,
    }


def collect_oversized_files(root_dir: Path, max_file_bytes: int, max_results: int = 5) -> list[tuple[str, int]]:
    found: list[tuple[str, int]] = []
    for path in root_dir.rglob("*"):
        if not path.is_file():
            continue
        try:
            size_bytes = path.stat().st_size
        except OSError:
            continue
        if size_bytes > max_file_bytes:
            relative = str(path.relative_to(root_dir)).replace("\\", "/")
            found.append((relative, size_bytes))
            if len(found) >= max_results:
                break
    return found


def format_size_mb(size_bytes: int) -> str:
    return f"{size_bytes / (1024 * 1024):.2f} MB"


def publish_tekla_firm_from_source(source_path: str, comment: str, cfg: dict, actor: str) -> dict:
    source_dir = Path(source_path.strip())
    if not source_dir.exists() or not source_dir.is_dir():
        raise HTTPException(status_code=400, detail="source_path must point to an existing directory")
    if not comment.strip():
        raise HTTPException(status_code=400, detail="comment is required")

    current_manifest = load_local_tekla_firm_manifest(cfg)
    settings = resolve_tekla_publish_settings(cfg, current_manifest)

    repo_worktree: Path = settings["repo_worktree"]
    repo_subdir = settings["repo_subdir"]
    git_executable = settings["git_executable"]
    repo_ref = settings["repo_ref"]
    destination_dir = repo_worktree / repo_subdir
    max_file_bytes = parse_int(cfg.get("tekla_firm_max_file_bytes", 95 * 1024 * 1024), 95 * 1024 * 1024)
    fetch_timeout_seconds = parse_int(cfg.get("tekla_firm_git_fetch_timeout_seconds", 120), 120)
    push_timeout_seconds = parse_int(cfg.get("tekla_firm_git_push_timeout_seconds", 420), 420)

    oversized = collect_oversized_files(source_dir, max_file_bytes=max_file_bytes)
    if oversized:
        details = "; ".join(f"{relative} ({format_size_mb(size)})" for relative, size in oversized)
        add_audit(
            event_type="tekla_firm_publish_blocked_large_files",
            device_id=None,
            actor=actor,
            details=f"source_path={source_dir}; max={max_file_bytes}; files={details}",
        )
        raise HTTPException(
            status_code=400,
            detail=(
                "Source contains oversized files for GitHub publishing. "
                f"Limit: {format_size_mb(max_file_bytes)}. Files: {details}"
            ),
        )

    remote_ref_exists = True
    try:
        run_git_command(git_executable, repo_worktree, ["fetch", "origin", repo_ref], timeout_seconds=fetch_timeout_seconds)
    except RuntimeError as exc:
        error_text = str(exc).lower()
        if "couldn't find remote ref" in error_text or "couldn't find remote branch" in error_text:
            remote_ref_exists = False
        else:
            raise

    if remote_ref_exists:
        run_git_command(git_executable, repo_worktree, ["checkout", repo_ref])
        run_git_command(
            git_executable,
            repo_worktree,
            ["pull", "--rebase", "origin", repo_ref],
            timeout_seconds=fetch_timeout_seconds,
        )
    else:
        try:
            run_git_command(git_executable, repo_worktree, ["checkout", repo_ref])
        except RuntimeError:
            run_git_command(git_executable, repo_worktree, ["checkout", "--orphan", repo_ref])

    replace_directory_contents(source_dir, destination_dir)

    status = run_git_command(git_executable, repo_worktree, ["status", "--porcelain", "--", repo_subdir])
    if not status.strip():
        add_audit(
            event_type="tekla_firm_publish_skipped",
            device_id=None,
            actor=actor,
            details=f"source_path={source_dir}; reason=no_changes",
        )
        return {
            "ok": True,
            "no_changes": True,
            "message": "Изменения не обнаружены. Публикация не выполнялась.",
            "version": str(current_manifest.get("version", "")),
            "revision": str(current_manifest.get("revision", "")),
            "manifest": current_manifest,
        }

    next_version = parse_tekla_next_version(str(current_manifest.get("version", "")))
    git_author_name = str(cfg.get("tekla_firm_git_author_name", "")).strip() or actor or "Structura Connector"
    git_author_email = str(cfg.get("tekla_firm_git_author_email", "")).strip() or "connector@local"
    head_before_commit = run_git_command(git_executable, repo_worktree, ["rev-parse", "HEAD"]).strip()
    run_git_command(git_executable, repo_worktree, ["add", "--", repo_subdir])
    run_git_command(
        git_executable,
        repo_worktree,
        [
            "-c",
            f"user.name={git_author_name}",
            "-c",
            f"user.email={git_author_email}",
            "commit",
            "-m",
            f"tekla firm {next_version}: {comment.strip()}",
        ],
    )
    rollback_done = False
    try:
        if remote_ref_exists:
            run_git_command(git_executable, repo_worktree, ["push", "origin", repo_ref], timeout_seconds=push_timeout_seconds)
        else:
            run_git_command(
                git_executable,
                repo_worktree,
                ["push", "-u", "origin", repo_ref],
                timeout_seconds=push_timeout_seconds,
            )
    except Exception as exc:
        try:
            run_git_command(git_executable, repo_worktree, ["reset", "--hard", head_before_commit], timeout_seconds=45)
            rollback_done = True
        except Exception as rollback_exc:
            add_audit(
                event_type="tekla_firm_publish_rollback_failed",
                device_id=None,
                actor=actor,
                details=f"source_path={source_dir}; rollback_error={str(rollback_exc)}",
            )
        if rollback_done:
            add_audit(
                event_type="tekla_firm_publish_rolled_back",
                device_id=None,
                actor=actor,
                details=f"source_path={source_dir}; reason=push_failed",
            )
        error_text = str(exc)
        if "gh001" in error_text.lower() or "large files detected" in error_text.lower() or "exceeds github's file size limit" in error_text.lower():
            raise HTTPException(
                status_code=400,
                detail=(
                    "GitHub отклонил публикацию из-за слишком больших файлов в runtime-папке. "
                    "Уберите крупные файлы из 01_XS_FIRM или вынесите их в дистрибутивы/базу знаний. "
                    f"Детали: {error_text}"
                ),
            ) from exc
        if "timed out after" in error_text.lower():
            raise HTTPException(
                status_code=504,
                detail=(
                    "Публикация заняла слишком много времени на шаге git push. "
                    "Сервер откатил локальный commit, можно повторить попытку. "
                    f"Детали: {error_text}"
                ),
            ) from exc
        raise HTTPException(
            status_code=500,
            detail=(
                "Ошибка git push при публикации XS_FIRM. Локальный commit откатан, можно повторить попытку. "
                f"Детали: {error_text}"
            ),
        ) from exc
    revision = run_git_command(git_executable, repo_worktree, ["rev-parse", "--short", "HEAD"])

    update_payload = TeklaManifestUpdateRequest(
        version=next_version,
        revision=revision,
        published_at=utc_now(),
        target_path=settings["target_path"],
        minimum_connector_version=settings["minimum_connector_version"],
        repo_url=settings["repo_url"],
        repo_ref=repo_ref,
        notes=comment.strip(),
    )
    manifest = save_local_tekla_firm_manifest(update_payload)

    add_audit(
        event_type="tekla_firm_publish_succeeded",
        device_id=None,
        actor=actor,
        details=(
            f"version={manifest.get('version', '')}; "
            f"revision={manifest.get('revision', '')}; "
            f"source_path={source_dir}; repo_ref={repo_ref}"
        ),
    )
    add_audit(
        event_type="tekla_manifest_updated",
        device_id=None,
        actor=actor,
        details=f"version={manifest.get('version', '')}; revision={manifest.get('revision', '')}",
    )

    return {
        "ok": True,
        "no_changes": False,
        "message": "Публикация выполнена успешно.",
        "version": str(manifest.get("version", "")),
        "revision": str(manifest.get("revision", "")),
        "manifest": manifest,
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


def clear_cached_github_manifest() -> None:
    global _github_manifest_cache_key, _github_manifest_cache_until, _github_manifest_cache_value
    _github_manifest_cache_key = ""
    _github_manifest_cache_until = 0.0
    _github_manifest_cache_value = None


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
            INSERT INTO device_tokens(device_id, token_hash, token_value, issued_to, created_at, last_used_at, revoked_at)
            VALUES (?, ?, ?, ?, ?, NULL, NULL)
            ON CONFLICT(device_id) DO UPDATE SET
                token_hash=excluded.token_hash,
                token_value=excluded.token_value,
                issued_to=excluded.issued_to,
                created_at=excluded.created_at,
                last_used_at=NULL,
                revoked_at=NULL
            """,
            (device_id, token_hash, raw_token, issued_to, utc_now()),
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


def get_managed_ports(cfg: dict) -> list[int]:
    raw_ports = cfg.get("managed_ports", [3389, 1238, 445])
    if not isinstance(raw_ports, list):
        raise HTTPException(status_code=500, detail="managed_ports must be a list")

    ports: list[int] = []
    for raw in raw_ports:
        try:
            port = int(raw)
        except (TypeError, ValueError):
            continue
        if 1 <= port <= 65535:
            ports.append(port)

    unique_ports = sorted(set(ports))
    if not unique_ports:
        raise HTTPException(status_code=500, detail="managed_ports is empty or invalid")
    return unique_ports


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


@app.get("/updates/tekla/firm/latest.json")
def updates_tekla_firm_latest_manifest() -> dict:
    return load_local_tekla_firm_manifest()


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
    upsert_tekla_client_state(heartbeat_payload)
    apply_firewall(device_id, client_ip, get_managed_ports(cfg))

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

    roles = get_admin_roles(device_id, cfg)

    return {
        "ok": True,
        "session_id": session_id,
        "device_id": device_id,
        "issued_to": issued_to,
        "public_ip": client_ip,
        "heartbeat_seconds": int(cfg.get("default_heartbeat_seconds", 60)),
        "update_manifest_url": str(cfg.get("update_manifest_url", "")).strip(),
        "is_system_admin": roles["is_system_admin"],
        "is_firm_admin": roles["is_firm_admin"],
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
    upsert_tekla_client_state(payload)
    touch_device_session(payload.device_id, payload.hostname, resolved_ip)
    apply_firewall(payload.device_id, resolved_ip, get_managed_ports(cfg))
    return {"ok": True}


@app.post("/connect/tekla/manifest")
def connect_publish_tekla_manifest(
    payload: TeklaManifestPublishRequest,
    request: Request,
    x_device_token: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    token = (x_device_token or "").strip()
    if not token:
        raise HTTPException(status_code=401, detail="Missing token")

    device_id, _, _ = resolve_device_by_token(token, cfg)
    ensure_device_firm_admin(device_id, cfg)
    actor = admin_actor_name(x_admin_actor) if x_admin_actor else device_id

    if not _tekla_publish_lock.acquire(blocking=False):
        raise HTTPException(status_code=409, detail="Tekla publish is already in progress")

    add_audit(
        event_type="tekla_firm_publish_started",
        device_id=device_id,
        actor=actor,
        details=f"source_path={payload.source_path.strip()}",
    )
    try:
        return publish_tekla_firm_from_source(
            source_path=payload.source_path,
            comment=payload.comment,
            cfg=cfg,
            actor=actor,
        )
    except HTTPException as exc:
        add_audit(
            event_type="tekla_firm_publish_failed",
            device_id=device_id,
            actor=actor,
            details=f"source_path={payload.source_path.strip()}; error={exc.detail}",
        )
        raise
    except Exception as exc:
        add_audit(
            event_type="tekla_firm_publish_failed",
            device_id=device_id,
            actor=actor,
            details=f"source_path={payload.source_path.strip()}; error={str(exc)}",
        )
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    finally:
        _tekla_publish_lock.release()


@app.post("/connect/services/restart-tekla")
def connect_restart_tekla_service(
    x_device_token: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    token = (x_device_token or "").strip()
    if not token:
        raise HTTPException(status_code=401, detail="Missing token")

    device_id, _, _ = resolve_device_by_token(token, cfg)
    roles = get_admin_roles(device_id, cfg)
    if not (roles["is_system_admin"] or roles["is_firm_admin"]):
        raise HTTPException(status_code=403, detail="System or firm admin role required for this device")
    actor = admin_actor_name(x_admin_actor) if x_admin_actor else device_id
    result = restart_tekla_service_internal()
    add_audit(
        event_type="service_restart_tekla",
        device_id=device_id,
        actor=actor,
        details=f"service={result.get('service_name', '')}; status={result.get('status', '')}",
    )
    return {"ok": True, "result": result}


@app.post("/connect/services/restart-revit")
def connect_restart_revit_service(
    x_device_token: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    token = (x_device_token or "").strip()
    if not token:
        raise HTTPException(status_code=401, detail="Missing token")

    device_id, _, _ = resolve_device_by_token(token, cfg)
    ensure_device_system_admin(device_id, cfg)
    actor = admin_actor_name(x_admin_actor) if x_admin_actor else device_id
    result = restart_revit_service_internal()
    add_audit(
        event_type="service_restart_revit",
        device_id=device_id,
        actor=actor,
        details=(
            f"services={len(result.get('services', []))}; "
            f"app_pools={len(result.get('app_pools', []))}; "
            f"was={result.get('was_status', '')}; w3svc={result.get('w3svc_status', '')}"
        ),
    )
    return {"ok": True, "result": result}


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
            SELECT t.device_id,
                   t.token_value,
                   t.issued_to,
                   t.created_at,
                   t.last_used_at,
                   t.revoked_at,
                   a.smb_login,
                   a.smb_username,
                   a.smb_password,
                   a.smb_share_unc,
                   a.updated_at,
                   d.public_ip,
                   d.hostname,
                   d.agent_version,
                   d.updated_at,
                   r.is_system_admin,
                   r.is_firm_admin,
                   s.installed_version,
                   s.target_version,
                   s.installed_revision,
                   s.target_revision,
                   s.last_error,
                   s.updated_at
            FROM device_tokens t
            LEFT JOIN device_access a ON a.device_id = t.device_id
            LEFT JOIN devices d ON d.device_id = t.device_id
            LEFT JOIN admin_user_roles r ON r.username = t.device_id
            LEFT JOIN tekla_client_state s ON s.device_id = t.device_id
            ORDER BY t.created_at DESC
            """
        ).fetchall()
    items = []
    for r in rows:
        items.append(
            {
                "device_id": r[0],
                "token": r[1],
                "issued_to": r[2],
                "created_at": r[3],
                "last_used_at": r[4],
                "revoked_at": r[5],
                "smb_login": r[6],
                "smb_username": r[7],
                "smb_password": r[8],
                "smb_share_unc": r[9],
                "smb_updated_at": r[10],
                "public_ip": r[11],
                "hostname": r[12],
                "agent_version": r[13],
                "device_updated_at": r[14],
                "is_system_admin": bool(r[15]),
                "is_firm_admin": bool(r[16]),
                "tekla_installed_version": r[17],
                "tekla_target_version": r[18],
                "tekla_installed_revision": r[19],
                "tekla_target_revision": r[20],
                "tekla_last_error": r[21],
                "tekla_updated_at": r[22],
            }
        )
    return {"items": items}


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


@app.get("/admin/firm-admins")
def admin_list_firm_admins(request: Request, x_admin_key: str | None = Header(default=None)) -> dict:
    cfg = load_config()
    require_system_admin_access(request, cfg, x_admin_key)

    default_admin = normalize_admin_username(str(cfg.get("admin_username", "admin")))
    items = []

    if default_admin:
        default_roles = get_admin_roles(default_admin, cfg)
        if default_roles["is_firm_admin"]:
            items.append(
                {
                    "username": default_admin,
                    "is_system_admin": default_roles["is_system_admin"],
                    "is_firm_admin": default_roles["is_firm_admin"],
                    "is_default_admin": True,
                }
            )

    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            """
            SELECT username, is_system_admin, is_firm_admin, created_at, updated_at
            FROM admin_user_roles
            WHERE is_firm_admin = 1
            ORDER BY username ASC
            """
        ).fetchall()

    for row in rows:
        username = normalize_admin_username(str(row[0]))
        if not username:
            continue
        if username == default_admin:
            continue
        items.append(
            {
                "username": username,
                "is_system_admin": bool(row[1]),
                "is_firm_admin": bool(row[2]),
                "created_at": row[3],
                "updated_at": row[4],
                "is_default_admin": False,
            }
        )

    return {"items": items}


@app.post("/admin/firm-admins/grant")
def admin_grant_firm_admin(
    payload: AdminRoleUpdateRequest,
    request: Request,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_system_admin_access(request, cfg, x_admin_key)
    result = set_firm_admin_role(payload.username, True, cfg)
    actor = admin_actor_name(x_admin_actor) if x_admin_actor else auth_user
    add_audit(
        event_type="firm_admin_granted",
        device_id=None,
        actor=actor,
        details=f"username={result['username']}",
    )
    return {"ok": True, "item": result}


@app.post("/admin/firm-admins/revoke")
def admin_revoke_firm_admin(
    payload: AdminRoleUpdateRequest,
    request: Request,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_system_admin_access(request, cfg, x_admin_key)
    result = set_firm_admin_role(payload.username, False, cfg)
    actor = admin_actor_name(x_admin_actor) if x_admin_actor else auth_user
    add_audit(
        event_type="firm_admin_revoked",
        device_id=None,
        actor=actor,
        details=f"username={result['username']}",
    )
    return {"ok": True, "item": result}


@app.get("/admin/tekla/clients")
def admin_tekla_clients(request: Request, x_admin_key: str | None = Header(default=None)) -> dict:
    cfg = load_config()
    require_admin_access(request, cfg, x_admin_key)
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            """
            SELECT t.device_id,
                   t.installed_version,
                   t.target_version,
                   t.installed_revision,
                   t.target_revision,
                   t.pending_after_close,
                   t.tekla_running,
                   t.last_check_utc,
                   t.last_success_utc,
                   t.last_error,
                   t.updated_at,
                   d.hostname,
                   d.public_ip,
                   tok.issued_to
            FROM tekla_client_state t
            LEFT JOIN devices d ON d.device_id = t.device_id
            LEFT JOIN device_tokens tok ON tok.device_id = t.device_id
            ORDER BY t.updated_at DESC
            """
        ).fetchall()

    return {
        "items": [
            {
                "device_id": r[0],
                "installed_version": r[1],
                "target_version": r[2],
                "installed_revision": r[3],
                "target_revision": r[4],
                "pending_after_close": bool(r[5]),
                "tekla_running": bool(r[6]),
                "last_check_utc": r[7],
                "last_success_utc": r[8],
                "last_error": r[9],
                "updated_at": r[10],
                "hostname": r[11],
                "public_ip": r[12],
                "issued_to": r[13],
            }
            for r in rows
        ]
    }


@app.post("/admin/tekla/manifest")
def admin_save_tekla_manifest(
    payload: TeklaManifestUpdateRequest,
    request: Request,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_firm_admin_access(request, cfg, x_admin_key)
    manifest = save_local_tekla_firm_manifest(payload)
    actor = admin_actor_name(x_admin_actor) if x_admin_actor else auth_user
    add_audit(
        event_type="tekla_manifest_updated",
        device_id=None,
        actor=actor,
        details=f"version={manifest.get('version', '')}; revision={manifest.get('revision', '')}",
    )
    return {"ok": True, "manifest": manifest}


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


@app.get("/admin/audit/firm")
def admin_firm_audit(
    request: Request,
    x_admin_key: str | None = Header(default=None),
    limit: int = 100,
    include_state: bool = True,
) -> dict:
    cfg = load_config()
    require_system_admin_access(request, cfg, x_admin_key)
    safe_limit = max(1, min(limit, 500))

    event_types = [
        "firm_admin_granted",
        "firm_admin_revoked",
        "tekla_firm_publish_started",
        "tekla_firm_publish_succeeded",
        "tekla_firm_publish_failed",
        "tekla_firm_publish_skipped",
        "tekla_firm_publish_blocked_large_files",
        "tekla_firm_publish_rolled_back",
        "tekla_firm_publish_rollback_failed",
        "tekla_manifest_updated",
    ]
    if include_state:
        event_types.append("tekla_client_state")

    placeholders = ",".join("?" for _ in event_types)
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            f"""
            SELECT id, event_type, device_id, actor, details, created_at
            FROM audit_events
            WHERE event_type IN ({placeholders})
            ORDER BY id DESC
            LIMIT ?
            """,
            (*event_types, safe_limit),
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


@app.post("/admin/updates/refresh")
def admin_refresh_updates_manifest(
    request: Request,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_admin_access(request, cfg, x_admin_key)
    clear_cached_github_manifest()
    manifest = load_update_manifest(cfg)
    actor = admin_actor_name(x_admin_actor) if x_admin_actor else auth_user
    add_audit(
        event_type="updates_manifest_refresh",
        device_id=None,
        actor=actor,
        details=f"version={manifest.get('version', '')}",
    )
    return {"ok": True, "manifest": manifest}


@app.post("/admin/services/restart-tekla")
def admin_restart_tekla_service(
    request: Request,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_admin_access(request, cfg, x_admin_key)

    result = restart_tekla_service_internal()

    actor = admin_actor_name(x_admin_actor) if x_admin_actor else auth_user
    add_audit(
        event_type="service_restart_tekla",
        device_id=None,
        actor=actor,
        details=f"service={result.get('service_name', '')}; status={result.get('status', '')}",
    )
    return {"ok": True, "result": result}


@app.post("/admin/services/restart-revit")
def admin_restart_revit_service(
    request: Request,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_admin_access(request, cfg, x_admin_key)

    result = restart_revit_service_internal()

    actor = admin_actor_name(x_admin_actor) if x_admin_actor else auth_user
    add_audit(
        event_type="service_restart_revit",
        device_id=None,
        actor=actor,
        details=(
            f"services={len(result.get('services', []))}; "
            f"app_pools={len(result.get('app_pools', []))}; "
            f"was={result.get('was_status', '')}; w3svc={result.get('w3svc_status', '')}"
        ),
    )
    return {"ok": True, "result": result}


@app.get("/admin/network/rules")
def admin_network_rules(request: Request, x_admin_key: str | None = Header(default=None)) -> dict:
    cfg = load_config()
    require_admin_access(request, cfg, x_admin_key)
    try:
        items = list_static_ip_firewall_rules()
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    return {"items": items}


@app.get("/admin/network/managed-ports")
def admin_network_managed_ports(request: Request, x_admin_key: str | None = Header(default=None)) -> dict:
    cfg = load_config()
    require_admin_access(request, cfg, x_admin_key)
    return {"managed_ports": get_managed_ports(cfg)}


@app.post("/admin/network/reapply-managed-ports")
def admin_network_reapply_managed_ports(
    request: Request,
    x_admin_key: str | None = Header(default=None),
    x_admin_actor: str | None = Header(default=None),
) -> dict:
    cfg = load_config()
    auth_user = require_admin_access(request, cfg, x_admin_key)
    ports = get_managed_ports(cfg)

    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            """
            SELECT t.device_id,
                   t.issued_to,
                   COALESCE(d.public_ip, s.public_ip) AS effective_ip
            FROM device_tokens t
            LEFT JOIN devices d ON d.device_id = t.device_id
            LEFT JOIN device_sessions s ON s.device_id = t.device_id
            WHERE t.revoked_at IS NULL
            ORDER BY t.created_at DESC
            """
        ).fetchall()

    applied: list[dict] = []
    skipped: list[dict] = []
    failed: list[dict] = []

    for row in rows:
        device_id = str(row[0])
        issued_to = row[1]
        ip = (row[2] or "").strip()

        if not ip:
            skipped.append({"device_id": device_id, "issued_to": issued_to, "reason": "no_ip"})
            continue

        try:
            safe_ip = normalize_ip(ip)
            apply_firewall(device_id, safe_ip, ports)
            applied.append({"device_id": device_id, "issued_to": issued_to, "ip": safe_ip})
        except Exception as exc:
            failed.append({"device_id": device_id, "issued_to": issued_to, "ip": ip, "error": str(exc)})

    actor = admin_actor_name(x_admin_actor) if x_admin_actor else auth_user
    add_audit(
        event_type="network_reapply_managed_ports",
        device_id=None,
        actor=actor,
        details=(
            f"ports={','.join(str(p) for p in ports)}; "
            f"applied={len(applied)}; skipped={len(skipped)}; failed={len(failed)}"
        ),
    )

    return {
        "ok": len(failed) == 0,
        "managed_ports": ports,
        "total_active_tokens": len(rows),
        "applied_count": len(applied),
        "skipped_count": len(skipped),
        "failed_count": len(failed),
        "applied": applied,
        "skipped": skipped,
        "failed": failed,
    }


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
    ports = payload.ports or get_managed_ports(cfg)
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
    ports = payload.ports or get_managed_ports(cfg)
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
