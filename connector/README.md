# Connector MVP

Lightweight "connector" system to auto-manage firewall allowlist entries for changing client IPs.

## Components
- `server/`: API that accepts device heartbeats and applies firewall rules on the server.
- `agent/`: client-side process for user PCs; reports current public IP periodically.

## Security model (MVP)
- Each device has its own token.
- API rejects requests with unknown token.
- Firewall rules are scoped per device and per port.
- Admin API key protects token issuing endpoints.
- Tokens are stored as hashes in DB; plaintext token is shown once on create/rotate.
- `issued_to` is optional metadata so you can track who received the token.
- On create/rotate, plaintext token details are also written to txt files in `\\62.113.36.107\BIM_Models\Tokens` (or `token_export_dir` from config).
- On create/rotate, server also provisions or updates a local SMB user and grants edit access to `\\62.113.36.107\BIM_Models`; login/password are returned once and exported to txt.

## Ports managed
- `3389` (RDP)
- `1238` (Tekla)
- `445` (SMB)

Adjust in `server/config.json`.

## Quick start (server side)
1. Install Python 3.11+ on VPS.
2. In `connector/server/`:
   - `pip install -r requirements.txt`
   - copy `config.example.json` to `config.json` and set values (`admin_api_key` is required).
   - run: `uvicorn app:app --host 0.0.0.0 --port 8080`

## Admin token flow (with optional "issued to")
Create token for device:

```bash
curl -X POST http://127.0.0.1:8080/admin/tokens \
  -H "Content-Type: application/json" \
  -H "X-Admin-Key: YOUR_ADMIN_KEY" \
  -d '{"device_id":"user-laptop-1","issued_to":"Ivan Petrov"}'
```

- `issued_to` is optional and stored for tracking.
- The response includes plaintext token and SMB credentials once; store them securely.

Rotate token for existing device:

```bash
curl -X POST http://127.0.0.1:8080/admin/tokens/user-laptop-1/rotate \
  -H "Content-Type: application/json" \
  -H "X-Admin-Key: YOUR_ADMIN_KEY" \
  -H "X-Admin-Actor: Alexander" \
  -d '{"issued_to":"Ivan Petrov"}'
```

Revoke token:

```bash
curl -X POST http://127.0.0.1:8080/admin/tokens/user-laptop-1/revoke \
  -H "X-Admin-Key: YOUR_ADMIN_KEY" \
  -H "X-Admin-Actor: Alexander"
```

Delete token record:

```bash
curl -X DELETE http://127.0.0.1:8080/admin/tokens/user-laptop-1 \
  -H "X-Admin-Key: YOUR_ADMIN_KEY" \
  -H "X-Admin-Actor: Alexander"
```

List issued tokens metadata:

```bash
curl http://127.0.0.1:8080/admin/tokens -H "X-Admin-Key: YOUR_ADMIN_KEY"
```

List devices with token metadata:

```bash
curl http://127.0.0.1:8080/admin/devices -H "X-Admin-Key: YOUR_ADMIN_KEY"
```

List audit events:

```bash
curl "http://127.0.0.1:8080/admin/audit?limit=200" -H "X-Admin-Key: YOUR_ADMIN_KEY"
```

## Admin UI
- Open `http://SERVER:8080/admin/ui`
- If not authorized, you are redirected to `http://SERVER:8080/admin/login`
- Enter admin login/password on the login page
- Optionally set `Admin Name (for audit)`
- Use Create/Rotate/Revoke and list buttons
- In the token table, each token has `Перевыдать` and `Удалить` buttons

Admin credentials source:
- Recommended: `admin_username` + `admin_password` in `server/config.json`
- Backward-compatible fallback: if `admin_password` is empty, password falls back to `admin_api_key`
- Login creates an HttpOnly admin session cookie (`connector_admin_session`)

## Ops UI (read-only monitoring)
- Open `http://SERVER:8080/ops/ui`
- Enter `Admin API Key`
- Check healthy/stale/revoked device counters and device list

## Desktop app updates
- Update manifest endpoint: `http://SERVER:8080/updates/latest.json`
- Update file endpoint: `http://SERVER:8080/updates/files/<MSI_FILE_NAME>`
- Manifest file is read from `server/updates/latest.json`
- Required manifest fields: `version`, `msiUrl`, `sha256` (optional `notes`)

## Tekla firm lightweight manifest
- Endpoint: `GET /updates/tekla/firm/latest.json`
- Local source file: `server/updates/tekla_firm_latest.json`
- Response fields: `version`, `revision`, `published_at`, `target_path`, `minimum_connector_version`, `repo_url`, `repo_ref`, `notes`
- Validation behavior:
  - Returns `404` when `server/updates/tekla_firm_latest.json` is missing
  - Returns `500` when JSON is invalid or required fields are missing
  - `target_path` can come from manifest file or fallback config key `tekla_firm_target_path`
- Optional config keys in `server/config.json`:
  - `tekla_firm_manifest_url` (default `https://server.structura-most.ru/updates/tekla/firm/latest.json`)
  - `tekla_firm_target_path` (default `C:\Company\TeklaFirm`)
  - `tekla_firm_repo_worktree` (local git working tree on server)
  - `tekla_firm_repo_subdir` (subfolder inside repo to mirror from source path)
  - `tekla_firm_repo_url` / `tekla_firm_repo_ref` (values written to published manifest)
  - `tekla_firm_minimum_connector_version` (default for published manifest)
  - `tekla_firm_max_file_bytes` (default `99614720`, blocks oversized files before git push)
  - `tekla_firm_git_fetch_timeout_seconds` / `tekla_firm_git_push_timeout_seconds` (timeouts for fetch/pull and push)
  - `tekla_firm_git_executable` (default `git`)

GitHub release mode (recommended):
- Configure `github_updates_repo` in `server/config.json` as `owner/repo`
- Optional: set `github_updates_asset_name` (default `Connector.Desktop.Setup.msi`)
- Optional: set `github_updates_cache_seconds` (default `300`) to reduce GitHub API calls
- Server returns the latest GitHub release as manifest data (`tag_name` -> `version`, release MSI asset -> `msiUrl`, asset digest -> `sha256`)
- If GitHub API fails and `github_updates_fallback_local_manifest=true`, server falls back to local `server/updates/latest.json`

Release automation:
- Workflow `.github/workflows/release-connector-desktop.yml` builds MSI and publishes it to GitHub Releases on tag push `v*`
- Recommended flow: update app version -> commit -> push tag `vX.Y.Z`

## Token bootstrap (desktop one-click connect)
- Endpoint: `POST /connect/bootstrap`
- Header: `X-Device-Token: <TOKEN>`
- Body fields: `hostname`, `agent_version` (optional `public_ip`)
- Server resolves client IP, applies firewall for managed ports, and returns:
  - `session_id` (active login session for this token/device)
  - `device_id`
  - `smb_access` (`login`, `password`, `share_unc`)
  - `heartbeat_seconds`
  - `update_manifest_url`
  - `is_firm_admin`

Single active device session:
- Heartbeat must include `X-Device-Session`
- New bootstrap replaces previous session for that token/device
- Previous client receives `409` on heartbeat and must reconnect

Example:

```bash
curl -X POST http://127.0.0.1:8080/connect/bootstrap \
  -H "Content-Type: application/json" \
  -H "X-Device-Token: DEVICE_TOKEN" \
  -d '{"hostname":"WORKSTATION-01","agent_version":"desktop-1.0.0"}'
```

## Desktop firm-admin publish API
- Endpoint: `POST /connect/tekla/manifest`
- Header: `X-Device-Token: <TOKEN>`
- Access: device must be granted `admin_firm` role
- Purpose: server-side auto publish for firm admins (`sync source -> git add/commit/push -> manifest update`)
- Request body:

```json
{
  "source_path": "\\\\62.113.36.107\\BIM_Models\\Tekla\\XS_FIRM",
  "comment": "Обновлены шаблоны и макросы"
}
```

## Admin network whitelist API
- List rules: `GET /admin/network/rules`
- Get current managed ports from server config: `GET /admin/network/managed-ports`
- Reapply current managed ports to all active tokens (using last known device IP): `POST /admin/network/reapply-managed-ports`
- Add/update IP rules: `POST /admin/network/allow-ip`
- Remove IP rules: `POST /admin/network/revoke-ip`

Request payload:

```json
{
  "ip": "193.27.41.84",
  "ports": [445, 3389, 80, 443]
}
```

If `ports` is omitted in allow/revoke requests, server uses `managed_ports` from `server/config.json`.

## Admin updates API
- Force refresh update manifest cache now: `POST /admin/updates/refresh`

Example:

```bash
curl -X POST http://127.0.0.1:8080/admin/updates/refresh \
  -H "X-Admin-Key: YOUR_ADMIN_KEY" \
  -H "X-Admin-Actor: release-bot"
```

## Admin Tekla API
- List Tekla standard client state from heartbeat data: `GET /admin/tekla/clients`
- Update Tekla firm manifest from admin UI/API: `POST /admin/tekla/manifest`

## Admin firm role API
- List users with `admin_firm` role: `GET /admin/firm-admins`
- Grant `admin_firm` role: `POST /admin/firm-admins/grant`
- Revoke `admin_firm` role: `POST /admin/firm-admins/revoke`
- Firm operations audit slice: `GET /admin/audit/firm?limit=120&include_state=true`

## Quick start (agent side)
1. Install Python 3.11+ on client PC.
2. Run installer script in `connector/agent/` (PowerShell as current user):

```powershell
.\install_agent.ps1 -ServerUrl "http://62.113.36.107:8080" -DeviceId "user-laptop-1" -DeviceToken "<TOKEN>" -HeartbeatSeconds 60
```

If Python is missing on client:

```powershell
.\install_agent.ps1 -ServerUrl "http://62.113.36.107:8080" -DeviceId "user-laptop-1" -DeviceToken "<TOKEN>" -InstallPythonIfMissing
```

What installer does:
- copies agent to `C:\ProgramData\ConnectorAgent`
- stores token in DPAPI-encrypted file `device_token.sec`
- creates scheduled task `ConnectorAgent` with auto-start at logon

Uninstall:

```powershell
.\uninstall_agent.ps1
```

## EXE installer
- Build setup exe:

```powershell
.\build_setup_exe.ps1
```

- Output file: `connector/agent/ConnectorAgentSetup.exe`
- End-user flow: run `ConnectorAgentSetup.exe`, paste device token when prompted.
