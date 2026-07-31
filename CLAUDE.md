# Structura Connector — заметки для Claude

> Этот файл подгружается автоматически в каждую сессию Claude Code в этом проекте. Он не заменяет код, а фиксирует то, что не видно из git/кода: инфраструктуру, конвенции, подводные камни и текущие риски.

## 1. Что это за продукт

Windows-продукт для удалённой работы с BIM-инфраструктурой. GitHub: `PROBIM55/ServerConnector`.

> ⚠️ **Активная ветка — `master`, не `main`.** `main` на GitHub помечен как default, но содержит только Initial commit с голым README.md. Весь код, история релизов и теги (v1.0.1…v1.0.20) живут на `master`. После `git clone` сразу делать `git checkout master`.

Состоит из двух частей:

- **`connector/server`** — FastAPI/Python (3.11+, FastAPI, **PostgreSQL** через psycopg2, PowerShell для firewall/SMB). До 2026-05-09 был SQLite — мигрирован на PG, см. § 4b.
- **`connector-desktop/Connector.Desktop`** — WPF/.NET 8, MSI через WiX, релизы через GitHub Actions.

Бизнес-логика в одном абзаце: пользователь вводит токен в desktop Connector → сервер возвращает device session, SMB-доступ и update endpoints → клиент работает в фоне (heartbeat, SMB mount, авто-обновление). Для Tekla пользователи работают не с сетевым runtime, а с локальной копией `C:\Company\TeklaFirm`, которую Connector держит в актуальном состоянии.

## 2. Прод-инфраструктура

| Компонент | Адрес/путь |
|---|---|
| VPS host | `62.113.36.107` (Windows Server 2022, hostname `VDSWIN2K22`) |
| Домен | `server.structura-most.ru` |
| Reverse proxy | Caddy → backend `127.0.0.1:8080` |
| SMB share | `\\62.113.36.107\BIM_Models` ↔ `C:\BIM_Models` |
| Server source (git checkout) | `C:\Connector\src\` (master, перезаписывается каждым деплоем) |
| Server runtime (persistent) | `C:\Connector\runtime\` (`config.json`, `.env` с `CONNECTOR_DB_URL`, `.venv\`, `logs\`, легаси `connector.db`) |
| Server БД | PostgreSQL `connector_prod` (см. § 4b) на shared `127.0.0.1:5432` |
| Server backup | `C:\Connector\backup\` (verified PostgreSQL dumps in `postgres\` + legacy SQLite snapshots + Scheduled Task XMLs + `server.archive-2026-05-08\`) |
| Tekla repo worktree | `C:\Connector\tekla-firm-standard` (`PROBIM55/tekla-firm-standard`, ref `main`) |
| Tekla на клиентах | `C:\Company\TeklaFirm` (локальная копия из subdir `XS_FIRM`) |
| Соседний сервис 1 | `budget.structura-most.ru` / `budzhetapp.ru` — FamilyBudget (Node.js, `C:\apps\family-budget`, порт 3001, Windows-служба `FamilyBudget`, использует PostgreSQL на :5432). **Не ломать.** |
| Соседний сервис 2 | Platform — продукт-генератор геометрии (мосты, ШЗЭ, армирование). Код в `C:\Platform\` (`bridge-runtime/`, `deploy/`, `dotnet/`, `releases/`, `server/current/`). Запуск через Scheduled Task `PlatformServerApp`. За Caddy: `server.structura-most.ru/alexander*` и fallback `/` (на `127.0.0.1:8090`). **Не ломать.** |
| Соседний сервис 3 | Tekla MultiUser xs_server (Trimble) — Windows-служба «Tekla Structures Multiuser Server (managed by AlwaysUpService)», порт 1238. **Сторонний бинарник, не трогать.** |
| Соседний сервис 4 | Autodesk Revit Server 2024 — установка `C:\Autodesk\Revit_Server_2024_Win_64bit_dlm`, Windows-служба `Revit Server AutoSync 2024`, работает поверх **IIS** (`W3SVC` + `WAS`). **Сторонний продукт, не трогать.** |

Managed firewall ports: `80, 443, 445, 1238, 3389`.

## 3. Ключевые API endpoints

| Endpoint | Метод | Auth | Назначение |
|---|---|---|---|
| `/health` | GET | — | Public health check |
| `/connect/bootstrap` | POST | `X-Device-Token` | Логин клиента → device_id, session_id, SMB, heartbeat interval, manifest URL |
| `/heartbeat` | POST | `X-Device-Token` + `X-Device-Session` | Один токен = одна активная сессия |
| `/updates/latest.json` | GET | — | Manifest desktop-обновлений (берёт из GitHub Releases) |
| `/devices` | GET | admin | Legacy-список устройств; публичный доступ запрещён |
| `/admin/login` | POST | JSON `{username,password}` → 303 | Логин админа |
| `/admin/tokens` | GET | **`X-Admin-Key`** или Basic Auth | Список токенов и SMB-кредов |
| `/admin/updates/refresh` | POST | admin | Сброс кэша update manifest |

> ⚠️ **Документационный дрейф:** в `CREDENTIALS.md` admin-заголовок назван `X-Admin-Api-Key`, но рабочий — **`X-Admin-Key`**. Сам API-ключ корректный, имя поля в доке устаревшее.

## 3a. Что Connector управляет — и что нет

**Connector ОТВЕЧАЕТ за** (всё внутри `C:\Connector\server\` или порождается его кодом):

- свой uvicorn-процесс на `127.0.0.1:8080` (Scheduled Task `ConnectorApi`, не Windows-служба);
- Scheduled Task `ConnectorPostgresBackup` (daily 02:30 SYSTEM) — создаёт и проверяет `pg_restore --list` дамп PostgreSQL в `C:\Connector\backup\postgres\`;
- Scheduled Task `ConnectorBackupRetention` (daily 04:00 SYSTEM) — чистит только известные DB/task snapshots в `C:\Connector\backup\` по политике 30 дней daily + 12 недель weekly + 6 месяцев monthly (`scripts/server_backup_retention.ps1`);
- Scheduled Task `ConnectorYandexBackup` (daily 03:00 SYSTEM) — encrypted backup `C:\BIM_Models` (≈9.2 GiB / 25k файлов) и PostgreSQL dumps на Yandex.Disk через `yandex_crypt:` с versioning через `--backup-dir archive/<date>/`, retention 90 дней. Скрипт `scripts/backup_bim_to_yandex.ps1`, конфиг и ключи `runtime/rclone.conf`, лог `runtime/logs/backup_yandex.log`. Полный runbook в `doc/YANDEX_BACKUP_RU.md`;
- локальные Windows-пользователи `bim_*` — создаёт/удаляет `smb_user_manager.ps1`, пароли отдаются через bootstrap;
- SMB-права на share `BIM_Models` — те же `bim_*` users;
- Windows Firewall allowlist managed-портов (80/443/445/1238/3389) — `firewall_manager.ps1`, allowlist по `public_ip` клиентов;
- git worktree `C:\Connector\tekla-firm-standard` — server-side `git add/commit/push` в `PROBIM55/tekla-firm-standard` (ref `main`);
- **PostgreSQL `connector_prod`** на shared `127.0.0.1:5432` — tokens, sessions, devices, Tekla state, audit. User `connector_user` (OWNER), connection string в `runtime/.env`. Schema управляется `connector/server/migrations/postgres/*.sql` (yoyo);
- кэш update manifest из `PROBIM55/ServerConnector` releases.

**Connector НЕ управляет** (трогать из нашего кода нельзя):

- **Caddy** — общий reverse-proxy на 80/443. Конфиг `C:\caddy\Caddyfile` правится вручную (рядом 8 `.bak`-файлов). Перезапуск Caddy оборвёт **все** домены сервера.
- **PostgreSQL процесс** (`:5432`, общий c FamilyBudget и Platform) — мы пользуемся им как клиент в `connector_prod`. Свою БД менять можно. **Рестарт самого процесса `postgres` оборвёт все три продукта** — согласовывать. Data dir `C:\apps\family-budget\postgres-data\`.
- **FamilyBudget** — соседний продукт (Node.js на `:3001`).
- **Platform / `PlatformServerApp`** — соседний продукт (геометрия мостов/ШЗЭ, `platform_prod` в том же PG).
- **Tekla MultiUser xs_server** — сторонний.
- **Autodesk Revit Server 2024** — сторонний продукт. **IIS** (`W3SVC`, `WAS`) запущен именно ради него — перезапуск IIS оборвёт Revit Server.
- **Zabbix Agent** (`:10050`), OpenSSH, RDP — инфра.

> ⚠️ **Caddy fallback на 8090 пустой.** В `Caddyfile`: `server.structura-most.ru/` (всё, что не матчит специфичные маршруты) → `127.0.0.1:8090`. Сейчас на 8090 никто не слушает (PlatformServerApp не запущен или crashed). Для всех Connector-маршрутов (`/admin*`, `/health`, `/connect/*`, `/heartbeat`, `/updates/*`, `/devices`, `/ops*`) это безопасно — они идут на 8080. Но для не-Connector-путей будет 502. Если увидишь репорт «server.structura-most.ru возвращает 502» — это про Platform, не про Connector.

## 4. Релиз-флоу

### 4a. Desktop клиент (MSI)

Триггер — push семвер-тега `v*` в `master`.

1. Локально собрать MSI (опционально, для проверки): `powershell -NoProfile -ExecutionPolicy Bypass -File "connector-desktop/build_msi.ps1"`
2. Закоммитить изменения.
3. `git tag vX.Y.Z`
4. `git push origin master vX.Y.Z`
5. GitHub Actions (`.github/workflows/release-connector-desktop.yml`) подменяет версию в `Connector.Desktop.csproj` и `Package.wxs`, собирает MSI, публикует GitHub Release с asset'ом, и POST'ит на `/admin/updates/refresh` чтобы сервер сразу подхватил новую версию.

### 4b. Серверный код (`connector/server/`)

Триггер — push в `master` с changes в `connector/server/**`, `scripts/server_deploy_step.ps1` или сам workflow.

1. Локально внести изменения, прогнать `pytest` в `connector/server/` (опционально, CI всё равно прогонит). Тесты по умолчанию используют SQLite через tmp_path — для PG-проверки задать `CONNECTOR_DB_URL` env vars в conftest при необходимости.
2. **DB schema changes:** добавить нумерованную миграцию в **обе** директории — `connector/server/migrations/postgres/<NNNN>_<name>.sql` (прод) и `connector/server/migrations/sqlite/<NNNN>_<name>.sql` (тесты + dev fallback). Никаких `CREATE TABLE`/`ALTER TABLE` в коде `app.py` — `init_db()` вызывает yoyo для applied миграций. Для PG используй `BIGSERIAL`/IDENTITY, для SQLite — `INTEGER PRIMARY KEY AUTOINCREMENT`.
3. `git push origin master`.
4. GitHub Actions (`.github/workflows/deploy-connector-server.yml`):
   - Job **`test`**: setup Python 3.11, ставит `requirements-dev.txt`, гоняет `pytest -v`. Без green test'а deploy не пускается.
   - Job **`deploy`** (нужен test): SSH на `62.113.36.107` под `opwork_admin`, запускает `C:\Connector\src\scripts\server_deploy_step.ps1 -CommitSha $github.sha`. Скрипт делает: `git fetch + reset --hard origin/master`, `pip install -r requirements.txt`, verified PostgreSQL dump (или SQLite snapshot в legacy-конфигурации), `python run_migrations.py`, рестарт Scheduled Task `ConnectorApi`, smoke `GET /health` (проверяет 200 + `version` совпадает с deployed SHA).
   - **Auto-rollback** при провале любого шага: `git reset --hard <prevSha>`, рестарт Task'а.
5. Concurrency group `deploy-connector-server` гарантирует один деплой за раз.

**Скрипты `scripts/remote_deploy_connector.ps1` / `remote_finish_connector_deploy.ps1` — DEPRECATED.** Они относятся к старой архитектуре копирования из `C:\Users\opwork_admin\connector\server` и могли уничтожить `connector.db`. Не использовать.

**Раскладка сервера** (после Phase 1 переезда 2026-05-08):
- `C:\Connector\src\` — git checkout `PROBIM55/ServerConnector` ветки `master`. Источник кода.
- `C:\Connector\runtime\` — БД, config, venv, логи, **`.env`** для секретов (`CONNECTOR_DB_URL`). Переживает деплой.
- `C:\Connector\backup\` — снапшоты `connector.db` перед каждым деплоем + один `connector.db.pre-pg-cutover-*` (legacy SQLite после миграции на PG).

**БД** (после миграции 2026-05-09): PostgreSQL `connector_prod` на `127.0.0.1:5432` (тот же экземпляр PG, что у FamilyBudget и Platform — у каждого своя БД и user). User `connector_user`, OWNER только `connector_prod`. Connection string в `C:\Connector\runtime\.env` через `CONNECTOR_DB_URL`. Старая SQLite `C:\Connector\runtime\connector.db` остаётся как safety-net на 2-3 недели — не используется. **Откат на SQLite:** удалить/переименовать `C:\Connector\runtime\.env`, рестартовать `ConnectorApi` — фолбэк сработает автоматически (но live-данные за время на PG останутся в `connector_prod`, не в SQLite).

### 4c. Tekla firm content

Не идёт через эти flow. См. п. 5 ниже — admin делает publish через desktop UI, сервер сам коммитит/пушит в `PROBIM55/tekla-firm-standard`.

## 5. Tekla publish — критичные правила

Реальный admin_firm flow в desktop UI отдаёт серверу только `source_path` + `comment`. Сервер сам делает sync/git add/commit/push, считает version и revision из git.

**Подводные камни:**

- `source_path` должен быть **server-visible** — UNC-путь, не клиентский mapped drive (например, `P:\Tekla\...` не сработает). Ошибка `source_path must point to an existing directory` = указан путь, который сервер не видит.
- Admin **не** вводит version/revision вручную — генерируются на сервере.
- Серверный корень: `\\62.113.36.107\BIM_Models\Tekla\02_ПАПКА ФИРМЫ`. Синхронизируется из подпапки `01_XS_FIRM`. Остальные подпапки (обучение, видео, плагины и т.д.) **не должны** синхронизироваться.
- Tekla runtime у пользователей умышленно локальный (`C:\Company\TeklaFirm`), не сетевой — стабильнее, не зависит от SMB во время работы.

## 6. Где лежат секреты (значения **не** в этом файле)

| Что | Где |
|---|---|
| SSH ключ к VPS | `.secrets/opwork_vps_key` (`opwork_admin@62.113.36.107`) |
| Admin login/password/API key | `CREDENTIALS.md` |
| Runtime config (источник правды) | `C:\Connector\runtime\config.json` на проде |
| Runtime env (CONNECTOR_DB_URL = PG connection string) | `C:\Connector\runtime\.env` на проде, читается `runtime_launch.ps1` |
| Yandex.Disk OAuth-токен | `C:\Connector\runtime\rclone.conf` на проде (gitignored через расположение в runtime\) |
| Runtime БД | PostgreSQL `connector_prod` (legacy SQLite `C:\Connector\runtime\connector.db` остаётся как safety-net) |
| GitHub auth | `gh auth status` (Account `Scccoco`, keyring) |
| GitHub Actions secrets | repo settings, plaintext недоступны (`DEPLOY_HOST/USER/PORT/SSH_KEY` существуют, см. § 4b) |

> ⚠️ **`CREDENTIALS.md` лежит в репозитории с реальными паролями.** Перед любым `git add`/`git push` проверить `.gitignore` и `git log --all -- CREDENTIALS.md`. Если он попал в push на `origin` — пароль/API-ключ/SMB-креды скомпрометированы и подлежат ротации.

## 7. История восстановления локальной копии

**Историческая справка (для контекста — сейчас всё ОК).** До 2026-05-08 локальная копия `E:\00_Cursor\18_Server` была частично уничтожена: много `*_RU.md` файлов обнулено (NUL-байты), исходный код `connector/server` и `connector-desktop` отсутствовал, `.git` имел corrupted loose object `6eb0752f…` и broken-name теги. 2026-05-08 переклонировали свежим из `https://github.com/PROBIM55/ServerConnector.git` (ветка `master`), старая копия переименована в `E:\00_Cursor\18_Server_archive` (на диске лежит, но не используется).

**Сейчас:** `E:\00_Cursor\18_Server` — здоровый клон, в синке с `origin/master`. Если увидишь обнулённые файлы или corrupted `.git` — это означает что-то пошло не так и нужно ситуацию диагностировать заново. Архив можно при необходимости открыть (там CREDENTIALS.md и .secrets/ скопированы во время переезда).

## 8. Известные инциденты и выводы (для будущих сессий)

- **«Медленный SMB»** часто = заполненный диск C: на VPS. При увеличении диска через панель VPS не забыть расширить том в Windows.
- **«Токены пропали в админке»** = UI-баг (`deviceIdValue` использовался до инициализации в `admin_ui.html`), а не потеря данных в БД. Различать UI-поломку и data loss.
- **`HTTP 401: Invalid token`** обычно = устаревший client-side cached token, не всегда runtime-проблема.
- **Нестабильный runtime сервера** проявлялся как дубли uvicorn / bind-ошибки на 8080 — держать один консистентный production launch path.
- **Ошибка SSH-команд через PowerShell-сессию:** при `ssh ... "cmd1; cmd2"` точки с запятой парсятся PS-сессией на удалённой стороне как разделители команд PS — может ломаться. Безопаснее: `ssh ... "cmd /c echo ..."` или однострочные команды.

## 9. Конвенции

- **Не класть реальные секреты в `.md` репозитория.** Только pointers и описание назначения.
- **Не полагаться на старую документацию,** если она расходится с реальным UI/кодом — проверять прод-факт через код + runtime + логи.
- **Документация — на русском.** Соответствуй стилю существующих `*_RU.md`.
- **Перед изменением `CREDENTIALS.md`** — убедиться, что он в `.gitignore` и не в публичной истории.
- **Не делать destructive git** (`reset --hard`, force-push, удаление веток) без явного подтверждения от пользователя.
- **Изменения схемы БД — только через миграции** в `connector/server/migrations/{postgres,sqlite}/`. Никаких `CREATE TABLE`/`ALTER TABLE` в коде сервера. Каждая миграция должна быть в обоих диалектах (или хотя бы в PG для прода).
- **Backend abstraction:** в коде использовать `db_connect()` из `app.py` (или `db.connect()` из `connector/server/db.py`), **не** `sqlite3.connect(...)` напрямую. Wrapper переводит `?`-параметры в `%s` для PG автоматически. Параметры запросов писать в sqlite-стиле с `?`.
- **Перезапуск процесса `postgres`** на VPS затрагивает 3 сервиса (Connector + Platform + FamilyBudget) — координировать с владельцами.

## 10. Полезные одноразовые проверки

```powershell
# Жив ли сервер + какая версия (commit SHA) задеплоена
curl.exe -sS https://server.structura-most.ru/health
# → {"ok":true,"version":"<short-sha>"}

# Полная версия (admin)
curl.exe -sS -H "X-Admin-Key: <key>" https://server.structura-most.ru/admin/version

# Текущая версия desktop MSI
curl.exe -sS https://server.structura-most.ru/updates/latest.json

# Список токенов (admin)
curl.exe -sS -H "X-Admin-Key: <key>" https://server.structura-most.ru/admin/tokens

# SSH (Windows OpenSSH)
ssh -i .secrets\opwork_vps_key opwork_admin@62.113.36.107 "cmd /c echo %USERNAME% %COMPUTERNAME%"

# GitHub
gh repo view PROBIM55/ServerConnector --json name,defaultBranchRef,latestRelease
gh run list -R PROBIM55/ServerConnector --workflow deploy-connector-server.yml --limit 5
```

## 11. Указатели на восстановленный контекст

Если нужно глубже — читать в порядке убывания актуальности:

1. `RECOVERED_CONTEXT_RU.md` — продукт, инфра, инциденты, релизы, Tekla-модель.
2. `RECOVERED_ACCESS_POINTERS_RU.md` — карта секретов и где они живут.
3. `CREDENTIALS.md` — фактические значения (с осторожностью, см. п. 6).
4. `SETUP_NOTES.md` — короткая выжимка по VPS/портам.

Остальные `*_RU.md` в корне и `connector-desktop/` сейчас обнулены — игнорировать локально, читать только с GitHub после переклонирования.
