# Backup BIM_Models в Я.Диск через rclone

## Что делает

Ежедневно в 03:00 MSK копирует `C:\BIM_Models` (≈9 GB, 25k файлов) и
проверенные дампы PostgreSQL на Yandex.Disk. Данные и имена файлов шифруются
локально через `rclone crypt`, поэтому Yandex не блокирует легитимные DLL
из BIM-дистрибутивов проверкой содержимого. Только изменившиеся файлы
заливаются заново; заменённые/удалённые сохраняются в `archive/<date>/`
для восстановления. Архив хранится 90 дней.

## Архитектура

| Компонент | Где |
|---|---|
| rclone бинарь | `C:\Tools\rclone\rclone.exe` |
| rclone config (с OAuth-токеном) | `C:\Connector\runtime\rclone.conf` (gitignore'д через расположение в runtime\) |
| PostgreSQL backup | `C:\Connector\backup\postgres\*.dump`, task `ConnectorPostgresBackup` daily 02:30 SYSTEM |
| Backup-скрипт | `C:\Connector\src\scripts\backup_bim_to_yandex.ps1` (в git) |
| Scheduled Task | `ConnectorYandexBackup`, daily 03:00 SYSTEM |
| Логи | `C:\Connector\runtime\logs\backup_yandex.log` |
| Статус последнего запуска | `C:\Connector\runtime\last_yandex_backup.json` |

## Раскладка на Я.Диске

```
Structura/
└── BIM_Backup_Encrypted/         ← зашифрованное содержимое, руками не редактировать
```

Логическая раскладка через remote `yandex_crypt:`:

```
yandex_crypt:
├── bim/
│   ├── current/                  ← mirror C:\BIM_Models
│   └── archive/<yyyy-MM-dd>/     ← заменённые и удалённые версии
└── connector-db/
    ├── current/                  ← verified PostgreSQL dumps
    └── archive/<yyyy-MM-dd>/
```

## Initial setup (одноразовый)

### 1. Получить OAuth-токен Yandex.Disk

OAuth-flow требует браузер, на VPS его нет. Делается на ЛОКАЛЬНОЙ Windows-машине
с браузером:

```powershell
# Скачать rclone (~25 MB, без установки):
$tmp = Join-Path $env:TEMP 'rclone-oauth'
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
Invoke-WebRequest 'https://downloads.rclone.org/rclone-current-windows-amd64.zip' -OutFile "$tmp\rclone.zip"
Expand-Archive "$tmp\rclone.zip" -DestinationPath $tmp -Force
$rclone = (Get-ChildItem $tmp -Recurse -Filter rclone.exe | Select -First 1).FullName

# Запустить OAuth flow — откроется браузер:
& $rclone authorize "yandex"
```

В браузере залогинься в нужный Я.Диск-аккаунт, разреши rclone доступ.
В консоли rclone напечатает однострочный JSON-токен:

```
{"access_token":"y0_AgA...","token_type":"OAuth","refresh_token":"...","expiry":"2026..."}
```

Скопируй **всю строку JSON** (она в одинарных кавычках после `Paste the following into your remote machine -->`). Это и есть твой токен.

### 2. Применить токен на VPS

Передай мне токен (один скопированный JSON-блок), либо сам выполни:

```powershell
# На VPS под opwork_admin (через SSH или RDP):
powershell -NoProfile -ExecutionPolicy Bypass `
    -File C:\Connector\src\scripts\setup_yandex_rclone_config.ps1 `
    -Token '<вставь сюда JSON-токен в одинарных кавычках>'
```

Скрипт создаст/обновит `yandex_disk:`, один раз создаст шифрованный remote
`yandex_crypt:` и сделает тестовые запросы. Существующие ключи шифрования
при повторном запуске сохраняются.

Сразу сохраните полную копию `C:\Connector\runtime\rclone.conf` в защищённом
месте. Без значений `password` и `password2` расшифровать резервные копии
невозможно.

### 3. Настроить локальный PostgreSQL backup

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
    -File C:\Connector\src\scripts\setup_connector_postgres_backup_task.ps1

powershell -NoProfile -ExecutionPolicy Bypass `
    -File C:\Connector\src\scripts\backup_connector_postgres.ps1 `
    -Reason initial
```

Дамп считается успешным только после `pg_restore --list`. Частично созданный
файл публиковаться не будет.

### 4. Initial off-site backup (manual)

Первый запуск — вручную, чтобы убедиться, что всё работает и оценить время:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
    -File C:\Connector\src\scripts\backup_bim_to_yandex.ps1
```

Первый sync 9 GB может занять десятки минут.
Лог пишется в реальном времени в `C:\Connector\runtime\logs\backup_yandex.log`.

### 5. Scheduled Task

После успешного initial backup'а:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
    -File C:\Connector\src\scripts\setup_yandex_backup_task.ps1
```

Scheduled Task `ConnectorYandexBackup` создан. Можно проверить:

```powershell
Get-ScheduledTask -TaskName ConnectorYandexBackup
```

С этого момента бэкап будет идти автоматически каждую ночь в 03:00 MSK.

## Восстановление файла из бэкапа

### Восстановить актуальную версию

Копировать нужно через `yandex_crypt:`, иначе на диске будут видны только
зашифрованные имена и содержимое:

```powershell
& 'C:\Tools\rclone\rclone.exe' --config C:\Connector\runtime\rclone.conf `
    copyto 'yandex_crypt:bim/current/<путь к файлу>' `
           'C:\BIM_Models_restored\<имя файла>'
```

### Восстановить старую версию (например, 5 дней назад)

Старая версия находится логически в
`yandex_crypt:bim/archive/2026-05-04/<путь>`.

### Массовое восстановление за день (rclone)

```powershell
& 'C:\Tools\rclone\rclone.exe' --config C:\Connector\runtime\rclone.conf `
    copy yandex_crypt:bim/archive/2026-05-04 `
         C:\BIM_Models_restored
```

## Мониторинг

### Проверить что последний backup прошёл успешно

```powershell
Get-Content C:\Connector\runtime\logs\backup_yandex.log -Tail 5
# Ищи "BACKUP DONE total=..." в конце

Get-Content C:\Connector\runtime\last_yandex_backup.json
# "ok": true
```

```powershell
Get-ScheduledTaskInfo -TaskName ConnectorYandexBackup
# LastTaskResult должен быть 0
```

### Альтернатива через rclone

```powershell
& 'C:\Tools\rclone\rclone.exe' --config C:\Connector\runtime\rclone.conf `
    size yandex_crypt:bim/current
# Покажет size+count папки current — должно ≈ соответствовать C:\BIM_Models
```

## Что в gitignore / что в репо

В git живут:
- `scripts/backup_bim_to_yandex.ps1` — main script
- `scripts/backup_connector_postgres.ps1` — verified PostgreSQL dump
- `scripts/setup_connector_postgres_backup_task.ps1` — task PostgreSQL backup
- `scripts/setup_yandex_rclone_config.ps1` — bootstrap rclone config из токена
- `scripts/setup_yandex_backup_task.ps1` — bootstrap Scheduled Task
- `doc/YANDEX_BACKUP_RU.md` — этот документ

В git **НЕ** живут (на сервере runtime\):
- `C:\Connector\runtime\rclone.conf` — содержит OAuth-токен Я.Диска
- `C:\Connector\runtime\logs\backup_yandex.log` — оперативные логи

## Что трогать нельзя

- **Не удалять `runtime/rclone.conf`** — придётся заново проходить OAuth.
- **Не запускать `rclone sync` руками с другим dest** — может уничтожить уже
  залитое (sync = mirror). Если нужно эксперимент — `rclone copy` (без sync).
- **Не отключать Scheduled Task** без причины — пропуск даже одной ночи
  означает потерю point-in-time снимка для archive/.

## Известные ограничения

- **Открытые/locked файлы** (например, активная Tekla-сессия пишет в `.db1`)
  — rclone скипает с warning, в следующий запуск (на следующий день, когда
  файл закрыт) — заберёт. Не критично, но в логе будут warning'и.
- **Yandex API rate limits** — при 25k файлов первый run может ловить 429.
  `--tpslimit 5` ограничивает 5 RPS — стандартный лимит без проблем.
- **Большие файлы (>2 GB)** — Я.Диск ограничивает single file 50 GB. У нас
  максимум 158 MB (по inventory), запас огромный.

## Будущие улучшения

- **Telegram alert** при failed backup — после реализации Этапа 1 из ROADMAP_RU.md (общая alert-инфра).
- **Метрики backup'а** в Prometheus после Этапа 1 ROADMAP — последний успех,
  размер последнего sync'а, длительность.
