# Structura Connector Desktop

## Текущая функциональность

- Токеновый вход: пользователь вводит только токен.
- Bootstrap параметров с сервера (сессия устройства, настройки доступа, проверка связи).
- Фоновая отправка служебного сигнала связи после успешного подключения.
- Локальное хранение чувствительных данных через DPAPI (`CurrentUser`).
- Работа в трее при закрытии окна.
- Автозапуск в Windows (`HKCU\...\Run`).

## Логика обновлений

- Manifest URL: `https://server.structura-most.ru/updates/latest.json`.
- Проверка обновлений при запуске и каждые 30 минут.
- Manifest и MSI принимаются только по HTTPS; перед запуском установщика
  Connector сверяет его SHA-256 с digest релизного артефакта.
- В интерфейсе одна кнопка действия:
  - `Проверить обновление`;
  - `Скачать и установить` (если найдена более новая версия).
- После обновления токен сохраняется, повторная авторизация не требуется.

## Ограничение по сессиям токена

- Один токен поддерживает одну активную машину в момент времени.
- Новое подключение тем же токеном деактивирует предыдущую сессию.
- Для параллельной работы на нескольких ПК необходимы разные токены.

## Релизы

- Публикация пользовательского MSI выполняется через GitHub Releases.
- Актуальный артефакт релиза: `Connector.Desktop.Setup.msi`.

## Локальная сборка MSI

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "connector-desktop/build_msi.ps1"
```

Во время сборки `build_msi.ps1` автоматически подготавливает встроенный bundled git (MinGit) в `Connector.Desktop/tools/git`, чтобы клиентам не требовалась отдельная установка Git.

## Smoke check Tekla Adapter

Быстрая проверка manifest и базового контура:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "connector-desktop/scripts/tekla_adapter_smoke_check.ps1" -ServerBaseUrl "https://server.structura-most.ru"
```

## Документы по Tekla Standard Adapter

- `connector-desktop/CONNECTOR_TEKLA_ADAPTER_PLAN_RU.md`
- `connector-desktop/CONNECTOR_TEKLA_ADAPTER_CHECKLIST_RU.md`
- `connector-desktop/CONNECTOR_TEKLA_ADAPTER_PHASES_RU.md`
- `connector-desktop/TEKLA_STANDARD_PUBLISH_GUIDE_RU.md`
- `connector-desktop/TEKLA_STANDARD_USER_GUIDE_RU.md`
- `connector-desktop/TEKLA_STANDARD_RUNBOOK_RU.md`
- `connector-desktop/TEKLA_XS_FIRM_SETUP_RU.md`
- `connector-desktop/TEKLA_STANDARD_TEST_PROTOCOL_RU.md`
