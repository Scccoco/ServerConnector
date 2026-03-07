# Structura Connector Desktop (MSI)

## Что реализовано
- Desktop-приложение WPF (`Structura Connector`).
- Настройка сервера, `device_id`, токена, SMB-доступа и интервала heartbeat.
- Ручная отправка heartbeat и автоотправка по таймеру.
- Локальное шифрование токена и SMB-пароля через DPAPI (`CurrentUser`).
- Сворачивание в трей при закрытии окна.
- Автозапуск через `HKCU\...\Run`.
- Проверка обновлений по `update manifest` и установка новой MSI из приложения.
- MSI-установщик с выбором пути установки и созданием ярлыков.

## Где лежит MSI
- `connector-desktop/artifacts/Connector.Desktop.Setup.msi`

## Сборка MSI
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "connector-desktop/build_msi.ps1"
```

## Что получает пользователь после установки
- Папка установки (по выбору в мастере MSI)
- Ярлык на рабочем столе: `Structura Connector`
- Ярлык в меню Пуск: `Structura Connector`

## Базовая проверка
1. Запустить `Structura Connector` с рабочего стола.
2. Вставить `Токен устройства` и при необходимости проверить `URL сервера`.
3. Нажать `Подключиться (по токену, автоматически)`.
4. Убедиться, что открыт SMB-доступ и включена автоотправка heartbeat.
5. В логе приложения увидеть `Heartbeat отправлен успешно`.
6. В админ-панели `/admin/ui` убедиться, что `updated_at` устройства обновляется.

## Формат update manifest
Приложение читает JSON по URL из поля `URL обновлений (json manifest)`.

Пример:

```json
{
  "version": "1.1.0",
  "msiUrl": "https://server.structura-most.ru/updates/files/Connector.Desktop.Setup.msi",
  "notes": "Исправления сети и улучшения подключения SMB"
}
```
