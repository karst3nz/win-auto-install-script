# Windows Auto-Install Script

Автоматический скрипт настройки Windows с установкой пакетов через Chocolatey.

## Требования

- **Windows** 10/11
- **PowerShell** 5.1 или выше
- **Права администратора** (скрипт требует запуска от имени администратора)
- **Интернет-соединение**

## Быстрый старт

```powershell
# Запустить от имени администратора
.\setup.ps1
```

## Параметры запуска

| Параметр | Описание |
|----------|----------|
| `-NoRestart` | Не предлагать перезагрузку после завершения |
| `-NoUpgrade` | Пропустить обновление пакетов и Chocolatey |
| `-DryRun` | Режим симуляции без реальных изменений |
| `-ConfigPath <путь>` | Путь к JSON-файлу конфигурации (по умолчанию: `.\packages.config.json`) |
| `-Uninstall` | Удалить пакеты из конфигурационного файла |
| `-ExportConfig <путь>` | Экспортировать установленные пакеты в JSON |
| `-SkipActivation` | Пропустить активацию Windows |

### Примеры использования

```powershell
# Обычная установка
.\setup.ps1

# Режим симуляции (без реальных изменений)
.\setup.ps1 -DryRun

# Экспорт установленных пакетов
.\setup.ps1 -ExportConfig "my-packages.json"

# Удаление пакетов из конфига
.\setup.ps1 -Uninstall

# Установка без перезагрузки и обновления
.\setup.ps1 -NoRestart -NoUpgrade

# Пропустить активацию Windows
.\setup.ps1 -SkipActivation

# Использовать альтернативный конфиг
.\setup.ps1 -ConfigPath "C:\path\to\custom-config.json"
```

## Конфигурационный файл

Файл `packages.config.json` содержит список пакетов для установки:

```json
{
  "packages": [
    {
      "id": "7zip",
      "name": "7-Zip"
    },
    {
      "id": "googlechrome",
      "name": "Google Chrome"
    },
    {
      "id": "vlc",
      "name": "VLC Media Player"
    }
  ]
}
```

### Обязательные поля

| Поле | Тип | Описание |
|------|-----|----------|
| `id` | string | Идентификатор пакета в репозитории Chocolatey |
| `name` | string | Отображаемое имя пакета |

### Поиск пакетов

Найти нужный пакет можно на [community.chocolatey.org/packages](https://community.chocolatey.org/packages)

## Что делает скрипт

1. **Проверяет интернет-соединение**
2. **Устанавливает/обновляет Chocolatey**
3. **Устанавливает пакеты** из `packages.config.json`
4. **Активирует Windows** (через MAS-скрипт)
5. **Включает тёмную тему**
6. **Добавляет пункты в контекстное меню:**
   - "Copy as Path" — копировать путь к файлу/папке
   - "Open in PowerShell" — открыть PowerShell в папке
7. **Обновляет все пакеты** (можно пропустить `-NoUpgrade`)
8. **Очищает старые логи** (старше 30 дней)

## Структура проекта

```
D:\win-auto-install-script\
├── setup.ps1              # Главный скрипт
├── packages.config.json   # Конфигурация пакетов
├── README.md              # Документация
├── setup_*.log            # Логи запусков
└── packages.config.json   # Пример конфига
```

## Логи

Каждый запуск создаёт лог-файл с меткой времени:
```
setup_20260327_103240.log
```

Логи хранятся 30 дней, затем автоматически удаляются.

## Безопасность

### Проверка хеша скриптов

Скрипт загружает MAS (Microsoft Activation Script) с проверкой:
- Скачивание с повторными попытками (до 3 раз)
- Вычисление SHA256-хеша загруженного файла
- Возможность верификации по известному хешу

Для установки ожидаемого хеша отредактируйте `$Configuration.MAS.ExpectedHash` в начале скрипта.

### Откат изменений

Скрипт поддерживает откат:
- Удаление установленных пакетов при ошибке
- Восстановление темы реестра

### Подпись скрипта

Для производственного использования рекомендуется подписать скрипт:

```powershell
# Создать самоподписанный сертификат
$cert = New-SelfSignedCertificate -DnsName "WinAutoInstall" -CertStoreLocation "Cert:\CurrentUser\My"

# Подписать скрипт
Set-AuthenticodeSignature -FilePath ".\setup.ps1" -Certificate $cert
```

## Устранение проблем

### Ошибка "Script execution is disabled"

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Ошибка установки пакета

1. Проверьте название пакета на community.chocolatey.org
2. Попробуйте установить вручную: `choco install <package-id>`
3. Проверьте лог `setup_*.log`

### Ошибка активации Windows

- Убедитесь, что есть интернет-соединение
- Попробуйте с `-SkipActivation` и активируйте вручную
- Проверьте [MAS GitHub](https://github.com/massgrave/MAS)

## Вклад в проект

### Добавление пакетов

Отредактируйте `packages.config.json`:

```json
{
  "packages": [
    {
      "id": "your-package-id",
      "name": "Your Package Name"
    }
  ]
}
```

### Сообщение об ошибке

Приложите:
1. Версию PowerShell: `$PSVersionTable.PSVersion`
2. Лог-файл `setup_*.log`
3. Описание проблемы

## Лицензия

MIT License — свободное использование и модификация.

## Ссылки

- [Chocolatey](https://chocolatey.org/)
- [MAS (Microsoft Activation Script)](https://github.com/massgrave/MAS)
- [PowerShell Documentation](https://docs.microsoft.com/powershell/)
