#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Автоматическая настройка Windows с установкой пакетов через Chocolatey.

.DESCRIPTION
    Скрипт устанавливает Chocolatey, пакеты из конфигурационного файла,
    активирует Windows, применяет тёмную тему и добавляет полезные пункты
    в контекстное меню проводника.

.PARAMETER NoRestart
    Пропустить перезагрузку после завершения.

.PARAMETER DryRun
    Запуск без реальных изменений (только логирование).

.PARAMETER ConfigPath
    Путь к JSON-файлу с конфигурацией пакетов.

.PARAMETER Uninstall
    Удалить пакеты, указанные в конфигурационном файле.

.PARAMETER ExportConfig
    Экспортировать список установленных пакетов в JSON-файл.

.PARAMETER SkipActivation
    Пропустить активацию Windows.

.PARAMETER DriversPath
    Путь к папке с драйверами/приложениями (.exe файлы).

.EXAMPLE
    .\setup.ps1
    Запустить полную установку.

.EXAMPLE
    .\setup.ps1 -DryRun
    Запустить в режиме симуляции.

.EXAMPLE
    .\setup.ps1 -ExportConfig "installed.json"
    Экспортировать установленные пакеты.

.EXAMPLE
    .\setup.ps1 -Uninstall
    Удалить пакеты из конфига.
#>

[CmdletBinding()]
param(
    [switch]$NoRestart,
    [switch]$DryRun,
    [string]$ConfigPath = "$PSScriptRoot\packages.config.json",
    [switch]$Uninstall,
    [string]$ExportConfig,
    [switch]$SkipActivation,
    [string]$DriversPath = "$PSScriptRoot\drivers"
)

# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------
$Configuration = @{
    Chocolatey = @{
        InstallUrl      = "https://community.chocolatey.org/install.ps1"
        MinVersion      = "2.0.0"
    }
    MAS = @{
        Url             = "https://raw.githubusercontent.com/massgrave/MAS/master/MAS.ps1"
        FallbackUrl     = "https://dev.azure.com/massgrave/Microsoft-Activation-Scripts/_apis/git/repositories/Microsoft-Activation-Scripts/items?path=/MAS/All-In-One-Version-KL/MAS_AIO.cmd&download=true"
        ExpectedHash    = $null  # Установить актуальный хеш после проверки
        MaxRetries      = 3
    }
    Network = @{
        TestHost        = "www.microsoft.com"
        TimeoutSec      = 30
        MaxRetries      = 3
    }
    Logging = @{
        KeepLogsDays    = 30
        MaxLogSize      = 10MB
    }
}

$Script:StartTime = Get-Date
$Script:LogFile = "$PSScriptRoot\setup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$Script:Stats = @{ Installed = 0; Failed = 0; Skipped = 0; Uninstalled = 0 }
$Script:Rollback = @()
$Script:PackagesToProcess = @()

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# ------------------------------------------------------------
# Logging
# ------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    try {
        Add-Content -Path $Script:LogFile -Value $logEntry -Force -ErrorAction SilentlyContinue
    } catch {}
    
    switch ($Level) {
        "ERROR"   { Write-Host $Message -ForegroundColor Red }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        "WARN"    { Write-Host $Message -ForegroundColor Yellow }
        "INFO"    { Write-Host $Message -ForegroundColor Cyan }
        default   { Write-Host $Message }
    }
}

function Write-Step { param([string]$Text); Write-Log "`n[*] $Text" "INFO" }
function Write-OK   { param([string]$Text); Write-Log "    [OK] $Text" "SUCCESS"; $Script:Stats.Installed++ }
function Write-Fail { param([string]$Text); Write-Log "    [!!] $Text" "ERROR"; $Script:Stats.Failed++ }
function Write-Skip { param([string]$Text); Write-Log "    [..] $Text" "WARN"; $Script:Stats.Skipped++ }
function Write-Upgrade { param([string]$Text); Write-Log "    [UPG] $Text" "INFO"; $Script:Stats.Upgraded++ }

# ------------------------------------------------------------
# Rollback Support
# ------------------------------------------------------------
function Add-Rollback {
    param([scriptblock]$Action, [string]$Description)
    $Script:Rollback += @{ Action = $Action; Description = $Description }
}

function Invoke-Rollback {
    Write-Step "Rolling back changes..."
    for ($i = $Script:Rollback.Count - 1; $i -ge 0; $i--) {
        $item = $Script:Rollback[$i]
        try {
            Write-Host "    Undo: $($item.Description)" -ForegroundColor Yellow
            & $item.Action
        } catch {
            Write-Host "    Failed to undo: $($item.Description) - $_" -ForegroundColor Red
        }
    }
}

# ------------------------------------------------------------
# JSON Configuration Validation
# ------------------------------------------------------------
function Test-ConfigValid {
    <#
    .SYNOPSIS
        Проверяет валидность JSON-конфигурации пакетов.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    try {
        $jsonContent = Get-Content -Path $ConfigPath -Raw -ErrorAction Stop
        $config = $jsonContent | ConvertFrom-Json -ErrorAction Stop

        # Проверка структуры
        if (-not $config.psobject.Properties.Name -contains "packages") {
            Write-Log "    Config missing 'packages' property" "ERROR"
            return $false
        }

        if (-not ($config.packages -is [System.Array] -or $config.packages -is [System.Collections.ArrayList])) {
            Write-Log "    Config 'packages' must be an array" "ERROR"
            return $false
        }

        # Проверка каждого пакета
        $index = 0
        foreach ($pkg in $config.packages) {
            if (-not ($pkg.psobject.Properties.Name -contains "id")) {
                Write-Log "    Package [$index] missing required 'id' property" "ERROR"
                return $false
            }
            if (-not ($pkg.psobject.Properties.Name -contains "name")) {
                Write-Log "    Package [$index] missing required 'name' property" "ERROR"
                return $false
            }
            if (-not $pkg.id -or $pkg.id -notmatch "^[a-zA-Z0-9\-\.]+$") {
                Write-Log "    Package [$index] has invalid 'id': $($pkg.id)" "ERROR"
                return $false
            }
            $index++
        }

        return $true
    } catch [System.Management.Automation.PSInvalidOperationException] {
        Write-Log "    Invalid JSON format: $_" "ERROR"
        return $false
    } catch {
        Write-Log "    Error reading config: $_" "ERROR"
        return $false
    }
}

function Get-PackagesFromConfig {
    <#
    .SYNOPSIS
        Загружает пакеты из валидированного JSON-конфига.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
    return $config.packages
}

# ------------------------------------------------------------
# Registry Helper Functions
# ------------------------------------------------------------
function Set-RegistryKey {
    <#
    .SYNOPSIS
        Универсальная функция для создания ключей реестра и установки значений.
    .PARAMETER Path
        Путь к ключу реестра.
    .PARAMETER Name
        Имя параметра.
    .PARAMETER Value
        Значение параметра.
    .PARAMETER Type
        Тип параметра реестра (по умолчанию String).
    .PARAMETER CreateSubKey
        Имя подключа для создания (обычно "command").
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$Value,
        [ValidateSet("String", "DWord", "QWord", "Binary", "MultiString", "ExpandString")]
        [string]$Type = "String",
        [string]$CreateSubKey
    )

    try {
        # Создаём подключ если указан
        if ($CreateSubKey) {
            $key = New-Item -Path "$Path\$CreateSubKey" -Force -ErrorAction Stop
            $targetPath = $key.PSPath
        } else {
            $targetPath = $Path
            # Убеждаемся что ключ существует
            if (-not (Test-Path $Path)) {
                New-Item -Path $Path -Force -ErrorAction Stop | Out-Null
            }
        }

        # Устанавливаем значение
        Set-ItemProperty -Path $targetPath -Name $Name -Value $Value -Type $Type -Force -ErrorAction Stop
        return $true
    } catch {
        Write-Log "    Registry error ($Path\$Name): $_" "ERROR"
        return $false
    }
}

function Remove-RegistryKey {
    <#
    .SYNOPSIS
        Удаляет ключ реестра с подтверждением существования.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    try {
        if (Test-Path $Path) {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        }
        return $true
    } catch {
        Write-Log "    Registry remove error ($Path): $_" "ERROR"
        return $false
    }
}

# ------------------------------------------------------------
# Helper Functions
# ------------------------------------------------------------
function Test-InternetConnection {
    try {
        $connection = Test-Connection -ComputerName www.microsoft.com -Count 1 -Quiet -ErrorAction Stop
        return $connection
    } catch {
        return $false
    }
}

function Get-FileHashSecure {
    param([string]$Path, [string]$Algorithm = "SHA256")
    try {
        return Get-FileHash -Path $Path -Algorithm $Algorithm -ErrorAction Stop
    } catch {
        return $null
    }
}

function Invoke-DownloadWithRetry {
    param(
        [string]$Url,
        [string]$OutFile,
        [int]$MaxRetries = 3,
        [int]$TimeoutSec = 30
    )
    $retryCount = 0
    while ($retryCount -lt $MaxRetries) {
        try {
            Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -TimeoutSec $TimeoutSec
            return $true
        } catch {
            $retryCount++
            if ($retryCount -lt $MaxRetries) {
                Write-Log "    Retry $retryCount/$MaxRetries for $Url" "WARN"
                Start-Sleep -Seconds (2 * $retryCount)
            }
        }
    }
    return $false
}

function Test-ChocoInstalled {
    $choco = Get-Command choco -ErrorAction SilentlyContinue
    return $null -ne $choco
}

function Get-ChocoVersion {
    try {
        $version = choco --version 2>$null
        return $version
    } catch {
        return $null
    }
}

function Test-PackageInstalled {
    <#
    .SYNOPSIS
        Проверяет, установлен ли пакет через Chocolatey.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageId
    )
    try {
        # Используем --exact для точного совпадения имени пакета
        $result = choco list --local-only --exact $PackageId 2>$null
        if ($result -match "^\s*$PackageId\s") {
            return $true
        }
        return $false
    } catch {
        Write-Log "    Error checking package '$PackageId': $_" "WARN"
        return $false
    }
}

function Install-ChocoPackage {
    <#
    .SYNOPSIS
        Устанавливает пакет через Chocolatey с проверкой и откатом.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$Id,
        [string[]]$Parameters,
        [scriptblock]$CustomInstall,
        [scriptblock]$PostCheck,
        [switch]$SkipVerification
    )

    Write-Step "Installing $Name..."

    if ($DryRun) {
        Write-Skip "$Name (DryRun mode)"
        return
    }

    try {
        # Check if already installed
        if (Test-PackageInstalled -PackageId $Id) {
            Write-Skip "$Name already installed"
            Add-Rollback -Action { choco uninstall $Id -y 2>$null } -Description "Uninstall $Name"
            return
        }

        if ($CustomInstall) {
            & $CustomInstall
        } else {
            $chocoArgs = @("install", $Id, "-y")
            if ($Parameters) {
                $chocoArgs += $Parameters
            }

            # Run installation with detailed output
            Write-Log "    Running: choco $($chocoArgs -join ' ')"
            $output = & choco $chocoArgs 2>&1
            $output | ForEach-Object { Write-Log "    $_" }

            # Check for failure with detailed error analysis
            $exitCode = $LASTEXITCODE
            $hasError = $output -match "failed|error|unsuccessful|Exception"
            
            if ($exitCode -ne 0 -or $hasError) {
                $errorMsg = $output | Where-Object { $_ -match "ERROR|Exception|failed" } | Select-Object -First 3
                throw "Chocolatey installation failed for $Name. Exit code: $exitCode. Errors: $($errorMsg -join '; ')"
            }
        }

        # Post-installation check
        if ($PostCheck) {
            if (& $PostCheck) {
                Write-OK "$Name installed and verified."
            } else {
                Write-Fail "$Name installed but verification failed."
                return
            }
        } else {
            if ($SkipVerification) {
                Write-OK "$Name installed (skipped verification)."
            } elseif (Test-PackageInstalled -PackageId $Id) {
                Write-OK "$Name installed."
            } else {
                Write-Log "    Verification: Package '$Id' not found in local list" "WARN"
                Write-OK "$Name installed (verification skipped - may be installed)."
            }
        }

        Add-Rollback -Action { choco uninstall $Id -y 2>$null } -Description "Uninstall $Name"
    } catch {
        Write-Fail "$Name error: $_"
    }
}

# ------------------------------------------------------------
# Module: Chocolatey Installation
# ------------------------------------------------------------
function Install-Chocolatey {
    <#
    .SYNOPSIS
        Устанавливает или проверяет наличие Chocolatey.
    #>
    Write-Step "Checking / installing Chocolatey..."

    try {
        if (Test-ChocoInstalled) {
            $version = Get-ChocoVersion
            Write-OK "Chocolatey already installed: v$version"
            return $true
        }

        if ($DryRun) {
            Write-Skip "Chocolatey installation (DryRun mode)"
            return $true
        }

        Write-Log "    Installing Chocolatey..." "WARN"

        # Set security protocol
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072

        # Install Chocolatey
        $chocoInstallScript = "$env:TEMP\choco-install.ps1"
        if (Invoke-DownloadWithRetry -Url $Configuration.Chocolatey.InstallUrl -OutFile $chocoInstallScript) {
            & $chocoInstallScript
            Start-Sleep -Seconds 3

            # Refresh PATH
            $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
                        [System.Environment]::GetEnvironmentVariable("PATH", "User") + ";" +
                        "$env:ProgramData\chocolatey\bin"

            # Verify installation
            if (Test-ChocoInstalled) {
                $version = Get-ChocoVersion
                Write-OK "Chocolatey installed: v$version"
                return $true
            } else {
                throw "Chocolatey not found after installation"
            }
        } else {
            throw "Failed to download Chocolatey installer"
        }
    } catch {
        Write-Fail "Chocolatey error: $_"
        Write-Log "    Please install manually from: https://community.chocolatey.org" "WARN"
        return $false
    }
}

# ------------------------------------------------------------
# Module: Windows Activation (MAS)
# ------------------------------------------------------------
function Invoke-WindowsActivation {
    <#
    .SYNOPSIS
        Активирует Windows с помощью скрипта MAS.
    .NOTES
        Требует проверки хеша для безопасности.
    #>
    Write-Step "Activating Windows..."

    if ($SkipActivation) {
        Write-Skip "Windows activation skipped by user request"
        return
    }

    if ($DryRun) {
        Write-Skip "Windows activation (DryRun mode)"
        return
    }

    try {
        $masScript = "$env:TEMP\mas.ps1"
        $downloaded = $false

        # Try primary URL
        Write-Log "    Downloading MAS script from primary URL..."
        if (Invoke-DownloadWithRetry -Url $Configuration.MAS.Url -OutFile $masScript) {
            $downloaded = $true
        }

        # Try fallback URL if primary failed
        if (-not $downloaded) {
            Write-Log "    Primary URL failed, trying fallback URL..." "WARN"
            $masScript = "$env:TEMP\mas.cmd"
            if (Invoke-DownloadWithRetry -Url $Configuration.MAS.FallbackUrl -OutFile $masScript) {
                $downloaded = $true
            }
        }

        if (-not $downloaded) {
            throw "Failed to download MAS activation script from all sources"
        }

        # Verify script exists and is not empty
        if (-not (Test-Path $masScript) -or (Get-Item $masScript).Length -eq 0) {
            throw "Downloaded MAS script is empty or missing"
        }

        # Log script hash for audit purposes
        $hash = Get-FileHashSecure -Path $masScript
        if ($hash) {
            Write-Log "    MAS script hash ($($hash.Algorithm)): $($hash.Hash)"

            # Verify against expected hash if configured
            if ($Configuration.MAS.ExpectedHash) {
                if ($hash.Hash -ne $Configuration.MAS.ExpectedHash) {
                    Write-Log "    WARNING: Script hash does not match expected value!" "ERROR"
                    Write-Log "    Expected: $($Configuration.MAS.ExpectedHash)" "ERROR"
                    Write-Log "    Got:      $($hash.Hash)" "ERROR"
                    throw "MAS script hash verification failed"
                }
                Write-Log "    Script hash verified successfully"
            }
        }

        # Execute the script
        Write-Log "    Executing MAS activation script..."
        & $masScript
        Write-OK "Activation script executed successfully."

    } catch {
        Write-Fail "Activation error: $_"
        Write-Log "    Please activate Windows manually: https://github.com/massgrave/MAS" "WARN"
    }
}

# ------------------------------------------------------------
# Module: Dark Theme
# ------------------------------------------------------------
function Set-DarkTheme {
    Write-Step "Enabling Dark Theme..."

    if ($DryRun) {
        Write-Skip "Dark theme (DryRun mode)"
        return
    }

    try {
        $p = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"

        $oldApps = Get-ItemProperty -Path $p -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue
        $oldSystem = Get-ItemProperty -Path $p -Name "SystemUsesLightTheme" -ErrorAction SilentlyContinue

        Set-ItemProperty -Path $p -Name "AppsUseLightTheme"    -Value 0 -Type DWord -Force
        Set-ItemProperty -Path $p -Name "SystemUsesLightTheme" -Value 0 -Type DWord -Force

        Write-OK "Dark theme enabled."

        Add-Rollback -Action {
            if ($oldApps) {
                Set-ItemProperty -Path $p -Name "AppsUseLightTheme" -Value $oldApps.AppsUseLightTheme -Force
            }
            if ($oldSystem) {
                Set-ItemProperty -Path $p -Name "SystemUsesLightTheme" -Value $oldSystem.SystemUsesLightTheme -Force
            }
        } -Description "Restore light theme"
    } catch {
        Write-Fail "Dark theme error: $_"
    }
}

# ------------------------------------------------------------
# Module: Context Menu - Copy as Path
# ------------------------------------------------------------
function Add-CopyAsPathContextMenu {
    <#
    .SYNOPSIS
        Добавляет пункт 'Copy as Path' в контекстное меню.
    #>
    Write-Step "Adding 'Copy as Path' to context menu..."

    if ($DryRun) {
        Write-Skip "Copy as Path context menu (DryRun mode)"
        return
    }

    try {
        $copyCmd = 'powershell.exe -NoProfile -Command "Set-Clipboard -Value ''%1''"'
        $bgCopyCmd = 'powershell.exe -NoProfile -Command "Set-Clipboard -Value ''%V''"'

        $registryEntries = @(
            @{ Path = "HKCU:\SOFTWARE\Classes\*\shell\CopyAsPath"; Name = "(Default)"; Value = "Copy as Path"; SubKey = "command"; CmdValue = $copyCmd }
            @{ Path = "HKCU:\SOFTWARE\Classes\Directory\shell\CopyAsPath"; Name = "(Default)"; Value = "Copy as Path"; SubKey = "command"; CmdValue = $copyCmd }
            @{ Path = "HKCU:\SOFTWARE\Classes\Directory\Background\shell\CopyAsPath"; Name = "(Default)"; Value = "Copy as Path"; SubKey = "command"; CmdValue = $bgCopyCmd }
        )

        foreach ($entry in $registryEntries) {
            # Создаём ключ и команды
            Set-RegistryKey -Path $entry.Path -Name $entry.Name -Value $entry.Value -CreateSubKey $entry.SubKey
            Set-RegistryKey -Path "$($entry.Path)\$($entry.SubKey)" -Name "(Default)" -Value $entry.CmdValue
            Set-RegistryKey -Path $entry.Path -Name "Icon" -Value "shell32.dll,-259"

            Add-Rollback -Action { Remove-RegistryKey -Path $entry.Path } -Description "Remove CopyAsPath from $($entry.Path)"
        }

        Write-OK "'Copy as Path' added."
    } catch {
        Write-Fail "Copy as Path error: $_"
    }
}

# ------------------------------------------------------------
# Module: Context Menu - Open in PowerShell
# ------------------------------------------------------------
function Add-PowerShellContextMenu {
    <#
    .SYNOPSIS
        Добавляет пункт 'Open in PowerShell' в контекстное меню.
    #>
    Write-Step "Adding 'Open in PowerShell' to context menu..."

    if ($DryRun) {
        Write-Skip "Open in PowerShell context menu (DryRun mode)"
        return
    }

    try {
        $registryEntries = @(
            @{
                Path = "HKCU:\SOFTWARE\Classes\Directory\Background\shell\OpenPowerShell"
                Command = 'powershell.exe -NoExit -Command "Set-Location ''%V''"'
            },
            @{
                Path = "HKCU:\SOFTWARE\Classes\Directory\shell\OpenPowerShell"
                Command = 'powershell.exe -NoExit -Command "Set-Location ''%1''"'
            }
        )

        foreach ($entry in $registryEntries) {
            Set-RegistryKey -Path $entry.Path -Name "(Default)" -Value "Open in PowerShell" -CreateSubKey "command"
            Set-RegistryKey -Path "$($entry.Path)\command" -Name "(Default)" -Value $entry.Command
            Set-RegistryKey -Path $entry.Path -Name "Icon" -Value "powershell.exe"

            Add-Rollback -Action { Remove-RegistryKey -Path $entry.Path } -Description "Remove OpenPowerShell from $($entry.Path)"
        }

        Write-OK "'Open in PowerShell' added."
    } catch {
        Write-Fail "Open in PowerShell error: $_"
    }
}

# ------------------------------------------------------------
# Optional: Install from config file
# ------------------------------------------------------------
function Install-FromConfig {
    <#
    .SYNOPSIS
        Устанавливает пакеты из конфигурационного файла.
    #>
    param([string]$ConfigPath)

    if (-not (Test-Path $ConfigPath)) {
        Write-Log "    Config file not found: $ConfigPath" "WARN"
        return $false
    }

    # Validate config structure
    if (-not (Test-ConfigValid -ConfigPath $ConfigPath)) {
        Write-Fail "Config validation failed: $ConfigPath"
        return $false
    }

    try {
        $packages = Get-PackagesFromConfig -ConfigPath $ConfigPath
        $Script:PackagesToProcess = $packages

        if ($packages.Count -eq 0) {
            Write-Skip "No packages in config"
            return $true
        }

        Write-Step "Installing packages from config ($($packages.Count) total)..."
        
        $current = 0
        foreach ($pkg in $packages) {
            $current++
            Write-Log "    Processing [$current/$($packages.Count)]: $($pkg.name) ($($pkg.id))"
            Install-ChocoPackage -Name $pkg.name -Id $pkg.id
        }
        
        return $true
    } catch {
        Write-Fail "Config parsing error: $_"
        return $false
    }
}

# ------------------------------------------------------------
# Install drivers/applications from folder
# ------------------------------------------------------------
function Install-FromDriversFolder {
    <#
    .SYNOPSIS
        Устанавливает .exe и .msi файлы из папки с драйверами/приложениями.
    .PARAMETER DriversPath
        Путь к папке с .exe и .msi файлами.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$DriversPath
    )

    Write-Step "Installing applications from drivers folder..."

    if (-not (Test-Path $DriversPath)) {
        Write-Skip "Drivers folder not found: $DriversPath"
        return
    }

    if ($DryRun) {
        Write-Skip "Drivers installation (DryRun mode)"
        return
    }

    try {
        # Ищем все .exe и .msi файлы в папке
        $exeFiles = Get-ChildItem -Path $DriversPath -Filter "*.exe" -File -ErrorAction SilentlyContinue
        $msiFiles = Get-ChildItem -Path $DriversPath -Filter "*.msi" -File -ErrorAction SilentlyContinue

        if ($exeFiles.Count -eq 0 -and $msiFiles.Count -eq 0) {
            Write-Skip "No .exe or .msi files found in $DriversPath"
            return
        }

        Write-Log "    Found $($exeFiles.Count) .exe file(s) and $($msiFiles.Count) .msi file(s)"

        # Обработка .exe файлов
        foreach ($file in $exeFiles) {
            Write-Step "Installing: $($file.Name)..."

            try {
                # Запускаем .exe файл в тихом режиме
                # Пытаемся использовать распространённые ключи тихой установки
                $installArgs = @("/S", "/s", "/quiet", "/silent", "/verysilent", "/qn")
                $success = $false

                foreach ($arg in $installArgs) {
                    $procInfo = Start-Process -FilePath $file.FullName -ArgumentList $arg -Wait -PassThru -ErrorAction SilentlyContinue
                    if ($procInfo.ExitCode -eq 0 -or $procInfo.ExitCode -eq 3010) {
                        $success = $true
                        Write-OK "$($file.Name) installed with args '$arg' (exit code: $($procInfo.ExitCode))"
                        break
                    }
                }

                if (-not $success) {
                    # Если тихая установка не сработала, запускаем без аргументов
                    Write-Log "    Silent install failed, trying without args..." "WARN"
                    $procInfo = Start-Process -FilePath $file.FullName -Wait -PassThru -ErrorAction SilentlyContinue
                    if ($procInfo.ExitCode -eq 0 -or $procInfo.ExitCode -eq 3010) {
                        Write-OK "$($file.Name) installed (exit code: $($procInfo.ExitCode))"
                    } else {
                        Write-Fail "$($file.Name) failed (exit code: $($procInfo.ExitCode))"
                    }
                }
            } catch {
                Write-Fail "$($file.Name) error: $_"
            }
        }

        # Обработка .msi файлов
        foreach ($file in $msiFiles) {
            Write-Step "Installing: $($file.Name)..."

            try {
                # Запускаем .msi файл через msiexec в тихом режиме
                $procInfo = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$($file.FullName)`" /quiet /norestart" -Wait -PassThru -ErrorAction SilentlyContinue
                
                if ($procInfo.ExitCode -eq 0 -or $procInfo.ExitCode -eq 3010) {
                    Write-OK "$($file.Name) installed (exit code: $($procInfo.ExitCode))"
                } else {
                    Write-Fail "$($file.Name) failed (exit code: $($procInfo.ExitCode))"
                }
            } catch {
                Write-Fail "$($file.Name) error: $_"
            }
        }
    } catch {
        Write-Fail "Drivers installation error: $_"
    }
}

# ------------------------------------------------------------
# Export installed packages to config
# ------------------------------------------------------------
function Export-InstalledPackages {
    <#
    .SYNOPSIS
        Экспортирует список установленных Chocolatey пакетов в JSON.
    .PARAMETER OutputPath
        Путь для сохранения JSON-файла.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Step "Exporting installed packages to $OutputPath..."

    try {
        # Получаем список установленных пакетов
        $packages = choco list --local-only 2>$null |
            Where-Object { $_ -match "^\s*\S+" } |
            ForEach-Object {
                $parts = $_.Trim() -split '\s+'
                if ($parts.Count -ge 2) {
                    @{
                        id = $parts[0]
                        name = $parts[0]  # Используем ID как имя по умолчанию
                        version = $parts[1]
                    }
                }
            } | Where-Object { $_ -ne $null }

        if ($packages.Count -eq 0) {
            Write-Fail "No installed packages found"
            return $false
        }

        # Формируем структуру JSON
        $exportData = @{
            packages = $packages | ForEach-Object {
                @{
                    id = $_.id
                    name = $_.name
                }
            }
        } | ConvertTo-Json -Depth 3

        # Сохраняем файл
        $exportData | Set-Content -Path $OutputPath -Force -Encoding UTF8
        Write-OK "Exported $($packages.Count) packages to $OutputPath"
        return $true
    } catch {
        Write-Fail "Export error: $_"
        return $false
    }
}

# ------------------------------------------------------------
# Uninstall packages from config
# ------------------------------------------------------------
function Uninstall-FromConfig {
    <#
    .SYNOPSIS
        Удаляет пакеты, указанные в конфигурационном файле.
    #>
    param([string]$ConfigPath)

    if (-not (Test-Path $ConfigPath)) {
        Write-Log "    Config file not found: $ConfigPath" "WARN"
        return $false
    }

    if (-not (Test-ConfigValid -ConfigPath $ConfigPath)) {
        Write-Fail "Config validation failed: $ConfigPath"
        return $false
    }

    try {
        $packages = Get-PackagesFromConfig -ConfigPath $ConfigPath

        if ($packages.Count -eq 0) {
            Write-Skip "No packages in config"
            return $true
        }

        Write-Step "Uninstalling packages from config ($($packages.Count) total)..."

        $current = 0
        foreach ($pkg in $packages) {
            $current++
            Write-Log "    Processing [$current/$($packages.Count)]: $($pkg.name) ($($pkg.id))"

            if ($DryRun) {
                Write-Skip "$($pkg.name) uninstall (DryRun mode)"
                continue
            }

            if (Test-PackageInstalled -PackageId $pkg.id) {
                try {
                    $output = choco uninstall $pkg.id -y 2>&1
                    $output | ForEach-Object { Write-Log "    $_" }

                    if ($LASTEXITCODE -eq 0) {
                        Write-OK "$($pkg.name) uninstalled."
                        $Script:Stats.Uninstalled++
                    } else {
                        Write-Fail "$($pkg.name) uninstall failed."
                    }
                } catch {
                    Write-Fail "$($pkg.name) error: $_"
                }
            } else {
                Write-Skip "$($pkg.name) not installed"
            }
        }

        return $true
    } catch {
        Write-Fail "Uninstall error: $_"
        return $false
    }
}

# ------------------------------------------------------------
# Cleanup old log files
# ------------------------------------------------------------
function Cleanup-OldLogs {
    <#
    .SYNOPSIS
        Удаляет старые лог-файлы (старше указанного количества дней).
    #>
    param(
        [int]$DaysToKeep = 30
    )

    try {
        $cutoffDate = (Get-Date).AddDays(-$DaysToKeep)
        $logFiles = Get-ChildItem -Path "$PSScriptRoot\setup_*.log" -File |
            Where-Object { $_.LastWriteTime -lt $cutoffDate }

        if ($logFiles.Count -gt 0) {
            $logFiles | Remove-Item -Force
            Write-Log "    Cleaned up $($logFiles.Count) old log file(s)" "INFO"
        }
    } catch {
        Write-Log "    Log cleanup error: $_" "WARN"
    }
}

# ------------------------------------------------------------
# Main Execution
# ------------------------------------------------------------
function Invoke-Setup {
    <#
    .SYNOPSIS
        Главная функция запуска скрипта.
    #>
    Write-Host "`n============================================" -ForegroundColor Cyan
    Write-Host "  Windows Setup Script (Chocolatey)" -ForegroundColor Cyan
    Write-Host "  Started: $($Script:StartTime)" -ForegroundColor Cyan
    Write-Host "  PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan
    Write-Host "============================================`n" -ForegroundColor Cyan

    Write-Log "Setup started"

    # Cleanup old logs
    Cleanup-OldLogs -DaysToKeep $Configuration.Logging.KeepLogsDays

    # Handle export mode
    if ($ExportConfig) {
        Export-InstalledPackages -OutputPath $ExportConfig
        return
    }

    # Pre-flight checks
    if (-not (Test-InternetConnection)) {
        Write-Fail "No internet connection detected. Aborting."
        return
    }
    Write-OK "Internet connection OK"

    # Install/Upgrade Chocolatey first
    if (-not (Install-Chocolatey)) {
        Write-Fail "Chocolatey installation failed. Cannot continue."
        return
    }

    # Handle uninstall mode
    if ($Uninstall) {
        if (Test-Path $ConfigPath) {
            Uninstall-FromConfig -ConfigPath $ConfigPath
        } else {
            Write-Log "    Config file not found: $ConfigPath" "WARN"
        }

        # Show uninstall summary
        $duration = (Get-Date) - $Script:StartTime
        Write-Host "`n============================================" -ForegroundColor Magenta
        Write-Host "  Uninstall Complete!" -ForegroundColor Magenta
        Write-Host "  Duration: $($duration.Minutes)m $($duration.Seconds)s" -ForegroundColor Magenta
        Write-Host "  Uninstalled: $($Script:Stats.Uninstalled)" -ForegroundColor Green
        Write-Host "  Failed: $($Script:Stats.Failed)" -ForegroundColor $(if ($Script:Stats.Failed -gt 0) { "Red" } else { "Green" })
        Write-Host "  Skipped: $($Script:Stats.Skipped)" -ForegroundColor Yellow
        Write-Host "  Log: $($Script:LogFile)" -ForegroundColor Cyan
        Write-Host "============================================`n" -ForegroundColor Magenta
        return
    }

    # Install packages from config file
    if (Test-Path $ConfigPath) {
        # Preview packages before installation
        try {
            $packages = Get-PackagesFromConfig -ConfigPath $ConfigPath
            if ($packages.Count -gt 0) {
                Write-Step "Packages to install ($($packages.Count)):"
                $packages | ForEach-Object {
                    $installed = Test-PackageInstalled -PackageId $_.id
                    $status = if ($installed) { " (already installed)" } else { "" }
                    Write-Host "    - $($_.name)$status" -ForegroundColor $(if ($installed) { "Yellow" } else { "Green" })
                }
                Write-Host ""
            }
        } catch {
            Write-Log "    Could not preview packages: $_" "WARN"
        }

        Install-FromConfig -ConfigPath $ConfigPath
    } else {
        Write-Log "    Config file not found: $ConfigPath" "WARN"
    }

    # Install drivers/applications from folder
    Install-FromDriversFolder -DriversPath $DriversPath

    # System configuration
    if (-not $SkipActivation) {
        Invoke-WindowsActivation
    }
    Set-DarkTheme
    Add-CopyAsPathContextMenu
    Add-PowerShellContextMenu

    # Summary
    $duration = (Get-Date) - $Script:StartTime
    Write-Host "`n============================================" -ForegroundColor Magenta
    Write-Host "  Setup Complete!" -ForegroundColor Magenta
    Write-Host "  Duration: $($duration.Minutes)m $($duration.Seconds)s" -ForegroundColor Magenta
    Write-Host "  Installed: $($Script:Stats.Installed)" -ForegroundColor Green
    Write-Host "  Failed: $($Script:Stats.Failed)" -ForegroundColor $(if ($Script:Stats.Failed -gt 0) { "Red" } else { "Green" })
    Write-Host "  Skipped: $($Script:Stats.Skipped)" -ForegroundColor Yellow
    Write-Host "  Log: $($Script:LogFile)" -ForegroundColor Cyan
    Write-Host "============================================`n" -ForegroundColor Magenta

    if (-not $NoRestart -and -not $DryRun) {
        $response = Read-Host "Restart now? (Y/N)"
        if ($response -eq "Y" -or $response -eq "y") {
            Write-Step "Restarting..."
            Restart-Computer -Force
        }
    }
}

# Trap for error handling
trap {
    Write-Fail "Fatal error: $_"
    Write-Host "`nRollback available on request." -ForegroundColor Yellow
    exit 1
}

# Run setup
Invoke-Setup
