[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$manifestUrl = "https://hedgef0g.github.io/research-insights-toolkit/manifest.xml"
$manifestPath = Join-Path $env:TEMP "ResearchSignal-Table.xml"
$logPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "ResearchSignal-Installer.log"

function Show-InstallError([string]$Message) {
  "ERROR: $Message" | Add-Content -LiteralPath $logPath -Encoding UTF8
  Write-Host "`nНе удалось подготовить установку ResearchSignal Table." -ForegroundColor Red
  Write-Host $Message
  Write-Host "Подробности сохранены: $logPath"
  return $false
}

try {
  "ResearchSignal Table installer started: $(Get-Date -Format o)" | Set-Content -LiteralPath $logPath -Encoding UTF8
  Write-Host "ResearchSignal Table — установка для Excel" -ForegroundColor Cyan
  Write-Host "Журнал установки: $logPath"
  Write-Host "Скачиваю официальный манифест надстройки..."
  Invoke-WebRequest -Uri $manifestUrl -OutFile $manifestPath -UseBasicParsing
  "Manifest downloaded: $manifestPath" | Add-Content -LiteralPath $logPath -Encoding UTF8

  [xml]$manifest = Get-Content -LiteralPath $manifestPath -Raw
  $ns = New-Object System.Xml.XmlNamespaceManager($manifest.NameTable)
  $ns.AddNamespace("o", "http://schemas.microsoft.com/office/appforoffice/1.1")
  $source = $manifest.SelectSingleNode("/o:OfficeApp/o:DefaultSettings/o:SourceLocation", $ns)
  if (-not $source -or $source.DefaultValue -notlike "https://hedgef0g.github.io/research-insights-toolkit/*") {
    throw "Скачанный файл не прошёл проверку манифеста. Проверьте ссылку или обратитесь к владельцу пилота."
  }

  $excelPath = $null
  $appPathKey = "Registry::HKEY_CLASSES_ROOT\Excel.Application\CurVer"
  if (Test-Path $appPathKey) {
    $excelVersion = (Get-ItemProperty -Path $appPathKey).'(default)'
    if ($excelVersion) {
      $commandKey = "Registry::HKEY_CLASSES_ROOT\$excelVersion\shell\open\command"
      if (Test-Path $commandKey) {
        $commandValue = (Get-ItemProperty -Path $commandKey).'(default)'
        if ($commandValue -match '^\s*"([^"]+EXCEL\.EXE)"') { $excelPath = $Matches[1] }
        elseif ($commandValue -match '^\s*(.+?EXCEL\.EXE)') { $excelPath = $Matches[1].Trim('"') }
      }
    }
  }

  if (-not $excelPath) {
    $excelCmd = Get-Command EXCEL.EXE -ErrorAction SilentlyContinue
    if ($excelCmd) { $excelPath = $excelCmd.Source }
  }

  if (-not $excelPath -or -not (Test-Path -LiteralPath $excelPath)) {
    throw "Настольный Excel не найден или не зарегистрирован в Windows. Если Excel установлен, запустите его один раз через меню Пуск, закройте и повторите установку."
  }

  "Excel executable: $excelPath" | Add-Content -LiteralPath $logPath -Encoding UTF8
  Write-Host "`nМанифест скачан и проверен: $manifestPath" -ForegroundColor Green
  Write-Host "Открываю Excel..."
  Write-Host "В Excel выберите Главная → Надстройки → Дополнительно → Загрузить мою надстройку и укажите этот XML."
  Write-Host "Если команда загрузки отсутствует, ваша организация отключила sideload: передайте манифест администратору Microsoft 365 для централизованного развёртывания."
  Start-Process -FilePath $excelPath -ErrorAction Stop
  try {
    Set-Clipboard -Value $manifestPath -ErrorAction Stop
    Write-Host "Путь к XML скопирован в буфер обмена." -ForegroundColor Yellow
  } catch {
    Write-Host "Не получилось скопировать путь автоматически. Путь к файлу: $manifestPath" -ForegroundColor Yellow
  }
  "Excel launch requested successfully." | Add-Content -LiteralPath $logPath -Encoding UTF8
  Write-Host "`nExcel запущен. Проверьте, что появилось его окно, затем завершите загрузку манифеста."
  Write-Host "Установщик не менял настройки безопасности Office."
} catch {
  Show-InstallError $_.Exception.Message | Out-Null
}

Write-Host ""
Read-Host "Нажмите Enter, чтобы закрыть установщик"
