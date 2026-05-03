$ErrorActionPreference = "Stop"

param(
    [switch]$NoRun
)

$appName = "KeywordOCR v3"
$installDir = Join-Path $env:LOCALAPPDATA "KeywordOcr"
$payloadZip = Join-Path $PSScriptRoot "payload.zip"

if (-not (Test-Path -LiteralPath $payloadZip)) {
    throw "payload.zip not found: $payloadZip"
}

$running = Get-Process -Name "KeywordOcr" -ErrorAction SilentlyContinue
foreach ($proc in $running) {
    try {
        $path = $proc.MainModule.FileName
        if ($path -and $path.StartsWith($installDir, [System.StringComparison]::OrdinalIgnoreCase)) {
            Stop-Process -Id $proc.Id -Force
        }
    } catch {
    }
}

New-Item -ItemType Directory -Force -Path $installDir | Out-Null
Expand-Archive -LiteralPath $payloadZip -DestinationPath $installDir -Force

$exePath = Join-Path $installDir "KeywordOcr.exe"
if (-not (Test-Path -LiteralPath $exePath)) {
    throw "KeywordOcr.exe not found after install: $exePath"
}

function New-AppShortcut([string]$shortcutPath) {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $exePath
    $shortcut.WorkingDirectory = $installDir
    $shortcut.IconLocation = "$exePath,0"
    $shortcut.Description = $appName
    $shortcut.Save()
}

$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "$appName.lnk"
New-AppShortcut $desktopShortcut

$startMenuDir = Join-Path ([Environment]::GetFolderPath("Programs")) $appName
New-Item -ItemType Directory -Force -Path $startMenuDir | Out-Null
New-AppShortcut (Join-Path $startMenuDir "$appName.lnk")

$uninstallScript = Join-Path $installDir "uninstall.ps1"
$uninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\KeywordOcr"
New-Item -Path $uninstallKey -Force | Out-Null
Set-ItemProperty -Path $uninstallKey -Name "DisplayName" -Value $appName
Set-ItemProperty -Path $uninstallKey -Name "DisplayVersion" -Value "3.0"
Set-ItemProperty -Path $uninstallKey -Name "Publisher" -Value "rkghrud-dev"
Set-ItemProperty -Path $uninstallKey -Name "InstallLocation" -Value $installDir
Set-ItemProperty -Path $uninstallKey -Name "DisplayIcon" -Value "$exePath,0"
Set-ItemProperty -Path $uninstallKey -Name "UninstallString" -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$uninstallScript`""
Set-ItemProperty -Path $uninstallKey -Name "QuietUninstallString" -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$uninstallScript`" -Quiet"
Set-ItemProperty -Path $uninstallKey -Name "NoModify" -Value 1 -Type DWord
Set-ItemProperty -Path $uninstallKey -Name "NoRepair" -Value 1 -Type DWord

$estimatedSize = [int]((Get-ChildItem -LiteralPath $installDir -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1KB)
Set-ItemProperty -Path $uninstallKey -Name "EstimatedSize" -Value $estimatedSize -Type DWord

Write-Host "$appName installed to $installDir"

if (-not $NoRun) {
    Start-Process -FilePath $exePath -WorkingDirectory $installDir
}
