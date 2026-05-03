$ErrorActionPreference = "Stop"

param(
    [switch]$Quiet
)

$appName = "KeywordOCR v3"
$installDir = Split-Path -Parent $MyInvocation.MyCommand.Path

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

$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "$appName.lnk"
if (Test-Path -LiteralPath $desktopShortcut) {
    Remove-Item -LiteralPath $desktopShortcut -Force
}

$startMenuDir = Join-Path ([Environment]::GetFolderPath("Programs")) $appName
if (Test-Path -LiteralPath $startMenuDir) {
    Remove-Item -LiteralPath $startMenuDir -Recurse -Force
}

$uninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\KeywordOcr"
if (Test-Path -LiteralPath $uninstallKey) {
    Remove-Item -LiteralPath $uninstallKey -Recurse -Force
}

$cleanupPath = Join-Path $env:TEMP ("KeywordOcr_cleanup_" + [Guid]::NewGuid().ToString("N") + ".ps1")
$escapedInstallDir = $installDir.Replace("'", "''")
$escapedCleanupPath = $cleanupPath.Replace("'", "''")
Set-Content -LiteralPath $cleanupPath -Encoding UTF8 -Value @"
Start-Sleep -Milliseconds 700
if (Test-Path -LiteralPath '$escapedInstallDir') {
    Remove-Item -LiteralPath '$escapedInstallDir' -Recurse -Force
}
Remove-Item -LiteralPath '$escapedCleanupPath' -Force -ErrorAction SilentlyContinue
"@

Start-Process -FilePath "powershell.exe" -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $cleanupPath) -WindowStyle Hidden

if (-not $Quiet) {
    Write-Host "$appName uninstalled."
}
