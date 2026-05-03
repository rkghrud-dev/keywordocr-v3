$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$project = Join-Path $root "KeywordOcr.App\KeywordOcr.App.csproj"
$setupProject = Join-Path $root "installer\SetupStub\KeywordOcr.Setup.csproj"
$distRoot = Join-Path $root "dist"
$appDist = Join-Path $distRoot "KeywordOcr"
$setupExe = Join-Path $distRoot "KeywordOcrSetup.exe"
$installerWork = Join-Path $distRoot "installer_work"
$setupStubDist = Join-Path $installerWork "setup_stub"
$payloadZip = Join-Path $installerWork "payload.zip"
$setupStubExe = Join-Path $setupStubDist "KeywordOcrSetup.exe"
$marker = [System.Text.Encoding]::ASCII.GetBytes("KOCRPAYLOADv1")

New-Item -ItemType Directory -Force -Path $distRoot | Out-Null
if (Test-Path -LiteralPath $appDist) {
    Remove-Item -LiteralPath $appDist -Recurse -Force
}
if (Test-Path -LiteralPath $installerWork) {
    Remove-Item -LiteralPath $installerWork -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $appDist | Out-Null
New-Item -ItemType Directory -Force -Path $installerWork | Out-Null

dotnet publish $project `
    --configuration Release `
    --runtime win-x64 `
    --self-contained true `
    --output $appDist `
    -p:PublishSingleFile=true `
    -p:AssemblyName=KeywordOcr

$initialFiles = @(
    "app_settings.json",
    "user_stopwords.json",
    "cafe24_upload_config.txt"
)

foreach ($fileName in $initialFiles) {
    $source = Join-Path $root $fileName
    $target = Join-Path $appDist $fileName
    if ((Test-Path -LiteralPath $source) -and -not (Test-Path -LiteralPath $target)) {
        Copy-Item -LiteralPath $source -Destination $target
    }
}

Copy-Item -LiteralPath (Join-Path $root "installer\uninstall.ps1") -Destination (Join-Path $appDist "uninstall.ps1") -Force

Compress-Archive -Path (Join-Path $appDist "*") -DestinationPath $payloadZip -Force

dotnet publish $setupProject `
    --configuration Release `
    --runtime win-x64 `
    --self-contained true `
    --output $setupStubDist `
    -p:PublishSingleFile=true

if (-not (Test-Path -LiteralPath $setupStubExe)) {
    throw "Setup stub was not built: $setupStubExe"
}

Copy-Item -LiteralPath $setupStubExe -Destination $setupExe -Force

$payloadBytes = [System.IO.File]::ReadAllBytes($payloadZip)
$lengthBytes = [System.BitConverter]::GetBytes([Int64]$payloadBytes.Length)
$out = [System.IO.File]::Open($setupExe, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write)
try {
    $out.Write($payloadBytes, 0, $payloadBytes.Length)
    $out.Write($lengthBytes, 0, $lengthBytes.Length)
    $out.Write($marker, 0, $marker.Length)
} finally {
    $out.Dispose()
}

Write-Host "Portable app: $appDist"
Write-Host "Installer: $setupExe"
