$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dist = Join-Path $root "dist"
$project = Join-Path $root "KeywordOcr.App\KeywordOcr.App.csproj"

New-Item -ItemType Directory -Force -Path $dist | Out-Null

dotnet publish $project `
  --configuration Release `
  --runtime win-x64 `
  --self-contained true `
  --output $dist `
  -p:PublishSingleFile=true `
  -p:AssemblyName=KeywordOcr

$initialFiles = @(
  "app_settings.json",
  "user_stopwords.json",
  "cafe24_upload_config.txt"
)

foreach ($fileName in $initialFiles) {
  $source = Join-Path $root $fileName
  $target = Join-Path $dist $fileName
  if ((Test-Path -LiteralPath $source) -and -not (Test-Path -LiteralPath $target)) {
    Copy-Item -LiteralPath $source -Destination $target
  }
}

Write-Host "Built: $(Join-Path $dist 'KeywordOcr.exe')"
Write-Host "Keys are loaded from: $(Join-Path $env:USERPROFILE 'Desktop\key')"
