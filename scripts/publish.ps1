param(
    [string]$Configuration = "Release",
    [string]$Version = "1.2.0",
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$artifacts = Join-Path $root "artifacts"
$publishRoot = Join-Path $artifacts "publish"

if (-not $SkipTests) {
    dotnet test (Join-Path $root "OutlookMcp.sln") -c $Configuration
    if ($LASTEXITCODE -ne 0) { throw "Tests failed." }
}

foreach ($rid in @("win-x64", "win-x86")) {
    $destination = Join-Path $publishRoot $rid
    if (Test-Path -LiteralPath $destination) {
        Remove-Item -LiteralPath $destination -Recurse -Force
    }
    dotnet publish (Join-Path $root "src\OutlookMcp.Server\OutlookMcp.Server.csproj") -c $Configuration -r $rid --self-contained true -p:PublishSingleFile=true -p:Version=$Version -o $destination
    if ($LASTEXITCODE -ne 0) { throw "Publish failed for $rid." }
    Copy-Item (Join-Path $root "README.md") $destination -Force
    Copy-Item (Join-Path $root "src\OutlookMcp.Server\config.sample.json") $destination -Force
    Copy-Item (Join-Path $root "scripts\install-mcp.ps1") $destination -Force
    Copy-Item (Join-Path $root "examples") (Join-Path $destination "examples") -Recurse -Force
    $zip = Join-Path $artifacts "EULE-Outlook-MCP-$Version-$rid.zip"
    Compress-Archive -Path (Join-Path $destination "*") -DestinationPath $zip -Force
}

$iscc = Get-Command "ISCC.exe" -ErrorAction SilentlyContinue
if ($null -ne $iscc) {
    $installerDirectory = Join-Path $artifacts "installer"
    New-Item -ItemType Directory -Path $installerDirectory -Force | Out-Null
    & $iscc.Source "/DMyAppVersion=$Version" "/DSourceRoot=$root" "/DOutputRoot=$installerDirectory" (Join-Path $root "installer\OutlookMcp.iss")
    if ($LASTEXITCODE -ne 0) { throw "Inno Setup compilation failed." }
} else {
    Write-Warning "ISCC.exe was not found; portable ZIP releases were created, but the optional installer was skipped."
}

Write-Output "Release artifacts: $artifacts"
