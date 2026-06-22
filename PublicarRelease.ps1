param(
    [string]$Version = "2.0.2"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Project = Join-Path $Root "Configuracion\PanaceaIEWrapper\PanaceaIEWrapper.csproj"
$Publish = Join-Path $Root "publish"
$Dist = Join-Path $Root "dist"
$Zip = Join-Path $Dist "RoBRips-v$Version.zip"

Write-Host "== Limpiando salida anterior =="
if (Test-Path $Publish) { Remove-Item $Publish -Recurse -Force }
if (Test-Path $Dist) { Remove-Item $Dist -Recurse -Force }
New-Item -ItemType Directory -Path $Publish | Out-Null
New-Item -ItemType Directory -Path $Dist | Out-Null

Write-Host "== Compilando PanaceaIEWrapper v$Version =="
msbuild $Project /restore /p:Configuration=Release /p:Platform="Any CPU"

$Required = @(
    "PanaceaIEWrapper.exe",
    "PanaceaIEWrapper.exe.config",
    "EPPlus.dll",
    "bot-config.json",
    "app.ico"
)

foreach ($f in $Required) {
    $path = Join-Path $Publish $f
    if (!(Test-Path $path)) {
        throw "Falta archivo obligatorio en publish: $f"
    }
}

Write-Host "== Creando ZIP para GitHub Release =="
Push-Location $Publish
Compress-Archive -Path $Required -DestinationPath $Zip -Force
Pop-Location

Write-Host "OK: $Zip"
Write-Host "Sube este ZIP como asset del release v$Version en GitHub."
