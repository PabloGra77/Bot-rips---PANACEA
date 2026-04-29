param()

$ErrorActionPreference = "Stop"

function Test-SilverlightInstalled {
    $key1 = "HKLM:\SOFTWARE\Microsoft\Silverlight"
    $key2 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Silverlight"
    if (Test-Path $key1 -PathType Any) {
        return $true
    }
    if (Test-Path $key2 -PathType Any) {
        return $true
    }
    return $false
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$rootDir = Split-Path $scriptDir -Parent
function New-Shortcut {
    param(
        [Parameter(Mandatory=$true)][string]$LinkPath,
        [Parameter(Mandatory=$true)][string]$TargetPath,
        [string]$IconPath
    )
    try {
        $ws = New-Object -ComObject WScript.Shell
        $sc = $ws.CreateShortcut($LinkPath)
        $sc.TargetPath = $TargetPath
        if ($IconPath -and (Test-Path $IconPath)) {
            $sc.IconLocation = "$IconPath,0"
        }
        $sc.WorkingDirectory = (Split-Path $TargetPath -Parent)
        $sc.Description = "Panacea"
        $sc.Save()
    } catch {
    }
}

if (-not (Test-SilverlightInstalled)) {
    Write-Host "Silverlight no está instalado. Iniciando instalación si el instalador está disponible..."
    $installerCandidates = @()
    $installerCandidates += (Join-Path $scriptDir "Silverlight.exe")

    $possibleFolders = @("silverlth", "Silverligth")
    foreach ($folder in $possibleFolders) {
        $folderPath = Join-Path $scriptDir $folder
        if (Test-Path $folderPath) {
            $installerCandidates += (Join-Path $folderPath "Silverlight.exe")
            $exeAny = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($exeAny) {
                $installerCandidates += $exeAny.FullName
            }
        }
    }

    $localInstaller = $null
    foreach ($path in $installerCandidates) {
        if ($path -and (Test-Path $path)) {
            $localInstaller = $path
            break
        }
    }

    if ($localInstaller) {
        Write-Host "Ejecutando instalador local de Silverlight: $localInstaller"
        $process = Start-Process -FilePath $localInstaller -PassThru
        $process.WaitForExit()
        if (-not (Test-SilverlightInstalled)) {
            Write-Host "La instalación de Silverlight no se completó correctamente." -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "No se encontró Silverlight.exe en la carpeta. Abriendo página de descarga..." -ForegroundColor Yellow
        Start-Process "https://www.microsoft.com/getsilverlight"
        exit 0
    }
}

$vbsCandidates = @(
    (Join-Path $rootDir "PanaceaIE.vbs"),
    (Join-Path $rootDir "Configuracion\PanaceaIE.vbs")
)
$vbsPath = $null
foreach ($p in $vbsCandidates) {
    if (Test-Path $p) { $vbsPath = $p; break }
}

$htaPath = Join-Path $rootDir "Panacea.hta"

if ($vbsPath) {
    $iconPath = $null
    $preferredIcon = Join-Path $rootDir "Configuracion\ico\panacea.ico"
    if (Test-Path $preferredIcon) {
        $iconPath = $preferredIcon
    } else {
        $iconFolders = @(
            (Join-Path $rootDir "Configuracion\icon"),
            (Join-Path $rootDir "Configuracion\ico")
        )
        foreach ($iconFolder in $iconFolders) {
            if (-not (Test-Path $iconFolder)) { continue }
            $ico = Get-ChildItem -Path $iconFolder -Filter "*.ico" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($ico) {
                $iconPath = $ico.FullName
                break
            }
            $iconAlt = Get-ChildItem -Path $iconFolder -Filter "*.icon" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($iconAlt) {
                $iconPath = $iconAlt.FullName
                break
            }
        }
    }
    $desktop = [Environment]::GetFolderPath("Desktop")
    $startMenu = [Environment]::GetFolderPath("Programs")
    $desktopLink = Join-Path $desktop "Panacea.lnk"
    $menuLink = Join-Path $startMenu "Panacea.lnk"
    if (Test-Path $desktopLink) { Remove-Item $desktopLink -Force -ErrorAction SilentlyContinue }
    if (Test-Path $menuLink) { Remove-Item $menuLink -Force -ErrorAction SilentlyContinue }
    New-Shortcut -LinkPath $desktopLink -TargetPath $vbsPath -IconPath $iconPath
    New-Shortcut -LinkPath $menuLink -TargetPath $vbsPath -IconPath $iconPath
    Start-Process "wscript.exe" "`"$vbsPath`""
    exit 0
}
elseif (Test-Path $htaPath) {
    Start-Process "mshta.exe" $htaPath
    exit 0
}

$panaceaUrl = "http://181.51.196.194/Panacea"
$iePath = Join-Path $env:ProgramFiles "Internet Explorer\iexplore.exe"

if (Test-Path $iePath) {
    Start-Process $iePath $panaceaUrl
} else {
    Start-Process $panaceaUrl
}
