@echo off
setlocal enabledelayedexpansion
title Instalador Panacea
color 0A

:: ============================================================
:: Solicitar permisos de administrador si no los tenemos
:: ============================================================
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b 0
)

set "BASEDIR=%~dp0"

echo.
echo  =====================================================
echo    INSTALADOR PANACEA - Configurando equipo nuevo
echo  =====================================================
echo.

:: ============================================================
:: 1. .NET Framework 4.7.2
:: ============================================================
echo [1/4] Verificando .NET Framework 4.7.2...
set "NET_OK=0"
for /f "tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release 2^>nul') do (
    if %%a geq 461808 set "NET_OK=1"
)
if "!NET_OK!"=="1" (
    echo       OK - .NET 4.7.2 ya esta instalado.
) else (
    echo       Descargando .NET Framework 4.7.2 ^(puede tardar varios minutos^)...
    powershell -Command "(New-Object Net.WebClient).DownloadFile('https://go.microsoft.com/fwlink/?LinkId=863262','%TEMP%\ndp472.exe')"
    if exist "%TEMP%\ndp472.exe" (
        echo       Instalando .NET 4.7.2...
        start /wait "%TEMP%\ndp472.exe" /passive /norestart
        del "%TEMP%\ndp472.exe" >nul 2>&1
        echo       OK - .NET 4.7.2 instalado.
    ) else (
        echo       ADVERTENCIA: No se pudo descargar .NET 4.7.2.
        echo       Descargalo manualmente desde:
        echo       https://go.microsoft.com/fwlink/?LinkId=863262
    )
)

:: ============================================================
:: 2. Internet Explorer (necesario para el motor WebBrowser)
:: ============================================================
echo [2/4] Verificando Internet Explorer...
DISM /Online /Get-FeatureInfo /FeatureName:Internet-Explorer-Optional-amd64 2>nul | findstr /i "Disabled" >nul
if %errorlevel% equ 0 (
    echo       Habilitando Internet Explorer...
    DISM /Online /Enable-Feature /FeatureName:Internet-Explorer-Optional-amd64 /NoRestart
    echo       OK - Internet Explorer habilitado.
) else (
    echo       OK - Internet Explorer disponible.
)

:: ============================================================
:: 3. Silverlight (necesario para la pagina web de Panacea)
:: ============================================================
echo [3/4] Verificando Silverlight...
set "SL_OK=0"
reg query "HKLM\SOFTWARE\Microsoft\Silverlight" >nul 2>&1 && set "SL_OK=1"
reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Silverlight" >nul 2>&1 && set "SL_OK=1"
if "!SL_OK!"=="1" (
    echo       OK - Silverlight ya esta instalado.
) else (
    echo       Instalando Silverlight...
    set "SL_EXE="
    for /r "%BASEDIR%Configuracion\Silverligth" %%f in (*.exe) do (
        if not defined SL_EXE set "SL_EXE=%%f"
    )
    if defined SL_EXE (
        start /wait "!SL_EXE!" /q
        echo       OK - Silverlight instalado.
    ) else (
        echo       ADVERTENCIA: No se encontro el instalador de Silverlight en Configuracion\Silverligth\
    )
)

:: ============================================================
:: 4. Clave de registro IE11 para el control WebBrowser
:: ============================================================
echo [4/5] Configurando emulacion IE11 para WebBrowser...
set "REG_EMUL=HKLM\SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
reg add "%REG_EMUL%" /v "PanaceaIEWrapper.exe" /t REG_DWORD /d 11001 /f >nul 2>&1
set "REG_EMUL32=HKLM\SOFTWARE\WOW6432Node\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
reg add "%REG_EMUL32%" /v "PanaceaIEWrapper.exe" /t REG_DWORD /d 11001 /f >nul 2>&1
echo       OK - IE11 emulacion registrada.

:: ============================================================
:: 5. Acceso directo en el Escritorio
:: ============================================================
echo [5/5] Creando acceso directo en el Escritorio...
set "EXE_PATH=%BASEDIR%Configuracion\PanaceaIEWrapper\bin\Release\net472\PanaceaIEWrapper.exe"
set "EXE_DIR=%BASEDIR%Configuracion\PanaceaIEWrapper\bin\Release\net472"
(
    echo $ws = New-Object -ComObject WScript.Shell
    echo $desk = [Environment]::GetFolderPath('Desktop'^)
    echo $sc = $ws.CreateShortcut("$desk\Panacea.lnk"^)
    echo $sc.TargetPath = "%EXE_PATH%"
    echo $sc.WorkingDirectory = "%EXE_DIR%"
    echo $sc.Description = "Panacea RIPS Bot"
    echo $ico = "%BASEDIR%Configuracion\ico\panacea.ico"
    echo if (Test-Path $ico^) { $sc.IconLocation = "$ico,0" }
    echo $sc.Save(^)
    echo Write-Host "      OK - Acceso directo creado."
) > "%TEMP%\panacea_lnk.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "%TEMP%\panacea_lnk.ps1"
del "%TEMP%\panacea_lnk.ps1" >nul 2>&1

echo.
echo  =====================================================
echo   Instalacion completada correctamente.
echo.
echo   SIGUIENTE PASO:
echo   1. Edita las credenciales del usuario en:
echo      %BASEDIR%Configuracion\PanaceaIEWrapper\bin\Release\net472\bot-config.json
echo      Campos: "username" y "password"
echo.
echo   2. Copia BASE RIPS.xlsx a la carpeta:
echo      %BASEDIR%Configuracion\PanaceaIEWrapper\bin\Release\net472\base\
echo.
echo   3. Ejecuta el acceso directo "Panacea" del
echo      Escritorio o abre PanaceaIEWrapper.exe en:
echo      %BASEDIR%Configuracion\PanaceaIEWrapper\bin\Release\net472\
echo  =====================================================
echo.
pause
exit /b 0
