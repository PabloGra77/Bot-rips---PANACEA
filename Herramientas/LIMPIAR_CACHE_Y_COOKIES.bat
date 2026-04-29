@echo off
setlocal enabledelayedexpansion
title Limpiar Cache y Cookies - Panacea

echo ============================================
echo   LIMPIAR CACHE Y COOKIES - PANACEA
echo ============================================
echo.

:: -------------------------------------------------
:: 1. Cerrar PanaceaIEWrapper si esta abierto
:: -------------------------------------------------
echo Cerrando Panacea si esta abierto...
taskkill /F /IM PanaceaIEWrapper.exe >nul 2>&1
taskkill /F /IM CefSharp.BrowserSubprocess.exe >nul 2>&1
timeout /t 1 /nobreak >nul

:: -------------------------------------------------
:: 2. Cache y logs internos de Panacea (CefSharp)
:: -------------------------------------------------
echo Limpiando cache de Panacea (CefSharp)...
if exist "%LOCALAPPDATA%\Panacea\CefCache" (
    rd /s /q "%LOCALAPPDATA%\Panacea\CefCache" >nul 2>&1
    echo   [OK] CefCache eliminado.
) else (
    echo   [--] CefCache no encontrado.
)

if exist "%LOCALAPPDATA%\Panacea\logs" (
    rd /s /q "%LOCALAPPDATA%\Panacea\logs" >nul 2>&1
    echo   [OK] Logs eliminados.
) else (
    echo   [--] Logs no encontrados.
)

:: -------------------------------------------------
:: 3. Cache y cookies de Internet Explorer / Edge Legacy
:: -------------------------------------------------
echo Limpiando cache y cookies de IE / Edge Legacy...
rundll32.exe InetCpl.cpl,ClearMyTracksByProcess 255 >nul 2>&1
echo   [OK] Datos de IE/Edge Legacy limpiados.

:: -------------------------------------------------
:: 4. Archivos temporales de Internet (%TEMP% e INetCache)
:: -------------------------------------------------
echo Limpiando archivos temporales de Internet...
if exist "%LOCALAPPDATA%\Microsoft\Windows\INetCache" (
    del /f /s /q "%LOCALAPPDATA%\Microsoft\Windows\INetCache\*.*" >nul 2>&1
    echo   [OK] INetCache limpiado.
)
if exist "%LOCALAPPDATA%\Microsoft\Windows\Temporary Internet Files" (
    del /f /s /q "%LOCALAPPDATA%\Microsoft\Windows\Temporary Internet Files\*.*" >nul 2>&1
    echo   [OK] Temporary Internet Files limpiados.
)

:: -------------------------------------------------
:: 5. Cookies de Windows (IE/Edge)
:: -------------------------------------------------
echo Limpiando cookies de Windows (IE/Edge)...
if exist "%APPDATA%\Microsoft\Windows\Cookies" (
    del /f /s /q "%APPDATA%\Microsoft\Windows\Cookies\*.*" >nul 2>&1
    echo   [OK] Cookies eliminadas.
)
if exist "%LOCALAPPDATA%\Microsoft\Windows\INetCookies" (
    del /f /s /q "%LOCALAPPDATA%\Microsoft\Windows\INetCookies\*.*" >nul 2>&1
    echo   [OK] INetCookies eliminadas.
)

echo.
echo ============================================
echo   Limpieza completada correctamente.
echo ============================================
echo.
pause
endlocal
