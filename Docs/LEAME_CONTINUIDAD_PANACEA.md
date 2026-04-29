# LEAME DE CONTINUIDAD - PANACEA

Fecha: 2026-04-23

## Objetivo del proyecto
Automatizar en IE el flujo de Panacea:
1. Login automático en http://181.51.196.194/Panacea
2. Después del login: seleccionar sede "SEDE ADMINISTRATIVA CALLE 74"
3. Seleccionar contingencia "FACTURACION"
4. Clic en "Aceptar"

## Estado actual
- Se corrigieron errores de compilación en MainForm.cs (bloques corruptos y llaves mal anidadas).
- El proyecto compila correctamente en Release.
- El llenado de credenciales para controles Infragistics se mejoró:
  - Campo visible *_I
  - Campo oculto *$CVS
- Se verificó en log que antes de enviar:
  - UserNameTextBox$CVS contiene el usuario
  - PasswordTextBox$CVS contiene la contraseña
- También se agregaron coordenadas ImageButton1.x / ImageButton1.y antes del submit.

## Hallazgo clave de hoy
La plataforma estaba en mantenimiento / con conectividad caída.
Evidencia:
- Test-NetConnection 181.51.196.194:80 => TcpTestSucceeded = False
- Invoke-WebRequest a /Panacea => "No es posible conectar con el servidor remoto"
- Captura de IE muestra "Internet Explorer no puede mostrar la página web"

Conclusión:
- El fallo principal observado hoy no es de lógica del bot sino de disponibilidad de plataforma/red.

## Archivos importantes
- Wrapper principal: Configuracion/PanaceaIEWrapper/MainForm.cs
- Entrada WinForms: Configuracion/PanaceaIEWrapper/Program.cs
- Lanzador VBS: PanaceaIE.vbs
- Log de ejecución: %LOCALAPPDATA%\Panacea\ui-autologin.log

## Cómo retomar cuando el servidor vuelva
1. Cerrar instancias abiertas:
   - Get-Process PanaceaIEWrapper -ErrorAction SilentlyContinue | Stop-Process -Force
2. Compilar:
   - dotnet build .\Configuracion\PanaceaIEWrapper.sln -c Release
3. Ejecutar con credenciales:
   - $env:PANACEA_USER='1016018747'
   - $env:PANACEA_PASS='Gole2025'
   - Start-Process .\Configuracion\PanaceaIEWrapper\bin\Release\net472\PanaceaIEWrapper.exe
4. Revisar resultados en log:
   - Buscar: "Sede seleccionada", "Contingencia resultado", "Aceptar resultado"

## Criterio de éxito al retomar
Debe aparecer en log:
- Sede seleccionada: SEDE ADMINISTRATIVA CALLE 74
- Contingencia resultado: CONT_SET (o CONT_IG_SET)
- Aceptar resultado: *CLICKED*
- Sin timeout en SelectSedeAndContingencia

## Nota
Si vuelve a aparecer rebote a login cuando ya no haya mantenimiento,
revisar nuevamente el mensaje real post-login (tooltip/mensaje server-side) porque podría ser política de autenticación del servidor.
