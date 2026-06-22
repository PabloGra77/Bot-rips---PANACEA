# Corrección para que el bot no quede pegado y sí se actualice por GitHub

## Causa encontrada

El ejecutable instalado compara la versión local contra el último release de GitHub:

- Local actual: `2.0.1`
- Release actual probable: `v2.0.1`

Si el release nuevo sigue siendo `v2.0.1`, el bot no descarga nada porque internamente evalúa que no hay una versión mayor.

## Qué se cambió

1. `MainForm.cs`
   - Se agregó fallback para `Buscar`, `Completar` y `Guardar`.
   - Si Panacea actualiza la grilla por AJAX y no dispara `DocumentCompleted`, el bot continúa después de unos segundos y procesa la grilla.

2. `Program.cs`
   - Si ya hay un `PanaceaIEWrapper.exe` abierto o bloqueado, ahora muestra un mensaje claro en vez de cerrarse en silencio.

3. `AutoUpdater.cs`
   - Ahora prioriza el asset ZIP de distribución llamado `RoBRips...zip`.
   - Esto evita que por error descargue un ZIP de código fuente y lo extraiga sobre la instalación.

4. `PanaceaIEWrapper.csproj`
   - Versión subida a `2.0.2`.

## Cómo publicar correctamente

### Opción A: desde Visual Studio / PowerShell en Windows

Ejecutar en la raíz del proyecto:

```powershell
.\PublicarRelease.ps1 -Version 2.0.2
```

Eso genera:

```text
dist\RoBRips-v2.0.2.zip
```

Luego en GitHub:

1. Ir a `Releases`.
2. Crear release nuevo con tag:

```text
v2.0.2
```

3. Subir como asset únicamente:

```text
RoBRips-v2.0.2.zip
```

4. Marcarlo como `Latest release`.

### Opción B: GitHub Actions

Este paquete incluye:

```text
.github/workflows/build-release.yml
```

Pasos:

1. Subir estos cambios al repo.
2. Ir a `Actions`.
3. Ejecutar `Build RoBRips Release`.
4. Usar versión `2.0.2`.

El workflow compila en Windows, crea el ZIP correcto y lo publica como release `v2.0.2`.

## Muy importante

No subir el ZIP completo del código fuente como asset del release.

El updater del bot espera un ZIP de distribución con estos archivos en la raíz:

```text
PanaceaIEWrapper.exe
PanaceaIEWrapper.exe.config
EPPlus.dll
bot-config.json
app.ico
```

Si se sube un ZIP con carpetas como `Configuracion`, `Docs`, `Herramientas`, etc., el autoupdater no instalará bien el bot.

## Después de publicar

En el equipo del usuario:

1. Abrir el bot instalado.
2. El bot debe detectar `v2.0.2`.
3. Aceptar actualización.
4. Se reemplaza el ejecutable en `%LOCALAPPDATA%\PanaceaRIPS`.

Si el bot viejo quedó abierto en segundo plano, cerrar primero `PanaceaIEWrapper.exe` desde el Administrador de tareas.
