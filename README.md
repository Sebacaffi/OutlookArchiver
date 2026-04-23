# Outlook Archiver

Herramienta de archivado automático de correos Outlook para entornos empresariales. Monitorea el tamaño del buzón y mueve correos antiguos a archivos `.pst` organizados por año, corriendo en segundo plano sin intervención del usuario.

---

## ¿Cómo funciona?

El programa monitorea diariamente el tamaño del archivo `.ost` (buzón local de Outlook). Cuando supera el umbral configurado, mueve los correos más antiguos a un archivo `.pst` llamado `Archivo YYYY.pst`, donde `YYYY` corresponde al año de los correos archivados. El archivo aparece en Outlook con el nombre **"Archivo YYYY"**.

### Lógica de fecha de corte

La fecha de corte se calcula automáticamente como **el mismo día del mes, un mes antes de la ejecución**:

| Fecha de ejecución | Fecha de corte | Último día archivado | PST destino |
|---|---|---|---|
| 22/04/2026 | 22/03/2026 | 21/03/2026 | Archivo 2026.pst |
| 01/02/2027 | 01/01/2027 | 31/12/2026 | Archivo 2026.pst |
| 15/02/2027 | 15/01/2027 | 14/01/2027 | Archivo 2027.pst |

> El archivado usa `ReceivedTime < fecha_de_corte`, por lo que el último día archivado es siempre el día anterior a la fecha de corte.

### Cambio de año automático

Cuando la fecha de corte cae en enero del año nuevo (por ejemplo al ejecutar el 01/02/2027), el código detecta que el último día archivado corresponde al año anterior y dirige los correos al PST de ese año (`Archivo 2026.pst`). A partir del siguiente archivado comienza a poblar `Archivo 2027.pst`.

### Rotación de archivos PST

Los archivos `.pst` tienen un límite técnico de ~50 GB en Outlook. La herramienta gestiona esto automáticamente con un límite configurable (por defecto 47 GB):

- Si el PST del año actual ya supera el límite al iniciar → rota al año siguiente y opcionalmente copia el PST lleno a OneDrive.
- Si el PST se llena durante el archivado → se detiene, registra advertencia en el log y opcionalmente copia el PST lleno a OneDrive.

### Respaldo a OneDrive

Cuando está habilitado, al rotar un PST lleno el programa:

1. Cierra Outlook (necesario para que el archivo no esté bloqueado)
2. Copia el PST a `OneDrive - Agencia de Aduanas I.P. Hardy y Cía. Ltda\{Subcarpeta configurada}`
3. Reabre Outlook automáticamente
4. Continúa archivando en el nuevo PST

La detección de OneDrive es automática: busca la carpeta `OneDrive - Agencia de Aduanas I.P. Hardy y Cía. Ltda` en el perfil del usuario. Si no existe, OneDrive no está configurado en ese equipo y el respaldo se omite. El estado se muestra en la interfaz principal.

> **Nota:** Si el nombre de la carpeta de OneDrive cambia, actualizar la constante `ONEDRIVE_FOLDER_NAME` en `archiver.py`.

---

## Instalación para el usuario final

1. Ejecutar `OutlookArchiver.exe` — no requiere Python ni ninguna instalación previa.
2. Se abre el **wizard de configuración** con 3 pasos:
   - **Bienvenida** — resumen de la herramienta y preview de la próxima fecha de archivado.
   - **Configuración** — umbral del buzón, carpeta de PSTs, límite por archivo, hora de ejecución, inicio con Windows y opciones de respaldo a OneDrive.
   - **Listo** — resumen de la configuración aplicada.
3. Al finalizar el wizard se registran automáticamente:
   - La tarea `OutlookArchiverTask` en el Programador de Windows.
   - La entrada de inicio automático en el registro de Windows (`HKCU`).
4. Se abre la ventana principal. Al cerrarla, el programa permanece activo en la **bandeja del sistema**.

> No se requieren permisos de administrador en ningún paso.

---

## Uso diario

Una vez configurado, el programa funciona sin intervención. El ícono en la bandeja del sistema permite:

- **Abrir configuración** — abre la ventana principal (doble clic en el ícono).
- **Archivar ahora** — ejecuta el archivado inmediatamente.
- **Salir** — cierra completamente el programa (la tarea programada sigue activa).

---

## Ventana principal

### Tarjeta de estado

Muestra en tiempo real:

| Indicador | Descripción |
|---|---|
| Estado de la tarea | Verde si `OutlookArchiverTask` está registrada, rojo si no |
| Inicio con Windows | Activo / Inactivo según el registro `HKCU` |
| Tamaño buzón | Tamaño actual del `.ost`, en rojo si supera el umbral |
| OneDrive | Configurado (con nombre de carpeta) o No encontrado |

### Tarjeta de próximo archivado

Muestra la fecha de corte calculada para hoy, el PST de destino y su tamaño actual.

### Botones disponibles

| Botón | Acción |
|---|---|
| **Guardar y programar** | Guarda la configuración y registra/actualiza la tarea en el Programador de Windows |
| **Archivar ahora** | Ejecuta el archivado manualmente de inmediato |
| **Desactivar tarea** | Elimina la tarea del Programador de Windows sin desinstalar |
| **Ver log** | Abre el archivo de log con el historial de ejecuciones |
| **Reconfigurar** | Vuelve a lanzar el wizard de configuración inicial |
| **Desinstalar** | Elimina la tarea, el autostart y la configuración (los PST no se tocan) |

---

## Parámetros configurables

| Parámetro | Descripción | Valor por defecto |
|---|---|---|
| Umbral del buzón (GB) | Tamaño a partir del cual se activa el archivado | 3 GB |
| Carpeta de PSTs | Directorio donde se crean los archivos `Archivo YYYY.pst` | `Documentos\ArchivosOutlook` |
| Límite por PST (GB) | Tamaño máximo antes de rotar al año siguiente | 47 GB |
| Hora de ejecución | Hora diaria en que corre el archivado | 20:00 |
| Inicio con Windows | Lanza el programa al iniciar sesión | Activado |
| Backup a OneDrive | Copia el PST lleno a OneDrive al rotar | Desactivado |
| Subcarpeta en OneDrive | Carpeta dentro de OneDrive para los PSTs | `Respaldo Correo` |

---

## Archivos generados

| Ruta | Descripción |
|---|---|
| `%APPDATA%\OutlookArchiver\config.json` | Configuración local del usuario |
| `%APPDATA%\OutlookArchiver\archiver.log` | Log rotativo de ejecuciones (máx. 2 MB, 3 backups) |
| `[Carpeta configurada]\Archivo YYYY.pst` | Archivos de correos por año |
| `OneDrive\[Subcarpeta]\Archivo YYYY.pst` | Copia de respaldo (si OneDrive está habilitado) |

---

## Modos de ejecución

| Comando | Efecto |
|---|---|
| `OutlookArchiver.exe` | Primera vez: wizard. Ya configurado: abre GUI |
| `OutlookArchiver.exe --run` | Archivado silencioso (llamado por el Programador de Windows) |
| `OutlookArchiver.exe --setup` | Fuerza el wizard de configuración inicial |

---

## Verificar la tarea programada

```powershell
schtasks /Query /TN OutlookArchiverTask
```

---

## Desinstalación

### Opción 1 — Desde la interfaz
Abrir el programa → botón **"Desinstalar"** (esquina inferior derecha, en rojo).

Esto elimina:
- La tarea `OutlookArchiverTask` del Programador de Windows.
- La entrada de inicio automático del registro.
- El archivo `config.json`.

Los archivos `.pst` y el `.exe` **no se eliminan**.

### Opción 2 — Manual vía PowerShell

```powershell
# Eliminar tarea programada
schtasks /Delete /F /TN OutlookArchiverTask

# Eliminar inicio con Windows
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v OutlookArchiver /f

# Eliminar configuración (opcional)
Remove-Item "$env:APPDATA\OutlookArchiver" -Recurse -Force
```

Luego eliminar el `.exe` manualmente.

---

## Requisitos (solo para desarrollo)

- Windows 10/11
- Python 3.10+
- Microsoft Outlook instalado y configurado

```bash
pip install -r requirements.txt
```

---

## Build — generar el .exe distribuible

```bash
python build.py
```

El ejecutable queda en `dist/OutlookArchiver.exe`. Incluye Python y todas las dependencias — solo hay que distribuir ese único archivo a cada equipo.

---

## Estructura del proyecto

```
OutlookArchiver/
├── main.py           # Entrada: orquesta wizard, GUI o modo silencioso
├── wizard.py         # Wizard de configuración inicial (3 pasos)
├── gui.py            # Ventana principal + bandeja del sistema
├── archiver.py       # Lógica de archivado via COM de Outlook
├── config.py         # Lectura/escritura de configuración en JSON
├── scheduler.py      # Registro de tarea en el Programador de Windows
├── startup.py        # Registro de inicio automático (HKCU)
├── logger.py         # Logging rotativo a archivo
├── build.py          # Empaquetado con PyInstaller
├── requirements.txt  # Dependencias Python
├── .gitignore
└── README.md
```

---

## Despliegue en múltiples equipos

Distribuir el `.exe` por red, carpeta compartida, email o GPO. Cada usuario lo ejecuta una vez para completar la configuración local. No se requiere intervención del administrador.

Para preconfigurar valores por defecto de empresa (umbral, carpeta de PSTs, límite, subcarpeta OneDrive), editar el diccionario `DEFAULTS` en `config.py` antes de generar el build.

Para cambiar el nombre de la carpeta de OneDrive (si varía entre usuarios), editar la constante `ONEDRIVE_FOLDER_NAME` en `archiver.py` antes del build.