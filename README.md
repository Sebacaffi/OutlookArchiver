# Outlook Archiver

Herramienta de archivado automático de correos Outlook para entornos empresariales.
Genera un archivo `.pst` local con correos antiguos para mantener el buzón bajo control.

---

## Flujo completo para el usuario final

```
1. Ejecuta OutlookArchiver.exe
        ↓
2. Wizard de 3 pasos (solo la primera vez)
   · Bienvenida
   · Configuración (umbral, meses, horario, ruta .pst)
   · Confirmación + opción "Iniciar con Windows"
        ↓
3. Se registran automáticamente:
   · Tarea en el Programador de Windows (archivado diario)
   · Entrada en el registro para inicio con Windows
        ↓
4. Se abre la ventana principal de configuración
        ↓
5. A partir de ese momento todo corre en segundo plano
```

---

## Requisitos

- Windows 10/11
- Microsoft Outlook instalado y configurado
- Python 3.10+ (solo para desarrollo / build)

---

## Instalación para desarrollo

```bash
pip install -r requirements.txt
python main.py
```

---

## Generar el .exe distribuible

```bash
python build.py
```

El ejecutable queda en `dist/OutlookArchiver.exe`.
Distribuye ese único archivo — no requiere Python en el equipo destino.

---

## Modos de ejecución

| Comando | Efecto |
|---|---|
| `OutlookArchiver.exe` | Primera vez: wizard. Ya configurado: abre GUI |
| `OutlookArchiver.exe --run` | Archivado silencioso (llamado por el Programador de Windows) |
| `OutlookArchiver.exe --setup` | Fuerza el wizard de configuración inicial |

---

## Estructura del proyecto

```
outlook_archiver/
├── main.py        # Entrada: orquesta wizard, GUI o modo silencioso
├── wizard.py      # Wizard de 3 pasos para primera configuración
├── gui.py         # Ventana principal de configuración (tkinter)
├── archiver.py    # Lógica de archivado via COM de Outlook
├── config.py      # Configuración local en JSON (%APPDATA%\OutlookArchiver\)
├── scheduler.py   # Tarea en el Programador de Windows (schtasks)
├── startup.py     # Inicio automático con Windows (registro HKCU)
├── logger.py      # Logging rotativo
├── build.py       # Build con PyInstaller → .exe
├── requirements.txt
└── README.md
```

---

## Archivos generados en el equipo del usuario

| Ruta | Contenido |
|---|---|
| `%APPDATA%\OutlookArchiver\config.json` | Configuración del usuario |
| `%APPDATA%\OutlookArchiver\archiver.log` | Log de ejecuciones (rotativo 2MB) |
| Ruta elegida por usuario | Archivo `.pst` con correos archivados |

---

## Desinstalar / desactivar

Desde la ventana principal:
- **"Desactivar tarea"** → elimina la tarea del Programador de Windows
- Desmarcar **"Iniciar con Windows"** y guardar → elimina la entrada del registro

Para desinstalar completamente, borra el `.exe` y la carpeta `%APPDATA%\OutlookArchiver\`.

---

## Despliegue en múltiples equipos

Distribuye el `.exe` por red, email o GPO.  
Cada usuario lo ejecuta una vez — el wizard guía la configuración local.  
No requiere permisos de administrador.

Para preconfigurar valores por defecto de empresa, edita `DEFAULTS` en `config.py` antes de hacer el build.
