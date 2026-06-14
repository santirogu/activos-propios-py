"""Helpers de rutas que funcionan tanto en modo dev como bundled (PyInstaller).

Centralizado para que `main`, `branding`, `sap_upload`, `sox_report`,
`extraer_activos_creados` y `subir_anexos` resuelvan paths de forma
consistente, y un solo punto cambia el comportamiento cuando se corre
desde el .exe en vez de `python src/main.py`.

Comportamiento por modo:

- **Dev mode** (`python src/main.py`):
  - `PROJECT_ROOT` = padre de `src/` (la raíz del repo).
  - Recursos en `<PROJECT_ROOT>/resources/`.
  - `salida/` en `<PROJECT_ROOT>/salida/`.

- **Bundled mode** (PyInstaller `--onefile`, `sys.frozen == True`):
  - `PROJECT_ROOT` = carpeta donde está el `.exe` (donde el usuario lo
    guardó). NO el temp `_MEIPASS`.
  - `salida/` queda al lado del `.exe` (mutable, output del usuario).
  - El logo se lee de `sys._MEIPASS/resources/` (embebido, read-only).
  - El `Formato_Dinamico_.xlsx` se lee de `<EXE_DIR>/resources/`
    (externo, editable por el usuario). Si no existe en el primer
    arranque, `asegurar_formato_dinamico()` lo copia desde el bundle
    como factory default.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path


# ---------------------------------------------------------------------------
# Resolución del project root
# ---------------------------------------------------------------------------

def project_root() -> Path:
    """Carpeta raíz donde viven `resources/` y `salida/`.

    En dev: padre de `src/`. En bundled: carpeta del `.exe`.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


PROJECT_ROOT = project_root()
SALIDA_DIR = PROJECT_ROOT / "salida"
RESOURCES_DIR = PROJECT_ROOT / "resources"


# ---------------------------------------------------------------------------
# Recursos bundleados (siempre read-only)
# ---------------------------------------------------------------------------

def bundled_resource_path(rel_path: str) -> Path:
    """Path a un archivo bundleado dentro del `.exe`, o su equivalente
    en el repo cuando estamos en dev.

    El path es **read-only**: no escribir aquí, porque `sys._MEIPASS` es
    una carpeta temporal que PyInstaller borra al cerrar la app.

    Para archivos que el usuario puede editar (ej. `Formato_Dinamico_`)
    usar `asegurar_formato_dinamico()` que copia del bundle a la carpeta
    externa al lado del `.exe`.
    """
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / rel_path
    return PROJECT_ROOT / rel_path


# ---------------------------------------------------------------------------
# Archivos externos editables
# ---------------------------------------------------------------------------

# Nombre del Excel del formato dinámico, referenciado desde main.py.
FORMATO_DINAMICO_NOMBRE = "Formato_Dinamico_.xlsx"


def formato_dinamico_path() -> Path:
    """Path al `Formato_Dinamico_.xlsx` que la app debe leer.

    Siempre apunta al archivo **externo** al lado del `.exe` (o del
    repo en dev). Llamar `asegurar_formato_dinamico()` antes para
    garantizar que existe.
    """
    return RESOURCES_DIR / FORMATO_DINAMICO_NOMBRE


def asegurar_formato_dinamico() -> tuple[Path, bool]:
    """Garantiza que `Formato_Dinamico_.xlsx` exista en `resources/`
    al lado del `.exe`. Si no existe, lo copia desde el bundle (factory
    default).

    Returns:
        (path al archivo externo, True si se acaba de crear)

    No falla si el bundle tampoco lo tiene: devuelve el path esperado
    para que el caller decida cómo reportar el error al usuario (típico:
    en dev mode si alguien borró el archivo del repo).
    """
    destino = formato_dinamico_path()
    if destino.exists():
        return destino, False

    origen = bundled_resource_path(f"resources/{FORMATO_DINAMICO_NOMBRE}")
    if not origen.exists():
        # Ni externo ni bundleado. Devolvemos el path esperado igual; el
        # caller (validador del flujo de extracción) reporta el error.
        return destino, False

    RESOURCES_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(origen, destino)
    return destino, True
