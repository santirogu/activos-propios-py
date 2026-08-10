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
  - `entrada/` queda al lado del `.exe` (mutable): es donde el usuario
    deja el `Formato_Dinamico.xlsm`. Se llama `entrada/` (no `resources/`)
    para no confundirla con la `resources/` interna del proyecto.
  - El logo se lee de `sys._MEIPASS/resources/` (embebido, read-only).
  - El `Formato_Dinamico.xlsm` se lee de `<EXE_DIR>/entrada/`
    (externo, editable por el usuario). Si esa carpeta no tiene ningún
    `.xlsm` en el primer arranque, `asegurar_formato_dinamico()` copia el
    bundleado como factory default. Dentro de `entrada/` debe haber UN
    solo `.xlsm`; si hay varios `validar_entrada_unica()` lo detecta.
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
# Carpeta externa/editable donde el usuario deja el `Formato_Dinamico.xlsm`.
# Se llama `entrada/` (no `resources/`) para no confundirla con la carpeta
# `resources/` interna del proyecto (logo, factory default, grabaciones VBS).
ENTRADA_DIR = PROJECT_ROOT / "entrada"


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

# Nombre canónico del Excel del formato dinámico. Es `.xlsm` (no `.xlsx`)
# porque el archivo contiene macros. Referenciado desde main.py y usado como
# nombre del factory default que se copia desde el bundle.
FORMATO_DINAMICO_NOMBRE = "Formato_Dinamico.xlsm"

# Advertencia mostrada al usuario cuando hay más de un `.xlsm` en `entrada/`.
MENSAJE_ENTRADA_MULTIPLE = (
    "Ten en cuenta que dentro de la carpeta «entrada» solo debe haber un "
    "archivo de Formato Dinámico. Se encontró más de un archivo .xlsm, lo "
    "que puede generar conflictos. Deja únicamente el archivo correcto y "
    "vuelve a intentarlo."
)


def listar_xlsm_entrada() -> list[Path]:
    """Lista los archivos `.xlsm` presentes en `entrada/` (orden alfabético).

    Devuelve `[]` si la carpeta no existe todavía. Es la fuente de verdad
    tanto para resolver el archivo a leer como para validar que haya uno
    y solo uno.
    """
    if not ENTRADA_DIR.exists():
        return []
    return sorted(ENTRADA_DIR.glob("*.xlsm"))


def formato_dinamico_path() -> Path:
    """Path al `Formato_Dinamico.xlsm` que la app debe leer.

    Apunta al archivo **externo** dentro de `entrada/` (al lado del `.exe`,
    o del repo en dev). Resolución:

    - Si existe el archivo con el nombre canónico, se prefiere ese.
    - Si no, pero hay algún otro `.xlsm`, se usa el primero (alfabético).
    - Si no hay ninguno, se devuelve el path canónico esperado (que el
      caller puede usar para reportar "archivo no encontrado").

    Cuando hay más de un `.xlsm` la elección es ambigua; usar
    `validar_entrada_unica()` para advertir al usuario antes de leer.
    """
    canonico = ENTRADA_DIR / FORMATO_DINAMICO_NOMBRE
    if canonico.exists():
        return canonico
    archivos = listar_xlsm_entrada()
    if archivos:
        return archivos[0]
    return canonico


def validar_entrada_unica() -> tuple[bool, str | None]:
    """Verifica que en `entrada/` haya como máximo un `.xlsm`.

    Returns:
        `(True, None)` si hay 0 o 1 archivo `.xlsm`.
        `(False, mensaje)` si hay 2 o más (situación de conflicto), donde
        `mensaje` es la advertencia lista para mostrar al usuario.
    """
    if len(listar_xlsm_entrada()) > 1:
        return False, MENSAJE_ENTRADA_MULTIPLE
    return True, None


def asegurar_formato_dinamico() -> tuple[Path, bool]:
    """Garantiza que exista un `.xlsm` en `entrada/` al lado del `.exe`.
    Si la carpeta no tiene ningún `.xlsm`, copia el bundleado como factory
    default (`Formato_Dinamico.xlsm`).

    Returns:
        (path al archivo externo, True si se acaba de crear)

    No sobrescribe si ya hay algún `.xlsm` (respeta el que el usuario dejó,
    aunque lo haya renombrado). No falla si el bundle tampoco lo tiene:
    devuelve el path canónico esperado para que el caller reporte el error.
    """
    # Si ya hay al menos un .xlsm, no copiamos nada: evita crear un segundo
    # archivo (que dispararía el conflicto de `validar_entrada_unica`).
    existentes = listar_xlsm_entrada()
    if existentes:
        return formato_dinamico_path(), False

    destino = ENTRADA_DIR / FORMATO_DINAMICO_NOMBRE
    origen = bundled_resource_path(f"resources/{FORMATO_DINAMICO_NOMBRE}")
    if not origen.exists():
        # Ni externo ni bundleado. Devolvemos el path esperado igual; el
        # caller (validador del flujo de extracción) reporta el error.
        return destino, False

    ENTRADA_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(origen, destino)
    return destino, True
