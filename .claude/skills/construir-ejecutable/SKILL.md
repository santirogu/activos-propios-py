---
name: construir-ejecutable
description: Empaqueta la app Tkinter en un .exe único usando PyInstaller, listo para entregar a usuarios NO técnicos. Maneja el bundling de resources/, el path detection (sys._MEIPASS vs __file__), y un build.spec reproducible.
---

# construir-ejecutable

Cuando el usuario invoque esta skill (típicamente al final de un ciclo de release), empaqueta el proyecto como un `.exe` standalone para Windows. El entregable final es un único archivo que el usuario doble-clica — sin Python instalado, sin pip, sin terminal.

## Pre-requisitos

### 1. Plataforma

PyInstaller **NO hace cross-compile**. Para generar el `.exe` de Windows hay que ejecutarse EN Windows.

```bash
# Detectar OS
python -c "import platform; print(platform.system())"
```

Si el output es:
- `Windows` → procede normal.
- `Darwin` (macOS) o `Linux` → AVISA al usuario que el build hay que hacerlo en una Windows. Puedes verificar la lógica del spec/script en estas plataformas pero el `.exe` generado no será válido para Windows.

### 2. Instalación de PyInstaller

```bash
.venv/bin/pip show pyinstaller >/dev/null 2>&1 || .venv/bin/pip install pyinstaller
```

## Ajustes al código ANTES del build (si no están aplicados)

PyInstaller hace que `__file__` apunte a un temp folder en `--onefile` mode. Esto rompe la convención del proyecto `PROJECT_ROOT = Path(__file__).resolve().parent.parent`. Hay que añadir un helper.

Edita `src/main.py` (cerca del tope, después de los imports):

```python
import sys
from pathlib import Path


def _resolver_project_root() -> Path:
    """Devuelve la carpeta donde están resources/ y salida/.

    - En modo dev (`python src/main.py`): el padre de `src/`.
    - En modo bundled (PyInstaller --onefile): la carpeta donde está
      el .exe (NO el temp `_MEIPASS` donde PyInstaller descomprime).
    """
    if getattr(sys, "frozen", False):
        # Bundled: usar la ruta del ejecutable
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


PROJECT_ROOT = _resolver_project_root()
```

Aplicar el MISMO ajuste en `src/sap_upload.py`, `src/sox_report.py`, `src/extraer_activos_creados.py`, `src/subir_anexos.py`, `src/branding.py` (cada uno tiene su propio `PROJECT_ROOT`).

Adicionalmente, los recursos bundleados (`resources/logo_hub_isa.png`, `resources/Formato_Dinamico_.xlsx`) hay que leerlos desde `sys._MEIPASS` cuando estén bundled. Helper:

```python
def _resource_path(rel_path: str) -> Path:
    """Path a un archivo bundled dentro del ejecutable."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / rel_path
    return PROJECT_ROOT / rel_path
```

Y usar `_resource_path("resources/logo_hub_isa.png")` en `branding.cargar_logo()`.

`salida/` SIEMPRE debe ir al lado del `.exe` (es output del usuario, mutable), así que `SALIDA_DIR = PROJECT_ROOT / "salida"` ya queda correcto con el helper de arriba.

## Spec file

Crea `GestionActivosFijos.spec` en la raíz del repo:

```python
# -*- mode: python ; coding: utf-8 -*-
"""Spec de PyInstaller para Gestión de Activos Fijos.

Generado por la skill construir-ejecutable. Editar manualmente si
añades dependencias nuevas con hidden imports.
"""

block_cipher = None

a = Analysis(
    ['src/main.py'],
    pathex=['src'],
    binaries=[],
    datas=[
        # Bundlear el logo y el Excel del formato dinámico
        ('resources/logo_hub_isa.png', 'resources'),
        ('resources/Formato_Dinamico_.xlsx', 'resources'),
        # Los VBS son referencia para devs, NO se bundlean
    ],
    hiddenimports=[
        # Pywin32 a veces necesita declarar submodules explícitamente
        'win32com.client',
        'pythoncom',
        # Pillow plugins que se cargan dinámicamente
        'PIL._tkinter_finder',
        # tkcalendar pulls babel/dateutil
        'babel.numbers',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Reducir tamaño excluyendo paquetes grandes que no usamos
        'matplotlib', 'numpy', 'pandas', 'scipy',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='GestionActivosFijos',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI app, no muestra consola
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='resources/logo_hub_isa.png',
)
```

## Comando de build

```bash
.venv/bin/pyinstaller GestionActivosFijos.spec --clean --noconfirm
```

- `--clean`: borra cache de builds previos.
- `--noconfirm`: no pregunta si sobrescribir el dist/.

El output queda en:
- `dist/GestionActivosFijos.exe` (típicamente ~80-120 MB).
- `build/` (artifacts intermedios, se puede borrar).

## Smoke test post-build

```bash
# Verificar tamaño razonable
ls -lh dist/GestionActivosFijos.exe

# En Windows: doble-clic en el .exe para abrirlo
# Verificar que:
# - La ventana se abre
# - El logo aparece (vino bundleado)
# - El menú principal muestra los 3 cards
# - La carpeta `salida/` se crea automáticamente al lado del .exe
# - Al menos "Extraer información en txt" funciona (no requiere SAP)
```

## Distribución

Recomendar al usuario:

1. **Compartir solo el `.exe`**. NO necesita compartir `resources/` ni nada más — todo está bundleado.

2. **El usuario final**:
   - Recibe `GestionActivosFijos.exe`.
   - Lo guarda en una carpeta (ej. `C:\Users\<user>\Documents\ActivosFijos\`).
   - Doble-clic → abre la app.
   - La primera ejecución crea `salida/` al lado del `.exe`.

3. **Antivirus warning**: algunos AV corporativos marcan ejecutables de PyInstaller como sospechosos (false positive). Mitigación: firmar el `.exe` con un certificado de la empresa. Sin firmar suele funcionar para uso interno con whitelisting.

4. **Actualizaciones**: cada release re-genera el `.exe`. El usuario solo lo reemplaza. Si guarda archivos en `salida/`, no se pierden.

## Pasos resumidos

1. Verificar plataforma Windows (avisar si no).
2. Instalar PyInstaller en venv si no está.
3. Aplicar (o verificar que ya están) los ajustes de `_resolver_project_root` y `_resource_path`.
4. Generar/actualizar `GestionActivosFijos.spec`.
5. Correr `pyinstaller GestionActivosFijos.spec --clean --noconfirm`.
6. Smoke test del `.exe`.
7. Reportar al usuario:
   - Path al `.exe`.
   - Tamaño.
   - Instrucciones de distribución.

## Qué NO hacer

- NO añadir `--onedir` por default — el proyecto pidió `.exe` único.
- NO incluir el código fuente (`src/`) como data — PyInstaller ya lo compila al binary.
- NO firmar el `.exe` con certificados propios sin pedir permiso — eso es decisión del equipo de seguridad de la empresa.
- NO subir el `.exe` al git — pesa demasiado (>80 MB). Confirmar que `dist/` está en `.gitignore`.
