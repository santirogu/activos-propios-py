# -*- mode: python ; coding: utf-8 -*-
"""Spec de PyInstaller para Gestión de Activos Fijos.

Genera `GestionActivosFijos.exe` en Windows (`--onefile`, sin consola).

Build:
    pyinstaller GestionActivosFijos.spec --clean --noconfirm

Output:
    dist/GestionActivosFijos.exe   (~80–120 MB típico, depende de upx)
    build/                         (artefactos intermedios; borrable)

Importante:
    - PyInstaller NO hace cross-compile. Para producir el `.exe` real
      hay que correr este spec EN Windows. Si lo corres en macOS/Linux
      vas a obtener un binario nativo de esa plataforma — útil sólo
      como dry-run para validar que el spec parsea y los `datas` /
      `hiddenimports` resuelven.
    - El usuario final solo necesita el `.exe`. El `Formato_Dinamico.xlsm`
      viene bundleado como factory default y la app lo extrae a
      `<EXE_DIR>/entrada/` en el primer arranque para que el usuario
      lo pueda editar después (ver `paths.asegurar_formato_dinamico`).
      La carpeta se llama `entrada/` (no `resources/`) para no confundirla
      con la `resources/` interna del proyecto.
    - El logo se lee como recurso bundled (read-only) desde
      `sys._MEIPASS` vía `paths.bundled_resource_path` (ver
      `branding.LOGO_PATH`).
    - `salida/` siempre se crea/lee al lado del `.exe` (es output del
      usuario, mutable), gracias a `paths.project_root()` que detecta
      `sys.frozen` y devuelve la carpeta del ejecutable.
"""

import sys as _sys


block_cipher = None


# El icono PNG solo lo respeta PyInstaller en Windows (lo convierte a
# .ico internamente). En macOS exigiría un .icns y en Linux se ignora.
# Lo dejamos None fuera de Windows para que el dry-run no falle por un
# tipo de archivo no soportado.
_icon = "resources/logo_hub_isa.png" if _sys.platform == "win32" else None


a = Analysis(
    ["src/main.py"],
    pathex=["src"],
    binaries=[],
    datas=[
        # Logo: read-only, bundleado en sys._MEIPASS. Branding.LOGO_PATH
        # lo resuelve via paths.bundled_resource_path.
        ("resources/logo_hub_isa.png", "resources"),
        # Formato dinámico: factory default (.xlsm con macros). La app lo
        # copia a <EXE_DIR>/entrada/ la primera vez (asegurar_formato_dinamico)
        # y a partir de ahí lee del externo editable. Se bundlea dentro de
        # `resources/` del _MEIPASS (read-only); la carpeta externa editable
        # es `entrada/`.
        ("resources/Formato_Dinamico.xlsm", "resources"),
        # Los .vbs de resources/ NO se bundlean: son referencia para
        # devs (grabaciones SAP) y no las necesita el runtime.
    ],
    hiddenimports=[
        # pywin32: PyInstaller a veces no detecta submódulos COM cuando
        # se cargan vía win32com.client.GetObject(...).
        "win32com.client",
        "pythoncom",
        # Pillow ImageTk se carga por un plugin que el análisis estático
        # pierde.
        "PIL._tkinter_finder",
        # tkcalendar arrastra babel para formatear meses/días al locale.
        "babel.numbers",
        "babel.dates",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Reducción de tamaño: paquetes pesados que el proyecto NO usa
        # pero que a veces se cuelan por análisis transitivo de Pillow
        # o openpyxl. Si alguno se importa accidentalmente el .exe
        # fallará en runtime con ImportError; quitarlo de aquí en ese
        # caso.
        "matplotlib",
        "numpy",
        "pandas",
        "scipy",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)


pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)


# Splash screen: imagen que el bootloader muestra INMEDIATAMENTE al doble-
# clic, antes de descomprimir el bundle de --onefile. Para el usuario eso
# es la diferencia entre "no pasa nada" (5–8 s en frío) y "ya está
# cargando". El runtime cierra el splash desde `main._cerrar_splash()`
# cuando la ventana Tk ya está lista.
#
# IMPORTANTE: PyInstaller solo soporta Splash en Windows y Linux. En macOS
# el build falla si se incluye, así que el dry-run salta el splash y arma
# el EXE sin él. En el build real en Windows el splash sí entra.
_soporta_splash = _sys.platform in ("win32", "linux")

if _soporta_splash:
    splash = Splash(
        "resources/logo_hub_isa.png",
        binaries=a.binaries,
        datas=a.datas,
        text_pos=(10, 260),       # x,y dentro de la imagen (469x286)
        text_size=10,
        text_color="#1A3A6C",     # ISA_AZUL — coincide con branding.py
        text_default="Cargando Gestión de Activos Fijos…",
        always_on_top=True,
    )
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        splash,
        splash.binaries,
        [],
        name="GestionActivosFijos",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        disable_windowed_traceback=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=_icon,
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        [],
        name="GestionActivosFijos",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        disable_windowed_traceback=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=_icon,
    )
