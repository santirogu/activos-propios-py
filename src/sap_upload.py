"""
sap_upload.py — Carga del .txt generado en salida/ a SAP vía LSMW.

Replica los pasos grabados en `resources/script_sap_base.txt` (flujo LSMW
completo) y en `resources/Script1.vbs` (configuración dinámica del archivo
de entrada en el paso "Specify Files") usando SAP GUI Scripting desde
Python (vía pywin32).

REQUISITOS DE EJECUCIÓN
=======================
Sistema operativo: Windows (única plataforma soportada por SAP GUI Scripting).

Dependencias Python:
    pip install pywin32

Configuración SAP (una sola vez por máquina):
1. Cliente — habilitar scripting en SAP GUI:
     Options > Accessibility & Scripting > Scripting > "Enable scripting"
   Recomendado desmarcar "Notify when a script attaches to SAP GUI" y
   "Notify when a script opens a connection".
2. Servidor — parámetro `sapgui/user_scripting = TRUE` (transacción RZ11).
3. El usuario debe haber iniciado sesión en SAP **antes** de correr este
   script. Este script NO autentica.
4. La transacción LSMW debe tener pre-cargado el proyecto/subproyecto/objeto
   correctos (basta con haberlos abierto manualmente al menos una vez en
   la sesión actual de SAP).

USO
===
    python src/sap_upload.py

El script:
1. Toma el .txt más reciente de `salida/`.
2. Conecta a la sesión SAP abierta.
3. Ejecuta el flujo LSMW: configura la ruta del archivo en "Specify Files"
   apuntando al .txt de salida/, luego Assign Files → Read Data →
   Display Read Data → Convert Data → Display Converted Data →
   Create Batch Input Session → Run Batch Input Session.
4. Procesa la sesión BDC creada (modo error, log completo, experto).
"""

from __future__ import annotations

import sys
import time
from pathlib import Path


# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SALIDA_DIR = PROJECT_ROOT / "salida"

# Identificadores tomados de las grabaciones VBS. Cambiar solo si el proyecto
# LSMW tiene un step list distinto.
LSMW_STEPLIST_TABLE = "wnd[0]/usr/tbl/SAPDMC/SAPLLSMW_OBJ_000TC_STEPLIST"
DEFAULT_SELECTED_ROW = 13  # SAP marca esta fila por default; deseleccionar primero.

# Filas del step list que requieren selección manual antes de pulsar Execute
# (F6 = btn[32]). Los pasos posteriores se ejecutan secuencialmente porque el
# cursor avanza solo después de cada Execute.
SPECIFY_FILES_ROW = 6
ASSIGN_FILES_ROW = 7
READ_DATA_ROW = 8

BDC_SESSION_TABLE = (
    "wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/"
    "ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI"
)


# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------

def _ejecutar(descripcion: str, fn, *args, **kwargs):
    """Ejecuta `fn(*args, **kwargs)` loguenado la operación. Si falla,
    re-lanza con un mensaje descriptivo que dice exactamente qué intentaba
    hacer — clave porque las excepciones COM del SAP Frontend Server
    suelen venir con descripción vacía o genérica."""
    _log(f"  → {descripcion}")
    try:
        return fn(*args, **kwargs)
    except Exception as exc:
        raise RuntimeError(
            f"Falló: {descripcion}\n"
            f"Detalle técnico SAP: {exc!r}"
        ) from exc


def _log(mensaje: str) -> None:
    """Imprime un mensaje con timestamp para seguimiento de la ejecución.

    Usa flush=True para que el log aparezca en tiempo real aunque stdout
    esté redirigido o la consola lo esté bufferizando.
    """
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {mensaje}", flush=True)


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------

def get_latest_txt(salida_dir: Path = SALIDA_DIR) -> Path:
    """Devuelve el archivo `LSMW_*.txt` más reciente en `salida/`."""
    if not salida_dir.exists():
        raise FileNotFoundError(
            f"No existe la carpeta {salida_dir}. "
            f"Ejecuta primero `python src/main.py` para generar el .txt."
        )
    txts = sorted(salida_dir.glob("LSMW_*.txt"), key=lambda p: p.stat().st_mtime)
    if not txts:
        raise FileNotFoundError(
            f"No hay archivos LSMW_*.txt en {salida_dir}. "
            f"Ejecuta primero `python src/main.py`."
        )
    return txts[-1]


def get_sap_session():
    """Conecta al SAP GUI Scripting Engine y devuelve la primera sesión activa.

    Raises:
        RuntimeError: si pywin32 no está instalado, SAP GUI no está corriendo,
            no hay conexión activa o no hay sesión iniciada.
    """
    _log("get_sap_session: importando win32com.client...")
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        _log(f"get_sap_session: ImportError → {exc!r}")
        raise RuntimeError(
            "Falta la dependencia pywin32. Instalar con: pip install pywin32"
        ) from exc

    _log("get_sap_session: llamando win32com.client.GetObject('SAPGUI')...")
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        _log("get_sap_session: GetObject('SAPGUI') OK")
    except Exception as exc:
        _log(f"get_sap_session: GetObject FALLÓ — {exc!r}")
        raise RuntimeError(
            "No se pudo conectar a SAP GUI. Verifica:\n"
            "  - SAP GUI for Windows está abierto y con sesión iniciada.\n"
            "  - SAP GUI Scripting habilitado en Options del cliente.\n"
            "  - sapgui/user_scripting = TRUE en el servidor SAP."
        ) from exc

    application = sap_gui_auto.GetScriptingEngine
    num_conex = application.Children.Count
    _log(f"get_sap_session: conexiones detectadas = {num_conex}")
    if num_conex == 0:
        raise RuntimeError("No hay conexiones SAP activas en este SAP GUI.")
    connection = application.Children(0)
    num_ses = connection.Children.Count
    _log(f"get_sap_session: sesiones en conexión[0] = {num_ses}")
    if num_ses == 0:
        raise RuntimeError(
            "No hay sesiones activas en la conexión SAP. "
            "Inicia sesión en el sistema SAP antes de correr este script."
        )
    _log("get_sap_session: devolviendo sesión[0][0]")
    return connection.Children(0)


def diagnosticar_conexion_sap() -> tuple[bool, str]:
    """Verifica el estado de la conexión SAP GUI sin ejecutar ningún flujo.

    Útil como botón "Test conexión SAP" para diagnosticar cuando los
    botones de carga (Subir a SAP / Generar Reporte SOX) fallan al
    conectar. Reporta exactamente qué encontró:
      - pywin32 ausente
      - SAP GUI no abierto / COM inaccesible
      - Scripting Engine deshabilitado
      - Sin conexiones activas / sin sesiones iniciadas
      - Conexiones y sesiones detectadas (con detalle de cada una)

    Returns:
        Tupla (ok, mensaje). `ok=True` si hay al menos una sesión SAP
        activa lista para usarse; en caso contrario `ok=False` con un
        mensaje detallado y accionable.
    """
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        return False, (
            "Falta la dependencia pywin32.\n\n"
            "Instalar con:\n"
            "    pip install pywin32\n\n"
            f"Detalle: {exc}"
        )

    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
    except Exception as exc:
        return False, (
            "No se pudo acceder al objeto COM 'SAPGUI'.\n\n"
            "Causas más probables:\n"
            "  • SAP GUI for Windows no está abierto en este momento.\n"
            "  • SAP GUI se cerró/crasheó silenciosamente.\n"
            "  • Python y SAP GUI corren en contextos distintos\n"
            "    (uno como administrador y el otro no — COM no los une).\n\n"
            f"Detalle técnico: {exc!r}"
        )

    try:
        application = sap_gui_auto.GetScriptingEngine
    except Exception as exc:
        return False, (
            "SAP GUI está corriendo pero el Scripting Engine no responde.\n\n"
            "Verifica en SAP GUI:\n"
            "  Options → Accessibility & Scripting → Scripting →\n"
            "  'Enable scripting' debe estar marcado.\n\n"
            f"Detalle: {exc!r}"
        )

    num_conexiones = application.Children.Count
    if num_conexiones == 0:
        return False, (
            "SAP GUI está corriendo y el scripting habilitado, pero NO\n"
            "hay ninguna conexión activa al servidor SAP.\n\n"
            "Abre una conexión desde el SAP Logon Pad."
        )

    lineas = [f"Conexiones SAP detectadas: {num_conexiones}"]
    total_sesiones = 0

    for i in range(num_conexiones):
        try:
            connection = application.Children(i)
            num_sesiones = connection.Children.Count
            total_sesiones += num_sesiones
            lineas.append(f"  Conexión [{i}]: {num_sesiones} sesión(es)")

            for j in range(num_sesiones):
                try:
                    session = connection.Children(j)
                    info = session.Info
                    lineas.append(
                        f"    Sesión [{j}]: "
                        f"sistema={info.SystemName}, "
                        f"client={info.Client}, "
                        f"user={info.User}"
                    )
                except Exception:
                    lineas.append(f"    Sesión [{j}]: (info no disponible)")
        except Exception as exc:
            lineas.append(f"  Conexión [{i}]: (no se pudo leer — {exc!r})")

    if total_sesiones == 0:
        return False, "\n".join(
            lineas
            + [
                "",
                "ATENCIÓN: hay conexión pero NO hay sesiones iniciadas.",
                "Inicia sesión en el sistema SAP antes de correr los flujos.",
            ]
        )

    return True, "\n".join(
        lineas
        + [
            "",
            f"OK — SAP está accesible. {total_sesiones} sesión(es) lista(s).",
            "Los botones 'Subir a SAP' y 'Generar Reporte SOX' pueden ejecutarse.",
        ]
    )


# ---------------------------------------------------------------------------
# FLUJO LSMW
# ---------------------------------------------------------------------------

def select_step_row(session, row: int) -> None:
    """Selecciona una fila concreta del step list, deseleccionando la default."""
    table = _ejecutar(
        f"Localizar tabla del step list ({LSMW_STEPLIST_TABLE})",
        session.findById, LSMW_STEPLIST_TABLE,
    )
    _ejecutar(
        f"Deseleccionar fila default [{DEFAULT_SELECTED_ROW}]",
        lambda: setattr(
            table.getAbsoluteRow(DEFAULT_SELECTED_ROW), "selected", False
        ),
    )
    _ejecutar(
        f"Seleccionar fila objetivo [{row}]",
        lambda: setattr(table.getAbsoluteRow(row), "selected", True),
    )
    cell_id = f"{LSMW_STEPLIST_TABLE}/txtGT_STEPLIST-STEPTEXT[0,{row}]"
    cell = _ejecutar(
        f"Localizar celda del paso [{row}] ({cell_id})",
        session.findById, cell_id,
    )
    _ejecutar(f"Foco en celda del paso [{row}]", cell.setFocus)
    _ejecutar(
        f"Cursor en celda del paso [{row}] (caretPosition=0)",
        lambda: setattr(cell, "caretPosition", 0),
    )


def open_lsmw(session) -> None:
    """Abre la T-code LSMW y ejecuta el proyecto pre-cargado."""
    _log("Paso 1/10: Abriendo transacción LSMW y proyecto pre-cargado...")
    wnd = _ejecutar(
        "Localizar ventana principal wnd[0]",
        session.findById, "wnd[0]",
    )
    _ejecutar("Maximizar ventana principal", wnd.maximize)
    okcd = _ejecutar(
        "Localizar casilla de comandos (wnd[0]/tbar[0]/okcd)",
        session.findById, "wnd[0]/tbar[0]/okcd",
    )
    _ejecutar(
        "Escribir T-code 'LSMW' en okcd",
        lambda: setattr(okcd, "text", "LSMW"),
    )
    _ejecutar("Enviar Enter (sendVKey 0)", wnd.sendVKey, 0)
    # F8 — entra al step list del proyecto pre-cargado
    boton_f8 = _ejecutar(
        "Localizar botón Ejecutar (F8 = wnd[0]/tbar[1]/btn[8])",
        session.findById, "wnd[0]/tbar[1]/btn[8]",
    )
    _ejecutar("Pulsar Ejecutar (F8) para abrir el step list", boton_f8.press)


def configurar_ruta_archivo(session, carpeta: str, nombre_archivo: str) -> None:
    """Configura dinámicamente la ruta del archivo en el paso "Specify Files".

    Replica `resources/Script1.vbs`. Entra al paso, edita la definición de
    archivo apuntándola a `carpeta/nombre_archivo`, guarda y vuelve al step
    list. Reemplaza la configuración manual previa de SAP_LSMW_INPUT_PATH.

    Args:
        session: sesión SAP GUI.
        carpeta: ruta absoluta de la carpeta (ej. r"C:\\Users\\xxx\\salida").
        nombre_archivo: nombre del archivo (ej. "LSMW_20260510_094838.txt").
    """
    _log(f"Paso 2/10: Configurando ruta del archivo en LSMW → {carpeta}\\{nombre_archivo}")
    # Foco en la celda del paso "Specify Files" (row 6) y F2 para abrirlo
    cell_id = f"{LSMW_STEPLIST_TABLE}/txtGT_STEPLIST-STEPTEXT[0,{SPECIFY_FILES_ROW}]"
    cell = _ejecutar(
        f"Localizar celda del paso 'Specify Files' [{SPECIFY_FILES_ROW}]",
        session.findById, cell_id,
    )
    _ejecutar("Foco en celda 'Specify Files'", cell.setFocus)
    _ejecutar(
        "Cursor en celda 'Specify Files' (caretPosition=5)",
        lambda: setattr(cell, "caretPosition", 5),
    )
    wnd = _ejecutar(
        "Localizar wnd[0] para enviar F2",
        session.findById, "wnd[0]",
    )
    _ejecutar("Pulsar F2 (sendVKey 2) para abrir el paso", wnd.sendVKey, 2)

    # Botón "Cambiar" (modo edición)
    btn_cambiar = _ejecutar(
        "Localizar botón 'Cambiar' (wnd[0]/tbar[1]/btn[25])",
        session.findById, "wnd[0]/tbar[1]/btn[25]",
    )
    _ejecutar("Pulsar 'Cambiar' (modo edición)", btn_cambiar.press)

    # Seleccionar la definición de archivo a editar
    file_def = _ejecutar(
        "Localizar definición de archivo (wnd[0]/usr/lbl[43,6])",
        session.findById, "wnd[0]/usr/lbl[43,6]",
    )
    _ejecutar("Foco en definición de archivo", file_def.setFocus)
    _ejecutar(
        "Cursor en definición (caretPosition=3)",
        lambda: setattr(file_def, "caretPosition", 3),
    )

    # Botón "Asignar archivo" — abre diálogo modal
    btn_asignar = _ejecutar(
        "Localizar botón 'Asignar archivo' (wnd[0]/tbar[1]/btn[27])",
        session.findById, "wnd[0]/tbar[1]/btn[27]",
    )
    _ejecutar("Pulsar 'Asignar archivo'", btn_asignar.press)

    # F4 en el modal para abrir el explorador de archivos del frontend
    wnd1 = _ejecutar(
        "Localizar diálogo modal (wnd[1])",
        session.findById, "wnd[1]",
    )
    _ejecutar(
        "Pulsar F4 (sendVKey 4) para abrir explorador",
        wnd1.sendVKey, 4,
    )

    # Ingresar ruta y nombre en el explorador (wnd[2])
    path_field = _ejecutar(
        "Localizar campo ruta del explorador (wnd[2]/usr/ctxtDY_PATH)",
        session.findById, "wnd[2]/usr/ctxtDY_PATH",
    )
    _ejecutar(
        f"Asignar ruta = '{carpeta}'",
        lambda: setattr(path_field, "text", carpeta),
    )
    filename_field = _ejecutar(
        "Localizar campo nombre del explorador (wnd[2]/usr/ctxtDY_FILENAME)",
        session.findById, "wnd[2]/usr/ctxtDY_FILENAME",
    )
    _ejecutar(
        f"Asignar nombre = '{nombre_archivo}'",
        lambda: setattr(filename_field, "text", nombre_archivo),
    )
    _ejecutar(
        "Cursor al final del nombre",
        lambda: setattr(filename_field, "caretPosition", len(nombre_archivo)),
    )

    # Confirmar diálogos: OK explorador → OK modal
    ok_explorador = _ejecutar(
        "Localizar OK del explorador (wnd[2]/tbar[0]/btn[0])",
        session.findById, "wnd[2]/tbar[0]/btn[0]",
    )
    _ejecutar("Pulsar OK del explorador", ok_explorador.press)
    ok_modal = _ejecutar(
        "Localizar OK del modal (wnd[1]/tbar[0]/btn[0])",
        session.findById, "wnd[1]/tbar[0]/btn[0]",
    )
    _ejecutar("Pulsar OK del modal", ok_modal.press)

    # Volver al step list. Si SAP detecta cambios pendientes muestra un
    # popup "¿Guardar?", pero si la ruta/nombre ya era el mismo (corrida
    # previa con el mismo archivo) SAP vuelve directo al step list sin
    # popup. Por eso el click en 'Sí' es condicional.
    btn_back = _ejecutar(
        "Localizar botón Back (wnd[0]/tbar[0]/btn[3])",
        session.findById, "wnd[0]/tbar[0]/btn[3]",
    )
    _ejecutar("Pulsar Back para volver al step list", btn_back.press)

    try:
        btn_guardar = session.findById("wnd[1]/usr/btnSPOP-OPTION1")
        _log("  → Popup de guardar cambios detectado, confirmando con 'Sí'")
        btn_guardar.press()
    except Exception:
        _log(
            "  → Sin popup de guardar (no había cambios pendientes), "
            "continuando al step list"
        )


def _confirmar_popup_opcional(session, descripcion: str) -> None:
    """Intenta enviar Enter a `wnd[1]` para confirmar un popup. Si el popup
    no existe (porque SAP no lo mostró esta vez), loguea y continúa sin
    romper. Patrón clave para hacer el flujo resistente a popups
    condicionales que aparecen solo en ciertos estados."""
    try:
        session.findById("wnd[1]").sendVKey(0)
        _log(f"  → {descripcion}: popup detectado y confirmado con Enter")
    except Exception:
        _log(f"  → {descripcion}: sin popup, continuando")


def step_assign_files(session) -> None:
    """Abre y cierra el paso "Assign Files"."""
    _log("Paso 3/10: Asignando archivo a estructura (Assign Files)...")
    select_step_row(session, ASSIGN_FILES_ROW)
    boton_exec = _ejecutar(
        "Localizar botón Execute (F6 = wnd[0]/tbar[1]/btn[32])",
        session.findById, "wnd[0]/tbar[1]/btn[32]",
    )
    _ejecutar("Pulsar Execute en step list", boton_exec.press)
    wnd = _ejecutar("Localizar wnd[0]", session.findById, "wnd[0]")
    _ejecutar("Pulsar F3 (Back, sendVKey 3)", wnd.sendVKey, 3)


def step_read_data(session) -> None:
    """Ejecuta el paso "Read Data" — lee el .txt desde la ruta configurada."""
    _log("Paso 4/10: Leyendo datos del archivo .txt (Read Data)...")
    select_step_row(session, READ_DATA_ROW)
    boton_exec = _ejecutar(
        "Localizar botón Execute (wnd[0]/tbar[1]/btn[32])",
        session.findById, "wnd[0]/tbar[1]/btn[32]",
    )
    _ejecutar("Pulsar Execute en step list", boton_exec.press)
    boton_f8 = _ejecutar(
        "Localizar F8 (wnd[0]/tbar[1]/btn[8])",
        session.findById, "wnd[0]/tbar[1]/btn[8]",
    )
    _ejecutar("Pulsar F8 para leer datos", boton_f8.press)
    wnd = _ejecutar("Localizar wnd[0]", session.findById, "wnd[0]")
    _ejecutar("Pulsar Back (sendVKey 3)", wnd.sendVKey, 3)
    _ejecutar("Pulsar Back otra vez (sendVKey 3)", wnd.sendVKey, 3)


def step_display_read_data(session) -> None:
    """Paso "Display Read Data" — confirma popup y vuelve."""
    _log("Paso 5/10: Verificando datos leídos (Display Read Data)...")
    boton_exec = _ejecutar(
        "Localizar botón Execute (wnd[0]/tbar[1]/btn[32])",
        session.findById, "wnd[0]/tbar[1]/btn[32]",
    )
    _ejecutar("Pulsar Execute en step list", boton_exec.press)
    _confirmar_popup_opcional(session, "Confirmar popup de visualización")
    wnd = _ejecutar("Localizar wnd[0]", session.findById, "wnd[0]")
    _ejecutar("Pulsar Back (sendVKey 3)", wnd.sendVKey, 3)


def step_convert_data(session) -> None:
    """Paso "Convert Data" — convierte los datos leídos."""
    _log("Paso 6/10: Convirtiendo datos al formato SAP (Convert Data)...")
    boton_exec = _ejecutar(
        "Localizar botón Execute (wnd[0]/tbar[1]/btn[32])",
        session.findById, "wnd[0]/tbar[1]/btn[32]",
    )
    _ejecutar("Pulsar Execute en step list", boton_exec.press)
    wnd = _ejecutar("Localizar wnd[0]", session.findById, "wnd[0]")
    _ejecutar("Pulsar F8 (sendVKey 8) para convertir", wnd.sendVKey, 8)
    _ejecutar("Pulsar Back (sendVKey 3)", wnd.sendVKey, 3)
    _ejecutar("Pulsar Back otra vez (sendVKey 3)", wnd.sendVKey, 3)


def step_display_converted_data(session) -> None:
    """Paso "Display Converted Data" — confirma popup y vuelve."""
    _log("Paso 7/10: Verificando datos convertidos (Display Converted Data)...")
    boton_exec = _ejecutar(
        "Localizar botón Execute (wnd[0]/tbar[1]/btn[32])",
        session.findById, "wnd[0]/tbar[1]/btn[32]",
    )
    _ejecutar("Pulsar Execute en step list", boton_exec.press)
    _confirmar_popup_opcional(session, "Confirmar popup de visualización")
    wnd = _ejecutar("Localizar wnd[0]", session.findById, "wnd[0]")
    _ejecutar("Pulsar Back (sendVKey 3)", wnd.sendVKey, 3)


def step_create_batch_input(session) -> None:
    """Paso "Create Batch Input Session" — marca P_KEEP y crea la sesión."""
    _log("Paso 8/10: Creando sesión Batch Input (Create BI Session)...")
    boton_exec = _ejecutar(
        "Localizar botón Execute (wnd[0]/tbar[1]/btn[32])",
        session.findById, "wnd[0]/tbar[1]/btn[32]",
    )
    _ejecutar("Pulsar Execute en step list", boton_exec.press)
    keep_checkbox = _ejecutar(
        "Localizar checkbox 'Keep batch input folder' (wnd[0]/usr/chkP_KEEP)",
        session.findById, "wnd[0]/usr/chkP_KEEP",
    )
    _ejecutar(
        "Marcar checkbox 'Keep batch input folder'",
        lambda: setattr(keep_checkbox, "selected", True),
    )
    _ejecutar("Foco en checkbox Keep", keep_checkbox.setFocus)
    boton_f8 = _ejecutar(
        "Localizar F8 (wnd[0]/tbar[1]/btn[8])",
        session.findById, "wnd[0]/tbar[1]/btn[8]",
    )
    _ejecutar("Pulsar F8 para crear la sesión BDC", boton_f8.press)
    _confirmar_popup_opcional(session, "Confirmar 'sesión BDC creada'")


def step_run_batch_input(session) -> None:
    """Paso "Run Batch Input Session" — abre el listado de sesiones BDC."""
    _log("Paso 9/10: Abriendo lista de sesiones BDC (Run BI Session)...")
    boton_exec = _ejecutar(
        "Localizar botón Execute (wnd[0]/tbar[1]/btn[32])",
        session.findById, "wnd[0]/tbar[1]/btn[32]",
    )
    _ejecutar("Pulsar Execute en step list", boton_exec.press)


def process_bdc_session(session) -> None:
    """Procesa la sesión BDC recién creada en modo error + log completo."""
    _log("Paso 10/10: Procesando sesión BDC en modo error + log completo...")
    table = _ejecutar(
        f"Localizar tabla de sesiones BDC ({BDC_SESSION_TABLE})",
        session.findById, BDC_SESSION_TABLE,
    )
    _ejecutar(
        "Seleccionar primera fila de la tabla BDC",
        lambda: setattr(table.getAbsoluteRow(0), "selected", True),
    )

    group_cell_id = f"{BDC_SESSION_TABLE}/txtITAB_APQI-GROUPID[0,0]"
    group_cell = _ejecutar(
        f"Localizar celda GROUPID ({group_cell_id})",
        session.findById, group_cell_id,
    )
    _ejecutar("Foco en celda GROUPID", group_cell.setFocus)
    _ejecutar(
        "Cursor en GROUPID (caretPosition=0)",
        lambda: setattr(group_cell, "caretPosition", 0),
    )

    boton_f8 = _ejecutar(
        "Localizar F8 (wnd[0]/tbar[1]/btn[8])",
        session.findById, "wnd[0]/tbar[1]/btn[8]",
    )
    _ejecutar("Pulsar F8 (procesar sesión)", boton_f8.press)

    # Diálogo de procesamiento: modo error + log all + expert
    rad_error = _ejecutar(
        "Localizar radio 'Error mode' (wnd[1]/usr/radD0300-ERROR)",
        session.findById, "wnd[1]/usr/radD0300-ERROR",
    )
    _ejecutar("Seleccionar modo Error", rad_error.select)
    chk_logall = _ejecutar(
        "Localizar checkbox 'Log All' (wnd[1]/usr/chkD0300-LOGALL)",
        session.findById, "wnd[1]/usr/chkD0300-LOGALL",
    )
    _ejecutar(
        "Marcar 'Log All'",
        lambda: setattr(chk_logall, "selected", True),
    )
    chk_expert = _ejecutar(
        "Localizar checkbox 'Expert mode' (wnd[1]/usr/chkD0300-EXPERT)",
        session.findById, "wnd[1]/usr/chkD0300-EXPERT",
    )
    _ejecutar(
        "Marcar 'Expert mode'",
        lambda: setattr(chk_expert, "selected", True),
    )
    _ejecutar("Foco en 'Expert mode'", chk_expert.setFocus)
    boton_ok = _ejecutar(
        "Localizar OK del diálogo (wnd[1]/tbar[0]/btn[0])",
        session.findById, "wnd[1]/tbar[0]/btn[0]",
    )
    _ejecutar("Pulsar OK del diálogo (1ra vez)", boton_ok.press)
    _ejecutar("Pulsar OK del diálogo (2da vez, confirmar)", boton_ok.press)


def run_lsmw_flow(session, carpeta: str, nombre_archivo: str) -> None:
    """Ejecuta el flujo LSMW completo apuntando al archivo dado.

    Args:
        session: sesión SAP GUI activa.
        carpeta: ruta absoluta de la carpeta donde está el .txt.
        nombre_archivo: nombre del archivo .txt a cargar.
    """
    inicio = time.monotonic()
    _log("=== Iniciando flujo LSMW (10 pasos) ===")
    open_lsmw(session)
    configurar_ruta_archivo(session, carpeta, nombre_archivo)
    step_assign_files(session)
    step_read_data(session)
    step_display_read_data(session)
    step_convert_data(session)
    step_display_converted_data(session)
    step_create_batch_input(session)
    step_run_batch_input(session)
    process_bdc_session(session)
    duracion = time.monotonic() - inicio
    _log(f"=== Flujo LSMW finalizado en {duracion:.1f}s ===")


# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------

def main() -> int:
    print("=" * 70, flush=True)
    print("Carga automatizada del .txt a SAP vía LSMW", flush=True)
    print("=" * 70, flush=True)

    _log("Buscando el archivo .txt más reciente en salida/...")
    try:
        latest = get_latest_txt()
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}", file=sys.stderr, flush=True)
        return 1
    _log(f"Archivo encontrado: {latest.name}")

    _log("Conectando a la sesión SAP abierta...")
    try:
        session = get_sap_session()
    except RuntimeError as exc:
        print(f"ERROR: {exc}", file=sys.stderr, flush=True)
        return 1
    _log("Sesión SAP obtenida correctamente.")

    try:
        run_lsmw_flow(session, str(latest.parent), latest.name)
    except Exception as exc:
        print(f"\nERROR durante el flujo LSMW: {exc}", file=sys.stderr, flush=True)
        print(
            "Revisa la pantalla de SAP para ver en qué paso se detuvo. "
            "Posibles causas: proyecto LSMW no pre-cargado, IDs de la pantalla "
            "distintos, definición de archivo en otra posición.",
            file=sys.stderr,
            flush=True,
        )
        return 1

    print(flush=True)
    print("=" * 70, flush=True)
    print("Carga completada. Revisa SM35 para ver el log de la sesión BDC.", flush=True)
    print("=" * 70, flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
