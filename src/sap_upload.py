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
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Falta la dependencia pywin32. Instalar con: pip install pywin32"
        ) from exc

    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
    except Exception as exc:
        raise RuntimeError(
            "No se pudo conectar a SAP GUI. Verifica:\n"
            "  - SAP GUI for Windows está abierto y con sesión iniciada.\n"
            "  - SAP GUI Scripting habilitado en Options del cliente.\n"
            "  - sapgui/user_scripting = TRUE en el servidor SAP."
        ) from exc

    application = sap_gui_auto.GetScriptingEngine
    if application.Children.Count == 0:
        raise RuntimeError("No hay conexiones SAP activas en este SAP GUI.")
    connection = application.Children(0)
    if connection.Children.Count == 0:
        raise RuntimeError(
            "No hay sesiones activas en la conexión SAP. "
            "Inicia sesión en el sistema SAP antes de correr este script."
        )
    return connection.Children(0)


# ---------------------------------------------------------------------------
# FLUJO LSMW
# ---------------------------------------------------------------------------

def select_step_row(session, row: int) -> None:
    """Selecciona una fila concreta del step list, deseleccionando la default."""
    table = session.findById(LSMW_STEPLIST_TABLE)
    table.getAbsoluteRow(DEFAULT_SELECTED_ROW).selected = False
    table.getAbsoluteRow(row).selected = True
    cell = session.findById(
        f"{LSMW_STEPLIST_TABLE}/txtGT_STEPLIST-STEPTEXT[0,{row}]"
    )
    cell.setFocus()
    cell.caretPosition = 0


def open_lsmw(session) -> None:
    """Abre la T-code LSMW y ejecuta el proyecto pre-cargado."""
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "LSMW"
    session.findById("wnd[0]").sendVKey(0)
    # F8 — entra al step list del proyecto pre-cargado
    session.findById("wnd[0]/tbar[1]/btn[8]").press()


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
    # Foco en la celda del paso "Specify Files" (row 6) y F2 para abrirlo
    cell_id = f"{LSMW_STEPLIST_TABLE}/txtGT_STEPLIST-STEPTEXT[0,{SPECIFY_FILES_ROW}]"
    cell = session.findById(cell_id)
    cell.setFocus()
    cell.caretPosition = 5
    session.findById("wnd[0]").sendVKey(2)

    # Botón "Cambiar" (modo edición)
    session.findById("wnd[0]/tbar[1]/btn[25]").press()

    # Seleccionar la definición de archivo a editar
    file_def = session.findById("wnd[0]/usr/lbl[43,6]")
    file_def.setFocus()
    file_def.caretPosition = 3

    # Botón "Asignar archivo" — abre diálogo modal
    session.findById("wnd[0]/tbar[1]/btn[27]").press()

    # F4 en el modal para abrir el explorador de archivos del frontend
    session.findById("wnd[1]").sendVKey(4)

    # Ingresar ruta y nombre en el explorador (wnd[2])
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = carpeta
    filename_field = session.findById("wnd[2]/usr/ctxtDY_FILENAME")
    filename_field.text = nombre_archivo
    filename_field.caretPosition = len(nombre_archivo)

    # Confirmar diálogos: OK explorador → OK modal
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Volver al step list y confirmar guardar cambios
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()


def step_assign_files(session) -> None:
    """Abre y cierra el paso "Assign Files"."""
    select_step_row(session, ASSIGN_FILES_ROW)
    session.findById("wnd[0]/tbar[1]/btn[32]").press()
    session.findById("wnd[0]").sendVKey(3)              # F3 — Back


def step_read_data(session) -> None:
    """Ejecuta el paso "Read Data" — lee el .txt desde la ruta configurada."""
    select_step_row(session, READ_DATA_ROW)
    session.findById("wnd[0]/tbar[1]/btn[32]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()   # F8 — ejecutar lectura
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def step_display_read_data(session) -> None:
    """Paso "Display Read Data" — confirma popup y vuelve."""
    session.findById("wnd[0]/tbar[1]/btn[32]").press()
    session.findById("wnd[1]").sendVKey(0)              # Confirma popup
    session.findById("wnd[0]").sendVKey(3)


def step_convert_data(session) -> None:
    """Paso "Convert Data" — convierte los datos leídos."""
    session.findById("wnd[0]/tbar[1]/btn[32]").press()
    session.findById("wnd[0]").sendVKey(8)              # F8 — ejecutar conversión
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def step_display_converted_data(session) -> None:
    """Paso "Display Converted Data" — confirma popup y vuelve."""
    session.findById("wnd[0]/tbar[1]/btn[32]").press()
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(3)


def step_create_batch_input(session) -> None:
    """Paso "Create Batch Input Session" — marca P_KEEP y crea la sesión."""
    session.findById("wnd[0]/tbar[1]/btn[32]").press()
    keep_checkbox = session.findById("wnd[0]/usr/chkP_KEEP")
    keep_checkbox.selected = True
    keep_checkbox.setFocus()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()   # F8 — crea la sesión BDC
    session.findById("wnd[1]").sendVKey(0)              # Confirma popup


def step_run_batch_input(session) -> None:
    """Paso "Run Batch Input Session" — abre el listado de sesiones BDC."""
    session.findById("wnd[0]/tbar[1]/btn[32]").press()


def process_bdc_session(session) -> None:
    """Procesa la sesión BDC recién creada en modo error + log completo."""
    table = session.findById(BDC_SESSION_TABLE)
    table.getAbsoluteRow(0).selected = True

    group_cell = session.findById(
        f"{BDC_SESSION_TABLE}/txtITAB_APQI-GROUPID[0,0]"
    )
    group_cell.setFocus()
    group_cell.caretPosition = 0

    session.findById("wnd[0]/tbar[1]/btn[8]").press()   # F8 — Process

    # Diálogo de procesamiento: modo error + log all + expert
    session.findById("wnd[1]/usr/radD0300-ERROR").select()
    session.findById("wnd[1]/usr/chkD0300-LOGALL").selected = True
    session.findById("wnd[1]/usr/chkD0300-EXPERT").selected = True
    session.findById("wnd[1]/usr/chkD0300-EXPERT").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()   # OK
    session.findById("wnd[1]/tbar[0]/btn[0]").press()   # OK confirmar


def run_lsmw_flow(session, carpeta: str, nombre_archivo: str) -> None:
    """Ejecuta el flujo LSMW completo apuntando al archivo dado.

    Args:
        session: sesión SAP GUI activa.
        carpeta: ruta absoluta de la carpeta donde está el .txt.
        nombre_archivo: nombre del archivo .txt a cargar.
    """
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


# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------

def main() -> int:
    print("=" * 70)
    print("Carga automatizada del .txt a SAP vía LSMW")
    print("=" * 70)

    try:
        latest = get_latest_txt()
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(f"Archivo origen: {latest}")

    try:
        session = get_sap_session()
    except RuntimeError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print("Conectado a SAP. Ejecutando flujo LSMW...")

    try:
        run_lsmw_flow(session, str(latest.parent), latest.name)
    except Exception as exc:
        print(f"\nERROR durante el flujo LSMW: {exc}", file=sys.stderr)
        print(
            "Revisa la pantalla de SAP para ver en qué paso se detuvo. "
            "Posibles causas: proyecto LSMW no pre-cargado, IDs de la pantalla "
            "distintos, definición de archivo en otra posición.",
            file=sys.stderr,
        )
        return 1

    print()
    print("=" * 70)
    print("Carga completada. Revisa SM35 para ver el log de la sesión BDC.")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    sys.exit(main())
