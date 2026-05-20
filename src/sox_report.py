"""
sox_report.py — Generación del Reporte SOX vía SAP GUI Scripting.

Replica los pasos grabados en `resources/Scriptsox.vbs`. El flujo es:
1. Maximizar ventana y navegar al nodo F00039 del árbol del menú SAP.
2. Llenar Sociedad (P_BUKRS), rango de fechas (S_DATUM-LOW / S_DATUM-HIGH).
3. Ejecutar el reporte (F8).
4. Exportar a Excel vía menú contextual del grid (&MB_EXPORT → &XXL).
5. Guardar el archivo en la carpeta destino (por default `salida/`).

REQUISITOS DE EJECUCIÓN
=======================
Sistema operativo: Windows con SAP GUI for Windows abierto y sesión
iniciada. Mismos requisitos que `sap_upload.py` (ver su docstring).

USO
===
    python src/sox_report.py SOCIEDAD DESDE HASTA

    Ejemplo:
        python src/sox_report.py ISA 01.05.2026 31.05.2026

También se puede invocar desde la GUI vía el botón "Control SOX" de
`main.py`.
"""

from __future__ import annotations

import sys
import time
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SALIDA_DIR = PROJECT_ROOT / "salida"

# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------

# Sociedades válidas (mismas opciones que el combo del formulario).
VALID_SOCIEDADES = (
    "TRAN", "ISA", "ITCH", "CEYBA", "CABA", "RPAE", "CTMP", "REPD", "ISAP",
)

# Formato esperado en los campos de fecha del formulario (y de SAP).
DATE_FORMAT_USER = "%d.%m.%Y"

# IDs SAP capturados de resources/Scriptsox.vbs.
TREE_SHELL = (
    "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell"
)
SOX_NODE_KEY = "F00039"

CAMPO_SOCIEDAD = "wnd[0]/usr/ctxtP_BUKRS"
CAMPO_FECHA_DESDE = "wnd[0]/usr/ctxtS_DATUM-LOW"
CAMPO_FECHA_HASTA = "wnd[0]/usr/ctxtS_DATUM-HIGH"

DOCS_GRID_SHELL = (
    "wnd[0]/usr/subDISPLAY:SAPLBANK_OBJ_CHDOC:0210/"
    "cntlCC_CHANGE_DOCUMENTS_SURVAY/shellcont/shell/shellcont[1]/shell"
)


# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------

def _log(mensaje: str) -> None:
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {mensaje}", flush=True)


# ---------------------------------------------------------------------------
# VALIDACIONES
# ---------------------------------------------------------------------------

def validar_sociedad(sociedad: str) -> str:
    """Verifica que la sociedad esté en VALID_SOCIEDADES.

    Devuelve la sociedad normalizada (uppercase + strip). Lanza ValueError
    si no es válida o está vacía.
    """
    if not isinstance(sociedad, str) or not sociedad.strip():
        raise ValueError("Debes seleccionar una sociedad.")
    norm = sociedad.strip().upper()
    if norm not in VALID_SOCIEDADES:
        raise ValueError(
            f"Sociedad inválida: '{sociedad}'. "
            f"Opciones válidas: {', '.join(VALID_SOCIEDADES)}."
        )
    return norm


def validar_fecha(fecha_str: str, etiqueta: str = "fecha") -> datetime:
    """Valida y parsea una fecha en formato dd.mm.aaaa.

    Args:
        fecha_str: cadena a parsear.
        etiqueta: nombre del campo (para mensajes de error).
    """
    if not isinstance(fecha_str, str) or not fecha_str.strip():
        raise ValueError(f"La {etiqueta} está vacía.")
    try:
        return datetime.strptime(fecha_str.strip(), DATE_FORMAT_USER)
    except ValueError as exc:
        raise ValueError(
            f"La {etiqueta} '{fecha_str}' no tiene el formato esperado dd.mm.aaaa."
        ) from exc


def validar_rango_fechas(desde: str, hasta: str) -> tuple[datetime, datetime]:
    """Valida ambas fechas y que `hasta >= desde`."""
    f_desde = validar_fecha(desde, etiqueta="fecha desde")
    f_hasta = validar_fecha(hasta, etiqueta="fecha hasta")
    if f_hasta < f_desde:
        raise ValueError(
            f"La fecha hasta ({hasta}) debe ser mayor o igual a la fecha desde ({desde})."
        )
    return f_desde, f_hasta


def validar_caracter_fecha(propuesto: str) -> bool:
    """Validación per-keystroke: solo dígitos y puntos, máx 10 caracteres.

    Se usa como `validatecommand` de los Entry de fecha para impedir que el
    usuario escriba letras u otros caracteres extraños.
    """
    if len(propuesto) > 10:
        return False
    return all(c.isdigit() or c == "." for c in propuesto)


# ---------------------------------------------------------------------------
# CONEXIÓN A SAP
# ---------------------------------------------------------------------------

def get_sap_session():
    """Conecta al SAP GUI Scripting Engine. Igual lógica que sap_upload."""
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
# PASOS DEL FLUJO SOX
# ---------------------------------------------------------------------------

def abrir_transaccion_sox(session) -> None:
    """Maximiza la ventana y abre la transacción haciendo doble clic en el
    nodo `F00039` del árbol del menú SAP."""
    _log("Paso 1/4: Abriendo transacción SOX (nodo F00039)...")
    session.findById("wnd[0]").maximize()
    session.findById(TREE_SHELL).doubleClickNode(SOX_NODE_KEY)


def ingresar_parametros(
    session, sociedad: str, fecha_desde: str, fecha_hasta: str
) -> None:
    """Llena P_BUKRS, S_DATUM-LOW, S_DATUM-HIGH y ejecuta el reporte (F8)."""
    _log(
        f"Paso 2/4: Ingresando sociedad='{sociedad}', "
        f"desde='{fecha_desde}', hasta='{fecha_hasta}'..."
    )
    session.findById(CAMPO_SOCIEDAD).text = sociedad
    session.findById(CAMPO_FECHA_DESDE).text = fecha_desde
    session.findById(CAMPO_FECHA_HASTA).text = fecha_hasta
    hasta_field = session.findById(CAMPO_FECHA_HASTA)
    hasta_field.setFocus()
    hasta_field.caretPosition = 5

    _log("Paso 3/4: Ejecutando reporte (F8)...")
    session.findById("wnd[0]/tbar[1]/btn[8]").press()


def exportar_a_excel(
    session, carpeta_destino: str, nombre_archivo: str
) -> None:
    """Exporta el grid resultante a Excel (XXL) y guarda en la ruta dada.

    El recording original termina con `wnd[1].close`. En la práctica, tras
    seleccionar &XXL SAP abre un diálogo de guardar archivo. Esta función
    intenta rellenarlo (DY_PATH/DY_FILENAME); si la estructura del diálogo
    difiere en otra instalación, hay que ajustar los IDs.
    """
    _log("Paso 4/4: Exportando grid a Excel (XXL)...")
    grid = session.findById(DOCS_GRID_SHELL)
    grid.pressToolbarContextButton("&MB_EXPORT")
    grid.selectContextMenuItem("&XXL")

    # Intento de rellenar el diálogo de guardar archivo.
    try:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = carpeta_destino
        nombre_field = session.findById("wnd[1]/usr/ctxtDY_FILENAME")
        nombre_field.text = nombre_archivo
        nombre_field.caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # OK
        _log(f"Archivo guardado en: {carpeta_destino}\\{nombre_archivo}")
    except Exception as exc:
        # Fallback: replicar el recording original que cierra wnd[1]. Útil
        # si SAP en esta instalación abre Excel directo en vez de mostrar
        # un diálogo de guardar.
        _log(
            f"(no se detectó diálogo DY_PATH/DY_FILENAME: {exc} — "
            f"cerrando wnd[1] como en el recording original)"
        )
        try:
            session.findById("wnd[1]").close()
        except Exception:
            pass


def generar_reporte_sox(
    session,
    sociedad: str,
    fecha_desde: str,
    fecha_hasta: str,
    carpeta_destino: str | None = None,
    nombre_archivo: str | None = None,
) -> tuple[str, str]:
    """Ejecuta el flujo SOX completo y devuelve (carpeta, nombre) usados.

    Args:
        session: sesión SAP GUI activa.
        sociedad: código de sociedad (debe estar en VALID_SOCIEDADES).
        fecha_desde: fecha inicial en formato dd.mm.aaaa.
        fecha_hasta: fecha final en formato dd.mm.aaaa.
        carpeta_destino: ruta donde guardar el .xlsx (default: salida/).
        nombre_archivo: nombre del .xlsx (default: SOX_{soc}_{ts}.xlsx).

    Returns:
        (carpeta, nombre): rutas usadas para el guardado.
    """
    sociedad_norm = validar_sociedad(sociedad)
    validar_rango_fechas(fecha_desde, fecha_hasta)

    if carpeta_destino is None:
        SALIDA_DIR.mkdir(parents=True, exist_ok=True)
        carpeta_destino = str(SALIDA_DIR)
    if nombre_archivo is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"SOX_{sociedad_norm}_{ts}.xlsx"

    inicio = time.monotonic()
    _log("=== Iniciando flujo SOX (4 pasos) ===")
    abrir_transaccion_sox(session)
    ingresar_parametros(session, sociedad_norm, fecha_desde, fecha_hasta)
    exportar_a_excel(session, carpeta_destino, nombre_archivo)
    duracion = time.monotonic() - inicio
    _log(f"=== Flujo SOX finalizado en {duracion:.1f}s ===")

    return carpeta_destino, nombre_archivo


# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    print("=" * 70, flush=True)
    print("Generación de Reporte SOX vía SAP GUI Scripting", flush=True)
    print("=" * 70, flush=True)

    if len(argv) != 3:
        print(
            "Uso: python src/sox_report.py SOCIEDAD DESDE HASTA\n"
            "Ejemplo: python src/sox_report.py ISA 01.05.2026 31.05.2026",
            file=sys.stderr,
        )
        return 2

    sociedad, desde, hasta = argv

    try:
        validar_sociedad(sociedad)
        validar_rango_fechas(desde, hasta)
    except ValueError as exc:
        print(f"ERROR de validación: {exc}", file=sys.stderr, flush=True)
        return 1

    try:
        session = get_sap_session()
    except RuntimeError as exc:
        print(f"ERROR: {exc}", file=sys.stderr, flush=True)
        return 1

    try:
        carpeta, nombre = generar_reporte_sox(session, sociedad, desde, hasta)
    except Exception as exc:
        print(f"\nERROR durante el flujo SOX: {exc}", file=sys.stderr, flush=True)
        return 1

    print(flush=True)
    print("=" * 70, flush=True)
    print(f"Reporte SOX generado: {carpeta}\\{nombre}", flush=True)
    print("=" * 70, flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
