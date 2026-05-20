"""
sox_report.py — Generación del Reporte SOX vía SAP GUI Scripting.

Replica los pasos grabados en `resources/Script2sox.vbs` (versión actual,
con T-code y calendario F4) — reemplaza la grabación inicial fragil con
nodos F00xxx del árbol que estaba en `resources/Scriptsox.vbs`.

Flujo:
1. Maximizar ventana y abrir la transacción SAP `AR15` vía okcd.
2. Llenar Sociedad (P_BUKRS) — texto directo.
3. Llenar Fecha Desde (S_DATUM-LOW) y Fecha Hasta (S_DATUM-HIGH) usando
   el calendario emergente (sendVKey 4 → focusDate + selectionInterval
   en formato yyyymmdd).
4. Ejecutar el reporte (F8).
5. (Opcional) Exportar a Excel vía menú contextual del grid
   (&MB_EXPORT → &XXL) y guardar en `salida/`.
   IMPORTANTE: el grid de AR15 usa otro shell ID que el del recording
   original. Si la exportación falla, hay que re-grabar ese paso y
   actualizar `DOCS_GRID_SHELL`.

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

# T-code SAP de la transacción del reporte SOX. Forma ROBUSTA de abrir la
# transacción (escribir el código en okcd y Enter) — no depende del árbol
# del menú, que tiene IDs (F00xxx) inestables entre usuarios y sesiones.
#
# Confirmado por `resources/Script2sox.vbs`: la T-code es AR15 (transacción
# estándar SAP de "Origen de altas de inmovilizado" / "Asset History").
# Si la T-code real de tu instalación es otra (variante Z*), ajustar aquí.
#
# Si se deja en None, el script hace fallback al árbol con SOX_NODE_KEY.
T_CODE_SOX: str | None = "AR15"

# Shell del calendario emergente que aparece al presionar F4 sobre un
# campo de fecha de SAP. Vía: setFocus + sendVKey(4) sobre el campo →
# foco/selección sobre este shell.
CALENDAR_SHELL = "wnd[1]/usr/cntlCONTAINER/shellcont/shell"

# Formato yyyymmdd que espera el calendario SAP para focusDate y
# selectionInterval (distinto del dd.mm.aaaa del formulario).
DATE_FORMAT_SAP_CALENDAR = "%Y%m%d"

CAMPO_SOCIEDAD = "wnd[0]/usr/ctxtP_BUKRS"
CAMPO_FECHA_DESDE = "wnd[0]/usr/ctxtS_DATUM-LOW"
CAMPO_FECHA_HASTA = "wnd[0]/usr/ctxtS_DATUM-HIGH"

# Shell del grid de resultados (sólo aplica al método "alv_grid").
# Este path viene del recording original (Scriptsox.vbs) — no aplica a
# AR15. Si en algún momento se necesita exportar desde un ALV grid con
# este método, ajustar este ID.
DOCS_GRID_SHELL = (
    "wnd[0]/usr/subDISPLAY:SAPLBANK_OBJ_CHDOC:0210/"
    "cntlCC_CHANGE_DOCUMENTS_SURVAY/shellcont/shell/shellcont[1]/shell"
)

# Método de exportación a archivo:
#   "pc_list"  → usa la T-code %PC (System > List > Save > File). Funciona
#                con listas SAP clásicas como AR15. Es el default actual.
#   "alv_grid" → método del recording original (&MB_EXPORT > &XXL sobre el
#                grid). Sólo aplica a ALV grids específicos.
#   None       → no exporta. Deja el reporte abierto en SAP y el usuario
#                guarda manualmente.
EXPORT_METHOD: str | None = "pc_list"


# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------

def _log(mensaje: str) -> None:
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {mensaje}", flush=True)


def _ejecutar(descripcion: str, fn, *args, **kwargs):
    """Ejecuta `fn(*args, **kwargs)` loguenado la operación. Si falla,
    re-lanza con un mensaje descriptivo que dice exactamente qué intentaba
    hacer — esto es clave porque las excepciones COM de SAP (`SAP Frontend
    Server`) suelen venir con descripción vacía.
    """
    _log(f"  → {descripcion}")
    try:
        return fn(*args, **kwargs)
    except Exception as exc:
        raise RuntimeError(
            f"Falló: {descripcion}\n"
            f"Detalle técnico SAP: {exc!r}"
        ) from exc


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

def _intentar_listar_nodos_arbol(tree) -> str:
    """Intenta enumerar los nodos visibles del árbol SAP para diagnóstico.

    Útil cuando un `doubleClickNode(...)` falla porque el ID grabado ya no
    aplica. Devuelve una cadena con los pares (key → texto) listos para
    incluir en el mensaje de error.
    """
    try:
        keys = tree.GetAllNodeKeys()
    except Exception as exc:
        return f"  (no se pudo enumerar el árbol: {exc!r})"

    keys_list = list(keys) if keys else []
    if not keys_list:
        return "  (árbol vacío o sin nodos visibles)"

    lineas = []
    for key in keys_list[:50]:
        try:
            texto = tree.GetNodeTextByKey(key)
            lineas.append(f"  {key} → {texto}")
        except Exception:
            lineas.append(f"  {key} → (no se pudo leer el texto del nodo)")
    if len(keys_list) > 50:
        lineas.append(f"  ... ({len(keys_list) - 50} nodos más)")
    return "\n".join(lineas)


def abrir_transaccion_sox(session) -> None:
    """Maximiza la ventana y abre la transacción del reporte SOX.

    Si `T_CODE_SOX` está configurado, navega vía la casilla de comandos
    (okcd) — esto es ROBUSTO entre usuarios y sesiones. Si no, intenta el
    fallback de doble-clic en el árbol con `SOX_NODE_KEY` (como el
    recording original), pero los IDs F00xxx del árbol son inestables.
    """
    _log("Paso 1/4: Abriendo transacción SOX...")

    wnd = _ejecutar(
        "Localizar ventana principal wnd[0]",
        session.findById, "wnd[0]",
    )
    _ejecutar("Maximizar ventana principal", wnd.maximize)

    # Camino preferido: T-code en la casilla de comandos.
    if T_CODE_SOX:
        _log(f"  Modo T-code (recomendado): usando '{T_CODE_SOX}'")
        okcd = _ejecutar(
            "Localizar casilla de comandos (wnd[0]/tbar[0]/okcd)",
            session.findById, "wnd[0]/tbar[0]/okcd",
        )
        _ejecutar(
            f"Escribir T-code '{T_CODE_SOX}' en okcd",
            lambda: setattr(okcd, "text", T_CODE_SOX),
        )
        _ejecutar("Enviar Enter (sendVKey 0)", wnd.sendVKey, 0)
        return

    # Fallback: navegación del árbol como en el recording.
    _log(
        f"  Modo árbol (fallback): doble-clic en nodo {SOX_NODE_KEY!r} — "
        f"frágil entre usuarios. Considera configurar T_CODE_SOX."
    )

    try:
        tree = _ejecutar(
            f"Localizar árbol del menú SAP ({TREE_SHELL})",
            session.findById, TREE_SHELL,
        )
    except RuntimeError as exc:
        raise RuntimeError(
            f"{exc}\n\n"
            f"PISTA: el árbol del menú no se encuentra. Verifica:\n"
            f"  • Estar logueado en SAP y en la pantalla SAP Easy Access.\n"
            f"  • Que el menú de roles del usuario sea visible (no minimizado).\n"
            f"  • Que la ruta del árbol coincida con tu instalación."
        ) from exc

    try:
        _ejecutar(
            f"Doble clic en el nodo {SOX_NODE_KEY!r} del árbol",
            tree.doubleClickNode, SOX_NODE_KEY,
        )
    except RuntimeError as exc:
        # Diagnóstico extra: listar los nodos disponibles para que el usuario
        # identifique cuál es el correcto en SU árbol.
        nodos_disponibles = _intentar_listar_nodos_arbol(tree)
        raise RuntimeError(
            f"{exc}\n\n"
            f"DIAGNÓSTICO: los IDs del árbol SAP (F00xxx) son posiciones\n"
            f"secuenciales asignadas cuando se renderiza el menú del usuario\n"
            f"que grabó el script. Cambian entre usuarios y sesiones.\n\n"
            f"SOLUCIÓN ROBUSTA: configurar la constante T_CODE_SOX al inicio\n"
            f"de src/sox_report.py con la T-code real de la transacción\n"
            f"(ej. T_CODE_SOX = 'ZTRX_SOX').\n\n"
            f"Para descubrir la T-code:\n"
            f"  1. En SAP, abre la transacción manualmente (como lo hacías).\n"
            f"  2. Ve a 'Sistema → Estado' o mira la barra de título.\n"
            f"  3. El campo 'Transacción' muestra la T-code (ej. ZSOX_REPORT).\n\n"
            f"Nodos visibles en tu árbol actual:\n{nodos_disponibles}"
        ) from exc


def _seleccionar_fecha_calendario(
    session, campo_id: str, fecha_str: str, etiqueta: str
) -> None:
    """Selecciona una fecha en SAP usando el calendario emergente F4.

    Replica `resources/Script2sox.vbs`: foco en el campo → sendVKey(4) abre
    el calendario en wnd[1] → setea focusDate y selectionInterval con la
    fecha en formato yyyymmdd.

    Args:
        session: sesión SAP GUI.
        campo_id: ID del campo de fecha (CAMPO_FECHA_DESDE / CAMPO_FECHA_HASTA).
        fecha_str: fecha en formato dd.mm.aaaa (se convierte internamente).
        etiqueta: nombre legible del campo para los logs (ej. "Desde").
    """
    fecha_sap = validar_fecha(fecha_str, etiqueta=etiqueta).strftime(
        DATE_FORMAT_SAP_CALENDAR
    )

    campo = _ejecutar(
        f"Localizar campo Fecha {etiqueta} ({campo_id})",
        session.findById, campo_id,
    )
    _ejecutar(f"Foco en campo Fecha {etiqueta}", campo.setFocus)
    _ejecutar(
        f"Posicionar cursor en Fecha {etiqueta} (caretPosition=0)",
        lambda: setattr(campo, "caretPosition", 0),
    )

    wnd = _ejecutar(
        "Localizar ventana principal wnd[0]",
        session.findById, "wnd[0]",
    )
    _ejecutar(
        f"Abrir calendario emergente F4 para Fecha {etiqueta}",
        wnd.sendVKey, 4,
    )

    calendario = _ejecutar(
        f"Localizar calendario emergente ({CALENDAR_SHELL})",
        session.findById, CALENDAR_SHELL,
    )
    _ejecutar(
        f"Enfocar fecha {fecha_sap} en calendario",
        lambda: setattr(calendario, "focusDate", fecha_sap),
    )
    _ejecutar(
        f"Seleccionar intervalo {fecha_sap},{fecha_sap}",
        lambda: setattr(
            calendario, "selectionInterval", f"{fecha_sap},{fecha_sap}"
        ),
    )


def ingresar_parametros(
    session, sociedad: str, fecha_desde: str, fecha_hasta: str
) -> None:
    """Llena P_BUKRS (texto directo), Fecha Desde/Hasta vía calendario F4 y
    ejecuta el reporte (F8). Replica el flujo grabado en `Script2sox.vbs`."""
    _log(
        f"Paso 2/4: Ingresando sociedad='{sociedad}', "
        f"desde='{fecha_desde}', hasta='{fecha_hasta}'..."
    )

    sociedad_field = _ejecutar(
        f"Localizar campo Sociedad ({CAMPO_SOCIEDAD})",
        session.findById, CAMPO_SOCIEDAD,
    )
    _ejecutar(
        f"Asignar Sociedad = '{sociedad}'",
        lambda: setattr(sociedad_field, "text", sociedad),
    )

    _seleccionar_fecha_calendario(
        session, CAMPO_FECHA_DESDE, fecha_desde, "Desde"
    )
    _seleccionar_fecha_calendario(
        session, CAMPO_FECHA_HASTA, fecha_hasta, "Hasta"
    )

    _log("Paso 3/4: Ejecutando reporte (F8)...")
    boton_f8 = _ejecutar(
        "Localizar botón Ejecutar (F8 = wnd[0]/tbar[1]/btn[8])",
        session.findById, "wnd[0]/tbar[1]/btn[8]",
    )
    _ejecutar("Pulsar Ejecutar (F8)", boton_f8.press)


def _rellenar_save_dialog(
    session, carpeta_destino: str, nombre_archivo: str
) -> None:
    """Llena el diálogo estándar de SAP "Save File" (DY_PATH + DY_FILENAME)
    y presiona OK. Usado por ambos métodos de exportación."""
    path_field = _ejecutar(
        "Localizar campo de ruta (wnd[1]/usr/ctxtDY_PATH)",
        session.findById, "wnd[1]/usr/ctxtDY_PATH",
    )
    _ejecutar(
        f"Asignar ruta = '{carpeta_destino}'",
        lambda: setattr(path_field, "text", carpeta_destino),
    )

    nombre_field = _ejecutar(
        "Localizar campo de nombre (wnd[1]/usr/ctxtDY_FILENAME)",
        session.findById, "wnd[1]/usr/ctxtDY_FILENAME",
    )
    _ejecutar(
        f"Asignar nombre = '{nombre_archivo}'",
        lambda: setattr(nombre_field, "text", nombre_archivo),
    )
    _ejecutar(
        "Posicionar cursor al final del nombre",
        lambda: setattr(nombre_field, "caretPosition", len(nombre_archivo)),
    )

    boton_ok = _ejecutar(
        "Localizar botón OK del diálogo (wnd[1]/tbar[0]/btn[0])",
        session.findById, "wnd[1]/tbar[0]/btn[0]",
    )
    _ejecutar("Pulsar OK para guardar", boton_ok.press)


def _exportar_via_pc_list(
    session, carpeta_destino: str, nombre_archivo: str
) -> None:
    """Exporta usando %PC (System > List > Save > File). Funciona para
    listas SAP clásicas como AR15.

    La estructura del diálogo que abre %PC varía entre versiones SAP:
      - Versión A: muestra primero un diálogo de selección de formato
        con OK en `tbar[0]/btn[0]`, después abre el save-as.
      - Versión B: muestra un popup distinto sin ese botón estándar.
      - Versión C: abre directamente el save-as (sin paso de formato).

    Estrategia robusta:
      1. Tras %PC, probar si ya hay un save-as en wnd[1] (buscar DY_PATH).
      2. Si no, enviar Enter (sendVKey 0) a wnd[1] — actúa como "OK"
         universal en cualquier diálogo modal sin depender del ID del
         botón. Después intentamos de nuevo el save-as.
    """
    _log("  Modo PC list: usando %PC en okcd...")
    wnd = _ejecutar(
        "Localizar ventana principal wnd[0]",
        session.findById, "wnd[0]",
    )
    okcd = _ejecutar(
        "Localizar casilla de comandos (wnd[0]/tbar[0]/okcd)",
        session.findById, "wnd[0]/tbar[0]/okcd",
    )
    _ejecutar(
        "Escribir '%PC' en okcd (System > List > Save > File)",
        lambda: setattr(okcd, "text", "%PC"),
    )
    _ejecutar("Enviar Enter (sendVKey 0)", wnd.sendVKey, 0)

    # ¿Save-as ya está abierto? Si DY_PATH existe en wnd[1], saltamos el
    # paso de "confirmar formato".
    save_dialog_listo = False
    try:
        session.findById("wnd[1]/usr/ctxtDY_PATH")
        save_dialog_listo = True
        _log("  → Save-as detectado directamente en wnd[1] (sin paso de formato)")
    except Exception:
        _log("  → wnd[1] no es el save-as todavía; enviando Enter para avanzar")

    if not save_dialog_listo:
        # Mandar Enter al diálogo de formato (cualquiera que sea su
        # estructura). Enter actúa como OK por default.
        wnd1 = _ejecutar(
            "Localizar diálogo intermedio (wnd[1])",
            session.findById, "wnd[1]",
        )
        _ejecutar("Confirmar formato con Enter (sendVKey 0)", wnd1.sendVKey, 0)

    # Diálogo de guardar archivo (debe estar abierto ahora)
    _rellenar_save_dialog(session, carpeta_destino, nombre_archivo)
    _log(f"Archivo guardado en: {carpeta_destino}\\{nombre_archivo}")


def _exportar_via_alv_grid(
    session, carpeta_destino: str, nombre_archivo: str
) -> None:
    """Exporta usando el menú contextual del ALV grid (&MB_EXPORT > &XXL).
    Sólo funciona si el grid tiene el shell ID configurado en
    DOCS_GRID_SHELL (no aplica a AR15)."""
    _log("  Modo ALV grid: usando &MB_EXPORT + &XXL...")
    grid = _ejecutar(
        f"Localizar grid de resultados ({DOCS_GRID_SHELL})",
        session.findById, DOCS_GRID_SHELL,
    )
    _ejecutar(
        "Abrir menú de exportación (&MB_EXPORT)",
        grid.pressToolbarContextButton, "&MB_EXPORT",
    )
    _ejecutar(
        "Seleccionar exportación a Excel (&XXL)",
        grid.selectContextMenuItem, "&XXL",
    )
    _rellenar_save_dialog(session, carpeta_destino, nombre_archivo)
    _log(f"Archivo guardado en: {carpeta_destino}\\{nombre_archivo}")


def exportar_a_excel(
    session, carpeta_destino: str, nombre_archivo: str
) -> None:
    """Exporta el reporte al archivo dado usando el método configurado en
    `EXPORT_METHOD`.

    - "pc_list" (default): usa %PC, recomendado para AR15.
    - "alv_grid": usa &MB_EXPORT > &XXL, requiere DOCS_GRID_SHELL válido.
    - None: no exporta; el reporte queda visible en SAP para guardar
      manualmente.
    """
    _log("Paso 4/4: Exportando reporte...")

    if EXPORT_METHOD is None:
        _log(
            "  EXPORT_METHOD=None → exportación omitida. "
            "Guarda el reporte manualmente desde SAP."
        )
        return

    if EXPORT_METHOD == "pc_list":
        _exportar_via_pc_list(session, carpeta_destino, nombre_archivo)
        return

    if EXPORT_METHOD == "alv_grid":
        _exportar_via_alv_grid(session, carpeta_destino, nombre_archivo)
        return

    raise ValueError(
        f"EXPORT_METHOD inválido: '{EXPORT_METHOD}'. "
        f"Usa 'pc_list', 'alv_grid' o None."
    )


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
