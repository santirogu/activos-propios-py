"""Pruebas unitarias para sox_report.py.

Las funciones que dialogan con SAP GUI Scripting se prueban con
`MockSAPSession` (copiada del estilo de test_sap_upload.py) para verificar
la secuencia exacta de llamadas findById/.method.
"""

import sys
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import sox_report  # noqa: E402
from sox_report import (  # noqa: E402
    CAMPO_FECHA_DESDE,
    CAMPO_FECHA_HASTA,
    CAMPO_SOCIEDAD,
    DOCS_GRID_SHELL,
    SOX_NODE_KEY,
    TREE_SHELL,
    VALID_SOCIEDADES,
    abrir_transaccion_sox,
    exportar_a_excel,
    generar_reporte_sox,
    get_sap_session,
    ingresar_parametros,
    validar_caracter_fecha,
    validar_fecha,
    validar_rango_fechas,
    validar_sociedad,
)


# ---------------------------------------------------------------------------
# Mock de sesión SAP (replica el de test_sap_upload.py)
# ---------------------------------------------------------------------------


class MockSAPSession:
    def __init__(self):
        self._elements: dict = {}
        self.actions: list = []

    def findById(self, sap_id):
        if sap_id not in self._elements:
            self._elements[sap_id] = _MockElement(self, sap_id)
        return self._elements[sap_id]


class _MockElement:
    def __init__(self, session, sap_id):
        self._session = session
        self._sap_id = sap_id
        self.text = ""
        self.caretPosition = 0

    def press(self):
        self._session.actions.append((self._sap_id, "press"))

    def maximize(self):
        self._session.actions.append((self._sap_id, "maximize"))

    def setFocus(self):
        self._session.actions.append((self._sap_id, "setFocus"))

    def close(self):
        self._session.actions.append((self._sap_id, "close"))

    def sendVKey(self, key):
        self._session.actions.append((self._sap_id, "sendVKey", key))

    def doubleClickNode(self, key):
        self._session.actions.append((self._sap_id, "doubleClickNode", key))

    def pressToolbarContextButton(self, key):
        self._session.actions.append(
            (self._sap_id, "pressToolbarContextButton", key)
        )

    def selectContextMenuItem(self, key):
        self._session.actions.append(
            (self._sap_id, "selectContextMenuItem", key)
        )


# ---------------------------------------------------------------------------
# Validaciones
# ---------------------------------------------------------------------------


class ValidarSociedadTest(unittest.TestCase):
    def test_accepts_all_valid_sociedades(self):
        for soc in VALID_SOCIEDADES:
            self.assertEqual(validar_sociedad(soc), soc)

    def test_normalizes_to_uppercase(self):
        self.assertEqual(validar_sociedad("isa"), "ISA")
        self.assertEqual(validar_sociedad("  tran  "), "TRAN")

    def test_rejects_invalid_value(self):
        with self.assertRaises(ValueError) as ctx:
            validar_sociedad("XYZ")
        self.assertIn("XYZ", str(ctx.exception))

    def test_rejects_empty_string(self):
        with self.assertRaises(ValueError):
            validar_sociedad("")

    def test_rejects_only_whitespace(self):
        with self.assertRaises(ValueError):
            validar_sociedad("   ")

    def test_rejects_non_string(self):
        with self.assertRaises(ValueError):
            validar_sociedad(None)  # type: ignore[arg-type]


class ValidarFechaTest(unittest.TestCase):
    def test_accepts_valid_date(self):
        result = validar_fecha("01.05.2026")
        self.assertEqual(result, datetime(2026, 5, 1))

    def test_accepts_with_surrounding_whitespace(self):
        self.assertEqual(validar_fecha("  31.12.2026 "), datetime(2026, 12, 31))

    def test_rejects_wrong_format(self):
        with self.assertRaises(ValueError):
            validar_fecha("2026-05-01")
        with self.assertRaises(ValueError):
            validar_fecha("01/05/2026")

    def test_rejects_invalid_day(self):
        with self.assertRaises(ValueError):
            validar_fecha("32.01.2026")

    def test_rejects_invalid_month(self):
        with self.assertRaises(ValueError):
            validar_fecha("01.13.2026")

    def test_rejects_empty(self):
        with self.assertRaises(ValueError):
            validar_fecha("")

    def test_rejects_alphabetic(self):
        with self.assertRaises(ValueError):
            validar_fecha("ab.cd.efgh")


class ValidarRangoFechasTest(unittest.TestCase):
    def test_accepts_desde_lower_than_hasta(self):
        d, h = validar_rango_fechas("01.05.2026", "31.05.2026")
        self.assertEqual(d, datetime(2026, 5, 1))
        self.assertEqual(h, datetime(2026, 5, 31))

    def test_accepts_equal_dates(self):
        d, h = validar_rango_fechas("15.05.2026", "15.05.2026")
        self.assertEqual(d, h)

    def test_rejects_hasta_before_desde(self):
        with self.assertRaisesRegex(ValueError, "mayor o igual"):
            validar_rango_fechas("31.05.2026", "01.05.2026")

    def test_propagates_format_error_from_desde(self):
        with self.assertRaises(ValueError):
            validar_rango_fechas("bad", "31.05.2026")

    def test_propagates_format_error_from_hasta(self):
        with self.assertRaises(ValueError):
            validar_rango_fechas("01.05.2026", "bad")


class ValidarCaracterFechaTest(unittest.TestCase):
    def test_accepts_digits_and_dots(self):
        self.assertTrue(validar_caracter_fecha("01.05.2026"))
        self.assertTrue(validar_caracter_fecha("123"))
        self.assertTrue(validar_caracter_fecha(""))
        self.assertTrue(validar_caracter_fecha("...."))

    def test_rejects_letters(self):
        self.assertFalse(validar_caracter_fecha("01a"))
        self.assertFalse(validar_caracter_fecha("hola"))

    def test_rejects_special_characters(self):
        self.assertFalse(validar_caracter_fecha("01-05"))
        self.assertFalse(validar_caracter_fecha("01/05"))
        self.assertFalse(validar_caracter_fecha("01 05"))
        self.assertFalse(validar_caracter_fecha("01,05"))

    def test_rejects_more_than_10_characters(self):
        self.assertFalse(validar_caracter_fecha("01.05.20266"))
        self.assertFalse(validar_caracter_fecha("12345678901"))


# ---------------------------------------------------------------------------
# get_sap_session
# ---------------------------------------------------------------------------


class GetSapSessionTest(unittest.TestCase):
    def test_raises_when_pywin32_missing(self):
        with patch.dict(sys.modules, {"win32com": None, "win32com.client": None}):
            with self.assertRaises(RuntimeError) as ctx:
                get_sap_session()
        self.assertIn("pywin32", str(ctx.exception))

    def test_returns_session_on_success(self):
        session = MagicMock(name="session")
        connection = MagicMock()
        connection.Children.Count = 1
        connection.Children.return_value = session
        engine = MagicMock()
        engine.Children.Count = 1
        engine.Children.return_value = connection
        sap_gui_auto = MagicMock()
        sap_gui_auto.GetScriptingEngine = engine
        fake_win32 = MagicMock()
        fake_win32.client.GetObject.return_value = sap_gui_auto

        with patch.dict(sys.modules, {
            "win32com": fake_win32,
            "win32com.client": fake_win32.client,
        }):
            result = get_sap_session()
        self.assertIs(result, session)


# ---------------------------------------------------------------------------
# Pasos del flujo SOX
# ---------------------------------------------------------------------------


class AbrirTransaccionSoxTest(unittest.TestCase):
    def setUp(self):
        # Asegurar que T_CODE_SOX está en None (modo árbol) para los tests
        # del comportamiento por defecto. Tests específicos del modo T-code
        # patchean este valor.
        self._t_code_original = sox_report.T_CODE_SOX
        sox_report.T_CODE_SOX = None

    def tearDown(self):
        sox_report.T_CODE_SOX = self._t_code_original

    def test_maximizes_and_double_clicks_node_when_no_tcode(self):
        session = MockSAPSession()
        abrir_transaccion_sox(session)

        self.assertIn(("wnd[0]", "maximize"), session.actions)
        self.assertIn(
            (TREE_SHELL, "doubleClickNode", SOX_NODE_KEY), session.actions
        )

    def test_uses_okcd_when_tcode_configured(self):
        """Cuando T_CODE_SOX está configurado, navega por la casilla de
        comandos en vez de tocar el árbol — más robusto."""
        sox_report.T_CODE_SOX = "ZSOX_REPORT"
        session = MockSAPSession()
        abrir_transaccion_sox(session)

        self.assertEqual(
            session._elements["wnd[0]/tbar[0]/okcd"].text, "ZSOX_REPORT"
        )
        self.assertIn(("wnd[0]", "sendVKey", 0), session.actions)
        # No se debe haber tocado el árbol
        self.assertFalse(
            any(a[0] == TREE_SHELL for a in session.actions),
            "No debería tocar el árbol cuando T_CODE_SOX está configurado",
        )

    def test_error_when_node_missing_includes_tree_diagnostic(self):
        """Cuando el doubleClickNode falla, el mensaje debe incluir
        pistas para resolver el problema y los nodos disponibles."""
        session = MockSAPSession()
        tree = session.findById(TREE_SHELL)
        tree.doubleClickNode = MagicMock(side_effect=Exception("node missing"))
        # Simular que el árbol expone GetAllNodeKeys/GetNodeTextByKey
        tree.GetAllNodeKeys = MagicMock(return_value=["F00001", "F00002"])
        tree.GetNodeTextByKey = MagicMock(side_effect=lambda k: f"Texto-{k}")

        with self.assertRaises(RuntimeError) as ctx:
            abrir_transaccion_sox(session)
        mensaje = str(ctx.exception)
        # Pista sobre T_CODE_SOX
        self.assertIn("T_CODE_SOX", mensaje)
        # Lista de nodos disponibles en el árbol
        self.assertIn("F00001", mensaje)
        self.assertIn("Texto-F00001", mensaje)
        self.assertIn("F00002", mensaje)

    def test_error_when_node_missing_handles_tree_enumeration_failure(self):
        """Si GetAllNodeKeys también falla, el diagnóstico no debe explotar."""
        session = MockSAPSession()
        tree = session.findById(TREE_SHELL)
        tree.doubleClickNode = MagicMock(side_effect=Exception("node missing"))
        tree.GetAllNodeKeys = MagicMock(side_effect=Exception("API no disponible"))

        with self.assertRaises(RuntimeError) as ctx:
            abrir_transaccion_sox(session)
        mensaje = str(ctx.exception)
        self.assertIn("T_CODE_SOX", mensaje)
        # El bloque de nodos debe estar presente aunque vacío/con error
        self.assertIn("no se pudo enumerar el árbol", mensaje)


class IngresarParametrosTest(unittest.TestCase):
    """Verifica el nuevo flujo del Script2sox.vbs: sociedad por texto, fechas
    vía calendario F4 (focusDate + selectionInterval en formato yyyymmdd)."""

    def test_sets_sociedad_as_text(self):
        session = MockSAPSession()
        ingresar_parametros(session, "ISA", "01.05.2026", "31.05.2026")

        self.assertEqual(session._elements[CAMPO_SOCIEDAD].text, "ISA")

    def test_opens_calendar_for_each_date_field(self):
        session = MockSAPSession()
        ingresar_parametros(session, "ISA", "01.05.2026", "31.05.2026")

        # Foco y caretPosition=0 en el campo Desde antes de F4
        self.assertIn((CAMPO_FECHA_DESDE, "setFocus"), session.actions)
        self.assertEqual(session._elements[CAMPO_FECHA_DESDE].caretPosition, 0)
        # Foco y caretPosition=0 en el campo Hasta antes de F4
        self.assertIn((CAMPO_FECHA_HASTA, "setFocus"), session.actions)
        self.assertEqual(session._elements[CAMPO_FECHA_HASTA].caretPosition, 0)
        # F4 (sendVKey 4) se envió dos veces: una para cada fecha
        f4_calls = [a for a in session.actions if a == ("wnd[0]", "sendVKey", 4)]
        self.assertEqual(len(f4_calls), 2)

    def test_sets_calendar_focus_and_selection_in_yyyymmdd_format(self):
        session = MockSAPSession()
        ingresar_parametros(session, "ISA", "01.05.2026", "31.05.2026")

        from sox_report import CALENDAR_SHELL
        calendario = session._elements[CALENDAR_SHELL]
        # El selectionInterval queda con la última fecha procesada (Hasta).
        # focusDate y selectionInterval se asignan dos veces (una por
        # cada fecha) — el valor final refleja la fecha hasta.
        self.assertEqual(calendario.focusDate, "20260531")
        self.assertEqual(calendario.selectionInterval, "20260531,20260531")

    def test_raises_when_date_format_invalid(self):
        session = MockSAPSession()
        with self.assertRaises(ValueError):
            ingresar_parametros(session, "ISA", "no-fecha", "31.05.2026")

    def test_presses_f8_to_execute(self):
        session = MockSAPSession()
        ingresar_parametros(session, "ISA", "01.05.2026", "31.05.2026")

        self.assertIn(("wnd[0]/tbar[1]/btn[8]", "press"), session.actions)


class ExportarAExcelTest(unittest.TestCase):
    def test_invokes_export_xxl_menu(self):
        session = MockSAPSession()
        exportar_a_excel(session, r"C:\salida", "SOX_ISA.xlsx")

        self.assertIn(
            (DOCS_GRID_SHELL, "pressToolbarContextButton", "&MB_EXPORT"),
            session.actions,
        )
        self.assertIn(
            (DOCS_GRID_SHELL, "selectContextMenuItem", "&XXL"),
            session.actions,
        )

    def test_fills_save_dialog_when_available(self):
        session = MockSAPSession()
        exportar_a_excel(session, r"C:\salida", "SOX_ISA_x.xlsx")

        self.assertEqual(
            session._elements["wnd[1]/usr/ctxtDY_PATH"].text, r"C:\salida"
        )
        self.assertEqual(
            session._elements["wnd[1]/usr/ctxtDY_FILENAME"].text, "SOX_ISA_x.xlsx"
        )
        self.assertEqual(
            session._elements["wnd[1]/usr/ctxtDY_FILENAME"].caretPosition,
            len("SOX_ISA_x.xlsx"),
        )
        self.assertIn(("wnd[1]/tbar[0]/btn[0]", "press"), session.actions)

    def test_falls_back_to_close_when_dialog_unavailable(self):
        # Sesión donde wnd[1]/usr/ctxtDY_PATH no se puede acceder: forzamos
        # que findById lance excepción para esa ruta específica.
        session = MockSAPSession()
        original_find = session.findById

        def find_with_error(sap_id):
            if sap_id == "wnd[1]/usr/ctxtDY_PATH":
                raise Exception("dialog inexistente")
            return original_find(sap_id)

        session.findById = find_with_error
        exportar_a_excel(session, r"C:\salida", "x.xlsx")

        # El recording original cierra wnd[1] como fallback.
        self.assertIn(("wnd[1]", "close"), session.actions)


# ---------------------------------------------------------------------------
# generar_reporte_sox (orquestador)
# ---------------------------------------------------------------------------


class StepErrorContextTest(unittest.TestCase):
    """Verifica que cuando una operación SAP falla durante los pasos del flujo
    SOX, la excepción re-lanzada contiene contexto suficiente para identificar
    la línea exacta que falló (clave porque las excepciones COM del SAP
    Frontend Server traen descripción vacía).
    """

    def setUp(self):
        # Algunos tests verifican el fallback al árbol — forzamos T_CODE_SOX
        # a None para que ese camino se ejecute. Los tests que verifican el
        # modo T-code patchean explícitamente.
        self._t_code_original = sox_report.T_CODE_SOX
        sox_report.T_CODE_SOX = None

    def tearDown(self):
        sox_report.T_CODE_SOX = self._t_code_original

    def test_abrir_transaccion_raises_with_context_when_maximize_fails(self):
        session = MockSAPSession()
        wnd = session.findById("wnd[0]")
        wnd.maximize = MagicMock(side_effect=Exception("COM error"))

        with self.assertRaisesRegex(RuntimeError, "Maximizar"):
            abrir_transaccion_sox(session)

    def test_abrir_transaccion_raises_with_context_when_tree_not_found(self):
        session = MockSAPSession()
        original = session.findById

        def find_with_error(sap_id):
            if sap_id == TREE_SHELL:
                raise Exception("tree not found")
            return original(sap_id)

        session.findById = find_with_error

        with self.assertRaises(RuntimeError) as ctx:
            abrir_transaccion_sox(session)
        # Mensaje incluye la ruta del árbol y pista para el usuario
        self.assertIn(TREE_SHELL, str(ctx.exception))
        self.assertIn("Easy Access", str(ctx.exception))

    def test_abrir_transaccion_raises_with_context_when_node_not_found(self):
        session = MockSAPSession()
        tree = session.findById(TREE_SHELL)
        tree.doubleClickNode = MagicMock(side_effect=Exception("node missing"))

        with self.assertRaises(RuntimeError) as ctx:
            abrir_transaccion_sox(session)
        msg = str(ctx.exception)
        self.assertIn(SOX_NODE_KEY, msg)
        # Incluye una pista para el usuario
        self.assertIn("menú", msg.lower())

    def test_ingresar_parametros_raises_when_sociedad_field_missing(self):
        session = MockSAPSession()
        original = session.findById

        def find_with_error(sap_id):
            if sap_id == CAMPO_SOCIEDAD:
                raise Exception("field missing")
            return original(sap_id)

        session.findById = find_with_error

        with self.assertRaisesRegex(RuntimeError, "Sociedad"):
            ingresar_parametros(session, "ISA", "01.05.2026", "31.05.2026")

    def test_ingresar_parametros_raises_when_f8_button_missing(self):
        session = MockSAPSession()
        original = session.findById

        def find_with_error(sap_id):
            if sap_id == "wnd[0]/tbar[1]/btn[8]":
                raise Exception("button missing")
            return original(sap_id)

        session.findById = find_with_error

        with self.assertRaisesRegex(RuntimeError, "Ejecutar"):
            ingresar_parametros(session, "ISA", "01.05.2026", "31.05.2026")

    def test_exportar_raises_when_grid_not_found(self):
        session = MockSAPSession()
        original = session.findById

        def find_with_error(sap_id):
            if sap_id == DOCS_GRID_SHELL:
                raise Exception("grid missing")
            return original(sap_id)

        session.findById = find_with_error

        with self.assertRaisesRegex(RuntimeError, "grid de resultados"):
            exportar_a_excel(session, r"C:\salida", "x.xlsx")


class GenerarReporteSoxTest(unittest.TestCase):
    def test_calls_all_steps_in_order(self):
        session = MockSAPSession()
        call_order = []

        def make_recorder(name):
            return lambda *args, **kwargs: call_order.append(name)

        with patch.multiple(
            "sox_report",
            abrir_transaccion_sox=make_recorder("abrir"),
            ingresar_parametros=make_recorder("ingresar"),
            exportar_a_excel=make_recorder("exportar"),
        ):
            generar_reporte_sox(
                session, "ISA", "01.05.2026", "31.05.2026",
                carpeta_destino="/tmp", nombre_archivo="x.xlsx",
            )

        self.assertEqual(call_order, ["abrir", "ingresar", "exportar"])

    def test_normalizes_sociedad_before_passing(self):
        session = MockSAPSession()
        with patch("sox_report.ingresar_parametros") as mock_ing, \
             patch("sox_report.abrir_transaccion_sox"), \
             patch("sox_report.exportar_a_excel"):
            generar_reporte_sox(
                session, "isa", "01.05.2026", "31.05.2026",
                carpeta_destino="/tmp", nombre_archivo="x.xlsx",
            )

        mock_ing.assert_called_once_with(
            session, "ISA", "01.05.2026", "31.05.2026"
        )

    def test_raises_for_invalid_sociedad(self):
        session = MockSAPSession()
        with self.assertRaises(ValueError):
            generar_reporte_sox(
                session, "XYZ", "01.05.2026", "31.05.2026",
            )

    def test_raises_for_invalid_date_range(self):
        session = MockSAPSession()
        with self.assertRaises(ValueError):
            generar_reporte_sox(
                session, "ISA", "31.05.2026", "01.05.2026",
            )

    def test_default_filename_includes_sociedad_and_timestamp(self):
        session = MockSAPSession()
        with patch("sox_report.abrir_transaccion_sox"), \
             patch("sox_report.ingresar_parametros"), \
             patch("sox_report.exportar_a_excel") as mock_export:
            carpeta, nombre = generar_reporte_sox(
                session, "ISA", "01.05.2026", "31.05.2026",
            )

        self.assertTrue(nombre.startswith("SOX_ISA_"))
        self.assertTrue(nombre.endswith(".xlsx"))
        self.assertRegex(nombre, r"^SOX_ISA_\d{8}_\d{6}\.xlsx$")
        # carpeta debe apuntar al directorio salida/ por default
        self.assertTrue(carpeta.endswith("salida"))
        mock_export.assert_called_once_with(session, carpeta, nombre)


# ---------------------------------------------------------------------------
# main() entry point
# ---------------------------------------------------------------------------


class MainEntryPointTest(unittest.TestCase):
    def test_returns_2_when_wrong_argument_count(self):
        self.assertEqual(sox_report.main(["ISA"]), 2)
        self.assertEqual(sox_report.main([]), 2)
        self.assertEqual(sox_report.main(["a", "b", "c", "d"]), 2)

    def test_returns_1_when_invalid_sociedad(self):
        self.assertEqual(
            sox_report.main(["XYZ", "01.05.2026", "31.05.2026"]), 1
        )

    def test_returns_1_when_invalid_date_range(self):
        self.assertEqual(
            sox_report.main(["ISA", "31.05.2026", "01.05.2026"]), 1
        )

    def test_returns_1_when_sap_session_fails(self):
        with patch(
            "sox_report.get_sap_session",
            side_effect=RuntimeError("SAP no abierto"),
        ):
            self.assertEqual(
                sox_report.main(["ISA", "01.05.2026", "31.05.2026"]), 1
            )

    def test_returns_0_on_happy_path(self):
        with patch("sox_report.get_sap_session", return_value=MagicMock()), \
             patch(
                 "sox_report.generar_reporte_sox",
                 return_value=("/tmp", "SOX_ISA.xlsx"),
             ) as mock_flow:
            self.assertEqual(
                sox_report.main(["ISA", "01.05.2026", "31.05.2026"]), 0
            )
            mock_flow.assert_called_once()


if __name__ == "__main__":
    unittest.main()
