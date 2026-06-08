"""Pruebas unitarias para extraer_activos_creados.py.

Replica el estilo de test_sox_report.py / test_sap_upload.py: usa un
`MockSAPSession` para verificar la secuencia exacta de llamadas
findById(...).method() sin necesidad de un SAP real.
"""

import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import extraer_activos_creados  # noqa: E402
from extraer_activos_creados import (  # noqa: E402
    BTN_CONFIRMAR_WND1,
    BTN_EXPORTAR_TBAR0,
    BTN_EXPORTAR_TBAR1,
    CAMPO_CREATOR,
    CAMPO_DY_FILENAME,
    CAMPO_DY_PATH,
    CELDA_PRIMER_REGISTRO,
    NOMBRE_EXTENSION,
    NOMBRE_PREFIJO,
    T_CODE_SM35P,
    _nombre_archivo_extraccion,
    abrir_primer_registro,
    abrir_sm35p,
    exportar_log,
    extraer_activos_creados as orquestador,
    filtrar_por_usuario,
    get_sap_session,
    validar_usuario_sap,
)


# ---------------------------------------------------------------------------
# Mock de sesión SAP (replica el de test_sap_upload.py / test_sox_report.py)
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

    def setFocus(self):
        self._session.actions.append((self._sap_id, "setFocus"))

    def maximize(self):
        self._session.actions.append((self._sap_id, "maximize"))

    def sendVKey(self, key):
        self._session.actions.append((self._sap_id, "sendVKey", key))

    def __setattr__(self, name, value):
        if name in ("_session", "_sap_id"):
            super().__setattr__(name, value)
            return
        super().__setattr__(name, value)
        if name in ("text", "caretPosition"):
            self._session.actions.append(
                (self._sap_id, f"set_{name}", value)
            )


# ---------------------------------------------------------------------------
# validar_usuario_sap
# ---------------------------------------------------------------------------


class ValidarUsuarioSapTest(unittest.TestCase):
    def test_accepts_numeric_id(self):
        self.assertEqual(validar_usuario_sap("1017209574"), "1017209574")

    def test_accepts_alphanumeric_id_preserving_casing(self):
        """IDs como INTC37089 deben pasar tal cual (no se fuerza casing)."""
        self.assertEqual(validar_usuario_sap("INTC37089"), "INTC37089")

    def test_strips_surrounding_whitespace(self):
        self.assertEqual(validar_usuario_sap("  USR123  "), "USR123")

    def test_rejects_empty_string(self):
        with self.assertRaises(ValueError):
            validar_usuario_sap("")

    def test_rejects_only_whitespace(self):
        with self.assertRaises(ValueError):
            validar_usuario_sap("   ")

    def test_rejects_non_string(self):
        for value in [None, 123, [], {}]:
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    validar_usuario_sap(value)


# ---------------------------------------------------------------------------
# get_sap_session
# ---------------------------------------------------------------------------


class GetSapSessionTest(unittest.TestCase):
    def test_raises_when_pywin32_missing(self):
        with patch.dict("sys.modules", {"win32com": None, "win32com.client": None}):
            with self.assertRaisesRegex(RuntimeError, "pywin32"):
                get_sap_session()

    def test_returns_first_session_when_sap_ok(self):
        fake_session = MagicMock()
        fake_application = MagicMock()
        fake_application.Children.Count = 1
        fake_connection = MagicMock()
        fake_connection.Children.Count = 1
        fake_connection.Children.return_value = fake_session
        fake_application.Children.return_value = fake_connection
        fake_sap_gui = MagicMock()
        fake_sap_gui.GetScriptingEngine = fake_application

        fake_win32com = MagicMock()
        fake_win32com.client.GetObject.return_value = fake_sap_gui

        with patch.dict("sys.modules", {
            "win32com": fake_win32com,
            "win32com.client": fake_win32com.client,
        }):
            result = get_sap_session()

        self.assertIs(result, fake_session)


# ---------------------------------------------------------------------------
# abrir_sm35p
# ---------------------------------------------------------------------------


class AbrirSm35pTest(unittest.TestCase):
    def test_maximizes_and_sends_okcd(self):
        session = MockSAPSession()
        abrir_sm35p(session)

        self.assertIn(("wnd[0]", "maximize"), session.actions)
        self.assertEqual(
            session._elements["wnd[0]/tbar[0]/okcd"].text, T_CODE_SM35P
        )
        self.assertIn(("wnd[0]", "sendVKey", 0), session.actions)

    def test_order_is_maximize_then_okcd_then_enter(self):
        session = MockSAPSession()
        abrir_sm35p(session)

        # Extraer las acciones relevantes en orden.
        idx_max = next(
            i for i, a in enumerate(session.actions) if a == ("wnd[0]", "maximize")
        )
        idx_text = next(
            i for i, a in enumerate(session.actions)
            if a == ("wnd[0]/tbar[0]/okcd", "set_text", T_CODE_SM35P)
        )
        idx_enter = next(
            i for i, a in enumerate(session.actions)
            if a == ("wnd[0]", "sendVKey", 0)
        )
        self.assertLess(idx_max, idx_text)
        self.assertLess(idx_text, idx_enter)


# ---------------------------------------------------------------------------
# filtrar_por_usuario
# ---------------------------------------------------------------------------


class FiltrarPorUsuarioTest(unittest.TestCase):
    def test_sets_creator_with_wildcard_prefix(self):
        session = MockSAPSession()
        filtrar_por_usuario(session, "1017209574")

        self.assertEqual(
            session._elements[CAMPO_CREATOR].text, "*1017209574"
        )

    def test_focuses_creator_and_sends_enter(self):
        session = MockSAPSession()
        filtrar_por_usuario(session, "USR1")

        self.assertIn((CAMPO_CREATOR, "setFocus"), session.actions)
        self.assertIn(("wnd[0]", "sendVKey", 0), session.actions)

    def test_caret_position_matches_filter_length(self):
        """caretPosition se setea al len(*<usuario>) para dejar el cursor
        al final del texto, replicando el recording original."""
        session = MockSAPSession()
        filtrar_por_usuario(session, "1017209574")

        # filtro = "*1017209574" tiene 11 chars
        self.assertEqual(
            session._elements[CAMPO_CREATOR].caretPosition, 11
        )


# ---------------------------------------------------------------------------
# abrir_primer_registro
# ---------------------------------------------------------------------------


class AbrirPrimerRegistroTest(unittest.TestCase):
    def test_focuses_first_row_cell_and_sends_f2(self):
        session = MockSAPSession()
        abrir_primer_registro(session)

        self.assertIn((CELDA_PRIMER_REGISTRO, "setFocus"), session.actions)
        # F2 = sendVKey 2
        self.assertIn(("wnd[0]", "sendVKey", 2), session.actions)

    def test_sets_caret_position_in_cell(self):
        session = MockSAPSession()
        abrir_primer_registro(session)

        self.assertEqual(
            session._elements[CELDA_PRIMER_REGISTRO].caretPosition, 5
        )


# ---------------------------------------------------------------------------
# exportar_log
# ---------------------------------------------------------------------------


class ExportarLogTest(unittest.TestCase):
    def test_presses_export_toolbar_chain(self):
        session = MockSAPSession()
        exportar_log(session, r"C:\salida", "ActivosCreados_USR_x.xlsx")

        self.assertIn((BTN_EXPORTAR_TBAR0, "press"), session.actions)
        self.assertIn((BTN_EXPORTAR_TBAR1, "press"), session.actions)
        self.assertIn((BTN_CONFIRMAR_WND1, "press"), session.actions)

    def test_skips_f4_picker_and_uses_direct_dy_path(self):
        """A diferencia del recording, exportar_log NO usa F4 + picker
        (wnd[2]). Inyecta DY_PATH y DY_FILENAME directamente en wnd[1] y
        confirma con btn[11]. Esto da control total del path de salida."""
        session = MockSAPSession()
        exportar_log(session, r"C:\salida", "ActivosCreados_USR_x.xlsx")

        # No debe haber F4 en wnd[1] ni press en wnd[2]/btn[11]
        self.assertNotIn(("wnd[1]", "sendVKey", 4), session.actions)
        self.assertNotIn(("wnd[2]/tbar[0]/btn[11]", "press"), session.actions)

    def test_fills_dy_path_and_dy_filename(self):
        session = MockSAPSession()
        exportar_log(session, r"C:\salida", "ActivosCreados_USR_x.xlsx")

        self.assertEqual(
            session._elements[CAMPO_DY_PATH].text, r"C:\salida"
        )
        self.assertEqual(
            session._elements[CAMPO_DY_FILENAME].text, "ActivosCreados_USR_x.xlsx"
        )
        self.assertEqual(
            session._elements[CAMPO_DY_FILENAME].caretPosition,
            len("ActivosCreados_USR_x.xlsx"),
        )

    def test_orden_secuencial(self):
        """btn[86] → btn[43] → set DY_PATH/DY_FILENAME → btn[11]."""
        session = MockSAPSession()
        exportar_log(session, r"C:\out", "x.xlsx")

        def idx(action_tuple):
            return session.actions.index(action_tuple)

        self.assertLess(idx((BTN_EXPORTAR_TBAR0, "press")),
                        idx((BTN_EXPORTAR_TBAR1, "press")))
        self.assertLess(idx((BTN_EXPORTAR_TBAR1, "press")),
                        idx((CAMPO_DY_PATH, "set_text", r"C:\out")))
        self.assertLess(idx((CAMPO_DY_PATH, "set_text", r"C:\out")),
                        idx((BTN_CONFIRMAR_WND1, "press")))


# ---------------------------------------------------------------------------
# extraer_activos_creados (orquestador)
# ---------------------------------------------------------------------------


class NombreArchivoExtraccionTest(unittest.TestCase):
    def test_pattern_is_prefix_usuario_timestamp_ext(self):
        nombre = _nombre_archivo_extraccion("1017209574")
        self.assertTrue(nombre.startswith(f"{NOMBRE_PREFIJO}_1017209574_"))
        self.assertTrue(nombre.endswith(NOMBRE_EXTENSION))

    def test_timestamp_format_is_yyyymmdd_hhmmss(self):
        nombre = _nombre_archivo_extraccion("USR1")
        # Patrón: ActivosCreados_USR1_YYYYMMDD_HHMMSS.xlsx
        import re
        self.assertRegex(
            nombre,
            r"^ActivosCreados_USR1_\d{8}_\d{6}\.xlsx$",
        )


class ExtraerActivosCreadosTest(unittest.TestCase):
    def test_calls_all_4_steps_in_order(self):
        session = MockSAPSession()
        call_order = []

        def make_recorder(name):
            return lambda *args, **kwargs: call_order.append(name)

        with patch.multiple(
            "extraer_activos_creados",
            abrir_sm35p=make_recorder("abrir"),
            filtrar_por_usuario=make_recorder("filtrar"),
            abrir_primer_registro=make_recorder("primer_registro"),
            exportar_log=make_recorder("exportar"),
        ):
            orquestador(session, "1017209574")

        self.assertEqual(
            call_order,
            ["abrir", "filtrar", "primer_registro", "exportar"],
        )

    def test_normalizes_usuario_before_passing_to_filtrar(self):
        """El usuario se pasa con strip() pero sin transformación de casing."""
        session = MockSAPSession()
        with patch("extraer_activos_creados.abrir_sm35p"), \
             patch("extraer_activos_creados.filtrar_por_usuario") as mock_filtrar, \
             patch("extraer_activos_creados.abrir_primer_registro"), \
             patch("extraer_activos_creados.exportar_log"):
            orquestador(session, "  INTC37089  ")

        mock_filtrar.assert_called_once_with(session, "INTC37089")

    def test_raises_for_invalid_usuario(self):
        session = MockSAPSession()
        with self.assertRaises(ValueError):
            orquestador(session, "")

    def test_returns_carpeta_and_nombre(self):
        """Devuelve la tupla (carpeta, nombre) usadas. Por default,
        carpeta = SALIDA_DIR y nombre = ActivosCreados_USR_TS.xlsx."""
        session = MockSAPSession()
        with patch("extraer_activos_creados.abrir_sm35p"), \
             patch("extraer_activos_creados.filtrar_por_usuario"), \
             patch("extraer_activos_creados.abrir_primer_registro"), \
             patch("extraer_activos_creados.exportar_log"):
            carpeta, nombre = orquestador(session, "1017209574")

        self.assertTrue(carpeta.endswith("salida"))
        self.assertRegex(
            nombre,
            r"^ActivosCreados_1017209574_\d{8}_\d{6}\.xlsx$",
        )

    def test_passes_carpeta_and_nombre_to_exportar_log(self):
        """exportar_log recibe la carpeta y el nombre del archivo final."""
        session = MockSAPSession()
        with patch("extraer_activos_creados.abrir_sm35p"), \
             patch("extraer_activos_creados.filtrar_por_usuario"), \
             patch("extraer_activos_creados.abrir_primer_registro"), \
             patch("extraer_activos_creados.exportar_log") as mock_exportar:
            orquestador(
                session, "USR1",
                carpeta_destino="/tmp/out", nombre_archivo="custom.xlsx",
            )

        mock_exportar.assert_called_once_with(
            session, "/tmp/out", "custom.xlsx"
        )


# ---------------------------------------------------------------------------
# main() entry point
# ---------------------------------------------------------------------------


class MainEntryPointTest(unittest.TestCase):
    def test_returns_2_when_wrong_argument_count(self):
        self.assertEqual(extraer_activos_creados.main([]), 2)
        self.assertEqual(extraer_activos_creados.main(["a", "b"]), 2)

    def test_returns_1_on_validation_error(self):
        self.assertEqual(extraer_activos_creados.main(["   "]), 1)

    def test_returns_1_when_sap_unavailable(self):
        with patch("extraer_activos_creados.get_sap_session",
                   side_effect=RuntimeError("no SAP")):
            self.assertEqual(extraer_activos_creados.main(["USR1"]), 1)

    def test_returns_0_on_happy_path(self):
        with patch("extraer_activos_creados.get_sap_session",
                   return_value=MagicMock()), \
             patch("extraer_activos_creados.extraer_activos_creados",
                   return_value=("/tmp/salida", "x.xlsx")):
            self.assertEqual(extraer_activos_creados.main(["USR1"]), 0)


if __name__ == "__main__":
    unittest.main()
