"""Pruebas unitarias para sap_upload.py.

Las funciones que dialogan con SAP GUI Scripting se prueban inyectando
MockSAPSession, una clase que registra cada llamada a findById/.method
para que los tests verifiquen la secuencia exacta de acciones.
"""

import os
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import sap_upload  # noqa: E402
from sap_upload import (  # noqa: E402
    ASSIGN_FILES_ROW,
    BDC_SESSION_TABLE,
    DEFAULT_SELECTED_ROW,
    LSMW_STEPLIST_TABLE,
    READ_DATA_ROW,
    SPECIFY_FILES_ROW,
    configurar_ruta_archivo,
    get_latest_txt,
    get_sap_session,
    open_lsmw,
    process_bdc_session,
    run_lsmw_flow,
    select_step_row,
    step_assign_files,
    step_convert_data,
    step_create_batch_input,
    step_display_converted_data,
    step_display_read_data,
    step_read_data,
    step_run_batch_input,
)


# ---------------------------------------------------------------------------
# Mock de sesión SAP
# ---------------------------------------------------------------------------


class MockSAPSession:
    """Mock que registra cada llamada findById/.method/.attr para inspección."""

    def __init__(self):
        self._elements: dict[str, "_MockElement"] = {}
        self.actions: list[tuple] = []

    def findById(self, sap_id):
        if sap_id not in self._elements:
            self._elements[sap_id] = _MockElement(self, sap_id)
        return self._elements[sap_id]


class _MockElement:
    def __init__(self, session: MockSAPSession, sap_id: str):
        self._session = session
        self._sap_id = sap_id
        self._rows: dict[int, _MockRow] = {}
        self.text = ""
        self.caretPosition = 0
        self.selected = False

    def press(self):
        self._session.actions.append((self._sap_id, "press"))

    def maximize(self):
        self._session.actions.append((self._sap_id, "maximize"))

    def setFocus(self):
        self._session.actions.append((self._sap_id, "setFocus"))

    def select(self):
        self._session.actions.append((self._sap_id, "select"))

    def sendVKey(self, key):
        self._session.actions.append((self._sap_id, "sendVKey", key))

    def getAbsoluteRow(self, idx):
        if idx not in self._rows:
            self._rows[idx] = _MockRow(self._session, self._sap_id, idx)
        return self._rows[idx]


class _MockRow:
    def __init__(self, session: MockSAPSession, parent_id: str, row_idx: int):
        self._session = session
        self._parent_id = parent_id
        self._row_idx = row_idx
        self._selected = False

    @property
    def selected(self):
        return self._selected

    @selected.setter
    def selected(self, value):
        self._selected = value
        self._session.actions.append(("row_selected", self._parent_id, self._row_idx, value))


# ---------------------------------------------------------------------------
# get_latest_txt
# ---------------------------------------------------------------------------


class GetLatestTxtTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.salida = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def test_raises_when_directory_missing(self):
        with self.assertRaises(FileNotFoundError):
            get_latest_txt(self.salida / "no_existe")

    def test_raises_when_no_lsmw_files(self):
        (self.salida / "otro.txt").write_text("x", encoding="utf-8")
        with self.assertRaises(FileNotFoundError):
            get_latest_txt(self.salida)

    def test_returns_most_recent_file_by_mtime(self):
        antiguo = self.salida / "LSMW_20260101_120000.txt"
        reciente = self.salida / "LSMW_20260102_120000.txt"
        antiguo.write_text("a", encoding="utf-8")
        reciente.write_text("b", encoding="utf-8")
        now = time.time()
        os.utime(antiguo, (now - 1000, now - 1000))
        os.utime(reciente, (now, now))

        self.assertEqual(get_latest_txt(self.salida), reciente)

    def test_ignores_files_not_matching_pattern(self):
        (self.salida / "otro.txt").write_text("x", encoding="utf-8")
        lsmw = self.salida / "LSMW_20260101_120000.txt"
        lsmw.write_text("ok", encoding="utf-8")

        self.assertEqual(get_latest_txt(self.salida), lsmw)


# ---------------------------------------------------------------------------
# get_sap_session
# ---------------------------------------------------------------------------


class GetSapSessionTest(unittest.TestCase):
    """Mockea sys.modules para inyectar un win32com falso."""

    def _make_fake_win32(self, engine):
        fake_win32 = MagicMock()
        sap_gui_auto = MagicMock()
        sap_gui_auto.GetScriptingEngine = engine
        fake_win32.client.GetObject.return_value = sap_gui_auto
        return fake_win32

    def test_raises_when_pywin32_not_installed(self):
        with patch.dict(sys.modules, {"win32com": None, "win32com.client": None}):
            with self.assertRaises(RuntimeError) as ctx:
                get_sap_session()
        self.assertIn("pywin32", str(ctx.exception))

    def test_raises_when_sap_gui_not_running(self):
        fake_win32 = MagicMock()
        fake_win32.client.GetObject.side_effect = Exception("COM error")
        with patch.dict(sys.modules, {
            "win32com": fake_win32,
            "win32com.client": fake_win32.client,
        }):
            with self.assertRaises(RuntimeError) as ctx:
                get_sap_session()
        self.assertIn("SAP GUI", str(ctx.exception))

    def test_raises_when_no_connections_active(self):
        engine = MagicMock()
        engine.Children.Count = 0
        fake_win32 = self._make_fake_win32(engine)
        with patch.dict(sys.modules, {
            "win32com": fake_win32,
            "win32com.client": fake_win32.client,
        }):
            with self.assertRaises(RuntimeError) as ctx:
                get_sap_session()
        self.assertIn("conexiones", str(ctx.exception).lower())

    def test_raises_when_no_sessions_active(self):
        connection = MagicMock()
        connection.Children.Count = 0
        engine = MagicMock()
        engine.Children.Count = 1
        engine.Children.return_value = connection
        fake_win32 = self._make_fake_win32(engine)
        with patch.dict(sys.modules, {
            "win32com": fake_win32,
            "win32com.client": fake_win32.client,
        }):
            with self.assertRaises(RuntimeError) as ctx:
                get_sap_session()
        self.assertIn("sesiones", str(ctx.exception).lower())

    def test_returns_first_session_on_success(self):
        session = MagicMock(name="session")
        connection = MagicMock()
        connection.Children.Count = 1
        connection.Children.return_value = session
        engine = MagicMock()
        engine.Children.Count = 1
        engine.Children.return_value = connection
        fake_win32 = self._make_fake_win32(engine)
        with patch.dict(sys.modules, {
            "win32com": fake_win32,
            "win32com.client": fake_win32.client,
        }):
            result = get_sap_session()
        self.assertIs(result, session)


# ---------------------------------------------------------------------------
# Pasos individuales del flujo LSMW
# ---------------------------------------------------------------------------


class DiagnosticarConexionSapTest(unittest.TestCase):
    """Pruebas para `diagnosticar_conexion_sap`, el helper del botón
    'Test conexión SAP' que devuelve (ok, mensaje detallado)."""

    def _make_fake_win32(self, engine):
        fake_win32 = MagicMock()
        sap_gui_auto = MagicMock()
        sap_gui_auto.GetScriptingEngine = engine
        fake_win32.client.GetObject.return_value = sap_gui_auto
        return fake_win32

    def test_returns_false_when_pywin32_missing(self):
        with patch.dict(sys.modules, {"win32com": None, "win32com.client": None}):
            ok, mensaje = sap_upload.diagnosticar_conexion_sap()
        self.assertFalse(ok)
        self.assertIn("pywin32", mensaje)
        self.assertIn("pip install pywin32", mensaje)

    def test_returns_false_when_sap_gui_not_running(self):
        fake_win32 = MagicMock()
        fake_win32.client.GetObject.side_effect = Exception("class not registered")
        with patch.dict(sys.modules, {
            "win32com": fake_win32,
            "win32com.client": fake_win32.client,
        }):
            ok, mensaje = sap_upload.diagnosticar_conexion_sap()
        self.assertFalse(ok)
        self.assertIn("COM", mensaje)
        self.assertIn("class not registered", mensaje)

    def test_returns_false_when_scripting_engine_fails(self):
        sap_gui_auto = MagicMock()
        # GetScriptingEngine returns a property-like object that raises
        type(sap_gui_auto).GetScriptingEngine = property(
            lambda self: (_ for _ in ()).throw(Exception("scripting disabled"))
        )
        fake_win32 = MagicMock()
        fake_win32.client.GetObject.return_value = sap_gui_auto
        with patch.dict(sys.modules, {
            "win32com": fake_win32,
            "win32com.client": fake_win32.client,
        }):
            ok, mensaje = sap_upload.diagnosticar_conexion_sap()
        self.assertFalse(ok)
        self.assertIn("Scripting Engine", mensaje)

    def test_returns_false_when_no_connections(self):
        engine = MagicMock()
        engine.Children.Count = 0
        fake_win32 = self._make_fake_win32(engine)
        with patch.dict(sys.modules, {
            "win32com": fake_win32,
            "win32com.client": fake_win32.client,
        }):
            ok, mensaje = sap_upload.diagnosticar_conexion_sap()
        self.assertFalse(ok)
        self.assertIn("conexión", mensaje.lower())
        self.assertIn("Logon Pad", mensaje)

    def test_returns_false_when_no_sessions(self):
        connection = MagicMock()
        connection.Children.Count = 0
        engine = MagicMock()
        engine.Children.Count = 1
        engine.Children.return_value = connection
        fake_win32 = self._make_fake_win32(engine)
        with patch.dict(sys.modules, {
            "win32com": fake_win32,
            "win32com.client": fake_win32.client,
        }):
            ok, mensaje = sap_upload.diagnosticar_conexion_sap()
        self.assertFalse(ok)
        self.assertIn("NO hay sesiones", mensaje)
        self.assertIn("Inicia sesión", mensaje)

    def test_returns_true_with_session_info(self):
        info = MagicMock()
        info.SystemName = "PRD"
        info.Client = "100"
        info.User = "SROCK"
        session = MagicMock()
        type(session).Info = property(lambda self: info)
        connection = MagicMock()
        connection.Children.Count = 1
        connection.Children.return_value = session
        engine = MagicMock()
        engine.Children.Count = 1
        engine.Children.return_value = connection
        fake_win32 = self._make_fake_win32(engine)
        with patch.dict(sys.modules, {
            "win32com": fake_win32,
            "win32com.client": fake_win32.client,
        }):
            ok, mensaje = sap_upload.diagnosticar_conexion_sap()
        self.assertTrue(ok)
        self.assertIn("OK", mensaje)
        self.assertIn("sistema=PRD", mensaje)
        self.assertIn("client=100", mensaje)
        self.assertIn("user=SROCK", mensaje)

    def test_returns_true_when_session_info_unavailable(self):
        """Si session.Info lanza excepción, debe seguir reportando OK
        siempre que haya al menos una sesión visible."""
        session = MagicMock()
        type(session).Info = property(
            lambda self: (_ for _ in ()).throw(Exception("denied"))
        )
        connection = MagicMock()
        connection.Children.Count = 1
        connection.Children.return_value = session
        engine = MagicMock()
        engine.Children.Count = 1
        engine.Children.return_value = connection
        fake_win32 = self._make_fake_win32(engine)
        with patch.dict(sys.modules, {
            "win32com": fake_win32,
            "win32com.client": fake_win32.client,
        }):
            ok, mensaje = sap_upload.diagnosticar_conexion_sap()
        self.assertTrue(ok)
        self.assertIn("info no disponible", mensaje)


class OpenLsmwTest(unittest.TestCase):
    def test_maximizes_window_sets_okcd_and_executes(self):
        session = MockSAPSession()
        open_lsmw(session)

        self.assertIn(("wnd[0]", "maximize"), session.actions)
        self.assertEqual(session._elements["wnd[0]/tbar[0]/okcd"].text, "LSMW")
        self.assertIn(("wnd[0]", "sendVKey", 0), session.actions)
        self.assertIn(("wnd[0]/tbar[1]/btn[8]", "press"), session.actions)

    def test_actions_in_correct_order(self):
        session = MockSAPSession()
        open_lsmw(session)

        order = [a for a in session.actions if a[0] in ("wnd[0]", "wnd[0]/tbar[1]/btn[8]")]
        # maximize → sendVKey 0 → press
        self.assertEqual(order[0], ("wnd[0]", "maximize"))
        self.assertEqual(order[1], ("wnd[0]", "sendVKey", 0))
        self.assertEqual(order[2], ("wnd[0]/tbar[1]/btn[8]", "press"))


class SelectStepRowTest(unittest.TestCase):
    def test_deselects_default_row_and_selects_target(self):
        session = MockSAPSession()
        select_step_row(session, 6)

        steplist = session._elements[LSMW_STEPLIST_TABLE]
        self.assertFalse(steplist._rows[DEFAULT_SELECTED_ROW]._selected)
        self.assertTrue(steplist._rows[6]._selected)

    def test_focuses_step_text_cell(self):
        session = MockSAPSession()
        select_step_row(session, 6)

        cell_id = f"{LSMW_STEPLIST_TABLE}/txtGT_STEPLIST-STEPTEXT[0,6]"
        self.assertIn((cell_id, "setFocus"), session.actions)
        self.assertEqual(session._elements[cell_id].caretPosition, 0)

    def test_works_for_different_rows(self):
        session = MockSAPSession()
        select_step_row(session, ASSIGN_FILES_ROW)
        select_step_row(session, READ_DATA_ROW)

        steplist = session._elements[LSMW_STEPLIST_TABLE]
        self.assertTrue(steplist._rows[ASSIGN_FILES_ROW]._selected)
        self.assertTrue(steplist._rows[READ_DATA_ROW]._selected)


class ConfigurarRutaArchivoTest(unittest.TestCase):
    """Verifica la replicación 1:1 de resources/Script1.vbs."""

    CARPETA = r"C:\Users\test\salida"
    NOMBRE = "LSMW_20260510_094838.txt"

    def _ejecutar(self):
        session = MockSAPSession()
        configurar_ruta_archivo(session, self.CARPETA, self.NOMBRE)
        return session

    def test_opens_specify_files_step_with_f2(self):
        session = self._ejecutar()
        cell_id = f"{LSMW_STEPLIST_TABLE}/txtGT_STEPLIST-STEPTEXT[0,{SPECIFY_FILES_ROW}]"
        self.assertIn((cell_id, "setFocus"), session.actions)
        self.assertEqual(session._elements[cell_id].caretPosition, 5)
        self.assertIn(("wnd[0]", "sendVKey", 2), session.actions)

    def test_presses_change_mode_and_assign_buttons(self):
        session = self._ejecutar()
        # btn[25] = modo edición, btn[27] = "Asignar archivo"
        self.assertIn(("wnd[0]/tbar[1]/btn[25]", "press"), session.actions)
        self.assertIn(("wnd[0]/tbar[1]/btn[27]", "press"), session.actions)

    def test_focuses_file_definition_label(self):
        session = self._ejecutar()
        self.assertIn(("wnd[0]/usr/lbl[43,6]", "setFocus"), session.actions)
        self.assertEqual(session._elements["wnd[0]/usr/lbl[43,6]"].caretPosition, 3)

    def test_sends_f4_to_open_file_browser(self):
        session = self._ejecutar()
        self.assertIn(("wnd[1]", "sendVKey", 4), session.actions)

    def test_sets_path_and_filename_in_browser(self):
        session = self._ejecutar()
        self.assertEqual(session._elements["wnd[2]/usr/ctxtDY_PATH"].text, self.CARPETA)
        filename_field = session._elements["wnd[2]/usr/ctxtDY_FILENAME"]
        self.assertEqual(filename_field.text, self.NOMBRE)
        self.assertEqual(filename_field.caretPosition, len(self.NOMBRE))

    def test_confirms_dialogs_back_and_saves(self):
        session = self._ejecutar()
        self.assertIn(("wnd[2]/tbar[0]/btn[0]", "press"), session.actions)
        self.assertIn(("wnd[1]/tbar[0]/btn[0]", "press"), session.actions)
        self.assertIn(("wnd[0]/tbar[0]/btn[3]", "press"), session.actions)
        self.assertIn(("wnd[1]/usr/btnSPOP-OPTION1", "press"), session.actions)

    def test_actions_in_correct_sequence(self):
        session = self._ejecutar()
        # Verificar el orden parcial: F2 → btn[25] → btn[27] → F4 → OK explorador
        order = [
            a for a in session.actions
            if a in [
                ("wnd[0]", "sendVKey", 2),
                ("wnd[0]/tbar[1]/btn[25]", "press"),
                ("wnd[0]/tbar[1]/btn[27]", "press"),
                ("wnd[1]", "sendVKey", 4),
                ("wnd[2]/tbar[0]/btn[0]", "press"),
                ("wnd[1]/usr/btnSPOP-OPTION1", "press"),
            ]
        ]
        self.assertEqual(order, [
            ("wnd[0]", "sendVKey", 2),
            ("wnd[0]/tbar[1]/btn[25]", "press"),
            ("wnd[0]/tbar[1]/btn[27]", "press"),
            ("wnd[1]", "sendVKey", 4),
            ("wnd[2]/tbar[0]/btn[0]", "press"),
            ("wnd[1]/usr/btnSPOP-OPTION1", "press"),
        ])


class StepAssignFilesTest(unittest.TestCase):
    def test_selects_row_executes_and_goes_back(self):
        session = MockSAPSession()
        step_assign_files(session)

        steplist = session._elements[LSMW_STEPLIST_TABLE]
        self.assertTrue(steplist._rows[ASSIGN_FILES_ROW]._selected)
        self.assertIn(("wnd[0]/tbar[1]/btn[32]", "press"), session.actions)
        self.assertIn(("wnd[0]", "sendVKey", 3), session.actions)


class StepReadDataTest(unittest.TestCase):
    def test_executes_read_then_returns_twice(self):
        session = MockSAPSession()
        step_read_data(session)

        steplist = session._elements[LSMW_STEPLIST_TABLE]
        self.assertTrue(steplist._rows[READ_DATA_ROW]._selected)
        self.assertIn(("wnd[0]/tbar[1]/btn[32]", "press"), session.actions)
        self.assertIn(("wnd[0]/tbar[1]/btn[8]", "press"), session.actions)

        backs = [a for a in session.actions if a == ("wnd[0]", "sendVKey", 3)]
        self.assertEqual(len(backs), 2)


class StepDisplayReadDataTest(unittest.TestCase):
    def test_executes_confirms_popup_and_returns(self):
        session = MockSAPSession()
        step_display_read_data(session)

        self.assertIn(("wnd[0]/tbar[1]/btn[32]", "press"), session.actions)
        self.assertIn(("wnd[1]", "sendVKey", 0), session.actions)
        self.assertIn(("wnd[0]", "sendVKey", 3), session.actions)


class StepConvertDataTest(unittest.TestCase):
    def test_executes_with_f8_and_returns_twice(self):
        session = MockSAPSession()
        step_convert_data(session)

        self.assertIn(("wnd[0]/tbar[1]/btn[32]", "press"), session.actions)
        self.assertIn(("wnd[0]", "sendVKey", 8), session.actions)
        backs = [a for a in session.actions if a == ("wnd[0]", "sendVKey", 3)]
        self.assertEqual(len(backs), 2)


class StepDisplayConvertedDataTest(unittest.TestCase):
    def test_executes_confirms_popup_and_returns(self):
        session = MockSAPSession()
        step_display_converted_data(session)

        self.assertIn(("wnd[0]/tbar[1]/btn[32]", "press"), session.actions)
        self.assertIn(("wnd[1]", "sendVKey", 0), session.actions)
        self.assertIn(("wnd[0]", "sendVKey", 3), session.actions)


class StepCreateBatchInputTest(unittest.TestCase):
    def test_marks_p_keep_creates_session_and_confirms(self):
        session = MockSAPSession()
        step_create_batch_input(session)

        self.assertIn(("wnd[0]/tbar[1]/btn[32]", "press"), session.actions)
        chk = session._elements["wnd[0]/usr/chkP_KEEP"]
        self.assertTrue(chk.selected)
        self.assertIn(("wnd[0]/usr/chkP_KEEP", "setFocus"), session.actions)
        self.assertIn(("wnd[0]/tbar[1]/btn[8]", "press"), session.actions)
        self.assertIn(("wnd[1]", "sendVKey", 0), session.actions)


class VolverAlStepListTest(unittest.TestCase):
    """_volver_al_step_list garantiza que la sesión esté en el step list,
    enviando Back hasta llegar (con límite de intentos)."""

    def test_no_op_when_already_on_step_list(self):
        """Si la tabla del step list está accesible, no se envía Back."""
        from sap_upload import _volver_al_step_list, LSMW_STEPLIST_TABLE

        session = MockSAPSession()
        # Asegura que LSMW_STEPLIST_TABLE existe
        session.findById(LSMW_STEPLIST_TABLE)
        _volver_al_step_list(session)

        # No se debe haber enviado ningún sendVKey 3
        backs = [a for a in session.actions if a == ("wnd[0]", "sendVKey", 3)]
        self.assertEqual(len(backs), 0)

    def test_sends_back_when_not_on_step_list(self):
        """Si LSMW_STEPLIST_TABLE no existe inicialmente, envía Back hasta
        encontrarla."""
        from sap_upload import _volver_al_step_list, LSMW_STEPLIST_TABLE

        session = MockSAPSession()
        original_find = session.findById
        attempts = [0]

        def find_failing_first_attempt(sap_id):
            if sap_id == LSMW_STEPLIST_TABLE and attempts[0] == 0:
                attempts[0] += 1
                raise Exception("no estamos en step list todavía")
            return original_find(sap_id)

        session.findById = find_failing_first_attempt
        _volver_al_step_list(session)

        backs = [a for a in session.actions if a == ("wnd[0]", "sendVKey", 3)]
        self.assertGreaterEqual(len(backs), 1)

    def test_raises_when_step_list_never_appears(self):
        """Si tras max_intentos seguimos sin step list, lanza RuntimeError
        con mensaje accionable."""
        from sap_upload import _volver_al_step_list, LSMW_STEPLIST_TABLE

        session = MockSAPSession()

        def always_fails(sap_id):
            if sap_id == LSMW_STEPLIST_TABLE:
                raise Exception("nunca llegamos")
            # Para que sendVKey funcione, devolver el wnd normalmente
            return MockSAPSession.findById(session, sap_id)

        # Necesitamos un wnd[0] real para sendVKey, pero hacer fallar
        # LSMW_STEPLIST_TABLE. Usamos un closure.
        original_find = MockSAPSession.findById

        def fake_find(self_sess, sap_id):
            if sap_id == LSMW_STEPLIST_TABLE:
                raise Exception("nunca llegamos")
            return original_find(self_sess, sap_id)

        session.findById = lambda sid: fake_find(session, sid)

        with self.assertRaisesRegex(RuntimeError, "step list de LSMW"):
            _volver_al_step_list(session, max_intentos=3)


class StepRunBatchInputTest(unittest.TestCase):
    def test_selects_run_bi_row_and_presses_execute(self):
        session = MockSAPSession()
        step_run_batch_input(session)

        # Debe seleccionar explícitamente la fila 13 (Run BI) y luego
        # presionar Execute. El select_step_row interno emite varias
        # acciones; verificamos las clave.
        from sap_upload import RUN_BI_ROW, LSMW_STEPLIST_TABLE
        steplist = session._elements[LSMW_STEPLIST_TABLE]
        self.assertTrue(steplist._rows[RUN_BI_ROW]._selected)
        self.assertIn(("wnd[0]/tbar[1]/btn[32]", "press"), session.actions)


class ProcessBdcSessionTest(unittest.TestCase):
    def test_selects_first_row_processes_with_error_mode_log_and_expert(self):
        session = MockSAPSession()
        process_bdc_session(session)

        bdc_table = session._elements[BDC_SESSION_TABLE]
        self.assertTrue(bdc_table._rows[0]._selected)

        group_id = f"{BDC_SESSION_TABLE}/txtITAB_APQI-GROUPID[0,0]"
        self.assertIn((group_id, "setFocus"), session.actions)
        self.assertEqual(session._elements[group_id].caretPosition, 0)

        self.assertIn(("wnd[0]/tbar[1]/btn[8]", "press"), session.actions)
        self.assertIn(("wnd[1]/usr/radD0300-ERROR", "select"), session.actions)
        self.assertTrue(session._elements["wnd[1]/usr/chkD0300-LOGALL"].selected)
        self.assertTrue(session._elements["wnd[1]/usr/chkD0300-EXPERT"].selected)
        self.assertIn(("wnd[1]/usr/chkD0300-EXPERT", "setFocus"), session.actions)

        ok_presses = [a for a in session.actions if a == ("wnd[1]/tbar[0]/btn[0]", "press")]
        self.assertEqual(len(ok_presses), 2)


# ---------------------------------------------------------------------------
# run_lsmw_flow
# ---------------------------------------------------------------------------


class RunLsmwFlowTest(unittest.TestCase):
    CARPETA = r"C:\test\salida"
    NOMBRE = "LSMW_test.txt"

    def test_calls_all_steps_in_order(self):
        session = MockSAPSession()
        call_order = []

        def make_recorder(name, accepts_path=False):
            if accepts_path:
                return lambda s, c, n: call_order.append((name, c, n))
            return lambda s: call_order.append(name)

        with patch.multiple(
            "sap_upload",
            open_lsmw=make_recorder("open_lsmw"),
            configurar_ruta_archivo=make_recorder("configurar_ruta", accepts_path=True),
            step_assign_files=make_recorder("assign_files"),
            step_read_data=make_recorder("read_data"),
            step_display_read_data=make_recorder("display_read"),
            step_convert_data=make_recorder("convert"),
            step_display_converted_data=make_recorder("display_converted"),
            step_create_batch_input=make_recorder("create_bi"),
            step_run_batch_input=make_recorder("run_bi"),
            process_bdc_session=make_recorder("process_bdc"),
        ):
            run_lsmw_flow(session, self.CARPETA, self.NOMBRE)

        self.assertEqual(
            call_order,
            [
                "open_lsmw",
                ("configurar_ruta", self.CARPETA, self.NOMBRE),
                "assign_files",
                "read_data",
                "display_read",
                "convert",
                "display_converted",
                "create_bi",
                "run_bi",
                "process_bdc",
            ],
        )

    def test_passes_session_to_each_step(self):
        session = MockSAPSession()
        single_arg_spies = {
            name: MagicMock()
            for name in [
                "open_lsmw",
                "step_assign_files",
                "step_read_data",
                "step_display_read_data",
                "step_convert_data",
                "step_display_converted_data",
                "step_create_batch_input",
                "step_run_batch_input",
                "process_bdc_session",
            ]
        }
        configurar_spy = MagicMock()
        with patch.multiple(
            "sap_upload",
            configurar_ruta_archivo=configurar_spy,
            **single_arg_spies,
        ):
            run_lsmw_flow(session, self.CARPETA, self.NOMBRE)

        for spy in single_arg_spies.values():
            spy.assert_called_once_with(session)
        configurar_spy.assert_called_once_with(session, self.CARPETA, self.NOMBRE)


# ---------------------------------------------------------------------------
# main()
# ---------------------------------------------------------------------------


class MainEntryPointTest(unittest.TestCase):
    def test_returns_1_when_no_txt_in_salida(self):
        with patch("sap_upload.get_latest_txt", side_effect=FileNotFoundError("no hay")):
            self.assertEqual(sap_upload.main(), 1)

    def test_returns_1_when_sap_session_fails(self):
        fake_path = Path("/tmp/fake.txt")
        with patch("sap_upload.get_latest_txt", return_value=fake_path), \
             patch("sap_upload.get_sap_session", side_effect=RuntimeError("no SAP")):
            self.assertEqual(sap_upload.main(), 1)

    def test_returns_0_on_happy_path(self):
        fake_path = Path("/tmp/LSMW_x.txt")
        session = MagicMock()
        with patch("sap_upload.get_latest_txt", return_value=fake_path), \
             patch("sap_upload.get_sap_session", return_value=session), \
             patch("sap_upload.run_lsmw_flow") as mock_flow:
            self.assertEqual(sap_upload.main(), 0)
            mock_flow.assert_called_once_with(session, str(fake_path.parent), fake_path.name)

    def test_passes_folder_and_filename_to_flow(self):
        fake_path = Path("/some/folder/LSMW_20260510_094838.txt")
        with patch("sap_upload.get_latest_txt", return_value=fake_path), \
             patch("sap_upload.get_sap_session", return_value=MagicMock()), \
             patch("sap_upload.run_lsmw_flow") as mock_flow:
            sap_upload.main()
            mock_flow.assert_called_once()
            args = mock_flow.call_args[0]
            self.assertEqual(args[1], "/some/folder")
            self.assertEqual(args[2], "LSMW_20260510_094838.txt")

    def test_returns_1_when_lsmw_flow_raises(self):
        fake_path = Path("/tmp/fake.txt")
        with patch("sap_upload.get_latest_txt", return_value=fake_path), \
             patch("sap_upload.get_sap_session", return_value=MagicMock()), \
             patch("sap_upload.run_lsmw_flow", side_effect=Exception("falla")):
            self.assertEqual(sap_upload.main(), 1)


if __name__ == "__main__":
    unittest.main()
