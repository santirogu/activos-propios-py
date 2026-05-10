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

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import sap_upload  # noqa: E402
from sap_upload import (  # noqa: E402
    ASSIGN_FILES_ROW,
    BDC_SESSION_TABLE,
    DEFAULT_SELECTED_ROW,
    LSMW_STEPLIST_TABLE,
    READ_DATA_ROW,
    SPECIFY_FILES_ROW,
    copy_to_sap_path,
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
    step_specify_files,
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
# copy_to_sap_path
# ---------------------------------------------------------------------------


class CopyToSapPathTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def test_copies_content_to_destination(self):
        src = self.tmp / "src.txt"
        src.write_text("contenido prueba", encoding="utf-8")
        dst = self.tmp / "dst.txt"

        result = copy_to_sap_path(src, str(dst))

        self.assertEqual(result, dst)
        self.assertEqual(dst.read_text(encoding="utf-8"), "contenido prueba")

    def test_creates_parent_directory_if_missing(self):
        src = self.tmp / "src.txt"
        src.write_text("x", encoding="utf-8")
        dst = self.tmp / "a" / "b" / "c" / "dst.txt"

        copy_to_sap_path(src, str(dst))

        self.assertTrue(dst.parent.is_dir())
        self.assertTrue(dst.exists())


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


class StepSpecifyFilesTest(unittest.TestCase):
    def test_selects_row_executes_and_returns(self):
        session = MockSAPSession()
        step_specify_files(session)

        steplist = session._elements[LSMW_STEPLIST_TABLE]
        self.assertTrue(steplist._rows[SPECIFY_FILES_ROW]._selected)
        self.assertIn(("wnd[0]/tbar[1]/btn[32]", "press"), session.actions)
        self.assertIn(("wnd[0]/tbar[0]/btn[3]", "press"), session.actions)


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


class StepRunBatchInputTest(unittest.TestCase):
    def test_presses_execute(self):
        session = MockSAPSession()
        step_run_batch_input(session)

        self.assertEqual(
            session.actions,
            [("wnd[0]/tbar[1]/btn[32]", "press")],
        )


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
    def test_calls_all_steps_in_order(self):
        session = MockSAPSession()
        call_order = []

        def make_recorder(name):
            return lambda s: call_order.append(name)

        with patch.multiple(
            "sap_upload",
            open_lsmw=make_recorder("open_lsmw"),
            step_specify_files=make_recorder("specify_files"),
            step_assign_files=make_recorder("assign_files"),
            step_read_data=make_recorder("read_data"),
            step_display_read_data=make_recorder("display_read"),
            step_convert_data=make_recorder("convert"),
            step_display_converted_data=make_recorder("display_converted"),
            step_create_batch_input=make_recorder("create_bi"),
            step_run_batch_input=make_recorder("run_bi"),
            process_bdc_session=make_recorder("process_bdc"),
        ):
            run_lsmw_flow(session)

        self.assertEqual(
            call_order,
            [
                "open_lsmw",
                "specify_files",
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
        spies = {
            name: MagicMock()
            for name in [
                "open_lsmw",
                "step_specify_files",
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
        with patch.multiple("sap_upload", **spies):
            run_lsmw_flow(session)

        for name, spy in spies.items():
            spy.assert_called_once_with(session)


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
             patch("sap_upload.SAP_LSMW_INPUT_PATH", None), \
             patch("sap_upload.get_sap_session", side_effect=RuntimeError("no SAP")):
            self.assertEqual(sap_upload.main(), 1)

    def test_returns_0_on_happy_path(self):
        fake_path = Path("/tmp/fake.txt")
        session = MagicMock()
        with patch("sap_upload.get_latest_txt", return_value=fake_path), \
             patch("sap_upload.SAP_LSMW_INPUT_PATH", None), \
             patch("sap_upload.get_sap_session", return_value=session), \
             patch("sap_upload.run_lsmw_flow") as mock_flow:
            self.assertEqual(sap_upload.main(), 0)
            mock_flow.assert_called_once_with(session)

    def test_copies_to_sap_path_when_configured(self):
        fake_path = Path("/tmp/fake.txt")
        session = MagicMock()
        with patch("sap_upload.get_latest_txt", return_value=fake_path), \
             patch("sap_upload.SAP_LSMW_INPUT_PATH", r"C:\sap\input.txt"), \
             patch("sap_upload.copy_to_sap_path") as mock_copy, \
             patch("sap_upload.get_sap_session", return_value=session), \
             patch("sap_upload.run_lsmw_flow"):
            self.assertEqual(sap_upload.main(), 0)
            mock_copy.assert_called_once_with(fake_path, r"C:\sap\input.txt")

    def test_returns_1_when_lsmw_flow_raises(self):
        fake_path = Path("/tmp/fake.txt")
        with patch("sap_upload.get_latest_txt", return_value=fake_path), \
             patch("sap_upload.SAP_LSMW_INPUT_PATH", None), \
             patch("sap_upload.get_sap_session", return_value=MagicMock()), \
             patch("sap_upload.run_lsmw_flow", side_effect=Exception("falla")):
            self.assertEqual(sap_upload.main(), 1)


if __name__ == "__main__":
    unittest.main()
