import sys
import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import main  # noqa: E402
from main import export_sheet_to_tsv, subir_a_sap  # noqa: E402


class ExportSheetToTsvTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.excel_path = self.tmp / "test.xlsx"
        self.output_dir = self.tmp / "out"

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _make_workbook(self, sheet_name: str, rows: list[list]) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name
        for r in rows:
            ws.append(r)
        wb.save(self.excel_path)

    def test_writes_tab_separated_content(self) -> None:
        self._make_workbook("LSMW ", [["A", "B", "C"], [1, 2, 3]])

        out_path, rows = export_sheet_to_tsv(self.excel_path, "LSMW ", self.output_dir)

        self.assertEqual(rows, 2)
        self.assertEqual(out_path.read_text(encoding="utf-8"), "A\tB\tC\n1\t2\t3\n")

    def test_none_values_become_empty_strings(self) -> None:
        self._make_workbook("S", [["x", None, "y"], [None, "z", None]])

        out_path, _ = export_sheet_to_tsv(self.excel_path, "S", self.output_dir)

        self.assertEqual(out_path.read_text(encoding="utf-8"), "x\t\ty\n\tz\t\n")

    def test_creates_output_directory_if_missing(self) -> None:
        self._make_workbook("S", [["a"]])
        nested = self.output_dir / "nivel1" / "nivel2"
        self.assertFalse(nested.exists())

        export_sheet_to_tsv(self.excel_path, "S", nested)

        self.assertTrue(nested.is_dir())

    def test_filename_has_timestamp_pattern(self) -> None:
        self._make_workbook("S", [["a"]])

        out_path, _ = export_sheet_to_tsv(self.excel_path, "S", self.output_dir)

        self.assertRegex(out_path.name, r"^LSMW_\d{8}_\d{6}\.txt$")

    def test_custom_file_prefix(self) -> None:
        self._make_workbook("S", [["a"]])

        out_path, _ = export_sheet_to_tsv(
            self.excel_path, "S", self.output_dir, file_prefix="EXPORT"
        )

        self.assertTrue(out_path.name.startswith("EXPORT_"))

    def test_missing_excel_raises_file_not_found(self) -> None:
        with self.assertRaises(FileNotFoundError):
            export_sheet_to_tsv(self.tmp / "no_existe.xlsx", "S", self.output_dir)

    def test_missing_sheet_raises_value_error(self) -> None:
        self._make_workbook("Existente", [["a"]])

        with self.assertRaisesRegex(ValueError, "NoExiste"):
            export_sheet_to_tsv(self.excel_path, "NoExiste", self.output_dir)

    def test_returns_row_count_matching_written_lines(self) -> None:
        self._make_workbook("S", [["a", "b"], ["c", "d"], ["e", "f"]])

        out_path, rows = export_sheet_to_tsv(self.excel_path, "S", self.output_dir)

        self.assertEqual(rows, 3)
        self.assertEqual(len(out_path.read_text(encoding="utf-8").splitlines()), 3)

    def test_does_not_overwrite_when_called_in_different_seconds(self) -> None:
        self._make_workbook("S", [["a"]])

        with patch("main.datetime") as mock_dt:
            mock_dt.now.return_value.strftime.return_value = "20260101_120000"
            first, _ = export_sheet_to_tsv(self.excel_path, "S", self.output_dir)
            mock_dt.now.return_value.strftime.return_value = "20260101_120001"
            second, _ = export_sheet_to_tsv(self.excel_path, "S", self.output_dir)

        self.assertNotEqual(first, second)
        self.assertTrue(first.exists())
        self.assertTrue(second.exists())


class RealWorkbookSmokeTest(unittest.TestCase):
    """Smoke test contra el Excel real del proyecto si está disponible."""

    REAL_EXCEL = (
        Path(__file__).resolve().parent.parent / "resources" / "Formato_Dinamico_.xlsx"
    )

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.output_dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_extracts_lsmw_sheet_from_real_file(self) -> None:
        if not self.REAL_EXCEL.exists():
            self.skipTest("Archivo Formato_Dinamico_.xlsx no disponible")

        out_path, rows = export_sheet_to_tsv(self.REAL_EXCEL, "LSMW ", self.output_dir)

        self.assertGreaterEqual(rows, 2)
        first_line = out_path.read_text(encoding="utf-8").splitlines()[0]
        self.assertIn("ANLKL", first_line)
        self.assertIn("BUKRS", first_line)
        self.assertEqual(first_line.count("\t"), 50)


class _SyncFakeThread:
    """Reemplaza threading.Thread para ejecutar el target de forma síncrona."""

    def __init__(self, target=None, daemon=None, **kwargs):
        self.target = target
        self.daemon = daemon

    def start(self):
        if self.target is not None:
            self.target()


class SubirASapTest(unittest.TestCase):
    """Pruebas para el handler del botón "Subir a SAP" en main.py."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        # root.after debe disparar el callback inmediatamente para que las
        # actualizaciones de UI del worker corran de forma síncrona.
        self.root.after = lambda delay, fn, *args: fn(*args)
        self.status_var = tk.StringVar(master=self.root)
        self.button = tk.Button(self.root)

    def tearDown(self) -> None:
        self.root.destroy()

    def _patch_sap_upload(self, **overrides):
        """Patches por defecto del módulo sap_upload con overrides opcionales.

        Guarda los mocks en `self.mocks` (patch.multiple no devuelve mocks
        pasados como valores explícitos — solo los marcados como DEFAULT).
        """
        self.mocks = {
            "get_latest_txt": MagicMock(return_value=Path("/tmp/LSMW_x.txt")),
            "copy_to_sap_path": MagicMock(),
            "get_sap_session": MagicMock(return_value=MagicMock(name="session")),
            "run_lsmw_flow": MagicMock(),
            "SAP_LSMW_INPUT_PATH": None,
        }
        self.mocks.update(overrides)
        return patch.multiple("sap_upload", **self.mocks)

    # ------------------------------------------------------------------ cancel

    def test_cancel_confirmation_does_not_start_thread(self) -> None:
        with patch("main.messagebox.askyesno", return_value=False), \
             patch("main.threading.Thread") as mock_thread:
            subir_a_sap(self.root, self.status_var, self.button)

        mock_thread.assert_not_called()
        self.assertEqual(str(self.button["state"]), "normal")

    def test_cancel_does_not_modify_status(self) -> None:
        self.status_var.set("estado previo")
        with patch("main.messagebox.askyesno", return_value=False):
            subir_a_sap(self.root, self.status_var, self.button)

        self.assertEqual(self.status_var.get(), "estado previo")

    # ----------------------------------------------------------------- confirm

    def test_confirmation_disables_button_before_starting_worker(self) -> None:
        captured_state = {}

        def capture(target=None, **kwargs):
            captured_state["before_start"] = str(self.button["state"])
            return _SyncFakeThread(target=target, **kwargs)

        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", side_effect=capture), \
             self._patch_sap_upload():
            subir_a_sap(self.root, self.status_var, self.button)

        self.assertEqual(captured_state["before_start"], "disabled")

    # --------------------------------------------------------------- happy path

    def test_worker_calls_full_flow_on_happy_path(self) -> None:
        session = MagicMock(name="session")
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo") as mock_info, \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload(
                 get_sap_session=MagicMock(return_value=session),
                 run_lsmw_flow=MagicMock(),
             ):
            subir_a_sap(self.root, self.status_var, self.button)

        self.mocks["get_latest_txt"].assert_called_once()
        self.mocks["get_sap_session"].assert_called_once()
        self.mocks["run_lsmw_flow"].assert_called_once_with(session)
        mock_info.assert_called_once()

    def test_worker_reenables_button_after_success(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload():
            subir_a_sap(self.root, self.status_var, self.button)

        self.assertEqual(str(self.button["state"]), "normal")

    def test_worker_updates_status_to_completion_message(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload():
            subir_a_sap(self.root, self.status_var, self.button)

        self.assertIn("completada", self.status_var.get().lower())

    def test_worker_skips_copy_when_sap_path_not_configured(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload(SAP_LSMW_INPUT_PATH=None):
            subir_a_sap(self.root, self.status_var, self.button)

        self.mocks["copy_to_sap_path"].assert_not_called()

    def test_worker_copies_to_sap_path_when_configured(self) -> None:
        fake_path = Path("/tmp/LSMW_x.txt")
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload(
                 get_latest_txt=MagicMock(return_value=fake_path),
                 SAP_LSMW_INPUT_PATH=r"C:\sap\input.txt",
             ):
            subir_a_sap(self.root, self.status_var, self.button)

        self.mocks["copy_to_sap_path"].assert_called_once_with(fake_path, r"C:\sap\input.txt")

    # ----------------------------------------------------------------- errores

    def test_worker_handles_missing_txt(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showerror") as mock_err, \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload(
                 get_latest_txt=MagicMock(side_effect=FileNotFoundError("no hay txt")),
             ):
            subir_a_sap(self.root, self.status_var, self.button)

        mock_err.assert_called_once()
        title, message = mock_err.call_args[0][:2]
        self.assertIn("no hay txt", message)
        self.assertEqual(str(self.button["state"]), "normal")

    def test_worker_handles_sap_connection_error(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showerror") as mock_err, \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload(
                 get_sap_session=MagicMock(side_effect=RuntimeError("no SAP")),
             ):
            subir_a_sap(self.root, self.status_var, self.button)

        mock_err.assert_called_once()
        self.assertIn("no SAP", mock_err.call_args[0][1])
        self.assertEqual(str(self.button["state"]), "normal")

    def test_worker_handles_lsmw_flow_error(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showerror") as mock_err, \
             patch("main.messagebox.showinfo") as mock_info, \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload(
                 run_lsmw_flow=MagicMock(side_effect=Exception("paso 5 falló")),
             ):
            subir_a_sap(self.root, self.status_var, self.button)

        mock_err.assert_called_once()
        mock_info.assert_not_called()
        self.assertIn("paso 5 falló", mock_err.call_args[0][1])
        self.assertEqual(str(self.button["state"]), "normal")

    def test_worker_resets_status_on_error(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showerror"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload(
                 get_sap_session=MagicMock(side_effect=RuntimeError("error")),
             ):
            subir_a_sap(self.root, self.status_var, self.button)

        self.assertEqual(self.status_var.get(), "")


if __name__ == "__main__":
    unittest.main()
