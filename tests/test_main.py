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


class ExtraerLsmwATxtTest(unittest.TestCase):
    """Pruebas para el handler del botón "Extraer información en txt" — en
    particular la lógica de confirmar antes de reemplazar un .txt existente.
    """

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        self.status_var = tk.StringVar(master=self.root)
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_salida = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()
        self.root.destroy()

    def _patch_output_dir(self):
        return patch("main.OUTPUT_DIR", self.tmp_salida)

    def test_proceeds_directly_when_no_existing_txt(self) -> None:
        with self._patch_output_dir(), \
             patch("main.export_sheet_to_tsv", return_value=(self.tmp_salida / "new.txt", 2)) as mock_export, \
             patch("main.messagebox.askyesno") as mock_ask, \
             patch("main.messagebox.showinfo"):
            main.extraer_lsmw_a_txt(self.status_var)

        mock_ask.assert_not_called()
        mock_export.assert_called_once()

    def test_asks_for_replacement_when_txt_exists(self) -> None:
        (self.tmp_salida / "LSMW_20260101_120000.txt").write_text("x", encoding="utf-8")

        with self._patch_output_dir(), \
             patch("main.export_sheet_to_tsv", return_value=(self.tmp_salida / "new.txt", 2)), \
             patch("main.messagebox.askyesno", return_value=True) as mock_ask, \
             patch("main.messagebox.showinfo"):
            main.extraer_lsmw_a_txt(self.status_var)

        mock_ask.assert_called_once()
        # El mensaje debe mencionar el archivo existente
        args = mock_ask.call_args[0]
        self.assertIn("LSMW_20260101_120000.txt", args[1])

    def test_yes_deletes_existing_and_creates_new(self) -> None:
        old_file = self.tmp_salida / "LSMW_20260101_120000.txt"
        old_file.write_text("contenido viejo", encoding="utf-8")

        with self._patch_output_dir(), \
             patch("main.export_sheet_to_tsv", return_value=(self.tmp_salida / "new.txt", 2)) as mock_export, \
             patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"):
            main.extraer_lsmw_a_txt(self.status_var)

        self.assertFalse(old_file.exists())
        mock_export.assert_called_once()

    def test_yes_deletes_all_existing_txt_files(self) -> None:
        files = [
            self.tmp_salida / "LSMW_20260101_120000.txt",
            self.tmp_salida / "LSMW_20260102_120000.txt",
            self.tmp_salida / "LSMW_20260103_120000.txt",
        ]
        for f in files:
            f.write_text("x", encoding="utf-8")

        with self._patch_output_dir(), \
             patch("main.export_sheet_to_tsv", return_value=(self.tmp_salida / "new.txt", 2)), \
             patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"):
            main.extraer_lsmw_a_txt(self.status_var)

        for f in files:
            self.assertFalse(f.exists(), f"{f.name} debió ser borrado")

    def test_no_keeps_existing_and_does_not_extract(self) -> None:
        old_file = self.tmp_salida / "LSMW_20260101_120000.txt"
        old_file.write_text("contenido viejo", encoding="utf-8")

        with self._patch_output_dir(), \
             patch("main.export_sheet_to_tsv") as mock_export, \
             patch("main.messagebox.askyesno", return_value=False), \
             patch("main.messagebox.showinfo") as mock_info:
            main.extraer_lsmw_a_txt(self.status_var)

        self.assertTrue(old_file.exists())
        self.assertEqual(old_file.read_text(encoding="utf-8"), "contenido viejo")
        mock_export.assert_not_called()
        mock_info.assert_not_called()

    def test_no_updates_status_with_cancellation_message(self) -> None:
        (self.tmp_salida / "LSMW_20260101_120000.txt").write_text("x", encoding="utf-8")

        with self._patch_output_dir(), \
             patch("main.export_sheet_to_tsv"), \
             patch("main.messagebox.askyesno", return_value=False), \
             patch("main.messagebox.showinfo"):
            main.extraer_lsmw_a_txt(self.status_var)

        self.assertIn("cancelad", self.status_var.get().lower())
        self.assertIn("conservó", self.status_var.get().lower())

    def test_ignores_non_lsmw_files_when_checking_existing(self) -> None:
        # Archivos con otro patrón no deben disparar el diálogo
        (self.tmp_salida / "otro.txt").write_text("x", encoding="utf-8")
        (self.tmp_salida / "README.md").write_text("x", encoding="utf-8")

        with self._patch_output_dir(), \
             patch("main.export_sheet_to_tsv", return_value=(self.tmp_salida / "new.txt", 2)), \
             patch("main.messagebox.askyesno") as mock_ask, \
             patch("main.messagebox.showinfo"):
            main.extraer_lsmw_a_txt(self.status_var)

        mock_ask.assert_not_called()


class ExtraerLsmwATxtErrorPathsTest(unittest.TestCase):
    """Verifica que toda excepción durante la extracción se muestre al usuario."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        self.status_var = tk.StringVar(master=self.root)
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_salida = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()
        self.root.destroy()

    def test_shows_error_when_excel_file_not_found(self) -> None:
        with patch("main.OUTPUT_DIR", self.tmp_salida), \
             patch(
                 "main.export_sheet_to_tsv",
                 side_effect=FileNotFoundError("Excel no existe"),
             ), \
             patch("main.messagebox.showerror") as mock_err, \
             patch("main.messagebox.showinfo"):
            main.extraer_lsmw_a_txt(self.status_var)

        mock_err.assert_called_once()
        title, message = mock_err.call_args[0][:2]
        self.assertEqual(title, "Archivo no encontrado")
        self.assertIn("Excel no existe", message)

    def test_shows_error_when_sheet_not_found(self) -> None:
        with patch("main.OUTPUT_DIR", self.tmp_salida), \
             patch(
                 "main.export_sheet_to_tsv",
                 side_effect=ValueError("Hoja no existe"),
             ), \
             patch("main.messagebox.showerror") as mock_err, \
             patch("main.messagebox.showinfo"):
            main.extraer_lsmw_a_txt(self.status_var)

        mock_err.assert_called_once()
        title, message = mock_err.call_args[0][:2]
        self.assertEqual(title, "Hoja no encontrada")
        self.assertIn("Hoja no existe", message)

    def test_shows_error_on_generic_export_failure(self) -> None:
        with patch("main.OUTPUT_DIR", self.tmp_salida), \
             patch(
                 "main.export_sheet_to_tsv",
                 side_effect=RuntimeError("disco lleno"),
             ), \
             patch("main.messagebox.showerror") as mock_err, \
             patch("main.messagebox.showinfo"):
            main.extraer_lsmw_a_txt(self.status_var)

        mock_err.assert_called_once()
        title, message = mock_err.call_args[0][:2]
        self.assertEqual(title, "Error al exportar")
        self.assertIn("disco lleno", message)

    def test_shows_error_on_unexpected_glob_failure(self) -> None:
        # Si OUTPUT_DIR.glob() falla (permisos, path inválido, etc.), la
        # red de seguridad de _show_unexpected_error debe mostrar el error.
        fake_output_dir = MagicMock()
        fake_output_dir.exists.return_value = True
        fake_output_dir.glob.side_effect = OSError("permiso denegado")

        with patch("main.OUTPUT_DIR", fake_output_dir), \
             patch("main.messagebox.showerror") as mock_err:
            main.extraer_lsmw_a_txt(self.status_var)

        mock_err.assert_called_once()
        title, message = mock_err.call_args[0][:2]
        self.assertEqual(title, "Error inesperado al extraer")
        self.assertIn("permiso denegado", message)
        # El mensaje debe incluir el traceback para diagnóstico
        self.assertIn("--- Detalle técnico ---", message)


class ShowUnexpectedErrorTest(unittest.TestCase):
    """_show_unexpected_error muestra dialog + log con traceback."""

    def test_displays_messagebox_with_exception_details(self) -> None:
        try:
            raise RuntimeError("algo falló")
        except RuntimeError as exc:
            with patch("main.messagebox.showerror") as mock_err:
                main._show_unexpected_error("Título de prueba", exc)

        mock_err.assert_called_once()
        title, message = mock_err.call_args[0][:2]
        self.assertEqual(title, "Título de prueba")
        self.assertIn("RuntimeError", message)
        self.assertIn("algo falló", message)
        self.assertIn("--- Detalle técnico ---", message)


class InstallTkExceptionHandlerTest(unittest.TestCase):
    """_install_tk_exception_handler reemplaza el handler default por uno
    que muestra diálogos en vez de imprimir silenciosamente a stderr."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()

    def tearDown(self) -> None:
        self.root.destroy()

    def test_sets_report_callback_exception_attribute(self) -> None:
        original_handler = self.root.report_callback_exception
        main._install_tk_exception_handler(self.root)
        self.assertIsNot(self.root.report_callback_exception, original_handler)
        self.assertTrue(callable(self.root.report_callback_exception))

    def test_handler_shows_dialog_when_invoked(self) -> None:
        main._install_tk_exception_handler(self.root)
        try:
            raise ValueError("uncaught en callback")
        except ValueError:
            with patch("main.messagebox.showerror") as mock_err:
                self.root.report_callback_exception(*sys.exc_info())

        mock_err.assert_called_once()
        title, message = mock_err.call_args[0][:2]
        self.assertEqual(title, "Error inesperado")
        self.assertIn("ValueError", message)
        self.assertIn("uncaught en callback", message)


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
        # Por defecto los tests asumen que hay .txt en salida/ (para que tras
        # completar el flujo el botón quede "normal"). Tests específicos
        # pueden sobreescribir el patch.
        self._hay_txt_patcher = patch("main._hay_txt_en_salida", return_value=True)
        self._hay_txt_patcher.start()
        main._upload_en_curso = False

    def tearDown(self) -> None:
        self._hay_txt_patcher.stop()
        main._upload_en_curso = False
        self.root.destroy()

    def _patch_sap_upload(self, **overrides):
        """Patches por defecto del módulo sap_upload con overrides opcionales.

        Guarda los mocks en `self.mocks` (patch.multiple no devuelve mocks
        pasados como valores explícitos — solo los marcados como DEFAULT).
        """
        self.mocks = {
            "get_latest_txt": MagicMock(return_value=Path("/tmp/LSMW_x.txt")),
            "get_sap_session": MagicMock(return_value=MagicMock(name="session")),
            "run_lsmw_flow": MagicMock(),
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
        fake_path = Path("/tmp/LSMW_test.txt")
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo") as mock_info, \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload(
                 get_latest_txt=MagicMock(return_value=fake_path),
                 get_sap_session=MagicMock(return_value=session),
                 run_lsmw_flow=MagicMock(),
             ):
            subir_a_sap(self.root, self.status_var, self.button)

        self.mocks["get_latest_txt"].assert_called_once()
        self.mocks["get_sap_session"].assert_called_once()
        self.mocks["run_lsmw_flow"].assert_called_once_with(
            session, str(fake_path.parent), fake_path.name
        )
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

    def test_worker_passes_folder_and_filename_to_run_lsmw_flow(self) -> None:
        fake_path = Path("/some/folder/LSMW_20260510_094838.txt")
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload(
                 get_latest_txt=MagicMock(return_value=fake_path),
             ):
            subir_a_sap(self.root, self.status_var, self.button)

        run_flow_call = self.mocks["run_lsmw_flow"].call_args
        self.assertEqual(run_flow_call[0][1], "/some/folder")
        self.assertEqual(run_flow_call[0][2], "LSMW_20260510_094838.txt")

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


# ---------------------------------------------------------------------------
# Estado dinámico del botón "Subir a SAP"
# ---------------------------------------------------------------------------


class HayTxtEnSalidaTest(unittest.TestCase):
    """Helper que detecta archivos LSMW_*.txt en salida/."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_salida = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_returns_false_when_directory_missing(self) -> None:
        with patch("main.OUTPUT_DIR", Path("/no/existe")):
            self.assertFalse(main._hay_txt_en_salida())

    def test_returns_false_when_directory_empty(self) -> None:
        with patch("main.OUTPUT_DIR", self.tmp_salida):
            self.assertFalse(main._hay_txt_en_salida())

    def test_returns_true_when_lsmw_txt_present(self) -> None:
        (self.tmp_salida / "LSMW_20260101_120000.txt").write_text("x", encoding="utf-8")
        with patch("main.OUTPUT_DIR", self.tmp_salida):
            self.assertTrue(main._hay_txt_en_salida())

    def test_returns_false_when_only_non_lsmw_files(self) -> None:
        (self.tmp_salida / "otro.txt").write_text("x", encoding="utf-8")
        (self.tmp_salida / "README.md").write_text("x", encoding="utf-8")
        with patch("main.OUTPUT_DIR", self.tmp_salida):
            self.assertFalse(main._hay_txt_en_salida())


class RefrescarEstadoBotonSubirTest(unittest.TestCase):
    """_refrescar_estado_boton_subir sincroniza el botón con salida/."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        self.button = tk.Button(self.root)
        main._upload_en_curso = False

    def tearDown(self) -> None:
        main._upload_en_curso = False
        self.root.destroy()

    def test_enables_button_when_txt_exists(self) -> None:
        self.button.config(state="disabled")
        with patch("main._hay_txt_en_salida", return_value=True):
            main._refrescar_estado_boton_subir(self.button)
        self.assertEqual(str(self.button["state"]), "normal")

    def test_disables_button_when_no_txt(self) -> None:
        self.button.config(state="normal")
        with patch("main._hay_txt_en_salida", return_value=False):
            main._refrescar_estado_boton_subir(self.button)
        self.assertEqual(str(self.button["state"]), "disabled")

    def test_skips_when_upload_in_progress(self) -> None:
        self.button.config(state="disabled")
        main._upload_en_curso = True
        with patch("main._hay_txt_en_salida", return_value=True):
            main._refrescar_estado_boton_subir(self.button)
        # Botón sigue deshabilitado pese a que hay archivo, porque
        # el worker controla el estado durante el upload.
        self.assertEqual(str(self.button["state"]), "disabled")


class PollEstadoBotonSubirTest(unittest.TestCase):
    """_poll_estado_boton_subir refresca y re-programa cada intervalo."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        self.button = tk.Button(self.root)

    def tearDown(self) -> None:
        self.root.destroy()

    def test_calls_refresh_and_schedules_next_poll(self) -> None:
        self.root.after = MagicMock()
        with patch("main._refrescar_estado_boton_subir") as mock_refresh:
            main._poll_estado_boton_subir(self.root, self.button)

        mock_refresh.assert_called_once_with(self.button)
        self.root.after.assert_called_once()
        scheduled_delay = self.root.after.call_args[0][0]
        self.assertEqual(scheduled_delay, main._POLL_INTERVAL_MS)


class SubirASapFlagTest(unittest.TestCase):
    """Verifica que _upload_en_curso se gestiona correctamente."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.after = lambda delay, fn, *args: fn(*args)
        self.status_var = tk.StringVar(master=self.root)
        self.button = tk.Button(self.root)
        self._hay_txt_patcher = patch("main._hay_txt_en_salida", return_value=True)
        self._hay_txt_patcher.start()
        main._upload_en_curso = False

    def tearDown(self) -> None:
        self._hay_txt_patcher.stop()
        main._upload_en_curso = False
        self.root.destroy()

    def _patch_sap_upload(self, **overrides):
        defaults = {
            "get_latest_txt": MagicMock(return_value=Path("/tmp/LSMW_x.txt")),
            "get_sap_session": MagicMock(return_value=MagicMock()),
            "run_lsmw_flow": MagicMock(),
        }
        defaults.update(overrides)
        return patch.multiple("sap_upload", **defaults)

    def test_flag_is_true_during_worker_execution(self) -> None:
        captured = []

        def capture_flag(*args, **kwargs):
            captured.append(main._upload_en_curso)

        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload(run_lsmw_flow=MagicMock(side_effect=capture_flag)):
            subir_a_sap(self.root, self.status_var, self.button)

        self.assertEqual(captured, [True])

    def test_flag_is_false_after_successful_upload(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload():
            subir_a_sap(self.root, self.status_var, self.button)

        self.assertFalse(main._upload_en_curso)

    def test_flag_is_false_after_lsmw_flow_error(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showerror"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload(
                 run_lsmw_flow=MagicMock(side_effect=Exception("falla"))
             ):
            subir_a_sap(self.root, self.status_var, self.button)

        self.assertFalse(main._upload_en_curso)

    def test_flag_not_set_when_user_cancels(self) -> None:
        with patch("main.messagebox.askyesno", return_value=False):
            subir_a_sap(self.root, self.status_var, self.button)

        self.assertFalse(main._upload_en_curso)

    def test_button_disabled_after_upload_when_no_txt_remains(self) -> None:
        # Simula que al final del flujo el .txt fue borrado/no existe.
        self._hay_txt_patcher.stop()
        with patch("main._hay_txt_en_salida", return_value=False), \
             patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             self._patch_sap_upload():
            subir_a_sap(self.root, self.status_var, self.button)

        self.assertEqual(str(self.button["state"]), "disabled")
        # Restaurar el patcher para que tearDown no falle
        self._hay_txt_patcher = patch("main._hay_txt_en_salida", return_value=True)
        self._hay_txt_patcher.start()


# ---------------------------------------------------------------------------
# Control SOX — diálogo y handler
# ---------------------------------------------------------------------------


class ControlSoxDialogTest(unittest.TestCase):
    """Verifica la construcción del diálogo Control SOX y sus validaciones
    declarativas (combobox readonly, valores válidos, tecla restringida en
    fechas)."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()

    def tearDown(self) -> None:
        self.root.destroy()

    def test_dialog_has_sociedad_combobox_with_valid_values(self) -> None:
        dialog = main.control_sox(self.root)
        try:
            from sox_report import VALID_SOCIEDADES
            self.assertEqual(
                tuple(dialog.sociedad_combo["values"]), VALID_SOCIEDADES
            )
        finally:
            dialog.destroy()

    def test_sociedad_combobox_is_readonly(self) -> None:
        dialog = main.control_sox(self.root)
        try:
            self.assertEqual(str(dialog.sociedad_combo["state"]), "readonly")
        finally:
            dialog.destroy()

    def test_dialog_exposes_form_state_variables(self) -> None:
        dialog = main.control_sox(self.root)
        try:
            self.assertIsInstance(dialog.sociedad_var, tk.StringVar)
            self.assertIsInstance(dialog.desde_var, tk.StringVar)
            self.assertIsInstance(dialog.hasta_var, tk.StringVar)
            self.assertIsInstance(dialog.status_var, tk.StringVar)
        finally:
            dialog.destroy()

    def test_dialog_title(self) -> None:
        dialog = main.control_sox(self.root)
        try:
            self.assertEqual(dialog.title(), "Control SOX")
        finally:
            dialog.destroy()


class GenerarReporteSoxHandlerTest(unittest.TestCase):
    """Pruebas del handler _generar_reporte_sox_handler:
    - validación previa muestra error si los inputs no son válidos
    - cancelar la confirmación no lanza el worker
    - el worker pasa los argumentos correctos al flujo SAP
    - errores del worker se muestran al usuario y reactivan el botón
    """

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        self.dialog = tk.Toplevel(self.root)
        self.dialog.withdraw()
        self.dialog.after = lambda delay, fn, *args: fn(*args)
        self.status_var = tk.StringVar(master=self.dialog)
        self.button = tk.Button(self.dialog)

    def tearDown(self) -> None:
        self.dialog.destroy()
        self.root.destroy()

    def test_shows_error_on_invalid_sociedad(self) -> None:
        with patch("main.messagebox.showerror") as mock_err, \
             patch("main.messagebox.askyesno") as mock_ask:
            main._generar_reporte_sox_handler(
                self.dialog, "XYZ", "01.05.2026", "31.05.2026",
                self.status_var, self.button,
            )

        mock_err.assert_called_once()
        title, message = mock_err.call_args[0][:2]
        self.assertEqual(title, "Datos inválidos")
        self.assertIn("XYZ", message)
        mock_ask.assert_not_called()

    def test_shows_error_on_invalid_date_format(self) -> None:
        with patch("main.messagebox.showerror") as mock_err, \
             patch("main.messagebox.askyesno"):
            main._generar_reporte_sox_handler(
                self.dialog, "ISA", "no-es-fecha", "31.05.2026",
                self.status_var, self.button,
            )

        mock_err.assert_called_once()
        self.assertEqual(mock_err.call_args[0][0], "Datos inválidos")

    def test_shows_error_when_hasta_before_desde(self) -> None:
        with patch("main.messagebox.showerror") as mock_err, \
             patch("main.messagebox.askyesno"):
            main._generar_reporte_sox_handler(
                self.dialog, "ISA", "31.05.2026", "01.05.2026",
                self.status_var, self.button,
            )

        mock_err.assert_called_once()
        message = mock_err.call_args[0][1]
        self.assertIn("mayor o igual", message)

    def test_cancel_confirmation_does_not_start_worker(self) -> None:
        with patch("main.messagebox.askyesno", return_value=False), \
             patch("main.threading.Thread") as mock_thread:
            main._generar_reporte_sox_handler(
                self.dialog, "ISA", "01.05.2026", "31.05.2026",
                self.status_var, self.button,
            )

        mock_thread.assert_not_called()

    def test_happy_path_calls_generar_reporte_sox_with_normalized_inputs(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("sox_report.get_sap_session", return_value=MagicMock()), \
             patch(
                 "sox_report.generar_reporte_sox",
                 return_value=("/tmp/salida", "SOX_ISA_x.xlsx"),
             ) as mock_flow:
            main._generar_reporte_sox_handler(
                self.dialog, "isa", "01.05.2026", "31.05.2026",
                self.status_var, self.button,
            )

        mock_flow.assert_called_once()
        args = mock_flow.call_args[0]
        # args: (session, sociedad, desde, hasta)
        self.assertEqual(args[1], "ISA")  # normalizada a uppercase
        self.assertEqual(args[2], "01.05.2026")
        self.assertEqual(args[3], "31.05.2026")

    def test_worker_disables_button_during_execution_and_reenables_after(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("sox_report.get_sap_session", return_value=MagicMock()), \
             patch(
                 "sox_report.generar_reporte_sox",
                 return_value=("/tmp", "x.xlsx"),
             ):
            main._generar_reporte_sox_handler(
                self.dialog, "ISA", "01.05.2026", "31.05.2026",
                self.status_var, self.button,
            )

        # tras el worker, debe estar habilitado de nuevo
        self.assertEqual(str(self.button["state"]), "normal")

    def test_worker_shows_error_when_sap_session_fails(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showerror") as mock_err, \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch(
                 "sox_report.get_sap_session",
                 side_effect=RuntimeError("no SAP"),
             ):
            main._generar_reporte_sox_handler(
                self.dialog, "ISA", "01.05.2026", "31.05.2026",
                self.status_var, self.button,
            )

        mock_err.assert_called_once()
        self.assertEqual(mock_err.call_args[0][0], "Error generando reporte SOX")
        self.assertIn("no SAP", mock_err.call_args[0][1])

    def test_worker_shows_error_when_flow_raises(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showerror") as mock_err, \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("sox_report.get_sap_session", return_value=MagicMock()), \
             patch(
                 "sox_report.generar_reporte_sox",
                 side_effect=Exception("paso 4 falló"),
             ):
            main._generar_reporte_sox_handler(
                self.dialog, "ISA", "01.05.2026", "31.05.2026",
                self.status_var, self.button,
            )

        mock_err.assert_called_once()
        self.assertIn("paso 4 falló", mock_err.call_args[0][1])


if __name__ == "__main__":
    unittest.main()
