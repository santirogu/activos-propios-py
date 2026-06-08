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


class SapComApartmentTest(unittest.TestCase):
    """Pruebas del context manager `_sap_com_apartment` que inicializa el
    apartamento COM antes de cualquier llamada SAP desde threads que no
    sean el main. Sin esto, GetObject('SAPGUI') falla en los workers."""

    def test_is_no_op_when_pythoncom_unavailable(self) -> None:
        # En Mac no existe pythoncom — el CM debe ser no-op (no lanzar).
        with patch.dict(sys.modules, {"pythoncom": None}):
            with main._sap_com_apartment():
                pass  # no debe lanzar

    def test_calls_co_initialize_and_co_uninitialize(self) -> None:
        fake_pythoncom = MagicMock()
        with patch.dict(sys.modules, {"pythoncom": fake_pythoncom}):
            with main._sap_com_apartment():
                fake_pythoncom.CoInitialize.assert_called_once()
                fake_pythoncom.CoUninitialize.assert_not_called()
            fake_pythoncom.CoUninitialize.assert_called_once()

    def test_calls_co_uninitialize_even_when_body_raises(self) -> None:
        fake_pythoncom = MagicMock()
        with patch.dict(sys.modules, {"pythoncom": fake_pythoncom}):
            with self.assertRaises(RuntimeError):
                with main._sap_com_apartment():
                    raise RuntimeError("boom")
        fake_pythoncom.CoUninitialize.assert_called_once()

    def test_swallows_co_uninitialize_errors(self) -> None:
        """Si CoUninitialize lanza, no debe propagarse (es cleanup)."""
        fake_pythoncom = MagicMock()
        fake_pythoncom.CoUninitialize.side_effect = Exception("denied")
        with patch.dict(sys.modules, {"pythoncom": fake_pythoncom}):
            # No debe lanzar pese al error en uninitialize
            with main._sap_com_apartment():
                pass


class TestConexionSapHandlerTest(unittest.TestCase):
    """Pruebas del handler del botón 'Test conexión SAP'."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()

    def tearDown(self) -> None:
        self.root.destroy()

    def test_shows_info_messagebox_on_success(self) -> None:
        with patch(
            "sap_upload.diagnosticar_conexion_sap",
            return_value=(True, "Todo OK\nSesión: PRD/100/SROCK"),
        ), patch("main.messagebox.showinfo") as mock_info, \
             patch("main.messagebox.showwarning") as mock_warn:
            main._test_conexion_sap_handler()

        mock_info.assert_called_once()
        title, message = mock_info.call_args[0][:2]
        self.assertEqual(title, "Test conexión SAP — OK")
        self.assertIn("Sesión: PRD/100/SROCK", message)
        mock_warn.assert_not_called()

    def test_shows_warning_messagebox_on_failure(self) -> None:
        with patch(
            "sap_upload.diagnosticar_conexion_sap",
            return_value=(False, "SAP no abierto"),
        ), patch("main.messagebox.showwarning") as mock_warn, \
             patch("main.messagebox.showinfo") as mock_info:
            main._test_conexion_sap_handler()

        mock_warn.assert_called_once()
        title, message = mock_warn.call_args[0][:2]
        self.assertEqual(title, "Test conexión SAP — Problema")
        self.assertIn("SAP no abierto", message)
        mock_info.assert_not_called()

    def test_catches_unexpected_exception_and_shows_error(self) -> None:
        with patch(
            "sap_upload.diagnosticar_conexion_sap",
            side_effect=RuntimeError("crash inesperado"),
        ), patch("main.messagebox.showerror") as mock_err:
            main._test_conexion_sap_handler()

        mock_err.assert_called_once()
        title, message = mock_err.call_args[0][:2]
        self.assertEqual(title, "Error en test de conexión SAP")
        self.assertIn("crash inesperado", message)


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

    def test_worker_runs_under_sap_com_apartment(self) -> None:
        """El worker debe envolverse en _sap_com_apartment para que las
        llamadas COM funcionen desde el thread de background."""
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("main._sap_com_apartment") as mock_cm, \
             self._patch_sap_upload():
            mock_cm.return_value.__enter__ = MagicMock(return_value=None)
            mock_cm.return_value.__exit__ = MagicMock(return_value=False)
            subir_a_sap(self.root, self.status_var, self.button)

        mock_cm.assert_called_once()


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


# Nota: la clase `PollEstadoBotonSubirTest` se eliminó al refactorizar
# main(). El polling ahora vive INLINE dentro de `abrir_activos_fijos`
# (scoped al frame, se cancela al destruirlo), por lo que ya no existe
# `_poll_estado_boton_subir`. La función `_refrescar_estado_boton_subir`
# sí se conserva — los tests de su comportamiento siguen en
# `RefrescarEstadoBotonSubirTest`.


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
    """Verifica que `control_sox` reemplaza la vista del menú (frame_menu)
    por el formulario SOX (frame_sox) en la misma ventana, expone las
    StringVars/widgets como atributos del frame, y que el botón Atrás
    revierte la vista al menú."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        # Replicamos la estructura que construye `main.main()`: el menú
        # principal vive dentro de un Frame que `control_sox` puede ocultar.
        self.frame_menu = tk.Frame(self.root)
        self.frame_menu.pack(fill="both", expand=True)

    def tearDown(self) -> None:
        self.root.destroy()

    def test_hides_frame_menu_when_invoked(self) -> None:
        """Al entrar al form SOX, el menú deja de estar gestionado por pack."""
        self.assertTrue(self.frame_menu.winfo_manager())  # actualmente packed
        frame_sox = main.control_sox(self.root, self.frame_menu)
        try:
            self.assertEqual(self.frame_menu.winfo_manager(), "")
        finally:
            frame_sox.destroy()

    def test_frame_has_sociedad_combobox_with_valid_values(self) -> None:
        frame_sox = main.control_sox(self.root, self.frame_menu)
        try:
            from sox_report import VALID_SOCIEDADES
            self.assertEqual(
                tuple(frame_sox.sociedad_combo["values"]), VALID_SOCIEDADES
            )
        finally:
            frame_sox.destroy()

    def test_sociedad_combobox_is_readonly(self) -> None:
        frame_sox = main.control_sox(self.root, self.frame_menu)
        try:
            self.assertEqual(str(frame_sox.sociedad_combo["state"]), "readonly")
        finally:
            frame_sox.destroy()

    def test_frame_exposes_form_state_variables(self) -> None:
        frame_sox = main.control_sox(self.root, self.frame_menu)
        try:
            self.assertIsInstance(frame_sox.sociedad_var, tk.StringVar)
            self.assertIsInstance(frame_sox.desde_var, tk.StringVar)
            self.assertIsInstance(frame_sox.hasta_var, tk.StringVar)
            self.assertIsInstance(frame_sox.status_var, tk.StringVar)
        finally:
            frame_sox.destroy()

    def test_frame_exposes_back_button(self) -> None:
        """El form debe tener un botón Atrás expuesto como `btn_atras`."""
        frame_sox = main.control_sox(self.root, self.frame_menu)
        try:
            self.assertIsInstance(frame_sox.btn_atras, tk.Button)
        finally:
            frame_sox.destroy()

    def test_back_button_destroys_form_and_reshows_menu(self) -> None:
        """Click en Atrás → frame_sox se destruye y frame_menu vuelve a
        mostrarse en la ventana principal."""
        frame_sox = main.control_sox(self.root, self.frame_menu)
        # menú oculto mientras estamos en el form
        self.assertEqual(self.frame_menu.winfo_manager(), "")

        frame_sox.btn_atras.invoke()

        self.assertFalse(frame_sox.winfo_exists())
        self.assertEqual(self.frame_menu.winfo_manager(), "pack")

    def test_date_fields_use_date_entry_calendar_widget(self) -> None:
        """Los campos Desde y Hasta deben ser DateEntry (tkcalendar) para
        que el usuario pueda elegir la fecha de un calendario emergente."""
        from tkcalendar import DateEntry

        frame_sox = main.control_sox(self.root, self.frame_menu)
        try:
            self.assertIsInstance(frame_sox.desde_entry, DateEntry)
            self.assertIsInstance(frame_sox.hasta_entry, DateEntry)
        finally:
            frame_sox.destroy()

    def test_date_entries_emit_value_in_ddmmyyyy_format(self) -> None:
        """El valor que escriben los DateEntry en la StringVar debe estar
        en formato dd.mm.aaaa, listo para validar_fecha."""
        from datetime import date

        frame_sox = main.control_sox(self.root, self.frame_menu)
        try:
            frame_sox.desde_entry.set_date(date(2026, 5, 1))
            frame_sox.hasta_entry.set_date(date(2026, 5, 31))
            self.assertEqual(frame_sox.desde_var.get(), "01.05.2026")
            self.assertEqual(frame_sox.hasta_var.get(), "31.05.2026")
        finally:
            frame_sox.destroy()

    def test_date_entries_initialize_with_today(self) -> None:
        """Al entrar al form, los DateEntry arrancan con la fecha actual."""
        from datetime import date

        frame_sox = main.control_sox(self.root, self.frame_menu)
        try:
            self.assertEqual(frame_sox.desde_entry.get_date(), date.today())
            self.assertEqual(frame_sox.hasta_entry.get_date(), date.today())
        finally:
            frame_sox.destroy()


class AbrirActivosFijosTest(unittest.TestCase):
    """`abrir_activos_fijos(root, frame_menu)` reemplaza la vista del menú
    por un sub-formulario con dos botones: Extraer (`extraer_lsmw_a_txt`)
    y Creación de Activo (`subir_a_sap`), y arranca polling scoped al
    frame para habilitar/deshabilitar Creación de Activo según haya
    LSMW_*.txt en salida/."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        # Stub after para evitar que el polling re-programe callbacks
        # en el event loop durante los tests.
        self.root.after = MagicMock(return_value="dummy_id")
        self.root.after_cancel = MagicMock()
        self.frame_menu = tk.Frame(self.root)
        self.frame_menu.pack(fill="both", expand=True)

    def tearDown(self) -> None:
        self.root.destroy()

    def test_hides_frame_menu_when_invoked(self) -> None:
        self.assertTrue(self.frame_menu.winfo_manager())
        frame = main.abrir_activos_fijos(self.root, self.frame_menu)
        try:
            self.assertEqual(self.frame_menu.winfo_manager(), "")
        finally:
            frame.destroy()

    def test_exposes_all_three_buttons_in_order(self) -> None:
        """La vista Activos Fijos tiene 3 botones primarios en orden:
        Extraer información en txt → Creación de Activo → Extraer Activos Creados."""
        frame = main.abrir_activos_fijos(self.root, self.frame_menu)
        try:
            self.assertIsInstance(frame.btn_extraer, tk.Button)
            self.assertEqual(
                frame.btn_extraer.cget("text"), "Extraer información en txt"
            )
            self.assertIsInstance(frame.btn_creacion, tk.Button)
            self.assertEqual(
                frame.btn_creacion.cget("text"), "Creación de Activo"
            )
            self.assertIsInstance(frame.btn_extraer_creados, tk.Button)
            self.assertEqual(
                frame.btn_extraer_creados.cget("text"),
                "Extraer Activos Creados",
            )
        finally:
            frame.destroy()

    def test_creacion_button_starts_disabled(self) -> None:
        """Mismo comportamiento que el viejo 'Subir a SAP': arranca disabled
        y el polling lo habilita cuando hay un LSMW_*.txt en salida/."""
        with patch("main._hay_txt_en_salida", return_value=False):
            frame = main.abrir_activos_fijos(self.root, self.frame_menu)
            try:
                self.assertEqual(str(frame.btn_creacion["state"]), "disabled")
            finally:
                frame.destroy()

    def test_creacion_button_enabled_when_txt_present(self) -> None:
        with patch("main._hay_txt_en_salida", return_value=True):
            frame = main.abrir_activos_fijos(self.root, self.frame_menu)
            try:
                self.assertEqual(str(frame.btn_creacion["state"]), "normal")
            finally:
                frame.destroy()

    def test_back_button_destroys_frame_and_reshows_menu(self) -> None:
        frame = main.abrir_activos_fijos(self.root, self.frame_menu)
        self.assertEqual(self.frame_menu.winfo_manager(), "")

        frame.btn_atras.invoke()

        self.assertFalse(frame.winfo_exists())
        self.assertEqual(self.frame_menu.winfo_manager(), "pack")


class AbrirExtraerCreadosTest(unittest.TestCase):
    """`abrir_extraer_creados(root, frame_activos)` reemplaza la vista de
    Activos Fijos por un sub-formulario con un campo "Usuario SAP" y un
    botón "Ejecutar". El botón Ejecutar aún no implementa lógica real —
    muestra un messagebox 'En desarrollo'."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        # Simular el padre (frame_activos) que estaría empacado cuando se
        # navega desde el menú principal.
        self.frame_activos = tk.Frame(self.root)
        self.frame_activos.pack(fill="both", expand=True)

    def tearDown(self) -> None:
        self.root.destroy()

    def test_hides_frame_activos_when_invoked(self) -> None:
        self.assertTrue(self.frame_activos.winfo_manager())
        frame = main.abrir_extraer_creados(self.root, self.frame_activos)
        try:
            self.assertEqual(self.frame_activos.winfo_manager(), "")
        finally:
            frame.destroy()

    def test_exposes_usuario_sap_field(self) -> None:
        frame = main.abrir_extraer_creados(self.root, self.frame_activos)
        try:
            self.assertIsInstance(frame.usuario_var, tk.StringVar)
            self.assertIsInstance(frame.usuario_entry, tk.Entry)
        finally:
            frame.destroy()

    def test_exposes_ejecutar_button(self) -> None:
        frame = main.abrir_extraer_creados(self.root, self.frame_activos)
        try:
            self.assertIsInstance(frame.btn_ejecutar, tk.Button)
            self.assertEqual(frame.btn_ejecutar.cget("text"), "Ejecutar")
        finally:
            frame.destroy()

    def test_ejecutar_button_calls_extraer_handler(self) -> None:
        """El botón Ejecutar invoca `_extraer_activos_creados_handler`
        pasándole el root, el valor del entry, el botón y el btn_atras."""
        frame = main.abrir_extraer_creados(self.root, self.frame_activos)
        try:
            frame.usuario_var.set("1017209574")
            with patch("main._extraer_activos_creados_handler") as mock_handler:
                frame.btn_ejecutar.invoke()
            mock_handler.assert_called_once_with(
                self.root, "1017209574", frame.btn_ejecutar, frame.btn_atras,
            )
        finally:
            frame.destroy()

    def test_back_button_returns_to_activos_fijos(self) -> None:
        frame = main.abrir_extraer_creados(self.root, self.frame_activos)
        self.assertEqual(self.frame_activos.winfo_manager(), "")

        frame.btn_atras.invoke()

        self.assertFalse(frame.winfo_exists())
        self.assertEqual(self.frame_activos.winfo_manager(), "pack")


class AbrirSoxMenuTest(unittest.TestCase):
    """`abrir_sox_menu(root, frame_menu)` reemplaza la vista del menú por
    un sub-formulario intermedio con un único botón "HUB.PPE.01 Creación
    de Activos Fijos" que abre el formulario clásico con parámetros."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        self.frame_menu = tk.Frame(self.root)
        self.frame_menu.pack(fill="both", expand=True)

    def tearDown(self) -> None:
        self.root.destroy()

    def test_hides_frame_menu_when_invoked(self) -> None:
        self.assertTrue(self.frame_menu.winfo_manager())
        frame = main.abrir_sox_menu(self.root, self.frame_menu)
        try:
            self.assertEqual(self.frame_menu.winfo_manager(), "")
        finally:
            frame.destroy()

    def test_exposes_hub_ppe_01_button(self) -> None:
        frame = main.abrir_sox_menu(self.root, self.frame_menu)
        try:
            self.assertIsInstance(frame.btn_hub_ppe_01, tk.Button)
            self.assertIn(
                "HUB.PPE.01", frame.btn_hub_ppe_01.cget("text")
            )
        finally:
            frame.destroy()

    def test_back_button_destroys_frame_and_reshows_menu(self) -> None:
        frame = main.abrir_sox_menu(self.root, self.frame_menu)
        self.assertEqual(self.frame_menu.winfo_manager(), "")

        frame.btn_atras.invoke()

        self.assertFalse(frame.winfo_exists())
        self.assertEqual(self.frame_menu.winfo_manager(), "pack")

    def test_hub_ppe_01_opens_control_sox(self) -> None:
        """Click en HUB.PPE.01 debe llamar a `control_sox(root, frame_sox_menu)`."""
        frame = main.abrir_sox_menu(self.root, self.frame_menu)
        try:
            with patch("main.control_sox") as mock_control_sox:
                frame.btn_hub_ppe_01.invoke()
            mock_control_sox.assert_called_once_with(self.root, frame)
        finally:
            if frame.winfo_exists():
                frame.destroy()


class GenerarReporteSoxHandlerTest(unittest.TestCase):
    """Pruebas del handler _generar_reporte_sox_handler:
    - validación previa muestra error si los inputs no son válidos
    - cancelar la confirmación no lanza el worker
    - el worker pasa los argumentos correctos al flujo SAP
    - errores del worker se muestran al usuario y reactivan los botones
    - el botón Atrás se deshabilita durante el worker y se reactiva después
    """

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        # Stub `root.after` para ejecutar callbacks inmediatamente (en vez
        # de programarlos en el event loop) — los callbacks thread-safe del
        # handler usan root.after en lugar del antiguo dialog.after.
        self.root.after = lambda delay, fn, *args: fn(*args)
        self.status_var = tk.StringVar(master=self.root)
        self.button = tk.Button(self.root)
        self.btn_atras = tk.Button(self.root)

    def tearDown(self) -> None:
        self.root.destroy()

    def test_shows_error_on_invalid_sociedad(self) -> None:
        with patch("main.messagebox.showerror") as mock_err, \
             patch("main.messagebox.askyesno") as mock_ask:
            main._generar_reporte_sox_handler(
                self.root, "XYZ", "01.05.2026", "31.05.2026",
                self.status_var, self.button, self.btn_atras,
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
                self.root, "ISA", "no-es-fecha", "31.05.2026",
                self.status_var, self.button, self.btn_atras,
            )

        mock_err.assert_called_once()
        self.assertEqual(mock_err.call_args[0][0], "Datos inválidos")

    def test_shows_error_when_hasta_before_desde(self) -> None:
        with patch("main.messagebox.showerror") as mock_err, \
             patch("main.messagebox.askyesno"):
            main._generar_reporte_sox_handler(
                self.root, "ISA", "31.05.2026", "01.05.2026",
                self.status_var, self.button, self.btn_atras,
            )

        mock_err.assert_called_once()
        message = mock_err.call_args[0][1]
        self.assertIn("mayor o igual", message)

    def test_cancel_confirmation_does_not_start_worker(self) -> None:
        with patch("main.messagebox.askyesno", return_value=False), \
             patch("main.threading.Thread") as mock_thread:
            main._generar_reporte_sox_handler(
                self.root, "ISA", "01.05.2026", "31.05.2026",
                self.status_var, self.button, self.btn_atras,
            )

        mock_thread.assert_not_called()

    def test_happy_path_calls_generar_reporte_sox_with_normalized_inputs(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("sox_report.get_sap_session", return_value=MagicMock()), \
             patch(
                 "sox_report.generar_reporte_sox",
                 return_value=("/tmp/salida", "Población_ISA_31.05.2026.xlsx"),
             ) as mock_flow:
            main._generar_reporte_sox_handler(
                self.root, "isa", "01.05.2026", "31.05.2026",
                self.status_var, self.button, self.btn_atras,
            )

        mock_flow.assert_called_once()
        args = mock_flow.call_args[0]
        # args: (session, sociedad, desde, hasta)
        self.assertEqual(args[1], "ISA")  # normalizada a uppercase
        self.assertEqual(args[2], "01.05.2026")
        self.assertEqual(args[3], "31.05.2026")

    def test_worker_disables_buttons_during_execution_and_reenables_after(self) -> None:
        """Tanto el botón Generar como el Atrás deben re-habilitarse tras
        el worker (no queremos dejar al usuario sin forma de volver)."""
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("sox_report.get_sap_session", return_value=MagicMock()), \
             patch(
                 "sox_report.generar_reporte_sox",
                 return_value=("/tmp", "x.xlsx"),
             ):
            main._generar_reporte_sox_handler(
                self.root, "ISA", "01.05.2026", "31.05.2026",
                self.status_var, self.button, self.btn_atras,
            )

        self.assertEqual(str(self.button["state"]), "normal")
        self.assertEqual(str(self.btn_atras["state"]), "normal")

    def test_back_button_disabled_during_worker(self) -> None:
        """El botón Atrás se deshabilita ANTES de lanzar el worker para
        evitar que el usuario vuelva al menú a mitad de un flujo SAP."""
        # Capturamos el estado del btn_atras justo después de pasar la
        # validación + confirmación, antes de que el worker (síncrono via
        # _SyncFakeThread) lo reactive al final.
        estado_durante_worker = []

        original_thread = _SyncFakeThread

        class CapturingFakeThread(_SyncFakeThread):
            def start(inner_self):
                estado_durante_worker.append(str(self.btn_atras["state"]))
                super().start()

        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", CapturingFakeThread), \
             patch("sox_report.get_sap_session", return_value=MagicMock()), \
             patch(
                 "sox_report.generar_reporte_sox",
                 return_value=("/tmp", "x.xlsx"),
             ):
            main._generar_reporte_sox_handler(
                self.root, "ISA", "01.05.2026", "31.05.2026",
                self.status_var, self.button, self.btn_atras,
            )

        self.assertEqual(estado_durante_worker, ["disabled"])

    def test_worker_shows_error_when_sap_session_fails(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showerror") as mock_err, \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch(
                 "sox_report.get_sap_session",
                 side_effect=RuntimeError("no SAP"),
             ):
            main._generar_reporte_sox_handler(
                self.root, "ISA", "01.05.2026", "31.05.2026",
                self.status_var, self.button, self.btn_atras,
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
                self.root, "ISA", "01.05.2026", "31.05.2026",
                self.status_var, self.button, self.btn_atras,
            )

        mock_err.assert_called_once()
        self.assertIn("paso 4 falló", mock_err.call_args[0][1])

    def test_worker_runs_under_sap_com_apartment(self) -> None:
        """El worker del SOX también debe envolverse en _sap_com_apartment."""
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("main._sap_com_apartment") as mock_cm, \
             patch("sox_report.get_sap_session", return_value=MagicMock()), \
             patch(
                 "sox_report.generar_reporte_sox",
                 return_value=("/tmp", "x.xlsx"),
             ):
            mock_cm.return_value.__enter__ = MagicMock(return_value=None)
            mock_cm.return_value.__exit__ = MagicMock(return_value=False)
            main._generar_reporte_sox_handler(
                self.root, "ISA", "01.05.2026", "31.05.2026",
                self.status_var, self.button, self.btn_atras,
            )

        mock_cm.assert_called_once()


class ExtraerActivosCreadosHandlerTest(unittest.TestCase):
    """Pruebas del handler _extraer_activos_creados_handler:
    - validación del Usuario SAP (vacío → error, válido → confirmación)
    - cancelar la confirmación no lanza el worker
    - el worker pasa el usuario normalizado a `extraer_activos_creados`
    - errores del worker se muestran y reactivan los botones
    - botón Atrás se deshabilita durante worker y se reactiva al final
    """

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        # Stub `root.after` para invocar callbacks inmediatamente.
        self.root.after = lambda delay, fn, *args: fn(*args)
        self.button = tk.Button(self.root)
        self.btn_atras = tk.Button(self.root)

    def tearDown(self) -> None:
        self.root.destroy()

    def test_shows_error_on_empty_usuario(self) -> None:
        with patch("main.messagebox.showerror") as mock_err, \
             patch("main.messagebox.askyesno") as mock_ask:
            main._extraer_activos_creados_handler(
                self.root, "", self.button, self.btn_atras,
            )

        mock_err.assert_called_once()
        title = mock_err.call_args[0][0]
        self.assertEqual(title, "Datos inválidos")
        mock_ask.assert_not_called()

    def test_cancel_confirmation_does_not_start_worker(self) -> None:
        with patch("main.messagebox.askyesno", return_value=False), \
             patch("main.threading.Thread") as mock_thread:
            main._extraer_activos_creados_handler(
                self.root, "1017209574", self.button, self.btn_atras,
            )

        mock_thread.assert_not_called()

    def test_happy_path_calls_extraer_with_normalized_usuario(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("extraer_activos_creados.get_sap_session",
                   return_value=MagicMock()), \
             patch("extraer_activos_creados.extraer_activos_creados",
                   return_value=("/tmp", "x.xlsx")) as mock_flow:
            main._extraer_activos_creados_handler(
                self.root, "  1017209574  ", self.button, self.btn_atras,
            )

        # Usuario debe llegar con strip() aplicado.
        mock_flow.assert_called_once()
        args = mock_flow.call_args[0]
        self.assertEqual(args[1], "1017209574")

    def test_worker_disables_buttons_and_reenables_after(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("extraer_activos_creados.get_sap_session",
                   return_value=MagicMock()), \
             patch("extraer_activos_creados.extraer_activos_creados",
                   return_value=("/tmp", "x.xlsx")):
            main._extraer_activos_creados_handler(
                self.root, "USR1", self.button, self.btn_atras,
            )

        # Tras el worker ambos botones deben estar habilitados de nuevo.
        self.assertEqual(str(self.button["state"]), "normal")
        self.assertEqual(str(self.btn_atras["state"]), "normal")

    def test_worker_shows_error_when_sap_session_fails(self) -> None:
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showerror") as mock_err, \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("extraer_activos_creados.get_sap_session",
                   side_effect=RuntimeError("no SAP")):
            main._extraer_activos_creados_handler(
                self.root, "USR1", self.button, self.btn_atras,
            )

        mock_err.assert_called_once()
        self.assertIn("no SAP", mock_err.call_args[0][1])

    def test_worker_runs_under_sap_com_apartment(self) -> None:
        """El worker se envuelve en _sap_com_apartment (igual que SOX)."""
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("main._sap_com_apartment") as mock_cm, \
             patch("extraer_activos_creados.get_sap_session",
                   return_value=MagicMock()), \
             patch("extraer_activos_creados.extraer_activos_creados",
                   return_value=("/tmp", "x.xlsx")):
            mock_cm.return_value.__enter__ = MagicMock(return_value=None)
            mock_cm.return_value.__exit__ = MagicMock(return_value=False)
            main._extraer_activos_creados_handler(
                self.root, "USR1", self.button, self.btn_atras,
            )

        mock_cm.assert_called_once()


if __name__ == "__main__":
    unittest.main()
