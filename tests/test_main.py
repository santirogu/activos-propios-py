import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from main import export_sheet_to_tsv  # noqa: E402


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


if __name__ == "__main__":
    unittest.main()
