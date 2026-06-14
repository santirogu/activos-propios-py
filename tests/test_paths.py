"""Tests para src/paths.py — helpers de modo dev / bundled (PyInstaller).

`paths` es el único módulo que distingue entre correr `python src/main.py`
y correr el `.exe` bundleado. Si esta lógica se rompe, en bundled mode
`salida/` y `resources/` podrían terminar dentro del temp `_MEIPASS`
(que PyInstaller borra al cerrar la app) en vez de al lado del `.exe`,
perdiéndose los outputs del usuario.
"""

import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import paths  # noqa: E402


class ProjectRootTest(unittest.TestCase):
    """`project_root()` decide dónde viven `salida/` y `resources/`.
    Si esto se rompe en frozen mode, los outputs del usuario van al
    `_MEIPASS` temporal y desaparecen al cerrar el .exe."""

    def test_dev_mode_returns_repo_root(self) -> None:
        """Sin sys.frozen, `project_root()` apunta al padre de src/."""
        # En este test la importación de `paths` ya fijó PROJECT_ROOT al
        # padre de src/; verificamos que `project_root()` (que se evalúa
        # cada vez) coincide con esa estructura.
        with patch.object(sys, "frozen", False, create=True):
            root = paths.project_root()
            self.assertTrue(
                (root / "src" / "paths.py").exists(),
                f"Esperaba ver src/paths.py bajo {root}",
            )

    def test_frozen_mode_returns_dirname_of_executable(self) -> None:
        """Con sys.frozen=True, `project_root()` apunta a la carpeta del
        .exe (NO al `_MEIPASS` temporal). Esto garantiza que salida/ y
        resources/ vivan al lado del binario y persistan entre runs."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fake_exe = Path(tmpdir) / "GestionActivosFijos.exe"
            fake_exe.touch()
            with patch.object(sys, "frozen", True, create=True), \
                 patch.object(sys, "executable", str(fake_exe)):
                root = paths.project_root()
                self.assertEqual(root, Path(tmpdir).resolve())


class BundledResourcePathTest(unittest.TestCase):
    """`bundled_resource_path()` lee de `sys._MEIPASS` (read-only, donde
    PyInstaller descomprime los datas) en bundled mode, o del repo en
    dev. Es lo que usa `branding.LOGO_PATH` para que el logo aparezca
    igual en ambos modos."""

    def test_without_meipass_uses_project_root(self) -> None:
        """En dev mode (sin _MEIPASS) el path resuelve al repo."""
        original = getattr(sys, "_MEIPASS", None)
        if original is not None:
            del sys._MEIPASS
        try:
            result = paths.bundled_resource_path("resources/logo_hub_isa.png")
            self.assertEqual(
                result, paths.PROJECT_ROOT / "resources/logo_hub_isa.png",
            )
        finally:
            if original is not None:
                sys._MEIPASS = original  # type: ignore[attr-defined]

    def test_with_meipass_uses_meipass(self) -> None:
        """En bundled mode el path resuelve al temp _MEIPASS."""
        with patch.object(sys, "_MEIPASS", "/fake/_MEIPASS_TEMP", create=True):
            result = paths.bundled_resource_path(
                "resources/logo_hub_isa.png"
            )
            self.assertEqual(
                result, Path("/fake/_MEIPASS_TEMP/resources/logo_hub_isa.png"),
            )


class AsegurarFormatoDinamicoTest(unittest.TestCase):
    """`asegurar_formato_dinamico()` implementa el "factory default":
    en el primer arranque del .exe copia el `Formato_Dinamico_` bundleado
    a `<EXE_DIR>/resources/` para que el usuario lo pueda editar sin
    rebuild. En arranques posteriores NO sobrescribe.
    """

    def setUp(self) -> None:
        self.tmpdir = Path(tempfile.mkdtemp())
        # Simulamos:
        #   <tmpdir>/                       = "carpeta del .exe"
        #   <tmpdir>/resources/             = externo, editable
        #   <tmpdir>/_MEIPASS_FAKE/resources/{archivo}  = bundleado read-only
        self.resources_dir = self.tmpdir / "resources"
        self.destino = self.resources_dir / paths.FORMATO_DINAMICO_NOMBRE

        bundle_resources = self.tmpdir / "_MEIPASS_FAKE" / "resources"
        bundle_resources.mkdir(parents=True)
        self.origen = bundle_resources / paths.FORMATO_DINAMICO_NOMBRE
        self.origen.write_bytes(b"factory_default_content")

        self._patches = [
            patch.object(paths, "RESOURCES_DIR", self.resources_dir),
            patch.object(paths, "formato_dinamico_path", lambda: self.destino),
            patch.object(
                paths, "bundled_resource_path",
                lambda rel: self.tmpdir / "_MEIPASS_FAKE" / rel,
            ),
        ]
        for p in self._patches:
            p.start()

    def tearDown(self) -> None:
        for p in self._patches:
            p.stop()
        shutil.rmtree(self.tmpdir)

    def test_first_run_copies_factory_default(self) -> None:
        """Primer arranque: archivo externo no existe → se copia del
        bundle y se devuelve (path, True)."""
        self.assertFalse(self.destino.exists())

        path, recien_creado = paths.asegurar_formato_dinamico()

        self.assertEqual(path, self.destino)
        self.assertTrue(recien_creado)
        self.assertTrue(self.destino.exists())
        self.assertEqual(
            self.destino.read_bytes(), b"factory_default_content"
        )

    def test_second_run_preserves_user_edits(self) -> None:
        """Si el archivo externo ya existe (usuario lo editó), NO se
        sobrescribe: se devuelve (path, False)."""
        self.resources_dir.mkdir()
        self.destino.write_bytes(b"edited_by_user")

        path, recien_creado = paths.asegurar_formato_dinamico()

        self.assertEqual(path, self.destino)
        self.assertFalse(recien_creado)
        self.assertEqual(
            self.destino.read_bytes(), b"edited_by_user",
            "El factory default NO debe sobrescribir ediciones del usuario.",
        )

    def test_bundle_missing_returns_destino_without_creating(self) -> None:
        """Si ni el archivo externo ni el bundle existen (caso raro: dev
        mode con el repo manipulado, o build corrupto), la función NO
        crashea — devuelve el path esperado para que el caller reporte
        un error claro al usuario."""
        self.origen.unlink()

        path, recien_creado = paths.asegurar_formato_dinamico()

        self.assertEqual(path, self.destino)
        self.assertFalse(recien_creado)
        self.assertFalse(self.destino.exists())


if __name__ == "__main__":
    unittest.main()
