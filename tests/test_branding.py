"""Tests para src/branding.py — paleta + RoundedButton.

`RoundedButton` es un widget custom basado en `tk.Canvas` que pinta el
botón con esquinas redondeadas (Tk no soporta `border-radius` nativo).
Espeja la API mínima de `tk.Button` que el proyecto necesita: text,
command, config(state=), cget(), invoke(), __getitem__.
"""

import sys
import tkinter as tk
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import branding  # noqa: E402


class RoundedButtonConstructorTest(unittest.TestCase):
    """Construcción básica + atributos visibles vía API."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()

    def tearDown(self) -> None:
        self.root.destroy()

    def test_stores_text_command_and_initial_state(self) -> None:
        clicked = []
        btn = branding.RoundedButton(
            self.root, text="Click me", command=lambda: clicked.append(True),
        )
        self.assertEqual(btn.cget("text"), "Click me")
        self.assertEqual(btn.cget("state"), "normal")
        self.assertIsNotNone(btn.cget("command"))

    def test_unknown_style_raises(self) -> None:
        with self.assertRaises(ValueError):
            branding.RoundedButton(self.root, text="X", style="frankenstein")

    def test_inherits_from_tk_canvas(self) -> None:
        """Heredar de Canvas le da geometry managers (pack/grid/place),
        bindings y rendering. tk.Widget como tipo base permite que los
        type hints `tk.Widget` en handlers acepten tanto tk.Button (tests)
        como RoundedButton (producción)."""
        btn = branding.RoundedButton(self.root, text="X")
        self.assertIsInstance(btn, tk.Canvas)
        self.assertIsInstance(btn, tk.Widget)


class RoundedButtonStylesTest(unittest.TestCase):
    """Los estilos `primary` y `tertiary` deben producir colores
    distinguibles. No verificamos hex exactos (pueden ajustarse en
    branding) sólo que no sean iguales — un test "primary != tertiary"
    detecta si alguien dejó los dos estilos con la misma paleta por
    accidente."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()

    def tearDown(self) -> None:
        self.root.destroy()

    def test_primary_uses_navy_background(self) -> None:
        btn = branding.RoundedButton(self.root, text="X", style="primary")
        bg = self.root.itemcget(btn._bg_item, "fill") if False else \
             btn.itemcget(btn._bg_item, "fill")
        # El fill debe coincidir con la paleta ISA_AZUL del primary.
        self.assertEqual(bg.lower(), branding.ISA_AZUL.lower())

    def test_tertiary_uses_light_background(self) -> None:
        btn = branding.RoundedButton(self.root, text="X", style="tertiary")
        bg = btn.itemcget(btn._bg_item, "fill")
        self.assertEqual(bg.lower(), branding.ISA_FONDO.lower())

    def test_styles_use_distinct_palettes(self) -> None:
        primary = branding.RoundedButton(self.root, text="A", style="primary")
        tertiary = branding.RoundedButton(self.root, text="A", style="tertiary")
        self.assertNotEqual(
            primary.itemcget(primary._bg_item, "fill"),
            tertiary.itemcget(tertiary._bg_item, "fill"),
        )

    def test_naranja_and_verde_styles_use_palette_constants(self) -> None:
        """Los estilos `naranja` y `verde` se usan en las cards del
        menú principal (Control SOX, Reportes). Ambos colores vienen
        de la paleta corporativa del logo (ISA_NARANJA del swoosh,
        ISA_VERDE_OK del estado de éxito) — este test fija el contrato
        de que los styles reflejan las constantes corporativas."""
        naranja = branding.RoundedButton(self.root, text="X", style="naranja")
        verde = branding.RoundedButton(self.root, text="X", style="verde")
        self.assertEqual(
            naranja.itemcget(naranja._bg_item, "fill").lower(),
            branding.ISA_NARANJA.lower(),
        )
        self.assertEqual(
            verde.itemcget(verde._bg_item, "fill").lower(),
            branding.ISA_VERDE_OK.lower(),
        )


class RoundedButtonStateTest(unittest.TestCase):
    """`config(state=...)` y la API `btn["state"]` deben ser consistentes,
    porque el código de polling y de los handlers SAP usa ambas formas
    indistintamente."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()

    def tearDown(self) -> None:
        self.root.destroy()

    def test_disabled_state_via_config(self) -> None:
        btn = branding.RoundedButton(self.root, text="X")
        btn.config(state="disabled")
        self.assertEqual(btn.cget("state"), "disabled")

    def test_getitem_returns_overridden_state(self) -> None:
        """REGRESIÓN: `tk.Misc.__getitem__` es alias directo de `cget`
        (no wrapper que re-dispatcha), así que sin override explícito,
        `btn["state"]` devolvería el valor C-level del Canvas (siempre
        "normal") ignorando nuestro estado custom."""
        btn = branding.RoundedButton(self.root, text="X")
        btn.config(state="disabled")
        self.assertEqual(str(btn["state"]), "disabled")
        btn.config(state="normal")
        self.assertEqual(str(btn["state"]), "normal")

    def test_setitem_routes_through_config(self) -> None:
        """`btn["state"] = "disabled"` debe ser equivalente a
        `btn.config(state="disabled")`."""
        btn = branding.RoundedButton(self.root, text="X")
        btn["state"] = "disabled"
        self.assertEqual(btn.cget("state"), "disabled")


class RoundedButtonInvokeTest(unittest.TestCase):
    """`invoke()` simula click programático — usado por tests del
    proyecto para disparar handlers sin event."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        self.calls = []

    def tearDown(self) -> None:
        self.root.destroy()

    def test_invoke_runs_command(self) -> None:
        btn = branding.RoundedButton(
            self.root, text="X", command=lambda: self.calls.append("ok"),
        )
        btn.invoke()
        self.assertEqual(self.calls, ["ok"])

    def test_invoke_when_disabled_is_noop(self) -> None:
        btn = branding.RoundedButton(
            self.root, text="X", command=lambda: self.calls.append("ok"),
        )
        btn.config(state="disabled")
        btn.invoke()
        self.assertEqual(self.calls, [])

    def test_invoke_with_no_command_is_safe(self) -> None:
        btn = branding.RoundedButton(self.root, text="X")
        # No debe levantar AttributeError ni TypeError.
        btn.invoke()


class RoundedButtonReconfigureTest(unittest.TestCase):
    """`config()` debe permitir actualizar text y command después de la
    construcción, igual que `tk.Button.config()`. El código del proyecto
    usa este patrón en `btn.config(command=lambda: handler(btn, ...))`
    para que el command pueda referenciar al botón mismo."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()

    def tearDown(self) -> None:
        self.root.destroy()

    def test_can_update_text_after_construction(self) -> None:
        btn = branding.RoundedButton(self.root, text="Antes")
        btn.config(text="Después")
        self.assertEqual(btn.cget("text"), "Después")

    def test_can_update_command_after_construction(self) -> None:
        marca = []
        btn = branding.RoundedButton(self.root, text="X")
        btn.config(command=lambda: marca.append("v2"))
        btn.invoke()
        self.assertEqual(marca, ["v2"])


if __name__ == "__main__":
    unittest.main()
