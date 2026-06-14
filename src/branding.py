"""Paleta corporativa "Hub de ISA" + helpers para aplicarla a widgets Tk.

Centraliza los colores y el logo para que `main.py` y la vista
`control_sox` mantengan apariencia consistente. La paleta aproxima los
colores del logo (navy + naranja). Si en algún momento se obtienen los
hex EXACTOS de la guía de marca, basta con ajustar las constantes de
abajo y todo el resto se actualiza automáticamente.

El archivo del logo se busca en `resources/logo_hub_isa.png`. Si no
existe, `cargar_logo()` devuelve None y la GUI debe seguir funcionando
mostrando sólo texto + colores (sin imagen).
"""

import tkinter as tk
from tkinter import font as tkfont

from paths import bundled_resource_path

# ---------------------------------------------------------------------------
# Paleta de colores
# ---------------------------------------------------------------------------

# Azul navy del cuerpo del logo y de "el Hub de ISA". Usado para titulares,
# botones primarios y acentos sobrios.
ISA_AZUL = "#1A3A6C"
# Variante más clara del navy, para hover/pressed de botones primarios.
ISA_AZUL_HOVER = "#2B5CA8"

# Naranja vibrante del swoosh que cruza la "H" del logo. Usado con
# moderación como acento (subrayados, separadores, status warnings).
ISA_NARANJA = "#F58220"
# Variante más oscura para hover/pressed si se usa como botón.
ISA_NARANJA_HOVER = "#D96D14"

# Grises de texto secundario.
ISA_GRIS = "#555555"
ISA_GRIS_CLARO = "#888888"

# Fondo y blanco puro.
ISA_BLANCO = "#FFFFFF"
ISA_FONDO = "#FFFFFF"

# Verde de éxito (mensaje de status OK), conservado del esquema anterior
# para no romper la lectura visual de "operación exitosa". Se reusa
# también como color del botón "Reportes" en el menú principal.
ISA_VERDE_OK = "#1a7f37"
ISA_VERDE_HOVER = "#15662C"

# Gris muy claro para bordes de cards (flat 1px en vez de drop-shadow,
# que Tk no soporta nativo).
ISA_GRIS_BORDE = "#E5E5E5"


# ---------------------------------------------------------------------------
# Logo
# ---------------------------------------------------------------------------

# El logo va bundleado dentro del .exe (read-only); en dev se lee de
# `<repo>/resources/`. Resuelto por `paths.bundled_resource_path` que
# detecta `sys._MEIPASS` cuando corremos como ejecutable.
LOGO_PATH = bundled_resource_path("resources/logo_hub_isa.png")
# Altura por defecto del logo en la GUI (en px). El ancho se calcula
# manteniendo el aspect ratio del archivo original.
LOGO_ALTURA_PX = 85


def cargar_logo(altura_px: int = LOGO_ALTURA_PX):
    """Carga el logo desde `LOGO_PATH` escalado a `altura_px` de alto.

    Returns:
        Un `PIL.ImageTk.PhotoImage` listo para `tk.Label(image=...)`,
        o `None` si el archivo no existe, PIL no está instalado, o
        Tk no está disponible (caso típico: Python sin tkinter).

    Loguea a stdout sólo si falla por una razón "rara" (exception al
    cargar/escalar), no para los casos esperados (archivo ausente, PIL
    sin Tk). Esto evita ruido en macOS dev pero da pistas en producción.

    IMPORTANTE: el caller DEBE guardar el PhotoImage devuelto en una
    variable de instancia (p.ej. `root._logo_ref = cargar_logo()`).
    Si la referencia se pierde, Python liberará la imagen por GC y Tk
    mostrará un cuadro vacío. Tk no retiene referencias a PhotoImage
    por sí mismo — es una "trampa" clásica.
    """
    try:
        from PIL import Image, ImageTk
    except ImportError:
        # Pillow no instalado (raro: está en requirements.txt) o tkinter
        # no disponible (Python sin Tk; afecta a ImageTk pero no a otras
        # partes del sistema). No es un error que ameriten log — es una
        # condición esperada en algunos entornos de dev/CI.
        return None

    if not LOGO_PATH.exists():
        return None

    try:
        img = Image.open(LOGO_PATH)
        ratio = altura_px / img.height
        nuevo_ancho = int(img.width * ratio)
        img = img.resize((nuevo_ancho, altura_px), Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception as exc:
        # Algo raro: el archivo existe, PIL+ImageTk están, pero algo
        # rompe al cargar/escalar. Loguear para diagnóstico.
        print(
            f"branding.cargar_logo: error cargando "
            f"{LOGO_PATH.name}: {type(exc).__name__}: {exc}",
            flush=True,
        )
        return None


# ---------------------------------------------------------------------------
# RoundedButton — botón con esquinas redondeadas (border-radius)
# ---------------------------------------------------------------------------
#
# Tk no soporta `border-radius` nativo en `tk.Button`. Esta clase pinta
# el botón sobre un `tk.Canvas`: un polígono `smooth=True` con anchor
# points repetidos en cada esquina (la curva sólo se aplica a las
# esquinas, los lados quedan rectos) + texto centrado. Los eventos
# `<Button-1>`/`<ButtonRelease-1>`/`<Enter>`/`<Leave>` simulan el
# comportamiento de un `tk.Button` (hover, click, disabled).
#
# API espejada de `tk.Button` — solo lo que el proyecto realmente usa:
#   - constructor: `text`, `command`, `style`, `radius`, `padx`, `pady`,
#     `font`, `width`, `height`, `wraplength`
#   - `config(state="normal"|"disabled", text=..., command=...)`
#   - `cget("text"|"state"|"command")`
#   - `invoke()`  (simula click programático — usado en tests)
#
# Estilos predefinidos vía el parámetro `style`:
#   - "primary"   — fondo navy + texto blanco bold (Extraer, Generar,
#                   Subir, los 3 cards del menú principal, etc.)
#   - "tertiary"  — fondo blanco + texto gris (← Atrás, Test conexión,
#                   Seleccionar archivos, Quitar seleccionado)

_BUTTON_STYLES = {
    "primary": dict(
        bg_normal=ISA_AZUL,
        bg_hover=ISA_AZUL_HOVER,
        bg_pressed=ISA_AZUL_HOVER,
        bg_disabled="#B8C3D4",
        fg_normal=ISA_BLANCO,
        fg_disabled=ISA_BLANCO,
        font=("Helvetica", 11, "bold"),
        padx=14,
        pady=8,
    ),
    "tertiary": dict(
        bg_normal=ISA_FONDO,
        bg_hover="#F0F0F0",
        bg_pressed="#E0E0E0",
        bg_disabled=ISA_FONDO,
        fg_normal=ISA_GRIS,
        fg_disabled=ISA_GRIS_CLARO,
        font=("Helvetica", 9),
        padx=10,
        pady=4,
    ),
    "naranja": dict(
        # Naranja del swoosh corporativo del logo ISA.
        bg_normal=ISA_NARANJA,
        bg_hover=ISA_NARANJA_HOVER,
        bg_pressed=ISA_NARANJA_HOVER,
        bg_disabled="#F8C99A",
        fg_normal=ISA_BLANCO,
        fg_disabled=ISA_BLANCO,
        font=("Helvetica", 11, "bold"),
        padx=14,
        pady=8,
    ),
    "verde": dict(
        # Verde del estado "operación exitosa" reusado como acento de
        # tercer nivel (Reportes) dentro de la paleta corporativa.
        bg_normal=ISA_VERDE_OK,
        bg_hover=ISA_VERDE_HOVER,
        bg_pressed=ISA_VERDE_HOVER,
        bg_disabled="#A8C9B0",
        fg_normal=ISA_BLANCO,
        fg_disabled=ISA_BLANCO,
        font=("Helvetica", 11, "bold"),
        padx=14,
        pady=8,
    ),
}


class RoundedButton(tk.Canvas):
    """`tk.Button` con esquinas redondeadas (default radius=4 px).

    Ver el comment-block superior para la motivación y la API
    soportada. Cualquier parámetro fuera de la lista documentada (ej.
    `relief`, `bd`, `activebackground`) NO aplica: el botón se pinta
    completamente con primitives de Canvas.
    """

    def __init__(
        self,
        parent,
        *,
        text: str = "",
        command=None,
        style: str = "primary",
        radius: int = 4,
        padx: int | None = None,
        pady: int | None = None,
        font=None,
        width: int | None = None,
        height: int | None = None,
        wraplength: int | None = None,
        **canvas_kwargs,
    ) -> None:
        if style not in _BUTTON_STYLES:
            raise ValueError(
                f"style debe ser uno de {list(_BUTTON_STYLES)!r}, recibí: {style!r}"
            )
        cfg = dict(_BUTTON_STYLES[style])
        if padx is not None:
            cfg["padx"] = padx
        if pady is not None:
            cfg["pady"] = pady
        if font is not None:
            cfg["font"] = font

        self._cfg = cfg
        self._text = text
        self._command = command
        self._radius = radius
        self._wraplength = wraplength
        self._state = "normal"

        # El bg del Canvas se hereda del parent para que las esquinas
        # redondeadas se mezclen sin marco visible. Si el caller pasa
        # `bg=`, lo respetamos.
        parent_bg = (
            canvas_kwargs.pop("bg")
            if "bg" in canvas_kwargs
            else parent.cget("bg")
        )
        super().__init__(
            parent,
            highlightthickness=0,
            bd=0,
            bg=parent_bg,
            **canvas_kwargs,
        )

        self._compute_size_and_draw(width=width, height=height)

        self.bind("<Button-1>", self._on_click)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    # ---- dibujo --------------------------------------------------------

    def _compute_size_and_draw(self, *, width=None, height=None) -> None:
        # Medimos el texto con tkfont.Font directamente — no requiere
        # que el Canvas haya sido layouteado (a diferencia de crear un
        # text item temporal y leer su bbox, que devuelve (0,0,0,0)
        # antes del primer update_idletasks).
        fnt = tkfont.Font(font=self._cfg["font"])
        text_w = (
            self._wraplength if self._wraplength
            else fnt.measure(self._text)
        )
        text_h = fnt.metrics("linespace")
        # Si hay wraplength, el alto puede crecer por wrapping. Estimamos
        # por longitud del texto vs ancho disponible.
        if self._wraplength:
            total_w = fnt.measure(self._text)
            lineas = max(1, -(-total_w // self._wraplength))  # ceil
            text_h *= lineas

        w = width if width is not None else int(text_w + 2 * self._cfg["padx"])
        h = height if height is not None else int(text_h + 2 * self._cfg["pady"])
        self.config(width=w, height=h)

        # Rectángulo redondeado: polígono smooth con puntos repetidos en
        # los lados rectos para que el spline curve SOLO las esquinas.
        self._bg_item = self._create_rounded_rect(
            1, 1, w - 1, h - 1, self._radius,
            fill=self._cfg["bg_normal"],
            outline="",
        )
        text_kwargs = dict(
            text=self._text,
            font=self._cfg["font"],
            fill=self._cfg["fg_normal"],
        )
        if self._wraplength:
            text_kwargs["width"] = self._wraplength
        self._text_item = self.create_text(w / 2, h / 2, **text_kwargs)

    def _create_rounded_rect(self, x1, y1, x2, y2, r, **kwargs):
        # Truco: con `smooth=True` Tk pasa un B-spline por TODOS los
        # puntos. Para que sólo se curven las esquinas, repetimos cada
        # punto de los lados rectos — el spline "se asienta" en esos
        # puntos repetidos y conserva la línea recta.
        points = [
            x1 + r, y1, x1 + r, y1,
            x2 - r, y1, x2 - r, y1,
            x2, y1,
            x2, y1 + r, x2, y1 + r,
            x2, y2 - r, x2, y2 - r,
            x2, y2,
            x2 - r, y2, x2 - r, y2,
            x1 + r, y2, x1 + r, y2,
            x1, y2,
            x1, y2 - r, x1, y2 - r,
            x1, y1 + r, x1, y1 + r,
            x1, y1,
        ]
        return self.create_polygon(points, smooth=True, **kwargs)

    # ---- eventos -------------------------------------------------------

    def _on_enter(self, _event) -> None:
        if self._state == "disabled":
            return
        self.itemconfig(self._bg_item, fill=self._cfg["bg_hover"])
        self.config(cursor="hand2")

    def _on_leave(self, _event) -> None:
        if self._state == "disabled":
            return
        self.itemconfig(self._bg_item, fill=self._cfg["bg_normal"])
        self.config(cursor="")

    def _on_click(self, _event) -> None:
        if self._state == "disabled":
            return
        self.itemconfig(self._bg_item, fill=self._cfg["bg_pressed"])

    def _on_release(self, _event) -> None:
        if self._state == "disabled":
            return
        self.itemconfig(self._bg_item, fill=self._cfg["bg_hover"])
        if self._command is not None:
            self._command()

    # ---- API tipo tk.Button --------------------------------------------

    def configure(self, **kwargs):  # alias estándar de Tk
        return self.config(**kwargs)

    def config(self, **kwargs):  # type: ignore[override]
        if "state" in kwargs:
            new_state = str(kwargs.pop("state"))
            self._state = new_state
            if new_state == "disabled":
                self.itemconfig(self._bg_item, fill=self._cfg["bg_disabled"])
                self.itemconfig(self._text_item, fill=self._cfg["fg_disabled"])
                self.config(cursor="")
            else:
                self.itemconfig(self._bg_item, fill=self._cfg["bg_normal"])
                self.itemconfig(self._text_item, fill=self._cfg["fg_normal"])
        if "text" in kwargs:
            self._text = kwargs.pop("text")
            self.itemconfig(self._text_item, text=self._text)
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if kwargs:
            return super().config(**kwargs)
        return None

    def cget(self, option):  # type: ignore[override]
        if option == "text":
            return self._text
        if option == "state":
            return self._state
        if option == "command":
            return self._command
        return super().cget(option)

    def __getitem__(self, key):
        # tk.Misc define `__getitem__ = cget` como alias directo (no
        # como wrapper que re-dispatcha), así que un override sólo de
        # `cget()` NO sería visible vía `btn["state"]`. Redirigimos
        # explícitamente para que ambas vías devuelvan el mismo valor.
        return self.cget(key)

    def __setitem__(self, key, value):
        self.config(**{key: value})

    def invoke(self) -> None:
        """Simula click programático (espejo de `tk.Button.invoke`).
        Usado por tests para disparar el `command` sin Event. No hace
        nada si el botón está disabled."""
        if self._state != "disabled" and self._command is not None:
            self._command()
