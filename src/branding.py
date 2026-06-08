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

from pathlib import Path

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
# para no romper la lectura visual de "operación exitosa".
ISA_VERDE_OK = "#1a7f37"


# ---------------------------------------------------------------------------
# Logo
# ---------------------------------------------------------------------------

PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOGO_PATH = PROJECT_ROOT / "resources" / "logo_hub_isa.png"
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
# Estilos de botón (aplicar después de crear el tk.Button)
# ---------------------------------------------------------------------------
#
# Nota macOS: `tk.Button` ignora `bg`/`activebackground` en macOS aqua
# por default (sólo respeta `fg`). En Windows los respeta normalmente.
# Como la app es Windows-only para los flujos SAP, no rompe nada — sólo
# se ve más "nativo" en macOS en lugar de coloreado.

def aplicar_estilo_primario(btn) -> None:
    """Botón de acción principal (Extraer, Subir SAP, Control SOX,
    Generar Reporte SOX). Fondo navy, texto blanco."""
    btn.config(
        bg=ISA_AZUL,
        fg=ISA_BLANCO,
        activebackground=ISA_AZUL_HOVER,
        activeforeground=ISA_BLANCO,
        relief="flat",
        bd=0,
        cursor="hand2",
        font=("Helvetica", 11, "bold"),
    )


def aplicar_estilo_terciario(btn) -> None:
    """Botón secundario/diagnóstico (Test conexión, ← Atrás). Estilo
    discreto en gris para no competir con los botones principales."""
    btn.config(
        bg=ISA_FONDO,
        fg=ISA_GRIS,
        activebackground="#F0F0F0",
        activeforeground=ISA_AZUL,
        relief="flat",
        bd=1,
        cursor="hand2",
    )
