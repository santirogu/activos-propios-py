---
name: agregar-vista-gui
description: Scaffolds una nueva sub-vista en la GUI Tkinter del proyecto siguiendo el patrón establecido (frame switching, _crear_header_form, handler con worker, branding ISA, tests). Útil cuando se agrega una opción nueva al menú.
---

# agregar-vista-gui

Cuando el usuario invoque esta skill (ej. "agrega una vista para X"), genera una nueva sub-vista en `src/main.py` siguiendo EXACTAMENTE el patrón de las vistas existentes: `abrir_activos_fijos`, `abrir_extraer_creados`, `abrir_subir_anexos`, `abrir_sox_menu`, `control_sox`.

## Contexto del proyecto (no preguntar)

- **Stack**: Tkinter clásico con `tk.Frame` + `pack()`, branding centralizado en `src/branding.py`.
- **Paleta**: `ISA_AZUL = "#1A3A6C"` (primario), `ISA_NARANJA = "#F58220"` (acento), `ISA_FONDO = "#FFFFFF"`.
- **Navegación**: por frames anidados, no Toplevels. Cada vista hace `parent.pack_forget()` para ocultar el padre y se pack-ea ella misma sobre `root`. Botón "← Atrás" destruye la actual y re-pack-ea el padre.
- **Helper estándar**: `_crear_header_form(root, frame, parent_frame, titulo, subtitulo=None)` construye el botón Atrás + logo + título + subtítulo y devuelve el botón Atrás (útil para deshabilitarlo durante workers SAP).
- **Estilos de botón**: `branding.aplicar_estilo_primario(btn)` para acciones principales (navy bg, blanco fg, bold), `branding.aplicar_estilo_terciario(btn)` para Atrás/Test (gris discreto).
- **Threading**: cualquier acción SAP corre en `threading.Thread(target=worker, daemon=True)` envuelta en `_sap_com_apartment()`. Comunicación thread-safe vía `root.after(0, ...)`.

## Pasos

### 1. Hacer las preguntas correctas al usuario

Si los detalles no están claros, pregunta con `AskUserQuestion`:

- **Nombre y label del botón** que abrirá la vista (ej. "Subir Anexos").
- **Dónde colocar el botón** en el padre: cards del menú principal, dentro de Activos Fijos, Control SOX intermedio, etc.
- **Campos del formulario**: Combobox (qué opciones), Entry (qué label), DateEntry, listbox, file picker, etc.
- **Acción principal**: ¿llama a un módulo SAP existente? ¿uno nuevo? ¿no hace nada todavía (placeholder)?
- **Si hay worker SAP**: ¿debería deshabilitar el botón Atrás durante la ejecución? (la convención del proyecto dice que SÍ — evita que el usuario navegue a media transacción SAP).

### 2. Modificar el padre para añadir el botón

Busca la función padre en `src/main.py` (ej. `abrir_activos_fijos`, `main`, etc.) y añade el botón nuevo. Ejemplo del patrón:

```python
btn_nueva_opcion = tk.Button(
    frame_padre,
    text="<Texto del botón>",
    padx=18, pady=8, width=24,
    command=lambda: abrir_nueva_vista(root, frame_padre),
)
branding.aplicar_estilo_primario(btn_nueva_opcion)
btn_nueva_opcion.pack(pady=(0, 8))
```

Y exponerlo como atributo: `frame_padre.btn_nueva_opcion = btn_nueva_opcion`.

### 3. Implementar la función de la nueva vista

Sigue ESTE esqueleto exacto:

```python
def abrir_<nombre>(root: tk.Tk, frame_padre: tk.Frame) -> tk.Frame:
    """Sub-formulario "<Título>" — accesible desde <Padre>.

    <Descripción breve de qué hace>.

    Form:
      - <Lista de campos>
      - Botón [<Acción>] (cableado a <handler>)
    """
    # Imports lazy si la vista usa cosas que no están al tope del archivo
    # (ej. VALID_SOCIEDADES de sox_report)

    frame_padre.pack_forget()

    frame_<nombre> = tk.Frame(root, bg=branding.ISA_FONDO)
    btn_atras = _crear_header_form(
        root, frame_<nombre>, frame_padre,
        titulo="<Título>",
        subtitulo="<Subtítulo opcional o None>",
    )

    # --- Form ---
    form = tk.Frame(frame_<nombre>, bg=branding.ISA_FONDO)
    form.pack(pady=(8, 8))

    # Para CADA campo del form:
    #   1. tk.Label con `bg=branding.ISA_FONDO`, `fg=branding.ISA_AZUL`, `font=("Helvetica", 10)`
    #   2. Widget del campo: ttk.Combobox state="readonly", tk.Entry, tkcalendar.DateEntry, etc.
    #   3. Grid layout con padx=4, pady=8

    # Ejemplo Combobox:
    tk.Label(
        form, text="<Label>:", anchor="e", width=10,
        bg=branding.ISA_FONDO, fg=branding.ISA_AZUL,
        font=("Helvetica", 10),
    ).grid(row=0, column=0, padx=4, pady=8, sticky="e")

    valor_var = tk.StringVar()
    combo = ttk.Combobox(
        form, textvariable=valor_var,
        values=list(<OPCIONES>), state="readonly", width=14,
    )
    combo.grid(row=0, column=1, padx=4, pady=8, sticky="w")

    # Ejemplo Entry:
    tk.Label(
        form, text="<Label>:", anchor="e", width=12,
        bg=branding.ISA_FONDO, fg=branding.ISA_AZUL,
        font=("Helvetica", 10),
    ).grid(row=1, column=0, padx=4, pady=8, sticky="e")

    texto_var = tk.StringVar()
    entry = tk.Entry(
        form, textvariable=texto_var, width=22,
        font=("Helvetica", 11),
    )
    entry.grid(row=1, column=1, padx=4, pady=8, sticky="w")

    # --- Status (si aplica) ---
    status_var = tk.StringVar(value="")

    # --- Botón de acción principal ---
    btn_accion = tk.Button(
        frame_<nombre>,
        text="<Texto del botón>",
        padx=18, pady=8, width=22,
    )
    btn_accion.config(
        command=lambda: _<nombre>_handler(
            root,
            <args del form>,
            status_var,  # si tiene status
            btn_accion,
            btn_atras,
        )
    )
    branding.aplicar_estilo_primario(btn_accion)
    btn_accion.pack(pady=(14, 6))

    # --- Label de status (si aplica) ---
    # IMPORTANTE: usar font ("Helvetica", 10, "bold") y pady=(10, 20) para
    # que NO quede recortado en el borde inferior de la ventana.
    tk.Label(
        frame_<nombre>,
        textvariable=status_var,
        font=("Helvetica", 10, "bold"),
        fg=branding.ISA_VERDE_OK,
        bg=branding.ISA_FONDO,
        wraplength=560,
    ).pack(pady=(10, 20))

    frame_<nombre>.pack(fill="both", expand=True)

    # --- Exponer atributos para tests ---
    frame_<nombre>.valor_var = valor_var
    frame_<nombre>.texto_var = texto_var
    frame_<nombre>.btn_accion = btn_accion
    frame_<nombre>.btn_atras = btn_atras
    # ... según lo que tests necesiten inspeccionar

    return frame_<nombre>
```

**Reglas críticas:**
- **TODO widget** (`Frame`, `Label`, etc.) lleva `bg=branding.ISA_FONDO` para mantener fondo consistente en blanco.
- **Labels de texto** llevan `fg=branding.ISA_AZUL` salvo los "ayuda" (gris claro: `branding.ISA_GRIS_CLARO`).
- **Status labels** llevan `fg=branding.ISA_VERDE_OK` (success), `font=("Helvetica", 10, "bold")`, y `pady=(10, 20)` para no quedar pegados al borde.
- **Ancho del frame** estándar: la ventana es `620x580`. Los wraplengths suelen ser 560 (deja margen).

### 4. Implementar el handler (si la vista tiene worker SAP)

Si la acción dispara SAP, sigue ESTE patrón (parallel a `_subir_anexos_handler`):

```python
def _<nombre>_handler(
    root: tk.Tk,
    <inputs>,
    status_var: tk.StringVar,
    button: tk.Button,
    btn_atras: tk.Button,
) -> None:
    """<Descripción>."""
    try:
        from <modulo_sap> import validar_<input>
    except ImportError as exc:
        messagebox.showerror("Error de import", str(exc))
        return

    # Validaciones previas con try/except ValueError
    try:
        normalizado = validar_<input>(<input>)
    except ValueError as exc:
        messagebox.showerror("Datos inválidos", str(exc))
        return

    # Confirmación con messagebox.askyesno
    if not messagebox.askyesno(
        "Confirmar <acción>",
        f"<Mensaje con detalles>\n\n¿Continuar?",
    ):
        return

    # Deshabilitar botones
    button.config(state="disabled")
    btn_atras.config(state="disabled")

    # Helpers thread-safe
    def update_status(text: str) -> None:
        root.after(0, status_var.set, text)

    def show_info(title: str, message: str) -> None:
        root.after(0, lambda: messagebox.showinfo(title, message))

    def show_error(title: str, message: str) -> None:
        root.after(0, lambda: messagebox.showerror(title, message))

    def reenable() -> None:
        def _do():
            button.config(state="normal")
            btn_atras.config(state="normal")
        root.after(0, _do)

    def worker() -> None:
        with _sap_com_apartment():
            try:
                try:
                    from <modulo_sap> import get_sap_session, <funcion_flujo>
                except ImportError as exc:
                    show_error("Error de import", str(exc))
                    return
                try:
                    update_status("Conectando a SAP...")
                    session = get_sap_session()
                    # llamar al flujo
                    resultado = <funcion_flujo>(session, ...)
                    # Mostrar éxito
                    show_info("Operación completada", f"<resultado>")
                except Exception as exc:
                    show_error("Error", str(exc))
            finally:
                reenable()

    threading.Thread(target=worker, daemon=True).start()
```

### 5. Generar tests en `tests/test_main.py`

Añade dos test classes al final del archivo (antes del `if __name__ == "__main__"`):

```python
class Abrir<Nombre>Test(unittest.TestCase):
    """`abrir_<nombre>(root, frame_padre)` muestra el form y expone widgets."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        self.frame_padre = tk.Frame(self.root)
        self.frame_padre.pack(fill="both", expand=True)

    def tearDown(self) -> None:
        self.root.destroy()

    def test_hides_padre_when_invoked(self):
        frame = main.abrir_<nombre>(self.root, self.frame_padre)
        try:
            self.assertEqual(self.frame_padre.winfo_manager(), "")
        finally:
            frame.destroy()

    def test_exposes_form_widgets(self):
        frame = main.abrir_<nombre>(self.root, self.frame_padre)
        try:
            self.assertIsInstance(frame.valor_var, tk.StringVar)
            self.assertIsInstance(frame.btn_accion, tk.Button)
        finally:
            frame.destroy()

    def test_back_button_returns_to_padre(self):
        frame = main.abrir_<nombre>(self.root, self.frame_padre)
        frame.btn_atras.invoke()
        self.assertFalse(frame.winfo_exists())
        self.assertEqual(self.frame_padre.winfo_manager(), "pack")


class <Nombre>HandlerTest(unittest.TestCase):
    """Pruebas del handler `_<nombre>_handler`."""

    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.after = lambda delay, fn, *args: fn(*args)
        self.status_var = tk.StringVar(master=self.root)
        self.button = tk.Button(self.root)
        self.btn_atras = tk.Button(self.root)

    def tearDown(self) -> None:
        self.root.destroy()

    def test_shows_error_on_invalid_input(self):
        with patch("main.messagebox.showerror") as mock_err, \
             patch("main.messagebox.askyesno"):
            main._<nombre>_handler(
                self.root, "<input inválido>",
                self.status_var, self.button, self.btn_atras,
            )
        mock_err.assert_called_once()

    def test_cancel_confirmation_does_not_start_worker(self):
        with patch("main.messagebox.askyesno", return_value=False), \
             patch("main.threading.Thread") as mock_thread:
            main._<nombre>_handler(
                self.root, "<input válido>",
                self.status_var, self.button, self.btn_atras,
            )
        mock_thread.assert_not_called()

    def test_happy_path_calls_sap_module(self):
        with patch("main.messagebox.askyesno", return_value=True), \
             patch("main.messagebox.showinfo"), \
             patch("main.threading.Thread", _SyncFakeThread), \
             patch("<modulo_sap>.get_sap_session", return_value=MagicMock()), \
             patch("<modulo_sap>.<funcion_flujo>") as mock_flow:
            main._<nombre>_handler(
                self.root, "<input válido>",
                self.status_var, self.button, self.btn_atras,
            )
        mock_flow.assert_called_once()
```

### 6. Sugerir siguientes pasos

- Invoca `correr-tests` para verificar que todo compila y los tests del handler pasan.
- Invoca `sincronizar-docs` para actualizar CLAUDE.md (tabla de sub-vistas) y README.md.

## Convenciones aprendidas (no romperlas)

1. **NUNCA cachees el padre `frame_padre` en un Toplevel separado** — el proyecto usa frame switching, no ventanas modales.
2. **El botón Atrás SIEMPRE viene de `_crear_header_form`** (devuelve el btn_atras) — no lo construyas manual.
3. **Logo NUNCA recargues** — usa `root._logo_ref` que ya está cargado.
4. **Threading**: el worker SIEMPRE va envuelto en `_sap_com_apartment()`. Comunicación con la GUI SIEMPRE vía `root.after(0, ...)`.
5. **Botón Atrás se deshabilita durante el worker SAP** (parte de la convención del proyecto — evita navegación a media transacción).
6. **macOS quirks**: `tk.Button` ignora `bg`/`activebackground` en aqua. La app es Windows-only para SAP así que no nos preocupa, pero si testeas en macOS los botones se ven nativos en vez de coloreados.
