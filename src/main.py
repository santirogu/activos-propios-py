import contextlib
import threading
import time
import traceback
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path
from datetime import datetime

import openpyxl
from tkcalendar import DateEntry

import branding

PROJECT_ROOT = Path(__file__).resolve().parent.parent
EXCEL_PATH = PROJECT_ROOT / "resources" / "Formato_Dinamico_.xlsx"
OUTPUT_DIR = PROJECT_ROOT / "salida"
SHEET_NAME = "LSMW "

# Intervalo (ms) con que se re-evalúa el estado del botón "Subir a SAP" para
# habilitarlo/deshabilitarlo según existan o no .txt en salida/.
_POLL_INTERVAL_MS = 1000

# Flag módulo-level: True mientras un worker de subir_a_sap está corriendo.
# Sirve para que el polling NO toque el estado del botón durante la carga
# (el worker tiene control exclusivo en ese momento).
_upload_en_curso = False


def _log(mensaje: str) -> None:
    """Imprime un mensaje con timestamp [HH:MM:SS] y flush=True para que
    aparezca en tiempo real al ejecutar `python src/main.py` desde terminal."""
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {mensaje}", flush=True)


def _show_unexpected_error(title: str, exc: BaseException) -> None:
    """Loguea la excepción completa y muestra un diálogo con el detalle.

    Sirve como red de seguridad para excepciones que ningún `except`
    específico capturó: el usuario verá un error en pantalla en vez de
    quedarse sin retroalimentación.
    """
    tb_text = "".join(
        traceback.format_exception(type(exc), exc, exc.__traceback__)
    )
    _log(f"ERROR — {title}: {exc}")
    print(tb_text, flush=True)
    messagebox.showerror(
        title,
        f"{type(exc).__name__}: {exc}\n\n--- Detalle técnico ---\n{tb_text}",
    )


@contextlib.contextmanager
def _sap_com_apartment():
    """Inicializa el apartamento COM del thread actual y lo libera al salir.

    Windows exige `pythoncom.CoInitialize()` antes de cualquier llamada COM
    desde un thread que no sea el main de la app — sin esto, `GetObject('SAPGUI')`
    en los workers de Subir a SAP / Generar Reporte SOX falla con un error
    genérico ("No se pudo conectar a SAP GUI") aunque SAP esté abierto.

    Loguea cada paso (CoInitialize / CoUninitialize) para diagnosticar
    problemas de conexión desde threads.

    No-op en sistemas sin pythoncom (Mac/Linux).
    """
    thread_id = threading.get_ident()
    try:
        import pythoncom  # type: ignore
    except ImportError:
        _log(f"_sap_com_apartment: pythoncom no disponible (thread={thread_id}, no-op)")
        yield
        return

    _log(f"_sap_com_apartment: llamando CoInitialize() en thread={thread_id}...")
    try:
        pythoncom.CoInitialize()
        _log(f"_sap_com_apartment: CoInitialize OK (thread={thread_id})")
    except Exception as exc:
        _log(f"_sap_com_apartment: CoInitialize FALLÓ — {exc!r}")
        raise

    try:
        yield
    finally:
        try:
            pythoncom.CoUninitialize()
            _log(f"_sap_com_apartment: CoUninitialize OK (thread={thread_id})")
        except Exception as exc:
            _log(f"_sap_com_apartment: CoUninitialize falló (ignorado) — {exc!r}")


def _install_tk_exception_handler(root: tk.Tk) -> None:
    """Reemplaza el handler default de Tkinter (que solo imprime a stderr)
    por uno que muestra un diálogo. Captura cualquier excepción no manejada
    en callbacks Tkinter — sin esto, los errores son invisibles para el
    usuario que abre la app por doble-clic."""

    def handler(exc_type, exc_value, tb) -> None:
        tb_text = "".join(traceback.format_exception(exc_type, exc_value, tb))
        _log(f"ERROR no manejado en callback Tkinter: {exc_value}")
        print(tb_text, flush=True)
        messagebox.showerror(
            "Error inesperado",
            f"{exc_type.__name__}: {exc_value}\n\n"
            f"--- Detalle técnico ---\n{tb_text}",
        )

    root.report_callback_exception = handler


def _hay_txt_en_salida() -> bool:
    """True si hay al menos un archivo LSMW_*.txt en salida/."""
    return OUTPUT_DIR.exists() and any(OUTPUT_DIR.glob("LSMW_*.txt"))


def _refrescar_estado_boton_subir(button: tk.Button) -> None:
    """Sincroniza el estado del botón con la presencia de .txt en salida/.

    Si hay un upload en curso, no toca el botón (el worker lo controla).
    """
    if _upload_en_curso:
        return
    button.config(state="normal" if _hay_txt_en_salida() else "disabled")


def export_sheet_to_tsv(
    excel_path: Path,
    sheet_name: str,
    output_dir: Path,
    file_prefix: str = "LSMW",
) -> tuple[Path, int]:
    """Lee `sheet_name` del workbook en `excel_path` y escribe un .txt
    separado por tabulación dentro de `output_dir`. Devuelve (ruta, filas)."""
    if not excel_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"La hoja '{sheet_name.strip()}' no existe en el archivo.")

    ws = wb[sheet_name]
    output_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"{file_prefix}_{ts}.txt"

    rows_written = 0
    with output_path.open("w", encoding="utf-8", newline="") as f:
        for row in ws.iter_rows(values_only=True):
            cells = ["" if v is None else str(v) for v in row]
            f.write("\t".join(cells) + "\n")
            rows_written += 1

    return output_path, rows_written


def extraer_lsmw_a_txt(status_var: tk.StringVar) -> None:
    try:
        _log("Botón 'Extraer información en txt' presionado")
        _log(f"OUTPUT_DIR = {OUTPUT_DIR}")
        _log(f"EXCEL_PATH = {EXCEL_PATH}")

        # Si ya existe(n) .txt previo(s) en salida/, pedir confirmación antes
        # de reemplazar.
        existentes = (
            sorted(OUTPUT_DIR.glob("LSMW_*.txt")) if OUTPUT_DIR.exists() else []
        )
        _log(f"Archivos LSMW_*.txt previos en salida/: {len(existentes)}")
        if existentes:
            reemplazar = messagebox.askyesno(
                "Archivo ya existente",
                f"Ya existe un .txt generado en salida/:\n"
                f"  {existentes[-1].name}\n\n"
                f"¿Deseas reemplazarlo por uno nuevo?",
            )
            if not reemplazar:
                _log("Usuario canceló el reemplazo. Conservando archivo existente.")
                status_var.set(
                    "Operación cancelada. Se conservó el archivo existente."
                )
                return
            for old in existentes:
                try:
                    old.unlink()
                    _log(f"Archivo borrado: {old.name}")
                except OSError as exc:
                    _log(f"Error al borrar {old.name}: {exc}")
                    messagebox.showerror(
                        "Error al borrar archivo",
                        f"No se pudo borrar {old.name}:\n{exc}",
                    )
                    return

        _log("Generando nuevo .txt desde la hoja LSMW...")
        try:
            output_path, rows_written = export_sheet_to_tsv(
                EXCEL_PATH, SHEET_NAME, OUTPUT_DIR
            )
        except FileNotFoundError as exc:
            _log(f"FileNotFoundError: {exc}")
            messagebox.showerror("Archivo no encontrado", str(exc))
            return
        except ValueError as exc:
            _log(f"ValueError: {exc}")
            messagebox.showerror("Hoja no encontrada", str(exc))
            return
        except Exception as exc:
            _log(f"Excepción durante export_sheet_to_tsv: {exc}")
            messagebox.showerror("Error al exportar", str(exc))
            return

        _log(f"Generado: {output_path.name} ({rows_written} filas)")
        status_var.set(f"Exportado: {output_path.name} ({rows_written} filas)")
        messagebox.showinfo(
            "Extracción completa",
            f"Se generó el archivo:\n{output_path}\n\nFilas exportadas: {rows_written}",
        )
    except Exception as exc:
        # Red de seguridad: cualquier excepción no prevista (acceso a
        # OUTPUT_DIR, errores de Tkinter, etc.) se muestra al usuario con
        # el traceback completo en consola.
        _show_unexpected_error("Error inesperado al extraer", exc)


def subir_a_sap(root: tk.Tk, status_var: tk.StringVar, button: tk.Button) -> None:
    """Lanza la carga LSMW a SAP en un hilo background.

    Confirma con el usuario, deshabilita el botón mientras corre y va
    actualizando `status_var` desde el hilo principal vía `root.after`.
    """
    confirmar = messagebox.askyesno(
        "Confirmar carga a SAP",
        "Esto tomará el .txt más reciente de salida/ y ejecutará el flujo "
        "LSMW en la sesión SAP abierta.\n\n"
        "Asegúrate de:\n"
        "  • Tener SAP abierto y con sesión iniciada.\n"
        "  • Tener el proyecto LSMW pre-cargado.\n"
        "  • No tocar SAP mientras se ejecuta el script.\n\n"
        "¿Continuar?",
    )
    if not confirmar:
        return

    global _upload_en_curso
    _upload_en_curso = True
    button.config(state="disabled")

    def update_status(text: str) -> None:
        root.after(0, status_var.set, text)

    def show_info(title: str, message: str) -> None:
        root.after(0, lambda: messagebox.showinfo(title, message))

    def show_error(title: str, message: str) -> None:
        root.after(0, lambda: messagebox.showerror(title, message))

    def worker() -> None:
        global _upload_en_curso
        with _sap_com_apartment():
            try:
                try:
                    from sap_upload import (
                        get_latest_txt,
                        get_sap_session,
                        run_lsmw_flow,
                    )
                except ImportError as exc:
                    show_error(
                        "Error de import",
                        f"No se pudo importar sap_upload:\n{exc}",
                    )
                    return

                try:
                    update_status("Buscando .txt más reciente en salida/...")
                    latest = get_latest_txt()

                    update_status("Conectando a la sesión SAP...")
                    session = get_sap_session()

                    update_status("Ejecutando flujo LSMW (no toques SAP)...")
                    run_lsmw_flow(session, str(latest.parent), latest.name)

                    update_status(
                        "Carga completada. Revisa SM35 para el log de la BDC."
                    )
                    show_info(
                        "Carga completada",
                        "Flujo LSMW ejecutado correctamente.\n\n"
                        "Revisa SM35 para ver el log de la sesión BDC.",
                    )
                except Exception as exc:
                    update_status("")
                    show_error("Error en carga SAP", str(exc))
            finally:
                _upload_en_curso = False
                root.after(0, lambda: _refrescar_estado_boton_subir(button))

    threading.Thread(target=worker, daemon=True).start()


# ---------------------------------------------------------------------------
# Control SOX — diálogo de generación de reporte
# ---------------------------------------------------------------------------

def _generar_reporte_sox_handler(
    root: tk.Tk,
    sociedad: str,
    fecha_desde: str,
    fecha_hasta: str,
    status_var: tk.StringVar,
    button: tk.Button,
    btn_atras: tk.Button,
) -> None:
    """Valida los inputs y lanza el worker que genera el reporte SOX.

    El form vive en `root` (no en un Toplevel); por eso usamos `root.after`
    para los callbacks thread-safe. Mientras el worker corre, deshabilita
    tanto el botón Generar como el botón Atrás (no queremos que el usuario
    vuelva al menú a mitad de un flujo SAP)."""
    try:
        from sox_report import validar_sociedad, validar_rango_fechas
    except ImportError as exc:
        messagebox.showerror(
            "Error de import", f"No se pudo importar sox_report:\n{exc}"
        )
        return

    try:
        sociedad_norm = validar_sociedad(sociedad)
        validar_rango_fechas(fecha_desde, fecha_hasta)
    except ValueError as exc:
        messagebox.showerror("Datos inválidos", str(exc))
        return

    if not messagebox.askyesno(
        "Confirmar generación del reporte SOX",
        f"Se generará el reporte SOX para:\n"
        f"  • Sociedad: {sociedad_norm}\n"
        f"  • Desde: {fecha_desde}\n"
        f"  • Hasta: {fecha_hasta}\n\n"
        f"El archivo se guardará en salida/.\n\n"
        f"Asegúrate de tener SAP abierto y con sesión iniciada.\n\n"
        f"¿Continuar?",
    ):
        return

    button.config(state="disabled")
    btn_atras.config(state="disabled")

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
                    from sox_report import generar_reporte_sox, get_sap_session
                except ImportError as exc:
                    show_error(
                        "Error de import",
                        f"No se pudo importar sox_report:\n{exc}",
                    )
                    return

                try:
                    update_status("Conectando a la sesión SAP...")
                    session = get_sap_session()

                    update_status(
                        f"Generando reporte SOX para {sociedad_norm} "
                        f"({fecha_desde} → {fecha_hasta})..."
                    )
                    carpeta, nombre = generar_reporte_sox(
                        session, sociedad_norm, fecha_desde, fecha_hasta
                    )

                    update_status(f"Reporte generado: {nombre}")
                    show_info(
                        "Reporte SOX generado",
                        f"Archivo guardado en:\n{carpeta}\\{nombre}",
                    )
                except Exception as exc:
                    update_status("")
                    show_error("Error generando reporte SOX", str(exc))
            finally:
                reenable()

    threading.Thread(target=worker, daemon=True).start()


def _crear_header_form(
    root: tk.Tk,
    frame: tk.Frame,
    parent_frame: tk.Frame,
    titulo: str,
    subtitulo: str | None = None,
) -> tk.Button:
    """Construye el encabezado consistente de cualquier sub-formulario:
    botón "← Atrás" arriba-izquierda + logo centrado + título + subtítulo.

    Devuelve el botón Atrás para que el caller pueda referenciarlo (p. ej.
    para deshabilitarlo durante un worker). El botón ya tiene cableado el
    comando que destruye `frame` y re-empaca `parent_frame`.
    """
    btn_atras = tk.Button(
        frame, text="← Atrás", font=("Helvetica", 9), padx=8, pady=2,
    )
    branding.aplicar_estilo_terciario(btn_atras)
    btn_atras.pack(anchor="w", padx=10, pady=(10, 0))

    def volver() -> None:
        frame.destroy()
        parent_frame.pack(fill="both", expand=True)

    btn_atras.config(command=volver)

    # Logo (reusa la referencia del root para no recargar la imagen).
    if getattr(root, "_logo_ref", None) is not None:
        tk.Label(
            frame, image=root._logo_ref, bg=branding.ISA_FONDO
        ).pack(pady=(4, 6))

    tk.Label(
        frame,
        text=titulo,
        font=("Helvetica", 14, "bold"),
        fg=branding.ISA_AZUL,
        bg=branding.ISA_FONDO,
    ).pack(pady=(4, 4))

    if subtitulo is not None:
        tk.Label(
            frame,
            text=subtitulo,
            font=("Helvetica", 10),
            fg=branding.ISA_GRIS,
            bg=branding.ISA_FONDO,
        ).pack(pady=(0, 12))

    return btn_atras


def abrir_activos_fijos(root: tk.Tk, frame_menu: tk.Frame) -> tk.Frame:
    """Sub-formulario "Activos Fijos" — accesible desde el menú principal.

    Contiene dos botones:
      - "Extraer información en txt": misma lógica que el botón homónimo
        anterior (función `extraer_lsmw_a_txt`).
      - "Creación de Activo": misma lógica que el viejo "Subir a SAP"
        (función `subir_a_sap`), incluyendo el polling que lo habilita
        sólo cuando hay un LSMW_*.txt en salida/.

    El polling se inicia al entrar a la vista y se cancela al salir
    (back button) para no dejar callbacks programados sobre widgets
    destruidos.

    Devuelve el `Frame` creado para que los tests puedan inspeccionar
    los widgets vía atributos.
    """
    frame_menu.pack_forget()

    frame_activos = tk.Frame(root, bg=branding.ISA_FONDO)
    btn_atras = _crear_header_form(
        root, frame_activos, frame_menu,
        titulo="Activos Fijos",
        subtitulo="Extracción de información y creación de activos en SAP",
    )

    status_var = tk.StringVar(value="")

    btn_extraer = tk.Button(
        frame_activos,
        text="Extraer información en txt",
        command=lambda: extraer_lsmw_a_txt(status_var),
        padx=18, pady=8, width=24,
    )
    branding.aplicar_estilo_primario(btn_extraer)
    btn_extraer.pack(pady=(8, 8))

    btn_creacion = tk.Button(
        frame_activos,
        text="Creación de Activo",
        padx=18, pady=8, width=24,
        state="disabled",
    )
    btn_creacion.config(
        command=lambda: subir_a_sap(root, status_var, btn_creacion)
    )
    branding.aplicar_estilo_primario(btn_creacion)
    btn_creacion.pack(pady=(0, 8))

    tk.Label(
        frame_activos,
        textvariable=status_var,
        font=("Helvetica", 9),
        fg=branding.ISA_VERDE_OK,
        bg=branding.ISA_FONDO,
        wraplength=440,
    ).pack(pady=(12, 0))

    # Polling scoped al frame: se inicia ahora, se cancela cuando el
    # frame se destruye (vía el binding <Destroy>). Esto evita callbacks
    # programados sobre un Button ya destruido cuando el usuario regresa
    # al menú.
    polling_id_holder: list[str | None] = [None]

    def poll_creacion_activo() -> None:
        if not btn_creacion.winfo_exists():
            return
        _refrescar_estado_boton_subir(btn_creacion)
        polling_id_holder[0] = root.after(
            _POLL_INTERVAL_MS, poll_creacion_activo
        )

    def on_frame_destroy(_event):
        # Cancelar polling para no dejar callbacks sueltos.
        if polling_id_holder[0] is not None:
            try:
                root.after_cancel(polling_id_holder[0])
            except Exception:
                pass

    frame_activos.bind("<Destroy>", on_frame_destroy)

    # Estado inicial + arrancar polling.
    _refrescar_estado_boton_subir(btn_creacion)
    polling_id_holder[0] = root.after(
        _POLL_INTERVAL_MS, poll_creacion_activo
    )

    frame_activos.pack(fill="both", expand=True)

    # Exponer atributos clave para los tests.
    frame_activos.status_var = status_var
    frame_activos.btn_extraer = btn_extraer
    frame_activos.btn_creacion = btn_creacion
    frame_activos.btn_atras = btn_atras

    return frame_activos


def abrir_sox_menu(root: tk.Tk, frame_menu: tk.Frame) -> tk.Frame:
    """Sub-formulario intermedio "Control SOX" — accesible desde el menú
    principal. Muestra un único botón "HUB.PPE.01 Creación de Activos
    Fijos" que abre el formulario con los parámetros (sociedad + fechas).

    Diseñado como contenedor de futuras opciones HUB.PPE.XX que se vayan
    agregando: hoy es un único botón pero la estructura permite añadir
    más sin mover la lógica del menú principal.
    """
    frame_menu.pack_forget()

    frame_sox_menu = tk.Frame(root, bg=branding.ISA_FONDO)
    btn_atras = _crear_header_form(
        root, frame_sox_menu, frame_menu,
        titulo="Control SOX",
        subtitulo="Procesos de control y auditoría",
    )

    btn_hub_ppe_01 = tk.Button(
        frame_sox_menu,
        text="HUB.PPE.01 Creación de Activos Fijos",
        padx=18, pady=8, width=32,
        command=lambda: control_sox(root, frame_sox_menu),
    )
    branding.aplicar_estilo_primario(btn_hub_ppe_01)
    btn_hub_ppe_01.pack(pady=(12, 8))

    frame_sox_menu.pack(fill="both", expand=True)

    # Exponer atributos clave para los tests.
    frame_sox_menu.btn_hub_ppe_01 = btn_hub_ppe_01
    frame_sox_menu.btn_atras = btn_atras

    return frame_sox_menu


def control_sox(root: tk.Tk, frame_menu: tk.Frame) -> tk.Frame:
    """Reemplaza la vista del menú principal por el formulario Control SOX
    en la misma ventana (sin abrir un Toplevel).

    Oculta `frame_menu` con `pack_forget` y muestra un nuevo `frame_sox` con
    un botón "← Atrás" arriba que destruye el form SOX y re-muestra el
    menú al ser presionado.

    Devuelve el `Frame` creado para que los tests puedan inspeccionar los
    widgets (las StringVars, los DateEntry, los botones).
    """
    from sox_report import VALID_SOCIEDADES, validar_caracter_fecha

    # Ocultar el menú — preservamos su estado (status_var, polling del
    # botón Subir, etc.) para no perder progreso si el usuario vuelve.
    frame_menu.pack_forget()

    frame_sox = tk.Frame(root, bg=branding.ISA_FONDO)

    # --- Botón Atrás (esquina superior izquierda, estilo discreto) ---
    btn_atras = tk.Button(
        frame_sox,
        text="← Atrás",
        font=("Helvetica", 9),
        padx=8,
        pady=2,
    )
    branding.aplicar_estilo_terciario(btn_atras)
    btn_atras.pack(anchor="w", padx=10, pady=(10, 0))

    def volver_al_menu() -> None:
        frame_sox.destroy()
        frame_menu.pack(fill="both", expand=True)

    btn_atras.config(command=volver_al_menu)

    # --- Logo arriba del título (reusa la referencia del root) ---
    if getattr(root, "_logo_ref", None) is not None:
        tk.Label(
            frame_sox, image=root._logo_ref, bg=branding.ISA_FONDO
        ).pack(pady=(4, 6))

    # --- Título ---
    tk.Label(
        frame_sox,
        text="Control SOX",
        font=("Helvetica", 13, "bold"),
        fg=branding.ISA_AZUL,
        bg=branding.ISA_FONDO,
    ).pack(pady=(4, 4))
    tk.Label(
        frame_sox,
        text="Genera el Reporte SOX con los parámetros indicados",
        font=("Helvetica", 10),
        fg=branding.ISA_GRIS,
        bg=branding.ISA_FONDO,
    ).pack(pady=(0, 12))

    form = tk.Frame(frame_sox, bg=branding.ISA_FONDO)
    form.pack(pady=(0, 12))

    # --- Sociedad (Combobox readonly) ---
    tk.Label(
        form, text="Sociedad:", anchor="e", width=10,
        bg=branding.ISA_FONDO, fg=branding.ISA_AZUL,
    ).grid(row=0, column=0, padx=4, pady=6, sticky="e")
    sociedad_var = tk.StringVar()
    sociedad_combo = ttk.Combobox(
        form,
        textvariable=sociedad_var,
        values=list(VALID_SOCIEDADES),
        state="readonly",
        width=14,
    )
    sociedad_combo.grid(row=0, column=1, padx=4, pady=6, sticky="w")

    # --- Fechas con calendario emergente (DateEntry de tkcalendar) ---
    # DateEntry abre un popup de calendario al hacer clic en la flecha. El
    # validatecommand sigue activo: aunque el usuario escriba a mano, solo
    # se aceptan dígitos y puntos (máx 10 caracteres).
    vcmd = (root.register(validar_caracter_fecha), "%P")
    fecha_hoy = datetime.now()

    tk.Label(
        form, text="Desde:", anchor="e", width=10,
        bg=branding.ISA_FONDO, fg=branding.ISA_AZUL,
    ).grid(row=1, column=0, padx=4, pady=6, sticky="e")
    desde_var = tk.StringVar()
    desde_entry = DateEntry(
        form,
        textvariable=desde_var,
        date_pattern="dd.mm.yyyy",
        width=14,
        background="#1a73e8",
        foreground="white",
        borderwidth=2,
        validate="key",
        validatecommand=vcmd,
        year=fecha_hoy.year,
        month=fecha_hoy.month,
        day=fecha_hoy.day,
    )
    desde_entry.grid(row=1, column=1, padx=4, pady=6, sticky="w")
    tk.Label(
        form, text="(dd.mm.aaaa)",
        fg=branding.ISA_GRIS_CLARO, bg=branding.ISA_FONDO,
    ).grid(row=1, column=2, padx=4)

    tk.Label(
        form, text="Hasta:", anchor="e", width=10,
        bg=branding.ISA_FONDO, fg=branding.ISA_AZUL,
    ).grid(row=2, column=0, padx=4, pady=6, sticky="e")
    hasta_var = tk.StringVar()
    hasta_entry = DateEntry(
        form,
        textvariable=hasta_var,
        date_pattern="dd.mm.yyyy",
        width=14,
        background="#1a73e8",
        foreground="white",
        borderwidth=2,
        validate="key",
        validatecommand=vcmd,
        year=fecha_hoy.year,
        month=fecha_hoy.month,
        day=fecha_hoy.day,
    )
    hasta_entry.grid(row=2, column=1, padx=4, pady=6, sticky="w")
    tk.Label(
        form, text="(dd.mm.aaaa)",
        fg=branding.ISA_GRIS_CLARO, bg=branding.ISA_FONDO,
    ).grid(row=2, column=2, padx=4)

    status_var = tk.StringVar()

    btn_generar = tk.Button(
        frame_sox,
        text="Generar Reporte SOX",
        padx=18,
        pady=8,
    )
    btn_generar.config(
        command=lambda: _generar_reporte_sox_handler(
            root,
            sociedad_var.get(),
            desde_var.get(),
            hasta_var.get(),
            status_var,
            btn_generar,
            btn_atras,
        )
    )
    branding.aplicar_estilo_primario(btn_generar)
    btn_generar.pack()

    tk.Label(
        frame_sox,
        textvariable=status_var,
        font=("Helvetica", 9),
        fg=branding.ISA_VERDE_OK,
        bg=branding.ISA_FONDO,
        wraplength=460,
    ).pack(pady=(12, 0))

    frame_sox.pack(fill="both", expand=True)

    # Exponer widgets clave en el frame para que los tests puedan inspeccionar.
    frame_sox.sociedad_var = sociedad_var
    frame_sox.desde_var = desde_var
    frame_sox.hasta_var = hasta_var
    frame_sox.status_var = status_var
    frame_sox.sociedad_combo = sociedad_combo
    frame_sox.desde_entry = desde_entry
    frame_sox.hasta_entry = hasta_entry
    frame_sox.btn_generar = btn_generar
    frame_sox.btn_atras = btn_atras

    return frame_sox


def _test_conexion_sap_handler() -> None:
    """Handler del botón "Test conexión SAP". Llama a
    `diagnosticar_conexion_sap` y muestra el resultado en un messagebox.
    """
    try:
        from sap_upload import diagnosticar_conexion_sap
    except ImportError as exc:
        messagebox.showerror(
            "Error de import",
            f"No se pudo importar sap_upload:\n{exc}",
        )
        return

    try:
        ok, mensaje = diagnosticar_conexion_sap()
    except Exception as exc:
        _show_unexpected_error("Error en test de conexión SAP", exc)
        return

    _log(f"Test conexión SAP → ok={ok}")
    print(mensaje, flush=True)
    if ok:
        messagebox.showinfo("Test conexión SAP — OK", mensaje)
    else:
        messagebox.showwarning("Test conexión SAP — Problema", mensaje)


def main() -> None:
    root = tk.Tk()
    _install_tk_exception_handler(root)
    root.title("Gestión de Activos Fijos")
    # Ventana algo más ancha para acomodar 3 cards horizontales.
    root.geometry("620x480")
    root.resizable(False, False)
    root.configure(bg=branding.ISA_FONDO)

    # Todos los widgets del menú principal viven dentro de `frame_menu`
    # para que las sub-vistas (Activos Fijos, Control SOX) puedan
    # ocultarlo con pack_forget y mostrar sus propios frames en la misma
    # ventana, preservando el estado del menú.
    frame_menu = tk.Frame(root, bg=branding.ISA_FONDO)

    # Logo Hub de ISA arriba. Si el archivo no existe, _logo_ref es None
    # y simplemente no se renderiza la imagen (graceful fallback). La
    # referencia se guarda en `root` para que GC no libere la imagen y
    # también para que `_crear_header_form` la reuse en sub-vistas.
    root._logo_ref = branding.cargar_logo()
    if root._logo_ref is not None:
        tk.Label(
            frame_menu, image=root._logo_ref, bg=branding.ISA_FONDO,
        ).pack(pady=(18, 8))

    tk.Label(
        frame_menu,
        text="Gestión de Activos Fijos",
        font=("Helvetica", 14, "bold"),
        fg=branding.ISA_AZUL,
        bg=branding.ISA_FONDO,
    ).pack(pady=(4, 20))

    # --- 3 cards horizontales (Activos Fijos | Control SOX | Reportes) ---
    # Cada card es un tk.Button con altura/ancho consistentes para que la
    # fila se vea pareja. `Reportes` queda en `state="disabled"` por ahora.
    cards_row = tk.Frame(frame_menu, bg=branding.ISA_FONDO)
    cards_row.pack(pady=(8, 0))

    def _crear_card(parent, texto, command=None, disabled=False) -> tk.Button:
        card = tk.Button(
            parent,
            text=texto,
            padx=12,
            pady=24,
            width=16,
            wraplength=140,
            command=command,
        )
        branding.aplicar_estilo_primario(card)
        if disabled:
            card.config(state="disabled")
        return card

    btn_card_activos = _crear_card(
        cards_row,
        "Activos Fijos",
        command=lambda: abrir_activos_fijos(root, frame_menu),
    )
    btn_card_activos.pack(side="left", padx=10)

    btn_card_sox = _crear_card(
        cards_row,
        "Control SOX",
        command=lambda: abrir_sox_menu(root, frame_menu),
    )
    btn_card_sox.pack(side="left", padx=10)

    btn_card_reportes = _crear_card(
        cards_row,
        "Reportes",
        disabled=True,
    )
    btn_card_reportes.pack(side="left", padx=10)

    # Botón de diagnóstico "Test conexión SAP": se conserva el código pero
    # NO se hace .pack para ocultarlo de la UI (puede reactivarse en el
    # futuro re-empaquetándolo). El handler `_test_conexion_sap_handler`
    # tampoco se borra.
    btn_test = tk.Button(
        frame_menu,
        text="Test conexión SAP",
        font=("Helvetica", 9),
        padx=10, pady=2,
        command=_test_conexion_sap_handler,
    )
    branding.aplicar_estilo_terciario(btn_test)
    # btn_test.pack()  # intencionalmente oculto

    frame_menu.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()
