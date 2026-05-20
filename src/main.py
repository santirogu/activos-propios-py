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

    No-op en sistemas sin pywin32 (Mac/Linux) — los workers fallarían de
    todas formas por la falta de SAP, pero al menos el módulo importa.
    """
    try:
        import pythoncom  # type: ignore
    except ImportError:
        yield
        return

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


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


def _poll_estado_boton_subir(root: tk.Tk, button: tk.Button) -> None:
    """Refresca el estado del botón y se re-programa cada `_POLL_INTERVAL_MS`."""
    _refrescar_estado_boton_subir(button)
    root.after(
        _POLL_INTERVAL_MS, lambda: _poll_estado_boton_subir(root, button)
    )


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
    dialog: tk.Toplevel,
    sociedad: str,
    fecha_desde: str,
    fecha_hasta: str,
    status_var: tk.StringVar,
    button: tk.Button,
) -> None:
    """Valida los inputs y lanza el worker que genera el reporte SOX."""
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

    def update_status(text: str) -> None:
        dialog.after(0, status_var.set, text)

    def show_info(title: str, message: str) -> None:
        dialog.after(0, lambda: messagebox.showinfo(title, message))

    def show_error(title: str, message: str) -> None:
        dialog.after(0, lambda: messagebox.showerror(title, message))

    def reenable() -> None:
        dialog.after(0, lambda: button.config(state="normal"))

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


def control_sox(root: tk.Tk) -> tk.Toplevel:
    """Abre el diálogo "Control SOX" con formulario (Sociedad + fechas).

    Devuelve la ventana `Toplevel` para que los tests puedan inspeccionarla.
    """
    from sox_report import VALID_SOCIEDADES, validar_caracter_fecha

    dialog = tk.Toplevel(root)
    dialog.title("Control SOX")
    dialog.geometry("500x340")
    dialog.resizable(False, False)
    dialog.transient(root)

    tk.Label(
        dialog, text="Control SOX", font=("Helvetica", 13, "bold")
    ).pack(pady=(18, 4))
    tk.Label(
        dialog,
        text="Genera el Reporte SOX con los parámetros indicados",
        font=("Helvetica", 10),
        fg="#555",
    ).pack(pady=(0, 12))

    form = tk.Frame(dialog)
    form.pack(pady=(0, 12))

    # --- Sociedad (Combobox readonly) ---
    tk.Label(form, text="Sociedad:", anchor="e", width=10).grid(
        row=0, column=0, padx=4, pady=6, sticky="e"
    )
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
    vcmd = (dialog.register(validar_caracter_fecha), "%P")
    fecha_hoy = datetime.now()

    tk.Label(form, text="Desde:", anchor="e", width=10).grid(
        row=1, column=0, padx=4, pady=6, sticky="e"
    )
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
    tk.Label(form, text="(dd.mm.aaaa)", fg="#777").grid(
        row=1, column=2, padx=4
    )

    tk.Label(form, text="Hasta:", anchor="e", width=10).grid(
        row=2, column=0, padx=4, pady=6, sticky="e"
    )
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
    tk.Label(form, text="(dd.mm.aaaa)", fg="#777").grid(
        row=2, column=2, padx=4
    )

    status_var = tk.StringVar()

    btn_generar = tk.Button(
        dialog,
        text="Generar Reporte SOX",
        font=("Helvetica", 11),
        padx=18,
        pady=6,
    )
    btn_generar.config(
        command=lambda: _generar_reporte_sox_handler(
            dialog,
            sociedad_var.get(),
            desde_var.get(),
            hasta_var.get(),
            status_var,
            btn_generar,
        )
    )
    btn_generar.pack()

    tk.Label(
        dialog,
        textvariable=status_var,
        font=("Helvetica", 9),
        fg="#1a7f37",
        wraplength=460,
    ).pack(pady=(12, 0))

    # Exponer widgets clave en el dialog para que los tests puedan inspeccionar.
    dialog.sociedad_var = sociedad_var
    dialog.desde_var = desde_var
    dialog.hasta_var = hasta_var
    dialog.status_var = status_var
    dialog.sociedad_combo = sociedad_combo
    dialog.desde_entry = desde_entry
    dialog.hasta_entry = hasta_entry
    dialog.btn_generar = btn_generar

    return dialog


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
    root.title("Creación Activos SAP")
    root.geometry("480x380")
    root.resizable(False, False)

    title = tk.Label(
        root,
        text="Creación de Activos Fijos en SAP",
        font=("Helvetica", 13, "bold"),
    )
    title.pack(pady=(18, 4))

    subtitle = tk.Label(
        root,
        text=f"Origen: resources/{EXCEL_PATH.name}\nDestino: salida/",
        font=("Helvetica", 10),
        fg="#555",
        justify="center",
    )
    subtitle.pack(pady=(0, 12))

    status_var = tk.StringVar(value="")

    btn_extraer = tk.Button(
        root,
        text="Extraer información en txt",
        command=lambda: extraer_lsmw_a_txt(status_var),
        font=("Helvetica", 11),
        padx=18,
        pady=6,
        width=24,
    )
    btn_extraer.pack(pady=(0, 8))

    btn_subir = tk.Button(
        root,
        text="Subir a SAP",
        font=("Helvetica", 11),
        padx=18,
        pady=6,
        width=24,
        state="disabled",
    )
    btn_subir.config(command=lambda: subir_a_sap(root, status_var, btn_subir))
    btn_subir.pack(pady=(0, 8))

    btn_sox = tk.Button(
        root,
        text="Control SOX",
        font=("Helvetica", 11),
        padx=18,
        pady=6,
        width=24,
        command=lambda: control_sox(root),
    )
    btn_sox.pack(pady=(0, 12))

    # Botón de diagnóstico: verifica si la conexión a SAP está disponible
    # sin ejecutar un flujo completo. Estilo secundario (más pequeño) para
    # marcar que es una herramienta de troubleshooting, no de uso normal.
    btn_test = tk.Button(
        root,
        text="Test conexión SAP",
        font=("Helvetica", 9),
        fg="#555",
        padx=10,
        pady=2,
        command=_test_conexion_sap_handler,
    )
    btn_test.pack()

    status = tk.Label(
        root,
        textvariable=status_var,
        font=("Helvetica", 9),
        fg="#1a7f37",
        wraplength=440,
    )
    status.pack(pady=(12, 0))

    # Polling: habilita el botón "Subir a SAP" cuando aparezca un .txt en
    # salida/ y lo deshabilita cuando no haya. Se re-programa cada segundo.
    _poll_estado_boton_subir(root, btn_subir)

    root.mainloop()


if __name__ == "__main__":
    main()
