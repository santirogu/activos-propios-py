import threading
import time
import traceback
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
from datetime import datetime

import openpyxl

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
        try:
            try:
                from sap_upload import (
                    get_latest_txt,
                    get_sap_session,
                    run_lsmw_flow,
                )
            except ImportError as exc:
                show_error("Error de import", f"No se pudo importar sap_upload:\n{exc}")
                return

            try:
                update_status("Buscando .txt más reciente en salida/...")
                latest = get_latest_txt()

                update_status("Conectando a la sesión SAP...")
                session = get_sap_session()

                update_status("Ejecutando flujo LSMW (no toques SAP)...")
                run_lsmw_flow(session, str(latest.parent), latest.name)

                update_status("Carga completada. Revisa SM35 para el log de la BDC.")
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


def main() -> None:
    root = tk.Tk()
    _install_tk_exception_handler(root)
    root.title("Creación Activos SAP")
    root.geometry("480x260")
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
    btn_subir.pack()

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
