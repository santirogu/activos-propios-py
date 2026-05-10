import threading
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
from datetime import datetime

import openpyxl

PROJECT_ROOT = Path(__file__).resolve().parent.parent
EXCEL_PATH = PROJECT_ROOT / "resources" / "Formato_Dinamico_.xlsx"
OUTPUT_DIR = PROJECT_ROOT / "salida"
SHEET_NAME = "LSMW "


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
        output_path, rows_written = export_sheet_to_tsv(
            EXCEL_PATH, SHEET_NAME, OUTPUT_DIR
        )
    except FileNotFoundError as exc:
        messagebox.showerror("Archivo no encontrado", str(exc))
        return
    except ValueError as exc:
        messagebox.showerror("Hoja no encontrada", str(exc))
        return
    except Exception as exc:
        messagebox.showerror("Error al exportar", str(exc))
        return

    status_var.set(f"Exportado: {output_path.name} ({rows_written} filas)")
    messagebox.showinfo(
        "Extracción completa",
        f"Se generó el archivo:\n{output_path}\n\nFilas exportadas: {rows_written}",
    )


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

    button.config(state="disabled")

    def update_status(text: str) -> None:
        root.after(0, status_var.set, text)

    def show_info(title: str, message: str) -> None:
        root.after(0, lambda: messagebox.showinfo(title, message))

    def show_error(title: str, message: str) -> None:
        root.after(0, lambda: messagebox.showerror(title, message))

    def reenable() -> None:
        root.after(0, lambda: button.config(state="normal"))

    def worker() -> None:
        try:
            from sap_upload import (
                SAP_LSMW_INPUT_PATH,
                copy_to_sap_path,
                get_latest_txt,
                get_sap_session,
                run_lsmw_flow,
            )
        except ImportError as exc:
            show_error("Error de import", f"No se pudo importar sap_upload:\n{exc}")
            reenable()
            return

        try:
            update_status("Buscando .txt más reciente en salida/...")
            latest = get_latest_txt()

            if SAP_LSMW_INPUT_PATH:
                update_status(f"Copiando archivo a {SAP_LSMW_INPUT_PATH}...")
                copy_to_sap_path(latest, SAP_LSMW_INPUT_PATH)

            update_status("Conectando a la sesión SAP...")
            session = get_sap_session()

            update_status("Ejecutando flujo LSMW (no toques SAP)...")
            run_lsmw_flow(session)

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
            reenable()

    threading.Thread(target=worker, daemon=True).start()


def main() -> None:
    root = tk.Tk()
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

    root.mainloop()


if __name__ == "__main__":
    main()
