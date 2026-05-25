# memory.md — Contexto del proyecto `activos-propios-py`

> Snapshot del proyecto al 2026-05-24. Generado leyendo `README.md`, `src/`, `tests/`, `resources/` y `requirements.txt`.

## 1. Propósito

Aplicación de escritorio en Python que **automatiza dos pasos del proceso de creación y capitalización de activos fijos en SAP**:

1. **Extracción** de la hoja `LSMW` del formato dinámico Excel a un `.txt` separado por tabulación.
2. **Carga** de ese `.txt` a SAP vía la transacción **LSMW**, ejecutando el flujo completo (Specify Files → Read Data → Convert Data → Create BI Session → Run BI) mediante **SAP GUI Scripting** (COM, pywin32).

Además, ofrece una tercera función:

3. **Generación del Reporte SOX** desde la transacción `AR15` de SAP, con exportación a Excel a `salida/`.

La autenticación SAP es manual; el script no logea al usuario, solo se conecta a una sesión ya abierta.

## 2. Stack y dependencias

- **Python 3.9+** con soporte Tkinter (en macOS, el Python de Homebrew 3.12 NO trae Tk → usar `python.org` o el del sistema).
- **openpyxl** ≥ 3.1 — lee/escribe Excel (multiplataforma).
- **tkcalendar** ≥ 1.6 — widget `DateEntry` para los campos de fecha SOX.
- **Pillow** ≥ 10.0 — captura de pantalla (`ImageGrab`) para las evidencias IPE del flujo SOX. Multiplataforma; en Windows captura desktop completo incluyendo barra de tareas.
- **pywin32** ≥ 306 — solo Windows (`platform_system == "Windows"`); requerido para "Subir a SAP" y "Generar Reporte SOX" (incluye `Shell.Application` COM para abrir el diálogo Propiedades del archivo SAP). Marker en `requirements.txt` lo evita en macOS/Linux.
- Pruebas: `unittest` (stdlib, sin dependencias extra).

## 3. Estructura del repo

```
.
├── src/
│   ├── main.py          # GUI Tkinter: 3 botones + test de conexión
│   ├── sap_upload.py    # Flujo LSMW completo (10 pasos) vía SAP GUI Scripting
│   ├── sox_report.py    # Flujo Reporte SOX (4 pasos) vía SAP GUI Scripting
│   └── branding.py      # Paleta corporativa Hub de ISA + helpers para Tk
├── tests/
│   ├── test_main.py         # 63 pruebas: extracción + handlers de botones + diálogo SOX
│   ├── test_sap_upload.py   # 36 pruebas: cada paso del flujo LSMW aislado
│   └── test_sox_report.py   # 42 pruebas: validaciones + flujo SOX
├── resources/
│   ├── Formato_Dinamico_.xlsx        # Excel maestro con hojas "Formato" y "LSMW "
│   ├── Población_ISA_31.03.2026.xlsx # Insumo del cliente
│   ├── script_sap_base.txt           # Grabación VBS del flujo LSMW (UTF-16)
│   ├── Script1.vbs / Script2.vbs     # Grabaciones VBS del flujo LSMW (paso Specify Files)
│   ├── Scriptsox.vbs                 # Grabación VBS original del SOX (árbol F00xxx, frágil)
│   └── Script2sox.vbs                # Grabación VBS actual del SOX (T-code AR15 + calendario F4)
├── docs/
│   └── flujo-proceso.png             # Diagrama del proceso end-to-end
├── salida/                           # Generada en runtime (ignorada por git)
├── requirements.txt
├── README.md                         # Documentación exhaustiva (≈370 líneas)
└── .gitignore                        # Ignora salida/, .venv/, __pycache__, .vscode/, .idea/, .DS_Store
```

Nota: la carpeta `salida/` está en `.gitignore` y se crea en runtime cuando se extraen .txt o se genera el reporte SOX.

## 3.5. Branding corporativo — `src/branding.py`

Centraliza paleta + logo para que `main.py` y `control_sox` mantengan apariencia consistente.

### Paleta (Hub de ISA)
- `ISA_AZUL = "#1A3A6C"` — navy del cuerpo del logo. Titulares, botones primarios.
- `ISA_AZUL_HOVER = "#2B5CA8"` — hover/pressed.
- `ISA_NARANJA = "#F58220"` — naranja del swoosh. Acento sobrio (no usado en botones primarios para no saturar).
- `ISA_NARANJA_HOVER = "#D96D14"`.
- `ISA_GRIS = "#555555"`, `ISA_GRIS_CLARO = "#888888"` — texto secundario.
- `ISA_BLANCO = "#FFFFFF"`, `ISA_FONDO = "#FFFFFF"`.
- `ISA_VERDE_OK = "#1a7f37"` — status success (conservado del esquema previo).

### Logo
- `LOGO_PATH = resources/logo_hub_isa.png` (PNG RGBA, 469×286 nativo).
- `cargar_logo(altura_px=55)` devuelve un `PIL.ImageTk.PhotoImage` escalado por aspect ratio, o `None` si: el archivo no existe / Pillow no está / Tk no está / error de carga. Loguea sólo el último caso (los demás son condiciones esperadas en dev).
- **Trampa Tk**: el caller DEBE guardar la referencia del PhotoImage en algún atributo persistente (ej. `root._logo_ref = cargar_logo()`). Si no, GC libera la imagen y Tk muestra cuadro vacío. Aplicado en `main()` y reusado en `control_sox`.

### Estilos de botón
- `aplicar_estilo_primario(btn)` — fondo navy, fg blanco, hover azul claro, `relief="flat"`, `bd=0`, `cursor="hand2"`, `font=("Helvetica", 11, "bold")`. Aplicado a Extraer, Subir a SAP, Control SOX, Generar Reporte SOX.
- `aplicar_estilo_terciario(btn)` — fondo blanco, fg gris, hover gris claro→fg navy, `relief="flat"`, `bd=1`. Aplicado a Test conexión SAP y ← Atrás.

Nota macOS: `tk.Button` ignora `bg`/`activebackground` en aqua por default (sólo respeta `fg`). En Windows funcionan. Como la app es Windows-only para SAP, en macOS los botones se ven nativos en vez de coloreados — aceptable.

## 4. GUI — `src/main.py`

Ventana principal (480x380, no redimensionable) con cuatro controles:

| Botón | Función | Plataforma |
|---|---|---|
| **Extraer información en txt** | `extraer_lsmw_a_txt` | Cualquier OS |
| **Subir a SAP** | `subir_a_sap` (arranca *disabled*; polling cada 1s habilita/deshabilita según `LSMW_*.txt` presentes en `salida/`) | Solo Windows |
| **Control SOX** | `control_sox(root, frame_menu)` — **reemplaza la vista del menú** en la misma ventana por un formulario con Sociedad + Desde + Hasta. Botón "← Atrás" arriba devuelve la vista al menú. **No abre Toplevel** (la versión original lo hacía). | Solo Windows |
| **Test conexión SAP** | `_test_conexion_sap_handler` (diagnóstico, estilo secundario) | Solo Windows |

### Detalles importantes

- `SHEET_NAME = "LSMW "` (con **espacio final** — así está nombrada la pestaña en el Excel).
- `_POLL_INTERVAL_MS = 1000` — polling para refrescar estado del botón "Subir a SAP".
- Flag módulo-level `_upload_en_curso` evita que el polling pise el estado del botón mientras corre el worker.
- Workers SAP corren en `threading.Thread(daemon=True)` para no congelar la GUI; comunican status vía `root.after(0, ...)` (thread-safe).
- `_sap_com_apartment()` — context manager que llama `pythoncom.CoInitialize()` / `CoUninitialize()` en cada worker. **Sin esto**, `GetObject('SAPGUI')` falla en threads no-main con error genérico aunque SAP esté abierto. No-op en macOS/Linux.
- `_install_tk_exception_handler(root)` reemplaza `root.report_callback_exception` para que excepciones en callbacks Tkinter abran un `messagebox.showerror` con traceback completo en vez de imprimir silenciosamente a stderr.
- `_show_unexpected_error(title, exc)` — red de seguridad para mostrar tipo + mensaje + traceback al usuario.
- `_log(mensaje)` imprime con timestamp `[HH:MM:SS]` y `flush=True`.

### Comportamiento del botón "Extraer"

- Si existe(n) `LSMW_*.txt` previos → diálogo SÍ/NO `messagebox.askyesno`. SÍ borra todos los previos y genera uno nuevo; NO conserva.
- Sin previos → genera directamente `LSMW_YYYYMMDD_HHMMSS.txt`.
- Validaciones manejadas explícitamente: `FileNotFoundError` (Excel ausente), `ValueError` (hoja ausente), `Exception` (genérica del export), y red de seguridad que muestra traceback completo.

### Vista Control SOX (frame embebido, no Toplevel)

- **Patrón de switching de vistas:** `main()` envuelve todos los widgets del menú en un `frame_menu` (en vez de poner los widgets directo en `root`). Cuando el usuario presiona "Control SOX", `control_sox(root, frame_menu)` hace `frame_menu.pack_forget()` y muestra un nuevo `frame_sox` con el formulario + un botón "← Atrás". El click en Atrás destruye `frame_sox` y re-empaca `frame_menu`. El estado del menú (status_var, polling, flag `_upload_en_curso`) se preserva porque sólo se oculta, no se destruye.
- **Sociedad**: `ttk.Combobox` en estado `readonly` (el usuario no puede escribir libre). Opciones: `TRAN, ISA, ITCH, CEYBA, CABA, RPAE, CTMP, REPD, ISAP`.
- **Desde/Hasta**: `DateEntry` de tkcalendar con `date_pattern="dd.mm.yyyy"`. Validación per-keystroke (`validar_caracter_fecha`) acepta solo dígitos y puntos, máx 10 caracteres. Inicializa con la fecha actual.
- Validaciones al pulsar **Generar Reporte SOX**:
  1. Sociedad en lista permitida (normaliza con `.strip().upper()`).
  2. Ambas fechas formato `dd.mm.aaaa` válido.
  3. `Hasta >= Desde`.
- **Durante el worker SOX**: tanto el botón Generar como el Atrás se deshabilitan; ambos se re-habilitan al finalizar (éxito o error). El usuario no puede volver al menú a mitad de un flujo SAP. Se logra porque `_generar_reporte_sox_handler` recibe ahora `(root, ..., button, btn_atras)` y usa `root.after` (no el viejo `dialog.after`) para callbacks thread-safe.

## 5. Flujo SAP LSMW — `src/sap_upload.py`

Replica las grabaciones VBS de SAP. Granularidad fina: cada paso es una función dedicada para poder testearlo aislado con `MockSAPSession`.

### Funciones de soporte
- `get_latest_txt(salida_dir)` → `Path` del `LSMW_*.txt` más reciente por mtime. Lanza `FileNotFoundError`.
- `get_sap_session()` → primera sesión activa vía `win32com.client.GetObject("SAPGUI")`. Lanza `RuntimeError` con mensajes accionables (pywin32 ausente, SAP cerrado, scripting deshabilitado, sin conexiones, sin sesiones).
- `diagnosticar_conexion_sap()` → tupla `(ok, mensaje)`. Detecta y reporta el estado paso a paso (con `SystemName/Client/User` de cada sesión).
- `_ejecutar(descripcion, fn, *args, **kwargs)` → wrapper que loguea y, si falla, re-lanza `RuntimeError` con descripción humana + repr de la excepción COM (que suele venir vacía).
- `_confirmar_popup_opcional(session, descripcion)` → intenta `wnd[1].sendVKey(0)` (Enter en popup). Si no hay popup, loguea y sigue. Clave para resistir popups condicionales.
- `_volver_al_step_list(session, max_intentos=3)` → garantiza retorno al step list buscando `LSMW_STEPLIST_TABLE`; si no, envía F3 (Back) iterativamente. Resuelve flakiness de SAP que a veces no auto-retorna tras confirmar popups.

### Orquestador
`run_lsmw_flow(session, carpeta, nombre_archivo)` ejecuta los 10 pasos secuencialmente.

### Mapeo del flujo LSMW (10 pasos)

| # | Función | Fila step list | Acciones SAP |
|---|---|---|---|
| 1 | `open_lsmw` | — | maximize + okcd="LSMW" + Enter + F8 |
| 2 | `configurar_ruta_archivo(carpeta, nombre)` | 6 (Specify Files) | F2 + btn[25] (Cambiar) + lbl[43,6] + btn[27] (Asignar) + F4 picker + DY_PATH/DY_FILENAME + 2×OK + Back + SPOP-OPTION1 (popup *opcional*) |
| 3 | `step_assign_files` | 7 | btn[32] + F3 |
| 4 | `step_read_data` | 8 | btn[32] + F8 + 2×F3 |
| 5 | `step_display_read_data` | (auto-avanza) | btn[32] + popup opcional + F3 |
| 6 | `step_convert_data` | (auto-avanza) | btn[32] + F8 + 2×F3 |
| 7 | `step_display_converted_data` | (auto-avanza) | btn[32] + popup opcional + F3 |
| 8 | `step_create_batch_input` | (auto-avanza) | btn[32] + chkP_KEEP=True + F8 + popup + `_volver_al_step_list` |
| 9 | `step_run_batch_input` | 13 (explícita) | `select_step_row` + btn[32] |
| 10 | `process_bdc_session` | (tabla BDC) | row[0] + GROUPID focus + F8 + radD0300-ERROR + chkLOGALL + chkEXPERT + 2×OK |

### Constantes clave
- `LSMW_STEPLIST_TABLE = "wnd[0]/usr/tbl/SAPDMC/SAPLLSMW_OBJ_000TC_STEPLIST"`
- `DEFAULT_SELECTED_ROW = 13` (SAP marca esta fila por default; `select_step_row` la deselecciona antes de elegir la objetivo).
- `BDC_SESSION_TABLE = "wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI"`
- Filas: `SPECIFY_FILES_ROW=6`, `ASSIGN_FILES_ROW=7`, `READ_DATA_ROW=8`, `RUN_BI_ROW=13`.

## 6. Flujo SOX — `src/sox_report.py`

Replica `resources/Script2sox.vbs` (versión actualizada con T-code AR15 + calendario F4, reemplazando la grabación inicial frágil con nodos F00xxx del árbol que sigue en `Scriptsox.vbs`).

### Constantes clave
- `T_CODE_SOX = "AR15"` — camino preferido (robusto). Si es `None`, hace fallback al árbol con `SOX_NODE_KEY = "F00039"` (frágil entre usuarios).
- `TREE_SHELL = "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell"` — árbol del menú SAP.
- `CALENDAR_SHELL = "wnd[1]/usr/cntlCONTAINER/shellcont/shell"` — calendario emergente F4.
- `DATE_FORMAT_USER = "%d.%m.%Y"` (formulario) y `DATE_FORMAT_SAP_CALENDAR = "%Y%m%d"` (calendario SAP).
- Campos: `CAMPO_SOCIEDAD = "wnd[0]/usr/ctxtP_BUKRS"`, `CAMPO_FECHA_DESDE = "wnd[0]/usr/ctxtS_DATUM-LOW"`, `CAMPO_FECHA_HASTA = "wnd[0]/usr/ctxtS_DATUM-HIGH"`.
- `EXPORT_METHOD = "alv_grid"` (default) — usa `&MB_EXPORT > &XXL` sobre `DOCS_GRID_SHELL`. Confirmado por `resources/Script2sox.vbs` para AR15 (que muestra un ALV grid, no lista clásica). Alternativas: `"pc_list"` (usa `%PC`, sólo aplica a listas clásicas — NO funciona en AR15) o `None` (deja al usuario guardar manualmente y omite el paso 5).
- `ALV_SAVE_DIALOG_OK_BTN = "btn[11]"` — el diálogo de save abierto por `&XXL` en AR15 confirma con `btn[11]` (Generar/Reemplazar). `btn[0]` no existe en ese diálogo y fue el origen del error `"The control could not be found by id"` cuando el default era `pc_list`. `_rellenar_save_dialog(..., boton_ok_id=...)` permite parametrizar cuál botón presionar.
- `STANDARD_FILE_PREFIX = "Población"` y `STANDARD_SHEET_NAME = "Original_SAP"` — usados por `generar_xlsx_poblacion` para nombrar el archivo final y la hoja interna. Patrón: `Población_{SOCIEDAD}_{FECHA_HASTA}.xlsx`.
- `CREADOS_SHEET_NAME = "Creados"`, `CREADOS_FILTRO_VALOR = "*** creado ***"`, `PATRON_AF = re.compile(r"^AF\s+(\d+)-(\d+)\s+(.+)$")`, `CREADOS_HEADERS`, `CREADOS_OBSERVACIONES` — usados por `generar_hoja_creados` para producir la segunda hoja del Población. Las columnas K y L son **fórmulas Excel** que se evalúan al abrir el archivo (no valores pre-calculados): K = `=MID(D{n},1,2)`, L = `=IF(K{n}="19","Intangible",IF(K{n}="20","Activo Construcción",IF(K{n}="14","Activo Construcción","PPE")))`. Las fórmulas se escriben en **inglés con `,` como separador** (estándar OOXML); Excel-ES las muestra automáticamente como `EXTRAE` y `SI` en la barra de fórmulas. Header de L = `"PPE o Intangible"`.
- `IPE_SHEET_NAME = "IPE"`, `IPE_SCREENSHOTS_INFO` (tupla de 5 `(filename, descripcion)`), `IPE_IMAGE_MAX_WIDTH = 1200` — usados por `generar_hoja_ipe` para construir la tercera hoja del Población con 5 capturas de pantalla embebidas: (1) parámetros ingresados antes de F8, (2) primer registro del grid AR15, (3) último registro (scroll al final), (4) status bar SAP con bytes exportados, (5) diálogo Propiedades del archivo SOX en Windows. Las capturas son **soft-fail**: si una falla (PIL ausente, scroll del grid no funciona, diálogo Propiedades no abre), se anota como "no disponible" en la hoja IPE y se sigue.

### Mapeo del flujo (7 etapas)

| # | Función | Acciones SAP / Python |
|---|---|---|
| 1 | `abrir_transaccion_sox` | maximize + okcd="AR15" + Enter. Fallback: `tree.doubleClickNode("F00039")` |
| 2a | `ingresar_parametros` | `P_BUKRS.text = sociedad` (ya NO incluye F8 — se split para permitir captura) |
| 2b | `_seleccionar_fecha_calendario` (Desde) | foco campo + caretPosition 0 + F4 → calendar.focusDate / selectionInterval con `yyyymmdd` |
| 2c | `_seleccionar_fecha_calendario` (Hasta) | Igual para S_DATUM-HIGH |
| — | `_capturar_pantalla` | **Screenshot 1** (parámetros llenados, antes de F8) → `01_parametros_ingresados.png` en tempdir |
| 3 | `ejecutar_reporte` | F8 (`tbar[1]/btn[8]`) — split de `ingresar_parametros` para permitir captura entre ambos |
| — | `_scroll_grid_a_primero` + `_capturar_pantalla` | **Screenshot 2** del primer registro del grid |
| — | `_scroll_grid_a_ultimo` + `_capturar_pantalla` | **Screenshot 3** del último registro (usa `grid.RowCount` y `firstVisibleRow`) |
| 4 | `exportar_a_excel` → `_exportar_via_alv_grid` (default) o `_exportar_via_pc_list` | ALV: `&MB_EXPORT` + `&XXL` sobre `DOCS_GRID_SHELL` + `_rellenar_save_dialog(..., boton_ok_id="btn[11]")`. PC: `%PC` + manejo de variantes A/B/C del save-as + `_rellenar_save_dialog(..., boton_ok_id="btn[0]")`. Produce `SOX_{SOC}_{YYYYMMDD_HHMMSS}.xlsx` intermedio. |
| — | `_capturar_pantalla` | **Screenshot 4** (status bar SAP con bytes exportados, tomada justo tras exportar_a_excel) |
| — | `_capturar_propiedades_archivo` | **Screenshot 5** — abre el diálogo Propiedades de Windows vía `Shell.Application` COM, captura, y cierra con Escape (`user32.keybd_event(0x1B)`) |
| 5 | `generar_xlsx_poblacion` (post-SAP, pure Python) | Lee el intermedio con openpyxl, copia su contenido (celda por celda, preservando `number_format` para que Fecha y Hora se vean como en SAP) a una nueva hoja `Original_SAP` y guarda como `Población_{SOC}_{FECHA_HASTA}.xlsx`. Si `EXPORT_METHOD=None`, este paso y los siguientes se omiten. |
| 6 | `generar_hoja_creados(poblacion)` (post-procesamiento, pure Python) | Abre el Población, lee `Original_SAP`, filtra filas con G == `*** creado ***`, parsea la columna D con `PATRON_AF` y produce una **segunda hoja `Creados`** con observaciones (filas 1-9), headers en bold (fila 10), datos desde fila 11. Columnas K y L como fórmulas Excel en inglés (estándar OOXML, Excel-ES traduce a EXTRAE/SI). |
| 7 | `generar_hoja_ipe(poblacion, screenshots_dir)` (paso final, pure Python) | Lee los 5 PNG del tempdir y los embebe en una **tercera hoja `IPE`** con título + descripción + imagen escalada a `IPE_IMAGE_MAX_WIDTH=1200px`. Soft-fail: capturas faltantes (porque PIL no estaba, el scroll falló, o el diálogo Propiedades no abrió) se anotan como "no disponible" pero el flujo continúa. Es el **deliverable final** que devuelve `generar_reporte_sox`. El tempdir se limpia automáticamente al salir del `with tempfile.TemporaryDirectory(...)`. |

### Helpers
- `validar_sociedad` / `validar_fecha` / `validar_rango_fechas` / `validar_caracter_fecha` — validaciones puras, testeables sin SAP.
- `_intentar_listar_nodos_arbol(tree)` — diagnóstico: enumera nodos del árbol SAP (`GetAllNodeKeys` + `GetNodeTextByKey`) cuando falla `doubleClickNode`. Mensaje de error sugiere descubrir la T-code real vía "Sistema → Estado".
- Salida: dos archivos en `salida/`:
  - **Intermedio:** `SOX_{SOCIEDAD}_{YYYYMMDD_HHMMSS}.xlsx` (lo que SAP exportó).
  - **Final / deliverable:** `Población_{SOCIEDAD}_{FECHA_HASTA}.xlsx` con **tres hojas**:
    - `Original_SAP`: copia 1:1 del intermedio (preserva `number_format`).
    - `Creados`: filas de `Original_SAP` filtradas por `G == "*** creado ***"`, con código/subnúmero/denominación parseados de la col D, + columnas K (prefijo del código como texto) y L (clasificación PPE/Intangible/Activo Construcción). Bloque de observaciones en filas 1-9.
    - `IPE`: 5 capturas de pantalla embebidas (parámetros, primer registro, último registro, status bar con bytes, propiedades del archivo) como evidencia visual del proceso. El handler GUI muestra este nombre al usuario.

### CLI
```bash
python src/sox_report.py ISA 01.05.2026 31.05.2026
```
Exit codes: 0 OK, 1 error (validación o SAP), 2 uso incorrecto.

## 7. Datos: Hoja LSMW del Excel

- La hoja `LSMW ` (con espacio) está cableada con fórmulas que referencian la hoja `Formato`.
- `openpyxl` lee con `data_only=True` los valores **cacheados** por Excel en el último guardado → si el usuario edita el formulario, **debe abrir y guardar el Excel** antes de extraer para que las fórmulas se recalculen.
- Celdas vacías referenciadas pueden aparecer como `0`.
- 51 columnas exportadas — algunos campos: `ANLKL` (clase de activo), `BUKRS` (sociedad), `TXT50` (denominación), `KOSTL` (centro de costo), `WERKS` (centro), `EAUFN` (orden de inversión), `POSNR` (elemento PEP), `ORD41`–`ORD44` y `GDLGRP` (criterios de clasificación 1–5).

## 8. Pruebas — 151 tests con `unittest`

### Estrategia de mocking SAP
- `MockSAPSession` registra cada llamada `findById(...).method()` en `session.actions` como tuplas `(sap_id, method, *args)`.
- Expone elementos vía `session._elements[id]` para inspeccionar propiedades (`text`, `selected`, `caretPosition`).
- Filas de tablas: `_MockRow` con setter que loguea cambios de `selected`.
- Permite verificar la secuencia exacta de IDs y métodos sin un SAP real.

### Estrategia de mocking GUI
- `_SyncFakeThread` reemplaza `threading.Thread` para ejecutar el worker síncrono.
- `root.after` se sobreescribe en `setUp` para invocar callbacks inmediatamente.
- `patch.multiple("sap_upload", ...)` inyecta mocks; se guardan en `self.mocks`.

### Distribución
- `tests/test_main.py` (75): handlers GUI, extracción TSV, vista SOX como frame embebido (`ControlSoxDialogTest` verifica el ocultado del menú, exposición del botón Atrás, y reversión al menú; `GenerarReporteSoxHandlerTest` verifica que ambos botones Generar+Atrás se deshabilitan durante el worker).
- `tests/test_sox_report.py` (105): validaciones puras + pasos del flujo SAP + `GenerarXlsxPoblacionTest` (11 tests del paso Población) + `PatronAfRegexTest` (6 tests del parseo de col D) + `GenerarHojaCreadosTest` (16 tests del paso post-procesamiento: filter, parsing, estructura, K y L como fórmulas =MID/=IF, replace idempotente, errores) + `EjecutarReporteTest` (2 tests del split de F8) + `GenerarHojaIpeTest` (8 tests de la hoja de evidencias: crear, embedding, soft-fail, replace, scaling) + `GenerarReporteSoxTest` (verifica orden de las 7 etapas incluyendo screenshots, paso a las funciones, `EXPORT_METHOD=None` salta Población/Creados/IPE).
- `tests/test_sap_upload.py` (46): cada paso del flujo LSMW + `MainEntryPointTest`.

### Cómo ejecutar
```bash
python -m unittest discover tests -v
python -m unittest tests.test_main -v
python -m unittest tests.test_main.SubirASapTest.test_worker_calls_full_flow_on_happy_path
```

## 9. Configuración SAP (una sola vez por máquina)

1. **Cliente** — habilitar scripting: *Options → Accessibility & Scripting → Scripting → "Enable scripting"*. Recomendado desmarcar los dos *"Notify when..."*.
2. **Servidor** — parámetro `sapgui/user_scripting = TRUE` (transacción RZ11). Pedir a Basis si no está activo.
3. **Iniciar sesión SAP** antes de ejecutar (el script no autentica).
4. **Pre-cargar el proyecto LSMW** una vez manualmente con Subproject + Object correctos (SAP recuerda la última selección).

## 10. Diagnóstico — errores comunes en la carga SAP

| Error | Causa probable | Solución |
|---|---|---|
| "No se pudo conectar a SAP GUI" | SAP no abierto o scripting deshabilitado | Abrir SAP, habilitar scripting |
| "No hay sesiones activas" | Pantalla de login | Iniciar sesión SAP |
| "Falta la dependencia pywin32" | Mac/Linux o deps no instaladas | `pip install pywin32` (solo Windows) |
| Falla en `select_step_row` | Proyecto LSMW incorrecto | Abrir LSMW manualmente con proyecto correcto |
| Falla en `configurar_ruta_archivo` | Definición de archivo en otra posición | Re-grabar `Script1.vbs` y ajustar IDs (`lbl[43,6]`, `btn[25]`, `btn[27]`) |
| Falla en `step_read_data` | Archivo no existe en ruta inyectada o sin permisos | Verificar `salida/<archivo>` y permisos SAP |

## 11. Git / estado actual

- **Rama actual:** `master` (también es la rama principal del repo).
- **Working tree:** limpio al 2026-05-24.
- **Últimos commits:**
  - `362d919` cambios
  - `0ba2774` cambios
  - `22800cc` fix(lsmw): garantizar retorno al step list entre pasos 8, 9 y 10
  - `f26abd3` fix(lsmw): instrumentar pasos 3-10 y hacer popups opcionales
  - `9cc7793` fix(lsmw): popup 'Sí guardar cambios' debe ser opcional
- Convención de commits observada: `fix(lsmw): …` o "cambios" libre. Los fixes recientes giran alrededor de hacer el flujo LSMW resistente a popups condicionales y al hecho de que SAP no siempre auto-retorna al step list.

## 12. Decisiones de diseño no obvias

- **Granularidad fina de funciones por paso SAP** — no es sobre-ingeniería sino que permite testear cada paso aislado con `MockSAPSession`. Sin esto, habría que mockear el flujo completo de 10 pasos para verificar uno.
- **Popups condicionales (`_confirmar_popup_opcional`)** — SAP a veces muestra popup, a veces no, según si hay cambios pendientes. La función intenta `wnd[1].sendVKey(0)` y captura el error sin romper.
- **Retorno al step list (`_volver_al_step_list`)** — SAP no siempre auto-retorna tras confirmar un popup; los pasos 8/9/10 son donde esto se manifestaba como flakiness. La función envía F3 hasta encontrar la tabla, con tope de intentos.
- **Selección explícita de la fila 13 en `step_run_batch_input`** — no se confía en el auto-advance del cursor de SAP, se selecciona explícitamente para hacer el flujo determinista entre corridas.
- **Apartamento COM en threads** — `pythoncom.CoInitialize()` obligatorio en threads no-main de Windows; sin esto el COM falla aunque SAP esté abierto.
- **Import lazy de `sap_upload` / `sox_report` dentro del worker** — permite que `main.py` arranque en macOS/Linux sin pywin32 (los botones SAP fallan al ejecutarse, pero la GUI carga).
- **Hoja `"LSMW "` con espacio final** — así está nombrada en el Excel; si se quita el espacio se rompe el lookup.
- **Polling cada 1s para habilitar "Subir a SAP"** — más simple que un watcher de filesystem y suficiente para la cadencia de uso esperada.
- **EXPORT_METHOD configurable** — `"alv_grid"` (default) para grids ALV con `&MB_EXPORT > &XXL`; `"pc_list"` para listas SAP clásicas vía `%PC`; `None` para no exportar. AR15 muestra un ALV grid, confirmado por `resources/Script2sox.vbs`.
- **Dos archivos en `salida/` por corrida del SOX** — el intermedio `SOX_*.xlsx` (lo que SAP produjo, nombre con timestamp) y el final `Población_*.xlsx` (lo que `generar_xlsx_poblacion` produjo a partir del anterior). Se conservan ambos; el handler GUI muestra el nombre del final al usuario.
- **`generar_xlsx_poblacion` es paso post-SAP, no SAP** — corre puro Python con openpyxl después de que SAP guardó su archivo. Esto permite testearlo sin mockear SAP (usar tempdir + .xlsx real) y separar la lógica de "qué hace SAP" de "cómo se llama el deliverable final".
- **Nombre estándar `Población_{SOC}_{FECHA_HASTA}.xlsx`** — formato fijo pedido por el cliente. La fecha es la `fecha_hasta` que el usuario ingresó (dd.mm.aaaa), re-formateada vía `validar_fecha→strftime` para normalizar whitespace y garantizar formato consistente.
- **openpyxl no preserva strings vacías en round-trip** — guarda celda vacía → lee `None`. Para el reporte SOX esto es funcionalmente equivalente; los tests usan datos sin strings vacías para evitar la confusión.
- **`generar_xlsx_poblacion` copia celda por celda (no `iter_rows(values_only=True)`)** — para preservar `number_format`. SAP usa `'mm-dd-yy'` para Fecha y `'[$-F400]h:mm:ss\ AM/PM'` para Hora; sin esto el usuario veía `2026-03-02 0:00:00` y `13:00:49` (defaults de openpyxl) en vez del formato corto + AM/PM del original SAP. Solo se copia `value` y `number_format`, no estilos completos (font/fill/border) — el archivo Población es más liviano y la única diferencia visual relevante eran las columnas de tiempo.
- **Control SOX usa frame switching, no Toplevel** — `main()` empaca todos los widgets del menú en `frame_menu`. `control_sox(root, frame_menu)` hace `pack_forget` del menú y muestra `frame_sox` en `root`; el botón "← Atrás" destruye `frame_sox` y re-empaca `frame_menu`. Razón: UX más fluida (una sola ventana, sin saltos de focus a un Toplevel). Las StringVars/widgets del menú se preservan porque sólo se oculta el frame, no se destruye — el polling del botón "Subir a SAP" sigue activo incluso mientras el usuario está en el form SOX.
- **`generar_hoja_creados` se ejecuta en el mismo workbook (no produce un archivo nuevo)** — añade la hoja `Creados` al `Población_*.xlsx` y guarda sobre sí mismo. Idempotente: si la hoja ya existe, se borra y recrea (no se hace append). Esto mantiene un solo deliverable para el usuario; no hay que coordinar dos archivos.
- **Filas que no matchean el regex se loguean y omiten, no rompen el flujo** — `PATRON_AF` puede fallar si SAP genera un texto inesperado en col D (raro pero posible). En vez de abortar todo el reporte, se cuentan en `filas_descartadas` y se loguean con el texto crudo para que el usuario pueda investigar. El conteo `filas_filtradas` vs `filas_escritas` permite detectar discrepancias.
- **Columnas K y L en `Creados` son fórmulas Excel, no valores pre-calculados** — `=MID(D{n},1,2)` y `=IF(...IF(...IF(...)))` anidado. Razón: cliente lo pidió explícitamente; mantener fórmulas permite que un usuario que edite D manualmente vea K y L recalcular en Excel. openpyxl serializa strings que empiezan con `=` como fórmulas (`data_type='f'` en el XLSX). Como openpyxl NO evalúa fórmulas, los tests sólo pueden verificar el string de la fórmula — la lógica de clasificación se valida visualmente al abrir en Excel.
- **Fórmulas en INGLÉS con `,` separador, NO en español** — el estándar OOXML del .xlsx exige nombres de funciones en inglés en el XML interno. Excel lee el XML, valida la sintaxis y TRADUCE al locale del usuario al mostrar (Excel-ES verá `EXTRAE` y `SI` en la barra de fórmulas). Escribir `EXTRAE`/`SI.CONJUNTO` directamente en el XML hace que Excel reporte el archivo como dañado ("Hemos encontrado un problema con contenido…"). Aprendido por la mala — la intuición de "openpyxl no traduce, escribe en español" estaba equivocada porque Excel valida el XML al abrirlo.
- **Usamos `IF` anidado en vez de `IFS`** — `IFS` es "future function" (Excel 2016+) y necesita prefijo `_xlfn.IFS` para ser válido en el XML. `IF` es universal y no requiere ningún prefijo, evitando esa fuente de bugs.
- **Las capturas de pantalla (IPE) son soft-fail** — si PIL no está, el grid no expone `RowCount`, o el diálogo Propiedades de Windows no abre, la captura se omite y se anota como "no disponible" en la hoja IPE. El flujo SOX nunca se rompe por una evidencia fallida. `generar_hoja_ipe` devuelve `{embedded, missing, missing_names}` para diagnóstico.
- **Las 5 capturas viven en un `tempfile.TemporaryDirectory`** — se generan durante el flujo SAP, se embeben en la hoja IPE del Población al final, y el tempdir se borra automáticamente al salir del `with`. Los PNG sueltos nunca se exponen al usuario; el deliverable es self-contained.
- **La ventana Tkinter de la app se minimiza ANTES de las capturas y se restaura al final** — sin esto, las screenshots IPE incluyen la UI de "Creación Activos SAP" tapando parte de la pantalla de SAP. `_minimizar_ventanas_app` usa `win32gui.EnumWindows` para encontrar HWNDs cuyo título esté en `TITULOS_VENTANA_APP` y los minimiza con `ShowWindow(SW_MINIMIZE)`, trackándolos en `_VENTANAS_MINIMIZADAS_PARA_CAPTURA`. `_restaurar_ventanas_app` los devuelve con `SW_RESTORE`. La restauración va en `try/finally` para que el usuario nunca se quede sin GUI si una etapa falla a media corrida. En macOS/Linux ambos helpers son no-op (no hay `win32gui`).
- **`TITULOS_VENTANA_APP = ("Creación Activos SAP",)` es una constante** — debe coincidir con `root.title(...)` en `main.py`. Si en el futuro se renombra la ventana hay que actualizar acá también; un comentario en `main.py` cerca del `root.title` ayudaría pero no se añadió para no contaminar el archivo de la GUI con dependencias de sox_report.
- **`_capturar_propiedades_archivo` usa `ShellExecuteExW` vía ctypes (no `Shell.Application` COM)** — el camino `Shell.Application.NameSpace + ParseName + InvokeVerb("Properties")` resultó inestable cuando el thread comparte apartment COM con SAP GUI Scripting: a veces caía al namespace del Escritorio en vez del archivo real y mostraba "Las propiedades para este archivo no están disponibles" en una caja de error. `ShellExecuteExW` con `lpVerb="properties"` + `fMask=SEE_MASK_INVOKEIDLIST` es la API canónica del shell, no usa COM, y resuelve el archivo correctamente. Se cierra el diálogo con ESC vía `ctypes.user32.keybd_event` (alternativa `pyautogui` rechazada porque añade dependencia cuando `ctypes` ya viene en stdlib).
- **`_esperar_archivo_listo(archivo, timeout=10s)` antes de invocar Propiedades** — SAP cierra el handle del archivo de forma asíncrona después de exportar (1-3s típicos), y abrir Propiedades sobre un archivo aún siendo escrito produce "propiedades no disponibles". El helper hace polling cada 0.5s al `stat().st_size` y devuelve True cuando el tamaño se mantiene constante entre dos ticks (= file ya no está creciendo). Si el archivo no se estabiliza en 10s se loguea pero se intenta abrir Propiedades igual.
- **`ingresar_parametros` y `ejecutar_reporte` están separadas** — split necesario para capturar el screenshot del estado del formulario ANTES de F8. Antes era una sola función con `_log("Paso 2/4")` + F8 al final.
- **Columna K usa `number_format="@"` como refuerzo declarativo** — `MID` siempre devuelve texto, así que el formato es redundante para el valor evaluado. Pero asegura que si alguien sobrescribe la fórmula con un valor pegado (ej. "08"), Excel siga interpretándolo como texto y no convierta a 8.
- **`generar_hoja_creados` usa openpyxl puro (no pandas)** — para no añadir una dependencia pesada (~50 MB) por una sola función. Para 500k filas en normal mode toma ~30-60 segundos pero es aceptable como paso síncrono del worker. Si la performance se vuelve un problema, evaluar `read_only=True` para la lectura + dos pasadas (read en read_only, write en normal); o añadir pandas como dependencia opcional.
