# memory.md — Contexto del proyecto `activos-propios-py`

> Snapshot del proyecto al 2026-08-09. Generado leyendo `README.md`, `src/`, `tests/`, `resources/` y `requirements.txt`.

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
│   ├── extraer_activos_creados.py  # Flujo SM35P (4 pasos): filtro por usuario + export log
│   ├── subir_anexos.py  # Flujo AS02 + GOS PCATTA_CREA: adjunta archivos a cada activo
│   ├── branding.py      # Paleta corporativa Hub de ISA + helpers para Tk
│   └── paths.py         # Resolución de rutas dev/bundled + factory default + carpeta entrada/
├── tests/
│   ├── test_main.py                       # 100 pruebas: extracción + botones + vistas + footer + splash + regresión calendario
│   ├── test_paths.py                      # 18 pruebas: helpers dev/bundled + resolución entrada/ + factory default + validación 1-xlsm
│   ├── test_branding.py                   # 15 pruebas: paleta + logo + estilos de botón
│   ├── test_sap_upload.py                 # 46 pruebas: cada paso del flujo LSMW aislado
│   ├── test_sox_report.py                 # 105 pruebas: validaciones + flujo SOX + Población + Creados + IPE
│   ├── test_extraer_activos_creados.py    # 48 pruebas
│   └── test_subir_anexos.py               # 29 pruebas
├── resources/                        # Recursos internos del proyecto (bundleados read-only en el .exe)
│   ├── Formato_Dinamico.xlsm         # Formato maestro (.xlsm con macros); factory default con hojas "Formato" y "LSMW "
│   ├── Población_ISA_31.03.2026.xlsx # Insumo del cliente
│   ├── logo_hub_isa.png              # Logo corporativo Hub de ISA (RGBA 469×286)
│   ├── script_sap_base.txt           # Grabación VBS del flujo LSMW (UTF-16)
│   ├── Script1.vbs / Script2.vbs     # Grabaciones VBS del flujo LSMW (paso Specify Files)
│   ├── Script2sox.vbs                # Grabación VBS del SOX (T-code AR15 + calendario F4)
│   ├── ScriptSM35P.vbs               # Grabación VBS del flujo Extraer Activos Creados
│   └── Scriptanexo.vbs               # Grabación VBS del flujo Subir Anexos
├── docs/
│   └── flujo-proceso.png             # Diagrama del proceso end-to-end
├── entrada/                          # Generada en runtime junto al .exe (ignorada por git): el Formato_Dinamico.xlsm editable
├── salida/                           # Generada en runtime (ignorada por git)
├── requirements.txt
├── README.md                         # Documentación exhaustiva (≈370 líneas)
└── .gitignore                        # Ignora salida/, entrada/, .venv/, __pycache__, .vscode/, .idea/, .DS_Store
```

Nota: las carpetas `salida/` y `entrada/` están en `.gitignore` y se crean en runtime. `salida/` recibe los .txt y reportes generados. `entrada/` es donde el usuario deja el `Formato_Dinamico.xlsm` (input del proceso): se llama `entrada/` — no `resources/` — para no confundirla con la `resources/` interna del proyecto (logo, factory default, grabaciones VBS). En el primer arranque `paths.asegurar_formato_dinamico()` copia el factory default desde `resources/` a `entrada/` si esta no tiene ningún `.xlsm`.

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

Ventana principal (620x480, no redimensionable, título "Gestión de Activos Fijos") con **3 cards horizontales** (mismo ancho, navy primario). Navegación por sub-frames (no Toplevel) con botón "← Atrás" en cada vista.

### Layout principal

```
┌─────────────────────────────────────┐
│           [LOGO 85px]               │
│      Gestión de Activos Fijos       │
│                                     │
│  ┌──────┐ ┌──────┐ ┌──────────┐    │
│  │Activos│ │Control│ │Reportes  │   │
│  │Fijos │ │ SOX  │ │(disabled)│    │
│  └──────┘ └──────┘ └──────────┘    │
└─────────────────────────────────────┘
```

### Cards del menú principal

| Card | Función | Plataforma |
|---|---|---|
| **Activos Fijos** | Abre `abrir_activos_fijos(root, frame_menu)` — sub-vista con Extraer + Creación de Activo | Cualquier OS / Solo Windows según botón |
| **Control SOX** | Abre `abrir_sox_menu(root, frame_menu)` — sub-vista intermedia con un único botón "HUB.PPE.01 Creación de Activos Fijos" que abre el form clásico de sociedad+fechas | Solo Windows |
| **Reportes** | `state="disabled"` — placeholder para futuras funcionalidades | — |

### Sub-vistas

| Sub-vista | Función | Contenido |
|---|---|---|
| **Activos Fijos** (`abrir_activos_fijos`) | ← Atrás + logo + título "Activos Fijos" | `[Extraer información en txt]` (función `extraer_lsmw_a_txt`) + `[Creación de Activo]` (función `subir_a_sap`, arranca *disabled*; polling **scoped al frame** cada 1s habilita/deshabilita según `LSMW_*.txt` en salida/; al destruirse el frame, el `<Destroy>` bind cancela el polling para no dejar callbacks sueltos sobre widgets liberados) + `[Extraer Activos Creados]` (abre `abrir_extraer_creados(root, frame_activos)`) |
| **Extraer Activos Creados** (`abrir_extraer_creados`) | ← Atrás + logo + título "Extraer Activos Creados" | Form: campo `Usuario SAP` (Entry) + botón `[Ejecutar]`. Cableado a `_extraer_activos_creados_handler` que valida el input, pide confirmación, deshabilita Ejecutar+Atrás durante el worker, y ejecuta el flujo `extraer_activos_creados.extraer_activos_creados()` en thread daemon (SAP COM apartment inicializado vía `_sap_com_apartment()`). El flujo abre SM35P, filtra por CREATOR=`*<usuario>`, abre el primer log y exporta el .xlsx vía la cadena `tbar[0]/btn[86]` → `tbar[1]/btn[43]` → F4 → confirmaciones del recording `ScriptSM35P.vbs`. |
| **Control SOX intermedio** (`abrir_sox_menu`) | ← Atrás + logo + título "Control SOX" | `[HUB.PPE.01 Creación de Activos Fijos]` → al click abre `control_sox(root, frame_sox_menu)` (el formulario con Sociedad + Desde + Hasta queda como sub-sub-vista; doble back devuelve al menú principal) |
| **Form Control SOX clásico** (`control_sox`) | ← Atrás + logo + título "Control SOX" | Formulario Sociedad/Desde/Hasta + botón Generar. **Sin cambios** desde el refactor anterior (sigue funcionando con el mismo signature) |
| **Test conexión SAP** | `_test_conexion_sap_handler` (botón creado pero **NO empaquetado** — oculto en la UI). Se conserva el código para reactivarlo en el futuro sin re-implementarlo | Solo Windows |

### Helpers

- `_crear_header_form(root, frame, parent_frame, titulo, subtitulo=None)` — construye el encabezado consistente de cualquier sub-vista (botón ← Atrás cableado a `frame.destroy() + parent_frame.pack(...)` + logo reusado de `root._logo_ref` + título navy + subtítulo gris opcional). Devuelve el botón Atrás para que el caller pueda deshabilitarlo durante workers.

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

- **Validación previa `entrada/` (bloqueante):** antes de leer, `validar_entrada_unica()` verifica que haya **uno y solo un** `.xlsm` en `entrada/`. Si hay 2 o más, se muestra un `messagebox.showwarning` con `MENSAJE_ENTRADA_MULTIPLE` y **se aborta** la extracción (no se puede continuar hasta que quede un solo archivo). El archivo a leer se resuelve dinámicamente con `formato_dinamico_path()` (prefiere el nombre canónico `Formato_Dinamico.xlsm`, si no el primer `.xlsm` alfabético).
- Si existe(n) `LSMW_*.txt` previos → diálogo SÍ/NO `messagebox.askyesno`. SÍ borra todos los previos y genera uno nuevo; NO conserva.
- Sin previos → genera directamente `LSMW_YYYYMMDD_HHMMSS.txt`.
- Validaciones manejadas explícitamente: `FileNotFoundError` (Excel ausente), `ValueError` (hoja ausente), `Exception` (genérica del export), y red de seguridad que muestra traceback completo.

### Vista Control SOX (frame embebido, no Toplevel)

- **Patrón de switching de vistas:** `main()` envuelve todos los widgets del menú en un `frame_menu` (en vez de poner los widgets directo en `root`). Cuando el usuario presiona "Control SOX", `control_sox(root, frame_menu)` hace `frame_menu.pack_forget()` y muestra un nuevo `frame_sox` con el formulario + un botón "← Atrás". El click en Atrás destruye `frame_sox` y re-empaca `frame_menu`. El estado del menú (status_var, polling, flag `_upload_en_curso`) se preserva porque sólo se oculta, no se destruye.
- **Sociedades (multiselect)**: `tk.Listbox` con `selectmode="multiple"` (clic simple toggle-a, sin Ctrl) + `exportselection=False`. Opciones: `TRAN, ISA, ITCH, CEYA, CABA, RPAE, CTMP, REPD, ISAP, XM`. El usuario marca **1 o varias**; se genera un reporte por cada una (no consolidado). El helper `_sociedades_seleccionadas()` (expuesto en el frame) devuelve las marcadas.
- **Desde/Hasta**: `DateEntry` de tkcalendar con `date_pattern="dd.mm.yyyy"`. Validación per-keystroke (`validar_caracter_fecha`) acepta solo dígitos y puntos, máx 10 caracteres. Inicializa con la fecha actual.
- Validaciones al pulsar **Generar Reporte SOX**:
  1. Al menos una sociedad seleccionada (lista no vacía).
  2. Cada sociedad en lista permitida (normaliza con `.strip().upper()`).
  3. Ambas fechas formato `dd.mm.aaaa` válido.
  4. `Hasta >= Desde`.
- **Worker multi-sociedad con soft-fail**: el worker hace loop por cada sociedad seleccionada llamando `generar_reporte_sox(session, soc, ...)`; si una falla se registra y **continúa con las demás** (no aborta). El status muestra `Generando N/total: SOC…` y al final un messagebox con resumen `X OK / Y con error` (showinfo si todas OK, showerror si hubo alguna con error). El fallo de conexión SAP inicial sí aborta antes del loop.
- **Durante el worker SOX**: tanto el botón Generar como el Atrás se deshabilitan; ambos se re-habilitan al finalizar (éxito o error). El usuario no puede volver al menú a mitad de un flujo SAP. Se logra porque `_generar_reporte_sox_handler` recibe `(root, sociedades, ..., button, btn_atras)` y usa `root.after` (no el viejo `dialog.after`) para callbacks thread-safe.

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

## 6.5. Flujo Extraer Activos Creados — `src/extraer_activos_creados.py`

Replica `resources/ScriptSM35P.vbs`. Filtra el Monitor de Logs BDC (T-code **SM35P**) por un Usuario SAP, abre el primer log de la tabla y exporta el detalle a un archivo (default .xlsx según el recording).

### Constantes clave
- `T_CODE_SM35P = "sm35p"` — Monitor de sesiones BDC.
- `CAMPO_CREATOR = "wnd[0]/usr/subSCR_INFO:RSBDC_PROTOCOL:0201/txtD0100-CREATOR"` — campo filtro por usuario. Se setea con `*<usuario_sap>` (el `*` lo añade el código).
- `CELDA_PRIMER_REGISTRO = "wnd[0]/usr/tabsTAB_PROTOCOL/tabpALL_PROT/ssubSCR_CONTENT:RSBDC_PROTOCOL:0200/tblRSBDC_PROTOCOLTC_PROTOCOL/txtLIST_BDCLD-EDATE[0,0]"` — primera celda de la tabla; F2 sobre ella abre el detalle.
- `BTN_EXPORTAR_TBAR0 = "wnd[0]/tbar[0]/btn[86]"`, `BTN_EXPORTAR_TBAR1 = "wnd[0]/tbar[1]/btn[43]"` — cadena de exportación que abre el diálogo de save (wnd[1]).
- `CAMPO_DY_PATH = "wnd[1]/usr/ctxtDY_PATH"`, `CAMPO_DY_FILENAME = "wnd[1]/usr/ctxtDY_FILENAME"`, `BTN_CONFIRMAR_WND1 = "wnd[1]/tbar[0]/btn[11]"` — campos y botón del diálogo "Save list as file". Se setean directamente saltando el F4/picker del recording.
- `NOMBRE_PREFIJO = "ActivosCreados"`, `NOMBRE_EXTENSION = ".xlsx"` — patrón del nombre de archivo: `ActivosCreados_{USUARIO}_{YYYYMMDD_HHMMSS}.xlsx`.
- `LOGS_SHEET_NAME = "Logs"`, `ACTIVOS_FIJOS_SHEET_NAME = "Activos Fijos"`, `ACTIVOS_FIJOS_HEADERS = ("Activos Fijos", "Subnúmero")`, `HEADER_MENSAJE_LOG = "Mensaje de log"`, `COL_MENSAJE_LOG_DEFAULT = 2` — constantes del post-procesamiento del .xlsx.
- `PATRON_ACTIVO_LOG = re.compile(r"act\.\s*fj\.\s+(\d+)\s+(\d+)", re.IGNORECASE)` — regex que extrae `(activo_fijo, subnúmero)` de mensajes tipo `"El act.fj. 8048124 0 se ha creado"`. Tolera espacio opcional entre `act.` y `fj.` y es case-insensitive.

### Mapeo del flujo (5 etapas)

| # | Función | Acciones SAP / Python |
|---|---|---|
| 1 | `abrir_sm35p(session)` | maximize + okcd="sm35p" + Enter |
| 2 | `filtrar_por_usuario(session, usuario)` | `CREATOR.text = "*<usuario>"` + setFocus + caretPosition + Enter |
| 3 | `abrir_primer_registro(session)` | setFocus en celda `EDATE[0,0]` + caretPosition + F2 (sendVKey 2) |
| 4 | `exportar_log(session, carpeta, nombre)` | btn[86] → btn[43] → set `DY_PATH` y `DY_FILENAME` directamente en wnd[1] → btn[11]. **Variante del recording**: salta el F4/picker (wnd[2]) inyectando los campos en wnd[1] para forzar que el archivo caiga en `salida/` con el nombre estándar. |
| 5 | `procesar_logs(archivo_path)` (post-SAP, pure Python) | Abre el .xlsx generado; renombra la hoja única (Sheet1) → `Logs`; parsea la columna "Mensaje de log" con `PATRON_ACTIVO_LOG` y extrae todos los pares `(activo_fijo, subnúmero)`; deduplica preservando orden; crea (o reemplaza) la hoja `Activos Fijos` con headers en bold + datos como ints. Idempotente. Espera a que el archivo esté listo en disco vía `_esperar_archivo_listo` (SAP puede tardar 1-3s en cerrar el handle tras la exportación). |

### Helpers
- `validar_usuario_sap(usuario)` — acepta strings no-vacíos tras strip, sin transformar casing (IDs SAP pueden ser numéricos como `1017209574` o alfanuméricos como `INTC37089`).
- `_nombre_archivo_extraccion(usuario)` — construye `ActivosCreados_{USUARIO}_{YYYYMMDD_HHMMSS}.xlsx`.
- `_esperar_archivo_listo(archivo, timeout=10s, poll=0.5s)` — duplicado de `sox_report._esperar_archivo_listo`. Hace polling al `stat().st_size` y devuelve True cuando el tamaño se mantiene constante entre dos ticks (= file ya no está creciendo).
- `get_sap_session()` — idéntico a `sap_upload.get_sap_session()`.

### Path de salida y estructura del .xlsx final
- **Path forzado a `<PROJECT_ROOT>/salida/`** vía inyección de `DY_PATH` en el diálogo wnd[1]. El picker F4 del recording era una conveniencia del usuario para navegar; programáticamente no se necesita.
- Si SAP rechaza el `DY_PATH` directo (porque la transacción exige picker), `findById(CAMPO_DY_PATH)` falla con `RuntimeError` y el error indica el control que no se encontró — habría que añadir fallback con F4 + picker.
- **Estructura final del archivo** (tras paso 5 / `procesar_logs`):
  - Hoja `Logs`: copia 1:1 de lo que SAP exportó, sólo renombrada desde "Sheet1". Headers en fila 1 (`Hora de log`, `Mensaje de log`, `Cód.transacción`, ...).
  - Hoja `Activos Fijos`: 2 columnas (`Activos Fijos`, `Subnúmero`) en bold en fila 1, una fila por par único `(activo, sub)` extraído de los mensajes. Ints, sin formato custom.

### Limitaciones conocidas
- Los índices `btn[86]` / `btn[43]` son específicos de la pantalla de detalle SM35P y NO son estándar SAP — si la transacción cambia layout, hay que re-grabar.
- El regex `PATRON_ACTIVO_LOG` requiere `(activo) (subnumero)` separados por espacio (formato observado en SAP). Si en producción aparecen variantes (ej. `act.fj. 12345` sin subnúmero, o con `/` como separador), hay que ajustar el regex.

### CLI
```bash
python src/extraer_activos_creados.py 1017209574
```
Exit codes: 0 OK, 1 error (validación o SAP), 2 uso incorrecto.

## 6.6. Flujo Subir Anexos — `src/subir_anexos.py`

Replica `resources/Scriptanexo.vbs` (sin la navegación manual de carpetas que el usuario hizo durante la grabación — inyectamos `DY_PATH` directamente). Para cada par `(activo, subnúmero)` leído de la hoja `Activos Fijos` del último `ActivosCreados_*.xlsx` en `salida/`, sube cada uno de los archivos seleccionados como adjunto SAP vía **AS02** (Cambio Activo Fijo) + **GOS PCATTA_CREA** (Crear Adjunto del Object Services).

### Constantes clave
- `T_CODE_AS02 = "/nas02"` — Cambio Activo Fijo. El prefijo `/n` fuerza navegación a transacción **fresca** desde cualquier estado previo. Sin él, si una iteración anterior dejó SAP en la pantalla detalle (porque falló a media ejecución), un `okcd = "as02"` crudo no resetea y la siguiente iteración falla con `findById(ANLN1) → "control not found"`.
- `GOS_MENU_SETTLE_SECONDS = 0.3` — pausa entre `pressContextButton("%GOS_TOOLBOX")` y `selectContextMenuItem("%GOS_PCATTA_CREA")`. Sin esta pausa el `selectContextMenuItem` falla con `"method got an invalid argument"` porque el menú se construye async y el item solicitado aún no existe.
- `CAMPO_ANLN1 = "wnd[0]/usr/ctxtANLA-ANLN1"` — número de activo.
- `CAMPO_ANLN2 = "wnd[0]/usr/ctxtANLA-ANLN2"` — subnúmero.
- `CAMPO_BUKRS = "wnd[0]/usr/ctxtANLA-BUKRS"` — sociedad.
- `SHELL_TITULAR = "wnd[0]/titl/shellcont/shell"` — shell del menú GOS (toolbox del título).
- `GOS_TOOLBOX = "%GOS_TOOLBOX"`, `GOS_PCATTA_CREA = "%GOS_PCATTA_CREA"` — context buttons del menú GOS.
- `CAMPO_DY_PATH = "wnd[2]/usr/ctxtDY_PATH"`, `CAMPO_DY_FILENAME = "wnd[2]/usr/ctxtDY_FILENAME"` — diálogo del path. Inyectamos directo aquí en vez de navegar.
- `BTN_CONFIRMAR_WND1/2 = "wnd[X]/tbar[0]/btn[0]"` — botones de confirmación de la cascada de vuelta.

### Mapeo del flujo (7 etapas por (activo × archivo))

| # | Función | Acciones SAP |
|---|---|---|
| 1 | `adjuntar_archivo` | okcd="as02" + Enter (abre AS02) |
| 2 | id. | Set ANLN1 + ANLN2 + BUKRS + setFocus + caretPosition + Enter (carga el activo) |
| 3 | id. | `shell.pressContextButton("%GOS_TOOLBOX")` (abre menú GOS) |
| 4 | id. | `shell.selectContextMenuItem("%GOS_PCATTA_CREA")` (Crear adjunto) |
| 5 | id. | UN solo F4 en wnd[1] → abre wnd[2] (diálogo del path) |
| 6 | id. | Set `wnd[2]/DY_PATH = ruta absoluta del archivo` + `DY_FILENAME = ""` + setFocus + caretPosition |
| 7 | id. | btn[0] wnd[2] → btn[0] wnd[1] (cascada de confirmación) |

**Importante:** el recording original mostraba 2 F4s y profundidad hasta wnd[3] porque el usuario navegó manualmente por carpetas del file picker. El recording actualizado (sin navegación, path tipeado directo) confirmó que la cadena correcta es **1 F4 + wnd[2] + 2 btn[0]** — fue un fix crítico tras el primer test en SAP real.

### Helpers
- `validar_sociedad`, `VALID_SOCIEDADES` — importadas de `sox_report` (única fuente de verdad para la lista de sociedades).
- `get_archivo_activos_mas_reciente(salida_dir)` — busca `ActivosCreados_*.xlsx` más reciente por mtime. Lanza `FileNotFoundError` si no hay (típicamente porque el usuario no corrió "Extraer Activos Creados" antes).
- `leer_activos_del_excel(archivo_path)` — abre con openpyxl la hoja `Activos Fijos` (constante `ACTIVOS_FIJOS_SHEET_NAME` reusada de `extraer_activos_creados`) y devuelve `list[tuple[int, int]]` con `(activo, subnúmero)`. Salta filas no-int para robustez.
- `validar_y_leer_activos_usuario(archivo_path)` — valida y lee el **`.xlsx` que el usuario sube manualmente** con activos ya creados/existentes. Estructura EXIGIDA (validación estricta): extensión exclusivamente `.xlsx`, **una sola hoja**, **fila 1 = encabezado** (2 celdas de texto, obligatorio), datos desde fila 2 en **exactamente 2 columnas** numéricas (Activo Fijo, Subnúmero; tolera números como texto o floats enteros), al menos una fila de datos. Lanza `ValueError` con mensaje claro y corto (apto para `messagebox`) ante cualquier incumplimiento. Devuelve `list[tuple[int, int]]`. Helpers privados: `_celda_vacia`, `_ancho_fila`, `_a_entero`, `_parece_encabezado`.
- `get_sap_session()` — idéntico al patrón de otros módulos SAP.

### Orquestador — `subir_anexos(session, sociedad, archivos, archivo_activos=None, activos=None, progress_callback=None)`
- Resolución de la lista de activos por prioridad: **(1)** `activos` si se pasa directo (el `.xlsx` validado del usuario) → NO lee `salida/`; **(2)** hoja `Activos Fijos` de `archivo_activos`; **(3)** `ActivosCreados_*.xlsx` más reciente en `salida/` (default). Con `activos` vacío lanza `ValueError`.
- Loop `for activo in activos: for archivo in archivos: adjuntar_archivo(...)`.
- **Soft-fail por iteración**: si `adjuntar_archivo` lanza, se acumula en `detalles_fallos` y se continúa con la siguiente combinación. NO aborta.
- `progress_callback(intento, total, descripcion)` se llama antes de cada attachment para que el handler GUI actualice el status label (el callback se llama dentro de try/except para que un bug en la GUI no rompa el flujo SAP).
- Devuelve `{exitosos, fallidos, total_intentos, detalles_fallos: list[(activo, sub, archivo, error_msg)]}`.

### GUI — Vista "Subir Anexos" (`abrir_subir_anexos(root, frame_activos)`)
- Accesible desde **Activos Fijos** vía el botón `Subir Anexos` (tercer botón).
- Layout en dos secciones tituladas (navy bold, para no confundir los dos selectores de archivo): combo `Sociedad` (readonly) → sección **"Lista de activos (opcional)"** (`[Cargar .xlsx]` + `[Quitar]` + label de estado) → sección **"Anexos a subir"** (`[Seleccionar archivos]` abre `filedialog.askopenfilenames` + `[Quitar seleccionado]` + Listbox de archivos elegidos) → botón `[Subir Anexos a SAP]` + status label.
- **Lista de activos existentes (opcional):** debajo de Sociedad. Al seleccionar un `.xlsx`, se **valida y parsea inmediatamente** con `validar_y_leer_activos_usuario`; si cumple, la lista queda cacheada en `estado_usuario` (dict) y el label muestra `nombre.xlsx — N activo(s)` en verde; si no cumple, `messagebox.showwarning` con el feedback y no se acepta. `[Quitar]` revierte al comportamiento por defecto. Si NO hay archivo cargado, el flujo usa el último `ActivosCreados_*.xlsx` (igual que antes).
- Handler `_subir_anexos_handler(root, sociedad, archivos, activos_usuario, nombre_archivo_usuario, status_var, button, btn_atras)` — valida sociedad/archivos, arma el texto de confirmación según el origen (archivo del usuario vs. ActivosCreados), deshabilita Subir+Atrás, ejecuta worker en thread daemon envuelto en `_sap_com_apartment()`, llama `subir_anexos(..., activos=activos_usuario)` (si `None`, fallback al default), pasa `progress_callback` que actualiza `status_var` vía `root.after`. Al final muestra messagebox con resumen `X OK / Y fallos` (showinfo si todos OK, showerror si hubo fallos).

### Limitaciones conocidas
- La cascada wnd[1]→wnd[2]→wnd[3] depende de que el GOS PCATTA_CREA abra exactamente esa profundidad de diálogos. Si la versión de SAP del cliente abre menos o más niveles, hay que ajustar el número de F4 / botones de confirmación.
- Cada attachment ejecuta el ciclo completo AS02 + GOS + cascada — esto es lento (~5-10s por iteración). Para subir, ej., 50 activos × 3 archivos = 150 iteraciones ≈ 15-25 minutos. NO interactuar con SAP durante el proceso.

## 7. Datos: Hoja LSMW del Excel

- La hoja `LSMW ` (con espacio) está cableada con fórmulas que referencian la hoja `Formato`.
- `openpyxl` lee con `data_only=True` los valores **cacheados** por Excel en el último guardado → si el usuario edita el formulario, **debe abrir y guardar el Excel** antes de extraer para que las fórmulas se recalculen.
- Celdas vacías referenciadas pueden aparecer como `0`.
- 51 columnas exportadas — algunos campos: `ANLKL` (clase de activo), `BUKRS` (sociedad), `TXT50` (denominación), `KOSTL` (centro de costo), `WERKS` (centro), `EAUFN` (orden de inversión), `POSNR` (elemento PEP), `ORD41`–`ORD44` y `GDLGRP` (criterios de clasificación 1–5).

## 8. Pruebas — 378 tests con `unittest`

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
- `tests/test_main.py` (105): handlers GUI, extracción TSV (incluye `test_warns_and_aborts_when_multiple_xlsm_in_entrada` — advertencia bloqueante ante múltiples `.xlsm` en `entrada/`), vista SOX como frame embebido (`ControlSoxDialogTest` incluye el multiselect de sociedades — Listbox `selectmode="multiple"`, helper `sociedades_seleccionadas`, presencia de `XM` — y `test_date_entries_do_not_use_key_validation`, regresión del bug del calendario donde `validate="key"` rompía el `_select` del popup), `GenerarReporteSoxHandlerTest` (deshabilitado de Generar+Atrás durante worker, **multiselect: un reporte por sociedad, soft-fail que continúa con las demás, lista vacía → error**), `FooterCopyrightTest` (3 tests: año actual, estilo discreto, packed `side="bottom"`), `CerrarSplashTest` (no-op silencioso de `_cerrar_splash` en dev mode).
- `tests/test_paths.py` (18): helpers de modo dev/bundled (PyInstaller). `ProjectRootTest` (`sys.frozen` vs dev), `BundledResourcePathTest` (lectura de `sys._MEIPASS`), `ListarXlsmEntradaTest` (glob de `.xlsm` en `entrada/`, ignora otras extensiones), `FormatoDinamicoPathTest` (resolución: prefiere canónico, si no el primero, si vacío devuelve canónico), `ValidarEntradaUnicaTest` (0/1 ok, ≥2 advierte con `MENSAJE_ENTRADA_MULTIPLE`), `AsegurarFormatoDinamicoTest` (factory default `Formato_Dinamico.xlsm`: primer arranque copia, segundo preserva ediciones, no copia si el usuario dejó un `.xlsm` con otro nombre, bundle ausente devuelve path sin crashear).
- `tests/test_branding.py` (15): paleta corporativa Hub de ISA, `cargar_logo` (escalado por aspect ratio, referencias persistentes), `aplicar_estilo_primario`/`aplicar_estilo_terciario`.
- `tests/test_sox_report.py` (105): validaciones puras + pasos del flujo SAP + `GenerarXlsxPoblacionTest` (11 tests del paso Población) + `PatronAfRegexTest` (6 tests del parseo de col D) + `GenerarHojaCreadosTest` (16 tests del paso post-procesamiento: filter, parsing, estructura, K y L como fórmulas =MID/=IF, replace idempotente, errores) + `EjecutarReporteTest` (2 tests del split de F8) + `GenerarHojaIpeTest` (8 tests de la hoja de evidencias: crear, embedding, soft-fail, replace, scaling) + `GenerarReporteSoxTest` (verifica orden de las 7 etapas incluyendo screenshots, paso a las funciones, `EXPORT_METHOD=None` salta Población/Creados/IPE).
- `tests/test_sap_upload.py` (46): cada paso del flujo LSMW + `MainEntryPointTest`.
- `tests/test_extraer_activos_creados.py` (48): cada paso del flujo SM35P + post-procesamiento del .xlsx.
- `tests/test_subir_anexos.py` (41): cada paso del flujo AS02+GOS PCATTA_CREA + orquestador soft-fail + `ValidarYLeerActivosUsuarioTest` (10 tests de la validación del `.xlsx` del usuario: válido, extensión, ≥2 hojas, >2 columnas, sin encabezado, datos no numéricos, solo encabezado, números como texto, floats enteros, filas vacías al final) + override `activos=` en el orquestador (prioridad sobre `salida/`, vacío → error).

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
- **Carpeta externa `entrada/` (no `resources/`) para el input del usuario** — el `Formato_Dinamico.xlsm` editable vive en `<EXE_DIR>/entrada/`, separada de la `resources/` interna del proyecto (logo, factory default, VBS) que se bundlea read-only. Sin esta separación el usuario veía dos carpetas conceptualmente distintas llamadas `resources` (la del repo/bundle vs. la externa editable) y era confuso. El factory default sigue viviendo en `resources/` y se copia a `entrada/` en el primer arranque.
- **El input es `.xlsm` (con macros), no `.xlsx`** — el nombre canónico es `Formato_Dinamico.xlsm` (constante `FORMATO_DINAMICO_NOMBRE`). openpyxl lee `.xlsm` sin problema con `load_workbook(..., data_only=True)` (emite un `UserWarning` inofensivo sobre "Data Validation extension"). El nombre no es configurable desde la UI.
- **Regla "un y solo un `.xlsm` en `entrada/`" bloqueante** — con 2+ archivos la elección del que se lee es ambigua, así que `validar_entrada_unica()` corre ANTES de leer y, si hay conflicto, muestra `MENSAJE_ENTRADA_MULTIPLE` y **aborta** la extracción (no advierte-y-sigue: no deja continuar hasta que quede un solo archivo). Además `asegurar_formato_dinamico()` no copia el factory default si ya hay algún `.xlsm` (aunque esté renombrado), para no auto-provocar el conflicto. La resolución (`formato_dinamico_path()`) y la validación son **dinámicas** (se evalúan en cada click de Extraer, no en el import de `main.py`), así el usuario puede corregir la carpeta sin reiniciar la app.
- **Subir Anexos: origen de activos con override opcional** — el usuario puede subir su propio `.xlsx` de activos ya creados/existentes en vez de usar el `ActivosCreados_*.xlsx` de "Extraer Activos Creados". El override se resuelve pasando la lista **ya parseada** (no la ruta) al orquestador vía `activos=`, que tiene prioridad y evita releer `salida/`. Se valida y parsea **al seleccionar** (no al subir): así el click de "Subir" es inmediato, no re-valida, y el error se muestra en el momento en que el usuario puede corregirlo. La validación es **estricta y bloqueante** (extensión `.xlsx` exclusiva, una sola hoja, encabezado obligatorio, exactamente 2 columnas numéricas) porque un archivo mal formado dispararía adjuntos a activos equivocados en SAP — un error caro y difícil de revertir. Si no hay archivo cargado, el comportamiento por defecto (leer `salida/`) queda intacto.
