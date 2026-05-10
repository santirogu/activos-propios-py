# Creación de Activos Fijos SAP

Aplicación de escritorio en Python que automatiza dos pasos del proceso de creación y capitalización de activos fijos en SAP:

1. **Extracción** de la hoja `LSMW` del formato dinámico Excel a un `.txt` separado por tabulación.
2. **Carga** de ese `.txt` a SAP vía la transacción LSMW, ejecutando el flujo completo (Specify Files → Read Data → Convert Data → Create BI Session → Run BI) mediante **SAP GUI Scripting**.

## Diagrama del proceso

![Diagrama de flujo del proceso](docs/flujo-proceso.png)

El proceso completo contempla:

1. **INICIO** — el usuario diligencia el formulario de creación en el archivo Excel maestro (`Formato_Dinamico_.xlsx`).
2. **Formulario de creación** — captura de los datos del activo en la hoja `Formato`.
3. **Extraer hoja LSMW en un `.txt`** — *automatizado por el botón "Extraer información en txt"*.
4. **Login en SAP** — paso manual realizado por el usuario.
5. **Correr `script.py`** — *automatizado por el botón "Subir a SAP"* (ejecuta `src/sap_upload.py`).
6. **Generar reporte** — log de la sesión BDC visible en la transacción SM35.
7. **FIN**.

> Esta aplicación cubre los **pasos 3 y 5** del flujo. La autenticación en SAP (paso 4) sigue siendo manual.

## Requerimientos

- **Python 3.9 o superior** con soporte para Tkinter (incluido por defecto; en macOS, el Python de Homebrew 3.12 **no** trae Tk — usa `python.org` o el del sistema).
- **openpyxl** — manipulación del Excel (multiplataforma).
- **pywin32** — solo necesario para el botón "Subir a SAP". Se instala automáticamente solo en Windows gracias al marcador `platform_system == "Windows"` en `requirements.txt`.
- El archivo `resources/Formato_Dinamico_.xlsx` debe existir.
- Para subir a SAP: SAP GUI for Windows abierto con sesión iniciada y `sapgui/user_scripting = TRUE` (ver sección "Configuración SAP" más abajo).

## Quick start

```bash
# 1. Clonar el repositorio
git clone https://github.com/santirogu/activos-propios-py.git
cd activos-propios-py

# 2. (Recomendado) Crear y activar un entorno virtual
python3 -m venv .venv
source .venv/bin/activate          # macOS / Linux
# .venv\Scripts\activate            # Windows

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar la app
python src/main.py
```

La ventana muestra dos botones:

- **Extraer información en txt** — funciona en cualquier sistema operativo.
- **Subir a SAP** — solo funciona en Windows con SAP GUI configurado.

## Cómo ejecutar la app

```bash
python src/main.py
```

### Botón "Extraer información en txt"

- Lee la hoja `LSMW` de `resources/Formato_Dinamico_.xlsx`.
- Crea la carpeta `salida/` en la raíz si no existe.
- Genera un archivo TSV con el patrón `LSMW_YYYYMMDD_HHMMSS.txt`.
- Muestra confirmación con la cantidad de filas exportadas.

### Botón "Subir a SAP"

- Pide confirmación antes de ejecutar (operación sensible que toma control de SAP).
- Toma el `.txt` más reciente de `salida/`.
- Conecta a la sesión SAP abierta vía SAP GUI Scripting (COM).
- Ejecuta el flujo LSMW completo: **configura dinámicamente la ruta del archivo en "Specify Files"** → Assign Files → Read Data → Display Read Data → Convert Data → Display Converted Data → Create Batch Input Session → Run Batch Input Session → Process BDC Session.
- El flujo corre en un hilo background, así la GUI no se congela.
- El status label muestra el progreso paso a paso.
- Al terminar, mostrar mensaje y sugerir revisar SM35.

También se puede ejecutar la carga sin GUI:

```bash
python src/sap_upload.py
```

### Notas sobre los datos exportados

La hoja `LSMW` está cableada con fórmulas que referencian la hoja `Formato`. `openpyxl` lee los valores que Excel **dejó cacheados** en el último guardado, por lo tanto:

- Si después de modificar el Excel quieres ver los nuevos valores en el TXT, **abre y guarda el Excel** antes de ejecutar la app (Excel recalcula y cachea las fórmulas al guardar).
- Las celdas referenciadas que estén vacías pueden aparecer como `0` (comportamiento estándar de Excel para referencias numéricas a celdas vacías).

## Configuración SAP (una sola vez por máquina)

Para que el botón "Subir a SAP" funcione:

1. **Cliente** — habilitar scripting en SAP GUI:
   *Options → Accessibility & Scripting → Scripting → "Enable scripting"*. Recomendado desmarcar los dos "Notify when..." para experiencia desatendida.
2. **Servidor** — parámetro `sapgui/user_scripting = TRUE` (transacción RZ11). Si no está habilitado, pídele al equipo Basis que lo active.
3. **Iniciar sesión SAP** antes de presionar el botón. El script no autentica.
4. **Pre-cargar el proyecto LSMW** — abrir LSMW manualmente al menos una vez con Subproject + Object correctos. SAP recuerda la última selección.

> **La ruta del archivo en LSMW ya no requiere configuración manual.** El script ahora la inyecta dinámicamente en cada corrida apuntando al `.txt` más reciente de `salida/`, replicando la grabación VBS de `resources/Script1.vbs`.

## Cómo ejecutar las pruebas

Las pruebas usan `unittest` (incluido en la librería estándar, sin dependencias adicionales).

```bash
# Toda la suite
python -m unittest discover tests -v

# Solo el módulo principal
python -m unittest tests.test_main -v

# Solo el módulo de carga SAP
python -m unittest tests.test_sap_upload -v

# Un test específico
python -m unittest tests.test_main.SubirASapTest.test_worker_calls_full_flow_on_happy_path
```

### Cobertura de pruebas

La suite contiene **57 pruebas** distribuidas en dos archivos:

#### `tests/test_main.py` (21 pruebas)

**`ExportSheetToTsvTest`** (9 pruebas) — lógica pura de extracción TSV: contenido tab-separated, manejo de `None`, creación de directorios, patrón de timestamp, prefijo configurable, errores de archivo/hoja faltantes, contador de filas, no-overwrite por timestamp.

**`RealWorkbookSmokeTest`** (1 prueba) — smoke test contra el Excel real del proyecto.

**`SubirASapTest`** (11 pruebas) — handler del botón "Subir a SAP":

| Test | Qué valida |
|---|---|
| `test_cancel_confirmation_does_not_start_thread` | Cancelar el diálogo no lanza el worker |
| `test_cancel_does_not_modify_status` | Cancelar no toca `status_var` |
| `test_confirmation_disables_button_before_starting_worker` | Botón deshabilitado antes del thread |
| `test_worker_calls_full_flow_on_happy_path` | `get_latest_txt` + `get_sap_session` + `run_lsmw_flow(session, carpeta, nombre)` |
| `test_worker_reenables_button_after_success` | Tras éxito el botón vuelve a `normal` |
| `test_worker_updates_status_to_completion_message` | Status final contiene "completada" |
| `test_worker_passes_folder_and_filename_to_run_lsmw_flow` | Carpeta y nombre del .txt llegan correctos al flujo |
| `test_worker_handles_missing_txt` | `FileNotFoundError` → error, botón se reactiva |
| `test_worker_handles_sap_connection_error` | `RuntimeError` SAP → error, botón se reactiva |
| `test_worker_handles_lsmw_flow_error` | Excepción del flujo → error, NO muestra info de éxito |
| `test_worker_resets_status_on_error` | `status_var` se vacía tras error |

#### `tests/test_sap_upload.py` (36 pruebas)

| Clase | Tests | Cobertura |
|---|---|---|
| `GetLatestTxtTest` | 4 | Directorio faltante, sin archivos, mtime más reciente, ignora otros patrones |
| `GetSapSessionTest` | 5 | pywin32 ausente, SAP no corre, sin conexiones, sin sesiones, devuelve sesión OK |
| `OpenLsmwTest` | 2 | maximize + okcd + sendVKey + btn[8], orden correcto |
| `SelectStepRowTest` | 3 | Deselecciona default, selecciona target, foco en celda |
| `ConfigurarRutaArchivoTest` | 7 | Replica `Script1.vbs`: F2 al paso, btn[25]/btn[27], lbl[43,6], F4 al picker, set path/filename, OK + Back + SPOP-OPTION1, secuencia correcta |
| `StepAssignFilesTest` | 1 | Row 7 + btn[32] + sendVKey(3) |
| `StepReadDataTest` | 1 | Row 8 + btn[32] + btn[8] + 2× sendVKey(3) |
| `StepDisplayReadDataTest` | 1 | btn[32] + popup confirm + back |
| `StepConvertDataTest` | 1 | btn[32] + sendVKey(8) + 2× sendVKey(3) |
| `StepDisplayConvertedDataTest` | 1 | btn[32] + popup confirm + back |
| `StepCreateBatchInputTest` | 1 | btn[32] + chkP_KEEP + btn[8] + popup |
| `StepRunBatchInputTest` | 1 | Solo btn[32] |
| `ProcessBdcSessionTest` | 1 | Tabla BDC + modo error + log all + expert + 2× OK |
| `RunLsmwFlowTest` | 2 | Orden completo de los 10 pasos, `configurar_ruta_archivo` recibe (carpeta, nombre) |
| `MainEntryPointTest` | 4 | Exit code 0/1 según escenario; pasa carpeta y nombre del `.txt` al flujo |

**Estrategia de mocking SAP**: `MockSAPSession` registra cada llamada `findById(...).method()` en `session.actions` como tuplas `(sap_id, method, *args)` y expone los elementos vía `session._elements[id]` para inspeccionar propiedades (`text`, `selected`, `caretPosition`). Las filas de tablas usan `_MockRow` con setter que loguea cambios de `selected`. Esto permite verificar la secuencia exacta de IDs y métodos SAP sin necesidad de un sistema SAP real.

**Estrategia de mocking GUI**: `_SyncFakeThread` reemplaza `threading.Thread` para ejecutar el worker síncrono; `root.after` se sobreescribe en `setUp` para invocar callbacks inmediatamente. `patch.multiple("sap_upload", ...)` inyecta los mocks de las funciones del módulo; los mocks se guardan en `self.mocks` para verificación.

## Estructura del proyecto

```
.
├── src/
│   ├── main.py                      # App GUI: 2 botones (extraer + subir a SAP)
│   └── sap_upload.py                # Lógica de carga LSMW vía SAP GUI Scripting
├── tests/
│   ├── test_main.py                 # 22 pruebas: extracción + botón SAP
│   └── test_sap_upload.py           # 32 pruebas: flujo LSMW completo
├── resources/
│   ├── Formato_Dinamico_.xlsx       # Formato maestro con catálogos y plantilla
│   └── script_sap_base.txt          # Grabación VBS de referencia (UTF-16)
├── docs/
│   └── flujo-proceso.png            # Diagrama del proceso completo
├── salida/                          # Carpeta generada con los .txt exportados
├── requirements.txt                 # openpyxl + pywin32 (Windows only)
└── README.md                        # Este archivo
```

## Arquitectura del código

### `src/main.py`

- **`export_sheet_to_tsv(excel_path, sheet_name, output_dir, file_prefix="LSMW")`** — función pura que realiza la extracción y devuelve `(ruta_archivo, filas_escritas)`. Lanza `FileNotFoundError` / `ValueError`. Es la pieza testeable de la extracción.
- **`extraer_lsmw_a_txt(status_var)`** — wrapper GUI del botón "Extraer", traduce excepciones a `messagebox`.
- **`subir_a_sap(root, status_var, button)`** — handler del botón "Subir a SAP". Pide confirmación, deshabilita el botón, lanza un hilo background que invoca las funciones de `sap_upload` y reporta progreso vía `root.after()` (thread-safe en Tkinter). El import de `sap_upload` es lazy dentro del worker para que `main.py` arranque sin pywin32 instalado.

### `src/sap_upload.py`

Replica los pasos grabados en dos VBS de SAP:
- `resources/script_sap_base.txt` — flujo LSMW completo (Read Data, Convert, Create BI, Run BI, BDC processing).
- `resources/Script1.vbs` — configuración dinámica del archivo de entrada en el paso "Specify Files".

Cada paso está en una función dedicada (`open_lsmw`, `configurar_ruta_archivo`, `step_assign_files`, `step_read_data`, `step_display_read_data`, `step_convert_data`, `step_display_converted_data`, `step_create_batch_input`, `step_run_batch_input`, `process_bdc_session`). El orquestador `run_lsmw_flow(session, carpeta, nombre_archivo)` los llama en secuencia inyectando la ruta del .txt.

Funciones de soporte:
- **`get_latest_txt(salida_dir)`** — devuelve el `LSMW_*.txt` más reciente por mtime.
- **`get_sap_session()`** — conecta al SAP GUI Scripting Engine vía `win32com.client` (importado lazy). Lanza `RuntimeError` con mensajes claros si pywin32 no está instalado, SAP GUI no corre, o no hay conexión/sesión activa.

Esta separación granular permite testear cada paso de forma aislada con un `MockSAPSession`.

## Mapeo del flujo LSMW

| Paso del proyecto | Fila step list | Función Python | Acciones SAP |
|---|---|---|---|
| Specify Files (configura ruta dinámica) | 6 | `configurar_ruta_archivo(session, carpeta, nombre)` | F2 + btn[25] + lbl[43,6] + btn[27] + F4 + DY_PATH/DY_FILENAME + 2×OK + Back + SPOP-OPTION1 |
| Assign Files | 7 | `step_assign_files` | btn[32] + VK3 |
| Read Data | 8 | `step_read_data` | btn[32] + btn[8] + 2×VK3 |
| Display Read Data | (auto-avanza) | `step_display_read_data` | btn[32] + popup + VK3 |
| Convert Data | (auto-avanza) | `step_convert_data` | btn[32] + VK8 + 2×VK3 |
| Display Converted Data | (auto-avanza) | `step_display_converted_data` | btn[32] + popup + VK3 |
| Create Batch Input Session | (auto-avanza) | `step_create_batch_input` | btn[32] + chkP_KEEP + btn[8] + popup |
| Run Batch Input Session | (auto-avanza) | `step_run_batch_input` | btn[32] |
| Procesar BDC Session | (en SM35-like) | `process_bdc_session` | row[0] + btn[8] + radError + chkLOGALL/EXPERT + 2×OK |

## Hoja LSMW: contenido exportado

La hoja `LSMW` mapea las columnas del formulario a los **nombres técnicos de campos SAP**. El TXT generado contiene 51 columnas con campos como:

- `ANLKL` (Clase de activo fijo)
- `BUKRS` (Sociedad)
- `TXT50` (Denominación del activo fijo)
- `KOSTL` (Centro de costo)
- `WERKS` (Centro)
- `EAUFN` (Orden de inversión)
- `POSNR` (Elemento PEP)
- `ORD41`–`ORD44`, `GDLGRP` (Criterios de clasificación 1–5)
- entre otros.

## Diagnóstico de errores comunes en la carga SAP

| Error | Causa probable | Solución |
|---|---|---|
| "No se pudo conectar a SAP GUI" | SAP no abierto o scripting deshabilitado | Abrir SAP GUI, habilitar scripting en Options |
| "No hay sesiones activas" | Estás en la pantalla de login | Iniciar sesión en el sistema SAP |
| "Falta la dependencia pywin32" | Estás en Mac/Linux o no instalaste deps | `pip install pywin32` (solo Windows) |
| Falla en `select_step_row` | Proyecto LSMW incorrecto pre-cargado | Abrir LSMW manualmente con el proyecto correcto |
| Falla en `configurar_ruta_archivo` | El proyecto LSMW tiene la definición de archivo en otra posición | Re-grabar `Script1.vbs` con tu proyecto y ajustar IDs (`lbl[43,6]`, `btn[25]`, `btn[27]`) |
| Falla en `step_read_data` | El archivo no existe en la ruta inyectada o no tiene permisos | Verifica que `salida/<archivo>` exista y SAP tenga acceso al disco |
