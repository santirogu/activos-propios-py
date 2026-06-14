---
name: convertir-vbs
description: Convierte una grabación VBS de SAP GUI Scripting (en resources/) en un módulo Python siguiendo las convenciones del proyecto (constantes, _ejecutar wrapper, MockSAPSession en tests, orquestador con stats dict, CLI con exit codes).
---

# convertir-vbs

Cuando el usuario invoque esta skill, normalmente con argumentos como `resources/ScriptX.vbs` y un nombre de módulo (ej. `subir_anexos`), genera el módulo Python equivalente siguiendo el patrón establecido en `src/sap_upload.py`, `src/sox_report.py`, `src/extraer_activos_creados.py` y `src/subir_anexos.py`.

## Pasos

### 1. Validar inputs

- Confirma que el archivo VBS existe en `resources/`.
- Pide al usuario un nombre de módulo si no lo dio (snake_case, sin extensión).
- Verifica que `src/<nombre_modulo>.py` no exista todavía (si existe, pregunta si reemplazar).

### 2. Leer y entender el VBS

Los VBS están en **UTF-16LE con CRLF**. Léelos con:

```bash
iconv -f UTF-16LE -t UTF-8 resources/ScriptX.vbs
```

Identifica:
- **T-code** del primer `okcd.text = "..."` (ej. `as02`, `sm35p`, `ar15`).
- **Campos del header** (textos en `wnd[0]/usr/ctxt*` o `wnd[0]/usr/txt*`).
- **Botones de toolbar** (`pressButton`, `pressContextButton`, `selectContextMenuItem`).
- **Diálogos** (`wnd[1]`, `wnd[2]`, etc.) con sus campos `DY_PATH`, `DY_FILENAME`, etc.
- **Confirmaciones** (`btn[0]`, `btn[11]`, etc.).
- **Setattr de `text`, `setFocus`, `caretPosition`, `sendVKey`**.

### 3. Generar `src/<nombre_modulo>.py`

Sigue el patrón EXACTO de [`src/subir_anexos.py`](src/subir_anexos.py) o [`src/extraer_activos_creados.py`](src/extraer_activos_creados.py):

```python
"""<nombre_modulo>.py — <descripción breve del flujo SAP>.

Replica `resources/ScriptX.vbs`. <Resumen 1-2 líneas de qué hace>.

REQUISITOS
==========
Windows con SAP GUI abierto y sesión iniciada.

USO CLI
=======
    python src/<nombre_modulo>.py <ARG1> [<ARG2>]
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

# Imports adicionales según el flujo:
# - openpyxl si lee/escribe Excel
# - sox_report para reutilizar VALID_SOCIEDADES + validar_sociedad
# - extraer_activos_creados para constantes compartidas

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SALIDA_DIR = PROJECT_ROOT / "salida"

# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------

# T-code (con prefijo "/n" para forzar transacción fresca entre iteraciones,
# importante cuando el mismo flujo se ejecuta varias veces y SAP podría
# quedar en otra pantalla)
T_CODE_X = "/n<tcode>"

# Constantes para cada campo y botón del VBS, con nombres descriptivos
CAMPO_X = "wnd[0]/usr/ctxt..."
SHELL_X = "wnd[0]/titl/shellcont/shell"
BTN_X = "wnd[1]/tbar[0]/btn[0]"


# ---------------------------------------------------------------------------
# LOGGING (idéntico en todos los módulos SAP)
# ---------------------------------------------------------------------------

def _log(mensaje: str) -> None:
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {mensaje}", flush=True)


def _ejecutar(descripcion: str, fn, *args, **kwargs):
    """Wrapper que loguea y re-lanza con contexto si falla."""
    _log(f"  → {descripcion}")
    try:
        return fn(*args, **kwargs)
    except Exception as exc:
        raise RuntimeError(
            f"Falló: {descripcion}\nDetalle técnico SAP: {exc!r}"
        ) from exc


# ---------------------------------------------------------------------------
# VALIDACIONES
# ---------------------------------------------------------------------------

def validar_<input>(valor: str) -> str:
    """Valida y normaliza el input del usuario."""
    if not isinstance(valor, str) or not valor.strip():
        raise ValueError("Debes ingresar <campo>.")
    return valor.strip()


# ---------------------------------------------------------------------------
# CONEXIÓN A SAP (copiar literal de sap_upload.get_sap_session)
# ---------------------------------------------------------------------------

def get_sap_session():
    """Conecta al SAP GUI Scripting Engine."""
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Falta la dependencia pywin32. Instalar con: pip install pywin32"
        ) from exc

    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
    except Exception as exc:
        raise RuntimeError(
            "No se pudo conectar a SAP GUI. Verifica:\n"
            "  - SAP GUI for Windows está abierto y con sesión iniciada.\n"
            "  - SAP GUI Scripting habilitado en Options del cliente.\n"
            "  - sapgui/user_scripting = TRUE en el servidor SAP."
        ) from exc

    application = sap_gui_auto.GetScriptingEngine
    if application.Children.Count == 0:
        raise RuntimeError("No hay conexiones SAP activas en este SAP GUI.")
    connection = application.Children(0)
    if connection.Children.Count == 0:
        raise RuntimeError(
            "No hay sesiones activas. Inicia sesión en SAP antes de correr."
        )
    return connection.Children(0)


# ---------------------------------------------------------------------------
# FLUJO SAP — cada función representa un paso lógico del .vbs
# ---------------------------------------------------------------------------

def paso_uno(session, ...):
    """<Descripción del paso>. Replica líneas X-Y del .vbs."""
    _log("Paso 1/N: <descripción>...")
    # Usar findById directo dentro de lambdas — NO cachear referencias.
    # SAP COM puede invalidar las referencias entre llamadas.
    _ejecutar(
        f"<descripción técnica>",
        lambda: session.findById(SHELL_X).pressButton(BTN_ID),
    )


def paso_dos(session, ...):
    ...


# ---------------------------------------------------------------------------
# ORQUESTADOR (con soft-fail si itera sobre múltiples elementos)
# ---------------------------------------------------------------------------

def <nombre_flujo>(session, <args>, progress_callback=None) -> dict | tuple:
    """Ejecuta el flujo completo. <Soft-fail por iteración si aplica>."""
    # Validar inputs
    # Inicializar contadores si soft-fail
    inicio = time.monotonic()
    _log("=== Iniciando <nombre> ===")

    # Llamar a los pasos en orden
    paso_uno(session, ...)
    paso_dos(session, ...)
    # ...

    duracion = time.monotonic() - inicio
    _log(f"=== Finalizado en {duracion:.1f}s ===")
    return ...  # tupla (carpeta, nombre) si genera archivo, o dict con stats si itera


# ---------------------------------------------------------------------------
# ENTRY POINT CLI
# ---------------------------------------------------------------------------

def main(argv=None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    print("=" * 70, flush=True)
    print("<Título del flujo>", flush=True)
    print("=" * 70, flush=True)

    if len(argv) < <N>:
        print(
            "Uso: python src/<nombre_modulo>.py <args>",
            file=sys.stderr,
        )
        return 2

    try:
        # Validar inputs
        pass
    except ValueError as exc:
        print(f"ERROR de validación: {exc}", file=sys.stderr)
        return 1

    try:
        session = get_sap_session()
    except RuntimeError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    try:
        resultado = <nombre_flujo>(session, ...)
    except Exception as exc:
        print(f"\nERROR durante el flujo: {exc}", file=sys.stderr)
        return 1

    print("=" * 70, flush=True)
    print(f"<resumen>", flush=True)
    print("=" * 70, flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

### Convenciones IMPORTANTES (aprendidas a la mala en este proyecto)

1. **NO cachear referencias de findById en variables**. SAP COM las invalida entre llamadas. Usa siempre `lambda: session.findById(X).method(...)` dentro del `_ejecutar`.

2. **NO añadir sleeps** entre llamadas del VBS (rompen el flujo, ej. menús contextuales se cierran solos).

3. **Match línea por línea con el VBS**. Cada `_ejecutar` debe corresponder a una línea del VBS. Comentar el número de línea ayuda a debuggear.

4. **Prefijo `/n` en T-codes** para forzar transacción fresca (`"/nas02"`, `"/nsm35p"`).

5. **ANLN2 y otros campos opcionales**: si el VBS no los setea, NO los setees (puedes ponerlos opcionales si la lógica lo requiere). Ejemplo: en `subir_anexos.py`, `ANLN2` solo se setea si `subnumero != 0`.

6. **`pressButton` vs `pressContextButton`** son métodos DIFERENTES. Lee el VBS con cuidado.

### 4. Generar `tests/test_<nombre_modulo>.py`

Sigue el patrón de [`tests/test_subir_anexos.py`](tests/test_subir_anexos.py) o [`tests/test_extraer_activos_creados.py`](tests/test_extraer_activos_creados.py):

- Copia la clase `MockSAPSession` y `_MockElement` (incluyendo `pressButton` y `pressContextButton` si el flujo los usa).
- Un test class por cada paso del flujo.
- Test class para el orquestador con mocks de los pasos individuales (verificando orden de llamadas).
- `MainEntryPointTest` para los exit codes 0/1/2.

### 5. Verificar

```bash
.venv/bin/python -c "import ast; ast.parse(open('src/<nombre_modulo>.py').read()); print('OK')"
.venv/bin/python -m unittest tests.test_<nombre_modulo>
```

### 6. Sugerir siguientes pasos al usuario

- Wirear el módulo a la GUI: invoca `agregar-vista-gui` si necesita interfaz.
- Sincronizar docs: invoca `sincronizar-docs` para actualizar CLAUDE.md/README.md.

## Qué NO hacer

- NO inventar IDs SAP que no estén en el VBS — si falta uno, pregúntale al usuario por una grabación actualizada.
- NO añadir lógica de retry/sleep "por si acaso" — el patrón del proyecto es match 1:1 con el VBS.
- NO cambiar el orden de las acciones del VBS — el orden importa (focus chain, auto-tab de SAP).
