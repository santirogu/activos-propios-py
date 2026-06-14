---
name: correr-tests
description: Corre el suite de tests del proyecto teniendo en cuenta la quirk del entorno macOS dev (Python 3.14 Homebrew sin tkinter → test_main.py no corre, solo valida sintaxis). Reporta resultados claros con conteos.
---

# correr-tests

Cuando el usuario invoque esta skill, corre los tests apropiados según el entorno donde estás. El proyecto tiene una particularidad: el Python local del usuario (Homebrew 3.14 en macOS) **NO trae tkinter**, así que `tests/test_main.py` no puede ejecutarse localmente — solo se valida su sintaxis con `ast.parse`.

## Detectar el entorno

```bash
# 1. Detectar venv del proyecto
test -x .venv/bin/python && PY=".venv/bin/python" || PY="python3"

# 2. Detectar si tkinter está disponible
$PY -c "import tkinter" 2>/dev/null && TK=1 || TK=0
```

## Estrategia según disponibilidad de Tk

### Caso A: Tk disponible (Windows, o macOS con python.org)

Corre el suite completo:

```bash
$PY -m unittest discover tests -v
```

Reporta el total de tests ejecutados, número de fallos, número de errores.

### Caso B: Tk NO disponible (caso típico en macOS dev)

Corre todos los suites EXCEPTO `test_main`:

```bash
$PY -m unittest tests.test_sap_upload tests.test_sox_report tests.test_extraer_activos_creados tests.test_subir_anexos
```

Luego valida que `test_main.py` parsea (no que pasa, solo que no tiene syntax errors):

```bash
$PY -c "import ast; ast.parse(open('tests/test_main.py').read()); print('test_main.py: sintaxis OK')"
```

Y los archivos `src/*.py`:

```bash
$PY -c "import ast; ast.parse(open('src/main.py').read()); print('main.py: sintaxis OK')"
```

## Selectividad por módulo (opcional)

Si el usuario invoca `/correr-tests <módulo>`, corre solo ese:

```bash
$PY -m unittest tests.test_<módulo>
```

Útil cuando se acaba de cambiar un módulo específico y quieres feedback rápido.

## Smoke test rápido del syntax check

Si el usuario quiere SOLO verificar que el código parsea (sin tests reales):

```bash
for f in src/*.py tests/*.py; do
    $PY -c "import ast; ast.parse(open('$f').read())" && echo "$f: OK"
done
```

## Reporte de resultados

Después de correr, reporta de forma estructurada:

```
## Resultados

### tests/test_sap_upload.py: 46 tests OK
### tests/test_sox_report.py: 105 tests OK
### tests/test_extraer_activos_creados.py: 48 tests OK
### tests/test_subir_anexos.py: 29 tests OK

**Total ejecutados: 228 — todos OK**

### tests/test_main.py (no ejecutado por falta de Tk)
- sintaxis OK (validado con ast.parse)
- Para correrlo: necesitas un Python con tkinter (python.org installer o
  python-tk via Homebrew). Ver README sección "Quick start".
```

Si hay fallos, lista los nombres exactos de los tests fallados y los assertion errors completos. NO trunques los tracebacks porque son la única pista útil.

## Particularidades del proyecto

- **Suite tarda ~5-20s** por los `time.sleep` en `subir_anexos` (`time.sleep(0.3)` × tests de `AdjuntarArchivoTest`). Es esperado, no es bug.
- **`tests/test_main.py` necesita Tk** porque crea `tk.Tk()` real en `setUp`. No hay manera de mockearlo razonablemente.
- **El venv del proyecto** está en `.venv/` con `openpyxl`, `Pillow`, `tkcalendar` instalados. Si no existe, sugerir crearlo:
  ```bash
  python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
  ```

## Cuándo invocar otras skills después

- Si hay tests fallando relacionados a un flujo SAP, sugiere invocar `comparar-vbs-python` para verificar el match con el VBS.
- Si los conteos cambiaron (ej. añadiste tests), sugiere `sincronizar-docs` para actualizar CLAUDE.md y README.md.

## Qué NO hacer

- NO ignorar tests que fallan diciendo "es algo del entorno" — siempre reporta exactamente cuáles fallaron.
- NO instalar dependencias sin avisar (`pip install X`) — pregunta antes.
- NO modificar tests para "hacerlos pasar" sin discutir con el usuario qué los está rompiendo.
- NO uses `pytest` — el proyecto usa `unittest` stdlib intencionalmente para no añadir deps.
