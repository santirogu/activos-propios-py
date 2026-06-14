---
name: sincronizar-docs
description: Después de un cambio de feature, sincroniza CLAUDE.md y README.md con el estado actual del proyecto (conteos de tests, tablas de sub-vistas, secciones arquitectónicas, módulos en el árbol de estructura). Evita docs desincronizados.
---

# sincronizar-docs

Cuando el usuario invoque esta skill después de una feature nueva o un cambio significativo, sincroniza la documentación con la realidad del código. Esta skill nació de los olvidos sistemáticos en actualizar conteos de tests y tablas tras cada feature.

## Cuándo usarla

- Después de añadir un módulo nuevo en `src/`.
- Después de añadir o eliminar tests.
- Después de añadir o cambiar una vista en la GUI.
- Después de añadir constantes/funciones que merezcan documentarse.

## Pasos

### 1. Detectar el cambio

Pregunta al usuario qué cambió si no es obvio:
- Módulo nuevo / modificado.
- Vista GUI nueva.
- Tests añadidos.
- Constantes/decisiones importantes nuevas.

Si el usuario dice "todo" o "última feature", usa `git diff HEAD~1` o `git log -1 --stat` para detectar.

### 2. Recalcular conteos de tests

Esto se rompe MUY seguido. Cuenta los `def test_*` por archivo:

```bash
grep -c "    def test_" tests/test_main.py tests/test_sap_upload.py tests/test_sox_report.py tests/test_extraer_activos_creados.py tests/test_subir_anexos.py
```

Suma total. Anota los conteos por archivo. Estos números aparecen en:

- **CLAUDE.md**: sección "8. Pruebas — XXX tests con `unittest`" → "Distribución" — busca las líneas `tests/test_X.py (NN)` y actualiza N.
- **README.md**: sección "### Cobertura de pruebas" — busca "La suite contiene **NNN pruebas**" y actualiza el total. Busca subsecciones `#### tests/test_X.py (NN pruebas)` y actualiza N.

### 3. Actualizar tabla de módulos

Si añadiste un módulo nuevo en `src/`, actualízalo en:

- **CLAUDE.md** sección "3. Estructura del repo" — el árbol ASCII del proyecto.

- **README.md** sección "Estructura del proyecto" — mismo árbol.

### 4. Actualizar tablas de sub-vistas (si tocaste la GUI)

- **CLAUDE.md** sección "4. GUI" → tabla de "Sub-vistas" con las funciones `abrir_*` y su contenido.
- **README.md** sección "### Card X" → si añadiste un botón nuevo, añade la sub-sección descriptiva.

### 5. Añadir sección arquitectónica nueva (si añadiste un módulo)

- **CLAUDE.md**: añade una sección "6.X. Flujo Nombre — `src/nombre.py`" con:
  - Constantes clave (T-code, paths SAP, IDs).
  - Mapeo del flujo en tabla (# / Función / Acciones SAP).
  - Helpers.
  - Limitaciones conocidas.
  - CLI con el comando exacto.

  Mira las secciones 6 / 6.5 / 6.6 para el formato exacto.

- **README.md**: añade una subsección dentro de la card relevante explicando el botón en lenguaje user-facing.

### 6. Añadir entrada en "Decisiones de diseño no obvias" (si aplica)

En **CLAUDE.md** sección "12. Decisiones de diseño no obvias", añade bullets nuevos si:
- Tomaste una decisión que va contra la intuición (ej. "no cachear referencias COM").
- Resolviste un bug específico que cambió el approach.
- Hay un workaround para una limitación de SAP/Tkinter/openpyxl.

Formato del bullet: `**Título corto** — explicación + por qué`.

### 7. Verificar coherencia entre CLAUDE.md y README.md

Detalles que suelen desincronizarse:
- **Conteos de tests** (total y por archivo).
- **Cantidad de botones** en una vista (ej. "tres botones" vs "cuatro").
- **Geometría de la ventana** (ej. `520x460` vs `620x580`).
- **Nombre del archivo final** (ej. `Población_*.xlsx` vs `Pob_*.xlsx`).
- **Color HEX** (`#1A3A6C` vs `#1a3a6c`).
- **T-codes** (`/nas02` vs `as02`).

Haz un grep cruzado para detectar:
```bash
grep -E "(\d+\s+(prueba|test)|wnd\[|wnd0|/n[a-z]+|620x|480x|580x)" CLAUDE.md README.md
```

### 8. Presentar el reporte y aplicar

Antes de editar, muestra al usuario un resumen:
```
## Cambios propuestos en docs

### CLAUDE.md
- L188: actualizar `tests/test_subir_anexos.py (27)` → `(29)`
- L189: actualizar "172 tests" → "180 tests"
- L240-260: añadir nueva sección "6.6. Flujo X"
- L350: nueva entrada en "Decisiones de diseño"

### README.md
- L42: "La suite contiene 175 pruebas" → "183 pruebas"
- L65: "tres botones" → "cuatro botones"
- L78: nueva subsección "Subir Anexos"
```

Pregunta confirmación. Si dice OK, aplica con Edit.

### 9. Verificar después de aplicar

```bash
# Comparar conteo real vs lo escrito
grep -c "    def test_" tests/*.py
grep -E "([0-9]+)\s+pruebas|tests" README.md CLAUDE.md
```

Si hay discrepancia, vuelve al paso 2.

## Patrones recurrentes que suelen romperse (checklist)

- [ ] Conteo total de tests en README header y CLAUDE sección 8 coinciden.
- [ ] Conteo por archivo en CLAUDE sección 8 coincide con `grep -c`.
- [ ] Estructura del repo en CLAUDE sección 3 y README "Estructura del proyecto" coinciden.
- [ ] Geometría de ventana mencionada en doc coincide con `root.geometry(...)` en main.py.
- [ ] Cantidad de botones en cada vista coincide con la realidad.
- [ ] Nombre de hoja/archivo Excel mencionado en doc coincide con constante en código.
- [ ] T-codes mencionados (con o sin /n) coinciden con `T_CODE_*` en código.

## Qué NO hacer

- NO añadir documentación de funciones triviales que ya son obvias por el nombre.
- NO copiar verbatim los docstrings del código a la doc — las docs explican el *qué* y *por qué*, no el *cómo*.
- NO escribir "se actualizó X" en CLAUDE.md — la doc vive como referencia, no como changelog (eso va al PR/commit).
- NO inventar fechas para "estado al 2026-XX-XX" — usa la del prompt o pregunta. Cuando dudes, déjala como estaba.
