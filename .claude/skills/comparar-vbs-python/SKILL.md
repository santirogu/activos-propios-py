---
name: comparar-vbs-python
description: Compara una grabación VBS de SAP contra un módulo Python existente, identifica discrepancias línea por línea (IDs, ventanas, métodos, orden), y propone los cambios concretos. Útil cuando un flujo SAP falla y hay que match-ear el recording actualizado.
---

# comparar-vbs-python

Cuando el usuario invoque esta skill (normalmente cuando un flujo SAP falla en producción), compara EXACTAMENTE el `.vbs` actual contra el módulo Python para encontrar discrepancias. Esta skill nació de ~5 ciclos de debug repetitivos en `subir_anexos.py` donde el VBS evolucionaba y mi código no.

## Cuándo usarla

- El usuario actualizó el archivo `.vbs` con un recording nuevo y dice "compara".
- Un flujo SAP falla con "control not found" / "invalid argument" / "method got an invalid argument".
- El usuario reporta que algún paso visible en SAP no se ve durante la ejecución.

## Pasos

### 1. Identificar los archivos

- VBS: típicamente en `resources/*.vbs` (UTF-16LE).
- Módulo Python: `src/<modulo>.py` (sap_upload, sox_report, extraer_activos_creados, subir_anexos, o un módulo nuevo).

Si el usuario no especificó, pregunta cuál par comparar.

### 2. Leer el VBS

```bash
iconv -f UTF-16LE -t UTF-8 resources/ScriptX.vbs | nl -ba
```

Numera las líneas para referenciarlas. Ignora las líneas de boilerplate (1-14 típicamente: `If Not IsObject(application)...`).

### 3. Extraer las acciones del VBS

Para cada línea efectiva, anota:
- **Tipo de acción**: `text =`, `setFocus`, `caretPosition =`, `sendVKey`, `press`, `pressButton`, `pressContextButton`, `selectContextMenuItem`, `maximize`.
- **Path completo del findById**: ej. `wnd[0]/usr/ctxtANLA-ANLN1`, `wnd[1]/tbar[0]/btn[0]`.
- **Argumentos**: el valor del text, el id del context button, el sendVKey número, etc.

### 4. Encontrar las acciones equivalentes en Python

En el módulo Python:
- Busca los `_ejecutar(...)` con `lambda` adentro.
- Mapea cada uno a una línea esperada del VBS por la combinación (path, método, argumento).
- Usa `Grep` o lee la función principal completa para no perder pasos.

### 5. Detectar discrepancias

Las **inconsistencias comunes** en este proyecto (aprendidas a la mala):

| Tipo | Ejemplo de error | Cómo detectar |
|---|---|---|
| **Wnd equivocado** | VBS: `wnd[1]/usr/ctxtDY_PATH`, Python: `wnd[2]/usr/ctxtDY_PATH` | Comparar el prefijo `wnd[N]` literal |
| **Método equivocado** | VBS: `pressButton`, Python: `pressContextButton` | Lectura literal del método |
| **ID equivocado** | VBS: `"CREATE_ATTA"` (sin prefijo), Python: `"%GOS_CREATE_ATTA"` | Comparar el string exacto |
| **Shell diferente** | VBS: `wnd[0]/shellcont/shell`, Python: `wnd[0]/titl/shellcont/shell` | Comparar paths completos |
| **F4 extra/faltante** | VBS no tiene `sendVKey 4`, Python sí (o viceversa) | Buscar todos los `sendVKey` |
| **Cascada de btn[0] de más** | VBS termina en btn[0] de wnd[1], Python sigue confirmando wnd[0] | Comparar el último press |
| **Sleep agregado por mí** | Python tiene `time.sleep(X)`, VBS no | Buscar `time.sleep` |
| **Variable cacheada** | Python guarda `shell = findById(...)`, VBS hace `findById` cada vez | Buscar asignaciones de findById |
| **Campo extra seteado** | Python setea `ANLN2 = "0"`, VBS solo setea ANLN1 | Comparar lista de fields tocados |
| **caretPosition con valor distinto** | VBS: `caretPosition = 3`, Python: `caretPosition = len(bukrs)` (pueden coincidir o no) | Verificar el valor |
| **T-code sin `/n`** | Python: `T_CODE = "as02"`, debería ser `/nas02` para iteraciones múltiples | Buscar `T_CODE_` y verificar prefijo |

### 6. Presentar el reporte

Formato sugerido:

```
## Discrepancias encontradas en `src/<modulo>.py` vs `resources/ScriptX.vbs`

### Línea VBS 23: pressButton "%GOS_TOOLBOX"
**Tu código**: pressContextButton (método diferente)
**Fix**: cambiar `session.findById(SHELL_TITULAR).pressContextButton(GOS_TOOLBOX)` a `pressButton(GOS_TOOLBOX)`

### Línea VBS 24: shell wnd[0]/shellcont/shell (sin /titl/)
**Tu código**: SHELL_TITULAR = "wnd[0]/titl/shellcont/shell"
**Fix**: añadir constante SHELL_GOS_BAR = "wnd[0]/shellcont/shell" y usarla aquí

### Línea VBS 26: campo está en wnd[1] (no wnd[2])
**Tu código**: CAMPO_DY_PATH = "wnd[2]/usr/ctxtDY_PATH"
**Fix**: CAMPO_DY_PATH = "wnd[1]/usr/ctxtDY_PATH"

### Pasos extra en Python que NO están en el VBS
- Línea Python `_ejecutar("F4 en wnd[1]"...)` — el VBS actualizado no usa F4
- Línea Python `setattr(... DY_FILENAME ...)` — el VBS solo toca DY_PATH

### Pasos del VBS faltantes en Python
(ninguno, o listar)

## Acciones recomendadas

1. Actualizar constantes en src/X.py:
   - SHELL_GOS_BAR = "wnd[0]/shellcont/shell" (nuevo)
   - CAMPO_DY_PATH = "wnd[1]/usr/ctxtDY_PATH" (era wnd[2])
2. Eliminar paso F4 en `adjuntar_archivo`
3. Eliminar el setattr de DY_FILENAME
4. Actualizar tests en tests/test_X.py que asuman wnd[2] o F4
```

### 7. NO hacer cambios automáticamente

Esta skill es de **análisis**, no de modificación. Después de presentar el reporte:
- Pregunta al usuario si quieres que apliques los cambios sugeridos.
- Si dice sí, edita los archivos.
- Después de editar, sugiere correr la skill `correr-tests`.

## Heurísticas adicionales

- **VBS suele ser autoritativo**. Si el VBS y el Python difieren, el VBS gana (es lo que SAP ejecuta correctamente cuando el usuario lo graba manualmente).
- **Excepción**: el T-code con `/n` prefix es un agregado nuestro para soportar iteraciones múltiples. Si el VBS dice `"as02"` y Python `/nas02`, eso NO es bug, es feature — el usuario solo grabó UNA ejecución.
- **Si encuentras un cambio fundamental** (ej. el VBS ahora usa una transacción totalmente distinta), avísale al usuario antes de proponer reemplazar todo — quizá sea por error o un nuevo flujo.

## Qué NO hacer

- NO proponer "ajustes razonables" (sleeps, retries, refresh de variables) — esas teorías especulativas fallaron en este proyecto históricamente.
- NO completar pasos faltantes del VBS asumiendo que "seguramente lo que sigue es...". Si el VBS no muestra el siguiente paso (el usuario truncó la grabación), pídeselo.
