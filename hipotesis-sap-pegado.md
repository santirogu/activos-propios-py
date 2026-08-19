# Hipótesis: SAP se "pega" cuando el ejecutable corre en horario laboral

> Documento de diagnóstico. Analiza por qué, al ejecutar los flujos del proyecto
> durante horario laboral, SAP parece quedarse pegado y varios procesos fallan o
> no avanzan, afectando a otras personas que están usando SAP.

## Contexto técnico relevante

- Los flujos del proyecto usan **SAP GUI Scripting** vía COM (pywin32). Esto **no es
  headless**: no abre un SAP invisible, sino que **maneja la GUI visible real** —
  teclea en los campos, presiona botones, navega transacciones sobre una ventana
  concreta.
- `get_sap_session()` ([src/sap_upload.py:158-167](../src/sap_upload.py#L158-L167))
  toma **`application.Children(0).Children(0)`**, es decir la **primera conexión,
  primera sesión**, sin filtrar cuál. Es la pista principal de varias hipótesis.

---

## Hipótesis #1 (la más probable): la automatización comparte la sesión SAP con humanos

Como el scripting maneja la GUI visible real y el código se engancha a "la primera
sesión que encuentre", durante horario laboral:

- Si en esa máquina/servidor la **sesión 0 es la que un humano está usando**, el
  robot le toma el control: le cambia de transacción, le escribe en los campos, le
  manda F8. Para esa persona SAP "se pega" o "hace cosas solo".
- Si varias personas comparten un **servidor terminal (Citrix/RDS)** o el mismo login
  SAP, el choque es sistémico: no hay aislamiento entre lo que hace el bot y lo que
  hacen ellos.
- En la práctica: **dos "manos" sobre la misma sesión** = race conditions, pantallas
  a medio llenar, popups que quedan abiertos.

Es, de lejos, el sospechoso #1.

## Hipótesis #2: bloqueos (enqueue locks) que el flujo deja tomados

Los flujos **bloquean objetos** en SAP:

- **AS02** (Subir Anexos) abre el activo para *cambio* → lock sobre `ANLN1`. Mientras
  el bot lo tiene abierto (o si falla a media corrida y no cierra), cualquiera que
  edite ese activo ve *"bloqueado por el usuario X"*.
- **LSMW / Batch Input (SM35)**: crear y correr la BI session bloquea rangos de
  numeración / datos maestros de la clase de activo. Si corre lento o queda a medias,
  otros que crean activos esperan.

El soft-fail de `subir_anexos` ayuda a no abortar, pero **no garantiza que la
transacción se cierre limpia** tras un fallo → el lock puede quedar colgado.

## Hipótesis #3: estado modal / sesiones BDC abandonadas

Si un paso falla en medio (muy probable bajo carga), puede quedar:

- Un **popup modal abierto** en la sesión → esa sesión queda inutilizable para el
  humano hasta que alguien lo cierre.
- Una **sesión de Batch Input a medio procesar** en SM35, consumiendo *work
  processes* de diálogo del servidor.

Los helpers `_confirmar_popup_opcional` / `_volver_al_step_list` existen justamente
porque SAP deja estados colgados — pero solo cubren *la propia* sesión, no la de
terceros.

## Hipótesis #4: carrera contra la latencia del servidor en hora pico

El código hace `findById(...)` **inmediatamente después** de una acción, asumiendo que
el servidor ya respondió, y usa `time.sleep` fijos y cortos (ej.
`GOS_MENU_SETTLE_SECONDS = 0.3`). En horario laboral el servidor está cargado y
responde más lento:

- El script **se adelanta**, el control aún no existe → *"control could not be
  found"*, pasos a medio ejecutar.
- Esos medio-ejecutados son los que dejan popups/locks de las hipótesis #2 y #3.

O sea: no es que la app "tumbe" SAP, sino que **la lentitud de SAP rompe el timing de
la app**, y los restos que deja afectan a todos.

## Hipótesis #5 (menor): robo de foco / notificaciones de scripting

SAP GUI Scripting roba foco de ventana, y si están activas las notificaciones
*"Notify when a script attaches / opens a connection"* aparece un popup que
**bloquea** hasta que alguien lo acepta. La app además minimiza/restaura ventanas y
toma screenshots (flujo SOX), lo que puede pelear el foco con un usuario concurrente.

---

## Lo que falta saber para afinar el diagnóstico

La topología cambia radicalmente cuál de estas es la causa raíz:

1. **¿Dónde corre SAP GUI?** ¿En el PC de cada persona por separado, o en un
   **servidor compartido (Citrix/RDS)** al que varios entran?
2. **¿El ejecutable corre en una máquina central** (una sola, que atiende a todos)
   **o cada quien corre su propio `.exe`** contra su propio SAP?
3. **¿Con qué usuario SAP** se autentica la sesión que el bot maneja — uno dedicado al
   robot, o el mismo de una persona real?
4. Cuando "se pega", ¿la gente ve *"bloqueado por usuario X"*, popups raros, o SAP
   simplemente lento?

## Recomendación preliminar

Sujeta a confirmar la topología (puntos 1-3), la dirección del fix es:

- **No engancharse a `Children(0).Children(0)` a ciegas.** El bot debería:
  - (a) correr en una **sesión SAP dedicada e identificable** (usuario/terminal
    propio), y
  - (b) **seleccionar la sesión correcta por criterio** (system / client / usuario
    esperado) en vez de "la primera".
- Operativamente lo más sano: **correr estos flujos fuera del horario pico** o contra
  un **usuario batch aislado**, porque SAP GUI Scripting, por diseño, no está pensado
  para compartir una sesión con un humano en simultáneo.
