# TASKS.md — plan de trabajo fichasComparables

## Objetivo general

Refactorizar `fichasComparables` para convertirla en una herramienta más visual, amable y robusta, con flujo guiado de introducción de comparables uno a uno, selector de tipología y estética corporativa.

---

## Fase 0 — Preparación

- [ ] Confirmar que se trabaja en una rama distinta de `main`.
- [ ] Confirmar que `git status` está limpio antes de empezar.
- [ ] Leer `SPEC.md`.
- [ ] Leer `CLAUDE.md`.
- [ ] Leer `TASKS.md`.
- [ ] Leer `app.py`.
- [ ] No modificar código todavía.
- [ ] Devolver ambigüedades, riesgos y plan de implementación.

---

## Fase 1 — Estado de sesión y modelo de datos

- [ ] Crear o adaptar `st.session_state.inputs`.
- [ ] Crear o adaptar `st.session_state.resultados`.
- [ ] Mantener separados los datos introducidos por el usuario y los datos extraídos por Claude.
- [ ] Evitar depender del orden entre imágenes, URLs y referencias en listas separadas.
- [ ] Mantener numeración clara de comparables.

Criterio de aceptación:

- La aplicación puede guardar internamente varios comparables como unidades independientes.

---

## Fase 2 — Nuevo flujo de entrada comparable a comparable

- [ ] Sustituir el `file_uploader` múltiple por entrada individual de comparable.
- [ ] Añadir campo de imagen por comparable.
- [ ] Añadir campo URL por comparable.
- [ ] Añadir campo referencia catastral opcional por comparable.
- [ ] Añadir selectbox de estado de conservación por comparable.
- [ ] Añadir botón `Guardar y añadir otro comparable`.
- [ ] Añadir botón `Finalizar carga del lote`.
- [ ] Permitir eliminar comparables ya añadidos.
- [ ] Mostrar resumen de comparables cargados.

Criterio de aceptación:

- El usuario puede introducir Comparable 1, luego Comparable 2, etc., sin mezclar datos.

---

## Fase 3 — Revisión previa del lote

- [ ] Mostrar resumen antes del análisis.
- [ ] Indicar imagen, URL, referencia y estado de cada comparable.
- [ ] Avisar si falta imagen.
- [ ] Avisar si falta estado de conservación.
- [ ] Deshabilitar análisis si el lote no es válido.
- [ ] Permitir limpiar todo el lote.

Criterio de aceptación:

- Antes de llamar a Claude, el usuario entiende claramente qué se va a procesar.

---

## Fase 4 — Adaptar análisis con Claude

- [ ] Adaptar el botón de análisis para recorrer `st.session_state.inputs`.
- [ ] Mantener la función `extract_comparable_from_image()` lo más estable posible.
- [ ] Pasar a Claude imagen, número, URL, fecha y referencia.
- [ ] Guardar resultados en `st.session_state.resultados`.
- [ ] Conservar nombre de imagen en cada resultado.
- [ ] Capturar errores por comparable sin romper todo el lote.

Criterio de aceptación:

- El análisis procesa correctamente los comparables guardados en el nuevo flujo.

---

## Fase 5 — Inyección de datos manuales en campos Excel

- [ ] Inyectar estado de conservación en `campos["H6"]`.
- [ ] Inyectar referencia catastral en `campos["B4"]`.
- [ ] Inyectar URL en `campos["B12"]`.
- [ ] Inyectar fecha global en `campos["H3"]`.
- [ ] Confirmar que estos datos llegan a `fill_comparable()`.

Criterio de aceptación:

- El Excel generado contiene los datos manuales correctos en sus celdas correspondientes.

---

## Fase 6 — Selector visual de tipología

- [ ] Añadir selector o botonera de tipología.
- [ ] Incluir opciones:
  - Vivienda
  - Local / Oficina
  - Garaje
  - Nave
  - Terreno
- [ ] Destacar visualmente la tipología activa.
- [ ] Actualizar título principal según tipología.
- [ ] Mantener operativa solo `Vivienda`.
- [ ] Mostrar `Ficha en preparación` para el resto.
- [ ] Deshabilitar análisis si la tipología no está operativa.

Criterio de aceptación:

- El usuario puede ver todas las tipologías, pero solo analizar vivienda.

---

## Fase 7 — Rediseño visual

- [ ] Cambiar fondo general a blanco.
- [ ] Usar texto principal negro o gris muy oscuro.
- [ ] Mantener `icono.png` en cabecera.
- [ ] Usar `#df7620` y `#dd4717` como colores corporativos.
- [ ] Usar texto blanco sobre fondos naranjas.
- [ ] Rediseñar panel izquierdo.
- [ ] Rediseñar tarjetas de comparables.
- [ ] Rediseñar botones principales.
- [ ] Mantener botón de descarga en verde Excel.
- [ ] Evitar estética oscura.
- [ ] Evitar interfaz recargada.

Criterio de aceptación:

- La app se percibe limpia, profesional, amable y corporativa.

---

## Fase 8 — Generación y descarga del Excel

- [ ] Mantener `build_xlsx_multisheet()`.
- [ ] Confirmar que un comparable genera un Excel válido.
- [ ] Confirmar que varios comparables generan un Excel multihoja válido.
- [ ] Mantener botón de descarga.
- [ ] Aplicar estilo verde Excel al botón de descarga.
- [ ] Conservar nombre de archivo configurable.

Criterio de aceptación:

- El usuario puede descargar un `.xlsx` final correctamente generado.

---

## Fase 9 — Validación manual

- [ ] Ejecutar `streamlit run app.py`.
- [ ] Probar lote con 1 comparable de vivienda.
- [ ] Probar lote con 2 o más comparables de vivienda.
- [ ] Probar eliminación de comparable.
- [ ] Probar limpiar todo.
- [ ] Probar tipologías no operativas.
- [ ] Verificar que no se han modificado archivos protegidos.
- [ ] Revisar `git diff`.
- [ ] Preparar commit.

---

## Prompt inicial recomendado para Claude Code

```text
@SPEC.md @CLAUDE.md @TASKS.md @app.py

Lee estos archivos y revisa si la especificación, las reglas del proyecto y las tareas son coherentes con el estado actual de app.py.

No modifiques código todavía.

Devuélveme:
1. Ambigüedades detectadas.
2. Riesgos técnicos.
3. Propuesta de orden de implementación.
4. Archivos que necesitarías leer antes de tocar código.
```

---

## Prompt para la primera implementación

```text
@SPEC.md @CLAUDE.md @TASKS.md @app.py

Implementa únicamente la Fase 1 y la Fase 2.

Objetivo:
Sustituir el flujo antiguo de carga múltiple + URLs/referencias en textareas por un flujo de entrada comparable a comparable usando `st.session_state.inputs`.

Restricciones:
- No modifiques `fill_template.py`.
- No modifiques `COMPARABLE_FICHA_v4.xlsx`.
- No modifiques `SYSTEM_PROMPT_COMPARABLES.txt`.
- Mantén estable la función `extract_comparable_from_image()` salvo necesidad mínima.
- Antes de tocar código, dame el plan.
```
