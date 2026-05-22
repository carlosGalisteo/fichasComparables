# CLAUDE.md — memoria operativa del proyecto fichasComparables

## Stack

- Lenguaje: Python
- Framework app: Streamlit
- IA: Anthropic Python SDK
- Modelo actual: `claude-opus-4-5`
- Imagen: Pillow
- Excel: openpyxl
- Despliegue: Streamlit Cloud
- Python objetivo: 3.14

---

## Archivos principales

- `app.py`: archivo principal de la aplicación. Es el archivo que se debe refactorizar.
- `fill_template.py`: lógica validada de escritura de Excel. No modificar.
- `COMPARABLE_FICHA_v4.xlsx`: plantilla validada para vivienda. No modificar.
- `SYSTEM_PROMPT_COMPARABLES.txt`: prompt de negocio validado. No modificar.
- `icono.png`: logo de empresa usado en cabecera.
- `requirements.txt`: dependencias.
- `.streamlit/config.toml`: configuración de Streamlit.

---

## Reglas duras

- No modificar `fill_template.py`.
- No modificar `COMPARABLE_FICHA_v4.xlsx`.
- No modificar `SYSTEM_PROMPT_COMPARABLES.txt`.
- No modificar la lógica de negocio validada salvo autorización expresa.
- No hacer un refactor masivo de una sola vez.
- No cambiar el stack tecnológico.
- No introducir frameworks externos de UI.
- No introducir dependencias nuevas sin justificarlo antes.
- No trabajar directamente sobre `main`.
- No hacer push a `main` sin confirmación expresa.
- No redesplegar Streamlit Cloud hasta que el usuario lo autorice.
- No cambiar el modelo de Claude sin confirmación.
- No cambiar nombres de campos/celdas Excel sin confirmación.
- No eliminar funcionalidad existente sin justificarlo.

---

## Reglas de trabajo con Claude Code

Antes de cualquier cambio no trivial:

1. Leer `SPEC.md`.
2. Leer `CLAUDE.md`.
3. Leer `TASKS.md`.
4. Leer los archivos necesarios.
5. Proponer un plan.
6. Esperar aprobación explícita.
7. Implementar solo la tarea aprobada.
8. Verificar.
9. Proponer commit.

Usar Plan Mode para:

- refactors,
- cambios de interfaz,
- cambios de estado de sesión,
- cambios en el flujo de análisis,
- cambios en generación de Excel,
- cualquier cambio que toque más de una zona del archivo.

---

## Convenciones

- Mantener nombres claros y explícitos.
- Preferir funciones pequeñas y comprensibles.
- Separar lógica de estado, renderizado UI, procesamiento y generación de Excel cuando sea razonable.
- Evitar duplicidad de lógica.
- Evitar soluciones frágiles basadas en el orden implícito de listas.
- Mantener comentarios útiles cuando aclaren decisiones de negocio.
- Evitar comentarios obvios.

---

## Reglas de UI

La app debe ser visualmente estética, amable y profesional.

- Fondo principal blanco.
- Texto principal negro o gris muy oscuro.
- Mantener `icono.png` junto al título.
- Usar naranja principal `#df7620`.
- Usar naranja secundario `#dd4717`.
- Sobre fondos naranjas, usar texto blanco.
- Usar verde tipo Excel para el botón de descarga `.xlsx`.
- Evitar estética oscura general.
- Evitar aspecto técnico o recargado.
- Priorizar claridad, jerarquía visual y usabilidad.
- Cada comparable debe verse como una unidad independiente.
- El usuario debe entender fácilmente qué imagen, URL, referencia y estado pertenecen a cada comparable.

---

## Flujo funcional deseado

La app debe pasar de un flujo basado en listas paralelas a un flujo guiado comparable a comparable.

Para cada comparable:

1. El usuario carga una imagen.
2. Introduce la URL.
3. Introduce referencia catastral opcional.
4. Elige estado de conservación.
5. Guarda el comparable.
6. Decide si añade otro o finaliza el lote.

Después:

1. Revisa el lote.
2. Lanza análisis con Claude.
3. Revisa resultados.
4. Descarga Excel.

---

## Tipologías

Opciones visibles:

- Vivienda
- Local / Oficina
- Garaje
- Nave
- Terreno

Solo `Vivienda` está operativa en esta iteración.

El resto deben aparecer como `Ficha en preparación`.

Si la tipología seleccionada no está operativa, el botón de análisis debe quedar deshabilitado.

---

## Datos que deben inyectarse manualmente en el resultado

El usuario introduce datos que no debe decidir Claude.

Antes de generar el Excel:

```python
campos["H6"] = estado_usuario
campos["B4"] = referencia_catastral_usuario
campos["B12"] = url_usuario
campos["H3"] = fecha_aportacion_global
```

---

## Comandos útiles

Ejecutar app local:

```powershell
streamlit run app.py
```

Comprobar estado Git:

```powershell
git status
```

Crear rama de trabajo:

```powershell
git checkout -b refactor/guiado-comparables
```

Ver diferencias:

```powershell
git diff
```

Añadir archivos concretos:

```powershell
git add app.py SPEC.md CLAUDE.md TASKS.md
```

Crear commit:

```powershell
git commit -m "refactor: rediseña flujo guiado de comparables"
```

---

## Criterios de aceptación

Una tarea no está terminada hasta que:

- La app arranca localmente.
- No se rompe la extracción con Claude.
- No se rompe la generación del Excel.
- Se mantiene la plantilla existente.
- Se puede generar un Excel de vivienda con 1 comparable.
- Se puede generar un Excel de vivienda con varios comparables.
- La interfaz es coherente con la identidad visual definida.
- No se han tocado los archivos protegidos.
