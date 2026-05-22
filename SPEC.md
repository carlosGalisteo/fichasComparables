# SPEC.md — fichasComparables

## 1. Visión del producto

`fichasComparables` es una aplicación Streamlit para generar fichas Excel de testigos de mercado inmobiliario a partir de capturas de anuncios.

La herramienta está orientada a un flujo profesional de valoración inmobiliaria en Tenerife, donde el usuario introduce comparables uno a uno, la aplicación extrae datos mediante Claude Vision API y finalmente genera un archivo `.xlsx` con una hoja por comparable.

La aplicación debe ser clara, visualmente amable, profesional y fácil de usar por un usuario no técnico.

---

## 2. Usuario principal y caso de uso

El usuario principal es un técnico que necesita preparar fichas de comparables inmobiliarios para valoración.

El usuario puede:

- Seleccionar la tipología del lote de comparables.
- Introducir comparables uno a uno.
- Aportar para cada comparable:
  - imagen del anuncio,
  - URL del anuncio,
  - referencia catastral opcional,
  - estado de conservación.
- Revisar los comparables cargados antes de analizarlos.
- Lanzar el análisis con Claude.
- Revisar los datos extraídos.
- Descargar un archivo Excel `.xlsx` con todas las fichas generadas.

---

## 3. Tipologías de ficha

La aplicación debe incluir una selección visual de tipología del lote.

Opciones:

- Vivienda
- Local / Oficina
- Garaje
- Nave
- Terreno

La selección es global para todo el lote. Un lote no debe mezclar tipologías.

En esta iteración solo está operativa la tipología `Vivienda`.

Las demás tipologías deben aparecer visibles, pero como `Ficha en preparación`, y no deben permitir lanzar el análisis hasta que exista su plantilla correspondiente.

Cuando se seleccione una tipología, el título principal de la pantalla debe actualizarse:

- `Comparables de vivienda`
- `Comparables de local / oficina`
- `Comparables de garaje`
- `Comparables de nave`
- `Comparables de terreno`

---

## 4. Flujo principal

La aplicación debe abandonar el flujo antiguo basado en:

- subida múltiple de imágenes,
- textarea de URLs,
- textarea de referencias catastrales,
- asociación por orden.

El nuevo flujo debe ser comparable a comparable.

Para cada comparable, el usuario introduce:

- Imagen del anuncio.
- URL del anuncio.
- Referencia catastral opcional.
- Estado de conservación:
  - Muy bueno
  - Bueno
  - Medio
  - Malo
  - Muy malo

Después de introducir los datos de un comparable, el usuario debe poder:

- Guardar y añadir otro comparable.
- Finalizar la carga del lote y analizar.

Antes de llamar a Claude, la aplicación debe mostrar un resumen del lote cargado.

El usuario debe poder eliminar un comparable antes del análisis.

---

## 5. Estructura general de la interfaz

La aplicación debe funcionar como un asistente visual para crear un lote de comparables.

Debe existir un panel principal izquierdo, plegable o fácilmente ocultable, donde se concentre la funcionalidad principal de carga y configuración.

Este panel debe incluir:

- Estado de API Key.
- Fecha global de aportación.
- Formulario del comparable actual.
- Botones para guardar, añadir otro comparable y finalizar.
- Resumen de comparables cargados.
- Opción de limpiar todo.

Debe existir otra zona visual específica para seleccionar la tipología de ficha mediante botonera o tarjetas laterales.

La tipología activa debe destacarse con color corporativo.

La zona central debe mostrar:

- Cabecera con `icono.png`.
- Título dinámico según tipología.
- Mensaje de estado del lote.
- Resumen de comparables cargados.
- Resultados del análisis.
- Botón de descarga del Excel.

---

## 6. Estado de sesión

La aplicación debe usar un modelo explícito de entrada de datos.

Estructura deseada:

```python
st.session_state.inputs = [
    {
        "num": 1,
        "file": uploaded_file,
        "url": "https://...",
        "ref": "",
        "estado": "Bueno",
    }
]
```

Los resultados extraídos por Claude deben almacenarse separadamente:

```python
st.session_state.resultados = [
    {
        "comparable": 1,
        "campos": {...},
        "flags": {...},
        "img_name": "captura.png",
    }
]
```

---

## 7. Reglas de datos

El estado de conservación lo introduce el usuario. Claude no debe extraerlo.

Antes de generar el Excel:

- El estado elegido debe escribirse en `campos["H6"]`.
- La referencia catastral debe escribirse en `campos["B4"]`.
- La URL debe escribirse en `campos["B12"]`.
- La fecha de aportación debe escribirse en `campos["H3"]`.

La fecha de aportación es global para todo el lote.

La plantilla actual de vivienda es:

```text
COMPARABLE_FICHA_v4.xlsx
```

---

## 8. Criterios visuales

La aplicación debe ser estética, amable y profesional.

Reglas visuales:

- Fondo general blanco.
- Texto principal negro o gris muy oscuro sobre fondo blanco.
- Mantener `icono.png` junto al título.
- Usar los colores corporativos:
  - Naranja principal: `#df7620`
  - Naranja secundario: `#dd4717`
- Los naranjas pueden usarse en:
  - paneles laterales,
  - botonera de tipología,
  - estados activos,
  - acentos visuales.
- Sobre fondos naranjas, el texto debe ser blanco.
- El botón de descarga del Excel debe usar verde asociado a Excel.
- Evitar una estética oscura general.
- Evitar una interfaz técnica o recargada.
- Priorizar claridad, jerarquía visual y facilidad de uso.

---

## 9. Arquitectura actual

Stack actual:

- Streamlit
- Anthropic Python SDK
- Claude Vision API
- openpyxl
- Pillow
- Streamlit Cloud
- Python 3.14

Archivos principales:

```text
app.py
fill_template.py
COMPARABLE_FICHA_v4.xlsx
SYSTEM_PROMPT_COMPARABLES.txt
icono.png
requirements.txt
.streamlit/config.toml
```

---

## 10. Restricciones

No modificar:

```text
fill_template.py
COMPARABLE_FICHA_v4.xlsx
SYSTEM_PROMPT_COMPARABLES.txt
```

Estos archivos se consideran lógica de negocio validada.

El archivo principal a refactorizar es:

```text
app.py
```

La lógica de extracción con Claude debe modificarse lo mínimo imprescindible.

La lógica de generación multihoja del Excel debe conservarse salvo necesidad justificada.

---

## 11. Requisitos de verificación

Antes de dar por buena una iteración deben probarse como mínimo estos casos:

- Ejecutar la app localmente con `streamlit run app.py`.
- Crear un lote con 1 comparable de vivienda.
- Crear un lote con varios comparables de vivienda.
- Confirmar que cada comparable conserva su propia imagen, URL, referencia y estado.
- Confirmar que el estado de conservación se escribe en `H6`.
- Confirmar que la URL se escribe en `B12`.
- Confirmar que la referencia catastral se escribe en `B4`.
- Confirmar que se genera un `.xlsx` descargable.
- Confirmar que Local / Oficina, Garaje, Nave y Terreno aparecen como opciones no operativas.
- Confirmar que la interfaz mantiene fondo blanco, texto oscuro, naranjas corporativos y botón Excel verde.
