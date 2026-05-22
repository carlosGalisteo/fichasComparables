# Fichas Comparables · Tenerife 🏠

Generador automático de fichas de testigos de mercado para valoraciones inmobiliarias en Tenerife.  
Extrae datos de capturas de anuncios (Idealista, Fotocasa, etc.) usando **Claude Vision** y genera un `.xlsx` listo para usar.

---

## 🚀 Despliegue en Streamlit Cloud (recomendado)

> Accesible desde cualquier lugar, sin instalación local.

### 1. Sube el proyecto a GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/TU_USUARIO/comparables-tenerife.git
git push -u origin main
```

### 2. Despliega en Streamlit Cloud

1. Ve a [share.streamlit.io](https://share.streamlit.io) e inicia sesión con GitHub.
2. Pulsa **"New app"** → selecciona tu repositorio → archivo: `app.py`.
3. En **Advanced settings → Secrets**, añade tu API key:

```toml
ANTHROPIC_API_KEY = "sk-ant-..."
```

> Si configuras el secret, la app lo usa automáticamente y no pide la key en pantalla.  
> Si no lo configuras, cada usuario introduce su propia key (modo multi-usuario).

4. Pulsa **Deploy** — en ~2 min la app está online con una URL pública.

---

## 💻 Ejecución local

```bash
# Crear entorno virtual
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar
streamlit run app.py
```

---

## 📁 Estructura del proyecto

```
comparables_app/
├── app.py                        # App principal Streamlit
├── fill_template.py              # Lógica de generación del XLSX
├── COMPARABLE_FICHA_v4.xlsx      # Plantilla base (no modificar)
├── SYSTEM_PROMPT_COMPARABLES.txt # Prompt de sistema para Claude
├── requirements.txt
└── .streamlit/
    └── config.toml               # Tema y configuración
```

---

## 🔑 API Key

La app requiere una API Key de Anthropic (`sk-ant-...`).  
Obtén la tuya en [console.anthropic.com](https://console.anthropic.com).

- **Modo equipo:** configura el secret `ANTHROPIC_API_KEY` en Streamlit Cloud → todos usan la misma key.
- **Modo personal:** cada usuario introduce su propia key → facturación individual.

---

## 🛠 Uso

1. Introduce la **API Key** en el panel lateral (si no está configurada como secret).
2. Establece la **fecha de aportación** del lote.
3. Pega las **URLs** de los anuncios (una por línea, en orden).
4. Sube las **capturas de pantalla** de los anuncios.
5. Pulsa **"Analizar con Claude"**.
6. Revisa los datos extraídos — los campos en rojo requieren revisión manual.
7. Descarga el **archivo `.xlsx`** con todas las fichas.

---

## ⚠️ Campos a revisar manualmente tras la descarga

| Celda | Campo | Motivo |
|-------|-------|--------|
| H6 | Estado de conservación | Juicio subjetivo del tasador |
| D9 | Precio de anejos | Requiere cotización separada |
| B4 | Referencia catastral | Si no se aportó |
| F9 | % Corrección | Ajustable según criterio |

---

## 📝 Notas técnicas

- Modelo: `claude-opus-4-5` (mejor precisión en visión)
- El `.xlsx` generado preserva todas las fórmulas, desplegables y estilos de la plantilla original.
- Los campos desconocidos se marcan en **rojo** (#C00000) en la ficha, igual que en la versión manual.
