# 🔍 Comparador de Precios · Cecotec

App en Streamlit que compara tus productos con alternativas de cecotec.es usando IA (Claude).

## Instalación

```bash
pip install -r requirements.txt
```

## Ejecución

```bash
# Exporta tu API key de Anthropic
export ANTHROPIC_API_KEY="sk-ant-..."

streamlit run app.py
```

## Uso

1. **Añade tus productos** (formulario manual, CSV/Excel o JSON)
2. Haz clic en **"Buscar alternativas en Cecotec"**
3. La app scrapeará cecotec.es y usará Claude para analizar y comparar
4. Descarga el resultado como **CSV**

## Columnas del CSV de entrada (opcional)

| Columna | Descripción |
|---|---|
| nombre | Nombre del producto |
| marca | Marca fabricante |
| precio | Precio en euros |
| categoria | Tipo de producto (ej: robot aspirador) |
| características | Especificaciones técnicas principales |

## Notas

- La app respeta una pausa de 0.5 s entre peticiones a cecotec.es
- El ASIN de Amazon lo infiere Claude; verifica siempre el enlace
- Requiere `ANTHROPIC_API_KEY` en el entorno
