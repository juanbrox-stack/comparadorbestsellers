import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import anthropic
import json
import re
import time

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Comparador Cecotec",
    page_icon="🔍",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Space Grotesk', sans-serif; }

.main { background: #f8f9fb; }

.hero {
    background: linear-gradient(135deg, #e63946 0%, #c1121f 100%);
    color: white;
    padding: 2rem 2.5rem;
    border-radius: 16px;
    margin-bottom: 2rem;
}
.hero h1 { font-size: 2.2rem; font-weight: 700; margin: 0; }
.hero p  { font-size: 1rem; opacity: .85; margin: .4rem 0 0; }

.card {
    background: white;
    border-radius: 12px;
    padding: 1.5rem;
    box-shadow: 0 2px 8px rgba(0,0,0,.07);
    margin-bottom: 1.5rem;
}

.badge-better  { background:#d4edda; color:#155724; padding:3px 10px; border-radius:20px; font-size:.8rem; font-weight:600; }
.badge-equal   { background:#fff3cd; color:#856404; padding:3px 10px; border-radius:20px; font-size:.8rem; font-weight:600; }
.badge-worse   { background:#f8d7da; color:#721c24; padding:3px 10px; border-radius:20px; font-size:.8rem; font-weight:600; }
.badge-nostock { background:#e2e3e5; color:#383d41; padding:3px 10px; border-radius:20px; font-size:.8rem; font-weight:600; }

.price-save { color: #198754; font-weight: 700; }
.price-orig { color: #6c757d; text-decoration: line-through; font-size:.85rem; }

div[data-testid="stDataFrame"] table { font-size: .87rem; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <h1>🔍 Comparador de Precios · Cecotec</h1>
  <p>Introduce tu listado de productos y encontramos alternativas Cecotec con mejor precio e iguales o mejores prestaciones</p>
</div>
""", unsafe_allow_html=True)

# ── Claude client ─────────────────────────────────────────────────────────────
client = anthropic.Anthropic()

# ── Helpers ───────────────────────────────────────────────────────────────────

def scrape_cecotec(query: str) -> str:
    """Fetch Cecotec search results page and return raw HTML snippet."""
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0 Safari/537.36"
        ),
        "Accept-Language": "es-ES,es;q=0.9",
    }
    url = f"https://www.cecotec.es/buscar?q={requests.utils.quote(query)}"
    try:
        r = requests.get(url, headers=headers, timeout=12)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        # Keep only product cards to limit tokens
        cards = soup.select(".product-item, .product-card, article, .item-product")
        if cards:
            return "\n".join(str(c)[:2000] for c in cards[:6])
        return r.text[:6000]
    except Exception as e:
        return f"ERROR_SCRAPING: {e}"


def ask_claude(system_prompt: str, user_msg: str, max_tokens: int = 2000) -> str:
    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=max_tokens,
        system=system_prompt,
        messages=[{"role": "user", "content": user_msg}],
    )
    return resp.content[0].text


def find_cecotec_alternative(product: dict) -> dict:
    """
    Given a product dict with keys: nombre, marca, precio, características
    1. Scrape cecotec.es
    2. Ask Claude to pick the best match
    Returns enriched dict with Cecotec data.
    """
    query_terms = f"{product.get('categoria', '')} {product.get('características', '')[:80]}"
    html_snippet = scrape_cecotec(query_terms)

    system = (
        "Eres un experto en electrónica de hogar. Se te dará información de un producto "
        "de referencia y fragmentos HTML de cecotec.es. Debes identificar el mejor producto "
        "Cecotec que iguale o supere las prestaciones del producto de referencia, tenga stock "
        "y sea más barato. Responde SOLO con JSON válido, sin markdown."
    )

    user = f"""
Producto de referencia:
{json.dumps(product, ensure_ascii=False, indent=2)}

HTML de cecotec.es (resultados de búsqueda):
{html_snippet[:5000]}

Devuelve un objeto JSON con estos campos exactos (usa null si no encuentras el dato):
{{
  "cecotec_nombre": "...",
  "cecotec_precio": 0.0,
  "cecotec_caracteristicas": "...",
  "cecotec_url": "...",
  "cecotec_referencia": "...",
  "cecotec_stock": true/false,
  "amazon_asin": "...",
  "ahorro": 0.0,
  "prestaciones": "mejor|igual|peor",
  "justificacion": "..."
}}
Si no existe ningún producto Cecotec adecuado, devuelve {{"no_encontrado": true, "justificacion": "..."}}.
"""
    raw = ask_claude(system, user)
    # Strip possible markdown fences
    raw = re.sub(r"```json|```", "", raw).strip()
    try:
        return json.loads(raw)
    except Exception:
        return {"no_encontrado": True, "justificacion": raw[:300]}


# ── Sidebar input ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuración")
    input_method = st.radio(
        "Método de entrada",
        ["Formulario manual", "Subir CSV/Excel", "Pegar JSON"],
        index=0,
    )

# ── Product list state ────────────────────────────────────────────────────────
if "products" not in st.session_state:
    st.session_state.products = []

# ─── Input: Manual form ───────────────────────────────────────────────────────
if input_method == "Formulario manual":
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("➕ Añadir producto")
    c1, c2, c3 = st.columns(3)
    with c1:
        p_nombre = st.text_input("Nombre del producto")
        p_marca  = st.text_input("Marca")
    with c2:
        p_precio = st.number_input("Precio (€)", min_value=0.0, step=0.01)
        p_cat    = st.text_input("Categoría (ej: robot aspirador)")
    with c3:
        p_caract = st.text_area("Características principales", height=100,
                                placeholder="Potencia, filtro, depósito, autonomía…")

    if st.button("Añadir a la lista", type="primary"):
        if p_nombre and p_precio:
            st.session_state.products.append({
                "nombre": p_nombre,
                "marca": p_marca,
                "precio": p_precio,
                "categoria": p_cat,
                "características": p_caract,
            })
            st.success(f"✅ '{p_nombre}' añadido.")
        else:
            st.warning("Nombre y precio son obligatorios.")
    st.markdown("</div>", unsafe_allow_html=True)

# ─── Input: CSV/Excel ─────────────────────────────────────────────────────────
elif input_method == "Subir CSV/Excel":
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📂 Subir archivo")
    st.caption("El archivo debe tener columnas: nombre, marca, precio, categoria, características")
    uploaded = st.file_uploader("CSV o Excel", type=["csv", "xlsx", "xls"])
    if uploaded:
        try:
            if uploaded.name.endswith(".csv"):
                df = pd.read_csv(uploaded)
            else:
                df = pd.read_excel(uploaded)
            st.dataframe(df, use_container_width=True)
            if st.button("Usar este listado", type="primary"):
                st.session_state.products = df.to_dict("records")
                st.success(f"✅ {len(df)} productos cargados.")
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

# ─── Input: JSON ──────────────────────────────────────────────────────────────
elif input_method == "Pegar JSON":
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📋 Pegar JSON")
    sample = json.dumps([
        {"nombre": "Roomba i3+", "marca": "iRobot", "precio": 299.99,
         "categoria": "robot aspirador",
         "características": "2000 Pa, HEPA, vaciado automático, 75 min autonomía"}
    ], ensure_ascii=False, indent=2)
    json_input = st.text_area("JSON (lista de productos)", value=sample, height=200)
    if st.button("Cargar JSON", type="primary"):
        try:
            st.session_state.products = json.loads(json_input)
            st.success(f"✅ {len(st.session_state.products)} productos cargados.")
        except Exception as e:
            st.error(f"JSON inválido: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

# ── Current product list ──────────────────────────────────────────────────────
if st.session_state.products:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader(f"📋 Listado actual — {len(st.session_state.products)} producto(s)")
    df_list = pd.DataFrame(st.session_state.products)
    st.dataframe(df_list, use_container_width=True, hide_index=True)
    col_clear, _ = st.columns([1, 5])
    with col_clear:
        if st.button("🗑️ Limpiar lista"):
            st.session_state.products = []
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# ── Search button ─────────────────────────────────────────────────────────────
    st.divider()
    if st.button("🚀 Buscar alternativas en Cecotec", type="primary", use_container_width=True):
        results = []
        progress = st.progress(0, text="Iniciando búsqueda…")
        total = len(st.session_state.products)

        for i, prod in enumerate(st.session_state.products):
            progress.progress((i) / total, text=f"Buscando alternativa para: {prod['nombre']}…")
            alt = find_cecotec_alternative(prod)
            time.sleep(0.5)  # polite scraping
            results.append({"producto_ref": prod, "alternativa": alt})
            progress.progress((i + 1) / total, text=f"✅ {prod['nombre']} procesado")

        progress.empty()
        st.session_state.results = results
        st.rerun()

# ── Results ───────────────────────────────────────────────────────────────────
if "results" in st.session_state and st.session_state.results:
    st.markdown("## 📊 Tabla Comparativa")

    rows = []
    for r in st.session_state.results:
        ref = r["producto_ref"]
        alt = r["alternativa"]

        if alt.get("no_encontrado"):
            rows.append({
                "Producto referencia": ref["nombre"],
                "Marca ref.": ref.get("marca", ""),
                "Precio ref. (€)": ref.get("precio", ""),
                "Alternativa Cecotec": "❌ No encontrado",
                "Precio Cecotec (€)": "—",
                "Ahorro (€)": "—",
                "Prestaciones": "—",
                "Stock": "—",
                "Características Cecotec": alt.get("justificacion", ""),
                "URL Cecotec": "—",
                "Referencia Cecotec": "—",
                "ASIN Amazon": "—",
            })
        else:
            prest = alt.get("prestaciones", "")
            badge = {"mejor": "✅ Mejor", "igual": "🟡 Igual", "peor": "🔴 Peor"}.get(prest, prest)
            stock = "✅ Sí" if alt.get("cecotec_stock") else "❌ No"
            ahorro = alt.get("ahorro") or (ref.get("precio", 0) - (alt.get("cecotec_precio") or 0))
            rows.append({
                "Producto referencia": ref["nombre"],
                "Marca ref.": ref.get("marca", ""),
                "Precio ref. (€)": ref.get("precio", ""),
                "Alternativa Cecotec": alt.get("cecotec_nombre", ""),
                "Precio Cecotec (€)": alt.get("cecotec_precio", ""),
                "Ahorro (€)": round(ahorro, 2) if isinstance(ahorro, (int, float)) else ahorro,
                "Prestaciones": badge,
                "Stock": stock,
                "Características Cecotec": alt.get("cecotec_caracteristicas", ""),
                "URL Cecotec": alt.get("cecotec_url", ""),
                "Referencia Cecotec": alt.get("cecotec_referencia", ""),
                "ASIN Amazon": alt.get("amazon_asin", ""),
            })

    df_res = pd.DataFrame(rows)
    st.dataframe(
        df_res,
        use_container_width=True,
        hide_index=True,
        column_config={
            "URL Cecotec": st.column_config.LinkColumn("URL Cecotec"),
            "Precio ref. (€)": st.column_config.NumberColumn(format="%.2f €"),
            "Precio Cecotec (€)": st.column_config.NumberColumn(format="%.2f €"),
            "Ahorro (€)": st.column_config.NumberColumn(format="%.2f €"),
        },
    )

    # Export
    csv = df_res.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Descargar CSV",
        data=csv,
        file_name="comparativa_cecotec.csv",
        mime="text/csv",
    )

    # Detail cards
    st.markdown("## 🔎 Detalle por producto")
    for r in st.session_state.results:
        ref = r["producto_ref"]
        alt = r["alternativa"]
        with st.expander(f"**{ref['nombre']}** — {ref.get('marca','')}"):
            col_ref, col_arrow, col_alt = st.columns([5, 1, 5])
            with col_ref:
                st.markdown("**📦 Producto referencia**")
                st.markdown(f"**{ref['nombre']}** · {ref.get('marca','')}")
                st.markdown(f"💶 **{ref.get('precio','')} €**")
                st.markdown(f"_{ref.get('características','')}_")
            with col_arrow:
                st.markdown("<div style='font-size:2rem;text-align:center;margin-top:40px'>→</div>",
                            unsafe_allow_html=True)
            with col_alt:
                if alt.get("no_encontrado"):
                    st.error("No se encontró alternativa adecuada.")
                    st.caption(alt.get("justificacion", ""))
                else:
                    st.markdown("**🟥 Alternativa Cecotec**")
                    st.markdown(f"**{alt.get('cecotec_nombre','')}**")
                    st.markdown(f"💶 **{alt.get('cecotec_precio','')} €**")
                    prest = alt.get("prestaciones", "")
                    badge_map = {"mejor": "badge-better", "igual": "badge-equal", "peor": "badge-worse"}
                    label_map = {"mejor": "✅ Mejores prestaciones", "igual": "🟡 Prestaciones equivalentes", "peor": "🔴 Prestaciones inferiores"}
                    css_cls = badge_map.get(prest, "badge-equal")
                    label = label_map.get(prest, prest)
                    st.markdown(f'<span class="{css_cls}">{label}</span>', unsafe_allow_html=True)
                    st.markdown(f"_{alt.get('cecotec_caracteristicas','')}_")
                    if alt.get("cecotec_url"):
                        st.markdown(f"[🔗 Ver en Cecotec.es]({alt['cecotec_url']})")
                    if alt.get("amazon_asin"):
                        asin = alt["amazon_asin"]
                        st.markdown(f"[🛒 Ver en Amazon](https://www.amazon.es/dp/{asin})")
                    if alt.get("justificacion"):
                        st.caption(f"💡 {alt['justificacion']}")

else:
    if not st.session_state.products:
        st.info("👆 Añade productos usando el formulario o sube un archivo para comenzar.")
