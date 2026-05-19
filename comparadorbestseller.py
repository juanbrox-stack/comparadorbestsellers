"""
Comparador Cecotec — Keepa Bestsellers (Hogar + Belleza)
Motor IA: Google Gemini Flash (gratuito via Google AI Studio)
"""

import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import json, re, time, os
from pathlib import Path
import google.generativeai as genai

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Comparador Cecotec · Keepa", page_icon="🔍", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.hero {
    background: linear-gradient(120deg,#c0392b,#e74c3c 60%,#e67e22);
    color:#fff; padding:2rem 2.5rem; border-radius:16px; margin-bottom:1.5rem;
}
.hero h1 { font-size:2rem; font-weight:700; margin:0; }
.hero p  { opacity:.88; margin:.3rem 0 0; font-size:.95rem; }
.kpi-row { display:flex; gap:1rem; margin-bottom:1.5rem; flex-wrap:wrap; }
.kpi { background:#fff; border-radius:12px; padding:1rem 1.4rem;
       box-shadow:0 2px 8px rgba(0,0,0,.07); flex:1; min-width:130px; }
.kpi .val { font-size:1.8rem; font-weight:700; color:#c0392b; }
.kpi .lbl { font-size:.78rem; color:#666; text-transform:uppercase; letter-spacing:.04em; }
.card { background:#fff; border-radius:12px; padding:1.4rem;
        box-shadow:0 2px 8px rgba(0,0,0,.07); margin-bottom:1.2rem; }
.tag-mejor { background:#d4edda; color:#155724; padding:2px 10px; border-radius:20px; font-size:.8rem; font-weight:600; }
.tag-igual { background:#fff3cd; color:#856404; padding:2px 10px; border-radius:20px; font-size:.8rem; font-weight:600; }
.tag-peor  { background:#f8d7da; color:#721c24; padding:2px 10px; border-radius:20px; font-size:.8rem; font-weight:600; }
.tag-skip  { background:#e2e3e5; color:#383d41; padding:2px 10px; border-radius:20px; font-size:.8rem; font-weight:600; }
</style>
""", unsafe_allow_html=True)

# ── API Key setup ─────────────────────────────────────────────────────────────
def get_gemini_key():
    # 1. session_state (introducida por el usuario en UI)
    if "GOOGLE_API_KEY" in st.session_state and st.session_state["GOOGLE_API_KEY"]:
        return st.session_state["GOOGLE_API_KEY"]
    # 2. Streamlit secrets
    try:
        k = st.secrets.get("GOOGLE_API_KEY", None)
        if k:
            return k
    except Exception:
        pass
    # 3. Env var
    return os.environ.get("GOOGLE_API_KEY", None)

api_key = get_gemini_key()

if not api_key:
    st.markdown("""
    <div class="hero">
      <h1>🔍 Comparador Cecotec · Bestsellers Amazon</h1>
      <p>Motor IA: Google Gemini (gratuito)</p>
    </div>
    """, unsafe_allow_html=True)
    st.warning("### 🔑 Configura tu API key de Google Gemini (gratuita)")
    st.markdown("""
    1. Ve a **[aistudio.google.com/apikey](https://aistudio.google.com/apikey)**
    2. Haz clic en **"Create API Key"** (es gratuito)
    3. Copia la key y pégala aquí abajo 👇
    """)
    key_input = st.text_input("Google API Key", type="password", placeholder="AIzaSy...")
    if st.button("✅ Guardar y continuar", type="primary"):
        if key_input and key_input.startswith("AIza"):
            st.session_state["GOOGLE_API_KEY"] = key_input
            st.rerun()
        else:
            st.error("Key inválida. Debe empezar por AIzaSy...")
    st.info("O configúrala en **Streamlit Cloud → Settings → Secrets**: `GOOGLE_API_KEY = 'AIzaSy...'`")
    st.stop()

genai.configure(api_key=api_key)

@st.cache_resource
def get_model():
    return genai.GenerativeModel("gemini-1.5-flash")

model = get_model()

# ── Constants ─────────────────────────────────────────────────────────────────
KEEPA_HOGAR   = "KeepaExport-2026-05-19-BestSellersList-9-599391031.xlsx"
KEEPA_BELLEZA = "BellezaKeepaExport-2026-05-19-BestSellersList-9-4347698031-9000.xlsx"

CECOTEC_CATS = {
    "Aspiradoras escoba", "Aspiradoras de mano", "Robots aspiradores",
    "Aspiradoras para alfombras", "Aspiradoras con bolsa",
    "Freidoras de aire", "Freidoras", "Hornos de sobremesa",
    "Tostadoras", "Sandwicheras y gofradoras", "Grills de contacto",
    "Batidoras de mano", "Batidoras de vaso", "Procesadores de alimentos",
    "Cafeteras italianas", "Cafeteras de filtro", "Cafeteras espresso",
    "Máquinas de café",
    "Planchas de vapor", "Planchas de vapor verticales para viaje",
    "Cepillos de vapor", "Centros de planchado",
    "Balanzas digitales", "Básculas de cocina", "Básculas de baño",
    "Batidoras amasadoras",
    "Purificadores de aire", "Humidificadores", "Ventiladores",
    "Aires acondicionados portátiles",
    "Televisores", "Monitores", "Proyectores", "Altavoces portátiles",
    "Desincrustantes",
    "Secadores de pelo", "Planchas para el pelo", "Rizadores",
    "Cepillos eléctricos para el cabello",
    "Afeitadoras eléctricas", "Depiladores",
    "Cepillos de dientes eléctricos", "Masajeadores",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9",
}

# ── Data loading ──────────────────────────────────────────────────────────────
@st.cache_data
def load_keepa_files(upload_dir: str) -> pd.DataFrame:
    dfs = []
    for fname, source in [(KEEPA_HOGAR, "Hogar"), (KEEPA_BELLEZA, "Belleza")]:
        path = Path(upload_dir) / fname
        if not path.exists():
            st.warning(f"Archivo no encontrado: {fname}")
            continue
        df = pd.read_excel(path)
        df["_source"] = source
        dfs.append(df)
    if not dfs:
        return pd.DataFrame()
    merged = pd.concat(dfs, ignore_index=True)
    merged = merged.rename(columns={
        "ASIN": "asin",
        "Título": "titulo",
        "Caja de Compra: Actual": "precio",
        "Categorías: Principal": "categoria_principal",
        "Categorías: Subcategoría": "subcategoria",
        "Clasificación de Ventas: Actual": "ranking",
        "Fabricante": "fabricante",
        "URL: Amazon": "url_amazon",
        "Códigos de producto: EAN": "ean",
        "Opiniones: Valoraciones": "rating",
        "Opiniones: Cantidad de valoraciones": "num_reviews",
        "Descripción & Características: Característica 1": "feat1",
        "Descripción & Características: Característica 2": "feat2",
        "Descripción & Características: Característica 3": "feat3",
        "Descripción & Características: Característica 4": "feat4",
        "Descripción & Características: Descripción breve": "descripcion_breve",
    })
    merged = merged.sort_values("ranking")
    merged = merged.groupby("_source").head(100).reset_index(drop=True)
    merged["_cecotec_relevant"] = merged["subcategoria"].apply(
        lambda s: any(cat.lower() in str(s).lower() for cat in CECOTEC_CATS)
    )
    def build_feats(row):
        parts = [str(row.get(f"feat{i}", "") or "") for i in range(1, 5)]
        parts.append(str(row.get("descripcion_breve", "") or ""))
        return " | ".join(p[:120] for p in parts if p.strip())[:500]
    merged["caracteristicas"] = merged.apply(build_feats, axis=1)
    return merged

# ── Scraping ──────────────────────────────────────────────────────────────────
def scrape_cecotec(query: str) -> str:
    url = f"https://www.cecotec.es/buscar?q={requests.utils.quote(query)}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=14)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        cards = soup.select(".product-item,.product-card,article,.item-product,[class*='product']")
        if cards:
            return "\n".join(str(c)[:1800] for c in cards[:5])
        return soup.get_text(separator=" ", strip=True)[:5000]
    except Exception as e:
        return f"SCRAPING_ERROR: {e}"

# ── Gemini analysis ───────────────────────────────────────────────────────────
PROMPT_TEMPLATE = """Eres un experto en electrónica de hogar y pequeño electrodoméstico español.

Producto de referencia Amazon:
{ref}

HTML de resultados cecotec.es:
{html}

Tu tarea: identifica el mejor producto Cecotec que:
1. Esté en stock
2. Iguale o supere las prestaciones del producto de referencia
3. Tenga precio INFERIOR al producto de referencia ({precio} €)

Responde SOLO con JSON válido, sin markdown, sin texto fuera del JSON.

Si encuentras alternativa válida:
{{
  "cecotec_nombre": "...",
  "cecotec_precio": 0.0,
  "cecotec_caracteristicas": "resumen specs clave en 100 chars",
  "cecotec_url": "https://www.cecotec.es/...",
  "cecotec_referencia": null,
  "cecotec_stock": true,
  "amazon_asin_cecotec": null,
  "ahorro_eur": 0.0,
  "prestaciones": "mejor|igual|peor",
  "justificacion": "max 120 chars"
}}

Si NO existe alternativa válida:
{{"no_encontrado": true, "motivo": "razón breve"}}"""

def analyze_with_gemini(row: pd.Series, html: str) -> dict:
    ref = {
        "titulo": row.get("titulo", ""),
        "fabricante": row.get("fabricante", ""),
        "subcategoria": row.get("subcategoria", ""),
        "caracteristicas": row.get("caracteristicas", ""),
    }
    prompt = PROMPT_TEMPLATE.format(
        ref=json.dumps(ref, ensure_ascii=False),
        html=html[:4000],
        precio=row.get("precio", 0),
    )
    try:
        resp = model.generate_content(prompt)
        raw = re.sub(r"```json|```", "", resp.text).strip()
        return json.loads(raw)
    except Exception as e:
        return {"no_encontrado": True, "motivo": str(e)[:150]}

def build_query(row: pd.Series) -> str:
    subcat = str(row.get("subcategoria", "")).split(",")[0].strip()
    feat   = str(row.get("feat1", "") or "")[:60]
    return f"{subcat} {feat}".strip()

# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <h1>🔍 Comparador Cecotec · Bestsellers Amazon</h1>
  <p>Top-100 Hogar &amp; Belleza · Alternativas Cecotec con mejor precio · Motor: Google Gemini ✨</p>
</div>
""", unsafe_allow_html=True)

UPLOAD_DIR = Path(__file__).parent
df_all = load_keepa_files(str(UPLOAD_DIR))

if df_all.empty:
    st.error("No se encontraron los archivos Keepa. Asegúrate de que los xlsx están en la misma carpeta que app.py.")
    st.stop()

df_relevant = df_all[df_all["_cecotec_relevant"]].copy()
df_skipped  = df_all[~df_all["_cecotec_relevant"]].copy()

total_hogar    = len(df_all[df_all["_source"] == "Hogar"])
total_belleza  = len(df_all[df_all["_source"] == "Belleza"])
total_relevant = len(df_relevant)

st.markdown(f"""
<div class="kpi-row">
  <div class="kpi"><div class="val">{total_hogar}</div><div class="lbl">Top-100 Hogar</div></div>
  <div class="kpi"><div class="val">{total_belleza}</div><div class="lbl">Top-100 Belleza</div></div>
  <div class="kpi"><div class="val">{total_relevant}</div><div class="lbl">Con equivalente Cecotec posible</div></div>
  <div class="kpi"><div class="val">{len(df_skipped)}</div><div class="lbl">Sin cobertura Cecotec</div></div>
</div>
""", unsafe_allow_html=True)

tab_search, tab_results, tab_raw = st.tabs(["🚀 Búsqueda", "📊 Resultados", "📋 Datos cargados"])

with tab_raw:
    st.subheader("Productos relevantes para Cecotec")
    cols_show = ["asin","titulo","fabricante","precio","subcategoria","_source","ranking"]
    st.dataframe(
        df_relevant[cols_show].rename(columns={"_source":"origen","ranking":"rank"}),
        use_container_width=True, hide_index=True,
        column_config={"precio": st.column_config.NumberColumn(format="%.2f €")}
    )
    st.caption(f"{len(df_skipped)} productos descartados (textil, cosmética sin aparatos, etc.)")
    with st.expander("Ver productos descartados"):
        st.dataframe(df_skipped[cols_show], use_container_width=True, hide_index=True)

with tab_search:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("⚙️ Configuración")
    col1, col2, col3 = st.columns(3)
    with col1:
        source_filter = st.multiselect("Origen", ["Hogar","Belleza"], default=["Hogar","Belleza"])
    with col2:
        pmax = float(df_relevant["precio"].max() or 500)
        precio_min, precio_max = st.slider("Precio Amazon (€)", 0.0, pmax, (0.0, pmax), step=5.0)
    with col3:
        max_prods = st.number_input("Máx. productos", 1, len(df_relevant), min(20, len(df_relevant)))
    st.markdown("</div>", unsafe_allow_html=True)

    df_to_process = df_relevant[
        df_relevant["_source"].isin(source_filter) &
        df_relevant["precio"].between(precio_min, precio_max)
    ].head(int(max_prods))

    st.info(f"Se procesarán **{len(df_to_process)}** productos (~{len(df_to_process)*5} segundos estimados).")
    run_btn = st.button("🚀 Iniciar comparación", type="primary", use_container_width=True)

if run_btn:
    results = []
    prog = st.progress(0, text="Preparando…")
    status_box = st.empty()
    total = len(df_to_process)
    for i, (_, row) in enumerate(df_to_process.iterrows()):
        prog.progress(i / total, text=f"[{i+1}/{total}] {row['titulo'][:55]}…")
        status_box.caption(f"🔍 Subcategoría: **{row['subcategoria']}**")
        html = scrape_cecotec(build_query(row))
        alt  = analyze_with_gemini(row, html)
        results.append({"ref": row.to_dict(), "alt": alt})
        time.sleep(0.5)
    prog.progress(1.0, text="✅ Listo")
    status_box.empty()
    st.session_state["results"] = results
    st.success(f"✅ {total} productos comparados. Ve a la pestaña **📊 Resultados**.")
    st.rerun()

with tab_results:
    if "results" not in st.session_state or not st.session_state["results"]:
        st.info("Ejecuta la búsqueda primero.")
        st.stop()

    results = st.session_state["results"]
    rows = []
    encontrados = 0
    ahorro_total = 0.0

    for r in results:
        ref = r["ref"]
        alt = r["alt"]
        if alt.get("no_encontrado"):
            rows.append({
                "ASIN Amazon": ref.get("asin",""),
                "Producto Amazon": ref.get("titulo","")[:80],
                "Marca": ref.get("fabricante",""),
                "Precio Amazon (€)": ref.get("precio"),
                "Subcategoría": ref.get("subcategoria",""),
                "Alternativa Cecotec": "❌ " + alt.get("motivo","No encontrado")[:70],
                "Precio Cecotec (€)": None,
                "Ahorro (€)": None,
                "Prestaciones": "—",
                "Stock": "—",
                "Referencia": "—",
                "URL Cecotec": "",
                "ASIN Cecotec": "",
            })
        else:
            encontrados += 1
            ahorro = alt.get("ahorro_eur") or max(0.0, float(ref.get("precio") or 0) - float(alt.get("cecotec_precio") or 0))
            ahorro_total += float(ahorro or 0)
            pmap = {"mejor":"✅ Mejor","igual":"🟡 Igual","peor":"🔴 Peor"}
            rows.append({
                "ASIN Amazon": ref.get("asin",""),
                "Producto Amazon": ref.get("titulo","")[:80],
                "Marca": ref.get("fabricante",""),
                "Precio Amazon (€)": ref.get("precio"),
                "Subcategoría": ref.get("subcategoria",""),
                "Alternativa Cecotec": alt.get("cecotec_nombre",""),
                "Precio Cecotec (€)": alt.get("cecotec_precio"),
                "Ahorro (€)": round(float(ahorro or 0), 2),
                "Prestaciones": pmap.get(alt.get("prestaciones",""), alt.get("prestaciones","")),
                "Stock": "✅ Sí" if alt.get("cecotec_stock") else "❌ No",
                "Referencia": alt.get("cecotec_referencia",""),
                "URL Cecotec": alt.get("cecotec_url",""),
                "ASIN Cecotec": alt.get("amazon_asin_cecotec",""),
            })

    df_res = pd.DataFrame(rows)

    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi"><div class="val">{encontrados}</div><div class="lbl">Alternativas encontradas</div></div>
      <div class="kpi"><div class="val">{len(results)-encontrados}</div><div class="lbl">Sin alternativa</div></div>
      <div class="kpi"><div class="val">{ahorro_total:.0f} €</div><div class="lbl">Ahorro total acumulado</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.dataframe(
        df_res, use_container_width=True, hide_index=True,
        column_config={
            "Precio Amazon (€)":  st.column_config.NumberColumn(format="%.2f €"),
            "Precio Cecotec (€)": st.column_config.NumberColumn(format="%.2f €"),
            "Ahorro (€)":         st.column_config.NumberColumn(format="%.2f €"),
            "URL Cecotec":        st.column_config.LinkColumn("URL Cecotec"),
        },
    )

    csv = df_res.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Descargar CSV", csv, "comparativa_cecotec_keepa.csv", "text/csv", use_container_width=True)

    st.markdown("---")
    st.subheader("🔎 Detalle por producto")
    for r in [r for r in results if not r["alt"].get("no_encontrado")]:
        ref = r["ref"]
        alt = r["alt"]
        prest = alt.get("prestaciones","")
        tag_css   = {"mejor":"tag-mejor","igual":"tag-igual","peor":"tag-peor"}.get(prest,"tag-skip")
        tag_label = {"mejor":"✅ Mejores prestaciones","igual":"🟡 Equivalente","peor":"🔴 Inferior"}.get(prest, prest)
        with st.expander(f"**{ref.get('titulo','')[:65]}** · {ref.get('fabricante','')} · {ref.get('precio','')}€"):
            c1, mid, c2 = st.columns([5,1,5])
            with c1:
                st.markdown("##### 📦 Amazon")
                st.markdown(f"**{ref.get('titulo','')}**")
                st.markdown(f"*{ref.get('fabricante','')}* · {ref.get('subcategoria','')}")
                st.markdown(f"💶 **{ref.get('precio','')} €**")
                st.markdown(f"_{ref.get('caracteristicas','')[:300]}_")
                if ref.get("url_amazon"):
                    st.markdown(f"[🔗 Amazon]({ref['url_amazon']})")
            with mid:
                st.markdown("<div style='font-size:2rem;text-align:center;margin-top:50px'>→</div>", unsafe_allow_html=True)
            with c2:
                st.markdown("##### 🟥 Cecotec")
                st.markdown(f"**{alt.get('cecotec_nombre','')}**")
                st.markdown(f"💶 **{alt.get('cecotec_precio','')} €**")
                st.markdown(f'<span class="{tag_css}">{tag_label}</span>', unsafe_allow_html=True)
                ahorro = alt.get("ahorro_eur") or max(0, float(ref.get("precio") or 0) - float(alt.get("cecotec_precio") or 0))
                if float(ahorro or 0) > 0:
                    st.markdown(f"💰 **Ahorro: {float(ahorro):.2f} €**")
                st.markdown(f"_{alt.get('cecotec_caracteristicas','')}_")
                if alt.get("cecotec_referencia"):
                    st.markdown(f"🏷️ `{alt['cecotec_referencia']}`")
                if alt.get("cecotec_url"):
                    st.markdown(f"[🔗 Cecotec.es]({alt['cecotec_url']})")
                if alt.get("amazon_asin_cecotec"):
                    st.markdown(f"[🛒 Amazon](https://www.amazon.es/dp/{alt['amazon_asin_cecotec']})")
                if alt.get("justificacion"):
                    st.caption(f"💡 {alt['justificacion']}")
