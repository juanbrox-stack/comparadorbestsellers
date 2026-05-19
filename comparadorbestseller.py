"""
Comparador Cecotec — Keepa Bestsellers (Hogar + Belleza)
Motor IA: Google Gemini Flash (gratuito via Google AI Studio)
Modos: Keepa automático · Producto manual · Subir fichero propio
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

# ── Cecotec brand palette ────────────────────────────────────────────────────
# Azul Cecotec: #3EB1C8  |  Negro: #141413  |  Fondo: #FAF9F5
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&family=Nunito+Sans:wght@400;600;700&display=swap');

/* ── Variables corporativas ── */
:root {
  --cec-blue:   #3EB1C8;
  --cec-blue-d: #2a8fa3;
  --cec-black:  #141413;
  --cec-bg:     #FAF9F5;
  --cec-white:  #ffffff;
  --cec-grey:   #e8e8e4;
  --cec-text:   #141413;
  --cec-muted:  #6b7280;
}

/* ── Reset & base ── */
html, body, [class*="css"] {
    font-family: 'Nunito Sans', sans-serif;
    background: var(--cec-bg) !important;
    color: var(--cec-black);
}
.stApp { background: var(--cec-bg) !important; }

/* ── Navbar ── */
.cec-navbar {
    background: var(--cec-black);
    padding: .7rem 2rem;
    display: flex;
    align-items: center;
    gap: 1.2rem;
    margin: -1rem -1rem 1.8rem -1rem;
    border-bottom: 3px solid var(--cec-blue);
}
.cec-navbar .logo-text {
    font-family: 'Nunito', sans-serif;
    font-size: 1.55rem;
    font-weight: 900;
    color: var(--cec-white);
    letter-spacing: -.5px;
}
.cec-navbar .logo-text span { color: var(--cec-blue); }
.cec-navbar .logo-sub {
    font-size: .72rem;
    color: rgba(255,255,255,.55);
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: .12em;
    border-left: 1px solid rgba(255,255,255,.2);
    padding-left: 1.1rem;
}
.cec-navbar .logo-badge {
    margin-left: auto;
    background: var(--cec-blue);
    color: var(--cec-black);
    font-size: .68rem;
    font-weight: 800;
    padding: 3px 10px;
    border-radius: 20px;
    text-transform: uppercase;
    letter-spacing: .08em;
}

/* ── Section titles ── */
.cec-section-title {
    font-family: 'Nunito', sans-serif;
    font-size: 1rem;
    font-weight: 800;
    color: var(--cec-black);
    text-transform: uppercase;
    letter-spacing: .08em;
    padding: .4rem 0 .4rem .7rem;
    border-left: 4px solid var(--cec-blue);
    margin: 1.4rem 0 .9rem;
    background: linear-gradient(90deg, rgba(62,177,200,.07) 0%, transparent 100%);
}

/* ── KPI row ── */
.kpi-row { display:flex; gap:.8rem; margin-bottom:1.6rem; flex-wrap:wrap; }
.kpi {
    background: var(--cec-white);
    border-radius: 10px;
    padding: 1rem 1.3rem;
    box-shadow: 0 1px 3px rgba(20,20,19,.08), 0 0 0 1px var(--cec-grey);
    flex: 1;
    min-width: 120px;
    position: relative;
    overflow: hidden;
}
.kpi::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 3px;
    background: var(--cec-blue);
}
.kpi .val {
    font-size: 2rem;
    font-weight: 900;
    color: var(--cec-blue);
    font-family: 'Nunito', sans-serif;
    line-height: 1;
}
.kpi .lbl {
    font-size: .68rem;
    color: var(--cec-muted);
    text-transform: uppercase;
    letter-spacing: .07em;
    margin-top: .3rem;
    font-weight: 600;
}

/* ── Cards ── */
.card {
    background: var(--cec-white);
    border-radius: 10px;
    padding: 1.2rem 1.4rem;
    box-shadow: 0 1px 3px rgba(20,20,19,.08), 0 0 0 1px var(--cec-grey);
    margin-bottom: 1rem;
}

/* ── Tags de prestaciones ── */
.tag-mejor {
    background: rgba(62,177,200,.12);
    color: #0a6a7a;
    padding: 3px 12px;
    border-radius: 4px;
    font-size: .76rem;
    font-weight: 800;
    border: 1px solid rgba(62,177,200,.35);
    text-transform: uppercase;
    letter-spacing: .04em;
}
.tag-igual {
    background: rgba(255,193,7,.12);
    color: #856404;
    padding: 3px 12px;
    border-radius: 4px;
    font-size: .76rem;
    font-weight: 800;
    border: 1px solid rgba(255,193,7,.35);
    text-transform: uppercase;
    letter-spacing: .04em;
}
.tag-peor {
    background: rgba(220,53,69,.1);
    color: #842029;
    padding: 3px 12px;
    border-radius: 4px;
    font-size: .76rem;
    font-weight: 800;
    border: 1px solid rgba(220,53,69,.3);
    text-transform: uppercase;
    letter-spacing: .04em;
}
.tag-skip {
    background: var(--cec-grey);
    color: var(--cec-muted);
    padding: 3px 12px;
    border-radius: 4px;
    font-size: .76rem;
    font-weight: 800;
    border: 1px solid #d0d0cc;
    text-transform: uppercase;
    letter-spacing: .04em;
}

/* ── Producto detail ── */
.prod-cec {
    background: rgba(62,177,200,.05);
    border: 2px solid var(--cec-blue);
    border-radius: 10px;
    padding: 1rem;
}

/* ── Streamlit overrides ── */
div[data-testid="stTabs"] button[data-baseweb="tab"] {
    font-weight: 700;
    font-size: .85rem;
    color: var(--cec-muted);
}
div[data-testid="stTabs"] button[aria-selected="true"] {
    color: var(--cec-blue) !important;
    border-bottom-color: var(--cec-blue) !important;
}
div[data-testid="stButton"] button[kind="primary"] {
    background: var(--cec-blue) !important;
    border-color: var(--cec-blue) !important;
    color: var(--cec-white) !important;
    font-weight: 800;
    border-radius: 6px;
    font-family: 'Nunito', sans-serif;
}
div[data-testid="stButton"] button[kind="primary"]:hover {
    background: var(--cec-blue-d) !important;
    border-color: var(--cec-blue-d) !important;
}
div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
div[data-testid="stDataFrame"] thead th {
    background: var(--cec-black) !important;
    color: var(--cec-white) !important;
    font-weight: 700;
    font-size: .78rem;
    text-transform: uppercase;
    letter-spacing: .05em;
}
.stProgress > div > div { background: var(--cec-blue) !important; }
div[data-testid="stMetric"] { background: var(--cec-white); border-radius:8px; padding:.8rem; }
</style>
""", unsafe_allow_html=True)

# ── Navbar ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="cec-navbar">
  <div class="logo-text">ceco<span>tec</span></div>
  <div class="logo-sub">Comparador de competencia</div>
  <div class="logo-badge">✨ Powered by IA</div>
</div>
""", unsafe_allow_html=True)

# ── API Key ───────────────────────────────────────────────────────────────────
def get_gemini_key():
    if "GOOGLE_API_KEY" in st.session_state and st.session_state["GOOGLE_API_KEY"]:
        return st.session_state["GOOGLE_API_KEY"]
    try:
        k = st.secrets.get("GOOGLE_API_KEY", None)
        if k: return k
    except Exception:
        pass
    return os.environ.get("GOOGLE_API_KEY", None)

api_key = get_gemini_key()

if not api_key:

    st.warning("### 🔑 Configura tu API key de Google Gemini (gratuita)")
    st.markdown("1. Ve a **[aistudio.google.com/apikey](https://aistudio.google.com/apikey)**\n2. Clic en **'Create API Key'**\n3. Pégala aquí 👇")
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
    return genai.GenerativeModel("gemini-2.0-flash")

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
    "Máquinas de café", "Planchas de vapor", "Planchas de vapor verticales para viaje",
    "Cepillos de vapor", "Centros de planchado",
    "Balanzas digitales", "Básculas de cocina", "Básculas de baño",
    "Batidoras amasadoras", "Purificadores de aire", "Humidificadores", "Ventiladores",
    "Aires acondicionados portátiles", "Televisores", "Monitores", "Proyectores",
    "Altavoces portátiles", "Desincrustantes",
    "Secadores de pelo", "Planchas para el pelo", "Rizadores",
    "Cepillos eléctricos para el cabello", "Afeitadoras eléctricas", "Depiladores",
    "Cepillos de dientes eléctricos", "Masajeadores",
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "es-ES,es;q=0.9",
}

# ── Data loading ──────────────────────────────────────────────────────────────
@st.cache_data
def load_keepa_files(upload_dir: str) -> pd.DataFrame:
    dfs = []
    for fname, source in [(KEEPA_HOGAR, "Hogar"), (KEEPA_BELLEZA, "Belleza")]:
        path = Path(upload_dir) / fname
        if not path.exists():
            continue
        df = pd.read_excel(path)
        df["_source"] = source
        dfs.append(df)
    if not dfs:
        return pd.DataFrame()
    return _process_keepa(pd.concat(dfs, ignore_index=True))

def load_custom_file(file) -> pd.DataFrame:
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    df["_source"] = "Fichero propio"
    return _process_keepa(df)

def _process_keepa(merged: pd.DataFrame) -> pd.DataFrame:
    merged = merged.rename(columns={
        "ASIN": "asin", "Título": "titulo",
        "Caja de Compra: Actual": "precio",
        "Categorías: Principal": "categoria_principal",
        "Categorías: Subcategoría": "subcategoria",
        "Clasificación de Ventas: Actual": "ranking",
        "Fabricante": "fabricante", "URL: Amazon": "url_amazon",
        "Códigos de producto: EAN": "ean",
        "Opiniones: Valoraciones": "rating",
        "Opiniones: Cantidad de valoraciones": "num_reviews",
        "Descripción & Características: Característica 1": "feat1",
        "Descripción & Características: Característica 2": "feat2",
        "Descripción & Características: Característica 3": "feat3",
        "Descripción & Características: Característica 4": "feat4",
        "Descripción & Características: Descripción breve": "descripcion_breve",
    })
    for col in ["asin","titulo","precio","subcategoria","ranking","fabricante","url_amazon","feat1","feat2","feat3","feat4","descripcion_breve"]:
        if col not in merged.columns:
            merged[col] = ""
    if "ranking" in merged.columns:
        merged = merged.sort_values("ranking")
        if "_source" in merged.columns:
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

# ── Scraping + AI ─────────────────────────────────────────────────────────────
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

PROMPT_TEMPLATE = """Eres un experto en electrónica de hogar y pequeño electrodoméstico español.

Producto de referencia:
{ref}

HTML de resultados cecotec.es:
{html}

Identifica el mejor producto Cecotec que: esté en stock, iguale o supere prestaciones y tenga precio INFERIOR a {precio} €.
Responde SOLO con JSON válido, sin markdown.

Si encuentras alternativa:
{{"cecotec_nombre":"...","cecotec_precio":0.0,"cecotec_caracteristicas":"specs clave","cecotec_url":"https://www.cecotec.es/...","cecotec_referencia":null,"cecotec_stock":true,"amazon_asin_cecotec":null,"ahorro_eur":0.0,"prestaciones":"mejor|igual|peor","justificacion":"max 120 chars"}}

Si NO hay alternativa válida:
{{"no_encontrado":true,"motivo":"razón breve"}}"""

def analyze_with_gemini(row, html: str) -> dict:
    if isinstance(row, pd.Series):
        ref = {"titulo": row.get("titulo",""), "fabricante": row.get("fabricante",""),
               "subcategoria": row.get("subcategoria",""), "caracteristicas": row.get("caracteristicas","")}
        precio = row.get("precio", 0)
    else:
        ref = row
        precio = row.get("precio", 0)
    prompt = PROMPT_TEMPLATE.format(ref=json.dumps(ref, ensure_ascii=False), html=html[:4000], precio=precio)
    try:
        resp = model.generate_content(prompt)
        raw = re.sub(r"```json|```", "", resp.text).strip()
        return json.loads(raw)
    except Exception as e:
        return {"no_encontrado": True, "motivo": str(e)[:150]}

def build_query(row) -> str:
    if isinstance(row, pd.Series):
        subcat = str(row.get("subcategoria","")).split(",")[0].strip()
        feat   = str(row.get("feat1","") or "")[:60]
    else:
        subcat = str(row.get("subcategoria","")).split(",")[0].strip()
        feat   = str(row.get("caracteristicas",""))[:60]
    return f"{subcat} {feat}".strip() or str(row.get("titulo",""))[:60]

# ── Shared results renderer ───────────────────────────────────────────────────
def render_results(results):
    rows = []
    encontrados = 0
    ahorro_total = 0.0
    for r in results:
        ref = r["ref"]
        alt = r["alt"]
        titulo = ref.get("titulo","") if isinstance(ref, dict) else ref.get("titulo","")
        if alt.get("no_encontrado"):
            rows.append({
                "ASIN Amazon": ref.get("asin",""), "Producto": titulo[:80],
                "Marca": ref.get("fabricante",""), "Precio ref. (€)": ref.get("precio"),
                "Subcategoría": ref.get("subcategoria",""),
                "Alternativa Cecotec": "❌ " + alt.get("motivo","No encontrado")[:70],
                "Precio Cecotec (€)": None, "Ahorro (€)": None,
                "Prestaciones":"—","Stock":"—","Referencia":"—","URL Cecotec":"","ASIN Cecotec":"",
            })
        else:
            encontrados += 1
            ahorro = alt.get("ahorro_eur") or max(0.0, float(ref.get("precio") or 0) - float(alt.get("cecotec_precio") or 0))
            ahorro_total += float(ahorro or 0)
            pmap = {"mejor":"✅ Mejor","igual":"🟡 Igual","peor":"🔴 Peor"}
            rows.append({
                "ASIN Amazon": ref.get("asin",""), "Producto": titulo[:80],
                "Marca": ref.get("fabricante",""), "Precio ref. (€)": ref.get("precio"),
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
    st.dataframe(df_res, use_container_width=True, hide_index=True, column_config={
        "Precio ref. (€)":    st.column_config.NumberColumn(format="%.2f €"),
        "Precio Cecotec (€)": st.column_config.NumberColumn(format="%.2f €"),
        "Ahorro (€)":         st.column_config.NumberColumn(format="%.2f €"),
        "URL Cecotec":        st.column_config.LinkColumn("URL Cecotec"),
    })
    csv = df_res.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Descargar CSV", csv, "comparativa_cecotec.csv", "text/csv", use_container_width=True)

    st.markdown("---")
    st.subheader("🔎 Detalle por producto")
    for r in [r for r in results if not r["alt"].get("no_encontrado")]:
        ref = r["ref"]
        alt = r["alt"]
        prest = alt.get("prestaciones","")
        tag_css   = {"mejor":"tag-mejor","igual":"tag-igual","peor":"tag-peor"}.get(prest,"tag-skip")
        tag_label = {"mejor":"✅ Mejores prestaciones","igual":"🟡 Equivalente","peor":"🔴 Inferior"}.get(prest, prest)
        titulo = ref.get("titulo","")
        with st.expander(f"**{titulo[:65]}** · {ref.get('fabricante','')} · {ref.get('precio','')}€"):
            c1, mid, c2 = st.columns([5,1,5])
            with c1:
                st.markdown("##### 📦 Referencia")
                st.markdown(f"**{titulo}**")
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

def run_search(df_to_process):
    results = []
    prog = st.progress(0, text="Preparando…")
    status_box = st.empty()
    total = len(df_to_process)
    for i, (_, row) in enumerate(df_to_process.iterrows()):
        prog.progress(i / total, text=f"[{i+1}/{total}] {str(row.get('titulo',''))[:55]}…")
        status_box.caption(f"🔍 **{row.get('subcategoria','')}**")
        html = scrape_cecotec(build_query(row))
        alt  = analyze_with_gemini(row, html)
        results.append({"ref": row.to_dict(), "alt": alt})
        time.sleep(0.5)
    prog.progress(1.0, text="✅ Listo")
    status_box.empty()
    return results

# ── HEADER ────────────────────────────────────────────────────────────────────


# ── TABS ──────────────────────────────────────────────────────────────────────
tab_keepa, tab_manual, tab_fichero, tab_resultados, tab_raw = st.tabs([
    "📦 Keepa Bestsellers",
    "✏️ Producto manual",
    "📂 Subir fichero",
    "📊 Resultados",
    "📋 Datos cargados",
])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 · KEEPA
# ═══════════════════════════════════════════════════════════════════════════════
with tab_keepa:
    UPLOAD_DIR = Path(__file__).parent
    df_all = load_keepa_files(str(UPLOAD_DIR))

    if df_all.empty:
        st.warning("No se encontraron los archivos Keepa en la carpeta del proyecto.")
    else:
        df_relevant = df_all[df_all["_cecotec_relevant"]].copy()
        df_skipped  = df_all[~df_all["_cecotec_relevant"]].copy()

        st.markdown('<div class="cec-section-title">📊 Bestsellers Amazon · Análisis de competencia</div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="kpi-row">
          <div class="kpi"><div class="val">{len(df_all[df_all['_source']=='Hogar'])}</div><div class="lbl">Top-100 Hogar</div></div>
          <div class="kpi"><div class="val">{len(df_all[df_all['_source']=='Belleza'])}</div><div class="lbl">Top-100 Belleza</div></div>
          <div class="kpi"><div class="val">{len(df_relevant)}</div><div class="lbl">Con equiv. Cecotec</div></div>
          <div class="kpi"><div class="val">{len(df_skipped)}</div><div class="lbl">Sin cobertura</div></div>
        </div>
        """, unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)
        with col1:
            source_filter = st.multiselect("Origen", ["Hogar","Belleza"], default=["Hogar","Belleza"], key="k_source")
        with col2:
            pmax = float(df_relevant["precio"].max() or 500)
            precio_min, precio_max = st.slider("Precio Amazon (€)", 0.0, pmax, (0.0, pmax), step=5.0, key="k_precio")
        with col3:
            max_prods = st.number_input("Máx. productos", 1, len(df_relevant), min(20, len(df_relevant)), key="k_max")

        df_to_process = df_relevant[
            df_relevant["_source"].isin(source_filter) &
            df_relevant["precio"].between(precio_min, precio_max)
        ].head(int(max_prods))

        st.info(f"Se procesarán **{len(df_to_process)}** productos (~{len(df_to_process)*5} s estimados).")

        if st.button("🚀 Buscar alternativas Cecotec", type="primary", use_container_width=True, key="btn_keepa"):
            results = run_search(df_to_process)
            st.session_state["results"] = results
            st.success(f"✅ {len(results)} productos comparados. Ve a **📊 Resultados**.")
            st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 · PRODUCTO MANUAL
# ═══════════════════════════════════════════════════════════════════════════════
with tab_manual:
    st.markdown('<div class="cec-section-title">✏️ Buscar alternativa para un producto concreto</div>', unsafe_allow_html=True)
    st.caption("Introduce los datos del producto competidor que quieres comparar con Cecotec.")

    c1, c2 = st.columns(2)
    with c1:
        m_titulo   = st.text_input("Nombre del producto *", placeholder="ej: Rowenta Pure Pop Cepillo vapor 1300W")
        m_marca    = st.text_input("Marca *", placeholder="ej: Rowenta")
        m_precio   = st.number_input("Precio (€) *", min_value=0.01, step=0.01, value=29.99)
    with c2:
        m_subcat   = st.text_input("Categoría / tipo de producto *", placeholder="ej: Planchas de vapor")
        m_caract   = st.text_area("Características principales", height=120,
                                  placeholder="ej: 1300W, 20 g/min vapor, cabezales reversibles, elimina pelusas")
        m_asin     = st.text_input("ASIN Amazon (opcional)", placeholder="ej: B0BY9592V9")

    if st.button("🔍 Buscar en Cecotec", type="primary", use_container_width=True, key="btn_manual"):
        if not m_titulo or not m_marca or not m_subcat:
            st.error("Rellena al menos nombre, marca y categoría.")
        else:
            with st.spinner(f"Buscando alternativa para **{m_titulo}**…"):
                row_dict = {
                    "titulo": m_titulo, "fabricante": m_marca, "precio": m_precio,
                    "subcategoria": m_subcat, "caracteristicas": m_caract,
                    "asin": m_asin, "url_amazon": f"https://www.amazon.es/dp/{m_asin}" if m_asin else "",
                    "feat1": m_caract,
                }
                html = scrape_cecotec(f"{m_subcat} {m_caract[:60]}")
                alt  = analyze_with_gemini(row_dict, html)
                result = [{"ref": row_dict, "alt": alt}]
                st.session_state["results_manual"] = result
            st.success("✅ Búsqueda completada. Resultado:")
            render_results(result)

    elif "results_manual" in st.session_state:
        render_results(st.session_state["results_manual"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 · SUBIR FICHERO
# ═══════════════════════════════════════════════════════════════════════════════
with tab_fichero:
    st.markdown('<div class="cec-section-title">📂 Subir fichero de productos competidores</div>', unsafe_allow_html=True)
    st.markdown("""
    Sube un **CSV o Excel** con tus propios productos. Columnas soportadas:
    - Las mismas que exporta Keepa (se detectan automáticamente)
    - O columnas simples: `titulo`, `fabricante`, `precio`, `subcategoria`, `caracteristicas`
    """)

    uploaded = st.file_uploader("CSV o Excel", type=["csv","xlsx","xls"], key="custom_file")

    if uploaded:
        try:
            df_custom = load_custom_file(uploaded)
            df_rel_custom = df_custom[df_custom["_cecotec_relevant"]].copy()
            df_all_custom = df_custom.copy()

            st.success(f"✅ {len(df_custom)} productos cargados · {len(df_rel_custom)} con posible equivalente Cecotec")
            st.dataframe(
                df_custom[["titulo","fabricante","precio","subcategoria","_cecotec_relevant"]].rename(
                    columns={"_cecotec_relevant":"relevante"}),
                use_container_width=True, hide_index=True,
                column_config={"precio": st.column_config.NumberColumn(format="%.2f €")}
            )

            col_a, col_b = st.columns(2)
            with col_a:
                solo_relevantes = st.checkbox("Solo productos con equivalente Cecotec posible", value=True)
            with col_b:
                max_custom = st.number_input("Máx. productos a procesar", 1,
                                             len(df_custom), min(20, len(df_custom)), key="max_custom")

            df_proc = (df_rel_custom if solo_relevantes else df_all_custom).head(int(max_custom))
            st.info(f"Se procesarán **{len(df_proc)}** productos.")

            if st.button("🚀 Buscar alternativas Cecotec", type="primary", use_container_width=True, key="btn_custom"):
                results = run_search(df_proc)
                st.session_state["results_custom"] = results
                st.success(f"✅ {len(results)} productos comparados. Ve a **📊 Resultados**.")
                st.rerun()

        except Exception as e:
            st.error(f"Error al leer el fichero: {e}")

    if "results_custom" in st.session_state and not uploaded:
        st.info("Sube un fichero para iniciar la comparación.")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4 · RESULTADOS
# ═══════════════════════════════════════════════════════════════════════════════
with tab_resultados:
    # Mostrar los resultados más recientes disponibles
    res_key = None
    if "results" in st.session_state and st.session_state["results"]:
        res_key = "results"
    if "results_custom" in st.session_state and st.session_state["results_custom"]:
        # Mostrar el más reciente; si ambos existen, dejar elegir
        if res_key:
            opcion = st.radio("Mostrar resultados de:", ["Keepa Bestsellers", "Fichero propio"], horizontal=True)
            res_key = "results" if opcion == "Keepa Bestsellers" else "results_custom"
        else:
            res_key = "results_custom"

    if not res_key:
        st.info("Ejecuta una búsqueda en alguna de las pestañas anteriores.")
    else:
        render_results(st.session_state[res_key])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 5 · DATOS CARGADOS
# ═══════════════════════════════════════════════════════════════════════════════
with tab_raw:
    UPLOAD_DIR2 = Path(__file__).parent
    df_all2 = load_keepa_files(str(UPLOAD_DIR2))
    if df_all2.empty:
        st.warning("No se encontraron los archivos Keepa.")
    else:
        df_relevant2 = df_all2[df_all2["_cecotec_relevant"]].copy()
        df_skipped2  = df_all2[~df_all2["_cecotec_relevant"]].copy()
        cols_show = ["asin","titulo","fabricante","precio","subcategoria","_source","ranking"]
        st.markdown('<div class="cec-section-title">Productos con equivalente Cecotec posible</div>', unsafe_allow_html=True)
        st.dataframe(
            df_relevant2[cols_show].rename(columns={"_source":"origen","ranking":"rank"}),
            use_container_width=True, hide_index=True,
            column_config={"precio": st.column_config.NumberColumn(format="%.2f €")}
        )
        st.caption(f"{len(df_skipped2)} productos descartados (textil, cosmética sin aparatos, etc.)")
        with st.expander("Ver productos descartados"):
            st.dataframe(df_skipped2[cols_show], use_container_width=True, hide_index=True)
