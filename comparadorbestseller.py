"""
Comparador Cecotec — v3
· Carga feed oficial Cecotec (feed_Espan_a.xlsx) como base de datos local
· Carga Keepa exports para productos competidores
· Matching via Gemini (solo para interpretar similitud, sin scraping)
· Sin límites de red · Sin 429s
"""

import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import json, re, time, os
from pathlib import Path
import google.generativeai as genai

st.set_page_config(page_title="Comparador Cecotec", page_icon="🔍", layout="wide")

# ── CSS corporativo Cecotec ───────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&family=Nunito+Sans:wght@400;600;700&display=swap');
:root {
  --cec-blue:#3EB1C8; --cec-blue-d:#2a8fa3;
  --cec-black:#141413; --cec-bg:#FAF9F5;
  --cec-white:#ffffff; --cec-grey:#e8e8e4; --cec-muted:#6b7280;
}
html,body,[class*="css"]{ font-family:'Nunito Sans',sans-serif; background:var(--cec-bg)!important; color:var(--cec-black); }
.stApp{ background:var(--cec-bg)!important; }
.cec-navbar{ background:var(--cec-black); padding:.7rem 2rem; display:flex; align-items:center; gap:1.2rem; margin:-1rem -1rem 1.8rem -1rem; border-bottom:3px solid var(--cec-blue); }
.cec-navbar .logo-text{ font-family:'Nunito',sans-serif; font-size:1.55rem; font-weight:900; color:#fff; letter-spacing:-.5px; }
.cec-navbar .logo-text span{ color:var(--cec-blue); }
.cec-navbar .logo-sub{ font-size:.72rem; color:rgba(255,255,255,.55); font-weight:600; text-transform:uppercase; letter-spacing:.12em; border-left:1px solid rgba(255,255,255,.2); padding-left:1.1rem; }
.cec-navbar .logo-badge{ margin-left:auto; background:var(--cec-blue); color:var(--cec-black); font-size:.68rem; font-weight:800; padding:3px 10px; border-radius:20px; text-transform:uppercase; letter-spacing:.08em; }
.cec-section-title{ font-family:'Nunito',sans-serif; font-size:1rem; font-weight:800; color:var(--cec-black); text-transform:uppercase; letter-spacing:.08em; padding:.4rem 0 .4rem .7rem; border-left:4px solid var(--cec-blue); margin:1.4rem 0 .9rem; background:linear-gradient(90deg,rgba(62,177,200,.07) 0%,transparent 100%); }
.kpi-row{ display:flex; gap:.8rem; margin-bottom:1.6rem; flex-wrap:wrap; }
.kpi{ background:var(--cec-white); border-radius:10px; padding:1rem 1.3rem; box-shadow:0 1px 3px rgba(20,20,19,.08),0 0 0 1px var(--cec-grey); flex:1; min-width:120px; position:relative; overflow:hidden; }
.kpi::after{ content:''; position:absolute; bottom:0; left:0; right:0; height:3px; background:var(--cec-blue); }
.kpi .val{ font-size:2rem; font-weight:900; color:var(--cec-blue); font-family:'Nunito',sans-serif; line-height:1; }
.kpi .lbl{ font-size:.68rem; color:var(--cec-muted); text-transform:uppercase; letter-spacing:.07em; margin-top:.3rem; font-weight:600; }
.tag-mejor{ background:rgba(62,177,200,.12); color:#0a6a7a; padding:3px 12px; border-radius:4px; font-size:.76rem; font-weight:800; border:1px solid rgba(62,177,200,.35); text-transform:uppercase; }
.tag-igual{ background:rgba(255,193,7,.12); color:#856404; padding:3px 12px; border-radius:4px; font-size:.76rem; font-weight:800; border:1px solid rgba(255,193,7,.35); text-transform:uppercase; }
.tag-peor{ background:rgba(220,53,69,.1); color:#842029; padding:3px 12px; border-radius:4px; font-size:.76rem; font-weight:800; border:1px solid rgba(220,53,69,.3); text-transform:uppercase; }
.tag-skip{ background:var(--cec-grey); color:var(--cec-muted); padding:3px 12px; border-radius:4px; font-size:.76rem; font-weight:800; text-transform:uppercase; }
div[data-testid="stTabs"] button[aria-selected="true"]{ color:var(--cec-blue)!important; border-bottom-color:var(--cec-blue)!important; }
div[data-testid="stButton"] button[kind="primary"]{ background:var(--cec-blue)!important; border-color:var(--cec-blue)!important; color:#fff!important; font-weight:800; border-radius:6px; }
div[data-testid="stButton"] button[kind="primary"]:hover{ background:var(--cec-blue-d)!important; }
div[data-testid="stDataFrame"] thead th{ background:var(--cec-black)!important; color:#fff!important; font-weight:700; font-size:.78rem; text-transform:uppercase; }
.stProgress > div > div{ background:var(--cec-blue)!important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="cec-navbar">
  <div class="logo-text">ceco<span>tec</span></div>
  <div class="logo-sub">Comparador de competencia</div>
  <div class="logo-badge">✨ Powered by IA</div>
</div>
""", unsafe_allow_html=True)

# ── API Key ───────────────────────────────────────────────────────────────────
def get_gemini_key():
    if st.session_state.get("GOOGLE_API_KEY"): return st.session_state["GOOGLE_API_KEY"]
    try:
        k = st.secrets.get("GOOGLE_API_KEY", None)
        if k: return k
    except Exception: pass
    return os.environ.get("GOOGLE_API_KEY", None)

api_key = get_gemini_key()
if not api_key:
    st.warning("### 🔑 Configura tu API key de Google Gemini")
    st.markdown("1. Ve a **[aistudio.google.com/apikey](https://aistudio.google.com/apikey)**\n2. Crea una key gratuita\n3. Pégala aquí:")
    k = st.text_input("Google API Key", type="password", placeholder="AIzaSy...")
    if st.button("✅ Guardar", type="primary"):
        if k and k.startswith("AIza"):
            st.session_state["GOOGLE_API_KEY"] = k
            st.rerun()
        else:
            st.error("Key inválida.")
    st.stop()

genai.configure(api_key=api_key)

@st.cache_resource
def get_model():
    return genai.GenerativeModel("gemini-2.0-flash")
model = get_model()

# ── Constantes ────────────────────────────────────────────────────────────────
KEEPA_HOGAR   = "KeepaExport-2026-05-19-BestSellersList-9-599391031.xlsx"
KEEPA_BELLEZA = "BellezaKeepaExport-2026-05-19-BestSellersList-9-4347698031-9000.xlsx"
FEED_FILE     = "feed_Espan_a.xlsx"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
    "Accept-Language": "es-ES,es;q=0.9",
}

CECOTEC_CATS = {
    "Aspiradoras escoba","Aspiradoras de mano","Robots aspiradores","Aspiradoras para alfombras",
    "Aspiradoras verticales","Freidoras de aire","Freidoras","Hornos de sobremesa","Tostadoras",
    "Sandwicheras","Grills de contacto","Batidoras de mano","Batidoras de vaso",
    "Procesadores de alimentos","Cafeteras italianas","Cafeteras de filtro","Cafeteras espresso",
    "Cafeteras express","Máquinas de café","Planchas de vapor","Planchas de vapor verticales para viaje",
    "Cepillos de vapor","Centros de planchado","Balanzas digitales","Básculas de cocina","Básculas de baño",
    "Batidoras amasadoras","Purificadores de aire","Humidificadores","Ventiladores","Ventiladores de pie",
    "Aires acondicionados portátiles","Televisores","Monitores","Proyectores","Altavoces portátiles",
    "Desincrustantes","Secadores de pelo","Planchas para el pelo","Planchas de pelo","Rizadores",
    "Cepillos eléctricos para el cabello","Afeitadoras eléctricas","Depiladores",
    "Cepillos de dientes eléctricos","Masajeadores","Freidoras sin aceite",
}

# ── Carga feed Cecotec ────────────────────────────────────────────────────────
@st.cache_data
def load_cecotec_feed(upload_dir: str) -> pd.DataFrame:
    p = Path(upload_dir.rstrip("/"))
    path = p if p.suffix in (".xlsx",".xls") else p / FEED_FILE
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_excel(path)
    # Limpiar HTML de descripciones
    def strip_html(s):
        return re.sub(r"<[^>]+>", " ", str(s or "")).strip()
    df["desc_clean"] = df["alternate_description"].apply(strip_html)
    df["desc_clean"] = df["desc_clean"].where(df["desc_clean"].str.len() > 10,
                        df["title"].astype(str))
    # Solo productos principales en stock (sin repuestos/recambios)
    df_stock = df[
        (df["availability"] == "in stock") &
        (~df["categories"].str.lower().str.contains(
            "repuesto|recambio", na=False))
    ].copy()
    df_stock["price"] = pd.to_numeric(df_stock["price"], errors="coerce")
    df_stock["sale_price"] = pd.to_numeric(df_stock["sale_price"], errors="coerce")
    df_stock["precio_final"] = df_stock["sale_price"].fillna(df_stock["price"])
    return df_stock.reset_index(drop=True)

# ── Carga Keepa ───────────────────────────────────────────────────────────────
@st.cache_data
def load_keepa_files(upload_dir: str) -> pd.DataFrame:
    dfs = []
    for fname, source in [(KEEPA_HOGAR,"Hogar"),(KEEPA_BELLEZA,"Belleza")]:
        p = Path(upload_dir) / fname
        if not p.exists(): continue
        df = pd.read_excel(p)
        df["_source"] = source
        dfs.append(df)
    if not dfs: return pd.DataFrame()
    return _process_keepa(pd.concat(dfs, ignore_index=True))

def load_custom_file(file) -> pd.DataFrame:
    df = pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
    df["_source"] = "Fichero propio"
    return _process_keepa(df)

def _process_keepa(merged: pd.DataFrame) -> pd.DataFrame:
    merged = merged.rename(columns={
        "ASIN":"asin","Título":"titulo","Caja de Compra: Actual":"precio",
        "Categorías: Principal":"categoria_principal","Categorías: Subcategoría":"subcategoria",
        "Clasificación de Ventas: Actual":"ranking","Fabricante":"fabricante",
        "URL: Amazon":"url_amazon","Códigos de producto: EAN":"ean",
        "Opiniones: Valoraciones":"rating","Opiniones: Cantidad de valoraciones":"num_reviews",
        "Descripción & Características: Característica 1":"feat1",
        "Descripción & Características: Característica 2":"feat2",
        "Descripción & Características: Característica 3":"feat3",
        "Descripción & Características: Característica 4":"feat4",
        "Descripción & Características: Descripción breve":"descripcion_breve",
    })
    for col in ["asin","titulo","precio","subcategoria","ranking","fabricante","url_amazon","feat1","feat2","feat3","feat4","descripcion_breve"]:
        if col not in merged.columns: merged[col] = ""
    if "ranking" in merged.columns:
        merged = merged.sort_values("ranking")
        if "_source" in merged.columns:
            merged = merged.groupby("_source").head(100).reset_index(drop=True)
    merged["_cecotec_relevant"] = merged["subcategoria"].apply(
        lambda s: any(cat.lower() in str(s).lower() for cat in CECOTEC_CATS)
    )
    def build_feats(row):
        parts = [str(row.get(f"feat{i}","") or "") for i in range(1,5)]
        parts.append(str(row.get("descripcion_breve","") or ""))
        return " | ".join(p[:120] for p in parts if p.strip())[:500]
    merged["caracteristicas"] = merged.apply(build_feats, axis=1)
    return merged

# ── Matching via Gemini ───────────────────────────────────────────────────────
MATCH_PROMPT = """Eres un experto en electrónica y electrodomésticos.

Producto de referencia (competidor):
{ref}

Lista de productos Cecotec disponibles en la misma categoría (con stock y precio inferior):
{candidatos}

Selecciona el producto Cecotec que mejor iguale o supere las prestaciones del producto de referencia.
Responde SOLO con JSON válido, sin markdown:

Si hay un buen match:
{{"idx": 0, "prestaciones": "mejor|igual|peor", "justificacion": "max 120 chars explicando por qué"}}

Si ningún producto Cecotec es adecuado:
{{"idx": -1, "justificacion": "razón breve"}}

El campo "idx" es el índice (0-based) del producto en la lista de candidatos."""

def find_best_match(ref: dict, df_cecotec: pd.DataFrame) -> dict:
    """Filter Cecotec candidates by category similarity + price, then ask Gemini to pick best."""
    precio_ref = float(ref.get("precio") or 0)
    subcat_ref = str(ref.get("subcategoria","")).lower()
    titulo_ref = str(ref.get("titulo","")).lower()

    # Filter candidates: same category family + cheaper
    def cat_score(row):
        cat = str(row["categories"]).lower()
        desc = str(row["desc_clean"]).lower()
        # keyword overlap between ref subcategory/title and cecotec category/desc
        words = set(re.findall(r'\w{4,}', subcat_ref + " " + titulo_ref))
        matches = sum(1 for w in words if w in cat or w in desc)
        return matches

    df_candidates = df_cecotec.copy()
    if precio_ref > 0:
        df_candidates = df_candidates[df_candidates["precio_final"] < precio_ref]

    if df_candidates.empty:
        return {"no_encontrado": True, "motivo": "No hay productos Cecotec más baratos en este rango"}

    df_candidates = df_candidates.copy()
    df_candidates["_score"] = df_candidates.apply(cat_score, axis=1)
    df_candidates = df_candidates[df_candidates["_score"] > 0].sort_values("_score", ascending=False).head(15)

    if df_candidates.empty:
        # Fallback: just use cheapest in rough category
        df_candidates = df_cecotec[df_cecotec["precio_final"] < precio_ref].head(10) if precio_ref > 0 else df_cecotec.head(10)

    if df_candidates.empty:
        return {"no_encontrado": True, "motivo": "Sin candidatos Cecotec disponibles"}

    # Build candidates list for Gemini
    candidatos_list = []
    for i, (_, row) in enumerate(df_candidates.iterrows()):
        candidatos_list.append(
            f"{i}. [{row['categories']}] {row['title']} — {row['precio_final']}€ | {row['desc_clean'][:120]}"
        )

    ref_str = json.dumps({
        "titulo": ref.get("titulo",""),
        "marca": ref.get("fabricante",""),
        "precio": precio_ref,
        "categoria": ref.get("subcategoria",""),
        "caracteristicas": ref.get("caracteristicas","")[:300],
    }, ensure_ascii=False)

    prompt = MATCH_PROMPT.format(ref=ref_str, candidatos="\n".join(candidatos_list))

    for intento in range(3):
        try:
            resp = model.generate_content(prompt)
            raw = re.sub(r"```json|```", "", resp.text).strip()
            result = json.loads(raw)
            idx = result.get("idx", -1)
            if idx == -1:
                return {"no_encontrado": True, "motivo": result.get("justificacion","Sin match")}
            # Get the matched product
            matched_row = df_candidates.iloc[idx]
            return {
                "cecotec_nombre": matched_row["title"],
                "cecotec_precio": float(matched_row["precio_final"]),
                "cecotec_precio_original": float(matched_row["price"]) if matched_row["price"] != matched_row["precio_final"] else None,
                "cecotec_caracteristicas": matched_row["desc_clean"][:200],
                "cecotec_url": matched_row["link"],
                "cecotec_referencia": str(matched_row.get("mpn","") or ""),
                "cecotec_stock": True,
                "cecotec_imagen": matched_row.get("image_link",""),
                "cecotec_categoria": matched_row["categories"],
                "ahorro_eur": round(precio_ref - float(matched_row["precio_final"]), 2) if precio_ref > 0 else 0,
                "prestaciones": result.get("prestaciones","igual"),
                "justificacion": result.get("justificacion",""),
            }
        except Exception as e:
            msg = str(e)
            if "429" in msg or "quota" in msg.lower():
                time.sleep(15 * (intento + 1))
                continue
            return {"no_encontrado": True, "motivo": msg[:150]}
    return {"no_encontrado": True, "motivo": "Límite de cuota API. Espera 1 min."}

# ── Render results ────────────────────────────────────────────────────────────
def render_results(results):
    rows, encontrados, ahorro_total = [], 0, 0.0
    for r in results:
        ref, alt = r["ref"], r["alt"]
        titulo = ref.get("titulo","")
        if alt.get("no_encontrado"):
            rows.append({
                "ASIN": ref.get("asin",""),
                "Producto competidor": titulo[:80],
                "Marca": ref.get("fabricante",""),
                "Precio comp. (€)": ref.get("precio"),
                "Categoría": ref.get("subcategoria",""),
                "Alternativa Cecotec": "❌ " + alt.get("motivo","")[:60],
                "Precio Cecotec (€)": None, "Ahorro (€)": None,
                "Prestaciones":"—","Ref. Cecotec":"—","URL Cecotec":"",
            })
        else:
            encontrados += 1
            ahorro = float(alt.get("ahorro_eur") or 0)
            ahorro_total += ahorro
            pmap = {"mejor":"✅ Mejor","igual":"🟡 Igual","peor":"🔴 Peor"}
            rows.append({
                "ASIN": ref.get("asin",""),
                "Producto competidor": titulo[:80],
                "Marca": ref.get("fabricante",""),
                "Precio comp. (€)": ref.get("precio"),
                "Categoría": ref.get("subcategoria",""),
                "Alternativa Cecotec": alt.get("cecotec_nombre",""),
                "Precio Cecotec (€)": alt.get("cecotec_precio"),
                "Ahorro (€)": round(ahorro, 2),
                "Prestaciones": pmap.get(alt.get("prestaciones",""), alt.get("prestaciones","")),
                "Ref. Cecotec": alt.get("cecotec_referencia",""),
                "URL Cecotec": alt.get("cecotec_url",""),
            })

    df_res = pd.DataFrame(rows)
    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi"><div class="val">{encontrados}</div><div class="lbl">Alternativas encontradas</div></div>
      <div class="kpi"><div class="val">{len(results)-encontrados}</div><div class="lbl">Sin alternativa</div></div>
      <div class="kpi"><div class="val">{ahorro_total:.0f} €</div><div class="lbl">Ahorro total acumulado</div></div>
    </div>""", unsafe_allow_html=True)

    st.dataframe(df_res, use_container_width=True, hide_index=True, column_config={
        "Precio comp. (€)":  st.column_config.NumberColumn(format="%.2f €"),
        "Precio Cecotec (€)":st.column_config.NumberColumn(format="%.2f €"),
        "Ahorro (€)":        st.column_config.NumberColumn(format="%.2f €"),
        "URL Cecotec":       st.column_config.LinkColumn("URL Cecotec"),
    })
    st.download_button("⬇️ Descargar CSV", df_res.to_csv(index=False).encode("utf-8"),
                       "comparativa_cecotec.csv","text/csv", use_container_width=True)

    st.markdown("---")
    st.markdown('<div class="cec-section-title">🔎 Detalle por producto</div>', unsafe_allow_html=True)
    for r in [r for r in results if not r["alt"].get("no_encontrado")]:
        ref, alt = r["ref"], r["alt"]
        prest = alt.get("prestaciones","")
        tag_css   = {"mejor":"tag-mejor","igual":"tag-igual","peor":"tag-peor"}.get(prest,"tag-skip")
        tag_label = {"mejor":"✅ Mejores prestaciones","igual":"🟡 Equivalente","peor":"🔴 Inferior"}.get(prest,prest)
        ahorro = float(alt.get("ahorro_eur") or 0)
        with st.expander(f"**{ref.get('titulo','')[:65]}** · {ref.get('fabricante','')} · {ref.get('precio','')}€  →  **{alt.get('cecotec_nombre','')}** · {alt.get('cecotec_precio','')}€"):
            c1, mid, c2 = st.columns([5,1,5])
            with c1:
                st.markdown("##### 📦 Producto competidor")
                st.markdown(f"**{ref.get('titulo','')}**")
                st.markdown(f"*{ref.get('fabricante','')}* · {ref.get('subcategoria','')}")
                st.markdown(f"💶 **{ref.get('precio','')} €**")
                st.markdown(f"_{ref.get('caracteristicas','')[:300]}_")
                if ref.get("url_amazon"):
                    st.markdown(f"[🔗 Ver en Amazon]({ref['url_amazon']})")
            with mid:
                st.markdown("<div style='font-size:2rem;text-align:center;margin-top:50px'>→</div>", unsafe_allow_html=True)
            with c2:
                st.markdown("##### 🟦 Alternativa Cecotec")
                st.markdown(f"**{alt.get('cecotec_nombre','')}**")
                p_orig = alt.get("cecotec_precio_original")
                if p_orig:
                    st.markdown(f"~~{p_orig}€~~ → 💶 **{alt.get('cecotec_precio','')} €**")
                else:
                    st.markdown(f"💶 **{alt.get('cecotec_precio','')} €**")
                st.markdown(f'<span class="{tag_css}">{tag_label}</span>', unsafe_allow_html=True)
                if ahorro > 0:
                    st.markdown(f"💰 **Ahorro: {ahorro:.2f} €**")
                st.markdown(f"_{alt.get('cecotec_caracteristicas','')}_")
                st.caption(f"📂 {alt.get('cecotec_categoria','')}  ·  🏷️ Ref: {alt.get('cecotec_referencia','')}")
                if alt.get("cecotec_url"):
                    st.markdown(f"[🔗 Ver en Cecotec.es]({alt['cecotec_url']})")
                if alt.get("justificacion"):
                    st.caption(f"💡 {alt['justificacion']}")

def run_search(df_to_process, df_cecotec):
    results = []
    prog = st.progress(0, text="Preparando…")
    status = st.empty()
    total = len(df_to_process)
    for i, (_, row) in enumerate(df_to_process.iterrows()):
        prog.progress(i/total, text=f"[{i+1}/{total}] {str(row.get('titulo',''))[:50]}…")
        status.caption(f"🔍 Buscando en feed Cecotec · **{row.get('subcategoria','')}**")
        alt = find_best_match(row.to_dict(), df_cecotec)
        results.append({"ref": row.to_dict(), "alt": alt})
        time.sleep(1.5)  # Rate limit: ~40 RPM safe
    prog.progress(1.0, text="✅ Completado")
    status.empty()
    return results

# ── Load data ─────────────────────────────────────────────────────────────────
UPLOAD_DIR = Path(__file__).parent
# Use session feed path if uploaded via UI, else look in project folder
_feed_path = st.session_state.get("feed_path")
if _feed_path and Path(_feed_path).exists():
    df_cecotec = load_cecotec_feed(_feed_path)
else:
    df_cecotec = load_cecotec_feed(str(UPLOAD_DIR))
df_keepa   = load_keepa_files(str(UPLOAD_DIR))

# Feed status bar
feed_ok = not df_cecotec.empty
keepa_ok = not df_keepa.empty

col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    if feed_ok:
        st.success(f"✅ Feed Cecotec: **{len(df_cecotec):,}** productos en stock")
    else:
        st.error("❌ Feed Cecotec no encontrado (`feed_Espan_a.xlsx`)")
with col_f2:
    if keepa_ok:
        df_rel = df_keepa[df_keepa["_cecotec_relevant"]]
        st.success(f"✅ Keepa: **{len(df_rel)}** productos comparables")
    else:
        st.warning("⚠️ Archivos Keepa no encontrados")
with col_f3:
    st.info(f"⚡ Modo: **Sin scraping** · Matching local + IA")

if not feed_ok:
    st.markdown('<div class="cec-section-title">📂 Subir catálogo Cecotec</div>', unsafe_allow_html=True)
    st.info("Sube el fichero `feed_Espan_a.xlsx` para iniciar la app. Solo se hace una vez — queda en memoria de sesión.")
    feed_upload = st.file_uploader("feed_Espan_a.xlsx", type=["xlsx","xls"], key="feed_upload")
    if feed_upload:
        with st.spinner("Cargando catálogo Cecotec…"):
            import tempfile, shutil
            tmp = Path(tempfile.mkdtemp()) / "feed_Espan_a.xlsx"
            tmp.write_bytes(feed_upload.read())
            df_cecotec = load_cecotec_feed(str(tmp.parent) + "/")
            # Guardar ruta temporal en session para persistencia dentro de la sesión
            st.session_state["feed_path"] = str(tmp)
        st.success(f"✅ Catálogo cargado: {len(df_cecotec):,} productos en stock")
        st.rerun()
    else:
        st.stop()

# ── TABS ──────────────────────────────────────────────────────────────────────
tab_keepa, tab_manual, tab_fichero, tab_resultados, tab_feed = st.tabs([
    "📦 Keepa Bestsellers", "✏️ Producto manual",
    "📂 Subir fichero", "📊 Resultados", "🗄️ Feed Cecotec",
])

# ═══════════════════════════════════════
# TAB 1 · KEEPA
# ═══════════════════════════════════════
with tab_keepa:
    if not keepa_ok:
        st.warning("Archivos Keepa no encontrados en la carpeta del proyecto.")
    else:
        df_relevant = df_keepa[df_keepa["_cecotec_relevant"]].copy()
        df_skipped  = df_keepa[~df_keepa["_cecotec_relevant"]].copy()
        st.markdown('<div class="cec-section-title">📊 Bestsellers Amazon · Análisis de competencia</div>', unsafe_allow_html=True)
        st.markdown(f"""<div class="kpi-row">
          <div class="kpi"><div class="val">{len(df_keepa[df_keepa['_source']=='Hogar'])}</div><div class="lbl">Top-100 Hogar</div></div>
          <div class="kpi"><div class="val">{len(df_keepa[df_keepa['_source']=='Belleza'])}</div><div class="lbl">Top-100 Belleza</div></div>
          <div class="kpi"><div class="val">{len(df_relevant)}</div><div class="lbl">Con equiv. Cecotec</div></div>
          <div class="kpi"><div class="val">{len(df_skipped)}</div><div class="lbl">Sin cobertura</div></div>
        </div>""", unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)
        with c1:
            src = st.multiselect("Origen", ["Hogar","Belleza"], default=["Hogar","Belleza"], key="k_src")
        with c2:
            pmax = float(df_relevant["precio"].max() or 500)
            pmin_v, pmax_v = st.slider("Precio competidor (€)", 0.0, pmax, (0.0, pmax), step=5.0, key="k_p")
        with c3:
            n = st.number_input("Máx. productos", 1, len(df_relevant), min(20, len(df_relevant)), key="k_n")

        df_proc = df_relevant[
            df_relevant["_source"].isin(src) &
            df_relevant["precio"].between(pmin_v, pmax_v)
        ].head(int(n))

        st.info(f"Se procesarán **{len(df_proc)}** productos (~{len(df_proc)*2} s · matching local sin scraping).")
        if st.button("🚀 Buscar alternativas Cecotec", type="primary", use_container_width=True, key="btn_k"):
            st.session_state["results"] = run_search(df_proc, df_cecotec)
            st.success("✅ Completado. Ve a **📊 Resultados**.")
            st.rerun()

# ═══════════════════════════════════════
# TAB 2 · MANUAL
# ═══════════════════════════════════════
with tab_manual:
    st.markdown('<div class="cec-section-title">✏️ Buscar alternativa para un producto concreto</div>', unsafe_allow_html=True)
    st.caption("Introduce el ASIN — el resto se obtiene automáticamente de Amazon.")

    col_asin, col_price = st.columns([3,1])
    with col_asin:
        m_asin = st.text_input("ASIN Amazon *", placeholder="ej: B0BY9592V9")
    with col_price:
        m_precio_override = st.number_input("Precio (€) opcional", min_value=0.0, step=0.01, value=0.0)

    with st.expander("➕ Datos adicionales (opcional — se autorellenan con el ASIN)"):
        cx1, cx2 = st.columns(2)
        with cx1:
            m_titulo = st.text_input("Nombre", placeholder="Se obtiene del ASIN")
            m_marca  = st.text_input("Marca", placeholder="Se obtiene del ASIN")
        with cx2:
            m_subcat = st.text_input("Categoría", placeholder="Se obtiene del ASIN")
            m_caract = st.text_area("Características", height=90, placeholder="Se obtienen del ASIN")

    if st.button("🔍 Buscar en Cecotec", type="primary", use_container_width=True, key="btn_m"):
        asin = m_asin.strip().upper()
        if not asin:
            st.error("El ASIN es obligatorio.")
        else:
            amz_data = {"titulo":m_titulo,"fabricante":m_marca,"subcategoria":m_subcat,
                        "caracteristicas":m_caract,"precio":m_precio_override or 0,
                        "asin":asin,"url_amazon":f"https://www.amazon.es/dp/{asin}","feat1":m_caract}
            with st.spinner(f"Obteniendo datos de Amazon para **{asin}**…"):
                try:
                    r = requests.get(f"https://www.amazon.es/dp/{asin}", headers=HEADERS, timeout=14)
                    soup = BeautifulSoup(r.text, "html.parser")
                    if not amz_data["titulo"]:
                        t = soup.select_one("#productTitle,#title")
                        if t: amz_data["titulo"] = t.get_text(strip=True)[:200]
                    if not amz_data["fabricante"]:
                        b = soup.select_one("#bylineInfo,.po-brand .a-span9")
                        if b: amz_data["fabricante"] = re.sub(r"Marca:|Visita la Store de","",b.get_text(strip=True)).strip()[:80]
                    if not amz_data["precio"]:
                        p = soup.select_one(".a-price .a-offscreen,#priceblock_ourprice")
                        if p:
                            try: amz_data["precio"] = float(re.sub(r"[^\d.]","", p.get_text().replace(",",".")))
                            except: pass
                    if not amz_data["caracteristicas"]:
                        bullets = soup.select("#feature-bullets .a-list-item")
                        amz_data["caracteristicas"] = " | ".join(b.get_text(strip=True) for b in bullets[:5])[:500]
                    if not amz_data["subcategoria"]:
                        crumbs = soup.select("#wayfinding-breadcrumbs_feature_div li")
                        if crumbs: amz_data["subcategoria"] = crumbs[-1].get_text(strip=True)
                    amz_data["feat1"] = amz_data["caracteristicas"]
                except Exception as e:
                    st.warning(f"Amazon no accesible ({e}). Usando datos manuales.")

            if amz_data["titulo"]:
                st.success(f"✅ **{amz_data['titulo'][:60]}** · {amz_data['fabricante']} · {amz_data['precio']}€")

            if not amz_data["titulo"] and not amz_data["subcategoria"]:
                st.error("Sin datos suficientes. Rellena nombre y categoría manualmente.")
            else:
                with st.spinner("Buscando en feed Cecotec…"):
                    alt = find_best_match(amz_data, df_cecotec)
                    result = [{"ref": amz_data, "alt": alt}]
                    st.session_state["results_manual"] = result
                render_results(result)

    elif "results_manual" in st.session_state:
        render_results(st.session_state["results_manual"])

# ═══════════════════════════════════════
# TAB 3 · FICHERO
# ═══════════════════════════════════════
with tab_fichero:
    st.markdown('<div class="cec-section-title">📂 Subir fichero de productos competidores</div>', unsafe_allow_html=True)
    st.markdown("Sube un **CSV o Excel** con tus productos. Columnas: `titulo`, `fabricante`, `precio`, `subcategoria`, `caracteristicas` (o formato Keepa completo).")
    uploaded = st.file_uploader("CSV o Excel", type=["csv","xlsx","xls"], key="cfile")
    if uploaded:
        try:
            df_custom = load_custom_file(uploaded)
            df_rel_c  = df_custom[df_custom["_cecotec_relevant"]].copy()
            st.success(f"✅ {len(df_custom)} productos · {len(df_rel_c)} con equivalente Cecotec posible")
            st.dataframe(df_custom[["titulo","fabricante","precio","subcategoria","_cecotec_relevant"]].rename(
                columns={"_cecotec_relevant":"relevante"}), use_container_width=True, hide_index=True,
                column_config={"precio":st.column_config.NumberColumn(format="%.2f €")})
            ca, cb = st.columns(2)
            with ca: solo = st.checkbox("Solo con equiv. Cecotec posible", value=True)
            with cb: mx = st.number_input("Máx. productos", 1, len(df_custom), min(20, len(df_custom)), key="mx_c")
            df_p = (df_rel_c if solo else df_custom).head(int(mx))
            st.info(f"Se procesarán **{len(df_p)}** productos.")
            if st.button("🚀 Buscar alternativas", type="primary", use_container_width=True, key="btn_c"):
                st.session_state["results_custom"] = run_search(df_p, df_cecotec)
                st.success("✅ Completado. Ve a **📊 Resultados**.")
                st.rerun()
        except Exception as e:
            st.error(f"Error: {e}")

# ═══════════════════════════════════════
# TAB 4 · RESULTADOS
# ═══════════════════════════════════════
with tab_resultados:
    res_key = None
    if st.session_state.get("results"):       res_key = "results"
    if st.session_state.get("results_custom"):
        if res_key:
            op = st.radio("Mostrar:", ["Keepa Bestsellers","Fichero propio"], horizontal=True)
            res_key = "results" if op == "Keepa Bestsellers" else "results_custom"
        else:
            res_key = "results_custom"
    if not res_key:
        st.info("Ejecuta una búsqueda en alguna de las pestañas anteriores.")
    else:
        render_results(st.session_state[res_key])

# ═══════════════════════════════════════
# TAB 5 · FEED CECOTEC
# ═══════════════════════════════════════
with tab_feed:
    st.markdown('<div class="cec-section-title">🗄️ Catálogo Cecotec cargado</div>', unsafe_allow_html=True)
    st.markdown(f"""<div class="kpi-row">
      <div class="kpi"><div class="val">{len(df_cecotec):,}</div><div class="lbl">Productos en stock</div></div>
      <div class="kpi"><div class="val">{df_cecotec['categories'].nunique()}</div><div class="lbl">Categorías</div></div>
      <div class="kpi"><div class="val">{df_cecotec['precio_final'].min():.0f}–{df_cecotec['precio_final'].max():.0f} €</div><div class="lbl">Rango de precios</div></div>
      <div class="kpi"><div class="val">{(df_cecotec['sale_price'] != df_cecotec['price']).sum()}</div><div class="lbl">En oferta</div></div>
    </div>""", unsafe_allow_html=True)

    cat_filter = st.multiselect("Filtrar por categoría", sorted(df_cecotec["categories"].unique()), key="feed_cat")
    search_term = st.text_input("Buscar en título/descripción", key="feed_search")
    df_show = df_cecotec.copy()
    if cat_filter: df_show = df_show[df_show["categories"].isin(cat_filter)]
    if search_term: df_show = df_show[df_show["title"].str.contains(search_term, case=False, na=False) | df_show["desc_clean"].str.contains(search_term, case=False, na=False)]

    st.caption(f"{len(df_show)} productos mostrados")
    st.dataframe(df_show[["title","categories","precio_final","price","link","desc_clean"]].rename(
        columns={"title":"Producto","categories":"Categoría","precio_final":"Precio (€)",
                 "price":"Precio orig.","link":"URL","desc_clean":"Descripción"}),
        use_container_width=True, hide_index=True,
        column_config={
            "Precio (€)":   st.column_config.NumberColumn(format="%.2f €"),
            "Precio orig.": st.column_config.NumberColumn(format="%.2f €"),
            "URL":          st.column_config.LinkColumn("URL"),
        })
