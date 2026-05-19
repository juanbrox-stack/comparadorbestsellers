"""
Comparador Cecotec v4
· Matching 100% local (pandas) — sin IA, sin scraping, sin límites
· Carga feed oficial Cecotec + Keepa exports
· Procesa TODOS los productos Keepa relevantes en segundos
"""

import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import re, time, os
from pathlib import Path
import google.generativeai as genai

st.set_page_config(page_title="Comparador Cecotec", page_icon="🔍", layout="wide")

# ── CSS corporativo ───────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&family=Nunito+Sans:wght@400;600;700&display=swap');
:root{--cec-blue:#3EB1C8;--cec-blue-d:#2a8fa3;--cec-black:#141413;--cec-bg:#FAF9F5;--cec-white:#fff;--cec-grey:#e8e8e4;--cec-muted:#6b7280;}
html,body,[class*="css"]{font-family:'Nunito Sans',sans-serif;background:var(--cec-bg)!important;color:var(--cec-black);}
.stApp{background:var(--cec-bg)!important;}
.cec-navbar{background:var(--cec-black);padding:.7rem 2rem;display:flex;align-items:center;gap:1.2rem;margin:-1rem -1rem 1.8rem -1rem;border-bottom:3px solid var(--cec-blue);}
.cec-navbar .logo-text{font-family:'Nunito',sans-serif;font-size:1.55rem;font-weight:900;color:#fff;letter-spacing:-.5px;}
.cec-navbar .logo-text span{color:var(--cec-blue);}
.cec-navbar .logo-sub{font-size:.72rem;color:rgba(255,255,255,.55);font-weight:600;text-transform:uppercase;letter-spacing:.12em;border-left:1px solid rgba(255,255,255,.2);padding-left:1.1rem;}
.cec-navbar .logo-badge{margin-left:auto;background:var(--cec-blue);color:var(--cec-black);font-size:.68rem;font-weight:800;padding:3px 10px;border-radius:20px;text-transform:uppercase;letter-spacing:.08em;}
.cec-section-title{font-family:'Nunito',sans-serif;font-size:1rem;font-weight:800;color:var(--cec-black);text-transform:uppercase;letter-spacing:.08em;padding:.4rem 0 .4rem .7rem;border-left:4px solid var(--cec-blue);margin:1.4rem 0 .9rem;background:linear-gradient(90deg,rgba(62,177,200,.07) 0%,transparent 100%);}
.kpi-row{display:flex;gap:.8rem;margin-bottom:1.6rem;flex-wrap:wrap;}
.kpi{background:var(--cec-white);border-radius:10px;padding:1rem 1.3rem;box-shadow:0 1px 3px rgba(20,20,19,.08),0 0 0 1px var(--cec-grey);flex:1;min-width:120px;position:relative;overflow:hidden;}
.kpi::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;background:var(--cec-blue);}
.kpi .val{font-size:2rem;font-weight:900;color:var(--cec-blue);font-family:'Nunito',sans-serif;line-height:1;}
.kpi .lbl{font-size:.68rem;color:var(--cec-muted);text-transform:uppercase;letter-spacing:.07em;margin-top:.3rem;font-weight:600;}
.tag-mejor{background:rgba(62,177,200,.12);color:#0a6a7a;padding:3px 12px;border-radius:4px;font-size:.76rem;font-weight:800;border:1px solid rgba(62,177,200,.35);text-transform:uppercase;}
.tag-igual{background:rgba(255,193,7,.12);color:#856404;padding:3px 12px;border-radius:4px;font-size:.76rem;font-weight:800;border:1px solid rgba(255,193,7,.35);text-transform:uppercase;}
.tag-peor{background:rgba(220,53,69,.1);color:#842029;padding:3px 12px;border-radius:4px;font-size:.76rem;font-weight:800;border:1px solid rgba(220,53,69,.3);text-transform:uppercase;}
.tag-skip{background:var(--cec-grey);color:var(--cec-muted);padding:3px 12px;border-radius:4px;font-size:.76rem;font-weight:800;text-transform:uppercase;}
div[data-testid="stTabs"] button[aria-selected="true"]{color:var(--cec-blue)!important;border-bottom-color:var(--cec-blue)!important;}
div[data-testid="stButton"] button[kind="primary"]{background:var(--cec-blue)!important;border-color:var(--cec-blue)!important;color:#fff!important;font-weight:800;border-radius:6px;}
div[data-testid="stButton"] button[kind="primary"]:hover{background:var(--cec-blue-d)!important;}
div[data-testid="stDataFrame"] thead th{background:var(--cec-black)!important;color:#fff!important;font-weight:700;font-size:.78rem;text-transform:uppercase;}
.stProgress > div > div{background:var(--cec-blue)!important;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="cec-navbar">
  <div class="logo-text">ceco<span>tec</span></div>
  <div class="logo-sub">Comparador de competencia</div>
  <div class="logo-badge">⚡ Matching local</div>
</div>
""", unsafe_allow_html=True)

# ── Constantes ────────────────────────────────────────────────────────────────
KEEPA_HOGAR   = "KeepaExport-2026-05-19-BestSellersList-9-599391031.xlsx"
KEEPA_BELLEZA = "BellezaKeepaExport-2026-05-19-BestSellersList-9-4347698031-9000.xlsx"
FEED_FILE     = "feed_Espan_a.xlsx"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
    "Accept-Language": "es-ES,es;q=0.9",
}

# Mapa de categorías Keepa → categorías Cecotec (puede ser lista para cubrir variantes)
CAT_MAP = {
    "aspiradoras escoba":            ["aspiradores verticales"],
    "aspiradoras de mano":           ["aspiradores de mano"],
    "robots aspiradores":            ["robots aspiradores"],
    "aspiradoras para alfombras":    ["aspiradores de trineo"],
    "aspiradoras con bolsa":         ["aspiradores de trineo"],
    "freidoras de aire":             ["freidoras sin aceite"],
    "freidoras":                     ["freidoras sin aceite"],
    "hornos de sobremesa":           ["microondas de sobremesa","hornos"],
    "tostadoras":                    ["tostadoras"],
    "sandwicheras":                  ["sandwicheras","grills"],
    "grills de contacto":            ["grills","sandwicheras"],
    "batidoras de mano":             ["batidoras de mano"],
    "batidoras de vaso":             ["batidoras de vaso"],
    "procesadores de alimentos":     ["batidoras / picadoras","robots de cocina"],
    "batidoras amasadoras":          ["amasadoras","robots de cocina"],
    "cafeteras italianas":           ["cafeteras express","cafeteras"],
    "cafeteras de filtro":           ["cafeteras de filtro","cafeteras"],
    "cafeteras espresso":            ["cafeteras express"],
    "máquinas de café":              ["cafeteras express","cafeteras"],
    "planchas de vapor":             ["centros de planchado","planchas de vapor"],
    "planchas de vapor verticales para viaje": ["planchas de vapor","centros de planchado"],
    "cepillos de vapor":             ["planchas de vapor"],
    "centros de planchado":          ["centros de planchado"],
    "balanzas digitales":            ["básculas de cocina","básculas"],
    "básculas de cocina":            ["básculas de cocina"],
    "básculas de baño":              ["básculas de baño"],
    "purificadores de aire":         ["purificadores de aire"],
    "humidificadores":               ["humidificadores"],
    "ventiladores":                  ["ventiladores de pie","ventiladores de techo","ventiladores"],
    "aires acondicionados portátiles":["aires acondicionados"],
    "televisores":                   ["televisores / smart tv"],
    "monitores":                     ["monitores"],
    "altavoces portátiles":          ["altavoces"],
    "desincrustantes":               ["repuestos cafeteras","accesorios"],
    "secadores de pelo":             ["secadores de pelo"],
    "planchas para el pelo":         ["planchas de pelo"],
    "planchas de pelo":              ["planchas de pelo"],
    "rizadores":                     ["rizadores"],
    "cepillos eléctricos para el cabello": ["cepillos alisadores","planchas de pelo"],
    "afeitadoras eléctricas":        ["afeitadoras","depilación"],
    "depiladores":                   ["depilación","depiladores"],
    "cepillos de dientes eléctricos":["cepillos de dientes"],
    "masajeadores":                  ["masajeadores"],
    "freidoras sin aceite":          ["freidoras sin aceite"],
    "aspiradoras verticales":        ["aspiradores verticales"],
    # Categorías adicionales Hogar
    "ventiladores de techo":         ["ventiladores de techo"],
    "microondas sencillos":          ["microondas de sobremesa"],
    "cafeteras individuales":        ["cafeteras express","cafeteras"],
    "deshumidificadores":            ["deshumidificadores"],
    "sartenes para freír":           ["sartenes","utensilios de cocina"],
    "juegos de sartenes":            ["sartenes","utensilios de cocina"],
    "hervidores":                    ["hervidores"],
    "robots de cocina":              ["robots de cocina"],
    "lavavajillas":                  ["lavavajillas"],
    "lavadoras":                     ["lavadoras"],
    "frigoríficos":                  ["frigoríficos combi","frigoríficos americanos"],
    "minibar":                       ["minibar / mini nevera"],
    "campanas extractoras":          ["campanas extractoras"],
    "vinotecas":                     ["vinoteca"],
    "hornos":                        ["hornos integrables","microondas de sobremesa"],
}

# ── Carga feed Cecotec ────────────────────────────────────────────────────────
@st.cache_data
def load_cecotec_feed(path_str: str) -> pd.DataFrame:
    p = Path(path_str.rstrip("/"))
    path = p if p.suffix in (".xlsx",".xls") else p / FEED_FILE
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_excel(path)
    def strip_html(s):
        return re.sub(r"<[^>]+>", " ", str(s or "")).strip()
    df["desc_clean"] = df["alternate_description"].apply(strip_html)
    df["desc_clean"] = df["desc_clean"].where(df["desc_clean"].str.len() > 10, df["title"].astype(str))
    df_stock = df[
        (df["availability"] == "in stock") &
        (~df["categories"].str.lower().str.contains("repuesto|recambio", na=False))
    ].copy()
    df_stock["price"]       = pd.to_numeric(df_stock["price"], errors="coerce")
    df_stock["sale_price"]  = pd.to_numeric(df_stock["sale_price"], errors="coerce")
    df_stock["precio_final"]= df_stock["sale_price"].fillna(df_stock["price"])
    df_stock["cat_lower"]   = df_stock["categories"].str.lower().fillna("")
    df_stock["title_lower"] = df_stock["title"].str.lower().fillna("")
    df_stock["desc_lower"]  = df_stock["desc_clean"].str.lower().fillna("")
    return df_stock.reset_index(drop=True)

# ── Carga Keepa (sin límite de 100) ──────────────────────────────────────────
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

def _process_keepa(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={
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
    for col in ["asin","titulo","precio","subcategoria","ranking","fabricante",
                "url_amazon","feat1","feat2","feat3","feat4","descripcion_breve"]:
        if col not in df.columns: df[col] = ""
    if "ranking" in df.columns:
        df = df.sort_values("ranking")
    # No limit — process ALL products
    df["_cecotec_relevant"] = df["subcategoria"].apply(
        lambda s: str(s).lower().strip() in CAT_MAP
    )
    def build_feats(row):
        parts = [str(row.get(f"feat{i}","") or "") for i in range(1,5)]
        parts.append(str(row.get("descripcion_breve","") or ""))
        return " | ".join(p[:120] for p in parts if p.strip())[:500]
    df["caracteristicas"] = df.apply(build_feats, axis=1)
    return df

# ── Matching local 100% pandas ────────────────────────────────────────────────
def find_best_match_local(ref: dict, df_cec: pd.DataFrame) -> dict:
    precio_ref  = float(ref.get("precio") or 0)
    subcat_ref  = str(ref.get("subcategoria","")).lower().strip()
    titulo_ref  = str(ref.get("titulo","")).lower()
    feats_ref   = str(ref.get("caracteristicas","")).lower()

    # 1. Get Cecotec categories for this Keepa subcategory
    cec_cats = CAT_MAP.get(subcat_ref, [])
    if not cec_cats:
        # Fuzzy fallback: find any partial word match
        words = [w for w in re.findall(r'\w{5,}', subcat_ref)]
        cec_cats = [c for c in df_cec["cat_lower"].unique()
                    if any(w in c for w in words)][:3]
    if not cec_cats:
        return {"no_encontrado": True, "motivo": f"Categoría '{subcat_ref}' sin equivalente en Cecotec"}

    # 2. Filter by category
    mask_cat = df_cec["cat_lower"].apply(lambda c: any(cc in c for cc in cec_cats))
    df_cat = df_cec[mask_cat].copy()

    if df_cat.empty:
        return {"no_encontrado": True, "motivo": f"Sin productos Cecotec en categorías: {', '.join(cec_cats)}"}

    # 3. Filter by price (cheaper than reference)
    if precio_ref > 0:
        df_cheap = df_cat[df_cat["precio_final"] < precio_ref].copy()
        if df_cheap.empty:
            # Relax: allow up to 10% more expensive
            df_cheap = df_cat[df_cat["precio_final"] <= precio_ref * 1.10].copy()
        df_cat = df_cheap

    if df_cat.empty:
        return {"no_encontrado": True, "motivo": f"No hay productos Cecotec más baratos en esta categoría"}

    # 4. Score by keyword overlap with title + features
    ref_words = set(re.findall(r'\w{4,}', titulo_ref + " " + feats_ref))
    def score_row(row):
        haystack = row["title_lower"] + " " + row["desc_lower"]
        return sum(1 for w in ref_words if w in haystack)

    df_cat = df_cat.copy()
    df_cat["_score"] = df_cat.apply(score_row, axis=1)
    df_cat = df_cat.sort_values(["_score","precio_final"], ascending=[False, True])

    best = df_cat.iloc[0]
    precio_cec = float(best["precio_final"])
    ahorro = round(precio_ref - precio_cec, 2) if precio_ref > 0 else 0.0

    # Prestaciones heuristic: if Cecotec is cheaper → "igual" by default
    # (conservative: we don't have full specs to claim "mejor")
    prestaciones = "igual"
    if ahorro > precio_ref * 0.2:
        prestaciones = "mejor"   # significantly cheaper → value win
    elif precio_cec > precio_ref:
        prestaciones = "peor"

    return {
        "cecotec_nombre":        best["title"],
        "cecotec_precio":        precio_cec,
        "cecotec_precio_original": float(best["price"]) if best["price"] != best["precio_final"] else None,
        "cecotec_caracteristicas": best["desc_clean"][:200],
        "cecotec_url":           best["link"],
        "cecotec_referencia":    str(best.get("mpn","") or ""),
        "cecotec_stock":         True,
        "cecotec_categoria":     best["categories"],
        "cecotec_imagen":        best.get("image_link",""),
        "ahorro_eur":            ahorro,
        "prestaciones":          prestaciones,
        "justificacion":         f"Mejor opción en '{best['categories']}' · score={int(best['_score'])}",
    }

# ── Render resultados ─────────────────────────────────────────────────────────
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
                "Subcategoría": ref.get("subcategoria",""),
                "Alternativa Cecotec": "❌ " + alt.get("motivo","")[:70],
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
                "Subcategoría": ref.get("subcategoria",""),
                "Alternativa Cecotec": alt.get("cecotec_nombre",""),
                "Precio Cecotec (€)": alt.get("cecotec_precio"),
                "Ahorro (€)": round(ahorro, 2),
                "Prestaciones": pmap.get(alt.get("prestaciones",""), alt.get("prestaciones","")),
                "Ref. Cecotec": alt.get("cecotec_referencia",""),
                "URL Cecotec": alt.get("cecotec_url",""),
            })

    df_res = pd.DataFrame(rows)
    st.markdown(f"""<div class="kpi-row">
      <div class="kpi"><div class="val">{encontrados}</div><div class="lbl">Alternativas encontradas</div></div>
      <div class="kpi"><div class="val">{len(results)-encontrados}</div><div class="lbl">Sin alternativa</div></div>
      <div class="kpi"><div class="val">{ahorro_total:.0f} €</div><div class="lbl">Ahorro total acumulado</div></div>
      <div class="kpi"><div class="val">{encontrados*100//len(results) if results else 0}%</div><div class="lbl">Tasa de cobertura</div></div>
    </div>""", unsafe_allow_html=True)

    st.dataframe(df_res, use_container_width=True, hide_index=True, column_config={
        "Precio comp. (€)":  st.column_config.NumberColumn(format="%.2f €"),
        "Precio Cecotec (€)":st.column_config.NumberColumn(format="%.2f €"),
        "Ahorro (€)":        st.column_config.NumberColumn(format="%.2f €"),
        "URL Cecotec":       st.column_config.LinkColumn("URL Cecotec"),
    })
    st.download_button("⬇️ Descargar CSV", df_res.to_csv(index=False).encode("utf-8"),
                       "comparativa_cecotec.csv","text/csv", use_container_width=True)

    found = [r for r in results if not r["alt"].get("no_encontrado")]
    if found:
        st.markdown("---")
        st.markdown('<div class="cec-section-title">🔎 Detalle por producto</div>', unsafe_allow_html=True)
        for r in found:
            ref, alt = r["ref"], r["alt"]
            prest = alt.get("prestaciones","")
            tag_css   = {"mejor":"tag-mejor","igual":"tag-igual","peor":"tag-peor"}.get(prest,"tag-skip")
            tag_label = {"mejor":"✅ Mejor precio y valor","igual":"🟡 Equivalente","peor":"🔴 Inferior"}.get(prest,prest)
            ahorro = float(alt.get("ahorro_eur") or 0)
            with st.expander(
                f"**{ref.get('titulo','')[:60]}** · {ref.get('precio','')}€  →  "
                f"**{alt.get('cecotec_nombre','')[:50]}** · {alt.get('cecotec_precio','')}€"
                + (f"  💰 -{ahorro:.2f}€" if ahorro > 0 else "")
            ):
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

def run_search_local(df_proc: pd.DataFrame, df_cec: pd.DataFrame) -> list:
    """Instant local matching — no API calls, no waiting."""
    results = []
    prog = st.progress(0)
    total = len(df_proc)
    for i, (_, row) in enumerate(df_proc.iterrows()):
        prog.progress((i+1)/total, text=f"[{i+1}/{total}] {str(row.get('titulo',''))[:60]}")
        alt = find_best_match_local(row.to_dict(), df_cec)
        results.append({"ref": row.to_dict(), "alt": alt})
    prog.empty()
    return results

# ── API key (solo para modo manual con ASIN) ─────────────────────────────────
def get_gemini_key():
    if st.session_state.get("GOOGLE_API_KEY"): return st.session_state["GOOGLE_API_KEY"]
    try:
        k = st.secrets.get("GOOGLE_API_KEY", None)
        if k: return k
    except Exception: pass
    return os.environ.get("GOOGLE_API_KEY", None)

# ── Load data ─────────────────────────────────────────────────────────────────
UPLOAD_DIR = Path(__file__).parent
_feed_path = st.session_state.get("feed_path")
if _feed_path and Path(_feed_path).exists():
    df_cecotec = load_cecotec_feed(_feed_path)
else:
    df_cecotec = load_cecotec_feed(str(UPLOAD_DIR))

df_keepa = load_keepa_files(str(UPLOAD_DIR))

# Status bar
feed_ok  = not df_cecotec.empty
keepa_ok = not df_keepa.empty
c1, c2, c3 = st.columns(3)
with c1:
    if feed_ok: st.success(f"✅ Feed Cecotec: **{len(df_cecotec):,}** productos en stock")
    else:       st.error("❌ Feed Cecotec no cargado")
with c2:
    if keepa_ok:
        df_rel = df_keepa[df_keepa["_cecotec_relevant"]]
        st.success(f"✅ Keepa: **{len(df_rel)}** productos relevantes (de {len(df_keepa)} totales)")
    else:
        st.warning("⚠️ Archivos Keepa no encontrados")
with c3:
    st.info("⚡ **Matching local** · Sin IA · Sin scraping · Instantáneo")

# Feed uploader si no está cargado
if not feed_ok:
    st.markdown('<div class="cec-section-title">📂 Subir catálogo Cecotec</div>', unsafe_allow_html=True)
    st.info("Sube `feed_Espan_a.xlsx` para iniciar la app.")
    feed_upload = st.file_uploader("feed_Espan_a.xlsx", type=["xlsx","xls"], key="feed_upload")
    if feed_upload:
        import tempfile
        tmp = Path(tempfile.mkdtemp()) / "feed_Espan_a.xlsx"
        tmp.write_bytes(feed_upload.read())
        df_cecotec = load_cecotec_feed(str(tmp))
        st.session_state["feed_path"] = str(tmp)
        st.success(f"✅ {len(df_cecotec):,} productos cargados")
        st.rerun()
    else:
        st.stop()

# ── TABS ──────────────────────────────────────────────────────────────────────
tab_keepa, tab_manual, tab_fichero, tab_resultados, tab_feed = st.tabs([
    "📦 Keepa Bestsellers","✏️ Producto manual",
    "📂 Subir fichero","📊 Resultados","🗄️ Feed Cecotec",
])

# ═══ TAB 1 · KEEPA ════════════════════════════════════════════════════════════
with tab_keepa:
    if not keepa_ok:
        st.warning("Archivos Keepa no encontrados.")
    else:
        df_relevant = df_keepa[df_keepa["_cecotec_relevant"]].copy()
        df_skipped  = df_keepa[~df_keepa["_cecotec_relevant"]].copy()
        st.markdown('<div class="cec-section-title">📊 Bestsellers Amazon · Análisis de competencia</div>', unsafe_allow_html=True)
        st.markdown(f"""<div class="kpi-row">
          <div class="kpi"><div class="val">{len(df_keepa[df_keepa['_source']=='Hogar']):,}</div><div class="lbl">Productos Hogar</div></div>
          <div class="kpi"><div class="val">{len(df_keepa[df_keepa['_source']=='Belleza']):,}</div><div class="lbl">Productos Belleza</div></div>
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
            procesar_todos = st.checkbox("Procesar TODOS los relevantes", value=True)
            if not procesar_todos:
                n = st.number_input("Máx. productos", 1, len(df_relevant), min(50, len(df_relevant)), key="k_n")

        df_proc = df_relevant[
            df_relevant["_source"].isin(src) &
            df_relevant["precio"].between(pmin_v, pmax_v)
        ]
        if not procesar_todos:
            df_proc = df_proc.head(int(n))

        st.info(f"Se procesarán **{len(df_proc)}** productos · matching local instantáneo ⚡")
        if st.button("🚀 Comparar todos con Cecotec", type="primary", use_container_width=True, key="btn_k"):
            with st.spinner("Calculando…"):
                st.session_state["results"] = run_search_local(df_proc, df_cecotec)
            st.success(f"✅ {len(df_proc)} productos comparados. Ve a **📊 Resultados**.")
            st.rerun()

# ═══ TAB 2 · MANUAL ═══════════════════════════════════════════════════════════
with tab_manual:
    st.markdown('<div class="cec-section-title">✏️ Buscar alternativa para un producto concreto</div>', unsafe_allow_html=True)
    st.caption("Introduce el ASIN — los datos se obtienen de Amazon automáticamente.")

    col_asin, col_price = st.columns([3,1])
    with col_asin:
        m_asin = st.text_input("ASIN Amazon *", placeholder="ej: B0BY9592V9")
    with col_price:
        m_precio_override = st.number_input("Precio (€) opcional", min_value=0.0, step=0.01, value=0.0)

    with st.expander("➕ Datos adicionales (opcional)"):
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
                    st.warning(f"Amazon no accesible: {e}")

            if amz_data.get("titulo"):
                st.success(f"✅ **{amz_data['titulo'][:60]}** · {amz_data.get('fabricante','')} · {amz_data.get('precio','')}€")
            if not amz_data["titulo"] and not amz_data["subcategoria"]:
                st.error("Sin datos. Rellena nombre y categoría manualmente.")
            else:
                with st.spinner("Buscando en feed Cecotec…"):
                    alt = find_best_match_local(amz_data, df_cecotec)
                    result = [{"ref": amz_data, "alt": alt}]
                    st.session_state["results_manual"] = result
                render_results(result)

    elif "results_manual" in st.session_state:
        render_results(st.session_state["results_manual"])

# ═══ TAB 3 · FICHERO ══════════════════════════════════════════════════════════
with tab_fichero:
    st.markdown('<div class="cec-section-title">📂 Subir fichero de productos competidores</div>', unsafe_allow_html=True)
    st.markdown("CSV o Excel con columnas: `titulo`, `fabricante`, `precio`, `subcategoria`, `caracteristicas` (o formato Keepa).")
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
            with cb: mx = st.number_input("Máx. productos", 1, len(df_custom), len(df_custom), key="mx_c")
            df_p = (df_rel_c if solo else df_custom).head(int(mx))
            st.info(f"Se procesarán **{len(df_p)}** productos · matching instantáneo ⚡")
            if st.button("🚀 Buscar alternativas", type="primary", use_container_width=True, key="btn_c"):
                with st.spinner("Calculando…"):
                    st.session_state["results_custom"] = run_search_local(df_p, df_cecotec)
                st.success("✅ Completado. Ve a **📊 Resultados**.")
                st.rerun()
        except Exception as e:
            st.error(f"Error: {e}")

# ═══ TAB 4 · RESULTADOS ═══════════════════════════════════════════════════════
with tab_resultados:
    res_key = None
    if st.session_state.get("results"):        res_key = "results"
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

# ═══ TAB 5 · FEED ═════════════════════════════════════════════════════════════
with tab_feed:
    st.markdown('<div class="cec-section-title">🗄️ Catálogo Cecotec cargado</div>', unsafe_allow_html=True)
    st.markdown(f"""<div class="kpi-row">
      <div class="kpi"><div class="val">{len(df_cecotec):,}</div><div class="lbl">Productos en stock</div></div>
      <div class="kpi"><div class="val">{df_cecotec['categories'].nunique()}</div><div class="lbl">Categorías</div></div>
      <div class="kpi"><div class="val">{df_cecotec['precio_final'].min():.0f}–{df_cecotec['precio_final'].max():.0f} €</div><div class="lbl">Rango de precios</div></div>
      <div class="kpi"><div class="val">{(df_cecotec['sale_price'] != df_cecotec['price']).sum()}</div><div class="lbl">En oferta</div></div>
    </div>""", unsafe_allow_html=True)
    cat_f = st.multiselect("Filtrar categoría", sorted(df_cecotec["categories"].unique()), key="feed_cat")
    srch  = st.text_input("Buscar en título/descripción", key="feed_srch")
    df_show = df_cecotec.copy()
    if cat_f:  df_show = df_show[df_show["categories"].isin(cat_f)]
    if srch:   df_show = df_show[df_show["title"].str.contains(srch, case=False, na=False) | df_show["desc_clean"].str.contains(srch, case=False, na=False)]
    st.caption(f"{len(df_show):,} productos")
    st.dataframe(df_show[["title","categories","precio_final","price","link","desc_clean"]].rename(
        columns={"title":"Producto","categories":"Categoría","precio_final":"Precio (€)",
                 "price":"Precio orig.","link":"URL","desc_clean":"Descripción"}),
        use_container_width=True, hide_index=True,
        column_config={
            "Precio (€)":   st.column_config.NumberColumn(format="%.2f €"),
            "Precio orig.": st.column_config.NumberColumn(format="%.2f €"),
            "URL":          st.column_config.LinkColumn("URL"),
        })
