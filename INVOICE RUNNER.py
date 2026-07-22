import streamlit as st
import pandas as pd
from datetime import datetime
import os

# ─────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────
TASA_CAMBIO = 36.6243
DB_FILE     = "DB.xlsx"
LOG_FILE    = "revisiones.xlsx"

DESTINOS = [
    "Managua","Tipitapa","Masaya","Matagalpa","Ticuantepe",
    "Chichigalpa","León","Estelí","Rivas","El Rama","Chontales"
]
DEST_COLS = {
    "Managua":"managua","Tipitapa":"tipitapa","Masaya":"masaya",
    "Matagalpa":"matagalpa","Ticuantepe":"ticuantepe","Chichigalpa":"chichigalpa",
    "León":"leon","Estelí":"esteli","Rivas":"rivas","El Rama":"el_rama","Chontales":"chontales"
}
CLIENTES_ESP = ["Ninguno","AMBER","IMPACTO","NLA","INDUCARIBE"]
CLIENTE_COLS = {"AMBER":"amber","IMPACTO":"impacto","NLA":"nla","INDUCARIBE":"inducaribe"}

# ─────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────
@st.cache_data
def cargar_tarifas():
    df = pd.read_excel(DB_FILE, sheet_name="Tarifas", header=1)
    rename = {
        "TRANSPORTE":            "transporte",
        "ESTADÍAS":              "estadias",
        "MOV. ADICIONAL":        "mov_adicional",
        "TRIPLE EJE":            "triple_eje",
        "AGILIZACIÓN\n≤26 ton":  "agil_hasta26",
        "AGILIZACIÓN\n>26 ton":  "agil_desde26",
        "SCANNER":               "scanner",
        "CL JULIA HERRERA":      "julia_herrera",
        "MANAGUA":               "managua",
        "TIPITAPA":              "tipitapa",
        "MASAYA":                "masaya",
        "MATAGALPA":             "matagalpa",
        "TICUANTEPE":            "ticuantepe",
        "CHICHIGALPA":           "chichigalpa",
        "LEÓN":                  "leon",
        "ESTELÍ":                "esteli",
        "RIVAS":                 "rivas",
        "EL RAMA":               "el_rama",
        "CHONTALES":             "chontales",
        "AMBER":                 "amber",
        "IMPACTO":               "impacto",
        "NLA":                   "nla",
        "INDUCARIBE":            "inducaribe",
        "OBSERVACIONES":         "observaciones",
    }
    df = df.rename(columns=rename)
    num_cols = ["estadias","mov_adicional","triple_eje","agil_hasta26","agil_desde26",
                "scanner","julia_herrera",
                "managua","tipitapa","masaya","matagalpa","ticuantepe","chichigalpa",
                "leon","esteli","rivas","el_rama","chontales",
                "amber","impacto","nla","inducaribe"]
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df["transporte"] = df["transporte"].astype(str).str.strip()
    return df

def guardar_revision(data: dict):
    if os.path.exists(LOG_FILE):
        df_log = pd.read_excel(LOG_FILE)
    else:
        df_log = pd.DataFrame()
    df_log = pd.concat([df_log, pd.DataFrame([data])], ignore_index=True)
    df_log.to_excel(LOG_FILE, index=False)

# ─────────────────────────────────────────
# ESTILO
# ─────────────────────────────────────────
st.set_page_config(page_title="Revisor de Facturas", page_icon="🧾", layout="centered")
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.block-container { max-width: 820px; padding-top: 2rem; }
.stSelectbox label, .stNumberInput label, .stTextInput label, .stDateInput label, .stRadio label {
    font-weight: 600; color: #ffffff; font-size: 0.8rem;
    letter-spacing: 0.05em; text-transform: uppercase;
}
.section-title {
    font-size: 0.7rem; font-weight: 700; letter-spacing: 0.1em;
    text-transform: uppercase; color: #6b7bab;
    margin: 1.6rem 0 0.4rem 0; border-bottom: 1px solid #e0e4f5; padding-bottom: 4px;
}
.desglose {
    background: #f7f9ff; border: 1px solid #dde2f5;
    border-radius: 12px; padding: 1.2rem 1.5rem; margin: 1rem 0;
}
.d-row {
    display:flex; justify-content:space-between; align-items:center;
    padding: 0.3rem 0; border-bottom: 1px solid #eef0f8;
    font-size: 0.92rem; color: #333;
}
.d-row:last-child { border-bottom: none; }
.d-label { color:#555; }
.d-detail { font-size:0.78rem; color:#999; margin-left:4px; }
.d-amount { font-weight:600; color:#1a1a2e; }
.d-total {
    display:flex; justify-content:space-between;
    padding: 0.7rem 0 0 0; margin-top:0.3rem;
    border-top: 2px solid #c5cdf5;
    font-size: 1.1rem; font-weight:700; color:#1a1a2e;
}
.obs-box {
    background:#fffbe6; border:1px solid #f0d060; border-radius:8px;
    padding:0.7rem 1rem; font-size:0.82rem; color:#5a4a00; margin:0.5rem 0;
}
.ok    { background:#e8f9f0; border:1.5px solid #2ecc71; border-radius:12px;
         padding:1rem 1.4rem; color:#1a6b3c; font-weight:600; font-size:0.95rem; margin-top:1rem; }
.error { background:#fff0f0; border:1.5px solid #e74c3c; border-radius:12px;
         padding:1rem 1.4rem; color:#7b1a1a; font-weight:600; font-size:0.95rem; margin-top:1rem; }
.warn  { background:#fff8e8; border:1.5px solid #f39c12; border-radius:12px;
         padding:1rem 1.4rem; color:#7a4a00; font-weight:600; font-size:0.95rem; margin-top:1rem; }
/* Título celeste claro */
h1 { color: #7ec8e3 !important; font-weight: 700 !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────
st.title("🧾 Verificador de Facturas - Oceanica Internacional")
st.caption("by AGVM")

try:
    df = cargar_tarifas()
except Exception as e:
    st.error(f"No se pudo leer {DB_FILE}: {e}")
    st.stop()

transportes = sorted(df["transporte"].dropna().unique().tolist())

# ─────────────────────────────────────────
# FORMULARIO
# ─────────────────────────────────────────
st.markdown('<p class="section-title">📋 Datos de la factura</p>', unsafe_allow_html=True)

# Fila 1: transportista y fecha (proveedor eliminado — se toma del transportista seleccionado)
c1, c2 = st.columns(2)
with c1: transporte  = st.selectbox("Transportista", transportes)
with c2: fecha_recib = st.date_input("Fecha recibida", value=datetime.today())

cf1, cf2 = st.columns(2)
with cf1: num_factura = st.text_input("N° Factura *", placeholder="E.g. 11985")
with cf2: posicion    = st.text_input("Posición", placeholder="E.g. IM05-01032-2026")

c3, c4 = st.columns(2)
with c3: destino   = st.selectbox("Destino", DESTINOS)
with c4: selectivo = st.selectbox("Selectivo", ["Verde", "Rojo", "DUCA"])

c5, c6 = st.columns(2)
with c5: cliente  = st.selectbox("Cliente especial", CLIENTES_ESP)
with c6: n_cont   = st.number_input("# Contenedores", min_value=1, max_value=50, value=1, step=1)

moneda = st.radio("Moneda", ["USD", "Córdobas (C$)"], horizontal=True)

st.markdown('<p class="section-title">➕ Cargos adicionales (cantidad de unidades)</p>', unsafe_allow_html=True)

ca1, ca2, ca3 = st.columns(3)
with ca1: q_estadias = st.number_input("Estadías",       min_value=0, value=0, step=1)
with ca2: q_triple   = st.number_input("Triple Eje",     min_value=0, value=0, step=1)
with ca3: q_mov_adic = st.number_input("Mov. Adicional", min_value=0, value=0, step=1)

ca4, ca5 = st.columns(2)
with ca4: q_agil26   = st.number_input("Agilización ≤26 TON", min_value=0, value=0, step=1)
with ca5: q_agil_m26 = st.number_input("Agilización >26 TON", min_value=0, value=0, step=1)

st.markdown('<p class="section-title"> Verificación</p>', unsafe_allow_html=True)
monto_fact = st.number_input("Monto en factura (USD)", min_value=0.0, value=0.0,
                              step=0.01, format="%.2f")

# ─────────────────────────────────────────
# CÁLCULO
# ─────────────────────────────────────────
fila = df[df["transporte"] == transporte]
if fila.empty:
    st.warning("No se encontraron datos o tarifas aplicables para ese transportista.")
    st.stop()
fila = fila.iloc[0]

def v(campo):
    val = fila.get(campo, None)
    try:    return float(val) if pd.notna(val) else 0.0
    except: return 0.0

def fmt(val):
    if moneda == "Córdobas (C$)":
        return f"C$ {val * TASA_CAMBIO:,.2f}"
    return f"$ {val:,.2f}"

dest_col    = DEST_COLS[destino]
tarifa_dest = v(dest_col)
tarifa_cli  = v(CLIENTE_COLS[cliente]) if cliente != "Ninguno" else 0.0

if tarifa_cli > 0:
    base, base_label = tarifa_cli, f"Tarifa {cliente}"
elif tarifa_dest > 0:
    base, base_label = tarifa_dest, f"Flete a {destino}"
else:
    base, base_label = 0.0, f"Flete a {destino} (sin tarifa)"

julia   = v("julia_herrera") if selectivo in ["Verde","DUCA"] else 0.0
scanner = v("scanner")       if selectivo == "DUCA"           else 0.0

c_estadias = v("estadias")      * q_estadias * n_cont
c_triple   = v("triple_eje")    * q_triple
c_mov_adic = v("mov_adicional") * q_mov_adic
c_agil26   = v("agil_hasta26")  * q_agil26
c_agil_m26 = v("agil_desde26")  * q_agil_m26

sub_base   = base * n_cont
sub_select = (julia + scanner) * n_cont
sub_extras = c_estadias + c_triple + c_mov_adic + c_agil26 + c_agil_m26
total_esp  = sub_base + sub_select + sub_extras

# ─────────────────────────────────────────
# DESGLOSE VISUAL
# ─────────────────────────────────────────
lineas = [(base_label, f"× {n_cont} cont.", sub_base)]
if julia      > 0: lineas.append(("Mov. Julia Herrera",   f"{fmt(julia)} × {n_cont}",             julia * n_cont))
if scanner    > 0: lineas.append(("Scanner",              f"{fmt(scanner)} × {n_cont}",           scanner * n_cont))
if c_estadias > 0: lineas.append(("Estadías",             f"{fmt(v('estadias'))} × {q_estadias}", c_estadias))
if c_triple   > 0: lineas.append(("Triple Eje",           f"{fmt(v('triple_eje'))} × {q_triple}", c_triple))
if c_mov_adic > 0: lineas.append(("Mov. Adicional",       f"{fmt(v('mov_adicional'))} × {q_mov_adic}", c_mov_adic))
if c_agil26   > 0: lineas.append(("Agilización ≤26 TON",  f"{fmt(v('agil_hasta26'))} × {q_agil26}",   c_agil26))
if c_agil_m26 > 0: lineas.append(("Agilización >26 TON",  f"{fmt(v('agil_desde26'))} × {q_agil_m26}", c_agil_m26))

rows_html = ""
for concepto, detalle, monto in lineas:
    rows_html += (f'<div class="d-row">'
                  f'<span class="d-label">{concepto} <span class="d-detail">({detalle})</span></span>'
                  f'<span class="d-amount">{fmt(monto)}</span></div>')
rows_html += f'<div class="d-total"><span>TOTAL ESPERADO</span><span>{fmt(total_esp)}</span></div>'
st.markdown(f'<div class="desglose">{rows_html}</div>', unsafe_allow_html=True)

obs = str(fila.get("observaciones","") or "").strip()
if obs and obs.lower() != "nan":
    partes   = [p.strip() for p in obs.replace("|||","||").split("||") if p.strip()]
    obs_html = " &nbsp;|&nbsp; ".join(f"⚠️ {p}" for p in partes)
    st.markdown(f'<div class="obs-box">{obs_html}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────
# RESULTADO DE VERIFICACIÓN
# ─────────────────────────────────────────
if monto_fact > 0:
    dif = monto_fact - total_esp
    if abs(dif) <= 0.01:
        st.markdown(f'<div class="ok">✅ Factura correcta, 📝 revisar Posición, nombre del cliente y a quién se le factura. &nbsp;|&nbsp; Esperado: {fmt(total_esp)} &nbsp;|&nbsp; Factura: {fmt(monto_fact)}</div>', unsafe_allow_html=True)
    elif dif > 0:
        st.markdown(f'<div class="error">❌ Factura con cobro de MÁS &nbsp;|&nbsp; Diferencia: {fmt(dif)} &nbsp;|&nbsp; Esperado: {fmt(total_esp)} &nbsp;|&nbsp; Factura: {fmt(monto_fact)}</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="warn">⚠️ Factura con cobro de MENOS &nbsp;|&nbsp; Diferencia: {fmt(abs(dif))} &nbsp;|&nbsp; Esperado: {fmt(total_esp)} &nbsp;|&nbsp; Factura: {fmt(monto_fact)}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────
# GUARDAR REVISIÓN
# ─────────────────────────────────────────
st.markdown("")
if st.button("💾 Guardar revisión", type="primary", use_container_width=True):
    if not num_factura.strip():
        st.warning("⚠️ El N° de Factura es obligatorio para guardar.")
        st.stop()
    guardar_revision({
        "N° Factura":          num_factura.strip(),
        "Posición":            posicion.strip(),
        "Fecha Recibida":      str(fecha_recib),
        "Fecha Revisada":      datetime.today().strftime("%Y-%m-%d %H:%M"),
        "Transportista":       transporte,
        "Destino":             destino,
        "Selectivo":           selectivo,
        "Cliente Especial":    cliente,
        "# Contenedores":      n_cont,
        "Total Esperado USD":  round(total_esp, 2),
        "Total Factura USD":   round(monto_fact, 2),
        "Diferencia USD":      round(monto_fact - total_esp, 2),
        "Resultado":           "OK" if abs(monto_fact - total_esp) <= 0.01 else "INCONSISTENCIA",
        "Tasa de Cambio":      TASA_CAMBIO,
    })
    st.success("✅ Verificación guardada en revisiones.xlsx")

# ─────────────────────────────────────────
# HISTORIAL
# ─────────────────────────────────────────
if os.path.exists(LOG_FILE):
    st.markdown('<p class="section-title">📂 Historial de revisiones</p>', unsafe_allow_html=True)
    df_log = pd.read_excel(LOG_FILE)
    st.dataframe(df_log[::-1], use_container_width=True, hide_index=True)
