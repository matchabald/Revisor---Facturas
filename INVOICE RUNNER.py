import streamlit as st
import pandas as pd
from datetime import datetime
import os

#Configuración de valores
TASA_CAMBIO = 36.6243          
DB_FILE     = "DB.xlsx"      
LOG_FILE    = "revisiones.xlsx"  

#Carga de datos en DB
@st.cache_data
def cargar_tarifas():
    df = pd.read_excel(DB_FILE, sheet_name="DB 1")
    df.columns = df.columns.str.strip()
    df = df.rename(columns={"P. Transporte ": "Transporte", "P. Transporte": "Transporte"})
    df["Transporte"] = df["Transporte"].str.strip()
    return df

def guardar_revision(data: dict):
    if os.path.exists(LOG_FILE):
        df_log = pd.read_excel(LOG_FILE)
    else:
        df_log = pd.DataFrame()
    nueva_fila = pd.DataFrame([data])
    df_log = pd.concat([df_log, nueva_fila], ignore_index=True)
    df_log.to_excel(LOG_FILE, index=False)

#Interfaz
st.set_page_config(page_title="Revisor de Facturas", page_icon="🧾", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .block-container { max-width: 780px; padding-top: 2rem; }
    .stSelectbox label, .stNumberInput label, .stTextInput label {
        font-weight: 600; color: #1a1a2e; font-size: 0.85rem; letter-spacing: 0.04em; text-transform: uppercase;
    }
    .tariff-card {
        background: #f8f9ff;
        border: 1px solid #e0e4f5;
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        margin: 1rem 0;
    }
    .tariff-row {
        display: flex; justify-content: space-between;
        padding: 0.35rem 0;
        border-bottom: 1px solid #eef0f8;
        font-size: 0.95rem;
        color: #333;
    }
    .tariff-row:last-child { border-bottom: none; }
    .tariff-total {
        display: flex; justify-content: space-between;
        padding: 0.6rem 0 0 0;
        font-size: 1.1rem;
        font-weight: 700;
        color: #1a1a2e;
    }
    .result-ok {
        background: #e8f9f0; border: 1.5px solid #2ecc71;
        border-radius: 12px; padding: 1.2rem 1.5rem; margin-top: 1rem;
        color: #1a6b3c; font-weight: 600; font-size: 1rem;
    }
    .result-error {
        background: #fff0f0; border: 1.5px solid #e74c3c;
        border-radius: 12px; padding: 1.2rem 1.5rem; margin-top: 1rem;
        color: #7b1a1a; font-weight: 600; font-size: 1rem;
    }
    .section-title {
        font-size: 0.75rem; font-weight: 700; letter-spacing: 0.08em;
        text-transform: uppercase; color: #6b7bab; margin: 1.5rem 0 0.5rem 0;
    }
    h1 { color: #1a1a2e !important; font-weight: 700 !important; }
</style>
""", unsafe_allow_html=True)

# Header 
st.title("🧾 Revisor de Facturas")
st.caption("by Alexa Vargas.")

#Cargar datos
try:
    df = cargar_tarifas()
except Exception as e:
    st.error(f"No se pudo leer DB.xlsx: {e}")
    st.stop()

transportes = sorted(df["Transporte"].dropna().unique().tolist())
col_num = ["Scanner", "mov. Julia Herrera", "DUCAT", "flete verde", "flete rojo",
           "IMPACTO", "ANBER", "INDUCARIBE", "Estadías", "triple Eje",
           "Agilización - 26", "Agilización +26", "Mov. Adicional"]
for c in col_num:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

# Formulario
st.markdown('<p class="section-title">Datos de la factura</p>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    proveedor_input = st.text_input("Proveedor / Empresa emisora", placeholder="Ej. Transportes del Norte")
with col2:
    fecha_recibida = st.date_input("Fecha recibida", value=datetime.today())

col3, col4 = st.columns(2)
with col3:
    transporte = st.selectbox("Transportista", transportes)
with col4:
    selectivo = st.selectbox("Selectivo", ["verde", "rojo", "ducat"])

col5, col6 = st.columns(2)
with col5:
    cliente = st.selectbox("Cliente especial", ["Ninguno", "IMPACTO", "ANBER", "INDUCARIBE"])
with col6:
    n_contenedores = st.number_input("# Contenedores", min_value=1, max_value=50, value=1, step=1)

st.markdown('<p class="section-title">Cargos adicionales</p>', unsafe_allow_html=True)

col7, col8, col9 = st.columns(3)
with col7:
    estadias    = st.number_input("Estadías", min_value=0, value=0, step=1)
with col8:
    triple_eje  = st.number_input("Triple Eje", min_value=0, value=0, step=1)
with col9:
    mov_adicional = st.number_input("Mov. Adicional", min_value=0, value=0, step=1)

col10, col11 = st.columns(2)
with col10:
    agil_26     = st.number_input("Agilización ≤26 TON", min_value=0, value=0, step=1)
with col11:
    agil_mas26  = st.number_input("Agilización >26 TON", min_value=0, value=0, step=1)

st.markdown('<p class="section-title">Verificación</p>', unsafe_allow_html=True)

col12, col13 = st.columns(2)
with col12:
    monto_factura = st.number_input("Monto en factura (USD)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
with col13:
    moneda = st.radio("Ver totales en", ["USD", "Córdobas"], horizontal=True)


# CALCULO EN VIVO - preliminar
fila = df[df["Transporte"] == transporte]

if fila.empty:
    st.warning("No se encontraron tarifas para ese transportista.")
    st.stop()

fila = fila.iloc[0]

def conv(val):
    return val * TASA_CAMBIO if moneda == "Córdobas" else val

def fmt(val):
    simbolo = "C$" if moneda == "Córdobas" else "$"
    return f"{simbolo} {conv(val):,.2f}"

# Tarifa base según selectivo y cliente
if cliente == "IMPACTO" and fila["IMPACTO"] > 0:
    base = fila["IMPACTO"]
    base_label = f"Tarifa IMPACTO ({transporte})"
elif cliente == "ANBER" and fila["ANBER"] > 0:
    base = fila["ANBER"]
    base_label = f"Tarifa ANBER ({transporte})"
elif cliente == "INDUCARIBE" and fila["INDUCARIBE"] > 0:
    base = fila["INDUCARIBE"]
    base_label = f"Tarifa INDUCARIBE ({transporte})"
else:
    if selectivo == "ducat":
        base = fila["DUCAT"]
    elif selectivo == "verde":
        base = fila["flete verde"]
    else:
        base = fila["flete rojo"]
    base_label = f"Flete {selectivo} ({transporte})"

# Cargos por selectivo
julia   = fila["mov. Julia Herrera"] if selectivo in ["verde", "ducat"] else 0
scanner = fila["Scanner"]            if selectivo == "ducat" else 0

# Cargos adicionales (x unidades ingresadas)
cargo_estadias      = fila["Estadías"]         * estadias
cargo_triple        = fila["triple Eje"]        * triple_eje
cargo_mov_adic      = fila["Mov. Adicional"]    * mov_adicional
cargo_agil_26       = fila["Agilización - 26"]  * agil_26
cargo_agil_mas26    = fila["Agilización +26"]   * agil_mas26

subtotal_base   = base * n_contenedores
subtotal_extras = cargo_estadias + cargo_triple + cargo_mov_adic + cargo_agil_26 + cargo_agil_mas26
subtotal_select = (julia + scanner) * n_contenedores
total_esperado  = subtotal_base + subtotal_select + subtotal_extras

# Desglose visual
st.markdown('<div class="tariff-card">', unsafe_allow_html=True)

lineas = [
    (base_label,                   f"{fmt(base)} × {n_contenedores}",  fmt(subtotal_base)),
]
if julia > 0:
    lineas.append(("Mov. Julia Herrera",           f"{fmt(julia)} × {n_contenedores}",  fmt(julia * n_contenedores)))
if scanner > 0:
    lineas.append(("Scanner",                      f"{fmt(scanner)} × {n_contenedores}", fmt(scanner * n_contenedores)))
if cargo_estadias > 0:
    lineas.append(("Estadías",                     f"{fmt(fila['Estadías'])} × {estadias}", fmt(cargo_estadias)))
if cargo_triple > 0:
    lineas.append(("Triple Eje",                   f"{fmt(fila['triple Eje'])} × {triple_eje}", fmt(cargo_triple)))
if cargo_mov_adic > 0:
    lineas.append(("Mov. Adicional",               f"{fmt(fila['Mov. Adicional'])} × {mov_adicional}", fmt(cargo_mov_adic)))
if cargo_agil_26 > 0:
    lineas.append(("Agilización ≤26 TON",          f"{fmt(fila['Agilización - 26'])} × {agil_26}", fmt(cargo_agil_26)))
if cargo_agil_mas26 > 0:
    lineas.append(("Agilización >26 TON",          f"{fmt(fila['Agilización +26'])} × {agil_mas26}", fmt(cargo_agil_mas26)))

html_rows = ""
for concepto, detalle, monto in lineas:
    html_rows += f'<div class="tariff-row"><span>{concepto} <small style="color:#888">({detalle})</small></span><span>{monto}</span></div>'

html_rows += f'<div class="tariff-total"><span>TOTAL ESPERADO</span><span>{fmt(total_esperado)}</span></div>'
st.markdown(html_rows + '</div>', unsafe_allow_html=True)

# Resultado de verificación
if monto_factura > 0:
    diferencia = monto_factura - total_esperado
    if abs(diferencia) <= 0.01:
        st.markdown(f'<div class="result-ok">✅ Factura correcta — Diferencia: {fmt(diferencia)}</div>', unsafe_allow_html=True)
    elif diferencia > 0:
        st.markdown(f'<div class="result-error">❌ Factura cobrada de MÁS — Diferencia: {fmt(diferencia)}</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="result-error">⚠️ Factura cobrada de MENOS — Diferencia: {fmt(diferencia)}</div>', unsafe_allow_html=True)

# GUARDAR EN EXCEL -Preeliminar
st.markdown("")
if st.button("💾 Guardar revisión", type="primary", use_container_width=True):
    if not proveedor_input.strip():
        st.warning("Ingresá el nombre del proveedor antes de guardar.")
    else:
        guardar_revision({
            "Fecha Recibida":    str(fecha_recibida),
            "Fecha Revisada":    datetime.today().strftime("%Y-%m-%d %H:%M"),
            "Proveedor":         proveedor_input.strip(),
            "Transportista":     transporte,
            "Selectivo":         selectivo,
            "Cliente":           cliente,
            "# Contenedores":    n_contenedores,
            "Total Esperado USD": round(total_esperado, 2),
            "Total Factura USD":  round(monto_factura, 2),
            "Diferencia USD":     round(monto_factura - total_esperado, 2),
            "Resultado":         "OK" if abs(monto_factura - total_esperado) <= 0.01 else "INCONSISTENCIA",
            "Tasa de Cambio":    TASA_CAMBIO,
        })
        st.success("✅ Revisión guardada en revisiones.xlsx")

# Historial - preeliminar tmb jajaj
if os.path.exists(LOG_FILE):
    st.markdown('<p class="section-title">Historial de revisiones</p>', unsafe_allow_html=True)
    df_log = pd.read_excel(LOG_FILE)
    st.dataframe(df_log[::-1], use_container_width=True, hide_index=True)