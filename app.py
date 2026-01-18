import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np
import random

# ==============================================================================
# 1. SETUP VISUAL Y MEMORIA
# ==============================================================================
st.set_page_config(page_title="Michelin Pilot Command Center V49", page_icon="🚛", layout="wide")

# Inicialización de Memoria (Para que no se borre al descargar)
if 'process_ready' not in st.session_state:
    st.session_state.process_ready = False
    st.session_state.m_fin = None
    st.session_state.d_fin = None
    st.session_state.log_tel = ""
    st.session_state.c_resp = ""

st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    .stButton>button { width: 100%; height: 3.5em; font-weight: bold; border-radius: 8px; }
    .auto-success { padding: 15px; background-color: #d4edda; color: #155724; border-left: 5px solid #28a745; margin-bottom: 15px;}
    .auto-info { padding: 15px; background-color: #e7f3ff; color: #004085; border-left: 5px solid #007bff; margin-bottom: 15px;}
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNCIONES DE LIMPIEZA
# ==============================================================================

def limpiar_id_v49(valor):
    if pd.isna(valor): return ""
    return re.sub(r'\D', '', str(valor).split('.')[0])

def tratar_telefono_v49(val):
    if pd.isna(val): return None
    s = str(val).strip()
    if not s or s in ["55", "+55"]: return None
    nums = re.sub(r'\D', '', s)
    if nums.startswith('55') and len(nums) >= 12: nums = nums[2:]
    if len(nums) >= 10: return f"+55{nums}"
    return None

def identificar_perfil(v):
    d = re.sub(r'\D', '', str(v))
    return "PEQUENO FROTISTA" if len(d) == 14 else "FRETEIRO"

# ==============================================================================
# 3. MOTOR DE NEGOCIO (VERTICALIZACIÓN Y ZIP)
# ==============================================================================

def generar_zip_v49(df_m, df_d, col_resp):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
        agentes = [a for a in df_m[col_resp].unique() if a not in ["SEM_MATCH", "SEM_DONO"]]
        for ag in agentes:
            safe = re.sub(r'[^a-zA-Z0-9]', '_', unicodedata.normalize('NFKD', str(ag)).encode('ASCII', 'ignore').decode('ASCII').upper())
            
            # Mailing
            m_ag = df_m[df_m[col_resp] == ag]
            for p, n in [("PEQUENO FROTISTA", "MANHA_Frotista"), ("FRETEIRO", "ALMOCO_Freteiro")]:
                sub = m_ag[m_ag['PERFIL_FINAL'] == p]
                if not sub.empty:
                    excel = io.BytesIO(); sub.to_excel(excel, index=False)
                    zf.writestr(f"{safe}_MAILING_{n}.xlsx", excel.getvalue())
            
            # Discador
            d_ag = df_d[df_d['RESPONSAVEL_FINAL'] == ag]
            for p, n in [("PEQUENO FROTISTA", "MANHA_Frotista"), ("FRETEIRO", "ALMOCO_Freteiro")]:
                sub_d = d_ag[d_ag['PERFIL_FINAL'] == p]
                if not sub_d.empty:
                    excel_d = io.BytesIO(); sub_d.to_excel(excel_d, index=False)
                    zf.writestr(f"{safe}_DISCADOR_{n}.xlsx", excel_d.getvalue())
    buf.seek(0)
    return buf

# ==============================================================================
# 4. INTERFAZ (FRONT-END)
# ==============================================================================

st.title("Command Center Michelin V49")

with st.sidebar:
    st.header("📂 Carga de Archivos")
    file = st.file_uploader("Subir Excel Mestre", type="xlsx")
    if st.session_state.process_ready:
        if st.button("🗑️ NUEVO PROCESO (RESET)"):
            st.session_state.process_ready = False
            st.rerun()

if file:
    if not st.session_state.process_ready:
        if st.button("🚀 PROCESAR AHORA"):
            xls = pd.ExcelFile(file)
            aba_d = next(s for s in xls.sheet_names if 'DISC' in s.upper())
            aba_m = next(s for s in xls.sheet_names if 'MAIL' in s.upper())
            
            df_m = pd.read_excel(file, aba_m)
            df_d = pd.read_excel(file, aba_d)

            # --- DETECCIÓN DE COLUMNAS ---
            c_id_m = next(c for c in df_m.columns if 'ID_CONTRATO' in c.upper())
            c_id_d = next(c for c in df_d.columns if 'ID_CONTRATO' in c.upper())
            c_resp_m = next(c for c in df_m.columns if 'RESP' in c.upper())
            c_doc_m = next(c for c in df_m.columns if 'CNPJ' in c.upper())
            c_tels_d = [c for c in df_d.columns if any(x in c.upper() for x in ['TEL', 'CEL', 'PRINCIPAL', 'COMERCIAL'])]

            # --- LÓGICA MAILING ---
            df_m['PERFIL_FINAL'] = df_m[c_doc_m].apply(identificar_perfil)
            
            # --- LÓGICA DISCADOR (Verticalización Quirúrgica) ---
            df_d['KEY'] = df_d[c_id_d].apply(limpar_id_v49)
            map_resp = dict(zip(df_m[c_id_m].apply(limpar_id_v49), df_m[c_resp_m]))
            map_perf = dict(zip(df_m[c_id_m].apply(limpar_id_v49), df_m['PERFIL_FINAL']))
            
            df_d['RESPONSAVEL_FINAL'] = df_d['KEY'].map(map_resp).fillna("SEM_MATCH")
            df_d['PERFIL_FINAL'] = df_d['KEY'].map(map_perf).fillna("FRETEIRO")
            
            # Melt y Limpieza de Teléfonos
            ids_fix = [c for c in df_d.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL_FINAL', 'PERFIL_FINAL'])]
            df_d_vert = df_d.melt(id_vars=ids_fix, value_vars=c_tels_d, value_name='Tel_Raw')
            df_d_vert['Telefone'] = df_d_vert['Tel_Raw'].apply(tratar_telefono_v49)
            
            # Audit Log
            basura = df_d_vert['Telefone'].isna().sum()
            df_d_vert = df_d_vert.dropna(subset=['Telefone']).drop_duplicates(subset=['KEY', 'Telefone'])
            
            # Guardar en Memoria
            st.session_state.log_tel = f"Audit: {basura} celdas basura eliminadas. Registros finales: {len(df_d_vert)}"
            st.session_state.m_fin = df_m
            st.session_state.d_fin = df_d_vert
            st.session_state.c_resp = c_resp_m
            st.session_state.process_ready = True
            st.rerun()

# --- MOSTRAR RESULTADOS (PERSISTENTES) ---
if st.session_state.process_ready:
    st.markdown('<div class="auto-success">✅ ¡Bases listas para descargar!</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📊 Distribución Mailing")
        st.dataframe(pd.crosstab(st.session_state.m_fin[st.session_state.c_resp], st.session_state.m_fin['PERFIL_FINAL'], margins=True))
    
    with col2:
        st.subheader("📞 Auditoría Discador")
        st.markdown(f'<div class="auto-info">{st.session_state.log_tel}</div>', unsafe_allow_html=True)
        
        # Botón de Descarga
        zip_final = generar_zip_v49(st.session_state.m_fin, st.session_state.d_fin, st.session_state.c_resp)
        st.download_button("📥 DESCARGAR KIT DIARIO (.ZIP)", zip_final, "Michelin_Kit_V49.zip", "application/zip", type="primary")
