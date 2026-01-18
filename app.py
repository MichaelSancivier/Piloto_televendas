import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np
import random

# ==============================================================================
# 1. SETUP DE MEMORIA (SESSION STATE)
# ==============================================================================
# Esto evita que la aplicación se reinicie y borre los datos al descargar el ZIP
if 'process_done' not in st.session_state:
    st.session_state.process_done = False
    st.session_state.m_fin = None
    st.session_state.d_fin = None
    st.session_state.audit_log = ""
    st.session_state.c_resp = ""

st.set_page_config(page_title="Michelin Pilot Command Center", page_icon="🚛", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    .stButton>button { width: 100%; height: 3.5em; font-weight: bold; border-radius: 8px; }
    .auto-success { padding: 15px; background-color: #d4edda; color: #155724; border-left: 5px solid #28a745; margin-bottom: 15px;}
    .auto-info { padding: 15px; background-color: #e7f3ff; color: #004085; border-left: 5px solid #007bff; margin-bottom: 15px;}
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNCIONES DE LIMPIEZA Y LÓGICA
# ==============================================================================

def limpiar_id_final(valor):
    if pd.isna(valor): return ""
    # Quita decimales .0 y deja solo números
    return re.sub(r'\D', '', str(valor).split('.')[0])

def tratar_telefono_final(val):
    if pd.isna(val): return None
    s = str(val).strip()
    if not s or s in ["55", "+55"]: return None
    # Mantener si ya viene formateado correctamente
    if s.startswith('+55') and len(re.sub(r'\D', '', s)) >= 12: return s
    # Limpiar y estandarizar
    nums = re.sub(r'\D', '', s)
    if nums.startswith('55') and len(nums) >= 12: nums = nums[2:]
    if len(nums) >= 10: return f"+55{nums}"
    return None

def obtener_perfil(v):
    d = re.sub(r'\D', '', str(v))
    return "PEQUENO FROTISTA" if len(d) == 14 else "FRETEIRO"

def peso_prioridad(v):
    s = str(v).upper()
    if "BACKLOG" in s: return 99
    num = re.findall(r'\d+', s)
    return int(num[0]) if num else 50

# ==============================================================================
# 3. GENERACIÓN DE ARCHIVOS (ZIP)
# ==============================================================================

def crear_paquete_zip(df_m, df_d, col_resp):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
        agentes = [a for a in df_m[col_resp].unique() if a not in ["SEM_MATCH", "SEM_DONO"]]
        
        for ag in agentes:
            # Nombre de carpeta/archivo seguro
            safe_name = re.sub(r'[^a-zA-Z0-9]', '_', unicodedata.normalize('NFKD', str(ag)).encode('ASCII', 'ignore').decode('ASCII').upper())
            
            # MAILING (Mañana y Almuerzo)
            m_ag = df_m[df_m[col_resp] == ag]
            for perfil, sufijo in [("PEQUENO FROTISTA", "MANHA_Frotista"), ("FRETEIRO", "ALMOCO_Freteiro")]:
                sub = m_ag[m_ag['PERFIL_FINAL'] == perfil]
                if not sub.empty:
                    out = io.BytesIO(); sub.to_excel(out, index=False)
                    zf.writestr(f"{safe_name}_MAILING_{sufijo}.xlsx", out.getvalue())
            
            # DISCADOR (Mañana y Almuerzo)
            d_ag = df_d[df_d['RESPONSAVEL_FINAL'] == ag]
            for perfil, sufijo in [("PEQUENO FROTISTA", "MANHA_Frotista"), ("FRETEIRO", "ALMOCO_Freteiro")]:
                sub_d = d_ag[d_ag['PERFIL_FINAL'] == perfil]
                if not sub_d.empty:
                    out_d = io.BytesIO(); sub_d.to_excel(out_d, index=False)
                    zf.writestr(f"{safe_name}_DISCADOR_{sufijo}.xlsx", out_d.getvalue())
    buf.seek(0)
    return buf

# ==============================================================================
# 4. INTERFAZ Y PROCESAMIENTO
# ==============================================================================

st.title("🚛 Michelin Pilot Command Center - V50")

with st.sidebar:
    st.header("📂 Entrada")
    file = st.file_uploader("Subir archivo Excel Mestre", type="xlsx")
    
    if st.session_state.process_done:
        st.markdown("---")
        if st.button("🗑️ NUEVO PROCESO"):
            st.session_state.process_done = False
            st.rerun()

if file:
    if not st.session_state.process_done:
        if st.button("🚀 PROCESAR Y SINCRONIZAR"):
            xls = pd.ExcelFile(file)
            # Detección de pestañas
            aba_d = next(s for s in xls.sheet_names if 'DISC' in s.upper())
            aba_m = next(s for s in xls.sheet_names if 'MAIL' in s.upper())
            
            df_m = pd.read_excel(file, aba_m)
            df_d = pd.read_excel(file, aba_d)

            # --- AUTO DETECCIÓN DE COLUMNAS ---
            c_id_m = next(c for c in df_m.columns if 'ID_CONTRATO' in c.upper())
            c_id_d = next(c for c in df_d.columns if 'ID_CONTRATO' in c.upper())
            c_resp_m = next(c for c in df_m.columns if 'RESP' in c.upper())
            c_doc_m = next(c for c in df_m.columns if 'CNPJ' in c.upper())
            c_prio_m = next((c for c in df_m.columns if 'PRIOR' in c.upper()), None)
            c_tels_d = [c for c in df_d.columns if any(x in c.upper() for x in ['TEL', 'CEL', 'PRINCIPAL', 'COMERCIAL'])]

            # --- 1. LÓGICA DE MAILING (MAESTRO) ---
            df_m['PERFIL_FINAL'] = df_m[c_doc_m].apply(obtener_perfil)
            df_m['PRIO_SCORE'] = df_m[c_prio_m].apply(peso_prioridad) if c_prio_m else 50
            
            # Balanceo Quirúrgico
            agentes = [a for a in df_m[c_resp_m].unique() if pd.notna(a) and "BACKLOG" not in str(a).upper()]
            # (Reparto simplificado: igualar cargas)
            mask_orfao = df_m[c_resp_m].isna() | (df_m[c_resp_m].astype(str).str.upper().str.contains("BACKLOG"))
            for idx in df_m[mask_orfao].index:
                counts = df_m[c_resp_m].value_counts()
                agente_menor = min(agentes, key=lambda x: counts.get(x, 0))
                df_m.at[idx, c_resp_m] = agente_menor

            # --- 2. LÓGICA DE DISCADOR (ESCLAVO) ---
            df_d['KEY'] = df_d[c_id_d].apply(limpiar_id_final)
            map_resp = dict(zip(df_m[c_id_m].apply(limpar_id_final), df_m[c_resp_m]))
            map_perf = dict(zip(df_m[c_id_m].apply(limpar_id_final), df_m['PERFIL_FINAL']))
            
            df_d['RESPONSAVEL_FINAL'] = df_d['KEY'].map(map_resp).fillna("SEM_MATCH")
            df_d['PERFIL_FINAL'] = df_d['KEY'].map(map_perf).fillna("FRETEIRO")
            
            # Verticalización
            ids_fixos = [c for c in df_d.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL_FINAL', 'PERFIL_FINAL'])]
            df_vert = df_d.melt(id_vars=ids_fixos, value_vars=c_tels_d, value_name='Tel_Raw')
            df_vert['Telefone'] = df_vert['Tel_Raw'].apply(tratar_telefono_final)
            
            # Auditoría
            basura = df_vert['Telefone'].isna().sum()
            df_vert = df_vert.dropna(subset=['Telefone']).drop_duplicates(subset=['KEY', 'Telefone'])
            
            # Guardar en memoria
            st.session_state.m_fin = df_m
            st.session_state.d_fin = df_vert
            st.session_state.audit_log = f"Audit: {basura} celdas vacías eliminadas. {len(df_vert)} números válidos listos."
            st.session_state.c_resp = c_resp_m
            st.session_state.process_done = True
            st.rerun()

# --- MOSTRAR RESULTADOS (SÓLO SI ESTÁ LISTO) ---
if st.session_state.process_done:
    st.markdown('<div class="auto-success">✅ ¡Bases procesadas y sincronizadas correctamente!</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📊 Resumen Mailing")
        st.dataframe(pd.crosstab(st.session_state.m_fin[st.session_state.c_resp], st.session_state.m_fin['PERFIL_FINAL'], margins=True))
    
    with col2:
        st.subheader("📞 Auditoría de Teléfonos")
        st.markdown(f'<div class="auto-info">{st.session_state.log_tel if "log_tel" in st.session_state else st.session_state.audit_log}</div>', unsafe_allow_html=True)
        
        # Generar ZIP para descarga
        zip_file = crear_paquete_zip(st.session_state.m_fin, st.session_state.d_fin, st.session_state.c_resp)
        st.download_button(
            label="📥 DESCARGAR KIT DIARIO (.ZIP)",
            data=zip_file,
            file_name="Michelin_Kit_Final.zip",
            mime="application/zip",
            type="primary"
        )
