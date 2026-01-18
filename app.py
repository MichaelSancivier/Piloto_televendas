import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np
import random

# ==============================================================================
# 1. ESTADO DE LA SESIÓN (LA MEMORIA DEL SISTEMA)
# ==============================================================================
if 'process_ready' not in st.session_state:
    st.session_state.update({
        'process_ready': False,
        'm_fin': None,
        'd_fin': None,
        'audit_cartera': None,
        'c_resp': "",
        'c_prio': "",
        'c_tels': [],
        'log_audit_tel': "",
        'refuerzo_data': None
    })

st.set_page_config(page_title="Michelin Pilot V56 - Full Operation", page_icon="🚛", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    .stButton>button { width: 100%; height: 3.8em; font-weight: bold; border-radius: 10px; background-color: #003366; color: white; border: 2px solid #FCE500; }
    .stButton>button:hover { background-color: #FCE500; color: #003366; }
    .auto-success { padding: 20px; background-color: #d4edda; color: #155724; border-left: 6px solid #28a745; border-radius: 5px; margin-bottom: 20px;}
    .auto-info { padding: 15px; background-color: #e7f3ff; color: #004085; border-left: 6px solid #007bff; border-radius: 5px; margin-bottom: 20px;}
    .auto-warning { padding: 15px; background-color: #fff3cd; color: #856404; border-left: 5px solid #ffeeba; border-radius: 5px; margin-bottom: 20px;}
    h3 { color: #003366; border-bottom: 3px solid #FCE500; padding-bottom: 8px; margin-top: 25px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNCIONES LÓGICAS CORE
# ==============================================================================

def limpiar_id_universal(v):
    if pd.isna(v): return ""
    return re.sub(r'\D', '', str(v).split('.')[0])

def tratar_tel_universal(v):
    if pd.isna(v): return None
    s = str(v).strip()
    if not s or s in ["55", "+55"]: return None
    if s.startswith('+55') and len(re.sub(r'\D', '', s)) >= 12: return s
    nums = re.sub(r'\D', '', s)
    if nums.startswith('55') and len(nums) >= 12: nums = nums[2:]
    return f"+55{nums}" if len(nums) >= 10 else None

def perfil_por_doc(v):
    d = re.sub(r'\D', '', str(v))
    return "PEQUENO FROTISTA" if len(d) == 14 else "FRETEIRO"

def prioridade_valor(v):
    s = str(v).upper()
    if "BACKLOG" in s: return 99
    n = re.findall(r'\d+', s)
    return int(n[0]) if n else 50

# ==============================================================================
# 3. MOTOR DE GENERACIÓN ZIP (LOS 8 ARCHIVOS + REFUERZO)
# ==============================================================================

def generar_kit_completo_v56(df_m, df_d, col_resp, col_tels_d, df_refuerzo=None):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
        agentes = [a for a in df_m[col_resp].unique() if a not in ["SEM_MATCH", "SEM_DONO"]]
        
        for ag in agentes:
            safe = re.sub(r'[^a-zA-Z0-9]', '_', unicodedata.normalize('NFKD', str(ag)).encode('ASCII', 'ignore').decode('ASCII').upper())
            
            # --- MAILING (2 archivos por agente) ---
            m_ag = df_m[df_m[col_resp] == ag]
            for perf, suf in [("PEQUENO FROTISTA", "MANHA_Frotista"), ("FRETEIRO", "ALMOCO_Freteiro")]:
                sub_m = m_ag[m_ag['PERFIL_FINAL'] == perf]
                ex_m = io.BytesIO(); sub_m.to_excel(ex_m, index=False)
                zf.writestr(f"{safe}_MAILING_{suf}.xlsx", ex_m.getvalue())
            
            # --- DISCADOR MAÑANA (2 archivos por agente) ---
            d_ag = df_d[df_d['RESPONSAVEL_FINAL'] == ag]
            for perf, suf in [("PEQUENO FROTISTA", "MANHA_Frotista"), ("FRETEIRO", "ALMOCO_Freteiro")]:
                sub_d = d_ag[d_ag['PERFIL_FINAL'] == perf]
                if not sub_d.empty:
                    ids_d = [c for c in sub_d.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL', 'PERFIL'])]
                    melted = sub_d.melt(id_vars=ids_d, value_vars=col_tels_d, value_name='TR')
                    melted['Telefone'] = melted['TR'].apply(tratar_tel_universal)
                    final_d = melted.dropna(subset=['Telefone']).drop_duplicates(subset=['Telefone'])
                    ex_d = io.BytesIO(); final_d.to_excel(ex_d, index=False)
                    zf.writestr(f"{safe}_DISCADOR_{suf}.xlsx", ex_d.getvalue())

            # --- REFUERZO TARDE (Si existe) ---
            if df_refuerzo is not None:
                r_ag = df_refuerzo[df_refuerzo['RESPONSAVEL_FINAL'] == ag]
                if not r_ag.empty:
                    ex_r = io.BytesIO(); r_ag.to_excel(ex_r, index=False)
                    zf.writestr(f"{safe}_DISCADOR_REFUERZO_TARDE.xlsx", ex_r.getvalue())
                    
    buf.seek(0)
    return buf

# ==============================================================================
# 4. INTERFAZ Y PROCESAMIENTO
# ==============================================================================

st.title("🚛 Command Center Michelin V56 - Operación 360°")

with st.sidebar:
    st.header("1️⃣ Mañana: Bases Maestras")
    file_mestre = st.file_uploader("Excel Mestre", type="xlsx")
    
    st.header("2️⃣ Tarde: Refuerzo")
    file_log = st.file_uploader("Log del Discador (CSV/XLSX)", type=["csv", "xlsx"])
    
    if st.session_state.process_ready:
        st.markdown("---")
        if st.button("🗑️ RESETEAR TODO"):
            st.session_state.process_ready = False
            st.rerun()

if file_mestre:
    if not st.session_state.process_ready:
        if st.button("🚀 INICIAR DÍA (MAÑANA)"):
            xls = pd.ExcelFile(file_mestre)
            aba_d = next(s for s in xls.sheet_names if 'DISC' in s.upper())
            aba_m = next(s for s in xls.sheet_names if 'MAIL' in s.upper())
            df_m, df_d = pd.read_excel(file_mestre, aba_m), pd.read_excel(file_mestre, aba_d)

            # --- DETECCIÓN ---
            c_id_m = next(c for c in df_m.columns if 'ID_CONTRATO' in c.upper())
            c_id_d = next(c for c in df_d.columns if 'ID_CONTRATO' in c.upper())
            c_resp_m = next(c for c in df_m.columns if 'RESP' in c.upper())
            c_doc_m = next(c for c in df_m.columns if 'CNPJ' in c.upper())
            c_prio_m = next(c for c in df_m.columns if 'PRIOR' in c.upper())
            c_tels_d = [c for c in df_d.columns if any(x in c.upper() for x in ['TEL', 'CEL', 'PRINCIPAL', 'COMERCIAL'])]

            # --- BALANCEO ---
            snap_ini = df_m[c_resp_m].value_counts().to_dict()
            df_m['PERFIL_FINAL'] = df_m[c_doc_m].apply(perfil_por_doc)
            df_m['PRIO_SCORE'] = df_m[c_prio_m].apply(prioridade_valor)
            
            agentes = [a for a in df_m[c_resp_m].unique() if pd.notna(a) and "BACKLOG" not in str(a).upper()]
            mask_orfao = df_m[c_resp_m].isna() | (df_m[c_resp_m].astype(str).str.upper().str.contains("BACKLOG"))
            for idx in df_m[mask_orfao].index:
                ag_menor = min(agentes, key=lambda x: df_m[c_resp_m].value_counts().get(x, 0))
                df_m.at[idx, c_resp_m] = ag_menor

            # Sincronía Discador
            df_d['KEY'] = df_d[c_id_d].apply(limpiar_id_universal)
            map_resp = dict(zip(df_m[c_id_m].apply(limpiar_id_universal), df_m[c_resp_m]))
            map_perf = dict(zip(df_m[c_id_m].apply(limpiar_id_universal), df_m['PERFIL_FINAL']))
            df_d['RESPONSAVEL_FINAL'] = df_d['KEY'].map(map_resp).fillna("SEM_MATCH")
            df_d['PERFIL_FINAL'] = df_d['KEY'].map(map_perf).fillna("FRETEIRO")

            # Auditoría Cartera
            snap_fin = df_m[c_resp_m].value_counts().to_dict()
            df_audit = pd.DataFrame([snap_ini, snap_fin], index=['Inicio', 'Final']).T.fillna(0)

            st.session_state.update({
                'm_fin': df_m, 'd_fin': df_d, 'audit_cartera': df_audit,
                'c_resp': c_resp_m, 'c_prio': c_prio_m, 'c_tels': c_tels_d,
                'process_ready': True
            })
            st.rerun()

# --- LÓGICA DE LA TARDE (DENTRO DE LA MEMORIA) ---
if st.session_state.process_ready and file_log:
    try:
        df_log = pd.read_csv(file_log, sep=None, engine='python') if file_log.name.endswith('.csv') else pd.read_excel(file_log)
        col_log_id = df_log.columns[0]
        ids_contactados = df_log[col_log_id].apply(limpiar_id_universal).unique()
        
        # Filtrar Discador para Tarde
        d_base = st.session_state.d_fin.copy()
        # Solo Frotistas y que NO estén en el log
        d_tarde = d_base[(d_base['PERFIL_FINAL'] == "PEQUENO FROTISTA") & (~d_base['KEY'].isin(ids_contactados))]
        
        # Verticalizar para Refuerzo
        ids_r = [c for c in d_tarde.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL', 'PERFIL'])]
        melt_r = d_tarde.melt(id_vars=ids_r, value_vars=st.session_state.c_tels, value_name='TR')
        melt_r['Telefone'] = melt_r['TR'].apply(tratar_tel_universal)
        st.session_state.refuerzo_data = melt_r.dropna(subset=['Telefone']).drop_duplicates(subset=['Telefone'])
        st.sidebar.success(f"✅ Log cargado: {len(st.session_state.refuerzo_data)} registros para la tarde.")
    except Exception as e:
        st.sidebar.error(f"Error en log: {e}")

# --- DESPLIEGUE DE DASHBOARD ---
if st.session_state.process_ready:
    st.markdown('<div class="auto-success">✅ Command Center Activo. Datos Sincronizados.</div>', unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["📊 Cartera y Perfiles", "🔢 Prioridades", "📥 Descargas"])
    
    with tab1:
        c1, c2 = st.columns(2)
        c1.write("**Equidad de Cartera (47 c/u)**"); c1.dataframe(st.session_state.audit_cartera.style.format("{:.0f}"), use_container_width=True)
        c2.write("**División Mañana/Almuerzo**"); c2.dataframe(pd.crosstab(st.session_state.m_fin[st.session_state.c_resp], st.session_state.m_fin['PERFIL_FINAL'], margins=True), use_container_width=True)
    
    with tab2:
        st.write("**Desglose de Prioridades**")
        st.dataframe(pd.crosstab(st.session_state.m_fin[st.session_state.c_resp], st.session_state.m_fin[st.session_state.c_prio], margins=True), use_container_width=True)
    
    with tab3:
        if st.session_state.refuerzo_data is not None:
            st.warning(f"⚠️ El KIT incluirá el REFUERZO DE LA TARDE ({len(st.session_state.refuerzo_data)} registros vivos).")
        
        zip_kit = generar_kit_completo_v56(st.session_state.m_fin, st.session_state.d_fin, st.session_state.c_resp, st.session_state.c_tels, st.session_state.refuerzo_data)
        st.download_button("📥 DESCARGAR KIT COMPLETO (8+ ARCHIVOS)", zip_kit, "Michelin_Full_Day_Kit.zip", "application/zip", type="primary")
