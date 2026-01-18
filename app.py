import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np
import random

# ==============================================================================
# 1. MEMORIA DE SESIÓN (SESSION STATE)
# ==============================================================================
if 'morning_done' not in st.session_state:
    st.session_state.update({
        'morning_done': False,
        'm_fin': None,
        'd_fin': None,
        'audit_cartera': None,
        'c_resp': "",
        'c_prio': "",
        'c_tels': [],
        'refuerzo_final': None
    })

st.set_page_config(page_title="Michelin Pilot V57 - Command Center", page_icon="🚛", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    .stButton>button { width: 100%; height: 3.8em; font-weight: bold; border-radius: 10px; background-color: #003366; color: white; border: 2px solid #FCE500; }
    .stButton>button:hover { background-color: #FCE500; color: #003366; }
    .auto-success { padding: 20px; background-color: #d4edda; color: #155724; border-left: 6px solid #28a745; border-radius: 5px; margin-bottom: 20px;}
    .auto-info { padding: 15px; background-color: #e7f3ff; color: #004085; border-left: 6px solid #007bff; border-radius: 5px; margin-bottom: 20px;}
    .auto-warning { padding: 15px; background-color: #fff3cd; color: #856404; border-left: 6px solid #ffeeba; border-radius: 5px; margin-bottom: 20px;}
    h3 { color: #003366; border-bottom: 3px solid #FCE500; padding-bottom: 8px; margin-top: 25px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNCIONES CORE (LIMPIEZA Y LOGICA)
# ==============================================================================

def clean_id(v):
    if pd.isna(v): return ""
    return re.sub(r'\D', '', str(v).split('.')[0])

def clean_tel(v):
    if pd.isna(v): return None
    s = str(v).strip()
    if not s or s in ["55", "+55"]: return None
    if s.startswith('+55') and len(re.sub(r'\D', '', s)) >= 12: return s
    nums = re.sub(r'\D', '', s)
    if nums.startswith('55') and len(nums) >= 12: nums = nums[2:]
    return f"+55{nums}" if len(nums) >= 10 else None

def get_perfil(v):
    d = re.sub(r'\D', '', str(v))
    return "PEQUENO FROTISTA" if len(d) == 14 else "FRETEIRO"

def get_prio_val(v):
    s = str(v).upper()
    if "BACKLOG" in s: return 99
    n = re.findall(r'\d+', s)
    return int(n[0]) if n else 50

# ==============================================================================
# 3. MOTOR DE EXPORTACIÓN (ZIP)
# ==============================================================================

def build_zip_v57(df_m, df_d, col_resp, col_tels_d, df_ref=None):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
        agentes = [a for a in df_m[col_resp].unique() if a not in ["SEM_MATCH", "SEM_DONO"]]
        for ag in agentes:
            safe = re.sub(r'[^a-zA-Z0-9]', '_', unicodedata.normalize('NFKD', str(ag)).encode('ASCII', 'ignore').decode('ASCII').upper())
            
            # --- MAILING ---
            m_ag = df_m[df_m[col_resp] == ag]
            for perf, suf in [("PEQUENO FROTISTA", "MANHA_Frotista"), ("FRETEIRO", "ALMOCO_Freteiro")]:
                sub_m = m_ag[m_ag['PERFIL_FINAL'] == perf]
                ex_m = io.BytesIO(); sub_m.to_excel(ex_m, index=False)
                zf.writestr(f"{safe}_MAILING_{suf}.xlsx", ex_m.getvalue())
            
            # --- DISCADOR MAÑANA ---
            d_ag = df_d[df_d['RESPONSAVEL_FINAL'] == ag]
            for perf, suf in [("PEQUENO FROTISTA", "MANHA_Frotista"), ("FRETEIRO", "ALMOCO_Freteiro")]:
                sub_d = d_ag[d_ag['PERFIL_FINAL'] == perf]
                if not sub_d.empty:
                    ids_d = [c for c in sub_d.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL', 'PERFIL'])]
                    melted = sub_d.melt(id_vars=ids_d, value_vars=col_tels_d, value_name='TR')
                    melted['Telefone'] = melted['TR'].apply(clean_tel)
                    final_d = melted.dropna(subset=['Telefone']).drop_duplicates(subset=['Telefone'])
                    ex_d = io.BytesIO(); final_d.to_excel(ex_d, index=False)
                    zf.writestr(f"{safe}_DISCADOR_{suf}.xlsx", ex_d.getvalue())

            # --- REFUERZO TARDE ---
            if df_ref is not None:
                r_ag = df_ref[df_ref['RESPONSAVEL_FINAL'] == ag]
                if not r_ag.empty:
                    ex_r = io.BytesIO(); r_ag.to_excel(ex_r, index=False)
                    zf.writestr(f"{safe}_DISCADOR_REFUERZO_TARDE.xlsx", ex_r.getvalue())
                    
    buf.seek(0)
    return buf

# ==============================================================================
# 4. INTERFAZ Y SELECTOR DE FASE
# ==============================================================================

st.title("🚛 Michelin Command Center V57")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    fase = st.radio("Fase da Operação:", ["🌅 Manhã (Geração)", "☀️ Tarde (Reforço)"])
    st.markdown("---")
    
    if fase == "🌅 Manhã (Geração)":
        file_master = st.file_uploader("Subir Excel Mestre", type="xlsx")
    else:
        file_log = st.file_uploader("Subir Log do Discador", type=["csv", "xlsx"])
        if not st.session_state.morning_done:
            st.warning("⚠️ Primeiro processe a Manhã.")

    if st.session_state.morning_done:
        if st.button("🗑️ RESETAR TUDO"):
            st.session_state.morning_done = False
            st.rerun()

# --- LÓGICA DE PROCESSAMENTO MANHÃ ---
if fase == "🌅 Manhã (Geração)" and file_master:
    if not st.session_state.morning_done:
        if st.button("🚀 PROCESSAR MAILING E DISCADOR"):
            xls = pd.ExcelFile(file_master)
            aba_d = next(s for s in xls.sheet_names if 'DISC' in s.upper())
            aba_m = next(s for s in xls.sheet_names if 'MAIL' in s.upper())
            df_m, df_d = pd.read_excel(file_master, aba_m), pd.read_excel(file_master, aba_d)

            c_id_m = next(c for c in df_m.columns if 'ID_CONTRATO' in c.upper())
            c_resp_m = next(c for c in df_m.columns if 'RESP' in c.upper())
            c_prio_m = next(c for c in df_m.columns if 'PRIOR' in c.upper())
            c_doc_m = next(c for c in df_m.columns if 'CNPJ' in c.upper())
            c_id_d = next(c for c in df_d.columns if 'ID_CONTRATO' in c.upper())
            c_tels_d = [c for c in df_d.columns if any(x in c.upper() for x in ['TEL', 'CEL', 'PRINCIPAL', 'COMERCIAL'])]

            snap_ini = df_m[c_resp_m].value_counts().to_dict()
            df_m['PERFIL_FINAL'] = df_m[c_doc_m].apply(get_perfil)
            df_m['PRIO_SCORE'] = df_m[c_prio_m].apply(get_prio_val)

            # Balanceo quirúrgico
            agentes = [a for a in df_m[c_resp_m].unique() if pd.notna(a) and "BACKLOG" not in str(a).upper()]
            mask_orfao = df_m[c_resp_m].isna() | (df_m[c_resp_m].astype(str).str.upper().str.contains("BACKLOG"))
            for idx in df_m[mask_orfao].index:
                ag_menor = min(agentes, key=lambda x: df_m[c_resp_m].value_counts().get(x, 0))
                df_m.at[idx, c_resp_m] = ag_menor

            # Sincronía
            df_d['KEY'] = df_d[c_id_d].apply(clean_id)
            map_resp = dict(zip(df_m[c_id_m].apply(clean_id), df_m[c_resp_m]))
            map_perf = dict(zip(df_m[c_id_m].apply(clean_id), df_m['PERFIL_FINAL']))
            df_d['RESPONSAVEL_FINAL'] = df_d['KEY'].map(map_resp).fillna("SEM_MATCH")
            df_d['PERFIL_FINAL'] = df_d['KEY'].map(map_perf).fillna("FRETEIRO")

            snap_fin = df_m[c_resp_m].value_counts().to_dict()
            df_audit = pd.DataFrame([snap_ini, snap_fin], index=['Início', 'Final']).T.fillna(0)

            st.session_state.update({
                'morning_done': True, 'm_fin': df_m, 'd_fin': df_d,
                'audit_cartera': df_audit, 'c_resp': c_resp_m, 'c_prio': c_prio_m, 'c_tels': c_tels_d
            })
            st.rerun()

# --- LÓGICA DE REFUERZO TARDE ---
if fase == "☀️ Tarde (Reforço)" and st.session_state.morning_done and 'file_log' in locals() and file_log:
    try:
        df_log = pd.read_csv(file_log, sep=None, engine='python') if file_log.name.endswith('.csv') else pd.read_excel(file_log)
        ids_contactados = df_log[df_log.columns[0]].apply(clean_id).unique()
        d_base = st.session_state.d_fin.copy()
        d_tarde = d_base[(d_base['PERFIL_FINAL'] == "PEQUENO FROTISTA") & (~d_base['KEY'].isin(ids_contactados))]
        
        # Verticalizar refuerzo
        ids_r = [c for c in d_tarde.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL', 'PERFIL'])]
        melt_r = d_tarde.melt(id_vars=ids_r, value_vars=st.session_state.c_tels, value_name='TR')
        melt_r['Telefone'] = melt_r['TR'].apply(clean_tel)
        st.session_state.refuerzo_final = melt_r.dropna(subset=['Telefone']).drop_duplicates(subset=['Telefone'])
    except Exception as e:
        st.error(f"Erro no Log: {e}")

# --- DASHBOARD VISUAL ---
if st.session_state.morning_done:
    st.markdown('<div class="auto-success">✅ Sistema Carregado. Use o seletor lateral para alternar entre Manhã e Tarde.</div>', unsafe_allow_html=True)
    
    if fase == "🌅 Manhã (Geração)":
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("⚖️ 1. Auditoria de Cartera")
            st.dataframe(st.session_state.audit_cartera.style.format("{:.0f}"), use_container_width=True)
        with c2:
            st.subheader("⏰ 2. Perfil Mañana/Almuerzo")
            st.dataframe(pd.crosstab(st.session_state.m_fin[st.session_state.c_resp], st.session_state.m_fin['PERFIL_FINAL'], margins=True), use_container_width=True)
        
        st.subheader("🔢 3. Detalhe por Prioridade")
        st.dataframe(pd.crosstab(st.session_state.m_fin[st.session_state.c_resp], st.session_state.m_fin[st.session_state.c_prio], margins=True), use_container_width=True)
    
    else:
        st.subheader("☀️ Reforço de Carteira (Tarde)")
        if st.session_state.refuerzo_final is not None:
            st.markdown(f'<div class="auto-info">🔍 <b>Reforço Ativo:</b> Foram encontrados {len(st.session_state.refuerzo_final)} números de Frotistas que não foram contactados hoje.</div>', unsafe_allow_html=True)
            st.dataframe(st.session_state.refuerzo_final.head(10), use_container_width=True)
        else:
            st.info("Suba o Log do Discador na lateral para gerar o reforço.")

    # DOWNLOAD UNIFICADO
    st.markdown("---")
    zip_kit = build_zip_v57(st.session_state.m_fin, st.session_state.d_fin, st.session_state.c_resp, st.session_state.c_tels, st.session_state.refuerzo_final)
    st.download_button("📥 DESCARREGAR KIT COMPLETO", zip_kit, "Michelin_Kit_V57.zip", "application/zip", type="primary")
