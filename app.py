import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np

# ==============================================================================
# 1. CONFIGURAÇÃO VISUAL
# ==============================================================================
st.set_page_config(page_title="Michelin Pilot V62 - Backlog 100", page_icon="🚛", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    .stButton>button { width: 100%; height: 3.8em; font-weight: bold; border-radius: 10px; background-color: #003366; color: white; border: 2px solid #FCE500; }
    .stButton>button:hover { background-color: #FCE500; color: #003366; }
    .auto-success { padding: 20px; background-color: #d4edda; color: #155724; border-left: 6px solid #28a745; border-radius: 5px; margin-bottom: 20px;}
    h3 { color: #003366; border-bottom: 3px solid #FCE500; padding-bottom: 8px; margin-top: 25px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNÇÕES DE LIMPEZA E PRIORIDADE (A REGRA DO 100)
# ==============================================================================

def super_limpador_tel(val):
    if pd.isna(val): return None
    nums = re.sub(r'\D', '', str(val).strip())
    while nums.startswith('0'): nums = nums[1:]
    while nums.startswith('55') and len(nums) > 11: nums = nums[2:]
    if len(nums) == 11: return f"+55{nums}"
    elif len(nums) == 10:
        if int(nums[2]) >= 6: return f"+55{nums[:2]}9{nums[2:]}"
        return f"+55{nums}"
    elif len(nums) > 11: return f"+55{nums[-11:]}"
    return None

def clean_id(v):
    if pd.isna(v): return ""
    return re.sub(r'\D', '', str(v).split('.')[0])

def get_perfil(v):
    d = re.sub(r'\D', '', str(v))
    return "PEQUENO FROTISTA" if len(d) == 14 else "FRETEIRO"

# --- A REGRA RESGATADA DO BACKLOG ---
def get_prio_score_v62(v):
    s = str(v).upper()
    if "BACKLOG" in s: return 100  # Peso máximo para ser movido primeiro
    nums = re.findall(r'\d+', s)
    return int(nums[0]) if nums else 50

# ==============================================================================
# 3. MOTOR DE GERAÇÃO ZIP
# ==============================================================================

def build_zip_v62(df_m, df_d, col_resp, col_tels_d, df_ref=None):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
        agentes = [a for a in df_m[col_resp].unique() if a not in ["SEM_MATCH", "SEM_DONO"]]
        for ag in agentes:
            safe = re.sub(r'[^a-zA-Z0-9]', '_', unicodedata.normalize('NFKD', str(ag)).encode('ASCII', 'ignore').decode('ASCII').upper())
            m_ag = df_m[df_m[col_resp] == ag]
            d_ag = df_d[df_d['RESPONSAVEL_FINAL'] == ag]
            
            for perf, suf in [("PEQUENO FROTISTA", "MANHA_Frotista"), ("FRETEIRO", "ALMOCO_Freteiro")]:
                # Mailing
                sub_m = m_ag[m_ag['PERFIL_FINAL'] == perf]
                ex_m = io.BytesIO(); sub_m.to_excel(ex_m, index=False)
                zf.writestr(f"{safe}_MAILING_{suf}.xlsx", ex_m.getvalue())

                # Discador
                sub_d = d_ag[d_ag['PERFIL_FINAL'] == perf]
                if not sub_d.empty:
                    ids_d = [c for c in sub_d.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL', 'PERFIL'])]
                    melted = sub_d.melt(id_vars=ids_d, value_vars=col_tels_d, value_name='TR')
                    melted['Telefone'] = melted['TR'].apply(super_limpador_tel)
                    final_d = melted.dropna(subset=['Telefone']).drop_duplicates(subset=['Telefone'])
                    ex_d = io.BytesIO(); final_d.to_excel(ex_d, index=False)
                    zf.writestr(f"{safe}_DISCADOR_{suf}.xlsx", ex_d.getvalue())

            if df_ref is not None:
                r_ag = df_ref[df_ref['RESPONSAVEL_FINAL'] == ag]
                if not r_ag.empty:
                    ex_r = io.BytesIO(); r_ag.to_excel(ex_r, index=False)
                    zf.writestr(f"{safe}_DISCADOR_REFORCO_TARDE.xlsx", ex_r.getvalue())
    buf.seek(0)
    return buf

# ==============================================================================
# 4. INTERFACE E LOGICA DE NIVELAMENTO
# ==============================================================================

st.title("🚛 Michelin Pilot V62 - Regra Backlog 100")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    fase = st.radio("Fase Operacional:", ["🌅 Manhã (Geração)", "☀️ Tarde (Reforço)"])
    file_master = st.file_uploader("Arquivo Mestre", type="xlsx")
    file_log = st.file_uploader("Log Discador (Tarde)", type=["csv", "xlsx"]) if fase == "☀️ Tarde (Reforço)" else None

if file_master:
    xls = pd.ExcelFile(file_master)
    df_m = pd.read_excel(file_master, next(s for s in xls.sheet_names if 'MAIL' in s.upper()))
    df_d = pd.read_excel(file_master, next(s for s in xls.sheet_names if 'DISC' in s.upper()))

    c_id_m = next(c for c in df_m.columns if 'ID_CONTRATO' in c.upper())
    c_resp_m = next(c for c in df_m.columns if 'RESP' in c.upper())
    c_prio_m = next(c for c in df_m.columns if 'PRIOR' in c.upper())
    c_doc_m = next(c for c in df_m.columns if 'CNPJ' in c.upper())
    c_tels_d = [c for c in df_d.columns if any(x in c.upper() for x in ['TEL', 'CEL', 'PRINCIPAL', 'COMERCIAL'])]

    snap_ini = df_m[c_resp_m].value_counts().to_dict()
    df_m['PERFIL_FINAL'] = df_m[c_doc_m].apply(get_perfil)
    df_m['PRIO_SCORE'] = df_m[c_prio_m].apply(get_prio_score_v62)

    # --- MOTOR DE BALANCEAMENTO V62 (CIRURGIA POR PRIORIDADE) ---
    agentes = [a for a in df_m[c_resp_m].unique() if pd.notna(a) and "BACKLOG" not in str(a).upper()]
    mask_orfao = df_m[c_resp_m].isna() | (df_m[c_resp_m].astype(str).str.upper().str.contains("BACKLOG"))
    for idx in df_m[mask_orfao].index:
        ag_menor = min(agentes, key=lambda x: df_m[c_resp_m].value_counts().get(x, 0))
        df_m.at[idx, c_resp_m] = ag_menor

    # Nivelamento forçado: Tira de quem tem mais e dá para quem tem menos
    for _ in range(100):
        cargas = df_m[c_resp_m].value_counts()
        ag_max, ag_min = max(agentes, key=lambda x: cargas.get(x,0)), min(agentes, key=lambda x: cargas.get(x,0))
        if (cargas[ag_max] - cargas[ag_min]) <= 1: break
        # Prioriza mover quem tem PRIO_SCORE mais alto (Backlog=100)
        idx_move = df_m[df_m[c_resp_m] == ag_max].sort_values('PRIO_SCORE', ascending=False).index[0]
        df_m.at[idx_move, c_resp_m] = ag_min

    # Sincronia Discador
    df_d['KEY'] = df_d[next(c for c in df_d.columns if 'ID_CONTRATO' in c.upper())].apply(clean_id)
    map_resp = dict(zip(df_m[c_id_m].apply(clean_id), df_m[c_resp_m]))
    map_perf = dict(zip(df_m[c_id_m].apply(clean_id), df_m['PERFIL_FINAL']))
    df_d['RESPONSAVEL_FINAL'] = df_d['KEY'].map(map_resp).fillna("SEM_MATCH")
    df_d['PERFIL_FINAL'] = df_d['KEY'].map(map_perf).fillna("FRETEIRO")

    # Lógica de Reforço
    df_ref = None
    if fase == "☀️ Tarde (Reforço)" and file_log:
        try:
            df_log = pd.read_csv(file_log, sep=None, engine='python') if file_log.name.endswith('.csv') else pd.read_excel(file_log)
            ids_cont = df_log[df_log.columns[0]].apply(clean_id).unique()
            d_tarde = df_d[(df_d['PERFIL_FINAL'] == "PEQUENO FROTISTA") & (~df_d['KEY'].isin(ids_cont))]
            ids_r = [c for c in d_tarde.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL', 'PERFIL'])]
            melt_r = d_tarde.melt(id_vars=ids_r, value_vars=c_tels_d, value_name='TR')
            melt_r['Telefone'] = melt_r['TR'].apply(super_limpador_tel)
            df_ref = melt_r.dropna(subset=['Telefone']).drop_duplicates(subset=['Telefone'])
            st.success(f"✅ Reforço Tarde: {len(df_ref)} registros.")
        except Exception as e: st.error(f"Erro Log: {e}")

    # --- DASHBOARD ---
    st.subheader("📊 Auditoria de Cartera V62")
    c1, c2 = st.columns(2)
    with c1:
        df_audit = pd.DataFrame([snap_ini, df_m[c_resp_m].value_counts().to_dict()], index=['Início', 'Final']).T.fillna(0)
        st.write("**Balanceamento Cirúrgico:**"); st.dataframe(df_audit.style.format("{:.0f}"), use_container_width=True)
    with c2:
        st.write("**Perfis (Ma manhã/Almoço):**"); st.dataframe(pd.crosstab(df_m[c_resp_m], df_m['PERFIL_FINAL'], margins=True), use_container_width=True)
    
    st.subheader("🔢 Detalhe Prioridades")
    st.dataframe(pd.crosstab(df_m[c_resp_m], df_m[c_prio_m], margins=True), use_container_width=True)

    zip_kit = build_zip_v62(df_m, df_d, c_resp_m, c_tels_d, df_ref)
    st.download_button("📥 DESCARREGAR KIT COMPLETO (V62)", zip_kit, "Michelin_Kit_V62.zip", "application/zip", type="primary")
