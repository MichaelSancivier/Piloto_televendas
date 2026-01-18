import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np

# ==============================================================================
# 1. CONFIGURAÇÃO VISUAL E ESTADO DA SESSÃO
# ==============================================================================
st.set_page_config(page_title="Michelin Pilot V64 - Final Production", page_icon="🚛", layout="wide")

if 'dados_prontos' not in st.session_state:
    st.session_state.dados_prontos = False

st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    .stButton>button { width: 100%; height: 3.8em; font-weight: bold; border-radius: 10px; background-color: #003366; color: white; border: 2px solid #FCE500; }
    .stButton>button:hover { background-color: #FCE500; color: #003366; }
    .btn-reset>div>button { background-color: #ff4b4b !important; color: white !important; border: none !important; }
    .auto-success { padding: 20px; background-color: #d4edda; color: #155724; border-left: 6px solid #28a745; border-radius: 5px; margin-bottom: 20px;}
    .auto-info { padding: 20px; background-color: #e7f3ff; color: #004085; border-left: 6px solid #007bff; border-radius: 5px; margin-bottom: 20px;}
    h3 { color: #003366; border-bottom: 3px solid #FCE500; padding-bottom: 8px; margin-top: 25px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNÇÕES DE INTELIGÊNCIA DE DADOS
# ==============================================================================

def super_limpador_tel(val):
    """Limpa e formata telefones removendo múltiplos prefixos 55 e zeros."""
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

def limpar_id(v):
    if pd.isna(v): return ""
    return re.sub(r'\D', '', str(v).split('.')[0])

def obter_perfil(v):
    d = re.sub(r'\D', '', str(v))
    return "PEQUENO FROTISTA" if len(d) == 14 else "FRETEIRO"

def converter_prio_v64(v):
    """Atribui peso 100 ao Backlog para priorizar no balanceamento."""
    s = str(v).upper()
    if "BACKLOG" in s: return 100 
    nums = re.findall(r'\d+', s)
    return int(nums[0]) if nums else 50

# ==============================================================================
# 3. MOTOR DE PROCESSAMENTO E GERAÇÃO
# ==============================================================================

def gerar_pacote_v64(df_m, df_d, col_resp, col_tels_d, df_ref=None):
    buf = io.BytesIO()
    total_tel_final = 0
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

                # Discador com Verticalização
                sub_d = d_ag[d_ag['PERFIL_FINAL'] == perf]
                if not sub_d.empty:
                    ids_d = [c for c in sub_d.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL', 'PERFIL'])]
                    melted = sub_d.melt(id_vars=ids_d, value_vars=col_tels_d, value_name='TR')
                    melted['Telefone'] = melted['TR'].apply(super_limpador_tel)
                    final_d = melted.dropna(subset=['Telefone']).drop_duplicates(subset=['Telefone'])
                    total_tel_final += len(final_d)
                    ex_d = io.BytesIO(); final_d.to_excel(ex_d, index=False)
                    zf.writestr(f"{safe}_DISCADOR_{suf}.xlsx", ex_d.getvalue())

            if df_ref is not None:
                r_ag = df_ref[df_ref['RESPONSAVEL_FINAL'] == ag]
                if not r_ag.empty:
                    ex_r = io.BytesIO(); r_ag.to_excel(ex_r, index=False)
                    zf.writestr(f"{safe}_DISCADOR_REFORCO_TARDE.xlsx", ex_r.getvalue())
    buf.seek(0)
    return buf, total_tel_final

# ==============================================================================
# 4. INTERFACE DO USUÁRIO
# ==============================================================================

st.title("🚛 Michelin Command Center - Dashboard de Produção")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    fase = st.radio("Momento da Operação:", ["🌅 Manhã (Geração)", "☀️ Tarde (Reforço)"])
    file_master = st.file_uploader("Subir Arquivo Mestre (Excel)", type="xlsx")
    file_log = st.file_uploader("Subir Log do Discador (Tarde)", type=["csv", "xlsx"]) if fase == "☀️ Tarde (Reforço)" else None

    if st.session_state.dados_prontos:
        st.markdown("---")
        st.markdown('<div class="btn-reset">', unsafe_allow_html=True)
        if st.button("LIMPAR RESULTADOS"):
            st.session_state.dados_prontos = False
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

if file_master:
    xls = pd.ExcelFile(file_master)
    df_m = pd.read_excel(file_master, next(s for s in xls.sheet_names if 'MAIL' in s.upper()))
    df_d = pd.read_excel(file_master, next(s for s in xls.sheet_names if 'DISC' in s.upper()))

    # Deteção Automática de Colunas
    c_id_m = next(c for c in df_m.columns if 'ID_CONTRATO' in c.upper())
    c_resp_m = next(c for c in df_m.columns if 'RESP' in c.upper())
    c_prio_m = next(c for c in df_m.columns if 'PRIOR' in c.upper())
    c_doc_m = next(c for c in df_m.columns if 'CNPJ' in c.upper())
    c_tels_d = [c for c in df_d.columns if any(x in c.upper() for x in ['TEL', 'CEL', 'PRINCIPAL', 'COMERCIAL'])]

    snap_ini = df_m[c_resp_m].value_counts().to_dict()
    df_m['PERFIL_FINAL'] = df_m[c_doc_m].apply(obter_perfil)
    df_m['PRIO_SCORE'] = df_m[c_prio_m].apply(converter_prio_v64)

    # Motor de Balanceamento Cirúrgico (Meta 47 por atendente)
    agentes = [a for a in df_m[c_resp_m].unique() if pd.notna(a) and "BACKLOG" not in str(a).upper()]
    mask_orfao = df_m[c_resp_m].isna() | (df_m[c_resp_m].astype(str).str.upper().str.contains("BACKLOG"))
    for idx in df_m[mask_orfao].index:
        ag_menor = min(agentes, key=lambda x: df_m[c_resp_m].value_counts().get(x, 0))
        df_m.at[idx, c_resp_m] = ag_menor

    for _ in range(100):
        cargas = df_m[c_resp_m].value_counts()
        ag_max, ag_min = max(agentes, key=lambda x: cargas.get(x,0)), min(agentes, key=lambda x: cargas.get(x,0))
        if (cargas[ag_max] - cargas[ag_min]) <= 1: break
        idx_move = df_m[df_m[c_resp_m] == ag_max].sort_values('PRIO_SCORE', ascending=False).index[0]
        df_m.at[idx_move, c_resp_m] = ag_min

    # Sincronização de Chaves
    df_d['KEY'] = df_d[next(c for c in df_d.columns if 'ID_CONTRATO' in c.upper())].apply(limpar_id)
    map_resp = dict(zip(df_m[c_id_m].apply(limpar_id), df_m[c_resp_m]))
    map_perf = dict(zip(df_m[c_id_m].apply(limpar_id), df_m['PERFIL_FINAL']))
    df_d['RESPONSAVEL_FINAL'] = df_d['KEY'].map(map_resp).fillna("SEM_MATCH")
    df_d['PERFIL_FINAL'] = df_d['KEY'].map(map_perf).fillna("FRETEIRO")

    df_ref = None
    if fase == "☀️ Tarde (Reforço)" and file_log:
        try:
            df_log = pd.read_csv(file_log, sep=None, engine='python') if file_log.name.endswith('.csv') else pd.read_excel(file_log)
            ids_cont = df_log[df_log.columns[0]].apply(limpar_id).unique()
            d_tarde = df_d[(df_d['PERFIL_FINAL'] == "PEQUENO FROTISTA") & (~df_d['KEY'].isin(ids_cont))]
            ids_r = [c for c in d_tarde.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL', 'PERFIL'])]
            melt_r = d_tarde.melt(id_vars=ids_r, value_vars=c_tels_d, value_name='TR')
            melt_r['Telefone'] = melt_r['TR'].apply(super_limpador_tel)
            df_ref = melt_r.dropna(subset=['Telefone']).drop_duplicates(subset=['Telefone'])
            st.sidebar.success(f"✅ Reforço Gerado: {len(df_ref)} registros.")
        except Exception as e: st.sidebar.error(f"Erro no Log: {e}")

    # Exibição de Resultados
    st.session_state.dados_prontos = True
    st.subheader("📊 Auditoria de Carteira e Telefones")
    zip_kit, total_telefones = gerar_pacote_v64(df_m, df_d, c_resp_m, c_tels_d, df_ref)
    
    st.markdown(f'<div class="auto-info">🔍 <b>Resumo Técnico:</b> A verticalização gerou <b>{total_telefones} números válidos</b> únicos para discagem.</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.write("**Balanceamento Cirúrgico:**")
        df_audit = pd.DataFrame([snap_ini, df_m[c_resp_m].value_counts().to_dict()], index=['Início', 'Final']).T.fillna(0)
        st.dataframe(df_audit.style.format("{:.0f}"), use_container_width=True)
    with c2:
        st.write("**Divisão de Perfis:**")
        st.dataframe(pd.crosstab(df_m[c_resp_m], df_m['PERFIL_FINAL'], margins=True), use_container_width=True)
    
    st.subheader("🔢 Detalhe de Prioridades")
    st.dataframe(pd.crosstab(df_m[c_resp_m], df_m[c_prio_m], margins=True), use_container_width=True)

    st.download_button("📥 BAIXAR KIT COMPLETO (V64)", zip_kit, "Michelin_Kit_V64_Producao.zip", "application/zip", type="primary")
