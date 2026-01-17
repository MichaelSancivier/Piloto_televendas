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
st.set_page_config(
    page_title="Michelin Pilot Command Center",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    h1 { color: #003366; font-family: 'Helvetica', sans-serif;}
    .stButton>button { 
        width: 100%; border-radius: 8px; height: 3.5em; font-weight: bold; 
        background-color: #003366; color: white; border: none;
    }
    .stButton>button:hover { background-color: #004080; color: white; }
    .stMetric { background-color: white; padding: 15px; border-radius: 8px; border-left: 6px solid #FCE500; box-shadow: 0 2px 5px rgba(0,0,0,0.05);}
    div[data-testid="stExpander"] { background-color: white; border-radius: 8px; }
    .audit-box-success { padding: 15px; background-color: #d4edda; color: #155724; border-radius: 8px; border: 1px solid #c3e6cb; margin-bottom: 20px; }
    .audit-box-danger { padding: 15px; background-color: #f8d7da; color: #721c24; border-radius: 8px; border: 1px solid #f5c6cb; margin-bottom: 20px; }
    .destaque-distribuicao { padding: 10px; background-color: #e8f4f8; border-radius: 5px; border-left: 5px solid #00aabb; margin-bottom: 20px;}
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. INTELIGÊNCIA LÓGICA (REGRA BINÁRIA)
# ==============================================================================

def normalizar_nome_arquivo(nome):
    if pd.isna(nome): return "SEM_CARTEIRA"
    nfkd = unicodedata.normalize('NFKD', str(nome))
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    limpo = re.sub(r'[^a-zA-Z0-9]', '_', sem_acento.upper())
    return re.sub(r'_+', '_', limpo).strip('_')

def identificar_perfil_pelo_doc(valor):
    """ 
    REGRA BINÁRIA (SIMPLIFICADA):
    - Tem 14 dígitos? -> PEQUENO FROTISTA (MANHÃ)
    - Qualquer outra coisa -> FRETEIRO (ALMOÇO)
    """
    if pd.isna(valor): return "FRETEIRO"
    
    doc_limpo = re.sub(r'\D', '', str(valor))
    
    if len(doc_limpo) == 14:
        return "PEQUENO FROTISTA"
    else:
        return "FRETEIRO"

def tratar_celular_discador(val):
    if pd.isna(val): return None, "Vazio"
    nums = re.sub(r'\D', '', str(val))
    if nums.startswith('55') and len(nums) >= 12: nums = nums[2:]
    
    tipo = "Inválido"
    final = None
    if len(nums) == 11 and int(nums[2]) == 9:
        tipo = "Celular"
        final = nums
    elif len(nums) == 10 and int(nums[2]) >= 6:
        tipo = "Celular (Corrigido)"
        final = nums[:2] + '9' + nums[2:]
    elif len(nums) == 10 and 2 <= int(nums[2]) <= 5:
        tipo = "Fixo"
        final = nums
        
    if final: return f"+55{final}", tipo
    return None, tipo

# ==============================================================================
# 3. MOTORES DE PROCESSAMENTO
# ==============================================================================

def distribuir_leads_orfãos(df, col_resp):
    df_proc = df.copy()
    termos_ignorar = ['CANAL TELEVENDAS', 'TIME', 'EQUIPE', 'TELEVENDAS', 'NULL', 'NAN', '']
    todos_resp = df_proc[col_resp].unique()
    agentes_humanos = [a for a in todos_resp if pd.notna(a) and str(a).strip().upper() not in termos_ignorar]
    
    if not agentes_humanos: return df_proc, "Sem agentes humanos.", 0

    mask_orfao = df_proc[col_resp].isna() | df_proc[col_resp].astype(str).str.strip().str.upper().isin(termos_ignorar)
    qtd_orfaos = mask_orfao.sum()
    
    if qtd_orfaos == 0: return df_proc, "Base completa.", 0

    atribuicoes = np.resize(agentes_humanos, qtd_orfaos)
    df_proc.loc[mask_orfao, col_resp] = atribuicoes
    return df_proc, f"Sucesso! {qtd_orfaos} leads redistribuídos.", qtd_orfaos

def processar_discador(df, col_id, cols_tel, col_resp, col_doc):
    df_trab = df.copy()
    # Aplica a regra Binária antes de tudo
    df_trab['ESTRATEGIA_PERFIL'] = df_trab[col_doc].apply(identificar_perfil_pelo_doc)
    
    cols_para_manter = col_id + [col_resp, 'ESTRATEGIA_PERFIL']
    df_melted = df_trab.melt(id_vars=cols_para_manter, value_vars=cols_tel, var_name='Origem', value_name='Telefone_Bruto')
    
    df_melted['Telefone_Tratado'], df_melted['Tipo'] = zip(*df_melted['Telefone_Bruto'].apply(tratar_celular_discador))
    df_final = df_melted.dropna(subset=['Telefone_Tratado'])
    df_final = df_final.drop_duplicates(subset=col_id + ['Telefone_Tratado'])
    return df_final

def processar_distribuicao_mailing(df, col_doc, col_resp):
    df_proc = df.copy()
    df_proc['PERFIL_CALCULADO'] = df_proc[col_doc].apply(identificar_perfil_pelo_doc)
    return df_proc

def gerar_zip_dinamico(df_dados, col_resp, col_segmentacao, modo="DISCADOR"):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        if col_resp not in df_dados.columns:
            df_dados['RESP_GERAL'] = 'EQUIPE'
            col_resp = 'RESP_GERAL'
            
        agentes = df_dados[col_resp].unique()

        for agente in agentes:
            nome_arquivo = normalizar_nome_arquivo(agente)
            if pd.isna(agente): df_agente = df_dados[df_dados[col_resp].isna()]
            else: df_agente = df_dados[df_dados[col_resp] == agente]
            if df_agente.empty: continue

            col_seg_uso = 'ESTRATEGIA_PERFIL' if modo == "DISCADOR" else col_segmentacao
            
            # --- SEPARAÇÃO ---
            df_frotista = df_agente[df_agente[col_seg_uso] == "PEQUENO FROTISTA"]
            df_freteiro = df_agente[df_agente[col_seg_uso] == "FRETEIRO"]
            
            prefixo = "DISCADOR" if modo == "DISCADOR" else "MAILING"
            pasta = nome_arquivo 
            
            # Salva FROTISTA (Manhã)
            if not df_frotista.empty:
                data = io.BytesIO()
                df_frotista.to_excel(data, index=False)
                zip_file.writestr(f"{pasta}/{prefixo}_1_MANHA_Frotista_CNPJ.xlsx", data.getvalue())
            
            # Salva FRETEIRO (Almoço) - Inclui tudo que não é Frotista
            if not df_freteiro.empty:
                data = io.BytesIO()
                df_freteiro.to_excel(data, index=False)
                zip_file.writestr(f"{pasta}/{prefixo}_2_ALMOCO_Freteiro_CPF.xlsx", data.getvalue())
            
            # Não existe mais a pasta "Outros/Backlog"

    zip_buffer.seek(0)
    return zip_buffer

# ==============================================================================
# 4. FRONTEND
# ==============================================================================

st.title("🚛 Michelin Pilot Command Center")
st.markdown("### Estratégia de Televentas & Logística")
st.markdown("---")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    st.header("📂 Entrada de Dados")
    uploaded_file = st.file_uploader("Arraste seu Excel aqui (.xlsx)", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    all_sheets = xls.sheet_names
    
    # --- AUDITORIA ---
    st.subheader("🕵️ Auditoria Inicial")
    aba_d_audit = next((s for s in all_sheets if 'DISCADOR' in s.upper()), None)
    aba_m_audit = next((s for s in all_sheets if 'MAILING' in s.upper()), None)
    
    if aba_d_audit and aba_m_audit:
        df_audit_d = pd.read_excel(uploaded_file, sheet_name=aba_d_audit)
        df_audit_m = pd.read_excel(uploaded_file, sheet_name=aba_m_audit)
        qtd_d = len(df_audit_d)
        qtd_m = len(df_audit_m)
        diff = abs(qtd_d - qtd_m)
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Reg. Discador", qtd_d)
        c2.metric("Reg. Mailing", qtd_m)
        if diff == 0:
            c3.metric("Status", "SINCRONIZADO", delta="OK")
            st.markdown(f'<div class="audit-box-success">✅ Massa consistente.</div>', unsafe_allow_html=True)
        else:
            c3.metric("Status", "DIVERGENTE", delta=f"-{diff}", delta_color="inverse")
            st.markdown(f'<div class="audit-box-danger">🚨 Atenção: Diferença de {diff} linhas.</div>', unsafe_allow_html=True)

    # --- BALANCEAMENTO ---
    st.markdown("### ⚖️ Distribuição de Carga")
    aplicar_balanceamento = st.checkbox("Distribuir leads sem dono entre a equipe?", value=True)
    
    tab1, tab2 = st.tabs(["🤖 DISCADOR", "👩‍💼 MAILING"])
    
    # ABA DISCADOR
    with tab1:
        col_sel, _ = st.columns([1,1])
        aba_d = col_sel.selectbox("Aba Discador:", all_sheets, index=all_sheets.index(aba_d_audit) if aba_d_audit else 0, key='d1')
        df_d = pd.read_excel(uploaded_file, sheet_name=aba_d)
        cols_d = df_d.columns.tolist()
        
        sug_tel = [c for c in cols_d if any(x in c.upper() for x in ['TEL','CEL','FONE'])]
        sug_id = [c for c in cols_d if any(x in c.upper() for x in ['ID','NOME'])]
        sug_doc = next((c for c in cols_d if 'CNPJ' in c.upper() or 'CPF' in c.upper()), None)
        sug_resp = next((c for c in cols_d if 'RESPONSAVEL' in c.upper()), cols_d[0])
        
        with st.expander("⚙️ Configurar Colunas"):
            c1, c2 = st.columns(2)
            sel_resp_d = c1.selectbox("Responsável:", cols_d, index=cols_d.index(sug_resp) if sug_resp in cols_d else 0, key='d_resp')
            sel_doc_d = c1.selectbox("CPF/CNPJ:", cols_d, index=cols_d.index(sug_doc) if sug_doc in cols_d else 0, key='d_doc')
            sel_id_d = c2.multiselect("ID Cliente:", cols_d, default=sug_id[:3], key='d_id')
            sel_tel_d = c2.multiselect("Telefones:", cols_d, default=sug_tel, key='d_tel')

        if st.button("🚀 PROCESSAR DISCADOR", key='btn_discador'):
            if not sel_tel_d: st.error("Selecione telefones.")
            else:
                with st.spinner("Processando..."):
                    df_trabalho = df_d.copy()
                    if aplicar_balanceamento:
                        df_trabalho, _, _ = distribuir_leads_orfãos(df_trabalho, sel_resp_d)
                    
                    df_res_d = processar_discador(df_trabalho, sel_id_d, sel_tel_d, sel_resp_d, sel_doc_d)
                    zip_d = gerar_zip_dinamico(df_res_d, sel_resp_d, None, "DISCADOR") 
                    st.success("Pronto!")
                    st.download_button("📥 BAIXAR DISCADOR", zip_d, "Discador.zip", "application/zip")

    # ABA MAILING
    with tab2:
        col_sel_m, _ = st.columns([1,1])
        aba_m = col_sel_m.selectbox("Aba Mailing:", all_sheets, index=all_sheets.index(aba_m_audit) if aba_m_audit else 0, key='m1')
        df_m = pd.read_excel(uploaded_file, sheet_name=aba_m)
        cols_m = df_m.columns.tolist()
        
        sug_doc_m = next((c for c in cols_m if 'CNPJ' in c.upper() or 'CPF' in c.upper()), None)
        sug_resp_m = next((c for c in cols_m if 'RESPONSAVEL' in c.upper()), cols_m[0])

        with st.expander("⚙️ Configurar Colunas"):
            c1, c2 = st.columns(2)
            sel_resp_m = c1.selectbox("Responsável:", cols_m, index=cols_m.index(sug_resp_m) if sug_resp_m in cols_m else 0, key='m_resp')
            sel_doc_m = c2.selectbox("CPF/CNPJ:", cols_m, index=cols_m.index(sug_doc_m) if sug_doc_m in cols_m else 0, key='m_doc')

        if st.button("📦 PROCESSAR MAILING", key='btn_mailing'):
            with st.spinner("Processando..."):
                df_trabalho_m = df_m.copy()
                if aplicar_balanceamento:
                    df_trabalho_m, _, _ = distribuir_leads_orfãos(df_trabalho_m, sel_resp_m)
                
                df_classificado = processar_distribuicao_mailing(df_trabalho_m, sel_doc_m, sel_resp_m)
                zip_m = gerar_zip_dinamico(df_classificado, sel_resp_m, "PERFIL_CALCULADO", "MAILING")
                st.success("Pronto!")
                st.download_button("📥 BAIXAR MAILING", zip_m, "Mailing.zip", "application/zip")
else:
    st.info("Aguardando arquivo...")
