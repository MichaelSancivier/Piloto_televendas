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
    .audit-box-success { padding: 10px; background-color: #d4edda; color: #155724; border-radius: 5px; border: 1px solid #c3e6cb; margin-bottom: 10px; }
    .audit-box-danger { padding: 10px; background-color: #f8d7da; color: #721c24; border-radius: 5px; border: 1px solid #f5c6cb; margin-bottom: 10px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. INTELIGÊNCIA LÓGICA (O CÉREBRO)
# ==============================================================================

def normalizar_nome_arquivo(nome):
    if pd.isna(nome): return "SEM_CARTEIRA"
    nfkd = unicodedata.normalize('NFKD', str(nome))
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    limpo = re.sub(r'[^a-zA-Z0-9]', '_', sem_acento.upper())
    return re.sub(r'_+', '_', limpo).strip('_')

def identificar_perfil_pelo_doc(valor):
    """ 
    REGRA DE OURO:
    - 14 Dígitos = Frotista (Manhã)
    - Qualquer outra coisa = Freteiro (Almoço)
    """
    if pd.isna(valor): return "FRETEIRO"
    doc_limpo = re.sub(r'\D', '', str(valor))
    if len(doc_limpo) == 14: return "PEQUENO FROTISTA"
    else: return "FRETEIRO"

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
    # 1. Classifica PRIMEIRO (Frotista vs Freteiro)
    df_trab['ESTRATEGIA_PERFIL'] = df_trab[col_doc].apply(identificar_perfil_pelo_doc)
    
    # 2. Verticaliza
    cols_para_manter = col_id + [col_resp, 'ESTRATEGIA_PERFIL']
    df_melted = df_trab.melt(id_vars=cols_para_manter, value_vars=cols_tel, var_name='Origem', value_name='Telefone_Bruto')
    
    # 3. Trata Telefones
    df_melted['Telefone_Tratado'], df_melted['Tipo'] = zip(*df_melted['Telefone_Bruto'].apply(tratar_celular_discador))
    df_final = df_melted.dropna(subset=['Telefone_Tratado'])
    df_final = df_final.drop_duplicates(subset=col_id + ['Telefone_Tratado'])
    return df_final

def processar_distribuicao_mailing(df, col_doc, col_resp):
    df_proc = df.copy()
    df_proc['PERFIL_CALCULADO'] = df_proc[col_doc].apply(identificar_perfil_pelo_doc)
    return df_proc

def processar_feedback_tarde(df_mestre, df_log, col_id_mestre, col_id_log, col_resp, col_doc):
    ids_trabalhados = df_log[col_id_log].astype(str).unique()
    col_id_uso = col_id_mestre[0]
    df_mestre['TEMP_ID_MATCH'] = df_mestre[col_id_uso].astype(str)
    
    df_pendente = df_mestre[~df_mestre['TEMP_ID_MATCH'].isin(ids_trabalhados)].copy()
    
    # Tarde = Apenas Frotistas
    df_pendente['ESTRATEGIA'] = df_pendente[col_doc].apply(identificar_perfil_pelo_doc)
    df_tarde = df_pendente[df_pendente['ESTRATEGIA'] == 'PEQUENO FROTISTA']
    
    return df_tarde, len(ids_trabalhados)

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

            # MODO TARDE (REFORÇO)
            if modo == "FEEDBACK_TARDE":
                data = io.BytesIO()
                df_agente.to_excel(data, index=False)
                zip_file.writestr(f"DISCADOR_{nome_arquivo}_REFORCO_TARDE_Frotista.xlsx", data.getvalue())
            
            # MODO MANHÃ (DISCADOR) E MAILING - SEPARAÇÃO OBRIGATÓRIA
            else:
                col_seg_uso = 'ESTRATEGIA_PERFIL' if modo == "DISCADOR" else col_segmentacao
                
                df_frotista = df_agente[df_agente[col_seg_uso] == "PEQUENO FROTISTA"]
                df_freteiro = df_agente[df_agente[col_seg_uso] == "FRETEIRO"]
                
                prefixo = "DISCADOR" if modo == "DISCADOR" else "MAILING"
                pasta = nome_arquivo
                
                # Gera arquivo 1: MANHÃ
                if not df_frotista.empty:
                    data = io.BytesIO()
                    df_frotista.to_excel(data, index=False)
                    zip_file.writestr(f"{pasta}/{prefixo}_1_MANHA_Frotista_CNPJ.xlsx", data.getvalue())
                
                # Gera arquivo 2: ALMOÇO
                if not df_freteiro.empty:
                    data = io.BytesIO()
                    df_freteiro.to_excel(data, index=False)
                    zip_file.writestr(f"{pasta}/{prefixo}_2_ALMOCO_Freteiro_CPF.xlsx", data.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# ==============================================================================
# 4. FRONTEND CENTRALIZADO
# ==============================================================================

st.title("🚛 Michelin Pilot Command Center")
st.markdown("### Estratégia de Televentas & Logística")
st.markdown("---")

# BARRA LATERAL
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    st.header("🎮 Controle da Missão")
    modo_operacao = st.radio("Qual o turno atual?", ("🌅 Manhã (Carga Inicial)", "☀️ Tarde (Reprocessamento)"))
    
    st.markdown("---")
    st.subheader("1. Arquivo Mestre")
    uploaded_file = st.file_uploader("Solte o Excel Geral aqui", type=["xlsx"], key="main_upload")
    
    uploaded_log = None
    if modo_operacao == "☀️ Tarde (Reprocessamento)":
        st.markdown("---")
        st.subheader("2. Log do Discador")
        st.info("Necessário para remover quem já foi atendido.")
        uploaded_log = st.file_uploader("Solte o CSV/Excel do Log aqui", type=["csv", "xlsx"], key="log_upload")

# ÁREA PRINCIPAL
if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    all_sheets = xls.sheet_names
    
    # AUDITORIA
    aba_d_audit = next((s for s in all_sheets if 'DISCADOR' in s.upper()), None)
    aba_m_audit = next((s for s in all_sheets if 'MAILING' in s.upper()), None)
    
    if aba_d_audit and aba_m_audit:
        df_audit_d = pd.read_excel(uploaded_file, sheet_name=aba_d_audit)
        df_audit_m = pd.read_excel(uploaded_file, sheet_name=aba_m_audit)
        diff = abs(len(df_audit_d) - len(df_audit_m))
        col_aud1, col_aud2 = st.columns([3,1])
        if diff != 0: col_aud1.warning(f"⚠️ Atenção: Divergência de {diff} linhas entre as abas.")
    
    # LÓGICA DE EXIBIÇÃO
    if modo_operacao == "🌅 Manhã (Carga Inicial)":
        tab1, tab2 = st.tabs(["🤖 DISCADOR (Manhã)", "👩‍💼 MAILING"])
        
        # DISCADOR MANHÃ
        with tab1:
            col_sel, _ = st.columns([1,1])
            aba_d = col_sel.selectbox("Aba Discador:", all_sheets, index=all_sheets.index(aba_d_audit) if aba_d_audit else 0, key='d1')
            df_d = pd.read_excel(uploaded_file, sheet_name=aba_d)
            cols_d = df_d.columns.tolist()
            
            sug_tel = [c for c in cols_d if any(x in c.upper() for x in ['TEL','CEL','FONE'])]
            sug_id = [c for c in cols_d if any(x in c.upper() for x in ['ID','NOME'])]
            sug_doc = next((c for c in cols_d if 'CNPJ' in c.upper() or 'CPF' in c.upper()), None)
            sug_resp = next((c for c in cols_d if 'RESPONSAVEL' in c.upper()), cols_d[0])
            
            with st.expander("⚙️ Configurar Colunas", expanded=True):
                c1, c2 = st.columns(2)
                sel_resp_d = c1.selectbox("Responsável:", cols_d, index=cols_d.index(sug_resp) if sug_resp in cols_d else 0, key='d_resp')
                sel_doc_d = c1.selectbox("CPF/CNPJ (Estratégia):", cols_d, index=cols_d.index(sug_doc) if sug_doc in cols_d else 0, key='d_doc')
                sel_id_d = c2.multiselect("ID Cliente:", cols_d, default=sug_id[:3], key='d_id')
                sel_tel_d = c2.multiselect("Telefones:", cols_d, default=sug_tel, key='d_tel')
                aplicar_bal_d = st.checkbox("Distribuir leads sem dono?", value=True, key='chk_bal_d')

            if st.button("🚀 GERAR PACK MANHÃ (Frotista + Freteiro)", key='btn_discador'):
                with st.spinner("Segmentando e Empacotando..."):
                    df_trabalho = df_d.copy()
                    if aplicar_bal_d: df_trabalho, _, _ = distribuir_leads_orfãos(df_d, sel_resp_d)
                    
                    df_res_d = processar_discador(df_trabalho, sel_id_d, sel_tel_d, sel_resp_d, sel_doc_d)
                    
                    # Contagem para tranquilizar o usuário
                    qtd_frot = len(df_res_d[df_res_d['ESTRATEGIA_PERFIL'] == 'PEQUENO FROTISTA'])
                    qtd_fret = len(df_res_d[df_res_d['ESTRATEGIA_PERFIL'] == 'FRETEIRO'])
                    
                    zip_d = gerar_zip_dinamico(df_res_d, sel_resp_d, None, "DISCADOR") 
                    st.success(f"Pack Gerado! {qtd_frot} Frotistas (Manhã) e {qtd_fret} Freteiros (Almoço).")
                    st.download_button("📥 BAIXAR PACK COMPLETO (.ZIP)", zip_d, "Discador_Manha_Completo.zip", "application/zip")

        # MAILING
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
                aplicar_bal_m = st.checkbox("Distribuir leads sem dono?", value=True, key='chk_bal_m')
            
            if st.button("📦 PROCESSAR MAILING", key='btn_mailing'):
                with st.spinner("Gerando Mailing..."):
                    df_trabalho_m = df_m.copy()
                    if aplicar_bal_m: df_trabalho_m, _, _ = distribuir_leads_orfãos(df_m, sel_resp_m)
                    df_classificado = processar_distribuicao_mailing(df_trabalho_m, sel_doc_m, sel_resp_m)
                    zip_m = gerar_zip_dinamico(df_classificado, sel_resp_m, "PERFIL_CALCULADO", "MAILING")
                    st.success("Mailing Pronto!")
                    st.download_button("📥 BAIXAR MAILING", zip_m, "Mailing.zip", "application/zip")

    elif modo_operacao == "☀️ Tarde (Reprocessamento)":
        st.subheader("🔄 Reprocessamento para Tarde (14:00)")
        if not uploaded_log:
            st.info("👈 Por favor, carregue o **Log do Discador** na barra lateral para continuar.")
        else:
            try:
                if uploaded_log.name.endswith('.csv'): df_log = pd.read_csv(uploaded_log, sep=None, engine='python')
                else: df_log = pd.read_excel(uploaded_log)
                st.markdown(f'<div class="audit-box-success">✅ Log Carregado: {len(df_log)} registros.</div>', unsafe_allow_html=True)
                
                aba_d = next((s for s in all_sheets if 'DISCADOR' in s.upper()), all_sheets[0])
                df_d = pd.read_excel(uploaded_file, sheet_name=aba_d)
                cols_d = df_d.columns.tolist()
                sug_tel = [c for c in cols_d if any(x in c.upper() for x in ['TEL','CEL','FONE'])]
                sug_id = [c for c in cols_d if any(x in c.upper() for x in ['ID','NOME'])]
                sug_doc = next((c for c in cols_d if 'CNPJ' in c.upper() or 'CPF' in c.upper()), None)
                sug_resp = next((c for c in cols_d if 'RESPONSAVEL' in c.upper()), cols_d[0])

                with st.expander("⚙️ Configuração de Match", expanded=True):
                    c1, c2 = st.columns(2)
                    st.markdown("**Base Mestre:**")
                    sel_resp_d = c1.selectbox("Responsável:", cols_d, index=cols_d.index(sug_resp) if sug_resp in cols_d else 0, key='t_resp')
                    sel_doc_d = c1.selectbox("CPF/CNPJ:", cols_d, index=cols_d.index(sug_doc) if sug_doc in cols_d else 0, key='t_doc')
                    sel_id_d = c1.multiselect("ID Cliente:", cols_d, default=sug_id[:1], key='t_id')
                    sel_tel_d = c1.multiselect("Telefones:", cols_d, default=sug_tel, key='t_tel')
                    col_id_log = c2.selectbox("Coluna ID no Log:", df_log.columns)

                if st.button("🔄 GERAR LISTA TARDE"):
                    with st.spinner("Processando..."):
                        df_trabalho, _, _ = distribuir_leads_orfãos(df_d, sel_resp_d)
                        df_tarde_limpa, qtd = processar_feedback_tarde(df_trabalho, df_log, sel_id_d, col_id_log, sel_resp_d, sel_doc_d)
                        df_res_tarde = processar_discador(df_tarde_limpa, sel_id_d, sel_tel_d, sel_resp_d, sel_doc_d)
                        zip_tarde = gerar_zip_dinamico(df_res_tarde, sel_resp_d, None, "FEEDBACK_TARDE")
                        st.success(f"Lista da Tarde Pronta! {qtd} contatos removidos.")
                        st.download_button("📥 BAIXAR PACK TARDE", zip_tarde, "Reforco_Tarde.zip", "application/zip")
            except Exception as e:
                st.error(f"Erro: {e}")
else:
    st.info("👈 Comece carregando o arquivo na barra lateral.")
