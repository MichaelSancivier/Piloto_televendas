import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np # Importante para a distribuição matemática

# ==============================================================================
# 1. CONFIGURAÇÃO VISUAL E TEMA
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
    .destaque-distribuicao { padding: 10px; background-color: #e8f4f8; border-radius: 5px; border-left: 5px solid #00aabb; margin-bottom: 20px;}
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
    """ 11 Dígitos = Freteiro (CPF) | 14 Dígitos = Frotista (CNPJ) """
    if pd.isna(valor): return "INDEFINIDO"
    doc_limpo = re.sub(r'\D', '', str(valor))
    if len(doc_limpo) == 11: return "FRETEIRO"
    elif len(doc_limpo) == 14: return "PEQUENO FROTISTA"
    else: return "OUTROS"

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
# 3. MOTOR DE DISTRIBUIÇÃO EQUITATIVA (NOVO!!!)
# ==============================================================================

def distribuir_leads_orfãos(df, col_resp):
    """
    Pega linhas onde o Responsável é Vazio, Nulo ou 'CANAL TELEVENDAS'
    e distribui igualmente entre os atendentes humanos encontrados.
    """
    df_proc = df.copy()
    
    # 1. Identificar Atendentes Válidos (Humanos)
    # Remove nulos e termos genéricos de sistema
    termos_ignorar = ['CANAL TELEVENDAS', 'TIME', 'EQUIPE', 'TELEVENDAS', 'NULL', 'NAN', '']
    
    # Pega todos os valores únicos da coluna
    todos_resp = df_proc[col_resp].unique()
    
    # Filtra apenas os que parecem ser pessoas
    agentes_humanos = [
        a for a in todos_resp 
        if pd.notna(a) and str(a).strip().upper() not in termos_ignorar
    ]
    
    if not agentes_humanos:
        return df_proc, "Nenhum atendente humano encontrado para receber a carga.", 0

    # 2. Identificar Órfãos (Quem precisa de dono)
    # Considera órfão quem é NaN ou está na lista de termos genéricos
    mask_orfao = df_proc[col_resp].isna() | df_proc[col_resp].astype(str).str.strip().str.upper().isin(termos_ignorar)
    
    qtd_orfaos = mask_orfao.sum()
    
    if qtd_orfaos == 0:
        return df_proc, "Nenhum lead órfão encontrado. A base já está completa.", 0

    # 3. A Distribuição (Round Robin)
    # Cria uma lista repetindo os agentes o suficiente para cobrir os órfãos
    # Ex: [Lilian, Susana, Lilian, Susana, Lilian...]
    atribuicoes = np.resize(agentes_humanos, qtd_orfaos)
    
    # Aplica na coluna
    df_proc.loc[mask_orfao, col_resp] = atribuicoes
    
    msg = f"Sucesso! {qtd_orfaos} leads sem dono foram distribuídos entre: {', '.join(map(str, agentes_humanos))}."
    return df_proc, msg, qtd_orfaos

# ==============================================================================
# 4. PROCESSADORES DE ARQUIVO
# ==============================================================================

def processar_discador(df, col_id, cols_tel, col_resp):
    df_melted = df.melt(id_vars=col_id + [col_resp], value_vars=cols_tel, 
                        var_name='Origem', value_name='Telefone_Bruto')
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
            
        agentes_encontrados = df_dados[col_resp].unique()

        for agente in agentes_encontrados:
            nome_arquivo = normalizar_nome_arquivo(agente)
            if pd.isna(agente):
                df_agente = df_dados[df_dados[col_resp].isna()]
            else:
                df_agente = df_dados[df_dados[col_resp] == agente]
            if df_agente.empty: continue

            if modo == "MAILING":
                df_frotista = df_agente[df_agente[col_segmentacao] == "PEQUENO FROTISTA"]
                df_freteiro = df_agente[df_agente[col_segmentacao] == "FRETEIRO"]
                ids_ok = df_frotista.index.union(df_freteiro.index)
                df_outros = df_agente.drop(ids_ok)
                
                if not df_frotista.empty:
                    data = io.BytesIO()
                    df_frotista.to_excel(data, index=False)
                    zip_file.writestr(f"{nome_arquivo}/1_MANHA_Frotistas_CNPJ.xlsx", data.getvalue())
                if not df_freteiro.empty:
                    data = io.BytesIO()
                    df_freteiro.to_excel(data, index=False)
                    zip_file.writestr(f"{nome_arquivo}/2_ALMOCO_Freteiros_CPF.xlsx", data.getvalue())
                if not df_outros.empty:
                    data = io.BytesIO()
                    df_outros.to_excel(data, index=False)
                    zip_file.writestr(f"{nome_arquivo}/3_BACKLOG_Verificar.xlsx", data.getvalue())
            else: 
                data = io.BytesIO()
                df_agente.to_excel(data, index=False)
                zip_file.writestr(f"DISCADOR_{nome_arquivo}.xlsx", data.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

# ==============================================================================
# 5. FRONTEND (INTERFACE DE COMANDO)
# ==============================================================================

st.title("🚛 Michelin Pilot Command Center")
st.markdown("### Estratégia de Televentas & Logística")
st.markdown("---")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    st.header("📂 Entrada de Dados")
    uploaded_file = st.file_uploader("Arraste seu Excel aqui (.xlsx)", type=["xlsx"])
    st.caption("O arquivo deve conter as abas do Discador e do Mailing.")

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    all_sheets = xls.sheet_names
    
    # --------------------------------------------------------------------------
    # NOVO: ÁREA DE PRÉ-PROCESSAMENTO (BALANCEAMENTO DE CARGA)
    # --------------------------------------------------------------------------
    st.markdown('<div class="destaque-distribuicao">', unsafe_allow_html=True)
    st.markdown("### ⚖️ Balanceamento de Carga (Opcional)")
    st.markdown("Antes de processar, deseja distribuir leads sem dono (Prioridades Novas) para a equipe?")
    
    col_bal_1, col_bal_2, col_bal_3 = st.columns([1, 2, 1])
    aplicar_balanceamento = col_bal_1.checkbox("Sim, distribuir leads órfãos", value=True)
    
    if aplicar_balanceamento:
        st.caption("O sistema irá procurar por 'CANAL TELEVENDAS' ou células vazias e dividir igualmente entre os nomes encontrados.")
    st.markdown('</div>', unsafe_allow_html=True)
    # --------------------------------------------------------------------------

    tab1, tab2 = st.tabs(["🤖 MODO ROBÔ (Discador)", "👩‍💼 MODO EQUIPE (Mailing)"])
    
    # ABA 1: DISCADOR
    with tab1:
        st.subheader("Preparação para Discador Automático")
        col_sel, _ = st.columns([1,1])
        aba_d = col_sel.selectbox("Selecione a aba do Discador:", all_sheets, index=0, key='d1')
        
        # Carrega DF
        df_d = pd.read_excel(uploaded_file, sheet_name=aba_d)
        cols_d = df_d.columns.tolist()
        
        # Auto-Sugestão
        sug_tel = [c for c in cols_d if any(x in c.upper() for x in ['TEL','CEL','FONE','MOV','COMERCIAL','RESIDENCIAL'])]
        sug_id = [c for c in cols_d if any(x in c.upper() for x in ['ID','NOME','CLIENTE'])]
        sug_resp = next((c for c in cols_d if 'RESPONSAVEL' in c.upper()), cols_d[0])
        
        with st.expander("⚙️ Conferir Colunas (Automático)", expanded=True):
            c1, c2 = st.columns(2)
            sel_id_d = c1.multiselect("Identificação:", cols_d, default=sug_id[:3], key='d_id')
            sel_resp_d = c1.selectbox("Responsável:", cols_d, index=cols_d.index(sug_resp) if sug_resp in cols_d else 0, key='d_resp')
            sel_tel_d = c2.multiselect("Telefones:", cols_d, default=sug_tel, key='d_tel')

        st.write("")
        if st.button("🚀 PROCESSAR DISCADOR", key='btn_discador'):
            if not sel_tel_d:
                st.error("Selecione os telefones.")
            else:
                with st.spinner("Processando..."):
                    
                    # 1. APLICA BALANCEAMENTO SE SOLICITADO
                    df_trabalho = df_d.copy()
                    if aplicar_balanceamento:
                        df_trabalho, msg_bal, qtd_orfaos = distribuir_leads_orfãos(df_trabalho, sel_resp_d)
                        if qtd_orfaos > 0: st.info(msg_bal)

                    # 2. PROCESSA
                    df_res_d = processar_discador(df_trabalho, sel_id_d, sel_tel_d, sel_resp_d)
                    zip_d = gerar_zip_dinamico(df_res_d, sel_resp_d, None, "DISCADOR") 
                    
                    st.success("Pronto!")
                    k1, k2, k3 = st.columns(3)
                    k1.metric("Linhas", len(df_res_d))
                    k2.metric("Celulares", len(df_res_d[df_res_d['Tipo'].str.contains('Celular')]))
                    k3.metric("Fixos", len(df_res_d[df_res_d['Tipo'] == 'Fixo']))
                    
                    st.download_button("📥 BAIXAR DISCADOR (.ZIP)", zip_d, "Pack_Discador.zip", "application/zip")

    # ABA 2: MAILING
    with tab2:
        st.subheader("Distribuição de Mailing (Equipe)")
        col_sel_m, _ = st.columns([1,1])
        aba_m = col_sel_m.selectbox("Selecione a aba de Mailing:", all_sheets, index=1 if len(all_sheets)>1 else 0, key='m1')
        
        df_m = pd.read_excel(uploaded_file, sheet_name=aba_m)
        cols_m = df_m.columns.tolist()
        
        sug_doc = next((c for c in cols_m if 'CNPJ' in c.upper() or 'CPF' in c.upper()), None)
        sug_resp_m = next((c for c in cols_m if 'RESPONSAVEL' in c.upper()), cols_m[0])

        with st.expander("⚙️ Conferir Colunas", expanded=True):
            c1, c2 = st.columns(2)
            sel_resp_m = c1.selectbox("Responsável:", cols_m, index=cols_m.index(sug_resp_m) if sug_resp_m in cols_m else 0, key='m_resp')
            sel_doc_m = c2.selectbox("CPF/CNPJ:", cols_m, index=cols_m.index(sug_doc) if sug_doc in cols_m else 0, key='m_doc')

        st.write("")
        if st.button("📦 DISTRIBUIR MAILING", key='btn_mailing'):
            with st.spinner("Distribuindo e Separando..."):
                
                # 1. APLICA BALANCEAMENTO SE SOLICITADO
                df_trabalho_m = df_m.copy()
                if aplicar_balanceamento:
                    df_trabalho_m, msg_bal, qtd_orfaos = distribuir_leads_orfãos(df_trabalho_m, sel_resp_m)
                    if qtd_orfaos > 0: st.info(msg_bal)

                # 2. PROCESSA
                df_classificado = processar_distribuicao_mailing(df_trabalho_m, sel_doc_m, sel_resp_m)
                zip_m = gerar_zip_dinamico(df_classificado, sel_resp_m, "PERFIL_CALCULADO", "MAILING")
                
                st.success("Sucesso!")
                met1, met2, met3 = st.columns(3)
                met1.metric("Atendentes", df_classificado[sel_resp_m].nunique())
                met2.metric("Frotistas (Manhã)", len(df_classificado[df_classificado['PERFIL_CALCULADO'] == 'PEQUENO FROTISTA']))
                met3.metric("Freteiros (Almoço)", len(df_classificado[df_classificado['PERFIL_CALCULADO'] == 'FRETEIRO']))

                st.download_button("📥 BAIXAR MAILING (.ZIP)", zip_m, "Pack_Mailing.zip", "application/zip")

else:
    st.markdown("""
    <div style="text-align: center; margin-top: 50px; color: #666;">
        <h1>Aguardando Arquivo...</h1>
    </div>
    """, unsafe_allow_html=True)
