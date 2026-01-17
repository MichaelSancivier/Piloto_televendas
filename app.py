import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata

# ==============================================================================
# 1. CONFIGURAÇÃO VISUAL E TEMA (O COCKPIT)
# ==============================================================================
st.set_page_config(
    page_title="Michelin Pilot Command Center",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS para dar cara de Aplicativo Corporativo Profissional
st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    h1 { color: #003366; font-family: 'Helvetica', sans-serif;}
    h2, h3 { color: #2c3e50; }
    .stButton>button { 
        width: 100%; 
        border-radius: 8px; 
        height: 3.5em; 
        font-weight: bold; 
        background-color: #003366; 
        color: white;
        border: none;
    }
    .stButton>button:hover { background-color: #004080; color: white; }
    .stMetric { background-color: white; padding: 15px; border-radius: 8px; border-left: 6px solid #FCE500; box-shadow: 0 2px 5px rgba(0,0,0,0.05);}
    div[data-testid="stExpander"] { background-color: white; border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. INTELIGÊNCIA LÓGICA (O CÉREBRO)
# ==============================================================================

def normalizar_nome_arquivo(nome):
    """
    Padroniza nomes para criar arquivos seguros (sem acentos ou espaços).
    Ex: 'Lílian Rodrigues ' -> 'LILIAN_RODRIGUES'
    """
    if pd.isna(nome): return "SEM_CARTEIRA"
    # Remove acentos
    nfkd = unicodedata.normalize('NFKD', str(nome))
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    # Remove caracteres especiais e espaços, converte para maiúsculo
    limpo = re.sub(r'[^a-zA-Z0-9]', '_', sem_acento.upper())
    # Remove _ repetidos e nas pontas
    return re.sub(r'_+', '_', limpo).strip('_')

def identificar_perfil_pelo_doc(valor):
    """ 
    A REGRA DE OURO DA LOGÍSTICA:
    11 Dígitos = CPF = FRETEIRO (Autônomo) -> Horário de Almoço
    14 Dígitos = CNPJ = PEQUENO FROTISTA (Empresário) -> Horário Manhã
    """
    if pd.isna(valor): return "INDEFINIDO"
    # Mantém apenas números
    doc_limpo = re.sub(r'\D', '', str(valor))
    
    if len(doc_limpo) == 11: return "FRETEIRO"
    elif len(doc_limpo) == 14: return "PEQUENO FROTISTA"
    else: return "OUTROS"

def tratar_celular_discador(val):
    """
    Aplica Regra do 9º Dígito, Formato +55 e distingue Fixo de Celular.
    """
    if pd.isna(val): return None, "Vazio"
    nums = re.sub(r'\D', '', str(val))
    
    # Remove DDI 55 se já vier no começo
    if nums.startswith('55') and len(nums) >= 12: nums = nums[2:]
    
    tipo = "Inválido"
    final = None
    
    # Celular Correto (11 dígitos, começa com 9)
    if len(nums) == 11 and int(nums[2]) == 9:
        tipo = "Celular"
        final = nums
    # Celular Antigo (10 dígitos, falta o 9) -> Começa com 6,7,8,9
    elif len(nums) == 10 and int(nums[2]) >= 6:
        tipo = "Celular (Corrigido)"
        final = nums[:2] + '9' + nums[2:]
    # Fixo (10 dígitos) -> Começa com 2,3,4,5
    elif len(nums) == 10 and 2 <= int(nums[2]) <= 5:
        tipo = "Fixo"
        final = nums
        
    if final: return f"+55{final}", tipo
    return None, tipo

# ==============================================================================
# 3. PROCESSADORES DE DADOS (OS MOTORES)
# ==============================================================================

def processar_discador(df, col_id, cols_tel, col_resp):
    # Melt: Transforma colunas de telefone em linhas
    df_melted = df.melt(id_vars=col_id + [col_resp], value_vars=cols_tel, 
                        var_name='Origem', value_name='Telefone_Bruto')
    
    # Aplica tratamento linha a linha
    df_melted['Telefone_Tratado'], df_melted['Tipo'] = zip(*df_melted['Telefone_Bruto'].apply(tratar_celular_discador))
    
    # Remove inválidos e duplicatas
    df_final = df_melted.dropna(subset=['Telefone_Tratado'])
    df_final = df_final.drop_duplicates(subset=col_id + ['Telefone_Tratado'])
    return df_final

def processar_distribuicao_mailing(df, col_doc, col_resp):
    df_proc = df.copy()
    # Cria coluna calculada baseada no CPF/CNPJ
    df_proc['PERFIL_CALCULADO'] = df_proc[col_doc].apply(identificar_perfil_pelo_doc)
    return df_proc

def gerar_zip_dinamico(df_dados, col_resp, col_segmentacao, modo="DISCADOR"):
    """
    Cria um arquivo ZIP contendo N arquivos Excel, separados por atendente e estratégia.
    """
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        
        # Garante coluna responsável
        if col_resp not in df_dados.columns:
            df_dados['RESP_GERAL'] = 'EQUIPE'
            col_resp = 'RESP_GERAL'
            
        # Descobre quem está na planilha (Auto-Discovery)
        agentes_encontrados = df_dados[col_resp].unique()

        for agente in agentes_encontrados:
            nome_arquivo = normalizar_nome_arquivo(agente)
            
            # Filtra dados do agente
            if pd.isna(agente):
                df_agente = df_dados[df_dados[col_resp].isna()]
            else:
                df_agente = df_dados[df_dados[col_resp] == agente]
            
            if df_agente.empty: continue

            # --- ESTRATÉGIA DE MAILING (MANHÃ/ALMOÇO) ---
            if modo == "MAILING":
                df_frotista = df_agente[df_agente[col_segmentacao] == "PEQUENO FROTISTA"] # CNPJ
                df_freteiro = df_agente[df_agente[col_segmentacao] == "FRETEIRO"]         # CPF
                
                # O resto vai para backlog
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
            
            # --- ESTRATÉGIA DE DISCADOR (ARQUIVO ÚNICO) ---
            else: 
                data = io.BytesIO()
                df_agente.to_excel(data, index=False)
                zip_file.writestr(f"DISCADOR_{nome_arquivo}.xlsx", data.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# ==============================================================================
# 4. INTERFACE DO USUÁRIO (FRONTEND)
# ==============================================================================

st.title("🚛 Michelin Pilot Command Center")
st.markdown("### Estratégia de Televentas & Logística")
st.markdown("---")

# --- SIDEBAR DE UPLOAD ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    st.header("📂 Entrada de Dados")
    uploaded_file = st.file_uploader("Arraste seu Excel aqui (.xlsx)", type=["xlsx"])
    st.caption("O arquivo deve conter as abas do Discador e do Mailing.")

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    all_sheets = xls.sheet_names
    
    # Abas de Navegação
    tab1, tab2 = st.tabs(["🤖 MODO ROBÔ (Discador)", "👩‍💼 MODO EQUIPE (Mailing)"])
    
    # ==========================================
    # ABA 1: DISCADOR
    # ==========================================
    with tab1:
        st.subheader("Preparação para Discador Automático")
        st.info("Este módulo verticaliza a base (1 linha por telefone), corrige números e separa arquivos por vendedor.")
        
        col_sel, col_vazio = st.columns([1,1])
        aba_d = col_sel.selectbox("Selecione a aba do Discador:", all_sheets, index=0, key='d1')
        
        df_d = pd.read_excel(uploaded_file, sheet_name=aba_d)
        cols_d = df_d.columns.tolist()
        
        # Sugestões Automáticas
        sug_tel = [c for c in cols_d if any(x in c.upper() for x in ['TEL','CEL','FONE','MOV','COMERCIAL','RESIDENCIAL'])]
        sug_id = [c for c in cols_d if any(x in c.upper() for x in ['ID','NOME','CLIENTE'])]
        sug_resp = next((c for c in cols_d if 'RESPONSAVEL' in c.upper()), cols_d[0])
        
        with st.expander("⚙️ Conferir Colunas (Automático)", expanded=True):
            c1, c2 = st.columns(2)
            sel_id_d = c1.multiselect("Identificação do Cliente:", cols_d, default=sug_id[:3])
            sel_resp_d = c1.selectbox("Responsável (Vendedor):", cols_d, index=cols_d.index(sug_resp) if sug_resp in cols_d else 0)
            sel_tel_d = c2.multiselect("Colunas de Telefone:", cols_d, default=sug_tel)

        st.write("")
        if st.button("🚀 PROCESSAR DISCADOR", key='btn_discador'):
            if not sel_tel_d:
                st.error("Selecione pelo menos uma coluna de telefone.")
            else:
                with st.spinner("O Robô está higienizando e verticalizando os dados..."):
                    df_res_d = processar_discador(df_d, sel_id_d, sel_tel_d, sel_resp_d)
                    zip_d = gerar_zip_dinamico(df_res_d, sel_resp_d, None, "DISCADOR") 
                    
                    st.success("Processamento Concluído!")
                    
                    # Métricas
                    k1, k2, k3 = st.columns(3)
                    k1.metric("Linhas Geradas", len(df_res_d))
                    k2.metric("Celulares (+55)", len(df_res_d[df_res_d['Tipo'].str.contains('Celular')]))
                    k3.metric("Fixos (+55)", len(df_res_d[df_res_d['Tipo'] == 'Fixo']))
                    
                    st.download_button("📥 BAIXAR ARQUIVOS DISCADOR (.ZIP)", zip_d, "Pack_Discador_Pronto.zip", "application/zip")

    # ==========================================
    # ABA 2: MAILING (DISTRIBUIÇÃO)
    # ==========================================
    with tab2:
        st.subheader("Distribuição Estratégica de Mailing")
        st.info("Este módulo lê o CPF/CNPJ e separa automaticamente os arquivos: Frotistas (Manhã) e Freteiros (Almoço).")
        
        col_sel_m, col_vazio_m = st.columns([1,1])
        aba_m = col_sel_m.selectbox("Selecione a aba de Mailing:", all_sheets, index=1 if len(all_sheets)>1 else 0, key='m1')
        
        df_m = pd.read_excel(uploaded_file, sheet_name=aba_m)
        cols_m = df_m.columns.tolist()
        
        # Sugestões Automáticas
        sug_doc = next((c for c in cols_m if 'CNPJ' in c.upper() or 'CPF' in c.upper()), None)
        sug_resp_m = next((c for c in cols_m if 'RESPONSAVEL' in c.upper()), cols_m[0])

        with st.expander("⚙️ Conferir Colunas (Automático)", expanded=True):
            c1, c2 = st.columns(2)
            sel_resp_m = c1.selectbox("Coluna Responsável (Quem recebe):", cols_m, index=cols_m.index(sug_resp_m) if sug_resp_m in cols_m else 0)
            sel_doc_m = c2.selectbox("Coluna CPF/CNPJ (O Robô analisa isso):", cols_m, index=cols_m.index(sug_doc) if sug_doc in cols_m else 0)

        st.write("")
        if st.button("📦 DISTRIBUIR MAILING", key='btn_mailing'):
            with st.spinner("Classificando clientes e separando pastas..."):
                
                # Processamento
                df_classificado = processar_distribuicao_mailing(df_m, sel_doc_m, sel_resp_m)
                zip_m = gerar_zip_dinamico(df_classificado, sel_resp_m, "PERFIL_CALCULADO", "MAILING")
                
                st.success("Distribuição finalizada com sucesso!")
                
                # Métricas da Operação
                total_agentes = df_classificado[sel_resp_m].nunique()
                frotistas = len(df_classificado[df_classificado['PERFIL_CALCULADO'] == 'PEQUENO FROTISTA'])
                freteiros = len(df_classificado[df_classificado['PERFIL_CALCULADO'] == 'FRETEIRO'])

                met1, met2, met3 = st.columns(3)
                met1.metric("Atendentes Encontrados", total_agentes)
                met2.metric("Frotistas (CNPJ - Manhã)", frotistas)
                met3.metric("Freteiros (CPF - Almoço)", freteiros)

                st.download_button("📥 BAIXAR PACK MAILING (.ZIP)", zip_m, "Pack_Mailing_Distribuido.zip", "application/zip")

else:
    # Tela de Boas Vindas (Vazia)
    st.markdown("""
    <div style="text-align: center; margin-top: 50px; color: #666;">
        <h1>Aguardando Arquivo...</h1>
        <p>Faça o upload do Excel na barra lateral esquerda para iniciar.</p>
    </div>
    """, unsafe_allow_html=True)
