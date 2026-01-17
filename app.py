
import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np

# ==============================================================================
# 1. CONFIGURAÇÃO VISUAL (CLEAN)
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
    
    /* Destaque para upload na Tab 3 */
    .upload-log-box { border: 2px dashed #003366; padding: 20px; border-radius: 10px; background-color: #e8f4f8; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. INTELIGÊNCIA LÓGICA
# ==============================================================================

def normalizar_nome_arquivo(nome):
    if pd.isna(nome): return "SEM_CARTEIRA"
    nfkd = unicodedata.normalize('NFKD', str(nome))
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    limpo = re.sub(r'[^a-zA-Z0-9]', '_', sem_acento.upper())
    return re.sub(r'_+', '_', limpo).strip('_')

def identificar_perfil_pelo_doc(valor):
    """ Regra Binária: 14 = Frotista | Outros = Freteiro """
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

def processar_feedback_tarde(df_mestre, df_log, col_id_mestre, col_id_log, col_resp, col_doc):
    ids_trabalhados = df_log[col_id_log].astype(str).unique()
    
    # Normaliza ID Mestre para string para garantir o match
    # Pega o primeiro ID da lista de IDs selecionados pelo usuário
    col_id_uso = col_id_mestre[0]
    df_mestre['TEMP_ID_MATCH'] = df_mestre[col_id_uso].astype(str)
    
    # Filtra: Mantém apenas quem NÃO está na lista de ids_trabalhados
    df_pendente = df_mestre[~df_mestre['TEMP_ID_MATCH'].isin(ids_trabalhados)].copy()
    
    # Filtra estratégia (Apenas Frotistas são o foco da Tarde, pois Freteiros já rodaram no almoço)
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

            if modo == "FEEDBACK_TARDE":
                data = io.BytesIO()
                df_agente.to_excel(data, index=False)
                zip_file.writestr(f"DISCADOR_{nome_arquivo}_REFORCO_TARDE.xlsx", data.getvalue())
            else:
                col_seg_uso = 'ESTRATEGIA_PERFIL' if modo == "DISCADOR" else col_segmentacao
                df_frotista = df_agente[df_agente[col_seg_uso] == "PEQUENO FROTISTA"]
                df_freteiro = df_agente[df_agente[col_seg_uso] == "FRETEIRO"]
                
                prefixo = "DISCADOR
