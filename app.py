import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np
import random

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
    
    /* Caixas de Status */
    .audit-box-success { padding: 15px; background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; border-radius: 5px; margin-bottom: 10px; }
    .audit-box-danger { padding: 15px; background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; border-radius: 5px; margin-bottom: 10px; }
    .preview-box { background-color: #ffffff; padding: 20px; border-radius: 10px; border: 1px solid #e0e0e0; margin-top: 20px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNÇÕES AUXILIARES (HIGIENE DE DADOS)
# ==============================================================================

def normalizar_nome_arquivo(nome):
    """Remove acentos e caracteres especiais para nome de arquivo seguro."""
    if pd.isna(nome): return "SEM_DONO"
    nfkd = unicodedata.normalize('NFKD', str(nome))
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    limpo = re.sub(r'[^a-zA-Z0-9]', '_', sem_acento.upper())
    return re.sub(r'_+', '_', limpo).strip('_')

def identificar_perfil_doc(valor):
    """
    REGRA DE NEGÓCIO:
    - 14 Dígitos = PEQUENO FROTISTA (Manhã)
    - Qualquer outra coisa = FRETEIRO (Almoço)
    """
    if pd.isna(valor): return "FRETEIRO"
    val_str = str(valor)
    if val_str.endswith('.0'): val_str = val_str[:-2]
    doc_limpo = re.sub(r'\D', '', val_str)
    
    if len(doc_limpo) == 14: return "PEQUENO FROTISTA"
    else: return "FRETEIRO"

def tratar_telefone(val):
    """Padroniza telefones para +55... e identifica tipo."""
    if pd.isna(val): return None, "Vazio"
    nums = re.sub(r'\D', '', str(val))
    if nums.startswith('55') and len(nums) >= 12: nums = nums[2:]
    
    final = None
    tipo = "Inválido"
    
    if len(nums) == 11 and int(nums[2]) == 9:
        tipo = "Celular"
        final = nums
    elif len(nums) == 10 and int(nums[2]) >= 6: # Celular antigo sem 9
        tipo = "Celular (Corrigido)"
        final = nums[:2] + '9' + nums[2:]
    elif len(nums) == 10 and 2 <= int(nums[2]) <= 5: # Fixo
        tipo = "Fixo"
        final = nums
        
    return (f"+55{final}", tipo) if final else (None, tipo)

# ==============================================================================
# 3. MOTOR DE DISTRIBUIÇÃO INTELIGENTE
# ==============================================================================

def distribuir_carga_equitativa(df, col_resp, col_prio=None):
    """
    Distribui leads 'sem dono' garantindo que cada atendente receba
    a mesma quantidade de Prioridade 1, 2, 3, etc.
    """
    df_proc = df.copy()
    
    # 1. Identifica Agentes Válidos
    termos_ignorar = ['CANAL TELEVENDAS', 'TIME', 'EQUIPE', 'TELEVENDAS', 'NULL', 'NAN', '', 'BACKLOG']
    todos_nomes = df_proc[col_resp].unique()
    agentes_validos = [n for n in todos_nomes if pd.notna(n) and str(n).strip().upper() not in termos_ignorar]
    
    if not agentes_validos:
        return df_proc, "⚠️ Erro: Nenhum atendente humano encontrado."

    # 2. Identifica Órfãos (Leads para distribuir)
    mask_orfao = df_proc[col_resp].isna() | df_proc[col_resp].astype(str).str.strip().str.upper().isin(termos_ignorar)
    
    if mask_orfao.sum() == 0:
        return df_proc, "Base já está totalmente atribuída."

    # 3. Lógica de Camadas (Prioridades)
    # Se não tiver prioridade, cria uma dummy para usar o mesmo loop
    use_prio = False
    if col_prio and col_prio in df_proc.columns:
        use_prio = True
        lista_prios = sorted(df_proc.loc[mask_orfao, col_prio].unique())
    else:
        lista_prios = [1] # Camada única
        
    for p in lista_prios:
        # Filtra órfãos desta camada específica
        if use_prio:
            mask_camada = mask_orfao & (df_proc[col_prio] == p)
        else:
            mask_camada = mask_orfao
            
        indices = df_proc[mask_camada].index.tolist()
        qtd = len(indices)
        
        if qtd > 0:
            # EMBARALHA a ordem dos índices para não viciar na ordem do Excel
            random.shuffle(indices)
            
            # EMBARALHA a equipe a cada rodada. 
            # Isso garante que se sobrar 1 lead Prio-1, ele caia aleatoriamente para alguém,
            # e não sempre para o primeiro da lista alfabética.
            equipe_rodada = agentes_validos.copy()
            random.shuffle(equipe_rodada)
            
            atribuicoes = np.resize(equipe_rodada, qtd)
            df_proc.loc[indices, col_resp] = atribuicoes
            
    msg = f"✅ Distribuição Concluída! {mask_orfao.sum()} leads alocados."
    if use_prio: msg += f" (Balanceado por {col_prio})"
    
    return df_proc, msg

# ==============================================================================
# 4. PROCESSADORES DE ARQUIVO
# ==============================================================================

def processar_verticalizacao_discador(df, col_id, cols_tel, col_resp, col_doc, col_prio):
    df_trab = df.copy()
    # Define Estratégia (Frotista/Freteiro) ANTES de verticalizar
    df_trab['ESTRATEGIA_PERFIL'] = df_trab[col_doc].apply(identificar_perfil_doc)
    
    # Colunas para manter no output
    cols_fixas = list(set(col_id + [col_resp, 'ESTRATEGIA_PERFIL'] + ([col_prio] if col_prio else [])))
    
    # Verticaliza (Melt)
    df_melt = df_trab.melt(id_vars=cols_fixas, value_vars=cols_tel, var_name='Origem', value_name='Telefone_Bruto')
    df_melt['Telefone_Tratado'], df_melt['Tipo'] = zip(*df_melt['Telefone_Bruto'].apply(tratar_telefone))
    
    # Limpa inválidos e duplicados
    df_final = df_melt.dropna(subset=['Telefone_Tratado']).drop_duplicates(subset=col_id + ['Telefone_Tratado'])
    return df_final

def processar_mailing_simples(df, col_doc):
    df_proc = df.copy()
    df_proc['PERFIL_CALCULADO'] = df_proc[col_doc].apply(identificar_perfil_doc)
    return df_proc

def processar_feedback_tarde(df_mestre, df_log, col_id_mestre, col_id_log, col_resp, col_doc):
    # IDs já trabalhados
    ids_trabalhados = df_log[col_id_log].astype(str).unique()
    
    # IDs Mestre
    col_id_uso = col_id_mestre[0]
    df_mestre['TEMP_ID_MATCH'] = df_mestre[col_id_uso].astype(str)
    
    # Filtra: Remove trabalhados
    df_pendente = df_mestre[~df_mestre['TEMP_ID_MATCH'].isin(ids_trabalhados)].copy()
    
    # Filtra: Apenas Frotistas para a Tarde
    df_pendente['ESTRATEGIA'] = df_pendente[col_doc].apply(identificar_perfil_doc)
    df_tarde = df_pendente[df_pendente['ESTRATEGIA'] == 'PEQUENO FROTISTA']
    
    return df_tarde, len(ids_trabalhados)

def gerar_zip_arquivos(df_discador, df_mailing, col_resp_d, col_resp_m, modo="MANHA"):
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
        
        # --- 1. ARQUIVOS DO DISCADOR ---
        if df_discador is not None and not df_discador.empty:
            agentes = df_discador[col_resp_d].unique()
            for agente in agentes:
                nome_safe = normalizar_nome_arquivo(agente)
                df_agente = df_discador[df_discador[col_resp_d] == agente]
                
                if modo == "TARDE":
                    # Tarde = Apenas 1 arquivo de reforço
                    buf = io.BytesIO()
                    df_agente.to_excel(buf, index=False)
                    zf.writestr(f"{nome_safe}_DISCADOR_REFORCO_TARDE.xlsx", buf.getvalue())
                else:
                    # Manhã = Separação Frotista vs Freteiro
                    frotistas = df_agente[df_agente['ESTRATEGIA_PERFIL'] == "PEQUENO FROTISTA"]
                    freteiros = df_agente[df_agente['ESTRATEGIA_PERFIL'] == "FRETEIRO"]
                    
                    if not frotistas.empty:
                        buf = io.BytesIO()
                        frotistas.to_excel(buf, index=False)
                        zf.writestr(f"{nome_safe}_DISCADOR_MANHA_Frotista.xlsx", buf.getvalue())
                    
                    if not freteiros.empty:
                        buf = io.BytesIO()
                        freteiros.to_excel(buf, index=False)
                        zf.writestr(f"{nome_safe}_DISCADOR_ALMOCO_Freteiro.xlsx", buf.getvalue())

        # --- 2. ARQUIVOS DE MAILING (Apenas se modo Manhã) ---
        if modo == "MANHA" and df_mailing is not None and not df_mailing.empty:
            agentes_m = df_mailing[col_resp_m].unique()
            for agente in agentes_m:
                nome_safe = normalizar_nome_arquivo(agente)
                df_agente_m = df_mailing[df_mailing[col_resp_m] == agente]
                
                if not df_agente_m.empty:
                    buf = io.BytesIO()
                    df_agente_m.to_excel(buf, index=False)
                    zf.writestr(f"{nome_safe}_MAILING_DISTRIBUIDO.xlsx", buf.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# ==============================================================================
# 5. FRONTEND CENTRALIZADO
# ==============================================================================

st.title("🚛 Michelin Pilot Command Center")
st.markdown("---")

# --- BARRA LATERAL ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    st.header("🎮 Controle de Missão")
    modo_operacao = st.radio("Selecione o Turno:", ("🌅 Manhã (Carga Inicial)", "☀️ Tarde (Reprocessamento)"))
    
    st.markdown("---")
    st.subheader("1. Arquivo Mestre")
    uploaded_file = st.file_uploader("Carregue o Excel Principal (.xlsx)", type=["xlsx"], key="main")
    
    uploaded_log = None
    if modo_operacao == "☀️ Tarde (Reprocessamento)":
        st.markdown("---")
        st.subheader("2. Log do Discador")
        st.info("Necessário para remover quem já foi atendido.")
        uploaded_log = st.file_uploader("Solte o Log aqui (.csv/.xlsx)", type=["csv", "xlsx"], key="log")

# --- LÓGICA PRINCIPAL ---
if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
    
    # Identificação Automática de Abas
    aba_d = next((s for s in sheets if 'DISCADOR' in s.upper()), None)
    aba_m = next((s for s in sheets if 'MAILING' in s.upper()), None)
    
    if not aba_d or not aba_m:
        st.error("❌ ERRO: O arquivo precisa ter abas nomeadas como 'Discador' e 'Mailing'.")
        st.stop()
        
    df_d = pd.read_excel(uploaded_file, sheet_name=aba_d)
    df_m = pd.read_excel(uploaded_file, sheet_name=aba_m)
    
    # --- AUDITORIA DE INTEGRIDADE (MATCH) ---
    col_kpi1, col_kpi2 = st.columns([3, 1])
    diff = len(df_d) - len(df_m)
    if diff == 0:
        col_kpi1.markdown(f'<div class="audit-box-success">✅ <b>AUDITORIA OK:</b> Discador e Mailing sincronizados com {len(df_d)} registros.</div>', unsafe_allow_html=True)
    else:
        col_kpi1.markdown(f'<div class="audit-box-danger">🚨 <b>ALERTA:</b> Discador tem {len(df_d)} linhas e Mailing tem {len(df_m)}. Diferença de {abs(diff)}.</div>', unsafe_allow_html=True)

    # --- MAPEAMENTO DE COLUNAS ---
    cols = df_d.columns.tolist()
    sug_resp = next((c for c in cols if 'RESP' in c.upper()), cols[0])
    sug_doc = next((c for c in cols if 'CPF' in c.upper() or 'CNPJ' in c.upper()), cols[0])
    sug_prio = next((c for c in cols if 'PRIOR' in c.upper() or 'AGING' in c.upper()), None)
    sug_id = [c for c in cols if 'ID' in c.upper()][:1]
    sug_tel = [c for c in cols if 'TEL' in c.upper() or 'CEL' in c.upper()]

    with st.expander("⚙️ Configurações de Colunas e Distribuição", expanded=True):
        c1, c2 = st.columns(2)
        sel_resp = c1.selectbox("Responsável (Atendente):", cols, index=cols.index(sug_resp))
        sel_doc = c1.selectbox("CPF/CNPJ:", cols, index=cols.index(sug_doc))
        sel_prio = c1.selectbox("Prioridade (Opcional):", ["(Sem Prioridade)"] + cols, index=cols.index(sug_prio)+1 if sug_prio else 0)
        if sel_prio == "(Sem Prioridade)": sel_prio = None
        
        sel_id = c2.multiselect("ID Cliente (Chave Única):", cols, default=sug_id)
        sel_tel = c2.multiselect("Telefones (Discador):", cols, default=sug_tel)
        
        aplicar_bal = st.checkbox("Distribuir leads 'sem dono'?", value=True)

    # ==========================================
    # MODO MANHÃ
    # ==========================================
    if modo_operacao == "🌅 Manhã (Carga Inicial)":
        if st.button("🚀 EXECUTAR PROCESSO MANHÃ"):
            with st.spinner("1/3 Balanceando Carga..."):
                # 1. Distribui no DISCADOR
                df_d_bal = df_d.copy()
                if aplicar_bal:
                    df_d_bal, msg_dist = distribuir_carga_equitativa(df_d_bal, sel_resp, sel_prio)
                    st.info(msg_dist)
                
                # 2. Sincroniza decisão com MAILING (Copia quem é dono de quem)
                mapa_donos = dict(zip(df_d_bal[sel_id[0]], df_d_bal[sel_resp]))
                df_m_bal = df_m.copy()
                # Atualiza Mailing apenas onde houve mudança ou estava vazio
                df_m_bal[sel_resp] = df_m_bal[sel_id[0]].map(mapa_donos).fillna(df_m_bal[sel_resp])
                
            with st.spinner("2/3 Processando Arquivos..."):
                # Verticaliza Discador
                df_d_final = processar_verticalizacao_discador(df_d_bal, sel_id, sel_tel, sel_resp, sel_doc, sel_prio)
                # Classifica Mailing
                df_m_final = processar_mailing_simples(df_m_bal, sel_doc)
            
            # 3. Conferência e Download
            st.markdown('<div class="preview-box">', unsafe_allow_html=True)
            st.subheader("🔎 Painel de Conferência")
            
            tab_conf1, tab_conf2 = st.tabs(["📊 Carga Horária", "💎 Qualidade (Prioridades)"])
            
            with tab_conf1:
                # Mostra Frotista vs Freteiro por atendente
                try:
                    resumo_perfil = df_d_final.groupby([sel_resp, 'ESTRATEGIA_PERFIL'])[sel_id[0]].nunique().unstack(fill_value=0)
                    st.write("**Clientes Únicos por Atendente e Perfil:**")
                    st.dataframe(resumo_perfil, use_container_width=True)
                except: st.write("Erro ao gerar tabela de perfil.")

            with tab_conf2:
                # Mostra Prioridades por atendente
                if sel_prio:
                    try:
                        # Conta IDs únicos para não duplicar telefones
                        resumo_prio = df_d_bal.groupby([sel_resp, sel_prio])[sel_id[0]].nunique().unstack(fill_value=0)
                        st.write(f"**Distribuição da Coluna '{sel_prio}':**")
                        st.dataframe(resumo_prio, use_container_width=True)
                    except: st.write("Erro ao gerar tabela de prioridade.")
                else:
                    st.info("Nenhuma prioridade selecionada.")
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Gera ZIP Único
            zip_final = gerar_zip_arquivos(df_d_final, df_m_final, sel_resp, sel_resp, "MANHA")
            
            st.success("Processo Finalizado!")
            st.download_button("📥 DOWNLOAD PACOTE COMPLETO (.ZIP)", zip_final, "Pack_Michelin_Manha.zip", "application/zip", type="primary")

    # ==========================================
    # MODO TARDE
    # ==========================================
    elif modo_operacao == "☀️ Tarde (Reprocessamento)":
        if not uploaded_log:
            st.warning("Aguardando upload do Log do Discador...")
        else:
            try:
                if uploaded_log.name.endswith('.csv'): df_log = pd.read_csv(uploaded_log, sep=None, engine='python')
                else: df_log = pd.read_excel(uploaded_log)
                
                col_id_log = st.selectbox("Selecione a coluna de ID no Log:", df_log.columns)
                
                if st.button("🔄 GERAR REFORÇO TARDE"):
                    with st.spinner("Cruzando dados..."):
                        # Reaplica distribuição para consistência
                        df_d_bal = df_d.copy()
                        if aplicar_bal: df_d_bal, _ = distribuir_carga_equitativa(df_d_bal, sel_resp, sel_prio)
                        
                        # Filtra IDs já trabalhados
                        df_tarde_limpa, qtd_removida = processar_feedback_tarde(
                            df_d_bal, df_log, sel_id, col_id_log, sel_resp, sel_doc
                        )
                        
                        # Verticaliza
                        df_d_final = processar_verticalizacao_discador(df_tarde_limpa, sel_id, sel_tel, sel_resp, sel_doc, sel_prio)
                        
                        st.success(f"Log processado: {qtd_removida} clientes removidos.")
                        
                        # Preview
                        st.write("Amostra da Tarde (Apenas Frotistas Pendentes):")
                        st.dataframe(df_d_final.head())
                        
                        # Zip
                        zip_tarde = gerar_zip_arquivos(df_d_final, None, sel_resp, None, "TARDE")
                        st.download_button("📥 DOWNLOAD REFORÇO TARDE (.ZIP)", zip_tarde, "Pack_Michelin_Tarde.zip", "application/zip", type="primary")

            except Exception as e:
                st.error(f"Erro ao ler Log: {e}")

else:
    st.info("👈 Comece carregando o arquivo Mestre na barra lateral.")
