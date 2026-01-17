import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np
import random

# ==============================================================================
# 1. SETUP VISUAL & CONFIGURAÇÕES GERAIS
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
    h1 { color: #003366; font-family: 'Helvetica', sans-serif; font-weight: bold; }
    h2, h3 { color: #2c3e50; }
    .stButton>button { 
        width: 100%; border-radius: 8px; height: 3.5em; font-weight: bold; 
        background-color: #003366; color: white; border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stButton>button:hover { background-color: #004080; color: white; transform: translateY(-2px); transition: all 0.2s; }
    
    /* Status Boxes */
    .audit-box-success { 
        padding: 15px; background-color: #d4edda; color: #155724; 
        border-left: 5px solid #28a745; border-radius: 4px; margin-bottom: 15px; font-weight: 500;
    }
    .audit-box-danger { 
        padding: 15px; background-color: #f8d7da; color: #721c24; 
        border-left: 5px solid #dc3545; border-radius: 4px; margin-bottom: 15px; font-weight: 500;
    }
    .preview-box { 
        background-color: #ffffff; padding: 25px; border-radius: 12px; 
        border: 1px solid #e0e0e0; margin-top: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); 
    }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNÇÕES AUXILIARES (LIMPEZA E REGRAS DE NEGÓCIO)
# ==============================================================================

def normalizar_nome_arquivo(nome):
    """Remove caracteres especiais para criar nomes de arquivos seguros."""
    if pd.isna(nome): return "SEM_DONO"
    nfkd = unicodedata.normalize('NFKD', str(nome))
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    limpo = re.sub(r'[^a-zA-Z0-9]', '_', sem_acento.upper())
    return re.sub(r'_+', '_', limpo).strip('_')

def identificar_perfil_doc(valor):
    """
    REGRA BINÁRIA:
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
    """Padroniza telefones para +55... e identifica tipo (Celular/Fixo)."""
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
# 3. MOTOR DE DISTRIBUIÇÃO INTELIGENTE (O CÉREBRO)
# ==============================================================================

def motor_distribuicao_sincronizado(df_mailing, df_discador, col_id_m, col_id_d, col_resp_m, col_prio_m):
    """
    1. Usa o Mailing como MESTRE para decidir a distribuição (baseado em Prioridade).
    2. Aplica a distribuição no Mailing.
    3. Replica essa distribuição para o Discador usando o ID como chave.
    """
    # Trabalha com cópias para não afetar o original
    mestre = df_mailing.copy()
    escravo = df_discador.copy()
    
    # Padroniza chaves para garantir o match (texto limpo)
    mestre['KEY_MATCH'] = mestre[col_id_m].astype(str).str.strip().str.upper()
    escravo['KEY_MATCH'] = escravo[col_id_d].astype(str).str.strip().str.upper()
    
    # 1. IDENTIFICAÇÃO DE AGENTES VÁLIDOS (NO MAILING)
    termos_ignorar = ['CANAL TELEVENDAS', 'TIME', 'EQUIPE', 'TELEVENDAS', 'NULL', 'NAN', '', 'BACKLOG', 'SEM DONO']
    todos_nomes = mestre[col_resp_m].unique()
    agentes_validos = [n for n in todos_nomes if pd.notna(n) and str(n).strip().upper() not in termos_ignorar]
    
    if not agentes_validos:
        return None, None, "❌ Erro: Nenhum atendente válido encontrado na coluna Responsável do Mailing."

    # 2. IDENTIFICAÇÃO DE ÓRFÃOS (LEADS SEM DONO NO MAILING)
    mask_orfao = mestre[col_resp_m].isna() | mestre[col_resp_m].astype(str).str.strip().str.upper().isin(termos_ignorar)
    total_orfaos = mask_orfao.sum()
    
    # 3. DISTRIBUIÇÃO JUSTA POR PRIORIDADE
    # Se o usuário selecionou uma coluna de prioridade, usamos ela. Se não, usamos uma dummy.
    usar_prio = True if col_prio_m and col_prio_m != "(Sem Prioridade)" else False
    
    if usar_prio:
        # Pega as prioridades existentes (ex: 1, 2, 3) e ordena
        try:
            lista_prios = sorted(mestre.loc[mask_orfao, col_prio_m].dropna().unique())
        except:
            lista_prios = mestre.loc[mask_orfao, col_prio_m].dropna().unique()
    else:
        lista_prios = [1] # Dummy
        
    for p in lista_prios:
        # Filtra os órfãos desta camada de prioridade
        if usar_prio:
            mask_camada = mask_orfao & (mestre[col_prio_m] == p)
        else:
            mask_camada = mask_orfao
            
        indices = mestre[mask_camada].index.tolist()
        qtd = len(indices)
        
        if qtd > 0:
            # EMBARALHAMENTO DUPLO (JUSTIÇA TOTAL)
            # 1. Embaralha os clientes para não pegar vício da ordem do Excel
            random.shuffle(indices)
            
            # 2. Embaralha a equipe a cada rodada. 
            # Se sobrar 1 lead, ele cai para alguém aleatório, não sempre para o "A" da lista.
            equipe_rodada = agentes_validos.copy()
            random.shuffle(equipe_rodada)
            
            atribuicoes = np.resize(equipe_rodada, qtd)
            mestre.loc[indices, col_resp_m] = atribuicoes

    # 4. SINCRONIZAÇÃO (APLICAÇÃO NO DISCADOR)
    # Cria um mapa {ID_CLIENTE -> NOVO_DONO} a partir do Mailing processado
    mapa_atribuicao = dict(zip(mestre['KEY_MATCH'], mestre[col_resp_m]))
    
    # Aplica no Discador
    # Cria coluna nova 'RESPONSAVEL_FINAL' para não perder histórico se quiser auditar
    escravo['RESPONSAVEL_FINAL'] = escravo['KEY_MATCH'].map(mapa_atribuicao)
    
    # Onde não houve match (cliente só existe no discador?), mantém o original ou marca erro
    # Aqui vamos tentar manter o original se não tiver match, ou marcar "SEM_MATCH_MAILING"
    col_resp_original_d = [c for c in escravo.columns if 'RESP' in c.upper()][0] # Tenta achar a original
    escravo['RESPONSAVEL_FINAL'] = escravo['RESPONSAVEL_FINAL'].fillna(escravo[col_resp_original_d])
    
    return mestre, escravo, f"✅ Sincronização Concluída! {total_orfaos} leads distribuídos via Mailing e aplicados no Discador."

# ==============================================================================
# 4. PROCESSAMENTO FINAL E ARQUIVOS
# ==============================================================================

def processar_verticalizacao_discador(df, col_id, cols_tel, col_resp_final, col_doc):
    """Prepara o arquivo do discador: Verticaliza telefones e limpa."""
    df_trab = df.copy()
    
    # Define Perfil (Frotista/Freteiro)
    df_trab['ESTRATEGIA_PERFIL'] = df_trab[col_doc].apply(identificar_perfil_doc)
    
    # Colunas para manter
    cols_fixas = list(set(col_id + [col_resp_final, 'ESTRATEGIA_PERFIL']))
    
    # Melt (Verticalizar)
    df_melt = df_trab.melt(id_vars=cols_fixas, value_vars=cols_tel, var_name='Origem', value_name='Telefone_Bruto')
    df_melt['Telefone_Tratado'], df_melt['Tipo'] = zip(*df_melt['Telefone_Bruto'].apply(tratar_telefone))
    
    # Limpa
    df_final = df_melt.dropna(subset=['Telefone_Tratado']).drop_duplicates(subset=col_id + ['Telefone_Tratado'])
    return df_final

def gerar_zip_sincronizado(df_mailing, df_discador, col_resp_m, col_resp_d, col_doc_d, modo="MANHA"):
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
        
        # --- 1. ARQUIVOS DO DISCADOR (SEPARADOS POR HORÁRIO) ---
        if df_discador is not None:
            agentes = df_discador[col_resp_d].unique()
            for agente in agentes:
                nome_safe = normalizar_nome_arquivo(agente)
                # Ignora nomes inválidos que possam ter sobrado
                if nome_safe in ["SEM_DONO", "NAO_ENCONTRADO", "NAN"]: continue
                
                df_agente = df_discador[df_discador[col_resp_d] == agente]
                
                if modo == "TARDE":
                    # Reforço Tarde (Arquivo Único)
                    buf = io.BytesIO()
                    df_agente.to_excel(buf, index=False)
                    zf.writestr(f"{nome_safe}_DISCADOR_REFORCO_TARDE.xlsx", buf.getvalue())
                else:
                    # Manhã (Frotista) vs Almoço (Freteiro)
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

        # --- 2. ARQUIVOS DO MAILING (DISTRIBUÍDO) ---
        # Só gera mailing na carga inicial (Manhã)
        if modo == "MANHA" and df_mailing is not None:
            agentes_m = df_mailing[col_resp_m].unique()
            for agente in agentes_m:
                nome_safe = normalizar_nome_arquivo(agente)
                if nome_safe in ["SEM_DONO", "NAO_ENCONTRADO", "NAN"]: continue
                
                df_agente_m = df_mailing[df_mailing[col_resp_m] == agente]
                
                if not df_agente_m.empty:
                    buf = io.BytesIO()
                    df_agente_m.to_excel(buf, index=False)
                    zf.writestr(f"{nome_safe}_MAILING_DISTRIBUIDO.xlsx", buf.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# ==============================================================================
# 5. INTERFACE DO USUÁRIO (FRONTEND)
# ==============================================================================

st.title("🚛 Michelin Pilot Command Center V31")
st.markdown("### 🎯 Sincronização Inteligente & Distribuição Justa")
st.markdown("---")

# BARRA LATERAL
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    st.header("🎮 Painel de Controle")
    modo_operacao = st.radio("Selecione a Operação:", ("🌅 Manhã (Carga & Distribuição)", "☀️ Tarde (Reprocessamento)"))
    
    st.markdown("---")
    st.subheader("1. Arquivo Mestre")
    uploaded_file = st.file_uploader("Carregue o Excel (Com abas Discador/Mailing)", type=["xlsx"])
    
    uploaded_log = None
    if modo_operacao == "☀️ Tarde (Reprocessamento)":
        st.markdown("---")
        st.subheader("2. Log do Discador")
        st.info("Para remover quem já foi atendido.")
        uploaded_log = st.file_uploader("Carregue o Log (.csv/.xlsx)", type=["csv", "xlsx"])

# LÓGICA PRINCIPAL
if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
    
    # 1. VERIFICAÇÃO DE ABAS
    aba_d = next((s for s in sheets if 'DISC' in s.upper()), None)
    aba_m = next((s for s in sheets if 'MAIL' in s.upper()), None)
    
    if not aba_d or not aba_m:
        st.error("❌ ERRO CRÍTICO: O arquivo Excel precisa ter as abas 'Discador' e 'Mailing'.")
        st.stop()
        
    df_d = pd.read_excel(uploaded_file, sheet_name=aba_d)
    df_m = pd.read_excel(uploaded_file, sheet_name=aba_m)
    
    # 2. AUDITORIA DE QUANTIDADE
    col_kpi1, col_kpi2 = st.columns([3, 1])
    diff = len(df_d) - len(df_m)
    
    if diff == 0:
        col_kpi1.markdown(f'<div class="audit-box-success">✅ <b>AUDITORIA OK:</b> Bases sincronizadas ({len(df_d)} registros).</div>', unsafe_allow_html=True)
    else:
        col_kpi1.markdown(f'<div class="audit-box-danger">⚠️ <b>ALERTA:</b> Discador ({len(df_d)}) e Mailing ({len(df_m)}) têm tamanhos diferentes. A sincronia será feita pelo ID.</div>', unsafe_allow_html=True)

    # 3. CONFIGURAÇÃO DE COLUNAS
    cols_m = df_m.columns.tolist()
    cols_d = df_d.columns.tolist()
    
    # Tentativa de auto-seleção
    sug_id_m = next((c for c in cols_m if 'ID' in c.upper() or 'CONTRATO' in c.upper()), cols_m[0])
    sug_resp_m = next((c for c in cols_m if 'RESP' in c.upper()), cols_m[0])
    sug_doc_m = next((c for c in cols_m if 'CPF' in c.upper() or 'DOC' in c.upper() or 'CNPJ' in c.upper()), cols_m[0])
    sug_prio_m = next((c for c in cols_m if 'PRIOR' in c.upper() or 'AGING' in c.upper()), None)
    
    sug_id_d = next((c for c in cols_d if 'ID' in c.upper() or 'CONTRATO' in c.upper()), cols_d[0])
    sug_doc_d = next((c for c in cols_d if 'CPF' in c.upper() or 'DOC' in c.upper() or 'CNPJ' in c.upper()), cols_d[0])
    sug_tel_d = [c for c in cols_d if 'TEL' in c.upper() or 'CEL' in c.upper()]

    with st.expander("⚙️ Configurações de Mapeamento (Crucial)", expanded=True):
        c1, c2 = st.columns(2)
        
        c1.markdown("#### 1. Mailing (Mestre)")
        sel_id_m = c1.selectbox("Chave Única (ID):", cols_m, index=cols_m.index(sug_id_m), key='idm')
        sel_resp_m = c1.selectbox("Responsável:", cols_m, index=cols_m.index(sug_resp_m), key='respm')
        sel_prio_m = c1.selectbox("Prioridade (Aging):", ["(Sem Prioridade)"] + cols_m, index=cols_m.index(sug_prio_m)+1 if sug_prio_m else 0, key='priom')
        
        c2.markdown("#### 2. Discador (Escravo)")
        sel_id_d = c2.selectbox("Chave Única (Deve bater com Mailing):", cols_d, index=cols_d.index(sug_id_d), key='idd')
        sel_doc_d = c2.selectbox("CPF/CNPJ (Para separar Manhã/Almoço):", cols_d, index=cols_d.index(sug_doc_d), key='docd')
        sel_tel_d = c2.multiselect("Colunas de Telefone:", cols_d, default=sug_tel_d, key='teld')

    # ==========================================
    # LÓGICA: MANHÃ
    # ==========================================
    if modo_operacao == "🌅 Manhã (Carga Inicial)":
        if st.button("🚀 INICIAR DISTRIBUIÇÃO E SINCRONIZAÇÃO"):
            with st.spinner("Analisando prioridades e distribuindo carga..."):
                
                # EXECUTA O MOTOR UNIFICADO
                df_m_final, df_d_final, msg = motor_distribuicao_sincronizado(
                    df_m, df_d, sel_id_m, sel_id_d, sel_resp_m, sel_prio_m
                )
                
                if df_m_final is not None:
                    # PROCESSA O DISCADOR (VERTICALIZAÇÃO)
                    # Usa a coluna 'RESPONSAVEL_FINAL' gerada pelo motor
                    df_d_vert = processar_verticalizacao_discador(
                        df_d_final, [sel_id_d], sel_tel_d, 'RESPONSAVEL_FINAL', sel_doc_d
                    )
                    
                    st.success(msg)
                    
                    # --- PAINEL DE RESULTADOS ---
                    st.markdown('<div class="preview-box">', unsafe_allow_html=True)
                    st.subheader("📊 Relatório de Distribuição")
                    
                    tab1, tab2 = st.tabs(["⚖️ Balanço de Prioridades", "⏰ Balanço de Horário (Perfil)"])
                    
                    with tab1:
                        if sel_prio_m != "(Sem Prioridade)":
                            st.write(f"**Justiça na Distribuição ({sel_prio_m}):**")
                            # Agrupa Mailing final para ver como ficou a distribuição das prioridades
                            pivot_prio = df_m_final.groupby([sel_resp_m, sel_prio_m]).size().unstack(fill_value=0)
                            st.dataframe(pivot_prio, use_container_width=True)
                        else:
                            st.info("Distribuição aleatória (sem prioridade definida).")

                    with tab2:
                        st.write("**O que vai para o Discador (Frotista/Manhã vs Freteiro/Almoço):**")
                        # Conta IDs únicos no discador verticalizado
                        pivot_perfil = df_d_vert.groupby(['RESPONSAVEL_FINAL', 'ESTRATEGIA_PERFIL'])[sel_id_d].nunique().unstack(fill_value=0)
                        st.dataframe(pivot_perfil, use_container_width=True)
                        
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # GERAÇÃO DO ZIP
                    zip_final = gerar_zip_sincronizado(
                        df_m_final, df_d_vert, 
                        sel_resp_m, 'RESPONSAVEL_FINAL', 
                        sel_doc_d, modo="MANHA"
                    )
                    
                    st.download_button(
                        label="📥 BAIXAR PACOTE COMPLETO (.ZIP)",
                        data=zip_final,
                        file_name="Michelin_Distribuicao_Oficial.zip",
                        mime="application/zip",
                        type="primary"
                    )

    # ==========================================
    # LÓGICA: TARDE
    # ==========================================
    elif modo_operacao == "☀️ Tarde (Reprocessamento)":
        if not uploaded_log:
            st.warning("⚠️ Carregue o Log do Discador na barra lateral.")
        else:
            # Leitura do Log
            try:
                if uploaded_log.name.endswith('.csv'): df_log = pd.read_csv(uploaded_log, sep=None, engine='python')
                else: df_log = pd.read_excel(uploaded_log)
                
                col_id_log = st.selectbox("Selecione a coluna de ID no Log:", df_log.columns)
                
                if st.button("🔄 GERAR REFORÇO TARDE"):
                    with st.spinner("Recalculando base pendente..."):
                        # 1. Refaz distribuição (para garantir consistência)
                        _, df_d_final, _ = motor_distribuicao_sincronizado(
                            df_m, df_d, sel_id_m, sel_id_d, sel_resp_m, sel_prio_m
                        )
                        
                        # 2. Filtra quem já foi trabalhado (Log)
                        ids_trabalhados = df_log[col_id_log].astype(str).str.strip().unique()
                        df_d_final['KEY_TEMP'] = df_d_final[sel_id_d].astype(str).str.strip()
                        
                        df_pendente = df_d_final[~df_d_final['KEY_TEMP'].isin(ids_trabalhados)].copy()
                        qtd_removida = len(df_d_final) - len(df_pendente)
                        
                        # 3. Filtra apenas FROTISTAS (Regra da Tarde)
                        df_pendente['ESTRATEGIA'] = df_pendente[sel_doc_d].apply(identificar_perfil_doc)
                        df_tarde = df_pendente[df_pendente['ESTRATEGIA'] == 'PEQUENO FROTISTA']
                        
                        # 4. Verticaliza
                        df_tarde_vert = processar_verticalizacao_discador(
                            df_tarde, [sel_id_d], sel_tel_d, 'RESPONSAVEL_FINAL', sel_doc_d
                        )
                        
                        st.success(f"Log processado! {qtd_removida} clientes removidos.")
                        
                        # 5. Zip
                        zip_tarde = gerar_zip_sincronizado(
                            None, df_tarde_vert, 
                            None, 'RESPONSAVEL_FINAL', 
                            sel_doc_d, modo="TARDE"
                        )
                        
                        st.download_button(
                            "📥 BAIXAR REFORÇO TARDE (.ZIP)",
                            zip_tarde,
                            "Michelin_Reforco_Tarde.zip",
                            "application/zip",
                            type="primary"
                        )

            except Exception as e:
                st.error(f"Erro ao processar log: {e}")
