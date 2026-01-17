import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np
import random

# ==============================================================================
# 1. SETUP VISUAL
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
    .stButton>button { 
        width: 100%; border-radius: 8px; height: 4em; font-weight: bold; font-size: 20px !important;
        background-color: #003366; color: white; border: 2px solid #004080; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.3); transition: all 0.3s;
    }
    .stButton>button:hover { background-color: #FCE500; color: #003366; border-color: #003366; transform: scale(1.02); }
    
    .audit-box-success { padding: 15px; background-color: #d4edda; color: #155724; border-left: 5px solid #28a745; border-radius: 4px; margin-bottom: 15px; }
    .audit-box-danger { padding: 15px; background-color: #f8d7da; color: #721c24; border-left: 5px solid #dc3545; border-radius: 4px; margin-bottom: 15px; }
    .info-box { padding: 20px; background-color: #e3f2fd; border-radius: 10px; border: 1px dashed #2196f3; text-align: center; margin-top: 20px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNÇÕES LÓGICAS
# ==============================================================================

def normalizar_nome_arquivo(nome):
    if pd.isna(nome): return "SEM_DONO"
    nfkd = unicodedata.normalize('NFKD', str(nome))
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    limpo = re.sub(r'[^a-zA-Z0-9]', '_', sem_acento.upper())
    return re.sub(r'_+', '_', limpo).strip('_')

def identificar_perfil_doc(valor):
    if pd.isna(valor): return "FRETEIRO"
    val_str = str(valor)
    if val_str.endswith('.0'): val_str = val_str[:-2]
    doc_limpo = re.sub(r'\D', '', val_str)
    if len(doc_limpo) == 14: return "PEQUENO FROTISTA"
    else: return "FRETEIRO"

def tratar_telefone(val):
    if pd.isna(val): return None, "Vazio"
    nums = re.sub(r'\D', '', str(val))
    if nums.startswith('55') and len(nums) >= 12: nums = nums[2:]
    
    final = None
    tipo = "Inválido"
    if len(nums) == 11 and int(nums[2]) == 9:
        tipo = "Celular"
        final = nums
    elif len(nums) == 10 and int(nums[2]) >= 6:
        tipo = "Celular (Corrigido)"
        final = nums[:2] + '9' + nums[2:]
    elif len(nums) == 10 and 2 <= int(nums[2]) <= 5:
        tipo = "Fixo"
        final = nums
    return (f"+55{final}", tipo) if final else (None, tipo)

# ==============================================================================
# 3. MOTOR DE DISTRIBUIÇÃO (SINCRONIZADO)
# ==============================================================================

def motor_distribuicao_sincronizado(df_mailing, df_discador, col_id_m, col_id_d, col_resp_m, col_prio_m):
    mestre = df_mailing.copy()
    escravo = df_discador.copy()
    
    mestre['KEY_MATCH'] = mestre[col_id_m].astype(str).str.strip().str.upper()
    escravo['KEY_MATCH'] = escravo[col_id_d].astype(str).str.strip().str.upper()
    
    termos_ignorar = ['CANAL TELEVENDAS', 'TIME', 'EQUIPE', 'TELEVENDAS', 'NULL', 'NAN', '', 'BACKLOG', 'SEM DONO']
    todos_nomes = mestre[col_resp_m].unique()
    agentes_validos = [n for n in todos_nomes if pd.notna(n) and str(n).strip().upper() not in termos_ignorar]
    
    if not agentes_validos: return None, None, "❌ Erro: Nenhum atendente válido encontrado."

    mask_orfao = mestre[col_resp_m].isna() | mestre[col_resp_m].astype(str).str.strip().str.upper().isin(termos_ignorar)
    total_orfaos = mask_orfao.sum()
    
    usar_prio = True if col_prio_m and col_prio_m != "(Sem Prioridade)" else False
    if usar_prio:
        try: lista_prios = sorted(mestre.loc[mask_orfao, col_prio_m].dropna().unique())
        except: lista_prios = mestre.loc[mask_orfao, col_prio_m].dropna().unique()
    else: lista_prios = [1]
        
    for p in lista_prios:
        if usar_prio: mask_camada = mask_orfao & (mestre[col_prio_m] == p)
        else: mask_camada = mask_orfao
            
        indices = mestre[mask_camada].index.tolist()
        if indices:
            random.shuffle(indices)
            equipe_rodada = agentes_validos.copy()
            random.shuffle(equipe_rodada)
            atribuicoes = np.resize(equipe_rodada, len(indices))
            mestre.loc[indices, col_resp_m] = atribuicoes

    mapa_atribuicao = dict(zip(mestre['KEY_MATCH'], mestre[col_resp_m]))
    escravo['RESPONSAVEL_FINAL'] = escravo['KEY_MATCH'].map(mapa_atribuicao)
    col_resp_original_d = [c for c in escravo.columns if 'RESP' in c.upper()][0]
    escravo['RESPONSAVEL_FINAL'] = escravo['RESPONSAVEL_FINAL'].fillna(escravo[col_resp_original_d])
    
    return mestre, escravo, f"✅ Sincronização OK! {total_orfaos} leads distribuídos."

# ==============================================================================
# 4. PROCESSAMENTO FINAL
# ==============================================================================

def processar_verticalizacao_discador(df, col_id, cols_tel, col_resp_final, col_doc):
    df_trab = df.copy()
    df_trab['ESTRATEGIA_PERFIL'] = df_trab[col_doc].apply(identificar_perfil_doc)
    cols_fixas = list(set(col_id + [col_resp_final, 'ESTRATEGIA_PERFIL']))
    df_melt = df_trab.melt(id_vars=cols_fixas, value_vars=cols_tel, var_name='Origem', value_name='Telefone_Bruto')
    df_melt['Telefone_Tratado'], df_melt['Tipo'] = zip(*df_melt['Telefone_Bruto'].apply(tratar_telefone))
    df_final = df_melt.dropna(subset=['Telefone_Tratado']).drop_duplicates(subset=col_id + ['Telefone_Tratado'])
    return df_final

def gerar_zip_sincronizado(df_mailing, df_discador, col_resp_m, col_resp_d, col_doc_d, modo="MANHA"):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
        if df_discador is not None:
            agentes = df_discador[col_resp_d].unique()
            for agente in agentes:
                nome_safe = normalizar_nome_arquivo(agente)
                if nome_safe in ["SEM_DONO", "NAO_ENCONTRADO", "NAN"]: continue
                df_agente = df_discador[df_discador[col_resp_d] == agente]
                
                if modo == "TARDE":
                    buf = io.BytesIO()
                    df_agente.to_excel(buf, index=False)
                    zf.writestr(f"{nome_safe}_DISCADOR_REFORCO_TARDE.xlsx", buf.getvalue())
                else:
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
# 5. FRONTEND (VISUAL)
# ==============================================================================

st.title("🚛 Michelin Pilot Command Center")
st.markdown("### Estratégia de Televentas & Logística")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    st.header("1. Configuração")
    modo_operacao = st.radio("Turno:", ("🌅 Manhã (Distribuição)", "☀️ Tarde (Reforço)"))
    st.markdown("---")
    st.subheader("2. Upload Arquivo")
    uploaded_file = st.file_uploader("Arraste o Excel Mestre (.xlsx)", type=["xlsx"])

# Lógica de Exibição (Só roda se tiver arquivo)
if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
    
    # Validação de Abas
    aba_d = next((s for s in sheets if 'DISC' in s.upper()), None)
    aba_m = next((s for s in sheets if 'MAIL' in s.upper()), None)
    
    if not aba_d or not aba_m:
        st.error(f"❌ ERRO: Faltam abas! Encontradas: {sheets}. Precisa ter 'Discador' e 'Mailing'.")
        st.stop()
        
    df_d = pd.read_excel(uploaded_file, sheet_name=aba_d)
    df_m = pd.read_excel(uploaded_file, sheet_name=aba_m)
    
    # Auditoria Visual
    st.markdown(f"""
        <div class="audit-box-success">
            ✅ <b>ARQUIVO CARREGADO COM SUCESSO!</b><br>
            • Discador: {len(df_d)} registros<br>
            • Mailing: {len(df_m)} registros
        </div>
    """, unsafe_allow_html=True)
    
    # Seleção de Colunas
    cols_m, cols_d = df_m.columns.tolist(), df_d.columns.tolist()
    
    with st.expander("⚙️ Conferir Colunas (Importante)", expanded=True):
        c1, c2 = st.columns(2)
        c1.markdown("#### Mailing (Mestre)")
        sel_id_m = c1.selectbox("Chave (ID):", cols_m, index=0, key='idm')
        sel_resp_m = c1.selectbox("Responsável:", cols_m, index=min(1, len(cols_m)-1), key='respm')
        sel_prio_m = c1.selectbox("Prioridade (Aging):", ["(Sem Prioridade)"] + cols_m, key='priom')
        
        c2.markdown("#### Discador (Escravo)")
        sel_id_d = c2.selectbox("Chave (ID Igual):", cols_d, index=0, key='idd')
        sel_doc_d = c2.selectbox("CPF/CNPJ:", cols_d, index=min(1, len(cols_d)-1), key='docd')
        sel_tel_d = c2.multiselect("Telefones:", cols_d, default=[c for c in cols_d if 'TEL' in c.upper() or 'CEL' in c.upper()])

    # ==========================
    # BOTÃO GRANDE DE AÇÃO
    # ==========================
    st.markdown("---")
    
    if modo_operacao == "🌅 Manhã (Distribuição)":
        
        # AQUI ESTÁ O BOTÃO! SE O ARQUIVO CARREGOU, ELE APARECE.
        if st.button("🚀 EXECUTAR DISTRIBUIÇÃO E GERAR ZIP"):
            with st.spinner("Processando... O Robô está pensando..."):
                # 1. Distribui
                df_m_final, df_d_final, msg = motor_distribuicao_sincronizado(
                    df_m, df_d, sel_id_m, sel_id_d, sel_resp_m, sel_prio_m
                )
                
                # 2. Verticaliza
                if df_m_final is not None:
                    df_d_vert = processar_verticalizacao_discador(
                        df_d_final, [sel_id_d], sel_tel_d, 'RESPONSAVEL_FINAL', sel_doc_d
                    )
                    
                    st.success(msg)
                    
                    # 3. Mostra Resultados
                    st.subheader("📊 Resultados da Distribuição")
                    resumo = df_m_final.groupby([sel_resp_m, sel_prio_m]).size().unstack(fill_value=0) if sel_prio_m != "(Sem Prioridade)" else None
                    if resumo is not None: st.dataframe(resumo, use_container_width=True)
                    
                    # 4. Download
                    zip_final = gerar_zip_sincronizado(
                        df_m_final, df_d_vert, sel_resp_m, 'RESPONSAVEL_FINAL', sel_doc_d, modo="MANHA"
                    )
                    st.download_button("📥 BAIXAR TUDO AGORA", zip_final, "Pack_Michelin_Final.zip", "application/zip", type="primary")

    elif modo_operacao == "☀️ Tarde (Reforço)":
        st.info("Para a tarde, suba o LOG na barra lateral.")
        # Lógica da tarde aqui (abre quando subir o log)

else:
    # ESTADO VAZIO (QUANDO ABRE O APP)
    st.markdown("""
        <div class="info-box">
            <h3>👋 Olá! Bem-vindo ao Michelin Pilot</h3>
            <p>Para começar, arraste seu arquivo Excel para a barra lateral esquerda.</p>
        </div>
    """, unsafe_allow_html=True)
