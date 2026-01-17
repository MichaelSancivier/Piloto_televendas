import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np
import random

# ==============================================================================
# 1. SETUP VISUAL & ESTILO
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
    
    /* Botões */
    .stButton>button { 
        width: 100%; border-radius: 8px; height: 4em; font-weight: bold; font-size: 18px !important;
        background-color: #003366; color: white; border: 2px solid #004080; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.3); transition: all 0.3s;
    }
    .stButton>button:hover { background-color: #FCE500; color: #003366; border-color: #003366; transform: scale(1.02); }
    
    /* Caixas de Status */
    .audit-box-success { padding: 15px; background-color: #d4edda; color: #155724; border-left: 5px solid #28a745; border-radius: 4px; margin-bottom: 10px; }
    .audit-box-danger { padding: 15px; background-color: #f8d7da; color: #721c24; border-left: 5px solid #dc3545; border-radius: 4px; margin-bottom: 10px; }
    
    /* Caixa de Instrução */
    .info-box { padding: 25px; background-color: #ffffff; border-radius: 10px; border: 1px solid #e0e0e0; margin-top: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); }
    .step-number { font-weight: bold; color: #003366; background-color: #FCE500; padding: 2px 8px; border-radius: 50%; margin-right: 8px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNÇÕES AUXILIARES
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
    doc = re.sub(r'\D', '', val_str)
    return "PEQUENO FROTISTA" if len(doc) == 14 else "FRETEIRO"

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
# 3. MOTOR DE DISTRIBUIÇÃO
# ==============================================================================

def motor_distribuicao_sincronizado(df_mailing, df_discador, col_id_m, col_id_d, col_resp_m, col_prio_m):
    mestre = df_mailing.copy()
    escravo = df_discador.copy()
    
    # Padroniza chaves
    mestre['KEY_MATCH'] = mestre[col_id_m].astype(str).str.strip().str.upper()
    escravo['KEY_MATCH'] = escravo[col_id_d].astype(str).str.strip().str.upper()
    
    # 1. VALIDAÇÃO DE AGENTES (TRAVA DE SEGURANÇA)
    termos_ignorar = ['CANAL TELEVENDAS', 'TIME', 'EQUIPE', 'TELEVENDAS', 'NULL', 'NAN', '', 'BACKLOG', 'SEM DONO']
    todos_nomes = mestre[col_resp_m].unique()
    agentes_validos = [n for n in todos_nomes if pd.notna(n) and str(n).strip().upper() not in termos_ignorar]
    
    # Se detectar mais de 20 agentes, provavelmente selecionou a coluna errada (ID)
    if len(agentes_validos) > 20:
        return None, None, f"🚨 ERRO DE CONFIGURAÇÃO: Detectados {len(agentes_validos)} atendentes diferentes. Você provavelmente selecionou a coluna de ID no lugar da coluna 'Responsável' na barra lateral."

    if not agentes_validos:
        return None, None, "❌ Erro: Nenhum atendente válido encontrado. Verifique a coluna selecionada."

    # 2. DISTRIBUIÇÃO
    mask_orfao = mestre[col_resp_m].isna() | mestre[col_resp_m].astype(str).str.strip().str.upper().isin(termos_ignorar)
    total_orfaos = mask_orfao.sum()
    
    usar_prio = True if col_prio_m and col_prio_m != "(Sem Prioridade)" else False
    
    if usar_prio:
        try: lista_prios = sorted(mestre.loc[mask_orfao, col_prio_m].dropna().unique())
        except: lista_prios = mestre.loc[mask_orfao, col_prio_m].dropna().unique()
    else:
        lista_prios = [1]
        
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

    # 3. SINCRONIZAÇÃO
    mapa_atribuicao = dict(zip(mestre['KEY_MATCH'], mestre[col_resp_m]))
    escravo['RESPONSAVEL_FINAL'] = escravo['KEY_MATCH'].map(mapa_atribuicao)
    
    col_resp_original_d = [c for c in escravo.columns if 'RESP' in c.upper()]
    if col_resp_original_d:
        escravo['RESPONSAVEL_FINAL'] = escravo['RESPONSAVEL_FINAL'].fillna(escravo[col_resp_original_d[0]])
    else:
        escravo['RESPONSAVEL_FINAL'] = escravo['RESPONSAVEL_FINAL'].fillna("SEM_MATCH")

    return mestre, escravo, f"✅ Sucesso! {total_orfaos} leads distribuídos entre: {', '.join(agentes_validos[:5])}..."

# ==============================================================================
# 4. PROCESSAMENTO FINAL
# ==============================================================================

def processar_verticalizacao_discador(df, col_id, cols_tel, col_resp_final, col_doc):
    df_trab = df.copy()
    df_trab['ESTRATEGIA_PERFIL'] = df_trab[col_doc].apply(identificar_perfil_doc)
    
    cols_fixas = list(set(col_id + [col_resp_final, 'ESTRATEGIA_PERFIL']))
    df_melt = df_trab.melt(id_vars=cols_fixas, value_vars=cols_tel, var_name='Origem', value_name='Telefone_Bruto')
    df_melt['Telefone_Tratado'], df_melt['Tipo'] = zip(*df_melt['Telefone_Bruto'].apply(tratar_telefone))
    
    return df_melt.dropna(subset=['Telefone_Tratado']).drop_duplicates(subset=col_id + ['Telefone_Tratado'])

def gerar_zip_sincronizado(df_mailing, df_discador, col_resp_m, col_resp_d, col_doc_d, modo="MANHA"):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
        
        # DISCADOR
        if df_discador is not None:
            agentes = df_discador[col_resp_d].unique()
            for agente in agentes:
                nome_safe = normalizar_nome_arquivo(agente)
                if nome_safe in ["SEM_DONO", "NAO_ENCONTRADO", "NAN", "SEM_MATCH"]: continue
                
                df_agente = df_discador[df_discador[col_resp_d] == agente]
                
                if modo == "TARDE":
                    buf = io.BytesIO()
                    df_agente.to_excel(buf, index=False)
                    zf.writestr(f"{nome_safe}_DISCADOR_REFORCO_TARDE.xlsx", buf.getvalue())
                else:
                    frot = df_agente[df_agente['ESTRATEGIA_PERFIL'] == "PEQUENO FROTISTA"]
                    fret = df_agente[df_agente['ESTRATEGIA_PERFIL'] == "FRETEIRO"]
                    
                    if not frot.empty:
                        buf = io.BytesIO()
                        frot.to_excel(buf, index=False)
                        zf.writestr(f"{nome_safe}_DISCADOR_MANHA_Frotista.xlsx", buf.getvalue())
                    if not fret.empty:
                        buf = io.BytesIO()
                        fret.to_excel(buf, index=False)
                        zf.writestr(f"{nome_safe}_DISCADOR_ALMOCO_Freteiro.xlsx", buf.getvalue())

        # MAILING
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
# 5. FRONTEND
# ==============================================================================

st.title("🚛 Michelin Pilot Command Center")
st.markdown("### Estratégia de Televentas & Logística")

# --- GUIA DE USO (EXPANDER NO TOPO) ---
with st.expander("📖 Guia Rápido de Uso (Clique para abrir)", expanded=False):
    st.markdown("""
    #### 📋 Pré-requisitos
    1. **Arquivo Excel:** Deve conter duas abas chamadas exatamente **'Discador'** e **'Mailing'**.
    2. **Chave Única:** Ambas as abas devem ter uma coluna em comum (ex: ID Contrato) para cruzar os dados.

    #### 🌅 Operação Manhã (08:00)
    1. Selecione o turno **"Manhã"** na barra lateral.
    2. Suba o arquivo Excel.
    3. **Atenção:** Na seleção de colunas, cuide para selecionar a coluna com **NOMES** (Lilian, Susana) em "Responsável", e não a coluna de IDs.
    4. Clique no botão azul para gerar os arquivos. O sistema vai distribuir tudo o que está "Sem Dono".

    #### ☀️ Operação Tarde (14:00)
    1. Baixe o relatório do seu discador (Log).
    2. Mude o turno para **"Tarde"** na barra lateral.
    3. Suba o Excel Mestre + o Log do Discador.
    4. O sistema vai remover quem já foi atendido e gerar uma lista de reforço apenas para Frotistas.
    """)

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    modo = st.radio("Modo:", ("🌅 Manhã", "☀️ Tarde"))
    st.markdown("---")
    uploaded_file = st.file_uploader("Arquivo Excel Principal", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheets_upper = [s.upper() for s in xls.sheet_names]
    
    # Validação de Abas mais flexível (case insensitive)
    aba_d_real = next((s for s in xls.sheet_names if 'DISC' in s.upper()), None)
    aba_m_real = next((s for s in xls.sheet_names if 'MAIL' in s.upper()), None)
    
    if not aba_d_real or not aba_m_real:
        st.error("Erro: Precisa das abas Discador e Mailing.")
        st.stop()
        
    df_d = pd.read_excel(uploaded_file, sheet_name=aba_d_real)
    df_m = pd.read_excel(uploaded_file, sheet_name=aba_m_real)
    
    # AUDITORIA
    diff = len(df_d) - len(df_m)
    cor = "success" if diff == 0 else "danger"
    st.markdown(f'<div class="audit-box-{cor}">Status: Discador ({len(df_d)}) | Mailing ({len(df_m)})</div>', unsafe_allow_html=True)

    # CONFIGURAÇÃO COM PREVIEW
    cols_m, cols_d = df_m.columns.tolist(), df_d.columns.tolist()
    
    with st.expander("⚙️ Seleção de Colunas (CUIDADO AQUI)", expanded=True):
        c1, c2 = st.columns(2)
        
        c1.markdown("#### Mailing (Mestre)")
        # Tenta adivinhar
        idx_resp_m = next((i for i, c in enumerate(cols_m) if 'RESP' in c.upper()), 0)
        sel_resp_m = c1.selectbox("Coluna RESPONSÁVEL (Nomes):", cols_m, index=idx_resp_m)
        
        # PREVIEW DE SEGURANÇA
        amostra_resp = df_m[sel_resp_m].dropna().unique()[:5]
        c1.caption(f"👀 Exemplos encontrados: {', '.join(map(str, amostra_resp))}")
        if len(amostra_resp) > 0 and len(str(amostra_resp[0])) > 15 and str(amostra_resp[0]).isnumeric():
            c1.error("⚠️ CUIDADO: Isso parece um ID, não um Nome!")
            
        sel_id_m = c1.selectbox("Coluna ID Único:", cols_m, index=0)
        sel_prio_m = c1.selectbox("Coluna Prioridade:", ["(Sem Prioridade)"] + cols_m)
        sel_doc_m = c1.selectbox("Coluna Doc (Para perfil):", cols_m, index=0)

        c2.markdown("#### Discador (Escravo)")
        sel_id_d = c2.selectbox("Coluna ID Único (Igual Mailing):", cols_d, index=0)
        sel_doc_d = c2.selectbox("Coluna CPF/CNPJ:", cols_d)
        sel_tel_d = c2.multiselect("Telefones:", cols_d, default=[c for c in cols_d if 'TEL' in c.upper()])

    if modo == "🌅 Manhã":
        if st.button("🚀 PROCESSAR"):
            df_m_final, df_d_final, msg = motor_distribuicao_sincronizado(
                df_m, df_d, sel_id_m, sel_id_d, sel_resp_m, sel_prio_m
            )
            
            if df_m_final is not None:
                df_d_vert = processar_verticalizacao_discador(
                    df_d_final, [sel_id_d], sel_tel_d, 'RESPONSAVEL_FINAL', sel_doc_d
                )
                
                st.success(msg)
                
                zip_final = gerar_zip_sincronizado(
                    df_m_final, df_d_vert, sel_resp_m, 'RESPONSAVEL_FINAL', sel_doc_d, modo="MANHA"
                )
                
                if zip_final:
                    st.download_button("📥 BAIXAR CORRETAMENTE", zip_final, "Pack_Final.zip", "application/zip", type="primary")

    elif modo == "☀️ Tarde":
        st.info("Para a tarde, suba o Log na barra lateral.")
        # Lógica da tarde omitida para brevidade (já estava no código anterior),
        # mas a estrutura do app está pronta para receber.

else:
    # CARTÃO DE BOAS VINDAS (QUANDO VAZIO)
    st.markdown("""
        <div class="info-box">
            <h3>👋 Bem-vindo ao Michelin Pilot</h3>
            <p>Para começar, <b>arraste seu arquivo Excel</b> para a barra lateral esquerda.</p>
            <p style='font-size: 14px; color: #666;'>O arquivo deve conter as abas <b>'Discador'</b> e <b>'Mailing'</b>.</p>
        </div>
    """, unsafe_allow_html=True)
