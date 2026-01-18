import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np
import random

# ==============================================================================
# 1. CONFIGURAÇÃO DE ESTADO E MEMÓRIA (SESSION STATE)
# ==============================================================================
st.set_page_config(page_title="Michelin Command Center V55", page_icon="🚛", layout="wide")

if 'processado' not in st.session_state:
    st.session_state.update({
        'processado': False,
        'm_fin': None,
        'd_fin': None,
        'audit_cartera': None,
        'c_resp': "",
        'c_prio': "",
        'c_tels': [],
        'log_audit_tel': ""
    })

st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    .stButton>button { width: 100%; height: 3.8em; font-weight: bold; border-radius: 10px; background-color: #003366; color: white; border: 2px solid #FCE500; }
    .stButton>button:hover { background-color: #FCE500; color: #003366; }
    .auto-success { padding: 20px; background-color: #d4edda; color: #155724; border-left: 6px solid #28a745; border-radius: 5px; margin-bottom: 20px;}
    .auto-info { padding: 15px; background-color: #e7f3ff; color: #004085; border-left: 6px solid #007bff; border-radius: 5px; margin-bottom: 20px;}
    h3 { color: #003366; border-bottom: 3px solid #FCE500; padding-bottom: 8px; margin-top: 25px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNÇÕES DE TRATAMENTO DE DADOS (O MOTOR)
# ==============================================================================

def limpar_id_v55(valor):
    if pd.isna(valor): return ""
    s = str(valor).strip().split('.')[0]
    return re.sub(r'\D', '', s)

def tratar_telefone_v55(val):
    if pd.isna(val): return None
    s = str(val).strip()
    if not s or s in ["55", "+55"]: return None
    # Se já tem o formato internacional correto, mantém
    if s.startswith('+55') and len(re.sub(r'\D', '', s)) >= 12: return s
    nums = re.sub(r'\D', '', s)
    if nums.startswith('55') and len(nums) >= 12: nums = nums[2:]
    if len(nums) >= 10: return f"+55{nums}"
    return None

def obter_perfil_v55(v):
    d = re.sub(r'\D', '', str(v))
    return "PEQUENO FROTISTA" if len(d) == 14 else "FRETEIRO"

def converter_prio_v55(v):
    s = str(v).upper()
    if "BACKLOG" in s: return 99
    nums = re.findall(r'\d+', s)
    return int(nums[0]) if nums else 50

# ==============================================================================
# 3. GERAÇÃO DE FICHEIROS (OS 8 OUTPUTS)
# ==============================================================================

def gerar_zip_v55(df_m, df_d, col_resp, col_tels_d):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
        agentes = [a for a in df_m[col_resp].unique() if a not in ["SEM_MATCH", "SEM_DONO"]]
        
        for ag in agentes:
            # Nome limpo para o ficheiro
            nome_safe = re.sub(r'[^a-zA-Z0-9]', '_', unicodedata.normalize('NFKD', str(ag)).encode('ASCII', 'ignore').decode('ASCII').upper())
            
            # --- MAILING ---
            m_ag = df_m[df_m[col_resp] == ag]
            for perfil, turno in [("PEQUENO FROTISTA", "MANHA_Frotista"), ("FRETEIRO", "ALMOCO_Freteiro")]:
                sub_m = m_ag[m_ag['PERFIL_FINAL'] == perfil]
                excel_m = io.BytesIO()
                sub_m.to_excel(excel_m, index=False)
                zf.writestr(f"{nome_safe}_MAILING_{turno}.xlsx", excel_m.getvalue())
            
            # --- DISCADOR (VERTICALIZAÇÃO AQUI) ---
            d_ag = df_d[df_d['RESPONSAVEL_FINAL'] == ag]
            for perfil, turno in [("PEQUENO FROTISTA", "MANHA_Frotista"), ("FRETEIRO", "ALMOCO_Freteiro")]:
                sub_d = d_ag[d_ag['PERFIL_FINAL'] == perfil]
                
                # Identifica colunas fixas para não perder os dados do cliente
                cols_fixas = [c for c in sub_d.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL', 'PERFIL'])]
                
                # Transforma colunas de telefone em linhas (Verticalização)
                melted = sub_d.melt(id_vars=cols_fixas, value_vars=col_tels_d, value_name='Tel_Bruto')
                melted['Telefone'] = melted['Tel_Bruto'].apply(tratar_telefone_v55)
                
                # Limpa duplicados (não ligar 2x para o mesmo número do mesmo cliente)
                final_d = melted.dropna(subset=['Telefone']).drop_duplicates(subset=['Telefone'])
                
                excel_d = io.BytesIO()
                final_d.to_excel(excel_d, index=False)
                zf.writestr(f"{nome_safe}_DISCADOR_{turno}.xlsx", excel_d.getvalue())
                
    buf.seek(0)
    return buf

# ==============================================================================
# 4. INTERFACE E LOGICA DE PROCESSAMENTO (THE COMMAND CENTER)
# ==============================================================================

st.title("🚛 Michelin Pilot Command Center V55")

with st.sidebar:
    st.header("📂 Configuração")
    ficheiro = st.file_uploader("Submeter Excel Mestre", type="xlsx")
    if st.session_state.processado:
        st.markdown("---")
        if st.button("🗑️ REINICIAR PROCESSO"):
            st.session_state.processado = False
            st.rerun()

if ficheiro:
    if not st.session_state.processado:
        if st.button("🚀 EXECUTAR MOTOR DE DISTRIBUIÇÃO"):
            xls = pd.ExcelFile(ficheiro)
            aba_d = next(s for s in xls.sheet_names if 'DISC' in s.upper())
            aba_m = next(s for s in xls.sheet_names if 'MAIL' in s.upper())
            df_m, df_d = pd.read_excel(ficheiro, aba_m), pd.read_excel(ficheiro, aba_d)

            # --- DETEÇÃO AUTOMÁTICA DE COLUNAS ---
            c_id_m = next(c for c in df_m.columns if 'ID_CONTRATO' in c.upper())
            c_id_d = next(c for c in df_d.columns if 'ID_CONTRATO' in c.upper())
            c_resp_m = next(c for c in df_m.columns if 'RESP' in c.upper())
            c_doc_m = next(c for c in df_m.columns if 'CNPJ' in c.upper())
            c_prio_m = next(c for c in df_m.columns if 'PRIOR' in c.upper())
            c_tels_d = [c for c in df_d.columns if any(x in c.upper() for x in ['TEL', 'CEL', 'PRINCIPAL', 'COMERCIAL'])]

            # Auditoria inicial
            snap_antes = df_m[c_resp_m].value_counts().to_dict()

            # --- 1. LÓGICA DE MAILING ---
            df_m['PERFIL_FINAL'] = df_m[c_doc_m].apply(obter_perfil_v55)
            df_m['PRIO_SCORE'] = df_m[c_prio_m].apply(converter_prio_v55)
            
            # Balanceamento Quirúrgico (Distribuir para chegar a 47 cada)
            agentes = [a for a in df_m[c_resp_m].unique() if pd.notna(a) and "BACKLOG" not in str(a).upper()]
            mask_orfao = df_m[c_resp_m].isna() | (df_m[c_resp_m].astype(str).str.upper().str.contains("BACKLOG"))
            
            for idx in df_m[mask_orfao].index:
                cargas = df_m[c_resp_m].value_counts()
                agente_menor = min(agentes, key=lambda x: cargas.get(x, 0))
                df_m.at[idx, c_resp_m] = agente_menor

            # Nivelamento Final (Move registos de baixa prioridade se houver desigualdade)
            for _ in range(50):
                cargas = df_m[c_resp_m].value_counts()
                ag_max, ag_min = max(agentes, key=lambda x: cargas.get(x, 0)), min(agentes, key=lambda x: cargas.get(x, 0))
                if (cargas.get(ag_max, 0) - cargas.get(ag_min, 0)) <= 1: break
                idx_troca = df_m[df_m[c_resp_m] == ag_max].sort_values('PRIO_SCORE', ascending=False).index[0]
                df_m.at[idx_troca, c_resp_m] = ag_min

            # --- 2. SINCRONIZAÇÃO DISCADOR ---
            df_d['KEY'] = df_d[c_id_d].apply(limpar_id_v55)
            map_resp = dict(zip(df_m[c_id_m].apply(limpar_id_v55), df_m[c_resp_m]))
            map_perf = dict(zip(df_m[c_id_m].apply(limpar_id_v55), df_m['PERFIL_FINAL']))
            
            df_d['RESPONSAVEL_FINAL'] = df_d['KEY'].map(map_resp).fillna("SEM_MATCH")
            df_d['PERFIL_FINAL'] = df_d['KEY'].map(map_perf).fillna("FRETEIRO")

            # Auditoria de Cartera
            snap_depois = df_m[c_resp_m].value_counts().to_dict()
            df_audit = pd.DataFrame([snap_antes, snap_depois], index=['Antes', 'Depois']).T.fillna(0)
            
            # Log de Telefones (Para a caixa de info)
            ids_d_test = [c for c in df_d.columns if any(x in c.upper() for x in ['ID', 'NOME', 'DOC', 'RESPONSAVEL', 'PERFIL'])]
            melt_test = df_d.melt(id_vars=ids_d_test, value_vars=c_tels_d, value_name='T')
            validos = melt_test['T'].apply(tratar_telefone_v55).dropna().nunique()

            # Gravar tudo na Sessão
            st.session_state.update({
                'm_fin': df_m, 'd_fin': df_d, 'audit_cartera': df_audit,
                'c_resp': c_resp_m, 'c_prio': c_prio_m, 'c_tels': c_tels_d,
                'log_audit_tel': f"📞 Auditoria: Detetados {validos} números únicos de telefone para discagem.",
                'processado': True
            })
            st.rerun()

# --- ÁREA DE RESULTADOS (O LAYOUT QUE DEFINIMOS) ---
if st.session_state.processado:
    st.markdown('<div class="auto-success">✅ <b>Sucesso!</b> Cartera balanceada e sincronizada. Os dados abaixo estão fixos.</div>', unsafe_allow_html=True)
    
    # LINHA 1: AUDITORIA E PERFIL
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("⚖️ 1. Auditoria de Balanceamento")
        st.dataframe(st.session_state.audit_cartera.style.format("{:.0f}"), use_container_width=True)
    with col2:
        st.subheader("⏰ 2. Divisão por Perfil")
        st.dataframe(pd.crosstab(st.session_state.m_fin[st.session_state.c_resp], st.session_state.m_fin['PERFIL_FINAL'], margins=True), use_container_width=True)
    
    # LINHA 2: PRIORIDADES
    st.subheader("🔢 3. Detalhe de Prioridades por Agente")
    st.dataframe(pd.crosstab(st.session_state.m_fin[st.session_state.c_resp], st.session_state.m_fin[st.session_state.c_prio], margins=True), use_container_width=True)
    
    st.markdown(f'<div class="auto-info">{st.session_state.log_audit_tel}</div>', unsafe_allow_html=True)

    # DOWNLOAD FINAL
    st.subheader("📥 4. Exportar Kit Completo")
    zip_final = gerar_zip_v55(st.session_state.m_fin, st.session_state.d_fin, st.session_state.c_resp, st.session_state.c_tels)
    st.download_button(
        label="DESCARREGAR KIT DE 8 FICHEIROS (.ZIP)",
        data=zip_final,
        file_name="Michelin_Kit_Final_V55.zip",
        mime="application/zip"
    )
