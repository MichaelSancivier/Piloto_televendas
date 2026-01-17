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
st.set_page_config(page_title="Michelin Pilot Command Center", page_icon="🚛", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    h1 { color: #003366; }
    .stButton>button { 
        width: 100%; height: 4em; font-weight: bold; font-size: 20px !important;
        background-color: #003366; color: white; border: 2px solid #004080; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.3); transition: all 0.3s;
    }
    .stButton>button:hover { background-color: #FCE500; color: #003366; transform: scale(1.02); }
    .auto-success { padding: 15px; background-color: #d4edda; color: #155724; border-left: 5px solid #28a745; margin-bottom: 10px; border-radius: 5px;}
    .auto-error { padding: 15px; background-color: #f8d7da; color: #721c24; border-left: 5px solid #dc3545; margin-bottom: 10px; border-radius: 5px;}
    div[data-testid="stDataFrame"] { width: 100%; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. AUTO-CONFIGURAÇÃO
# ==============================================================================

def buscar_coluna_smart(df, keywords_primarias):
    cols_upper = {c.upper(): c for c in df.columns}
    for kw in keywords_primarias:
        if kw in cols_upper: return cols_upper[kw]
        match = next((real for upper, real in cols_upper.items() if kw in upper), None)
        if match: return match
    return None

def validar_coluna_responsavel(df, col_name):
    if not col_name: return False, "No encontrada"
    amostra = df[col_name].dropna().unique()[:5]
    if len(amostra) == 0: return False, "Columna vacía"
    primeiro = str(amostra[0]).replace('.', '').replace('-', '').strip()
    if primeiro.isdigit() and len(primeiro) > 6:
        return False, f"⚠️ La columna '{col_name}' parece contener IDs ({primeiro}...), no Nombres."
    return True, "OK"

# ==============================================================================
# 3. LÓGICA DE NEGOCIO
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
    final = None; tipo = "Inválido"
    if len(nums) == 11 and int(nums[2]) == 9: tipo = "Celular"; final = nums
    elif len(nums) == 10 and int(nums[2]) >= 6: tipo = "Celular (Corrigido)"; final = nums[:2] + '9' + nums[2:]
    elif len(nums) == 10 and 2 <= int(nums[2]) <= 5: tipo = "Fixo"; final = nums
    return (f"+55{final}", tipo) if final else (None, tipo)

def calcular_peso_prioridade(val):
    val_str = str(val).upper().strip()
    if "BACKLOG" in val_str: return 99
    if val_str == "NAN" or val_str == "" or val_str == "NONE": return 100
    nums = re.findall(r'\d+', val_str)
    if nums: return int(nums[0])
    return 50

def motor_distribuicao_quirurgico(df_m, df_d, col_id_m, col_id_d, col_resp_m, col_prio_m, col_doc_m):
    mestre = df_m.copy()
    escravo = df_d.copy()
    
    # Standardize Keys
    mestre['KEY'] = mestre[col_id_m].astype(str).str.strip().str.upper()
    escravo['KEY'] = escravo[col_id_d].astype(str).str.strip().str.upper()
    
    # 1. CALCULATE PROFILE IN MAILING (THE MASTER SOURCE)
    # This ensures consistency. We calculate here and pass it down.
    mestre['PERFIL_FINAL'] = mestre[col_doc_m].apply(identificar_perfil_doc)
    
    ignorar = ['CANAL TELEVENDAS', 'TIME', 'EQUIPE', 'TELEVENDAS', 'NULL', 'NAN', '', 'BACKLOG', 'SEM DONO']
    agentes = [n for n in mestre[col_resp_m].unique() if pd.notna(n) and str(n).strip().upper() not in ignorar]
    
    if len(agentes) > 25: return None, None, f"🚨 ALERTA: {len(agentes)} agentes. Error columna ID."
    
    snapshot_inicial = mestre[col_resp_m].value_counts().to_dict()

    # --- ASIGNACIÓN ---
    mestre['PRIO_SCORE'] = mestre[col_prio_m].apply(calcular_peso_prioridade) if col_prio_m else 1
    mask_orfao = mestre[col_resp_m].isna() | mestre[col_resp_m].astype(str).str.strip().str.upper().isin(ignorar)
    orfaos_idxs = mestre[mask_orfao].index.tolist()
    random.shuffle(orfaos_idxs)
    
    for idx in orfaos_idxs:
        cargas = mestre[col_resp_m].value_counts()
        agente_menor_carga = min(agentes, key=lambda x: cargas.get(x, 0))
        mestre.at[idx, col_resp_m] = agente_menor_carga

    # --- NIVELACIÓN ---
    for _ in range(len(agentes) * 100): 
        cargas = mestre[col_resp_m].value_counts()
        agente_max = max(agentes, key=lambda x: cargas.get(x, 0))
        agente_min = min(agentes, key=lambda x: cargas.get(x, 0))
        if (cargas.get(agente_max, 0) - cargas.get(agente_min, 0)) <= 1: break
        candidatos = mestre[mestre[col_resp_m] == agente_max].sort_values('PRIO_SCORE', ascending=False)
        if not candidatos.empty:
            mestre.at[candidatos.index[0], col_resp_m] = agente_min

    # --- SYNC (NOMBRE + PERFIL) ---
    # Sync Responsible
    mapa_resp = dict(zip(mestre['KEY'], mestre[col_resp_m]))
    escravo['RESPONSAVEL_FINAL'] = escravo['KEY'].map(mapa_resp).fillna("SEM_MATCH")
    
    # Sync Profile (Here is the Fix!)
    mapa_perfil = dict(zip(mestre['KEY'], mestre['PERFIL_FINAL']))
    escravo['PERFIL_FINAL'] = escravo['KEY'].map(mapa_perfil)
    # If missing match, default to Freteiro (safer)
    escravo['PERFIL_FINAL'] = escravo['PERFIL_FINAL'].fillna("FRETEIRO")
    
    snapshot_final = mestre[col_resp_m].value_counts().to_dict()
    df_audit = pd.DataFrame([snapshot_inicial, snapshot_final], index=['Inicio', 'Final']).T.fillna(0)
    df_audit['Cambio'] = df_audit['Final'] - df_audit['Inicio']
    
    return mestre, escravo, df_audit

# ==============================================================================
# 4. GENERACIÓN DE ARCHIVOS
# ==============================================================================

def processar_vertical(df, col_id, cols_tel, col_resp_col, col_perfil_col):
    # We use the synced columns 'RESPONSAVEL_FINAL' and 'PERFIL_FINAL'
    cols = list(set(col_id + [col_resp_col, col_perfil_col]))
    melt = df.melt(id_vars=cols, value_vars=cols_tel, var_name='Orig', value_name='Tel_Raw')
    melt['Tel_Clean'], _ = zip(*melt['Tel_Raw'].apply(tratar_telefone))
    return melt.dropna(subset=['Tel_Clean']).drop_duplicates(subset=col_id + ['Tel_Clean'])

def gerar_zip_completo(df_m, df_d_vert, col_resp_m, col_resp_d_final, col_perfil_d_final, modo="MANHA"):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
        
        # --- DISCADOR ---
        if df_d_vert is not None:
            for ag in df_d_vert[col_resp_d_final].unique():
                safe = normalizar_nome_arquivo(ag)
                if safe in ["SEM_DONO", "SEM_MATCH"]: continue
                df_a = df_d_vert[df_d_vert[col_resp_d_final] == ag]
                
                if modo == "TARDE":
                    d = io.BytesIO(); df_a.to_excel(d, index=False)
                    zf.writestr(f"{safe}_DISCADOR_REFORCO_TARDE.xlsx", d.getvalue())
                else:
                    # Using the Synced Profile Column
                    fr = df_a[df_a[col_perfil_d_final] == "PEQUENO FROTISTA"]
                    fre = df_a[df_a[col_perfil_d_final] == "FRETEIRO"]
                    
                    if not fr.empty:
                        d = io.BytesIO(); fr.to_excel(d, index=False)
                        zf.writestr(f"{safe}_DISCADOR_MANHA_Frotista.xlsx", d.getvalue())
                    if not fre.empty:
                        d = io.BytesIO(); fre.to_excel(d, index=False)
                        zf.writestr(f"{safe}_DISCADOR_ALMOCO_Freteiro.xlsx", d.getvalue())

        # --- MAILING ---
        if modo == "MANHA" and df_m is not None:
            # Mailing already has PERFIL_FINAL calculated in motor
            for ag in df_m[col_resp_m].unique():
                safe = normalizar_nome_arquivo(ag)
                if safe in ["SEM_DONO"]: continue
                df_a = df_m[df_m[col_resp_m] == ag]
                
                fr = df_a[df_a['PERFIL_FINAL'] == "PEQUENO FROTISTA"]
                fre = df_a[df_a['PERFIL_FINAL'] == "FRETEIRO"]
                
                if not fr.empty:
                    d = io.BytesIO(); fr.to_excel(d, index=False)
                    zf.writestr(f"{safe}_MAILING_MANHA_Frotista.xlsx", d.getvalue())
                if not fre.empty:
                    d = io.BytesIO(); fre.to_excel(d, index=False)
                    zf.writestr(f"{safe}_MAILING_ALMOCO_Freteiro.xlsx", d.getvalue())
    buf.seek(0)
    return buf

# ==============================================================================
# 5. FRONTEND
# ==============================================================================

st.title("🚛 Michelin Pilot V42 (Sync Fix)")
st.markdown("---")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    modo = st.radio("Modo:", ("🌅 Manhã", "☀️ Tarde"))
    file_main = st.file_uploader("Archivo Mestre (.xlsx)", type=["xlsx"])
    file_log = st.file_uploader("Log Discador", type=["csv","xlsx"]) if modo == "☀️ Tarde" else None

if file_main:
    xls = pd.ExcelFile(file_main)
    aba_d = next((s for s in xls.sheet_names if 'DISC' in s.upper()), None)
    aba_m = next((s for s in xls.sheet_names if 'MAIL' in s.upper()), None)
    if not aba_d or not aba_m: st.error("❌ Faltan pestañas."); st.stop()
    
    df_d = pd.read_excel(file_main, sheet_name=aba_d)
    df_m = pd.read_excel(file_main, sheet_name=aba_m)

    col_resp_m = buscar_coluna_smart(df_m, ['RESPONSAVEL', 'RESPONSÁVEL', 'AGENTE'])
    col_id_m = buscar_coluna_smart(df_m, ['ID_CONTRATO', 'ID_CLIENTE', 'CONTRATO'])
    col_prio_m = buscar_coluna_smart(df_m, ['PRIORIDADES', 'PRIORIDADE', 'AGING', 'SCORE'])
    col_doc_m = buscar_coluna_smart(df_m, ['CNPJ_CPF', 'CPF', 'CNPJ', 'DOCUMENTO'])
    
    col_id_d = buscar_coluna_smart(df_d, ['ID_CONTRATO', 'EXTERNAL ID', 'CONTRATO'])
    # Documento Discador es menos relevante ahora, usamos Mailing
    col_tels_d = [c for c in df_d.columns if any(x in c.upper() for x in ['TEL', 'CEL', 'FONE'])]

    val_resp, msg_resp = validar_coluna_responsavel(df_m, col_resp_m)
    
    if val_resp:
        st.markdown(f"""<div class="auto-success">✅ <b>Detectado:</b> Resp: {col_resp_m} | ID: {col_id_m}</div>""", unsafe_allow_html=True)
        
        if modo == "🌅 Manhã":
            if st.button("🚀 EJECUTAR"):
                # PASAMOS col_doc_m AL MOTOR PARA CALCULAR PERFIL MASTER
                m, d, audit = motor_distribuicao_quirurgico(df_m, df_d, col_id_m, col_id_d, col_resp_m, col_prio_m, col_doc_m)
                
                if m is not None:
                    st.subheader("📊 Auditoría")
                    c1, c2 = st.columns(2)
                    c1.write("**Cartera:**"); c1.dataframe(audit.style.format("{:.0f}"), use_container_width=True)
                    
                    c2.write("**Distribución Perfil (A = B):**")
                    # Usamos PERFIL_FINAL que viene del Motor
                    res_perfil = pd.crosstab(m[col_resp_m], m['PERFIL_FINAL'], margins=True, margins_name="TOTAL")
                    c2.dataframe(res_perfil, use_container_width=True)
                    
                    if col_prio_m:
                        st.write("**Prioridades:**")
                        res_prio = pd.crosstab(m[col_resp_m], m[col_prio_m], margins=True, margins_name="TOTAL")
                        st.dataframe(res_prio, use_container_width=True)
                    
                    # Verticalizamos usando las columnas SINCRONIZADAS
                    d_vert = processar_vertical(d, [col_id_d], col_tels_d, 'RESPONSAVEL_FINAL', 'PERFIL_FINAL')
                    
                    zip_f = gerar_zip_completo(m, d_vert, col_resp_m, 'RESPONSAVEL_FINAL', 'PERFIL_FINAL', "MANHA")
                    st.success("✅ ¡Archivos generados correctamente!")
                    st.download_button("📥 DESCARGAR KIT", zip_f, "Michelin_Kit_V42.zip", "application/zip", type="primary")

        elif modo == "☀️ Tarde":
            if file_log:
                if st.button("🔄 REFUERZO"):
                    m, d, _ = motor_distribuicao_quirurgico(df_m, df_d, col_id_m, col_id_d, col_resp_m, col_prio_m, col_doc_m)
                    try:
                        df_log = pd.read_csv(file_log, sep=None, engine='python') if file_log.name.endswith('.csv') else pd.read_excel(file_log)
                        col_log_id = df_log.columns[0]
                        ids_out = df_log[col_log_id].astype(str).unique()
                        
                        d['KEY_TEMP'] = d[col_id_d].astype(str)
                        d_tarde = d[~d['KEY_TEMP'].isin(ids_out)].copy()
                        # Usamos PERFIL_FINAL sincronizado
                        d_tarde = d_tarde[d_tarde['PERFIL_FINAL'] == 'PEQUENO FROTISTA']
                        
                        d_vert = processar_vertical(d_tarde, [col_id_d], col_tels_d, 'RESPONSAVEL_FINAL', 'PERFIL_FINAL')
                        st.write(f"Refuerzo: {len(d_vert)} registros.")
                        zip_t = gerar_zip_completo(None, d_vert, None, 'RESPONSAVEL_FINAL', 'PERFIL_FINAL', "TARDE")
                        st.download_button("📥 DESCARGAR", zip_t, "Refuerzo_Tarde.zip", "application/zip", type="primary")
                    except Exception as e: st.error(f"Error: {e}")
            else: st.info("Sube el Log.")
    else:
        st.markdown(f'<div class="auto-error">⚠️ {msg_resp}</div>', unsafe_allow_html=True)
        with st.expander("🛠️ Manual", expanded=True):
            col_resp_m = st.selectbox("Columna NOMBRES:", df_m.columns.tolist())
