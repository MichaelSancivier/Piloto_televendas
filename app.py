import streamlit as st
import pandas as pd
import io
import re
import zipfile
import unicodedata
import numpy as np
import random

# ==============================================================================
# 1. CONFIGURACIÓN VISUAL
# ==============================================================================
st.set_page_config(page_title="Michelin Pilot Command Center", page_icon="🚛", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    h1 { color: #003366; }
    
    /* Botón Gigante */
    .stButton>button { 
        width: 100%; border-radius: 8px; height: 4em; font-weight: bold; font-size: 20px !important;
        background-color: #003366; color: white; border: 2px solid #004080; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.3); transition: all 0.3s;
    }
    .stButton>button:hover { background-color: #FCE500; color: #003366; transform: scale(1.02); }
    
    /* Cajas de Estado */
    .auto-success { padding: 15px; background-color: #d4edda; color: #155724; border-left: 5px solid #28a745; margin-bottom: 10px; border-radius: 5px;}
    .auto-error { padding: 15px; background-color: #f8d7da; color: #721c24; border-left: 5px solid #dc3545; margin-bottom: 10px; border-radius: 5px;}
    
    /* Tablas */
    .preview-box { background-color: #ffffff; padding: 20px; border-radius: 10px; border: 1px solid #ddd; margin-top: 20px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. CEREBRO AUTOMÁTICO
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
    final = None
    tipo = "Inválido"
    if len(nums) == 11 and int(nums[2]) == 9:
        tipo = "Celular"; final = nums
    elif len(nums) == 10 and int(nums[2]) >= 6:
        tipo = "Celular (Corrigido)"; final = nums[:2] + '9' + nums[2:]
    elif len(nums) == 10 and 2 <= int(nums[2]) <= 5:
        tipo = "Fixo"; final = nums
    return (f"+55{final}", tipo) if final else (None, tipo)

def motor_distribuicao(df_m, df_d, col_id_m, col_id_d, col_resp_m, col_prio_m):
    mestre, escravo = df_m.copy(), df_d.copy()
    mestre['KEY'] = mestre[col_id_m].astype(str).str.strip().str.upper()
    escravo['KEY'] = escravo[col_id_d].astype(str).str.strip().str.upper()
    
    ignorar = ['CANAL TELEVENDAS', 'TIME', 'EQUIPE', 'TELEVENDAS', 'NULL', 'NAN', '', 'BACKLOG']
    agentes = [n for n in mestre[col_resp_m].unique() if pd.notna(n) and str(n).strip().upper() not in ignorar]
    
    if len(agentes) > 25: return None, None, f"🚨 ALERTA: Se detectaron {len(agentes)} agentes. Revisa la columna Responsable."

    mask_orfao = mestre[col_resp_m].isna() | mestre[col_resp_m].astype(str).str.strip().str.upper().isin(ignorar)
    
    # Distribución
    if mask_orfao.sum() > 0:
        prios = sorted(mestre.loc[mask_orfao, col_prio_m].dropna().unique()) if col_prio_m else [1]
        for p in prios:
            mask = mask_orfao & (mestre[col_prio_m] == p) if col_prio_m else mask_orfao
            idxs = mestre[mask].index.tolist()
            if idxs:
                random.shuffle(idxs)
                eq = agentes.copy(); random.shuffle(eq)
                mestre.loc[idxs, col_resp_m] = np.resize(eq, len(idxs))

    mapa = dict(zip(mestre['KEY'], mestre[col_resp_m]))
    escravo['RESPONSAVEL_FINAL'] = escravo['KEY'].map(mapa).fillna("SEM_MATCH")
    return mestre, escravo, f"✅ Procesado: {mask_orfao.sum()} leads distribuidos."

def processar_vertical(df, col_id, cols_tel, col_resp, col_doc):
    df['PERFIL'] = df[col_doc].apply(identificar_perfil_doc)
    cols = list(set(col_id + [col_resp, 'PERFIL']))
    melt = df.melt(id_vars=cols, value_vars=cols_tel, var_name='Orig', value_name='Tel_Raw')
    melt['Tel_Clean'], _ = zip(*melt['Tel_Raw'].apply(tratar_telefone))
    return melt.dropna(subset=['Tel_Clean']).drop_duplicates(subset=col_id + ['Tel_Clean'])

def gerar_zip(df_m, df_d_vert, col_resp_m, col_resp_d, modo="MANHA"):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
        if df_d_vert is not None:
            for ag in df_d_vert[col_resp_d].unique():
                safe = normalizar_nome_arquivo(ag)
                if safe in ["SEM_DONO", "SEM_MATCH"]: continue
                df_a = df_d_vert[df_d_vert[col_resp_d] == ag]
                
                if modo == "TARDE":
                    d = io.BytesIO(); df_a.to_excel(d, index=False)
                    zf.writestr(f"{safe}_DISCADOR_REFORCO_TARDE.xlsx", d.getvalue())
                else:
                    fr = df_a[df_a['PERFIL'] == "PEQUENO FROTISTA"]
                    fre = df_a[df_a['PERFIL'] == "FRETEIRO"]
                    if not fr.empty:
                        d = io.BytesIO(); fr.to_excel(d, index=False)
                        zf.writestr(f"{safe}_DISCADOR_MANHA_Frotista.xlsx", d.getvalue())
                    if not fre.empty:
                        d = io.BytesIO(); fre.to_excel(d, index=False)
                        zf.writestr(f"{safe}_DISCADOR_ALMOCO_Freteiro.xlsx", d.getvalue())
        
        if modo == "MANHA" and df_m is not None:
            for ag in df_m[col_resp_m].unique():
                safe = normalizar_nome_arquivo(ag)
                if safe in ["SEM_DONO"]: continue
                df_a = df_m[df_m[col_resp_m] == ag]
                if not df_a.empty:
                    d = io.BytesIO(); df_a.to_excel(d, index=False)
                    zf.writestr(f"{safe}_MAILING_DISTRIBUIDO.xlsx", d.getvalue())
    buf.seek(0)
    return buf

# ==============================================================================
# 4. FRONTEND AUTOMÁTICO
# ==============================================================================

st.title("🚛 Michelin Pilot V37 (Completo)")
st.markdown("---")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    modo = st.radio("Modo:", ("🌅 Manhã", "☀️ Tarde"))
    st.markdown("---")
    file_main = st.file_uploader("Archivo Mestre (.xlsx)", type=["xlsx"])
    file_log = st.file_uploader("Log Discador (.csv/.xlsx)", type=["csv","xlsx"]) if modo == "☀️ Tarde" else None

if file_main:
    xls = pd.ExcelFile(file_main)
    aba_d = next((s for s in xls.sheet_names if 'DISC' in s.upper()), None)
    aba_m = next((s for s in xls.sheet_names if 'MAIL' in s.upper()), None)
    
    if not aba_d or not aba_m: st.error("❌ Faltan pestañas 'Discador' o 'Mailing'."); st.stop()
        
    df_d = pd.read_excel(file_main, sheet_name=aba_d)
    df_m = pd.read_excel(file_main, sheet_name=aba_m)

    # AUTO-CONFIG
    st.subheader("🤖 Configuración Automática")
    col_resp_m = buscar_coluna_smart(df_m, ['RESPONSAVEL', 'RESPONSÁVEL', 'AGENTE'])
    col_id_m = buscar_coluna_smart(df_m, ['ID_CONTRATO', 'ID_CLIENTE', 'CONTRATO'])
    col_prio_m = buscar_coluna_smart(df_m, ['PRIORIDADES', 'PRIORIDADE', 'AGING', 'SCORE'])
    col_doc_m = buscar_coluna_smart(df_m, ['CNPJ_CPF', 'CPF', 'CNPJ', 'DOCUMENTO'])
    
    col_id_d = buscar_coluna_smart(df_d, ['ID_CONTRATO', 'EXTERNAL ID', 'CONTRATO'])
    col_doc_d = buscar_coluna_smart(df_d, ['CNPJ_CPF', 'CPF', 'CNPJ'])
    col_tels_d = [c for c in df_d.columns if any(x in c.upper() for x in ['TEL', 'CEL', 'FONE'])]

    val_resp, msg_resp = validar_coluna_responsavel(df_m, col_resp_m)
    
    if val_resp:
        st.markdown(f"""
        <div class="auto-success">
            <b>✅ ¡Columnas Detectadas!</b><br>
            Resp: {col_resp_m} | ID: {col_id_m} | Prio: {col_prio_m}
        </div>
        """, unsafe_allow_html=True)
        
        if modo == "🌅 Manhã":
            if st.button("🚀 EJECUTAR PILOTO AUTOMÁTICO"):
                with st.spinner("Procesando..."):
                    m, d, msg = motor_distribuicao(df_m, df_d, col_id_m, col_id_d, col_resp_m, col_prio_m)
                    
                    if m is not None:
                        d_vert = processar_vertical(d, [col_id_d], col_tels_d, 'RESPONSAVEL_FINAL', col_doc_d)
                        st.success(msg)
                        
                        # --- AQUÍ ESTÁN LAS TABLAS QUE FALTABAN ---
                        st.markdown('<div class="preview-box">', unsafe_allow_html=True)
                        st.subheader("📊 Auditoría de Resultados")
                        t1, t2 = st.tabs(["⚖️ Justicia (Prioridades)", "⏰ Carga Horaria (Perfil)"])
                        
                        with t1:
                            if col_prio_m:
                                st.write(f"Distribución basada en **{col_prio_m}**:")
                                res_prio = m.groupby([col_resp_m, col_prio_m]).size().unstack(fill_value=0)
                                st.dataframe(res_prio, use_container_width=True)
                            else: st.info("Sin prioridad detectada.")
                            
                        with t2:
                            st.write("Distribución para el Discador (Frotista/Freteiro):")
                            res_perf = d_vert.groupby(['RESPONSAVEL_FINAL', 'PERFIL'])[col_id_d].nunique().unstack(fill_value=0)
                            st.dataframe(res_perf, use_container_width=True)
                        st.markdown('</div>', unsafe_allow_html=True)
                        # ------------------------------------------

                        zip_f = gerar_zip(m, d_vert, col_resp_m, 'RESPONSAVEL_FINAL', "MANHA")
                        st.download_button("📥 DESCARGAR PACK", zip_f, "Michelin_Pack.zip", "application/zip", type="primary")

        elif modo == "☀️ Tarde":
            if file_log:
                if st.button("🔄 GENERAR REFUERZO"):
                    m, d, _ = motor_distribuicao(df_m, df_d, col_id_m, col_id_d, col_resp_m, col_prio_m)
                    try:
                        df_log = pd.read_csv(file_log, sep=None, engine='python') if file_log.name.endswith('.csv') else pd.read_excel(file_log)
                        col_log_id = df_log.columns[0]
                        ids_out = df_log[col_log_id].astype(str).unique()
                        
                        d['KEY_TEMP'] = d[col_id_d].astype(str)
                        d_tarde = d[~d['KEY_TEMP'].isin(ids_out)].copy()
                        d_tarde['PERFIL'] = d_tarde[col_doc_d].apply(identificar_perfil_doc)
                        d_tarde = d_tarde[d_tarde['PERFIL'] == 'PEQUENO FROTISTA']
                        
                        d_vert = processar_vertical(d_tarde, [col_id_d], col_tels_d, 'RESPONSAVEL_FINAL', col_doc_d)
                        
                        # --- TABLA TARDE ---
                        st.write("Resumen Tarde (Pendientes):")
                        st.dataframe(d_vert.groupby(['RESPONSAVEL_FINAL'])[col_id_d].nunique(), use_container_width=True)
                        # -------------------

                        zip_t = gerar_zip(None, d_vert, None, 'RESPONSAVEL_FINAL', "TARDE")
                        st.download_button("📥 DESCARGAR TARDE", zip_t, "Refuerzo_Tarde.zip", "application/zip", type="primary")
                    except Exception as e: st.error(f"Error Log: {e}")
            else: st.info("Sube el Log.")
            
    else:
        st.markdown(f'<div class="auto-error">⚠️ {msg_resp}</div>', unsafe_allow_html=True)
        with st.expander("🛠️ Configuración Manual", expanded=True):
            col_resp_m = st.selectbox("Columna NOMBRES:", df_m.columns.tolist())
