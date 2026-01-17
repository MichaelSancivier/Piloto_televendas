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
    .metric-card { background-color: white; padding: 15px; border-radius: 8px; border: 1px solid #ddd; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 2. INTELIGENCIA AUTOMÁTICA (COLUMNAS)
# ==============================================================================

def buscar_coluna_smart(df, keywords_primarias):
    cols_upper = {c.upper(): c for c in df.columns}
    for kw in keywords_primarias:
        # Busca exacta
        if kw in cols_upper: return cols_upper[kw]
        # Busca parcial
        match = next((real for upper, real in cols_upper.items() if kw in upper), None)
        if match: return match
    return None

def validar_coluna_responsavel(df, col_name):
    if not col_name: return False, "No encontrada"
    amostra = df[col_name].dropna().unique()[:5]
    if len(amostra) == 0: return False, "Columna vacía"
    # Chequeo anti-ID (evita seleccionar columnas numéricas)
    primeiro = str(amostra[0]).replace('.', '').replace('-', '').strip()
    if primeiro.isdigit() and len(primeiro) > 6:
        return False, f"⚠️ La columna '{col_name}' parece contener IDs ({primeiro}...), no Nombres."
    return True, "OK"

# ==============================================================================
# 3. LÓGICA DE NEGOCIO (QUIRÚRGICA)
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
    """
    Convierte prioridades a números para ordenar 'quirúrgicamente'.
    Menor número = Más importante (Se conserva).
    Mayor número = Menos importante (Se redistribuye).
    """
    val_str = str(val).upper().strip()
    if "BACKLOG" in val_str: return 99
    if val_str == "NAN" or val_str == "" or val_str == "NONE": return 100
    
    # Intenta sacar número (Ej: "Prioridade 1" -> 1)
    nums = re.findall(r'\d+', val_str)
    if nums: return int(nums[0])
    
    return 50 # Valor medio por defecto

def motor_distribuicao_quirurgico(df_m, df_d, col_id_m, col_id_d, col_resp_m, col_prio_m):
    """
    Algoritmo de Balanceo Quirúrgico:
    1. Respeta dueños actuales.
    2. Calcula meta ideal.
    3. Llena huecos con huérfanos.
    4. Si hay desbalance, quita sobrante empezando por Backlog/Prio Baja.
    """
    mestre = df_m.copy()
    escravo = df_d.copy()
    
    # Claves de Match
    mestre['KEY'] = mestre[col_id_m].astype(str).str.strip().str.upper()
    escravo['KEY'] = escravo[col_id_d].astype(str).str.strip().str.upper()
    
    # 1. IDENTIFICAR AGENTES Y CARGA ACTUAL
    ignorar = ['CANAL TELEVENDAS', 'TIME', 'EQUIPE', 'TELEVENDAS', 'NULL', 'NAN', '', 'BACKLOG', 'SEM DONO']
    agentes = [n for n in mestre[col_resp_m].unique() if pd.notna(n) and str(n).strip().upper() not in ignorar]
    
    if len(agentes) > 25: return None, None, f"🚨 ALERTA: {len(agentes)} agentes detectados. Posible error de columna ID."
    
    # Calcular Meta
    total_registros = len(mestre)
    meta_ideal = total_registros // len(agentes)
    
    # Mapear Prioridad Numérica (Para ordenar)
    mestre['PRIO_SCORE'] = mestre[col_prio_m].apply(calcular_peso_prioridade) if col_prio_m else 1
    
    # Snapshot Inicial para Auditoría
    snapshot_inicial = mestre[col_resp_m].value_counts().to_dict()

    # --- FASE 1: ASIGNAR HUÉRFANOS ---
    mask_orfao = mestre[col_resp_m].isna() | mestre[col_resp_m].astype(str).str.strip().str.upper().isin(ignorar)
    orfaos_idxs = mestre[mask_orfao].index.tolist()
    
    # Barajar huérfanos para aleatoriedad en misma prioridad
    random.shuffle(orfaos_idxs)
    
    # Loop de asignación inteligente (llenar primero a los que tienen menos)
    for idx in orfaos_idxs:
        # Recalcular cargas actuales
        cargas = mestre[col_resp_m].value_counts()
        # Encontrar agente con MENOS carga
        agente_menor_carga = min(agentes, key=lambda x: cargas.get(x, 0))
        mestre.at[idx, col_resp_m] = agente_menor_carga

    # --- FASE 2: NIVELACIÓN QUIRÚRGICA (SI ES NECESARIA) ---
    # Si alguien supera la meta por mucho, le quitamos lo "peor" (Backlog) para dar a otros
    
    # Ordenamos el DF: Prioridad Alta (1) arriba, Backlog (99) abajo.
    # Así, cuando iteramos, iteramos desde los VIPs. Pero queremos quitar los NO-VIPs.
    # Estrategia: Identificar excedentes y moverlos.
    
    for _ in range(len(agentes) * 2): # Iteraciones de seguridad
        cargas = mestre[col_resp_m].value_counts()
        agente_max = max(agentes, key=lambda x: cargas.get(x, 0))
        agente_min = min(agentes, key=lambda x: cargas.get(x, 0))
        
        carga_max = cargas.get(agente_max, 0)
        carga_min = cargas.get(agente_min, 0)
        
        # Si la diferencia es pequeña (ej: < 2), paramos. Equilibrio logrado.
        if (carga_max - carga_min) <= 1:
            break
            
        # Si hay desequilibrio, quitamos 1 registro al MAX y damos al MIN
        # CRITERIO QUIRÚRGICO: Buscar el registro con PEOR prioridad (Mayor Score) de agente_max
        # Ordenamos descendente (99, 98...) para tomar el backlog primero
        candidatos = mestre[mestre[col_resp_m] == agente_max].sort_values('PRIO_SCORE', ascending=False)
        
        if not candidatos.empty:
            idx_a_mover = candidatos.index[0] # El peor registro (Backlog)
            mestre.at[idx_a_mover, col_resp_m] = agente_min

    # --- FASE 3: SINCRONIZACIÓN ---
    mapa = dict(zip(mestre['KEY'], mestre[col_resp_m]))
    escravo['RESPONSAVEL_FINAL'] = escravo['KEY'].map(mapa).fillna("SEM_MATCH")
    
    # Auditoría Final
    snapshot_final = mestre[col_resp_m].value_counts().to_dict()
    df_audit = pd.DataFrame([snapshot_inicial, snapshot_final], index=['Inicio', 'Final']).T.fillna(0)
    df_audit['Cambio'] = df_audit['Final'] - df_audit['Inicio']
    
    return mestre, escravo, df_audit

# ==============================================================================
# 4. GENERADORES DE ARCHIVOS (4 POR CABEZA)
# ==============================================================================

def processar_vertical(df, col_id, cols_tel, col_resp, col_doc):
    df['PERFIL'] = df[col_doc].apply(identificar_perfil_doc)
    cols = list(set(col_id + [col_resp, 'PERFIL']))
    melt = df.melt(id_vars=cols, value_vars=cols_tel, var_name='Orig', value_name='Tel_Raw')
    melt['Tel_Clean'], _ = zip(*melt['Tel_Raw'].apply(tratar_telefone))
    return melt.dropna(subset=['Tel_Clean']).drop_duplicates(subset=col_id + ['Tel_Clean'])

def gerar_zip_completo(df_m, df_d_vert, col_resp_m, col_resp_d, modo="MANHA"):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
        
        # 1. DISCADOR (4 Archivos: Mañana/Tarde x Frotista/Freteiro)
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

        # 2. MAILING (También separado por Perfil para análisis del atendente)
        if modo == "MANHA" and df_m is not None:
            # Añadir perfil al mailing para separar
            col_doc_temp = [c for c in df_m.columns if 'CNPJ' in c.upper() or 'CPF' in c.upper()][0]
            df_m['PERFIL_TEMP'] = df_m[col_doc_temp].apply(identificar_perfil_doc)
            
            for ag in df_m[col_resp_m].unique():
                safe = normalizar_nome_arquivo(ag)
                if safe in ["SEM_DONO"]: continue
                df_a = df_m[df_m[col_resp_m] == ag]
                
                fr = df_a[df_a['PERFIL_TEMP'] == "PEQUENO FROTISTA"]
                fre = df_a[df_a['PERFIL_TEMP'] == "FRETEIRO"]
                
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

st.title("🚛 Michelin Pilot V38 (Surgical Logic)")
st.markdown("---")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Michelin_Logo.svg/1200px-Michelin_Logo.svg.png", width=150)
    modo = st.radio("Modo:", ("🌅 Manhã", "☀️ Tarde"))
    st.markdown("---")
    file_main = st.file_uploader("Archivo Mestre (.xlsx)", type=["xlsx"])
    file_log = st.file_uploader("Log Discador", type=["csv","xlsx"]) if modo == "☀️ Tarde" else None

if file_main:
    xls = pd.ExcelFile(file_main)
    # Auto-detectar Pestañas
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
            <b>✅ Columnas Detectadas:</b> Resp: <i>{col_resp_m}</i> | ID: <i>{col_id_m}</i> | Prio: <i>{col_prio_m}</i>
        </div>
        """, unsafe_allow_html=True)
        
        if modo == "🌅 Manhã":
            if st.button("🚀 EJECUTAR DISTRIBUCIÓN QUIRÚRGICA"):
                with st.spinner("1. Respetando cartera... 2. Asignando Huérfanos... 3. Nivelando Backlog..."):
                    m, d, audit = motor_distribuicao_quirurgico(df_m, df_d, col_id_m, col_id_d, col_resp_m, col_prio_m)
                    
                    if m is not None:
                        # Tablas de Auditoría
                        st.subheader("📊 Auditoría de Balanceo")
                        c1, c2 = st.columns(2)
                        c1.write("**Cambios en la Cartera:**")
                        c1.dataframe(audit.style.format("{:.0f}"), use_container_width=True)
                        
                        if col_prio_m:
                            c2.write("**Distribución Final por Prioridad:**")
                            res_prio = m.groupby([col_resp_m, col_prio_m]).size().unstack(fill_value=0)
                            c2.dataframe(res_prio, use_container_width=True)
                        
                        # Procesar Salida
                        d_vert = processar_vertical(d, [col_id_d], col_tels_d, 'RESPONSAVEL_FINAL', col_doc_d)
                        zip_f = gerar_zip_completo(m, d_vert, col_resp_m, 'RESPONSAVEL_FINAL', "MANHA")
                        
                        st.success("✅ ¡Proceso completado con éxito!")
                        st.download_button("📥 DESCARGAR KIT COMPLETO (.ZIP)", zip_f, "Michelin_Kit_Diario.zip", "application/zip", type="primary")

        elif modo == "☀️ Tarde":
            if file_log:
                if st.button("🔄 GENERAR REFUERZO TARDE"):
                    m, d, _ = motor_distribuicao_quirurgico(df_m, df_d, col_id_m, col_id_d, col_resp_m, col_prio_m)
                    try:
                        df_log = pd.read_csv(file_log, sep=None, engine='python') if file_log.name.endswith('.csv') else pd.read_excel(file_log)
                        col_log_id = df_log.columns[0]
                        
                        ids_out = df_log[col_log_id].astype(str).unique()
                        d['KEY_TEMP'] = d[col_id_d].astype(str)
                        d_tarde = d[~d['KEY_TEMP'].isin(ids_out)].copy()
                        d_tarde['PERFIL'] = d_tarde[col_doc_d].apply(identificar_perfil_doc)
                        d_tarde = d_tarde[d_tarde['PERFIL'] == 'PEQUENO FROTISTA']
                        
                        d_vert = processar_vertical(d_tarde, [col_id_d], col_tels_d, 'RESPONSAVEL_FINAL', col_doc_d)
                        
                        st.write(f"Total Refuerzo Tarde: {len(d_vert)} registros.")
                        zip_t = gerar_zip_completo(None, d_vert, None, 'RESPONSAVEL_FINAL', "TARDE")
                        st.download_button("📥 DESCARGAR TARDE", zip_t, "Refuerzo_Tarde.zip", "application/zip", type="primary")
                    except Exception as e: st.error(f"Error Log: {e}")
            else: st.info("Sube el Log.")
    else:
        st.markdown(f'<div class="auto-error">⚠️ {msg_resp}</div>', unsafe_allow_html=True)
        with st.expander("🛠️ Configuración Manual", expanded=True):
            col_resp_m = st.selectbox("Columna NOMBRES:", df_m.columns.tolist())
