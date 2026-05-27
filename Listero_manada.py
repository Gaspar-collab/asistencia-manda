import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import os
import datetime
import json

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA (¡SIEMPRE PRIMERO!)
# ==========================================
# Configuración de página optimizada para celulares
st.set_page_config(page_title="Asistencia Manada 🐾", page_icon="📝", layout="centered")

# ==========================================
# 2. CONFIGURACIÓN INICIAL
# ==========================================
SPREADSHEET_ID = "1vAZpsEgxJNGTlZGh8UzhfGz05KTc8WRPpAB-lugeAT0"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def autenticar():
    # Lee el JSON del robot directo desde los secretos de la nube
    creds_dict = json.loads(st.secrets["google"]["service_account"])
    # Se conecta a Google usando las credenciales del robot
    return gspread.service_account_from_dict(creds_dict)

# Convertir número de columna a letras de Excel (Ej: 8 -> H, 28 -> AB)
def col_to_letter(col):
    letter = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter

# Cargar nombres dinámicamente desde el Excel de Asistencia
@st.cache_data(ttl=60)
def cargar_personas_asistencia():
    try:
        cliente_sheets = autenticar()
        sheet = cliente_sheets.open_by_key(SPREADSHEET_ID)
        
        # 1. Cargar Niños
        hoja_niños = sheet.worksheet("Asistencia Niños")
        valores_niños = hoja_niños.get_all_values()
        niños = []
        for i in range(2, len(valores_niños)):
            nombre = valores_niños[i][0].strip()
            if not nombre or nombre == "0" or nombre.startswith("0,") or nombre == "Nombre Lobato":
                continue
            niños.append(nombre)
            
        # 2. Cargar Dirigentes (Equipo)
        hoja_equipo = sheet.worksheet("Asistencia Equipo")
        valores_equipo = hoja_equipo.get_all_values()
        equipo = []
        for i in range(2, len(valores_equipo)):
            nombre = valores_equipo[i][0].strip()
            if "Fecha reunión" in nombre or not nombre or nombre == "0" or nombre == "Nombre Adulto":
                break
            equipo.append(nombre)
            
        return niños, equipo
    except Exception as e:
        st.error(f"Error al conectar con la planilla de asistencia: {e}")
        st.exception(e)
        return [], []

def guardar_asistencia_hoja(sheet, fecha_col_str, asistencia_dict, tipo_grupo, es_reunion=False):
    todas_las_filas = sheet.get_all_values()
    if not todas_las_filas:
        return
    
    fila_cabecera = todas_las_filas[0]
    
    col_idx = None
    for idx, cell in enumerate(fila_cabecera):
        if cell.strip() == fecha_col_str:
            col_idx = idx + 1
            break
            
    es_nueva_columna = False
    if col_idx is None:
        es_nueva_columna = True
        last_date_idx = None
        
        for idx, cell in enumerate(fila_cabecera):
            if "/" in cell:
                last_date_idx = idx
        
        if last_date_idx is not None:
            col_idx = last_date_idx + 2  
        else:
            col_idx = 2

    letra_col = col_to_letter(col_idx)
    
    if es_nueva_columna:
        sheet.insert_cols([[""] * len(todas_las_filas)], col=col_idx)
        sheet.update_cell(1, col_idx, fecha_col_str)
        if tipo_grupo == "Equipo":
            val_fila2 = "reu" if es_reunion else "sáb"
        else:
            val_fila2 = ""
        sheet.update_cell(2, col_idx, val_fila2)
        
    updates = []
    en_seccion_reunion = False
    
    for i in range(2, len(todas_las_filas)):
        nombre_celda = todas_las_filas[i][0].strip()
        
        if tipo_grupo == "Equipo":
            if "Fecha reunión" in nombre_celda:
                en_seccion_reunion = True
                continue
            
            if not es_reunion and en_seccion_reunion:
                break
            if es_reunion and not en_seccion_reunion:
                continue
        
        if not nombre_celda or nombre_celda == "0" or "Nombre" in nombre_celda:
            continue
            
        if nombre_celda in asistencia_dict:
            estado = asistencia_dict[nombre_celda]
            num_fila = i + 1
            updates.append({
                'range': f"{letra_col}{num_fila}",
                'values': [[estado]]
            })
            
    if updates:
        sheet.batch_update(updates, value_input_option='USER_ENTERED')

# ==========================================
# 3. INTERFAZ GRÁFICA (STREAMLIT)
# ==========================================
st.title("🐾 Asistencia Manada 2026")
st.write("Pasa la asistencia de sábados o reuniones de equipo de forma rápida.")
st.caption(f"ID de Planilla en uso: `{SPREADSHEET_ID}`")

# Cargar listas desde Google Sheets
lista_niños, lista_dirigentes = cargar_personas_asistencia()

# Selector de fecha automático
fecha_ingresada = st.date_input("📆 Selecciona la fecha de la actividad:", datetime.date.today())
fecha_formateada = fecha_ingresada.strftime("%d/%m")

# AUTOMATIZACIÓN: 5 significa Sábado. Si es distinto, es reunión de equipo.
es_reunion_equipo = fecha_ingresada.weekday() != 5

if es_reunion_equipo:
    st.info(f"👥 **Día de semana detectado.** Se registrará como **Reunión de Equipo** (Sección de abajo en Excel).")
else:
    st.success(f"📊 **Sábado detectado.** Se registrará como **Actividad Normal de Manada** (Sección de arriba en Excel).")

st.info(f"📍 Las asistencias se registrarán en la columna **{fecha_formateada}**")
st.caption("**Guía de marcado:** 🟢 **A** = Asistió | 🔴 **N** = No Justificó | 🟡 **J** = Justificó")

asistencia_niños_final = {}
asistencia_dirigentes_final = {}

# Crear las dos pestañas en la pantalla
tab_niños, tab_dirigentes = st.tabs(["👦 Lobatos (Niños)", "🧑‍💼 Dirigentes (Equipo)"])

with tab_niños:
    if es_reunion_equipo:
        st.warning("📭 Los días de semana no hay actividad con los Lobatos. Esta sección no se registrará en el Excel.")
    elif not lista_niños:
        st.warning("No se encontraron niños en la planilla.")
    else:
        st.write("### Asistencia de Niños")
        for idx, nues_niño in enumerate(lista_niños):
            col_nom, col_opc = st.columns([3, 2])
            with col_nom:
                st.write(f"**{nues_niño}**")
            with col_opc:
                opcion = st.radio(
                    f"Asist-{nues_niño}", ["A", "N", "J"], 
                    index=0, horizontal=True, key=f"n_{idx}_{nues_niño}",
                    label_visibility="collapsed"
                )
                asistencia_niños_final[nues_niño] = opcion

with tab_dirigentes:
    if not lista_dirigentes:
        st.warning("No se encontraron dirigentes en la planilla.")
    else:
        if es_reunion_equipo:
            st.write("### Asistencia: Reunión de Equipo (Dirigentes)")
        else:
            st.write("### Asistencia: Sábado de Manada (Dirigentes)")
            
        for idx, nues_dirigente in enumerate(lista_dirigentes):
            col_nom, col_opc = st.columns([3, 2])
            with col_nom:
                st.write(f"**{nues_dirigente}**")
            with col_opc:
                opcion = st.radio(
                    f"Asist-{nues_dirigente}", ["A", "N", "J"], 
                    index=0, horizontal=True, key=f"d_{idx}_{nues_dirigente}",
                    label_visibility="collapsed"
                )
                asistencia_dirigentes_final[nues_dirigente] = opcion

# ==========================================
# 4. GUARDAR CAMBIOS
# ==========================================
st.write("---")
if st.button("💾 GUARDAR ASISTENCIA COMPLETA", use_container_width=True):
    with st.spinner("Sincronizando con Google Sheets... ⏳"):
        try:
            cliente_sheets = autenticar()
            sheet = cliente_sheets.open_by_key(SPREADSHEET_ID)
            
            if lista_niños and not es_reunion_equipo:
                hoja_niños = sheet.worksheet("Asistencia Niños")
                guardar_asistencia_hoja(hoja_niños, fecha_formateada, asistencia_niños_final, "Niños", es_reunion=False)
                
            if lista_dirigentes:
                hoja_equipo = sheet.worksheet("Asistencia Equipo")
                guardar_asistencia_hoja(hoja_equipo, fecha_formateada, asistencia_dirigentes_final, "Equipo", es_reunion=es_reunion_equipo)
                
            st.success(f"🎉 ¡Asistencia del día ({fecha_formateada}) subida correctamente al Excel!")
            st.balloons()
            st.cache_data.clear()
            
        except Exception as err:
            st.error(f"❌ Ocurrió un error inesperado al guardar los datos: {err}")