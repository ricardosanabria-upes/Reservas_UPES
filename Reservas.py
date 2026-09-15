import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import pytz

# ============================================================================
# CONFIGURACIÓN STREAMLIT
# ============================================================================
st.set_page_config(page_title="Disponibilidad Espacios UPES", layout="wide", initial_sidebar_state="collapsed")

# ============================================================================
# CONSTANTES
# ============================================================================
ESPACIOS_SIN_CLASES = {"SUM", "Sala de juntas", "Pasillos", "Biblioteca"}
CICLO_FRANJAS = [
    ("06:00", "07:40"),
    ("08:00", "09:40"),
    ("10:00", "11:40"),
    ("13:10", "14:50"),
    ("16:50", "18:30"),
    ("18:30", "20:10"),
]



# ============================================================================
# FUNCIONES DE DATOS
# ============================================================================

@st.cache_data
def cargar_horario_ciclo():
    """Carga Excel de horario desde GitHub y lo pivotea a formato largo"""
    try:
        url = "https://raw.githubusercontent.com/ricardosanabria-upes/Consulta_Disponibilidad_UPES/main/DETALLE%20AULAS%20CICLO%20ACTUAL.xlsx"
        df_raw = pd.read_excel(url)
        
        # Pivotear: columnas de aulas → filas
        # ID vars: Dia, Hora. Variable name: Aula. Value name: Curso
        aulas = [col for col in df_raw.columns if col not in ['Dia', 'Hora']]
        df = pd.melt(df_raw, id_vars=['Dia', 'Hora'], value_vars=aulas, 
                     var_name='Aula', value_name='Curso')
        
        # Filtrar filas vacías (NaN)
        df = df.dropna(subset=['Curso'])
        df['Curso'] = df['Curso'].astype(str).str.strip()
        
        # Parsear hora de inicio y fin
        def parse_horas(hora_str):
            try:
                partes = hora_str.replace('–', '-').split('-')
                h_ini = partes[0].strip()
                h_fin = partes[1].strip() if len(partes) > 1 else h_ini
                return h_ini, h_fin
            except:
                return None, None
        
        df[['Hora Inicio', 'Hora Fin']] = df['Hora'].apply(
            lambda x: pd.Series(parse_horas(x))
        )
        df = df.dropna(subset=['Hora Inicio', 'Hora Fin'])
        df['Ocupada'] = True
        
        return df
    except Exception as e:
        st.error(f"Error al cargar horario: {e}")
        return None

@st.cache_data
def cargar_reservas():
    """Carga reservas desde Google Sheets"""
    try:
        # Verificar si existen las credenciales
        if "gcp_service_account" not in st.secrets or "SHEETS_URL" not in st.secrets:
            return pd.DataFrame()
        
        SHEETS_URL = st.secrets["SHEETS_URL"]
        # Usar gspread para leer Google Sheets
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(SHEETS_URL).sheet1
        reservas = sheet.get_all_records()
        df = pd.DataFrame(reservas)
        return df
    except Exception as e:
        st.warning(f"No se pudieron cargar reservas: {e}")
        return pd.DataFrame()

# ============================================================================
# FUNCIONES DE LÓGICA
# ============================================================================

def tiempo_a_minutos(hora_str):
    """Convierte HH:MM a minutos desde 00:00"""
    try:
        h, m = map(int, hora_str.split(":"))
        return h * 60 + m
    except:
        return 0

def minutos_a_tiempo(minutos):
    """Convierte minutos a HH:MM"""
    h = minutos // 60
    m = minutos % 60
    return f"{h:02d}:{m:02d}"

def tiene_clases(instalacion, df_horario):
    """Detecta si una instalación tiene clases registradas"""
    if df_horario is None or df_horario.empty:
        return False
    df_inst = df_horario[
        (df_horario["Aula"].str.strip() == instalacion.strip()) &
        (df_horario["Ocupada"] == True)
    ]
    return len(df_inst) > 0

def obtener_bloques_dia(instalacion, df_horario, df_reservas):
    """
    Obtiene bloques de clases y reservas para una instalación.
    Retorna: {
        'clases': [(inicio_min, fin_min, curso), ...],
        'reservas': [(inicio_min, fin_min, responsable, motivo), ...],
        'traslapes': [(inicio_min, fin_min, duracion_min), ...]
    }
    """
    bloques = {
        'clases': [],
        'reservas': [],
        'traslapes': []
    }
    
    # Filtrar clases para esta instalación
    if df_horario is not None and not df_horario.empty:
        df_inst = df_horario[
            (df_horario["Aula"].str.strip() == instalacion.strip()) &
            (df_horario["Ocupada"] == True)
        ]
        for _, row in df_inst.iterrows():
            try:
                inicio_min = tiempo_a_minutos(str(row["Hora Inicio"]))
                fin_min = tiempo_a_minutos(str(row["Hora Fin"]))
                curso = row.get("Curso", "")
                bloques['clases'].append((inicio_min, fin_min, curso))
            except:
                pass
    
    # Filtrar reservas para esta instalación
    if df_reservas is not None and not df_reservas.empty:
        df_res = df_reservas[
            (df_reservas.get("Aula", "").str.strip() == instalacion.strip())
        ]
        for _, row in df_res.iterrows():
            try:
                inicio_min = tiempo_a_minutos(str(row.get("Hora Inicio", "00:00")))
                fin_min = tiempo_a_minutos(str(row.get("Hora Fin", "00:00")))
                responsable = row.get("Responsable", "")
                motivo = row.get("Motivo", "")
                bloques['reservas'].append((inicio_min, fin_min, responsable, motivo))
            except:
                pass
    
    # Detectar traslapes (reservas que se superponen con clases)
    traslapes_detectados = []
    for r_inicio, r_fin, _, _ in bloques['reservas']:
        for c_inicio, c_fin, _ in bloques['clases']:
            # ¿Se solapan?
            overlap_inicio = max(r_inicio, c_inicio)
            overlap_fin = min(r_fin, c_fin)
            if overlap_inicio < overlap_fin:
                duracion = overlap_fin - overlap_inicio
                traslapes_detectados.append((overlap_inicio, overlap_fin, duracion))
    
    bloques['traslapes'] = traslapes_detectados
    return bloques

def calcular_disponibles(inicio_dia, fin_dia, clases, reservas):
    """
    Calcula bloques disponibles entre clases y reservas.
    Retorna lista de (inicio_min, fin_min)
    """
    # Ordenar todos los eventos
    eventos = sorted(
        [(c[0], c[1], 'clase') for c in clases] +
        [(r[0], r[1], 'reserva') for r in reservas],
        key=lambda x: x[0]
    )
    
    disponibles = []
    cursor = inicio_dia
    
    for evento_inicio, evento_fin, _ in eventos:
        if cursor < evento_inicio:
            disponibles.append((cursor, evento_inicio))
        cursor = max(cursor, evento_fin)
    
    if cursor < fin_dia:
        disponibles.append((cursor, fin_dia))
    
    return disponibles

# ============================================================================
# RENDERIZADO SVG
# ============================================================================

def generar_svg_calendario(instalacion, fecha, bloques, disponibles):
    """
    Genera SVG del calendario vertical (Google Calendar style).
    1 hora = 25px de altura
    """
    inicio_dia = tiempo_a_minutos("06:00")  # 6 AM
    fin_dia = tiempo_a_minutos("20:10")     # 8:10 PM
    
    altura_total = (fin_dia - inicio_dia) / 60 * 25  # px por hora * horas
    ancho = 600
    
    svg_lines = []
    svg_lines.append(f'<svg width="100%" viewBox="0 0 {ancho + 100} {altura_total + 150}" xmlns="http://www.w3.org/2000/svg">')
    svg_lines.append(f'<style>')
    svg_lines.append(f'.time-label {{ font-size: 11px; fill: var(--text-secondary); text-anchor: end; }}')
    svg_lines.append(f'.bloque-text {{ font-size: 10px; }}')
    svg_lines.append(f'.bloque-sub {{ font-size: 8px; }}')
    svg_lines.append(f'a {{ cursor: pointer; }}')
    svg_lines.append(f'</style>')
    
    # Título
    svg_lines.append(f'<text x="40" y="25" style="font-size: 16px; font-weight: bold;">{instalacion} — {fecha.strftime("%a %d/%m/%Y")}</text>')
    
    y_base = 60
    
    # Líneas de horas
    for hora in range(6, 21):
        minutos_desde_inicio = (hora * 60) - inicio_dia
        y_pos = y_base + (minutos_desde_inicio / 60) * 25
        svg_lines.append(f'<text x="35" y="{y_pos + 4}" class="time-label">{hora:02d}:00</text>')
        svg_lines.append(f'<line x1="40" y1="{y_pos}" x2="{ancho + 40}" y2="{y_pos}" stroke="var(--border)" stroke-width="0.5"/>')
    
    # Línea final
    y_fin = y_base + altura_total
    svg_lines.append(f'<line x1="40" y1="{y_fin}" x2="{ancho + 40}" y2="{y_fin}" stroke="var(--border)" stroke-width="1"/>')
    
    # Dibujar clases
    for c_inicio, c_fin, curso in bloques['clases']:
        y_inicio = y_base + ((c_inicio - inicio_dia) / 60) * 25
        altura = ((c_fin - c_inicio) / 60) * 25
        svg_lines.append(f'<rect x="50" y="{y_inicio}" width="{ancho - 20}" height="{altura}" fill="var(--bg-danger)" stroke="var(--border-danger)" stroke-width="1" rx="4"/>')
        svg_lines.append(f'<text x="60" y="{y_inicio + 14}" class="bloque-text" style="fill: var(--text-danger); font-weight: bold;">🔴 {curso[:30]}</text>')
        svg_lines.append(f'<text x="60" y="{y_inicio + altura - 4}" class="bloque-sub" style="fill: var(--text-danger);">{minutos_a_tiempo(c_inicio)}–{minutos_a_tiempo(c_fin)}</text>')
    
    # Dibujar reservas
    for r_inicio, r_fin, responsable, motivo in bloques['reservas']:
        y_inicio = y_base + ((r_inicio - inicio_dia) / 60) * 25
        altura = ((r_fin - r_inicio) / 60) * 25
        svg_lines.append(f'<rect x="50" y="{y_inicio}" width="{ancho - 20}" height="{altura}" fill="var(--bg-warning)" stroke="var(--border-warning)" stroke-width="1" rx="4"/>')
        svg_lines.append(f'<text x="60" y="{y_inicio + 14}" class="bloque-text" style="fill: var(--text-warning); font-weight: bold;">🟡 {responsable[:30]}</text>')
        svg_lines.append(f'<text x="60" y="{y_inicio + altura - 4}" class="bloque-sub" style="fill: var(--text-warning);">{minutos_a_tiempo(r_inicio)}–{minutos_a_tiempo(r_fin)} ({motivo[:20]})</text>')
    
    # Dibujar traslapes (OPCIÓN C: borde grueso rojo + fondo claro)
    for t_inicio, t_fin, _ in bloques['traslapes']:
        y_inicio = y_base + ((t_inicio - inicio_dia) / 60) * 25
        altura = ((t_fin - t_inicio) / 60) * 25
        svg_lines.append(f'<rect x="50" y="{y_inicio}" width="{ancho - 20}" height="{altura}" fill="#ffcccc" stroke="var(--border-danger)" stroke-width="3" rx="4"/>')
        svg_lines.append(f'<text x="60" y="{y_inicio + altura/2}" class="bloque-text" style="fill: var(--text-danger); font-weight: bold;">⚠ TRASLAPE</text>')
    
    # Dibujar disponibles
    for disp_inicio, disp_fin in disponibles:
        y_inicio = y_base + ((disp_inicio - inicio_dia) / 60) * 25
        altura = ((disp_fin - disp_inicio) / 60) * 25
        duracion_min = disp_fin - disp_inicio
        duracion_str = f"{duracion_min // 60}h {duracion_min % 60}m" if duracion_min >= 60 else f"{duracion_min}m"
        
        svg_lines.append(f'<rect x="50" y="{y_inicio}" width="{ancho - 20}" height="{altura}" fill="var(--bg-success)" stroke="var(--border-success)" stroke-width="2" rx="4"/>')
        svg_lines.append(f'<text x="60" y="{y_inicio + 14}" class="bloque-text" style="fill: var(--text-success); font-weight: bold;">✅ Disponible ({duracion_str})</text>')
    
    svg_lines.append(f'</svg>')
    return "\n".join(svg_lines)

# ============================================================================
# INTERFAZ STREAMLIT
# ============================================================================

st.title("📅 Disponibilidad de Espacios UPES")

# Cargar datos
df_horario = cargar_horario_ciclo()
df_reservas = cargar_reservas()

if df_horario is None or df_horario.empty:
    st.error("No se pudo cargar el horario del ciclo")
    st.stop()

# Obtener lista de instalaciones únicas
instalaciones = sorted(set(df_horario["Aula"].unique()))

# Agregar espacios sin clases
ESPACIOS_ADICIONALES = {"SUM", "Sala de juntas", "Pasillos", "Biblioteca"}
instalaciones = sorted(set(instalaciones) | ESPACIOS_ADICIONALES)

# Sidebar: Filtros
st.sidebar.title("Filtros")
instalacion_sel = st.sidebar.selectbox("Seleccionar instalación", instalaciones, index=0)

# Obtener bloques (sin fecha, solo del horario del ciclo)
bloques = obtener_bloques_dia(instalacion_sel, df_horario, df_reservas)
clases = bloques['clases']
reservas = bloques['reservas']
traslapes = bloques['traslapes']

# Calcular disponibles
inicio_dia = tiempo_a_minutos("06:00")
fin_dia = tiempo_a_minutos("20:10")
disponibles = calcular_disponibles(inicio_dia, fin_dia, clases, reservas)

# Métricas en header
col1, col2, col3, col4 = st.columns(4)

libres = len(disponibles)
col1.metric("✅ Bloques libres", libres)
col2.metric("🔴 Clases", len(clases))
col3.metric("🟡 Reservas", len(reservas))

duracion_total_traslapes = sum(t[2] for t in traslapes)
if traslapes:
    duracion_str = f"{duracion_total_traslapes // 60}h {duracion_total_traslapes % 60}m" if duracion_total_traslapes >= 60 else f"{duracion_total_traslapes}m"
    col4.metric("⚠ Traslape", f"{len(traslapes)} - {duracion_str}")
else:
    col4.metric("⚠ Traslape", "0")

# Renderizar calendario
st.subheader(f"Calendario del ciclo: {instalacion_sel}")
fecha_hoy = datetime.now()
svg_html = generar_svg_calendario(instalacion_sel, fecha_hoy, bloques, disponibles)
st.write(svg_html, unsafe_allow_html=True)