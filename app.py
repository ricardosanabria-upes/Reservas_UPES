import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime, date, time, timedelta

# ============================================================================
# CONFIGURACIÓN
# ============================================================================
st.set_page_config(
    page_title="Disponibilidad Espacios UPES",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
    .titulo { font-size: 2rem; font-weight: 700; color: #1e1b4b; margin-bottom: 0; }
    .sub { color: #64748b; font-size: 0.95rem; margin-bottom: 1rem; }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# CONSTANTES
# ============================================================================
try:
    SHEETS_URL = st.secrets["SHEETS_URL"]
except Exception:
    SHEETS_URL = os.environ.get("SHEETS_URL", "")

EXCEL_GITHUB_URL = "https://raw.githubusercontent.com/ricardosanabria-upes/Consulta_Disponibilidad_UPES/main/DETALLE%20AULAS%20CICLO%20ACTUAL.xlsx?v=2"

ESPACIOS_ADICIONALES = {"SUM", "Sala de juntas", "Pasillos", "Biblioteca"}

DIA_SEMANA = {
    0: "1.Lunes", 1: "2.Martes", 2: "3.Miercoles",
    3: "4.Jueves", 4: "5.Viernes", 5: "6.Sabado", 6: "7.Domingo",
}

def normalizar_aula(aula: str) -> str:
    mapeo = {
        "A-21": "A-21 C/Acondicionado",
        "A-22": "A-22 C/Acondicionado",
        "A-34": "A-34 (Mesas de dibujo)",
    }
    return mapeo.get(aula.strip(), aula.strip())

# ============================================================================
# CARGAR DATOS
# ============================================================================

@st.cache_data(ttl=300)
def cargar_reservas():
    """Carga reservas desde Google Sheets - LEE DIRECTAMENTE LA URL COMO CSV"""
    try:
        if not SHEETS_URL:
            return None

        # EXACTAMENTE COMO EL CÓDIGO QUE FUNCIONA: pd.read_csv(SHEETS_URL, header=1)
        df = pd.read_csv(SHEETS_URL, header=1)
        df.columns = df.columns.str.strip()

        # Mapear nombres de columnas
        rename = {}
        for col in df.columns:
            cl = col.lower().strip()
            if "instalación solicitada" in cl or "instalacion solicitada" in cl:
                rename[col] = "instalacion"
            elif "fecha del evento" in cl:
                rename[col] = "fecha"
            elif "hora de inicio" in cl:
                rename[col] = "hora_inicio"
            elif "hora de finalización" in cl or "finalización exacta" in cl or "finalizacion" in cl:
                rename[col] = "hora_fin"
            elif "nombre completo del solicitante" in cl:
                rename[col] = "nombre"
            elif "nombre y descripción" in cl or "descripción de la actividad" in cl or "descripcion de la actividad" in cl:
                rename[col] = "actividad"

        df = df.rename(columns=rename)

        # Parsear fecha
        if "fecha" in df.columns:
            df["fecha_date"] = pd.to_datetime(df["fecha"], dayfirst=True, errors="coerce").dt.date

        # Parsear hora
        def parse_hora(val):
            s = str(val).strip()
            for fmt in ["%H:%M:%S", "%H:%M"]:
                try:
                    return datetime.strptime(s, fmt).time()
                except:
                    pass
            return None

        if "hora_inicio" in df.columns:
            df["hora_inicio_t"] = df["hora_inicio"].apply(parse_hora)
        if "hora_fin" in df.columns:
            df["hora_fin_t"] = df["hora_fin"].apply(parse_hora)

        return df
    except Exception as e:
        st.warning(f"Error cargando reservas: {str(e)[:50]}")
        return None

@st.cache_data(ttl=300)
def cargar_horario():
    """Carga horario desde GitHub"""
    try:
        import requests
        resp = requests.get(EXCEL_GITHUB_URL)
        resp.raise_for_status()
        df_raw = pd.read_excel(io.BytesIO(resp.content))
        df_raw.columns = df_raw.columns.str.strip()
        df_raw["Dia"] = df_raw["Dia"].ffill()
        df_raw["Hora"] = df_raw["Hora"].ffill()

        aulas = [c for c in df_raw.columns if c not in ["Dia", "Hora"]]
        filas = []

        for _, row in df_raw.iterrows():
            dia = str(row["Dia"]).strip()
            hora = str(row["Hora"]).strip()

            if not dia or dia == "nan" or not hora or hora == "nan":
                continue

            try:
                partes_h = hora.replace("–", "-").split("-")
                h_ini = datetime.strptime(partes_h[0].strip(), "%H:%M").time()
                h_fin = datetime.strptime(partes_h[1].strip(), "%H:%M").time()
            except:
                h_ini = h_fin = None

            for aula in aulas:
                val = row[aula]
                ocupada = not (pd.isna(val) or str(val).strip() == "")

                if ocupada:
                    texto = str(val).strip()
                    partes = texto.split()
                    codigo = partes[0] if partes else ""
                    seccion = partes[-1] if len(partes) > 1 and partes[-1].isdigit() else ""
                    nombre = " ".join(partes[1:-1]) if seccion else " ".join(partes[1:])
                else:
                    codigo = seccion = nombre = ""

                filas.append({
                    "Dia": dia, "Hora": hora,
                    "HoraInicio": h_ini, "HoraFin": h_fin,
                    "Aula": normalizar_aula(aula),
                    "Materia": nombre, "Codigo": codigo, "Seccion": seccion,
                    "Ocupada": ocupada
                })

        return pd.DataFrame(filas)
    except Exception as e:
        st.error(f"Error cargando horario: {e}")
        return None

# Cargar datos
df_horario = cargar_horario()
df_reservas = cargar_reservas()

# ============================================================================
# LÓGICA
# ============================================================================

def obtener_bloques_dia(instalacion, fecha, df_horario, df_reservas):
    """Obtiene bloques de clases y reservas para una instalación"""
    dia_semana = DIA_SEMANA.get(fecha.weekday(), "")
    bloques = []
    reservas = []

    # Clases
    if df_horario is not None:
        df_inst = df_horario[
            (df_horario["Aula"] == instalacion) &
            (df_horario["Dia"] == dia_semana)
        ].sort_values("HoraInicio", na_position='last')

        for _, row in df_inst.iterrows():
            if row["Ocupada"] and row["HoraInicio"] is not None:
                bloques.append({
                    "hora": row["Hora"],
                    "tipo": "clase",
                    "detalle": f"{row['Codigo']} {row['Materia']} — Sección {row['Seccion']}",
                    "h_ini": row["HoraInicio"],
                    "h_fin": row["HoraFin"]
                })
            elif not row["Ocupada"] and row["HoraInicio"] is not None:
                bloques.append({
                    "hora": row["Hora"],
                    "tipo": "libre",
                    "detalle": "Disponible",
                    "h_ini": row["HoraInicio"],
                    "h_fin": row["HoraFin"]
                })

    # Reservas
    if df_reservas is not None and "fecha_date" in df_reservas.columns:
        c_inst = next((c for c in df_reservas.columns if c == "instalacion"), None)
        c_nom = next((c for c in df_reservas.columns if c == "nombre"), None)
        c_act = next((c for c in df_reservas.columns if c == "actividad"), None)

        if c_inst:
            filtradas = df_reservas[
                (df_reservas[c_inst].astype(str).str.strip() == instalacion) &
                (df_reservas["fecha_date"] == fecha)
            ]

            for _, row in filtradas.iterrows():
                ini_r = row.get("hora_inicio_t")
                fin_r = row.get("hora_fin_t")

                if ini_r is None or fin_r is None:
                    continue

                nom = str(row[c_nom]) if c_nom else "—"
                act = str(row[c_act]) if c_act else ""
                hora_exacta = f"{ini_r.strftime('%H:%M')} – {fin_r.strftime('%H:%M')}"
                detalle = f"{hora_exacta} — {nom}"

                if act and act not in ("nan", ""):
                    detalle += f" — {act}"

                reservas.append({
                    "hora": hora_exacta,
                    "tipo": "reserva",
                    "detalle": detalle,
                    "h_ini": ini_r,
                    "h_fin": fin_r
                })

    bloques.sort(key=lambda b: b["h_ini"] or time(0, 0))
    reservas.sort(key=lambda r: r["h_ini"] or time(0, 0))
    return bloques, reservas

# ============================================================================
# RENDERIZADO SVG
# ============================================================================

def generar_svg_calendario(instalacion, fecha, bloques, reservas):
    """Genera SVG calendario vertical"""
    inicio_dia = datetime.strptime("06:00", "%H:%M").time()
    fin_dia = datetime.strptime("20:10", "%H:%M").time()

    inicio_min = inicio_dia.hour * 60 + inicio_dia.minute
    fin_min = fin_dia.hour * 60 + fin_dia.minute

    altura_total = (fin_min - inicio_min) / 60 * 25
    ancho = 600

    svg_lines = []
    svg_lines.append(f'<svg width="100%" viewBox="0 0 {ancho + 100} {altura_total + 150}" xmlns="http://www.w3.org/2000/svg">')
    svg_lines.append(f'<style>.time-label {{ font-size: 11px; fill: #666; text-anchor: end; }} .bloque-text {{ font-size: 11px; font-weight: bold; }} .bloque-sub {{ font-size: 9px; }}</style>')

    # Título
    dia_nombre = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"][fecha.weekday()]
    svg_lines.append(f'<text x="40" y="25" style="font-size: 16px; font-weight: bold; fill: #1a1a1a;">{instalacion} — {dia_nombre} {fecha.strftime("%d/%m/%Y")}</text>')

    y_base = 60

    # Líneas de horas
    for hora in range(6, 21):
        minutos_desde_inicio = (hora * 60) - inicio_min
        y_pos = y_base + (minutos_desde_inicio / 60) * 25
        svg_lines.append(f'<text x="35" y="{y_pos + 4}" class="time-label">{hora:02d}:00</text>')
        svg_lines.append(f'<line x1="40" y1="{y_pos}" x2="{ancho + 40}" y2="{y_pos}" stroke="#e5e5e5" stroke-width="0.5"/>')

    y_fin = y_base + altura_total
    svg_lines.append(f'<line x1="40" y1="{y_fin}" x2="{ancho + 40}" y2="{y_fin}" stroke="#999" stroke-width="1"/>')

    # Clases (rojo)
    for bloque in bloques:
        if bloque["tipo"] == "clase":
            h_ini = bloque["h_ini"]
            h_fin = bloque["h_fin"]
            ini_min_b = h_ini.hour * 60 + h_ini.minute
            fin_min_b = h_fin.hour * 60 + h_fin.minute

            y_inicio = y_base + ((ini_min_b - inicio_min) / 60) * 25
            altura = ((fin_min_b - ini_min_b) / 60) * 25

            svg_lines.append(f'<rect x="50" y="{y_inicio}" width="{ancho - 20}" height="{altura}" fill="#ffebee" stroke="#ef5350" stroke-width="2" rx="4"/>')
            svg_lines.append(f'<text x="60" y="{y_inicio + 14}" class="bloque-text" style="fill: #c62828;">🔴 {bloque["detalle"][:25]}</text>')
            svg_lines.append(f'<text x="60" y="{y_inicio + altura - 5}" class="bloque-sub" style="fill: #c62828;">{bloque["hora"]}</text>')

    # Reservas (amarillo)
    for reserva in reservas:
        h_ini = reserva["h_ini"]
        h_fin = reserva["h_fin"]
        ini_min = h_ini.hour * 60 + h_ini.minute
        fin_min_r = h_fin.hour * 60 + h_fin.minute

        y_inicio = y_base + ((ini_min - inicio_min) / 60) * 25
        altura = ((fin_min_r - ini_min) / 60) * 25

        svg_lines.append(f'<rect x="50" y="{y_inicio}" width="{ancho - 20}" height="{altura}" fill="#fffde7" stroke="#fdd835" stroke-width="2" rx="4"/>')
        svg_lines.append(f'<text x="60" y="{y_inicio + 14}" class="bloque-text" style="fill: #f57f17;">🟡 {reserva["detalle"][:30]}</text>')
        svg_lines.append(f'<text x="60" y="{y_inicio + altura - 5}" class="bloque-sub" style="fill: #f57f17;">{reserva["hora"]}</text>')

    # Disponibles (verde)
    # Calcular espacios libres
    eventos = sorted(
        [(b["h_ini"], b["h_fin"]) for b in bloques if b["tipo"] == "clase"] +
        [(r["h_ini"], r["h_fin"]) for r in reservas],
        key=lambda x: x[0]
    )

    cursor = inicio_dia
    disponibles = []

    for evento_ini, evento_fin in eventos:
        if cursor < evento_ini:
            disponibles.append((cursor, evento_ini))
        cursor = max(cursor, evento_fin)

    if cursor < fin_dia:
        disponibles.append((cursor, fin_dia))

    for disp_ini, disp_fin in disponibles:
        ini_min_d = disp_ini.hour * 60 + disp_ini.minute
        fin_min_d = disp_fin.hour * 60 + disp_fin.minute

        y_inicio = y_base + ((ini_min_d - inicio_min) / 60) * 25
        altura = ((fin_min_d - ini_min_d) / 60) * 25
        duracion_min = fin_min_d - ini_min_d
        duracion_str = f"{duracion_min // 60}h {duracion_min % 60}m" if duracion_min >= 60 else f"{duracion_min}m"

        svg_lines.append(f'<rect x="50" y="{y_inicio}" width="{ancho - 20}" height="{altura}" fill="#e8f5e9" stroke="#66bb6a" stroke-width="2" rx="4"/>')
        svg_lines.append(f'<text x="60" y="{y_inicio + 16}" class="bloque-text" style="fill: #2e7d32; font-size: 12px;">✅ Disponible</text>')
        svg_lines.append(f'<text x="60" y="{y_inicio + altura - 5}" class="bloque-sub" style="fill: #2e7d32; font-weight: bold;">{duracion_str}</text>')

    svg_lines.append('</svg>')
    return "\n".join(svg_lines)

# ============================================================================
# INTERFAZ
# ============================================================================

col_icon, col_title = st.columns([1, 9])
with col_icon:
    st.markdown("<div style='font-size:2.8rem;padding-top:0.3rem'>📅</div>", unsafe_allow_html=True)
with col_title:
    st.markdown("<h1 class='titulo'>Disponibilidad de Espacios UPES</h1>", unsafe_allow_html=True)

# Estado
with st.sidebar:
    st.markdown("## 📡 Estado")
    st.markdown("---")

    if df_horario is not None:
        aulas = sorted(set(df_horario["Aula"].unique().tolist()) | ESPACIOS_ADICIONALES)
        st.success(f"✅ Horario cargado\n{len(aulas)} aulas")
    else:
        st.error("❌ Error horario")
        aulas = []

    if df_reservas is not None:
        st.success(f"✅ Reservas cargadas\n{len(df_reservas)} registros")
    else:
        st.warning("⚠️ No se pudieron cargar reservas")
        if not SHEETS_URL:
            st.error("❌ SHEETS_URL no en secrets")

# Filtros
c1, c2 = st.columns(2)
with c1:
    if aulas:
        instalacion = st.selectbox("Instalación", aulas)
    else:
        instalacion = "SUM"

with c2:
    fecha = st.date_input("Fecha", value=date.today())

# Bloques
if df_horario is not None:
    bloques, reservas_dia = obtener_bloques_dia(instalacion, fecha, df_horario, df_reservas)

    # Métricas
    col1, col2, col3 = st.columns(3)
    libres = sum(1 for b in bloques if b["tipo"] == "libre")
    clases = sum(1 for b in bloques if b["tipo"] == "clase")

    col1.metric("✅ Libres", libres)
    col2.metric("🔴 Clases", clases)
    col3.metric("🟡 Reservas", len(reservas_dia))

    # Calendario
    st.subheader(f"Calendario: {instalacion}")
    svg_html = generar_svg_calendario(instalacion, fecha, bloques, reservas_dia)
    st.write(svg_html, unsafe_allow_html=True)
else:
    st.error("No se pudo cargar el horario")
