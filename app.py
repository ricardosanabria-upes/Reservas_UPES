"""
Sistema de Reserva de Instalaciones
Replica exacta del formulario Google Forms institucional.
Valida contra:
  1. Excel DETALLE_AULAS_CICLO_ACTUAL.xlsx (clases del ciclo)
  2. Hoja de respuestas del formulario (reservas ya registradas)
"""

import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime, date, time, timedelta
from pathlib import Path

st.set_page_config(
    page_title="Reserva de Instalaciones",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }

    .titulo { font-size: 2rem; font-weight: 700; color: #1e1b4b; margin-bottom: 0; }
    .sub    { color: #64748b; font-size: 0.95rem; margin-bottom: 1rem; }
    .badge  { display:inline-block; background:#ede9fe; color:#4c1d95;
              border-radius:999px; padding:2px 14px; font-size:0.78rem;
              font-weight:600; margin-bottom:1rem; }
    .seccion { font-size:1.05rem; font-weight:700; color:#1e1b4b;
               border-left:4px solid #7c3aed; padding-left:10px;
               margin:1.5rem 0 0.8rem 0; }
    .aviso  { background:#fffbeb; border:1px solid #fcd34d; border-radius:10px;
               padding:0.9rem 1.1rem; font-size:0.88rem; color:#78350f; margin-bottom:1rem; }
    .libre  { background:#f0fdf4; border:1px solid #86efac; border-radius:8px;
               padding:8px 12px; margin:3px 0; font-size:0.83rem; color:#166534; }
    .clase  { background:#fef2f2; border:1px solid #fca5a5; border-radius:8px;
               padding:8px 12px; margin:3px 0; font-size:0.83rem; color:#991b1b; }
    .reserva{ background:#fefce8; border:1px solid #fde047; border-radius:8px;
               padding:8px 12px; margin:3px 0; font-size:0.83rem; color:#713f12; }
    .ok-box { background:#f0fdf4; border:1px solid #86efac; border-radius:12px;
               padding:1.2rem 1.5rem; margin-top:1rem; }
    .err-box{ background:#fef2f2; border:1px solid #fca5a5; border-radius:12px;
               padding:1rem 1.2rem; margin-top:0.5rem; }
    div[data-testid="stMetric"] { background:#f8fafc; border:1px solid #e2e8f0;
               border-radius:10px; padding:0.75rem 1rem; }
    .stButton>button { background:#4c1d95; color:white; border:none; border-radius:8px;
               font-weight:600; padding:0.65rem 1.5rem; font-size:0.97rem; width:100%; }
    .stButton>button:hover { background:#3b0764; }
    .stTextInput>div>div>input, .stTextArea>div>textarea {
        border-radius: 8px !important; }
</style>
""", unsafe_allow_html=True)


# ─── INSTALACIONES (igual al menú del formulario real) ────────────────────────
INSTALACIONES_AULAS = [
    "A-11", "A-12", "A-13", "A-14", "A-15", "A-16",
    "A-21 C/Acondicionado", "A-22 C/Acondicionado",
    "A-31", "A-32", "A-33", "A-34 (Mesas de dibujo)", "A-35", "A-36",
    "A-41", "A-42", "A-43", "A-44", "A-45", "A-46",
    "SUM", "Sala de juntas",
]
INSTALACIONES_OTROS = ["Pasillos", "Biblioteca"]

TODAS_INSTALACIONES = (
    ["── Aulas y salones ──"] + INSTALACIONES_AULAS +
    ["── Otros espacios ──"] + INSTALACIONES_OTROS
)

# Días de semana (Python weekday: 0=lunes … 6=domingo)
DIA_SEMANA = {
    0: "1.Lunes", 1: "2.Martes", 2: "3.Miercoles",
    3: "4.Jueves", 4: "5.Viernes", 5: "6.Sabado", 6: "7.Domingo",
}


# ─── FUNCIONES ────────────────────────────────────────────────────────────────

@st.cache_data
def procesar_horario(contenido_bytes: bytes) -> pd.DataFrame:
    df_raw = pd.read_excel(io.BytesIO(contenido_bytes))
    df_raw.columns = df_raw.columns.str.strip()
    df_raw["Dia"]  = df_raw["Dia"].ffill()
    df_raw["Hora"] = df_raw["Hora"].ffill()

    # Normalizar nombres de aula del Excel al formato del formulario
    # Ej: "A-21" → "A-21 C/Acondicionado" si aplica
    aulas_raw = [c for c in df_raw.columns if c not in ["Dia", "Hora"]]
    filas = []

    for _, row in df_raw.iterrows():
        dia  = str(row["Dia"]).strip()
        hora = str(row["Hora"]).strip()
        if not dia or dia == "nan" or not hora or hora == "nan":
            continue
        try:
            partes_h = hora.replace("–", "-").split("-")
            h_ini = datetime.strptime(partes_h[0].strip(), "%H:%M").time()
            h_fin = datetime.strptime(partes_h[1].strip(), "%H:%M").time()
        except Exception:
            h_ini = h_fin = None

        for aula in aulas_raw:
            val = row[aula]
            ocupada = not (pd.isna(val) or str(val).strip() == "")
            if ocupada:
                texto  = str(val).strip()
                partes = texto.split()
                codigo  = partes[0] if partes else ""
                seccion = partes[-1] if len(partes) > 1 and partes[-1].isdigit() else ""
                nombre  = " ".join(partes[1:-1]) if seccion else " ".join(partes[1:])
            else:
                codigo = seccion = nombre = ""

            filas.append({
                "Dia": dia, "Hora": hora,
                "HoraInicio": h_ini, "HoraFin": h_fin,
                "Aula": aula, "AulaNorm": normalizar_aula(aula),
                "Materia": nombre, "Codigo": codigo, "Seccion": seccion,
                "Ocupada": ocupada
            })
    return pd.DataFrame(filas)


def normalizar_aula(aula: str) -> str:
    """Mapea nombres del Excel a nombres del formulario."""
    aula = aula.strip()
    mapeo = {
        "A-21": "A-21 C/Acondicionado",
        "A-22": "A-22 C/Acondicionado",
        "A-34": "A-34 (Mesas de dibujo)",
        "A-25-26": "A-25-26",
        "Área Básica": "Área Básica",
        "Centro de Cómputo": "Centro de Cómputo",
        "Centro de Innovación Tecnológi": "Centro de Innovación Tecnológica",
        "Escuela de Civil": "Escuela de Civil",
        "Escuela de Computación": "Escuela de Computación",
        "Escuela de Eléctrica": "Escuela de Eléctrica",
        "Escuela de Industrial": "Escuela de Industrial",
    }
    return mapeo.get(aula, aula)


@st.cache_data
def cargar_respuestas(contenido_bytes: bytes, extension: str) -> pd.DataFrame:
    if extension == ".csv":
        df = pd.read_csv(io.BytesIO(contenido_bytes))
    else:
        # El Excel exportado de Google Forms tiene una fila extra al inicio
        # Intentar con header=1 primero, si falla usar header=0
        try:
            df_test = pd.read_excel(io.BytesIO(contenido_bytes), header=1)
            # Verificar si la primera fila parece encabezado real
            if "Marca temporal" in df_test.columns or "Nombre completo" in str(df_test.columns.tolist()):
                df = df_test
            else:
                df = pd.read_excel(io.BytesIO(contenido_bytes), header=0)
        except Exception:
            df = pd.read_excel(io.BytesIO(contenido_bytes), header=0)

    df.columns = df.columns.str.strip()

    # Renombrar columnas largas del formulario real a nombres cortos
    rename = {}
    for col in df.columns:
        cl = col.lower().strip()
        if "instalación solicitada" in cl or "instalacion solicitada" in cl:
            rename[col] = "instalacion"
        elif "fecha del evento" in cl:
            rename[col] = "fecha"
        elif "hora de inicio" in cl:
            rename[col] = "hora_inicio"
        elif "hora de finalización" in cl or "hora de finalizacion" in cl or "finalización exacta" in cl:
            rename[col] = "hora_fin"
        elif "nombre completo del solicitante" in cl:
            rename[col] = "nombre"
        elif "nombre y descripción" in cl or "nombre y descripcion" in cl:
            rename[col] = "actividad"
        elif cl in ["día", "dia"]:
            rename[col] = "dia_semana"
        elif "total de horas" in cl:
            rename[col] = "total_horas"
        elif "detalles y requerimientos" in cl:
            rename[col] = "especificaciones"
        elif "instrucción técnica" in cl or "instruccion tecnica" in cl:
            rename[col] = "fechas_recurrentes"

    df = df.rename(columns=rename)
    return df


def detectar_col(df, opciones):
    # Buscar primero coincidencia exacta con nombres cortos renombrados
    for col in df.columns:
        if col in opciones:
            return col
    # Luego buscar por contenido parcial
    for col in df.columns:
        if any(o in col.lower() for o in opciones):
            return col
    return None


def hay_traslape(ini1: time, fin1: time, ini2: time, fin2: time) -> bool:
    return ini1 < fin2 and ini2 < fin1


def conflictos_excel(df_h: pd.DataFrame, instalacion: str,
                     dia_semana: str, h_ini: time, h_fin: time) -> list[dict]:
    """Bloques del Excel que chocan con la solicitud."""
    # Buscar por nombre normalizado
    df_inst = df_h[
        (df_h["AulaNorm"] == instalacion) &
        (df_h["Dia"] == dia_semana) &
        (df_h["Ocupada"] == True)
    ].dropna(subset=["HoraInicio", "HoraFin"])

    resultado = []
    for _, row in df_inst.iterrows():
        if hay_traslape(h_ini, h_fin, row["HoraInicio"], row["HoraFin"]):
            resultado.append({
                "Bloque": row["Hora"],
                "Materia": f"{row['Codigo']} {row['Materia']} (Sección {row['Seccion']})"
            })
    return resultado


def conflictos_reservas(df_r: pd.DataFrame, instalacion: str,
                         fecha: date, h_ini: time, h_fin: time) -> list[dict]:
    """
    Reservas existentes que chocan con la solicitud.
    Bloqueo total: sin importar quién reservó.
    """
    c_inst  = detectar_col(df_r, ["instalacion", "instalación", "aula", "salon", "espacio"])
    c_fecha = detectar_col(df_r, ["fecha", "date", "día del evento", "dia del evento"])
    c_ini   = detectar_col(df_r, ["hora_inicio", "hora de inicio", "hora inicio", "inicio"])
    c_fin   = detectar_col(df_r, ["hora_fin", "hora de fin", "hora fin", "finalización", "fin"])
    c_nom   = detectar_col(df_r, ["nombre", "nombre completo", "solicitante", "name"])

    if not (c_inst and c_fecha and c_ini and c_fin):
        return []

    resultado = []
    for _, row in df_r.iterrows():
        if str(row[c_inst]).strip() != instalacion:
            continue
        try:
            fecha_r = pd.to_datetime(str(row[c_fecha]), dayfirst=True).date()
        except Exception:
            continue
        if fecha_r != fecha:
            continue
        try:
            ini_r = datetime.strptime(str(row[c_ini]).strip()[:5], "%H:%M").time()
            fin_r = datetime.strptime(str(row[c_fin]).strip()[:5], "%H:%M").time()
        except Exception:
            continue
        if hay_traslape(h_ini, h_fin, ini_r, fin_r):
            quien = str(row[c_nom]) if c_nom else "otro solicitante"
            resultado.append({
                "Reservado por": quien,
                "Horario": f"{ini_r.strftime('%H:%M')} – {fin_r.strftime('%H:%M')}"
            })
    return resultado


def parse_hora(texto: str):
    """Parsea HH:MM desde texto."""
    try:
        return datetime.strptime(texto.strip(), "%H:%M").time()
    except Exception:
        return None


# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Archivos")
    st.markdown("---")

    excel_file = st.file_uploader(
        "📊 Horario del ciclo actual",
        type=["xlsx", "xls"],
        help="DETALLE_AULAS_CICLO_ACTUAL.xlsx — sube el nuevo cada ciclo"
    )
    ciclo_nombre = st.text_input("🗓 Ciclo", value="Ciclo I - 2025")

    st.markdown("---")
    st.markdown("**📋 Hoja de respuestas del formulario**")
    resp_file = st.file_uploader(
        "CSV o Excel exportado de Google Forms",
        type=["csv", "xlsx", "xls"],
        help="Las respuestas ya registradas para detectar choques"
    )
    st.markdown("---")
    st.markdown(
        "<small style='color:#94a3b8'>"
        "Cada ciclo sube el nuevo Excel.<br>"
        "La hoja de respuestas se actualiza<br>"
        "conforme llegan nuevas solicitudes."
        "</small>", unsafe_allow_html=True
    )


# ─── CARGAR DATOS ─────────────────────────────────────────────────────────────
df_horario  = None
df_resp     = None
aulas_excel = []

if excel_file:
    df_horario  = procesar_horario(excel_file.read())
    aulas_excel = sorted(df_horario["AulaNorm"].unique().tolist())

if resp_file:
    ext = Path(resp_file.name).suffix.lower()
    df_resp = cargar_respuestas(resp_file.read(), ext)
    if df_resp is not None:
        df_resp.columns = df_resp.columns.str.strip()


# ─── HEADER ──────────────────────────────────────────────────────────────────
cl, ch = st.columns([1, 9])
with cl:
    st.markdown("<div style='font-size:2.8rem;padding-top:0.3rem'>🏫</div>", unsafe_allow_html=True)
with ch:
    st.markdown("<h1 class='titulo'>Solicitud de Reserva de Instalaciones</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub'>Por favor, responda a todas las preguntas de forma exhaustiva para el éxito de su evento.</p>", unsafe_allow_html=True)

if df_horario is not None:
    st.markdown(f"<span class='badge'>📅 {ciclo_nombre}</span>", unsafe_allow_html=True)

st.markdown("---")

tab_form, tab_disp, tab_resp_tab, tab_ayuda = st.tabs([
    "📝 Solicitar Reserva", "🔍 Consultar Disponibilidad",
    "📋 Reservas Registradas", "❓ Ayuda"
])


# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 1 – FORMULARIO
# ═══════════════════════════════════════════════════════════════════════════════
with tab_form:

    # ── Aviso de revisión previa (igual al formulario real) ───────────────────
    st.markdown(
        "<div class='aviso'>"
        "📌 <b>REVISAR:</b> Antes de proceder, le invitamos a consultar la pestaña "
        "<b>Consultar Disponibilidad</b> para confirmar que el espacio y horario "
        "deseado esté disponible antes de enviar su solicitud."
        "</div>", unsafe_allow_html=True
    )

    # ── Sección 1: Solicitante ────────────────────────────────────────────────
    st.markdown("<div class='seccion'>👤 Datos del Solicitante</div>", unsafe_allow_html=True)
    nombre_sol = st.text_input("Nombre completo del Solicitante *", placeholder="Ej. María López Martínez")
    nombre_act = st.text_area("Nombre y descripción de la actividad *", height=80,
                               placeholder="Ej. Reunión de coordinación académica — revisión de plan de estudios 2025")

    # ── Sección 2: Fechas y Horario ───────────────────────────────────────────
    st.markdown("<div class='seccion'>📅 Fecha y Horario</div>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        fecha_ev = st.date_input(
            "Fecha del evento/actividad *",
            min_value=date.today(),
            value=date.today() + timedelta(days=1),
            help="En caso de evento recurrente, indique la primera fecha y detalle las demás abajo."
        )
    with c2:
        st.markdown("")  # espaciador

    fechas_recurrentes = st.text_input(
        "Detalles de fechas y días (para eventos recurrentes)",
        placeholder="Ej. del 2 al 11 de Marzo / Lunes 2, Miércoles 4, Lunes 9, Miércoles 11 de Marzo",
        help="Si la reserva contempla varios días o fechas específicas, especifíquelas aquí."
    )

    c3, c4, c5 = st.columns(3)
    with c3:
        hora_ini_txt = st.text_input(
            "Hora de inicio (formato 24h) *",
            placeholder="Ej. 14:00 o 08:00",
            help="Si son las 2:00 PM coloque 14:00. Para las 8:00 AM coloque 08:00."
        )
    with c4:
        hora_fin_txt = st.text_input(
            "Hora de finalización exacta (formato 24h) *",
            placeholder="Ej. 16:00 o 10:30",
            help="Si son las 2:00 PM coloque 14:00. Para las 8:00 AM coloque 08:00."
        )
    with c5:
        total_horas = st.text_input(
            "Total de horas de la reserva *",
            placeholder="Ej. 2 horas",
        )

    # Parsear horas
    hora_ini = parse_hora(hora_ini_txt) if hora_ini_txt.strip() else None
    hora_fin = parse_hora(hora_fin_txt) if hora_fin_txt.strip() else None

    # ── Sección 3: Instalación ────────────────────────────────────────────────
    st.markdown("<div class='seccion'>🏛 Instalación Solicitada</div>", unsafe_allow_html=True)
    st.markdown(
        "<small style='color:#64748b'>Para esto debe haber revisado la disponibilidad "
        "y respetar las reservas de Clases y eventos.</small>",
        unsafe_allow_html=True
    )

    instalacion = st.selectbox(
        "Instalación solicitada *",
        TODAS_INSTALACIONES,
        format_func=lambda x: x
    )
    # Si seleccionó separador, no es válido
    inst_valida = not instalacion.startswith("──")

    # Preview de disponibilidad en tiempo real
    if inst_valida and hora_ini and hora_fin and df_horario is not None:
        dia_semana = DIA_SEMANA.get(fecha_ev.weekday(), "")
        dia_nombre = dia_semana.split(".")[1] if "." in dia_semana else str(fecha_ev)

        if hora_ini >= hora_fin:
            st.warning("⚠️ La hora de fin debe ser posterior a la hora de inicio.")
        else:
            cx_excel    = []
            cx_reservas = []

            if instalacion in aulas_excel:
                cx_excel = conflictos_excel(df_horario, instalacion, dia_semana, hora_ini, hora_fin)

            if df_resp is not None:
                cx_reservas = conflictos_reservas(df_resp, instalacion, fecha_ev, hora_ini, hora_fin)

            if cx_excel or cx_reservas:
                st.markdown("**⚠️ Se detectaron conflictos para esta selección:**")
                for c in cx_excel:
                    st.markdown(
                        f"<div class='clase'>🔴 <b>Clase del ciclo:</b> {c['Bloque']} — {c['Materia']}</div>",
                        unsafe_allow_html=True
                    )
                for r in cx_reservas:
                    st.markdown(
                        f"<div class='reserva'>🟡 <b>Ya reservada:</b> {r['Horario']} por {r['Reservado por']}</div>",
                        unsafe_allow_html=True
                    )
            else:
                st.markdown(
                    f"<div class='libre'>✅ <b>{instalacion}</b> disponible el "
                    f"<b>{fecha_ev.strftime('%d/%m/%Y')} ({dia_nombre})</b> "
                    f"de <b>{hora_ini_txt}</b> a <b>{hora_fin_txt}</b></div>",
                    unsafe_allow_html=True
                )

    # ── Sección 4: Especificaciones ───────────────────────────────────────────
    st.markdown("<div class='seccion'>📋 Especificaciones del Espacio</div>", unsafe_allow_html=True)
    st.markdown(
        "<small style='color:#64748b'>Indique la instalación requerida y los recursos necesarios. "
        "<b>Nota importante:</b> La adecuación del espacio se programa con el personal de logística "
        "inmediatamente después de recibir su solicitud.<br><br>"
        "Por favor detalle si requiere: <b>Mantelería, cafetera, número de tazas, PC, cañón (proyector) "
        "o una ubicación específica del mobiliario.</b></small>",
        unsafe_allow_html=True
    )
    especificaciones = st.text_area(
        "Especificaciones y apoyo requerido",
        height=110,
        placeholder=(
            "Ej. Requiero proyector (cañón), 30 sillas en semicírculo, "
            "1 mesa principal, cafetera con 20 tazas, mantelería blanca..."
        )
    )

    # Aviso de requerimientos técnicos
    st.markdown(
        "<div class='aviso'>"
        "🔧 <b>Requerimientos Técnicos y Equipo:</b> Si su evento requiere instalación de "
        "regletas, extensiones o equipo de sonido, por favor realice la solicitud adicional "
        "a la <b>Unidad de T.I.</b> por los canales correspondientes."
        "</div>", unsafe_allow_html=True
    )

    # ── Botón enviar ─────────────────────────────────────────────────────────
    st.markdown("")
    enviar = st.button("✅ Enviar Solicitud de Reserva")

    if enviar:
        errores = []

        # Campos obligatorios
        if not nombre_sol.strip():
            errores.append("El **nombre del solicitante** es obligatorio.")
        if not nombre_act.strip():
            errores.append("El **nombre y descripción de la actividad** es obligatorio.")
        if not hora_ini_txt.strip():
            errores.append("La **hora de inicio** es obligatoria.")
        elif not hora_ini:
            errores.append("La **hora de inicio** no tiene formato válido (use HH:MM, ej. 14:00).")
        if not hora_fin_txt.strip():
            errores.append("La **hora de finalización** es obligatoria.")
        elif not hora_fin:
            errores.append("La **hora de finalización** no tiene formato válido (use HH:MM, ej. 16:00).")
        if hora_ini and hora_fin and hora_ini >= hora_fin:
            errores.append("La **hora de fin** debe ser posterior a la hora de inicio.")
        if not total_horas.strip():
            errores.append("El **total de horas** es obligatorio.")
        if not inst_valida:
            errores.append("Selecciona una **instalación** válida.")

        # Validación 1: choque con clase del ciclo (Excel)
        if not errores and df_horario is not None and inst_valida and instalacion in aulas_excel:
            dia_semana = DIA_SEMANA.get(fecha_ev.weekday(), "")
            cx = conflictos_excel(df_horario, instalacion, dia_semana, hora_ini, hora_fin)
            for c in cx:
                errores.append(
                    f"🔴 **Choque con clase del ciclo:** {instalacion} tiene "
                    f"**{c['Materia']}** en el bloque {c['Bloque']} "
                    f"({dia_semana.split('.')[1] if '.' in dia_semana else dia_semana})."
                )

        # Validación 2: choque con reserva existente (bloqueo total)
        if not errores and df_resp is not None and inst_valida:
            cx_r = conflictos_reservas(df_resp, instalacion, fecha_ev, hora_ini, hora_fin)
            for r in cx_r:
                errores.append(
                    f"🟡 **Instalación ya reservada** en ese horario "
                    f"({r['Horario']}) por **{r['Reservado por']}**. "
                    f"Ninguna otra solicitud puede ocupar ese espacio."
                )

        # ── Mostrar resultado ─────────────────────────────────────────────────
        if errores:
            st.markdown("<div class='err-box'>", unsafe_allow_html=True)
            st.markdown("#### ⚠️ No se puede procesar la solicitud")
            for e in errores:
                st.markdown(f"- {e}")
            st.markdown("</div>", unsafe_allow_html=True)

        else:
            dia_semana_txt = DIA_SEMANA.get(fecha_ev.weekday(), "")
            dia_nombre = dia_semana_txt.split(".")[1] if "." in dia_semana_txt else ""

            st.markdown("<div class='ok-box'>", unsafe_allow_html=True)
            st.markdown("### 🎉 Solicitud validada y generada exitosamente")
            st.markdown(f"""
| Campo | Detalle |
|---|---|
| 👤 Solicitante | {nombre_sol.strip()} |
| 📌 Actividad | {nombre_act.strip()} |
| 📅 Fecha | {fecha_ev.strftime('%d/%m/%Y')} ({dia_nombre}) |
| 🔁 Fechas recurrentes | {fechas_recurrentes.strip() or '—'} |
| ⏰ Horario | {hora_ini_txt} – {hora_fin_txt} |
| ⌛ Total horas | {total_horas.strip()} |
| 🏛 Instalación | {instalacion} |
| 📋 Especificaciones | {especificaciones.strip() or '—'} |
| 📋 Ciclo | {ciclo_nombre} |
| 🕐 Generado | {datetime.now().strftime('%d/%m/%Y %H:%M')} |
""")
            st.markdown("</div>", unsafe_allow_html=True)

            # Descargar comprobante
            registro = pd.DataFrame([{
                "ciclo":               ciclo_nombre,
                "nombre":              nombre_sol.strip(),
                "actividad":           nombre_act.strip(),
                "fecha":               fecha_ev.strftime("%Y-%m-%d"),
                "dia_semana":          dia_nombre,
                "fechas_recurrentes":  fechas_recurrentes.strip(),
                "hora_inicio":         hora_ini_txt,
                "hora_fin":            hora_fin_txt,
                "total_horas":         total_horas.strip(),
                "instalacion":         instalacion,
                "especificaciones":    especificaciones.strip(),
                "fecha_solicitud":     datetime.now().strftime("%Y-%m-%d %H:%M"),
            }])
            st.download_button(
                "⬇️ Descargar comprobante CSV",
                data=registro.to_csv(index=False).encode("utf-8"),
                file_name=f"reserva_{instalacion.replace(' ','-')}_{fecha_ev}.csv",
                mime="text/csv"
            )


# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 2 – CONSULTAR DISPONIBILIDAD
# ═══════════════════════════════════════════════════════════════════════════════
with tab_disp:
    st.markdown("### 🔍 Consultar disponibilidad")

    inst_opciones = [i for i in TODAS_INSTALACIONES if not i.startswith("──")]
    c1, c2 = st.columns(2)
    with c1:
        inst_q  = st.selectbox("Instalación", inst_opciones, key="iq")
    with c2:
        fecha_q = st.date_input("Fecha", value=date.today() + timedelta(days=1), key="fq")

    dia_q      = DIA_SEMANA.get(fecha_q.weekday(), "")
    dia_q_nom  = dia_q.split(".")[1] if "." in dia_q else str(fecha_q)

    st.markdown(f"**{inst_q} — {dia_q_nom} {fecha_q.strftime('%d/%m/%Y')}**")

    # Clases del ciclo
    if df_horario is not None and inst_q in aulas_excel:
        df_inst = df_horario[
            (df_horario["AulaNorm"] == inst_q) &
            (df_horario["Dia"] == dia_q)
        ].sort_values("HoraInicio")

        st.markdown("**Horario del ciclo:**")
        if df_inst.empty:
            st.success("✅ Sin clases asignadas este día.")
        for _, row in df_inst.iterrows():
            if row["Ocupada"]:
                st.markdown(
                    f"<div class='clase'>🔴 {row['Hora']} — {row['Codigo']} {row['Materia']} (Sección {row['Seccion']})</div>",
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    f"<div class='libre'>✅ {row['Hora']} — Libre</div>",
                    unsafe_allow_html=True
                )
    elif df_horario is None:
        st.info("Sube el Excel del ciclo en el panel lateral para ver el horario de aulas.")

    # Reservas registradas para esa fecha
    if df_resp is not None:
        st.markdown("**Reservas registradas:**")
        c_inst  = detectar_col(df_resp, ["instalacion", "instalación", "aula", "salon", "espacio"])
        c_fecha = detectar_col(df_resp, ["fecha", "date", "día del evento", "dia del evento"])
        c_ini   = detectar_col(df_resp, ["hora_inicio", "hora de inicio", "hora inicio", "inicio"])
        c_fin   = detectar_col(df_resp, ["hora_fin", "hora de fin", "hora fin", "finalización", "fin"])
        c_nom   = detectar_col(df_resp, ["nombre", "nombre completo", "solicitante", "name"])

        if c_inst and c_fecha:
            try:
                filtradas = df_resp[
                    (df_resp[c_inst].astype(str).str.strip() == inst_q) &
                    (pd.to_datetime(df_resp[c_fecha], dayfirst=True, errors="coerce").dt.date == fecha_q)
                ]
            except Exception:
                filtradas = pd.DataFrame()

            if filtradas.empty:
                st.success("✅ Sin reservas registradas para esta fecha.")
            else:
                for _, row in filtradas.iterrows():
                    ini = str(row[c_ini]) if c_ini else "?"
                    fin = str(row[c_fin]) if c_fin else "?"
                    nom = str(row[c_nom]) if c_nom else "—"
                    st.markdown(
                        f"<div class='reserva'>🟡 {ini} – {fin} — {nom}</div>",
                        unsafe_allow_html=True
                    )
        else:
            st.info("La hoja de respuestas necesita columnas de instalación y fecha.")
    else:
        st.info("Sube la hoja de respuestas en el panel lateral para ver reservas registradas.")


# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 3 – RESERVAS REGISTRADAS
# ═══════════════════════════════════════════════════════════════════════════════
with tab_resp_tab:
    if df_resp is None:
        st.info("Sube la hoja de respuestas exportada de Google Forms en el panel lateral.")
    else:
        st.markdown(f"### 📋 Reservas registradas — {len(df_resp)} registros")
        busq = st.text_input("🔍 Buscar", placeholder="Nombre, instalación, fecha...")
        drv  = df_resp.copy()
        if busq:
            mask = drv.apply(lambda r: r.astype(str).str.contains(busq, case=False, na=False).any(), axis=1)
            drv  = drv[mask]
        st.dataframe(drv, use_container_width=True, height=400)
        st.caption(f"Mostrando {len(drv)} de {len(df_resp)} registros")

        # Validación cruzada
        if df_horario is not None:
            st.markdown("---")
            st.markdown("#### 🔄 Validación cruzada con el horario del ciclo actual")
            c_inst  = detectar_col(df_resp, ["instalacion", "instalación", "aula", "salon"])
            c_fecha = detectar_col(df_resp, ["fecha", "date", "día del evento"])
            c_ini   = detectar_col(df_resp, ["hora_inicio", "hora de inicio", "hora inicio", "inicio"])
            c_fin   = detectar_col(df_resp, ["hora_fin", "hora de fin", "hora fin", "finalización", "fin"])
            c_nom   = detectar_col(df_resp, ["nombre", "nombre completo", "solicitante"])

            if c_inst and c_fecha and c_ini and c_fin:
                choque = []
                for _, row in df_resp.iterrows():
                    inst_r = str(row[c_inst]).strip()
                    if inst_r not in aulas_excel:
                        continue
                    try:
                        fecha_r = pd.to_datetime(str(row[c_fecha]), dayfirst=True).date()
                        dia_r   = DIA_SEMANA.get(fecha_r.weekday(), "")
                        ini_r   = datetime.strptime(str(row[c_ini]).strip()[:5], "%H:%M").time()
                        fin_r   = datetime.strptime(str(row[c_fin]).strip()[:5], "%H:%M").time()
                    except Exception:
                        continue
                    cx = conflictos_excel(df_horario, inst_r, dia_r, ini_r, fin_r)
                    for c in cx:
                        choque.append({
                            "Solicitante":  str(row[c_nom]) if c_nom else "—",
                            "Instalación":  inst_r,
                            "Fecha":        fecha_r.strftime("%d/%m/%Y"),
                            "Horario":      f"{ini_r.strftime('%H:%M')}–{fin_r.strftime('%H:%M')}",
                            "Choca con":    c["Materia"],
                        })
                if choque:
                    st.warning(f"⚠️ {len(choque)} reserva(s) colisionan con clases del ciclo actual:")
                    st.dataframe(pd.DataFrame(choque), use_container_width=True)
                else:
                    st.success("✅ Ninguna reserva conflictúa con el horario del ciclo actual.")
            else:
                st.info("La hoja necesita columnas: instalación, fecha, hora inicio, hora fin.")


# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 4 – AYUDA
# ═══════════════════════════════════════════════════════════════════════════════
with tab_ayuda:
    st.markdown("""
### 📝 Campos del formulario

| Campo | Obligatorio | Descripción |
|---|---|---|
| Nombre del solicitante | ✅ | Nombre completo |
| Nombre y descripción de la actividad | ✅ | Qué es el evento |
| Fecha del evento | ✅ | Fecha exacta (se convierte a día de semana automáticamente) |
| Detalles de fechas recurrentes | — | Para eventos en varios días |
| Hora de inicio | ✅ | Formato 24h — ej. 14:00 |
| Hora de finalización | ✅ | Formato 24h — ej. 16:00 |
| Total de horas | ✅ | Ej. 2 horas |
| Instalación | ✅ | Menú desplegable con aulas y otros espacios |
| Especificaciones | — | Proyector, sillas, mantelería, etc. |

---

### 🔴🟡✅ Lógica de validación

**🔴 Choque con clase del ciclo**
La fecha se convierte al día de semana y se cruza con el Excel.
Si hay una materia asignada en ese bloque horario → bloqueado.

**🟡 Choque con reserva existente**
Si alguien ya reservó esa instalación en esa fecha y el horario se traslapa
→ bloqueado para **todos**, sin importar quién solicite.

**✅ Disponible**
Ni el Excel ni las reservas registradas tienen conflicto.

---

### 📋 Columnas esperadas en la hoja de respuestas (Google Forms)

| Dato | Nombres de columna aceptados |
|---|---|
| Instalación | `instalación`, `instalacion`, `aula`, `salon`, `espacio` |
| Fecha | `fecha`, `date`, `día del evento`, `dia del evento` |
| Hora inicio | `hora de inicio`, `hora inicio`, `inicio`, `hora_inicio` |
| Hora fin | `hora de fin`, `hora fin`, `finalización`, `fin`, `hora_fin` |
| Nombre | `nombre completo`, `nombre`, `solicitante`, `name` |

---

### 🔄 Flujo por ciclo

1. Al inicio del ciclo sube el nuevo **DETALLE_AULAS_CICLO_ACTUAL.xlsx** — sin configuración adicional
2. La app detecta automáticamente todas las aulas y horarios
3. La hoja de respuestas se actualiza exportándola de Google Forms en CSV/Excel
""")
