import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import io
import os

# URLs - intenta secrets primero, luego variables de entorno
try:
    GITHUB_EXCEL_URL = st.secrets["github_excel_url"]
except Exception:
    GITHUB_EXCEL_URL = os.environ.get("GITHUB_EXCEL_URL", "https://raw.githubusercontent.com/ricardosanabria-upes/Consulta_Disponibilidad_UPES/main/DETALLE%20AULAS%20CICLO%20ACTUAL.xlsx")

try:
    SHEETS_URL = st.secrets["sheets_url"]
except Exception:
    SHEETS_URL = os.environ.get("SHEETS_URL", "")

@st.cache_data(ttl=300)
def cargar_horario():
    try:
        response = requests.get(GITHUB_EXCEL_URL)
        response.raise_for_status()
        df = pd.read_excel(io.BytesIO(response.content))
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Error cargando horario: {e}")
        return None

@st.cache_data(ttl=300)
def cargar_reservas():
    if not SHEETS_URL:
        return None
    try:
        df = pd.read_csv(SHEETS_URL, header=1)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Error cargando reservas: {e}")
        return None

def normalizar_dia(dia_str):
    """Convierte '1.Lunes' -> 'LUNES'"""
    dia_str = str(dia_str).strip()
    # Si tiene el formato "N.Dia", extrae la parte después del punto
    if "." in dia_str:
        dia_str = dia_str.split(".", 1)[1]
    # Mapea nombres en español a nombres en mayúsculas
    dias_map = {
        "lunes": "LUNES",
        "martes": "MARTES",
        "miércoles": "MIÉRCOLES",
        "miercoles": "MIÉRCOLES",
        "jueves": "JUEVES",
        "viernes": "VIERNES",
        "sábado": "SÁBADO",
        "sabado": "SÁBADO",
        "domingo": "DOMINGO",
    }
    return dias_map.get(dia_str.lower(), dia_str.upper())

def obtener_bloques_dia(df_horario, df_reservas, aula, fecha):
    fecha_str = fecha.strftime("%d/%m/%Y")
    dia_nombre = fecha.strftime("%A").upper()

    # Traducir nombres de días de Python
    dias_es = {
        "MONDAY": "LUNES",
        "TUESDAY": "MARTES",
        "WEDNESDAY": "MIÉRCOLES",
        "THURSDAY": "JUEVES",
        "FRIDAY": "VIERNES",
        "SATURDAY": "SÁBADO",
        "SUNDAY": "DOMINGO"
    }
    dia_nombre_es = dias_es.get(dia_nombre, dia_nombre)

    clases = []
    if df_horario is not None:
        # Buscar filas que coincidan con el día (después de normalizar)
        for idx, row in df_horario.iterrows():
            dia_normalizado = normalizar_dia(row["Dia"])
            if dia_normalizado == dia_nombre_es and aula in df_horario.columns:
                valor = row[aula]
                if pd.notna(valor) and str(valor).strip() != "" and str(valor).strip() != "None":
                    clases.append({
                        "type": "clase",
                        "hora": row["Hora"],
                        "contenido": str(valor)
                    })

    reservas = []
    if df_reservas is not None:
        try:
            df_reservas_aula = df_reservas[
                (df_reservas.iloc[:, 0].astype(str).str.strip() == aula) &
                (df_reservas.iloc[:, 1].astype(str).str.strip() == fecha_str)
            ]
            for idx, row in df_reservas_aula.iterrows():
                reservas.append({
                    "type": "reserva",
                    "hora": row.iloc[2],
                    "contenido": row.iloc[3] if len(row) > 3 else "Reserva"
                })
        except:
            pass

    return clases + reservas

def generar_svg_calendario(bloques, aula, fecha):
    HORAS = [f"{h:02d}:00" for h in range(6, 21)]
    ANCHO = 700
    ALTO_FILA = 50
    MARGIN_LEFT = 100
    MARGIN_TOP = 80

    # Colores
    COLOR_CLASE = "#FF6B6B"
    COLOR_RESERVA = "#FFD93D"

    svg_content = f"""
    <svg width="{ANCHO + 120}" height="{len(HORAS) * ALTO_FILA + MARGIN_TOP + 40}" xmlns="http://www.w3.org/2000/svg">
        <style>
            .hora-label {{ font-size: 12px; font-weight: bold; color: #333; }}
            .bloque-text {{ font-size: 12px; font-weight: 500; fill: #000; text-anchor: start; }}
            .titulo {{ font-size: 16px; font-weight: bold; color: #1e1b4b; }}
            .linea-hora {{ stroke: #e5e7eb; stroke-width: 1; }}
            rect {{ filter: drop-shadow(0 1px 3px rgba(0,0,0,0.1)); }}
        </style>

        <!-- Fondo alterno -->
    """

    for i, hora in enumerate(HORAS):
        if i % 2 == 0:
            y = MARGIN_TOP + i * ALTO_FILA
            svg_content += f'<rect x="0" y="{y}" width="{ANCHO + MARGIN_LEFT + 20}" height="{ALTO_FILA}" fill="#f9fafb"/>\n'

    # Título
    svg_content += f'<text x="10" y="35" class="titulo">{aula} - {fecha.strftime("%d/%m/%Y")}</text>\n'

    # Horas y líneas de cuadrícula
    for i, hora in enumerate(HORAS):
        y = MARGIN_TOP + i * ALTO_FILA
        svg_content += f'<text x="10" y="{y + 35}" class="hora-label">{hora}</text>\n'
        svg_content += f'<line x1="{MARGIN_LEFT}" y1="{y}" x2="{ANCHO + MARGIN_LEFT}" y2="{y}" class="linea-hora"/>\n'

    # Renderizar bloques (clases y reservas)
    for bloque in bloques:
        hora_str = str(bloque["hora"]).strip()
        try:
            # Extrae la hora inicial del rango (ej: "06:00-07:40" -> "06:00")
            if "-" in hora_str:
                hora_inicio = hora_str.split("-")[0].strip()
            else:
                hora_inicio = hora_str

            # Convierte "06:00" a índice (6-6=0, 7-6=1, etc.)
            hora_parts = hora_inicio.split(":")
            hora_idx = int(hora_parts[0]) - 6

            if 0 <= hora_idx < len(HORAS):
                y = MARGIN_TOP + hora_idx * ALTO_FILA + 8

                if bloque["type"] == "clase":
                    color = COLOR_CLASE
                    emoji = "🔴"
                else:
                    color = COLOR_RESERVA
                    emoji = "🟡"

                svg_content += f'<rect x="{MARGIN_LEFT}" y="{y}" width="550" height="38" fill="{color}" rx="4" opacity="0.85"/>\n'
                svg_content += f'<text x="{MARGIN_LEFT + 12}" y="{y + 24}" class="bloque-text">{emoji} {bloque["contenido"][:60]}</text>\n'
        except Exception as e:
            pass

    svg_content += "</svg>"
    return svg_content

def main():
    st.set_page_config(page_title="Disponibilidad UPES", layout="wide")
    st.title("📅 Consulta de Disponibilidad - UPES")

    # Cargar datos
    df_horario = cargar_horario()
    df_reservas = cargar_reservas()

    with st.sidebar:
        st.header("⚙️ Configuración")

        # Estado de conexión
        if df_horario is not None:
            st.success("✅ Horario cargado desde GitHub")
        else:
            st.error("❌ Error al cargar horario")
            return

        if df_reservas is not None:
            st.success("✅ Reservas cargadas desde Google Sheets")
        else:
            st.warning("⚠️ Reservas no configuradas")

        st.divider()

        if df_horario is not None:
            aulas = sorted([col for col in df_horario.columns if col not in ["Dia", "Hora"]])
            aula_seleccionada = st.selectbox("Selecciona un aula:", aulas)
        else:
            st.error("No hay aulas disponibles")
            return

        fecha_seleccionada = st.date_input("Selecciona una fecha:", datetime.now())

        if st.button("🔄 Actualizar datos"):
            st.cache_data.clear()
            st.rerun()

    # Obtener bloques del día
    bloques = obtener_bloques_dia(df_horario, df_reservas, aula_seleccionada, fecha_seleccionada)

    # Calcular métricas
    clases = [b for b in bloques if b["type"] == "clase"]
    reservas_list = [b for b in bloques if b["type"] == "reserva"]
    horas_totales = 15  # 6:00 a 21:00
    horas_libres = horas_totales - len(clases) - len(reservas_list)

    # Mostrar métricas
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Horas Libres", horas_libres)
    with col2:
        st.metric("Clases", len(clases))
    with col3:
        st.metric("Reservas", len(reservas_list))
    with col4:
        ocupacion = round((len(clases) + len(reservas_list)) / horas_totales * 100) if horas_totales > 0 else 0
        st.metric("Ocupación", f"{ocupacion}%")

    st.divider()

    # Mostrar calendario
    svg_html = generar_svg_calendario(bloques, aula_seleccionada, fecha_seleccionada)
    st.html(svg_html)

    # Leyenda
    st.subheader("Leyenda")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("🔴 **Clase programada**")
    with col2:
        st.markdown("🟡 **Reserva**")

    # Detalles
    if clases or reservas_list:
        st.subheader("Detalles")
        if clases:
            st.write("**Clases:**")
            for clase in clases:
                st.text(f"  {clase['hora']} - {clase['contenido']}")
        if reservas_list:
            st.write("**Reservas:**")
            for reserva in reservas_list:
                st.text(f"  {reserva['hora']} - {reserva['contenido']}")
    else:
        st.info("No hay clases ni reservas para esta fecha y aula")

if __name__ == "__main__":
    main()
