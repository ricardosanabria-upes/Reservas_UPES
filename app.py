import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import io

# URLs (reemplaza con las tuyas)
GITHUB_EXCEL_URL = "https://raw.githubusercontent.com/ricardosanabria-upes/Consulta_Disponibilidad_UPES/main/DETALLE%20AULAS%20CICLO%20ACTUAL.xlsx"
SHEETS_URL = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/export?format=csv&gid=0"

@st.cache_data(ttl=300)
def cargar_horario():
    try:
        response = requests.get(GITHUB_EXCEL_URL)
        response.raise_for_status()
        df = pd.read_excel(io.BytesIO(response.content))
        return df
    except Exception as e:
        st.error(f"Error cargando horario: {e}")
        return None

@st.cache_data(ttl=300)
def cargar_reservas():
    try:
        df = pd.read_csv(SHEETS_URL, header=1)
        return df
    except Exception as e:
        st.error(f"Error cargando reservas: {e}")
        return None

def obtener_bloques_dia(df_horario, df_reservas, aula, fecha):
    fecha_str = fecha.strftime("%d/%m/%Y")
    dia_nombre = fecha.strftime("%A").upper()

    # Traducir nombres de días
    dias_es = {
        "MONDAY": "LUNES",
        "TUESDAY": "MARTES",
        "WEDNESDAY": "MIÉRCOLES",
        "THURSDAY": "JUEVES",
        "FRIDAY": "VIERNES",
        "SATURDAY": "SÁBADO",
        "SUNDAY": "DOMINGO"
    }
    dia_nombre = dias_es.get(dia_nombre, dia_nombre)

    clases = []
    if df_horario is not None:
        df_dia = df_horario[df_horario["Dia"] == dia_nombre]
        if not df_dia.empty and aula in df_horario.columns:
            for idx, row in df_dia.iterrows():
                if pd.notna(row[aula]) and str(row[aula]).strip() != "":
                    clases.append({
                        "type": "clase",
                        "hora": row["Hora"],
                        "contenido": str(row[aula])
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
    HORAS = [f"{h:02d}:00" for h in range(7, 20)]
    ANCHO = 600
    ALTO_FILA = 40
    MARGIN_LEFT = 80
    MARGIN_TOP = 60

    # Colores
    COLOR_CLASE = "#FF6B6B"
    COLOR_RESERVA = "#FFD93D"

    svg_content = f"""
    <svg width="{ANCHO + 100}" height="{len(HORAS) * ALTO_FILA + MARGIN_TOP + 40}" xmlns="http://www.w3.org/2000/svg">
        <style>
            .hora-label {{ font-size: 12px; font-weight: bold; }}
            .bloque-text {{ font-size: 11px; font-weight: bold; fill: #000; text-anchor: start; }}
            .titulo {{ font-size: 14px; font-weight: bold; }}
            .linea-hora {{ stroke: #e0e0e0; stroke-width: 1; }}
        </style>

        <!-- Título -->
        <text x="10" y="30" class="titulo">{aula} - {fecha.strftime('%d/%m/%Y')}</text>

        <!-- Horas y líneas de cuadrícula -->
    """

    for i, hora in enumerate(HORAS):
        y = MARGIN_TOP + i * ALTO_FILA
        svg_content += f'<text x="10" y="{y + 25}" class="hora-label">{hora}</text>\n'
        svg_content += f'<line x1="{MARGIN_LEFT}" y1="{y}" x2="{ANCHO + MARGIN_LEFT}" y2="{y}" class="linea-hora"/>\n'

    # Renderizar bloques (clases y reservas)
    for bloque in bloques:
        hora_str = str(bloque["hora"]).strip()
        try:
            if ":" in hora_str:
                hora_parts = hora_str.split(":")
                hora_idx = int(hora_parts[0]) - 7
            else:
                hora_idx = int(hora_str) - 7

            if 0 <= hora_idx < len(HORAS):
                y = MARGIN_TOP + hora_idx * ALTO_FILA + 5

                if bloque["type"] == "clase":
                    color = COLOR_CLASE
                    emoji = "🔴"
                else:
                    color = COLOR_RESERVA
                    emoji = "🟡"

                svg_content += f'<rect x="{MARGIN_LEFT}" y="{y}" width="450" height="30" fill="{color}" rx="3"/>\n'
                svg_content += f'<text x="{MARGIN_LEFT + 10}" y="{y + 20}" class="bloque-text">{emoji} {bloque["contenido"]}</text>\n'
        except:
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

        if df_horario is not None:
            aulas = [col for col in df_horario.columns if col not in ["Dia", "Hora"]]
            aula_seleccionada = st.selectbox("Selecciona un aula:", aulas)
        else:
            st.error("No se pudo cargar la lista de aulas")
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
    horas_totales = 13  # 7:00 a 19:00
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
        st.metric("Ocupación", f"{round((len(clases) + len(reservas_list)) / horas_totales * 100)}%")

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

if __name__ == "__main__":
    main()
