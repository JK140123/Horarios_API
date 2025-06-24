import streamlit as st
import pandas as pd
from processor.asignador import procesar_horarios
import openpyxl
import io
import numpy as np
import random

# --- Limpiar caché ---
st.cache_data.clear()

import streamlit as st
import pandas as pd
from processor.asignador import procesar_horarios
import openpyxl
import io
import base64

# --- Imagen de fondo local + overlay oscuro ---
def set_background_local(image_path, opacity=0.6):
    with open(image_path, "rb") as file:
        encoded = base64.b64encode(file.read()).decode()
    st.markdown(
        f"""
        <style>
            .stApp {{
                background-image: linear-gradient(rgba(0,0,0,{opacity}), rgba(0,0,0,{opacity})), 
                                  url("data:image/jpg;base64,{encoded}");
                background-size: cover;
                background-position: center;
                background-attachment: fixed;
            }}
        </style>
        """,
        unsafe_allow_html=True
    )

# --- Aplica el fondo personalizado ---
set_background_local("images/fondo.jpg", opacity=0.8)  # Puedes ajustar el nivel de oscuridad (0 a 1)

# --- Configuración general de la página ---
st.set_page_config(page_title="Asignador de Clases", layout="wide")

# --- Estilos globales con colores Pantone ---
st.markdown("""
    <style>
    /* Botón normal */
    .stButton>button, .stDownloadButton>button {
        background-color: #00205B;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 10px 20px;
        border: none;
        transition: background-color 0.3s ease;
    }

    /* Hover: Azul claro */
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #98B1D3 !important; /* Pantone 292 */
        color: black !important;
    }

    /* Active (clic sostenido) */
    .stButton>button:active, .stDownloadButton>button:active {
        background-color: #DCE5F3 !important; /* Pantone 2707 */
        color: black !important;
    }

    /* Opcional: efecto de foco (cuando se selecciona con tab) */
    .stButton>button:focus, .stDownloadButton>button:focus {
        outline: 2px solid #98B1D3;
    }
</style>

""", unsafe_allow_html=True)

# --- Título principal ---
st.markdown("<h1>📅 Asignador de Clases por Salón</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitle'>Sube un archivo Excel con la plantilla de clases para generar el cronograma y visualizar posibles reubicaciones.</p>", unsafe_allow_html=True)

# --- Resto del flujo original (resumido aquí) ---
archivo = st.file_uploader("📤 Sube el archivo Excel", type=["xlsx"])


# Paleta de colores exacta de asignador.py
colores_programa = {
    "medicina": "#FFFF00",    # Amarillo
    "enfermería": "#C9DAF8",  # Azul claro
    "fisioterapia": "#EAD1DC", # Rosa pálido
    "psicología": "#9370DB",  # Lila claro
    "educación continua": "#93C47D", # Verde
    "otros": "#93C47D"        # Verde (mismo que educación continua)
}

def normalizar(texto):
    if not isinstance(texto, str):
        return ""
    texto = texto.lower().strip()
    texto = texto.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    texto = texto.replace("ñ", "n")
    return texto

def obtener_color_programa(programa):
    programa = normalizar(programa)
    if "enfermeria" in programa:
        return colores_programa["enfermería"]
    elif "fisioterapia" in programa:
        return colores_programa["fisioterapia"]
    elif "psicologia" in programa:
        return colores_programa["psicología"]
    elif "educacion continua" in programa or "educacion" in programa:
        return colores_programa["educación continua"]
    elif "medicina" in programa:
        return colores_programa["medicina"]
    return colores_programa["otros"]

# --- Cache de vistas previas por hoja ---
@st.cache_data(show_spinner=False)
def generar_vista_previa(nombre_hoja, _workbook, df_original):
    hoja = _workbook[nombre_hoja]
    data = list(hoja.values)
    columnas = data[0]
    df_raw = pd.DataFrame(data[1:], columns=columnas)

    df_vista = df_raw.copy()
    columnas_horas = df_vista.columns[1:]
    
    # Diccionario para mapear asignatura -> programa -> color
    asignaturas_programas = {}
    
    for col in columnas_horas:
        i = 0
        while i < len(df_vista):
            valor = df_vista.at[i, col]
            if pd.notna(valor):
                contenido_actual = valor
                duracion_real = 1

                asignatura = None
                profesor = None

                if isinstance(valor, str) and " - " in valor:
                    partes = valor.split(" - ")
                    if len(partes) >= 2:
                        asignatura = partes[0].strip()
                        profesor = partes[1].strip()

                if asignatura and profesor:
                    hora_inicio_str = df_vista.iloc[i, 0]
                    if isinstance(hora_inicio_str, str) and ":" in hora_inicio_str:
                        try:
                            hora_inicio_int = int(hora_inicio_str.split(":")[0])
                        except:
                            hora_inicio_int = None
                    else:
                        hora_inicio_int = None

                    fila_clase = df_original[
                        df_original["Asignatura"].str.strip() == asignatura
                        ]
                    fila_clase = fila_clase[
                        fila_clase["Profesor"].str.strip() == profesor
                    ]
                    
                    # Guardar relación asignatura-programa
                    if not fila_clase.empty:
                        programa = fila_clase.iloc[0]["Programa"]
                        if pd.notna(programa):
                            asignaturas_programas[asignatura] = programa

                    if hora_inicio_int is not None and not fila_clase.empty:
                        # Buscar coincidencia exacta con la hora
                        for _, row in fila_clase.iterrows():
                            h_ini = row["Hora de inicio"]
                            if pd.notna(h_ini) and int(str(h_ini).split(":")[0]) == hora_inicio_int:
                                h_fin = row["Hora de finalización"]
                                if pd.notna(h_fin):
                                    try:
                                        h_fin_int = int(str(h_fin).split(":")[0])
                                        duracion_real = h_fin_int - hora_inicio_int
                                        break
                                    except:
                                        pass

                for j in range(1, duracion_real):
                    if i + j < len(df_vista):
                        if pd.isna(df_vista.at[i + j, col]):
                            df_vista.at[i + j, col] = contenido_actual
                        else:
                            break
                i += duracion_real
            else:
                i += 1
    
    # Crear mapeo de asignatura a color basado en el programa
    colores_asignaturas = {}
    for asignatura, programa in asignaturas_programas.items():
        colores_asignaturas[asignatura] = obtener_color_programa(programa)
    
    return df_vista, colores_asignaturas

# --- Procesamiento principal ---
if archivo:
    try:
        df = pd.read_excel(archivo, sheet_name="Plantilla")
        st.success("Archivo cargado correctamente.")

        if st.button("🚀 Generar Horarios"):
            with st.spinner("Procesando..."):
                excel_bytes = procesar_horarios(df)
                st.session_state["excel_bytes"] = excel_bytes
                st.session_state["workbook"] = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
                st.session_state["procesado"] = True
                st.success("Archivo procesado exitosamente.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")

# --- Vista previa con colores y descarga ---
if st.session_state.get("procesado", False):
    workbook = st.session_state["workbook"]
    hojas = workbook.sheetnames
    df_original = pd.read_excel(archivo, sheet_name="Plantilla")

    hoja_seleccionada = st.selectbox("👀 Selecciona una hoja para vista previa", hojas)
    df_vista, colores_asignaturas = generar_vista_previa(hoja_seleccionada, workbook, df_original)

    st.subheader(f"📄 Vista previa: {hoja_seleccionada}")
    
    # Función para aplicar estilos de color según la paleta
    def color_celda(val):
        if pd.isna(val):
            return ''
        
        asignatura = None
        if isinstance(val, str) and " - " in val:
            partes = val.split(" - ")
            if len(partes) >= 2:
                asignatura = partes[0].strip()
        
        if asignatura and asignatura in colores_asignaturas:
            color = colores_asignaturas[asignatura]
            return f'background-color: {color}; color: black; border: 1px solid #ddd;'
        return ''
    
    # Aplicar estilos al DataFrame
    styled_df = df_vista.style.applymap(color_celda)
    
    # Mostrar el DataFrame con colores
    st.dataframe(styled_df, use_container_width=True, height=600)
    
    # Mostrar leyenda de colores con la paleta exacta
    st.subheader("🎨 Leyenda de Programas")
    
    programas_colores = {
        "Medicina": colores_programa["medicina"],
        "Enfermería": colores_programa["enfermería"],
        "Fisioterapia": colores_programa["fisioterapia"],
        "Psicología": colores_programa["psicología"],
        "Educación Continua": colores_programa["educación continua"],
        "Otros": colores_programa["otros"]
    }
    
    cols = st.columns(3)
    for i, (programa, color) in enumerate(programas_colores.items()):
        with cols[i % 3]:
            st.markdown(
                f"""<div style='background-color:{color}; color:black; padding:10px; border-radius:5px; 
                    margin-bottom:10px; text-align:center;'><strong>{programa}</strong></div>""",
                unsafe_allow_html=True
            )

    st.download_button(
        label="📥 Descargar Cronograma",
        data=st.session_state["excel_bytes"],
        file_name="Cronograma_Clases_Reubicadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )