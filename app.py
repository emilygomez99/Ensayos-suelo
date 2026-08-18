import streamlit as st
from pathlib import Path
import pandas as pd

# Carpeta raíz donde viven todas las subcarpetas del árbol (Hincado > ... > Excel)
DATA_DIR = Path(__file__).parent / "data"

st.set_page_config(page_title="Ensayos en Suelo a Escala", layout="wide")
st.title("🧪 Ensayos en Suelo a Escala – Acero Corrugado")

if not DATA_DIR.exists():
    st.error(f"No encuentro la carpeta '{DATA_DIR.name}'. Crea una carpeta 'data' junto a app.py "
              "y pon ahí tu árbol de subcarpetas con los Excel.")
    st.stop()


def listar_subcarpetas(path: Path):
    return sorted([p for p in path.iterdir() if p.is_dir()], key=lambda p: p.name)


def listar_archivos(path: Path):
    return sorted(
        [p for p in path.iterdir() if p.suffix.lower() in (".xlsx", ".xls")],
        key=lambda p: p.name,
    )


current = DATA_DIR
breadcrumb = []

# Va bajando de nivel en nivel mientras existan subcarpetas
nivel = 0
while True:
    subfolders = listar_subcarpetas(current)
    if not subfolders:
        break
    etiqueta = "Categoría" if nivel == 0 else current.name.replace("_", " ")
    opciones = [f.name.replace("_", " ") for f in subfolders]
    seleccion = st.selectbox(f"Nivel {nivel + 1} — {etiqueta}", opciones, key=f"nivel_{nivel}")
    # recupera la carpeta real correspondiente a la opción mostrada
    current = subfolders[opciones.index(seleccion)]
    breadcrumb.append(seleccion)
    nivel += 1

if breadcrumb:
    st.caption(" ➜ ".join(breadcrumb))

archivos = listar_archivos(current)

if not archivos:
    st.info("No hay archivos Excel en esta rama del árbol todavía.")
else:
    nombres = [f.name for f in archivos]
    archivo_sel = st.selectbox("Archivo", nombres)
    ruta = archivos[nombres.index(archivo_sel)]

    try:
        hojas = pd.read_excel(ruta, sheet_name=None)
        hoja_sel = st.selectbox("Hoja", list(hojas.keys())) if len(hojas) > 1 else list(hojas.keys())[0]
        df = hojas[hoja_sel]
        st.dataframe(df, use_container_width=True)

        # --- Sección de gráficas ---
        columnas_numericas = df.select_dtypes(include="number").columns.tolist()
        todas_columnas = df.columns.tolist()

        if len(todas_columnas) >= 2:
            st.markdown("### 📈 Graficar")
            col1, col2, col3 = st.columns(3)
            with col1:
                eje_x = st.selectbox("Eje X", todas_columnas, key="eje_x")
            with col2:
                opciones_y = [c for c in columnas_numericas if c != eje_x] or columnas_numericas
                eje_y = st.selectbox("Eje Y", opciones_y, key="eje_y")
            with col3:
                tipo = st.selectbox("Tipo de gráfica", ["Línea", "Dispersión", "Barras"], key="tipo_grafica")

            datos_grafica = df[[eje_x, eje_y]].dropna().sort_values(by=eje_x)

            if tipo == "Línea":
                st.line_chart(datos_grafica, x=eje_x, y=eje_y)
            elif tipo == "Dispersión":
                st.scatter_chart(datos_grafica, x=eje_x, y=eje_y)
            else:
                st.bar_chart(datos_grafica, x=eje_x, y=eje_y)
        else:
            st.caption("Esta hoja no tiene suficientes columnas para graficar.")

    except Exception as e:
        st.warning(f"No pude previsualizar el archivo ({e}), pero sí puedes descargarlo.")

    with open(ruta, "rb") as f:
        st.download_button("⬇️ Descargar Excel", f, file_name=archivo_sel)
