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
        st.dataframe(hojas[hoja_sel], use_container_width=True)
    except Exception as e:
        st.warning(f"No pude previsualizar el archivo ({e}), pero sí puedes descargarlo.")

    with open(ruta, "rb") as f:
        st.download_button("⬇️ Descargar Excel", f, file_name=archivo_sel)
