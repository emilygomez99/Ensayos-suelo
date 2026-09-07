import streamlit as st
from pathlib import Path
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.drawing.image import Image as XLImage
import matplotlib.pyplot as plt

# Carpeta raíz donde viven todas las subcarpetas del árbol (Hincado > ... > Excel)
DATA_DIR = Path(__file__).parent / "data"
st.set_page_config(page_title="Ensayos en Suelo a Escala", layout="wide")
st.title("🧪 Ensayos en Suelo a Escala – Acero Corrugado")

ESQUEMA_PATH = Path(__file__).parent / "esquema_ensayos.png"
if ESQUEMA_PATH.exists():
    with st.expander("🗺️ Ver esquema completo del árbol de ensayos (guía de navegación)", expanded=False):
        ancho_esquema = st.slider(
            "Tamaño del esquema", min_value=300, max_value=2000, value=700, step=50,
            key="ancho_esquema",
        )
        st.image(str(ESQUEMA_PATH), width=ancho_esquema)
        st.caption(
            "Usa este esquema como referencia: los selectores de abajo recorren estas mismas ramas "
            "(Mortero / Acero Corrugado → Hincado / Arrancamiento / Normal / Volcamiento → bulbos → diámetro → longitud)."
        )

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


# --- Funciones para la gráfica combinada suavizada de varios ensayos ---

def _leer_carga_desplazamiento(
    archivo,
    hoja: str = "Hoja1",
    col_carga: int = 5,          # columna E
    col_desplazamiento: int = 7,  # columna G
):
    """
    Lee un archivo .xlsx (ruta o bytes) y devuelve (desplazamiento, carga)
    como arrays de numpy, ordenados por desplazamiento y sin celdas vacías.
    """
    if hasattr(archivo, "read"):
        contenido = archivo.read()
        wb = openpyxl.load_workbook(io.BytesIO(contenido), data_only=True)
    else:
        wb = openpyxl.load_workbook(archivo, data_only=True)

    if hoja not in wb.sheetnames:
        raise ValueError(f"La hoja '{hoja}' no existe en este archivo.")

    ws = wb[hoja]

    desplazamiento, carga = [], []
    for fila in range(1, ws.max_row + 1):
        d = ws.cell(row=fila, column=col_desplazamiento).value
        c = ws.cell(row=fila, column=col_carga).value
        if d is None or c is None:
            continue
        # Ignora filas de encabezado o cualquier celda que no sea numérica
        try:
            d_num = float(d)
            c_num = float(c)
        except (TypeError, ValueError):
            continue
        desplazamiento.append(d_num)
        carga.append(c_num)

    if not desplazamiento:
        return np.array([]), np.array([])

    desplazamiento = np.array(desplazamiento, dtype=float)
    carga = np.array(carga, dtype=float)

    orden = np.argsort(desplazamiento)
    return desplazamiento[orden], carga[orden]


def _obtener_hojas(archivo):
    """Devuelve la lista de nombres de hoja reales de un archivo .xlsx."""
    wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
    hojas = wb.sheetnames
    wb.close()
    return hojas


def _obtener_columnas_con_encabezado(archivo, hoja: str, max_col: int = 30):
    """
    Devuelve una lista de opciones 'LETRA - encabezado' para las columnas
    de una hoja que tengan algún valor en las primeras filas, junto con
    un diccionario que mapea cada opción a su letra de columna.
    """
    wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
    ws = wb[hoja]

    opciones = []
    letra_por_opcion = {}
    max_c = min(ws.max_column or 1, max_col)
    for col_idx in range(1, max_c + 1):
        letra = openpyxl.utils.get_column_letter(col_idx)
        encabezado = None
        for fila in range(1, min(ws.max_row or 1, 5) + 1):
            valor = ws.cell(row=fila, column=col_idx).value
            if valor is not None:
                encabezado = str(valor)
                break
        etiqueta = f"{letra} - {encabezado}" if encabezado else letra
        opciones.append(etiqueta)
        letra_por_opcion[etiqueta] = letra

    wb.close()
    return opciones, letra_por_opcion


def _promedio_movil(y: np.ndarray, ventana: int = 5) -> np.ndarray:
    """Suaviza una serie con un promedio móvil centrado de tamaño `ventana`."""
    if ventana <= 1:
        return y
    kernel = np.ones(ventana) / ventana
    y_pad = np.pad(y, (ventana // 2, ventana // 2), mode="edge")
    return np.convolve(y_pad, kernel, mode="valid")[: len(y)]


def graficar_ensayos_combinados(
    archivos,
    nombres=None,
    ventana_suavizado: int = 5,
    hoja: str = "Hoja1",
    col_carga: int = 5,
    col_desplazamiento: int = 7,
    titulo: str = "Carga vs. Desplazamiento",
    xlabel: str = "Desplazamiento (mm)",
    ylabel: str = "Carga (kN)",
):
    """
    Genera una figura de matplotlib con las curvas suavizadas de N ensayos
    superpuestas. Retorna (fig, errores) donde errores es una lista de
    (nombre, mensaje) para los archivos que no se pudieron leer.
    """
    if nombres is None:
        nombres = [f"Ensayo {i+1}" for i in range(len(archivos))]

    colores = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#4a3aa7"]

    fig, ax = plt.subplots(figsize=(9, 5.5))
    errores = []

    for i, (archivo, nombre) in enumerate(zip(archivos, nombres)):
        try:
            x, y = _leer_carga_desplazamiento(
                archivo, hoja=hoja, col_carga=col_carga, col_desplazamiento=col_desplazamiento
            )
            if len(x) == 0:
                errores.append((nombre, "No se encontraron datos válidos."))
                continue
            y_suave = _promedio_movil(y, ventana=ventana_suavizado)
            ax.plot(x, y_suave, label=nombre, color=colores[i % len(colores)], linewidth=2)
        except Exception as e:
            errores.append((nombre, str(e)))

    ax.set_title(titulo)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.grid(True, color="#e1e0d9", linewidth=0.8)
    ax.legend(frameon=False)
    fig.tight_layout()

    return fig, errores


# --- Navegación por árbol de carpetas ---

current = DATA_DIR
breadcrumb = []

nivel = 0
while True:
    subfolders = listar_subcarpetas(current)
    if not subfolders:
        break
    etiqueta = "Categoría" if nivel == 0 else current.name.replace("_", " ")
    opciones = [f.name.replace("_", " ") for f in subfolders]
    seleccion = st.selectbox(f"Nivel {nivel + 1} — {etiqueta}", opciones, key=f"nivel_{nivel}")
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

    # --- Vista de un solo archivo (comportamiento original) ---
    archivo_sel = st.selectbox("Archivo", nombres)
    ruta = archivos[nombres.index(archivo_sel)]

    try:
        hojas = pd.read_excel(ruta, sheet_name=None)
        hoja_sel = st.selectbox("Hoja", list(hojas.keys())) if len(hojas) > 1 else list(hojas.keys())[0]
        df = hojas[hoja_sel]
        st.dataframe(df, use_container_width=True)

        # --- Sección de gráficas (un solo archivo) ---
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

    # --- Comparar varios ensayos en una sola gráfica suavizada ---
    st.markdown("---")
    st.markdown("### 🔗 Comparar varios ensayos (curva suavizada)")
    st.caption(
        "Selecciona 2 o más archivos de esta misma carpeta para superponerlos en una sola "
        "gráfica de Carga vs. Desplazamiento, igual a la hoja 'Hoja1' de cada Excel."
    )

    archivos_comparar = st.multiselect(
        "Archivos a comparar",
        nombres,
        default=nombres[: min(3, len(nombres))],
        key="archivos_comparar",
    )

    ventana = st.slider("Suavizado (ventana del promedio móvil)", 1, 15, 5, key="ventana_suavizado")

    # --- Configuración individual (hoja + columnas) por cada archivo seleccionado ---
    config_por_archivo = {}

    if archivos_comparar:
        st.markdown("#### ⚙️ Configura cada archivo")
        for nombre_archivo in archivos_comparar:
            ruta_archivo = archivos[nombres.index(nombre_archivo)]

            with st.expander(f"📄 {nombre_archivo}", expanded=True):
                try:
                    hojas_disponibles = _obtener_hojas(ruta_archivo)
                except Exception as e:
                    st.warning(f"No pude leer las hojas de este archivo: {e}")
                    continue

                hoja_sel = st.selectbox(
                    "Hoja con los datos",
                    hojas_disponibles,
                    key=f"hoja_{nombre_archivo}",
                )

                try:
                    opciones_col, letra_por_opcion = _obtener_columnas_con_encabezado(
                        ruta_archivo, hoja_sel
                    )
                except Exception as e:
                    st.warning(f"No pude leer las columnas de la hoja '{hoja_sel}': {e}")
                    continue

                col_x, col_y = st.columns(2)
                with col_x:
                    opcion_carga = st.selectbox(
                        "Columna de carga (kN)",
                        opciones_col,
                        key=f"col_carga_{nombre_archivo}",
                    )
                with col_y:
                    opcion_desp = st.selectbox(
                        "Columna de desplazamiento (mm)",
                        opciones_col,
                        key=f"col_desp_{nombre_archivo}",
                    )

                config_por_archivo[nombre_archivo] = {
                    "ruta": ruta_archivo,
                    "hoja": hoja_sel,
                    "col_carga": letra_por_opcion[opcion_carga],
                    "col_desplazamiento": letra_por_opcion[opcion_desp],
                }

    if st.button("Generar gráfica combinada", key="btn_combinar"):
        if len(archivos_comparar) < 2:
            st.warning("Selecciona al menos 2 archivos.")
        else:
            colores = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#4a3aa7"]
            fig, ax = plt.subplots(figsize=(9, 5.5))
            errores = []
            datos_por_ensayo = {}

            for i, nombre_archivo in enumerate(archivos_comparar):
                cfg = config_por_archivo.get(nombre_archivo)
                if cfg is None:
                    errores.append((nombre_archivo, "No se pudo configurar este archivo."))
                    continue
                try:
                    col_carga_idx = openpyxl.utils.column_index_from_string(cfg["col_carga"])
                    col_desp_idx = openpyxl.utils.column_index_from_string(cfg["col_desplazamiento"])
                    x, y = _leer_carga_desplazamiento(
                        cfg["ruta"],
                        hoja=cfg["hoja"],
                        col_carga=col_carga_idx,
                        col_desplazamiento=col_desp_idx,
                    )
                    if len(x) == 0:
                        errores.append((nombre_archivo, "No se encontraron datos válidos."))
                        continue
                    y_suave = _promedio_movil(y, ventana=ventana)
                    ax.plot(x, y_suave, label=nombre_archivo, color=colores[i % len(colores)], linewidth=2)
                    datos_por_ensayo[f"Desplazamiento_{nombre_archivo}"] = pd.Series(x)
                    datos_por_ensayo[f"Carga_suavizada_{nombre_archivo}"] = pd.Series(y_suave)
                except Exception as e:
                    errores.append((nombre_archivo, str(e)))

            ax.set_title("Carga vs. Desplazamiento")
            ax.set_xlabel("Desplazamiento (mm)")
            ax.set_ylabel("Carga (kN)")
            ax.grid(True, color="#e1e0d9", linewidth=0.8)
            ax.legend(frameon=False)
            fig.tight_layout()

            # Guardar una imagen del gráfico para incrustarla luego en el Excel
            img_buffer = io.BytesIO()
            fig.savefig(img_buffer, format="png", dpi=150, bbox_inches="tight")
            img_buffer.seek(0)

            # Persistimos todo en session_state para que no se pierda cuando
            # el usuario le dé clic al botón de descarga (eso recarga la app
            # y "Generar gráfica combinada" volvería a estar sin presionar).
            st.session_state["combinada_fig"] = fig
            st.session_state["combinada_errores"] = errores
            st.session_state["combinada_df"] = pd.DataFrame(datos_por_ensayo)
            st.session_state["combinada_img"] = img_buffer.getvalue()

    # --- Mostrar el resultado guardado (persiste entre reruns) ---
    if "combinada_fig" in st.session_state:
        st.pyplot(st.session_state["combinada_fig"])

        if st.session_state["combinada_errores"]:
            for nombre, msg in st.session_state["combinada_errores"]:
                st.warning(f"{nombre}: {msg}")

        df_export = st.session_state["combinada_df"]
        if not df_export.empty:
            st.markdown("#### 💾 Descargar datos de la gráfica combinada")
            nombre_archivo_salida = st.text_input(
                "Nombre del archivo Excel (con o sin .xlsx)",
                value="grafica_combinada.xlsx",
                key="nombre_excel_combinado",
            )
            if not nombre_archivo_salida.lower().endswith(".xlsx"):
                nombre_archivo_salida += ".xlsx"

            # Construir el Excel: datos crudos + la imagen del gráfico incrustada
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                df_export.to_excel(writer, index=False, sheet_name="Datos")
            excel_buffer.seek(0)

            wb_final = openpyxl.load_workbook(excel_buffer)
            ws_final = wb_final["Datos"]
            img_final = XLImage(io.BytesIO(st.session_state["combinada_img"]))
            col_img = openpyxl.utils.get_column_letter(len(df_export.columns) + 2)
            img_final.anchor = f"{col_img}2"
            ws_final.add_image(img_final)

            final_buffer = io.BytesIO()
            wb_final.save(final_buffer)
            final_buffer.seek(0)

            st.download_button(
                "⬇️ Descargar Excel",
                data=final_buffer,
                file_name=nombre_archivo_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="descargar_excel_combinado",
            )
            st.caption(
                "El archivo se descarga con el nombre que escribiste arriba. La **carpeta** de destino "
                "la decide tu navegador, no la app (por seguridad, ninguna app web puede elegir carpetas "
                "de tu computador). Si quieres que te pregunte dónde guardar cada vez, actívalo en tu "
                "navegador: en Chrome/Edge → Configuración → Descargas → 'Preguntar dónde guardar cada "
                "archivo antes de descargarlo'."
            )
