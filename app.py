import streamlit as st
from pathlib import Path
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.drawing.image import Image as XLImage
import matplotlib.pyplot as plt
import requests
import base64
import json
import datetime

# Carpeta raíz donde viven todas las subcarpetas del árbol (Hincado > ... > Excel)
DATA_DIR = Path(__file__).parent / "data"
st.set_page_config(page_title="Ensayos en Suelo a Escala", layout="wide")
st.title("🧪 Ensayos en Suelo a Escala – Acero Corrugado")

# --- Configuración para guardar archivos directo en GitHub ---
GITHUB_REPO = "emilygomez99/Ensayos-suelo"  # owner/repo


def subir_archivo_a_github(path_en_repo: str, contenido_bytes: bytes, mensaje_commit: str):
    """
    Crea o actualiza un archivo en el repo de GitHub usando la API de contenidos.
    Requiere un token con permiso de escritura guardado en st.secrets['GITHUB_TOKEN'].
    """
    token = st.secrets.get("GITHUB_TOKEN") if hasattr(st, "secrets") else None
    if not token:
        raise RuntimeError(
            "No encontré GITHUB_TOKEN en los secretos de la app. "
            "Configúralo en Streamlit Cloud: Settings → Secrets (ver instrucciones)."
        )

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path_en_repo}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
    }

    # Si el archivo ya existe en esa ruta, GitHub exige mandar su "sha" para poder sobrescribirlo
    sha_existente = None
    resp_get = requests.get(api_url, headers=headers, timeout=15)
    if resp_get.status_code == 200:
        sha_existente = resp_get.json().get("sha")

    payload = {
        "message": mensaje_commit,
        "content": base64.b64encode(contenido_bytes).decode("utf-8"),
    }
    if sha_existente:
        payload["sha"] = sha_existente

    resp = requests.put(api_url, headers=headers, json=payload, timeout=20)
    if resp.status_code not in (200, 201):
        raise RuntimeError(f"GitHub respondió {resp.status_code}: {resp.text}")
    return resp.json()


# --- Historial de comparaciones (para poder comparar ensayos guardados en momentos distintos) ---
HISTORIAL_PATH_EN_REPO = "historial_comparaciones.json"


def cargar_historial():
    """Lee el historial de comparaciones guardadas desde el JSON en GitHub. Si no existe, devuelve lista vacía."""
    token = st.secrets.get("GITHUB_TOKEN") if hasattr(st, "secrets") else None
    if not token:
        return []
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{HISTORIAL_PATH_EN_REPO}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
    resp = requests.get(api_url, headers=headers, timeout=15)
    if resp.status_code != 200:
        return []
    contenido_b64 = resp.json().get("content", "")
    try:
        contenido = base64.b64decode(contenido_b64).decode("utf-8")
        return json.loads(contenido).get("entradas", [])
    except Exception:
        return []


def guardar_historial(entradas):
    """Sobrescribe el JSON de historial en GitHub con la lista completa de entradas."""
    contenido_json = json.dumps({"entradas": entradas}, ensure_ascii=False, indent=2).encode("utf-8")
    subir_archivo_a_github(
        HISTORIAL_PATH_EN_REPO, contenido_json, "Actualiza historial de comparaciones"
    )

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


def curva_promedio(series, n_puntos: int = 200):
    """
    Recibe una lista de series {"x":[...], "y":[...]} de una misma categoría,
    las interpola a una malla común de desplazamiento y devuelve (x, y) del
    promedio punto a punto. Si solo hay una serie, la devuelve tal cual.
    """
    series_validas = [s for s in series if len(s["x"]) > 1]
    if not series_validas:
        return [], []
    if len(series_validas) == 1:
        return series_validas[0]["x"], series_validas[0]["y"]

    x_min = max(min(s["x"]) for s in series_validas)
    x_max = min(max(s["x"]) for s in series_validas)
    if x_min >= x_max:  # los rangos no se solapan; usar el rango completo como respaldo
        x_min = min(min(s["x"]) for s in series_validas)
        x_max = max(max(s["x"]) for s in series_validas)

    malla = np.linspace(x_min, x_max, n_puntos)
    interpoladas = [np.interp(malla, s["x"], s["y"]) for s in series_validas]
    y_prom = np.mean(interpoladas, axis=0)
    return malla.tolist(), y_prom.tolist()


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

            st.markdown("#### ☁️ O guardarlo directo en tu repo de GitHub")
            carpeta_repo = st.text_input(
                "Carpeta dentro del repo (déjalo vacío para guardarlo en la raíz)",
                value="resultados",
                key="carpeta_github",
            )
            carpeta_repo = carpeta_repo.strip().strip("/")
            ruta_completa = f"{carpeta_repo}/{nombre_archivo_salida}" if carpeta_repo else nombre_archivo_salida
            st.caption(f"Se guardará en: `{ruta_completa}` dentro de `{GITHUB_REPO}`")

            if st.button("💾 Guardar en GitHub", key="btn_guardar_github"):
                try:
                    resultado = subir_archivo_a_github(
                        ruta_completa,
                        final_buffer.getvalue(),
                        f"Agrega {nombre_archivo_salida} desde la app",
                    )
                    url_archivo = resultado.get("content", {}).get("html_url", "")
                    st.success(f"¡Listo! Se guardó en `{ruta_completa}`.")
                    if url_archivo:
                        st.markdown(f"[Ver el archivo en GitHub]({url_archivo})")
                except Exception as e:
                    st.error(f"No pude guardar en GitHub: {e}")

            # --- Guardar esta comparación en el historial de la app ---
            st.markdown("#### 📌 Agregar esta comparación al historial (para comparar después)")
            etiqueta_default = " / ".join(breadcrumb) if breadcrumb else "Comparación"
            etiqueta_historial = st.text_input(
                "Etiqueta para reconocer esta comparación después",
                value=etiqueta_default,
                key="etiqueta_historial",
            )
            if st.button("➕ Agregar al historial", key="btn_agregar_historial"):
                try:
                    series = []
                    for col in df_export.columns:
                        if col.startswith("Desplazamiento_"):
                            nombre_ensayo = col[len("Desplazamiento_"):]
                            col_carga = f"Carga_suavizada_{nombre_ensayo}"
                            if col_carga in df_export.columns:
                                series.append({
                                    "nombre": nombre_ensayo,
                                    "x": df_export[col].dropna().tolist(),
                                    "y": df_export[col_carga].dropna().tolist(),
                                })
                    entradas = cargar_historial()
                    entradas.append({
                        "id": datetime.datetime.now().strftime("%Y%m%d%H%M%S"),
                        "etiqueta": etiqueta_historial,
                        "fecha": datetime.datetime.now().isoformat(timespec="seconds"),
                        "series": series,
                    })
                    guardar_historial(entradas)
                    st.success(f"¡Agregado! Ya puedes verlo en '📊 Historial de comparaciones' más abajo.")
                except Exception as e:
                    st.error(f"No pude guardar en el historial: {e}")

# --- Historial de comparaciones guardadas (siempre visible, sin importar en qué carpeta estés) ---
st.markdown("---")
st.markdown("## 📊 Historial de comparaciones")
st.caption(
    "Aquí se acumulan todas las comparaciones que hayas guardado con el botón "
    "'➕ Agregar al historial', sin importar de qué carpeta (diámetro/longitud/ubicación) vinieran. "
    "Elige varias para verlas juntas en una sola gráfica."
)

historial_entradas = cargar_historial()

if not historial_entradas:
    st.info("Todavía no has guardado ninguna comparación al historial.")
else:
    etiquetas_disponibles = [
        f"{e['etiqueta']} ({e['fecha'][:16].replace('T', ' ')})" for e in historial_entradas
    ]
    seleccion_historial = st.multiselect(
        "Comparaciones guardadas",
        etiquetas_disponibles,
        default=etiquetas_disponibles[-min(3, len(etiquetas_disponibles)):],
        key="seleccion_historial",
    )

    if seleccion_historial:
        colores_hist = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#4a3aa7",
                        "#8e44ad", "#16a085", "#c0392b", "#2c3e50", "#f39c12", "#7f8c8d"]
        estilos_linea = ["-", "--", ":", "-."]

        modo_representativo = st.checkbox(
            "🎯 Dejar solo una curva representativa por categoría (para sacar conclusiones)",
            key="modo_representativo",
        )

        curvas_a_graficar = []  # cada item: {"etiqueta":, "nombre":, "x":, "y":}

        if modo_representativo:
            st.caption(
                "Por defecto se calcula el **promedio** de los ensayos de cada categoría "
                "(interpolados a una malla común). Si prefieres una curva puntual en vez del "
                "promedio, elígela en el desplegable de esa categoría."
            )
            for etiqueta_completa in seleccion_historial:
                idx = etiquetas_disponibles.index(etiqueta_completa)
                entrada = historial_entradas[idx]
                opciones_rep = ["Promedio de todos"] + [s["nombre"] for s in entrada["series"]]
                eleccion = st.selectbox(
                    f"Curva representativa para «{entrada['etiqueta']}»",
                    opciones_rep,
                    key=f"rep_{idx}",
                )
                if eleccion == "Promedio de todos":
                    x_rep, y_rep = curva_promedio(entrada["series"])
                else:
                    serie_elegida = next(s for s in entrada["series"] if s["nombre"] == eleccion)
                    x_rep, y_rep = serie_elegida["x"], serie_elegida["y"]
                if x_rep:
                    curvas_a_graficar.append({
                        "etiqueta": entrada["etiqueta"], "nombre": eleccion, "x": x_rep, "y": y_rep,
                    })

            # Filtro final: de las representativas ya calculadas, cuáles dejar visibles
            etiquetas_calculadas = [c["etiqueta"] for c in curvas_a_graficar]
            seleccion_final = st.multiselect(
                "¿Cuáles representativas quieres dejar para tus conclusiones?",
                etiquetas_calculadas,
                default=etiquetas_calculadas,
                key="seleccion_final_representativas",
            )
            curvas_a_graficar = [c for c in curvas_a_graficar if c["etiqueta"] in seleccion_final]
        else:
            for etiqueta_completa in seleccion_historial:
                idx = etiquetas_disponibles.index(etiqueta_completa)
                entrada = historial_entradas[idx]
                for serie in entrada["series"]:
                    curvas_a_graficar.append({
                        "etiqueta": entrada["etiqueta"], "nombre": serie["nombre"],
                        "x": serie["x"], "y": serie["y"],
                    })

        if curvas_a_graficar:
            # Un color por categoría (mismo color aunque haya varias curvas en esa categoría)
            etiquetas_unicas = []
            for c in curvas_a_graficar:
                if c["etiqueta"] not in etiquetas_unicas:
                    etiquetas_unicas.append(c["etiqueta"])
            color_por_etiqueta = {
                et: colores_hist[i % len(colores_hist)] for i, et in enumerate(etiquetas_unicas)
            }

            fig_hist, ax_hist = plt.subplots(figsize=(10, 6))
            datos_hist_export = {}
            contador_por_etiqueta = {}

            for c in curvas_a_graficar:
                j = contador_por_etiqueta.get(c["etiqueta"], 0)
                linestyle = "-" if modo_representativo else estilos_linea[j % len(estilos_linea)]
                ax_hist.plot(
                    c["x"], c["y"],
                    label=f"{c['etiqueta']} · {c['nombre']}",
                    color=color_por_etiqueta[c["etiqueta"]],
                    linestyle=linestyle,
                    linewidth=2.5 if modo_representativo else 2,
                )
                contador_por_etiqueta[c["etiqueta"]] = j + 1

                col_base = f"{c['etiqueta']}_{c['nombre']}".replace(" ", "_")
                datos_hist_export[f"Desplazamiento_{col_base}"] = pd.Series(c["x"])
                datos_hist_export[f"Carga_{col_base}"] = pd.Series(c["y"])

            ax_hist.set_title("Comparación entre ensayos guardados")
            ax_hist.set_xlabel("Desplazamiento (mm)")
            ax_hist.set_ylabel("Carga (kN)")
            ax_hist.grid(True, color="#e1e0d9", linewidth=0.8)
            ax_hist.legend(frameon=False, fontsize=8)
            fig_hist.tight_layout()
            st.pyplot(fig_hist)

        # --- Descargar esta comparación del historial (Excel local o directo a GitHub) ---
        if curvas_a_graficar:
            img_hist_buffer = io.BytesIO()
            fig_hist.savefig(img_hist_buffer, format="png", dpi=150, bbox_inches="tight")
            img_hist_buffer.seek(0)

            df_hist_export = pd.DataFrame(datos_hist_export)

            st.markdown("#### 💾 Descargar esta comparación")
            nombre_excel_hist = st.text_input(
                "Nombre del archivo Excel (con o sin .xlsx)",
                value="comparacion_historial.xlsx",
                key="nombre_excel_historial",
            )
            if not nombre_excel_hist.lower().endswith(".xlsx"):
                nombre_excel_hist += ".xlsx"

            excel_hist_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_hist_buffer, engine="openpyxl") as writer:
                df_hist_export.to_excel(writer, index=False, sheet_name="Datos")
            excel_hist_buffer.seek(0)

            wb_hist = openpyxl.load_workbook(excel_hist_buffer)
            ws_hist = wb_hist["Datos"]
            img_hist_final = XLImage(io.BytesIO(img_hist_buffer.getvalue()))
            col_img_hist = openpyxl.utils.get_column_letter(len(df_hist_export.columns) + 2)
            img_hist_final.anchor = f"{col_img_hist}2"
            ws_hist.add_image(img_hist_final)

            final_hist_buffer = io.BytesIO()
            wb_hist.save(final_hist_buffer)
            final_hist_buffer.seek(0)

            st.download_button(
                "⬇️ Descargar Excel",
                data=final_hist_buffer,
                file_name=nombre_excel_hist,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="descargar_excel_historial",
            )

            st.markdown("##### ☁️ O guardarlo directo en tu repo de GitHub")
            carpeta_repo_hist = st.text_input(
                "Carpeta dentro del repo (déjalo vacío para la raíz)",
                value="resultados",
                key="carpeta_github_historial",
            )
            carpeta_repo_hist = carpeta_repo_hist.strip().strip("/")
            ruta_completa_hist = (
                f"{carpeta_repo_hist}/{nombre_excel_hist}" if carpeta_repo_hist else nombre_excel_hist
            )
            st.caption(f"Se guardará en: `{ruta_completa_hist}` dentro de `{GITHUB_REPO}`")

            if st.button("💾 Guardar en GitHub", key="btn_guardar_github_historial"):
                try:
                    resultado_hist = subir_archivo_a_github(
                        ruta_completa_hist,
                        final_hist_buffer.getvalue(),
                        f"Agrega {nombre_excel_hist} (comparación de historial) desde la app",
                    )
                    url_archivo_hist = resultado_hist.get("content", {}).get("html_url", "")
                    st.success(f"¡Listo! Se guardó en `{ruta_completa_hist}`.")
                    if url_archivo_hist:
                        st.markdown(f"[Ver el archivo en GitHub]({url_archivo_hist})")
                except Exception as e:
                    st.error(f"No pude guardar en GitHub: {e}")

    with st.expander("🗑️ Administrar historial (borrar comparaciones guardadas)"):
        etiqueta_borrar = st.selectbox(
            "Elige cuál borrar", ["(ninguna)"] + etiquetas_disponibles, key="etiqueta_borrar"
        )
        if etiqueta_borrar != "(ninguna)" and st.button("Borrar esta comparación", key="btn_borrar_historial"):
            idx_borrar = etiquetas_disponibles.index(etiqueta_borrar)
            nuevas_entradas = [e for i, e in enumerate(historial_entradas) if i != idx_borrar]
            guardar_historial(nuevas_entradas)
            st.success("Borrada. Recarga la página para ver el historial actualizado.")
