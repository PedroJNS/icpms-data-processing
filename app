"""
===============================================================================
Aplicación: Analizador ICP-MS - Concentración (% wt) (Versión Streamlit)
Desarrollador: Pedro J. Navarrete Segado
Institución: Universidad de Jaén (UJA)
Contacto: pnsegado@ujaen.es
Licencia: GNU General Public License v3.0 (GPL-3.0) (Copyright (c) Pedro J. Navarrete Segado)
===============================================================================
"""

import io
import warnings
import matplotlib.pyplot as plt
import matplotlib
import pandas as pd
import streamlit as st

# Configuración de página de Streamlit
st.set_page_config(
    page_title="Analizador ICP-MS - Concentración (% wt)",
    page_icon="🧪",
    layout="wide",
)

# Ocultar advertencias
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
try:
  warnings.filterwarnings("ignore", category=pd.errors.PerformanceWarning)
except AttributeError:
  pass


# --- FUNCIONES DE PROCESAMIENTO ---
def procesar_archivo_raw(uploaded_file):
  """Lee el archivo cargado y localiza los encabezados y columnas clave."""
  if uploaded_file.name.lower().endswith(".csv"):
    df_raw = pd.read_csv(uploaded_file, header=None)
  else:
    excel_obj = pd.ExcelFile(uploaded_file)
    hoja = (
        "icpms"
        if "icpms" in [s.lower() for s in excel_obj.sheet_names]
        else excel_obj.sheet_names[0]
    )
    df_raw = pd.read_excel(excel_obj, sheet_name=hoja, header=None)

  fila_encabezado = None
  col_sample_idx = None
  col_type_idx = None
  col_date_idx = None

  for r_idx, row in df_raw.iterrows():
    for c_idx, val in enumerate(row):
      val_str = str(val).strip().lower()
      if val_str in ["sample name", "nombre muestra"]:
        fila_encabezado = r_idx
        col_sample_idx = c_idx
      elif val_str in ["type", "tipo"]:
        col_type_idx = c_idx
      elif any(
          k in val_str for k in ["acq", "date", "time", "fecha", "hora"]
      ):
        col_date_idx = c_idx

    if fila_encabezado is not None:
      break

  if fila_encabezado is not None:
    headers = (
        df_raw.iloc[fila_encabezado].fillna("").astype(str).tolist()
    )
    df_data = df_raw.iloc[fila_encabezado + 1 :].copy()
    df_data.columns = headers
  else:
    df_data = df_raw.copy()
    col_sample_idx = 3

  if (
      col_date_idx is None
      and col_type_idx is not None
      and col_type_idx > 0
  ):
    col_date_idx = col_type_idx - 1

  return (
      df_raw,
      df_data,
      fila_encabezado,
      col_sample_idx,
      col_type_idx,
      col_date_idx,
  )


def preparar_mapeo_columnas(df_raw, fila_encabezado, col_sample_idx):
  """Mapea las columnas de elementos evitando los ISTD."""
  element_map = {}
  curr_elem = ""
  for c_idx in range(df_raw.shape[1]):
    val0 = str(df_raw.iloc[0, c_idx]).strip()
    if val0 and val0.lower() != "nan":
      curr_elem = val0
    element_map[c_idx] = curr_elem

  cols_interes = []
  for c_idx in range(df_raw.shape[1]):
    param = str(df_raw.iloc[fila_encabezado, c_idx]).strip()
    elem = element_map[c_idx]

    if "( ISTD )" in elem or "istd" in elem.lower():
      continue

    if c_idx == col_sample_idx:
      cols_interes.append((c_idx, "Sample Name", False, "none"))
    elif "conc" in param.lower():
      unit = "ppm" if "ppm" in param.lower() else "ppb"
      nombre_elem_clean = (
          elem.split("[")[0].strip() if "[" in elem else elem.strip()
      )
      nombre_elem = (
          f"{nombre_elem_clean} (% wt)"
          if nombre_elem_clean
          else f"Col_{c_idx} (% wt)"
      )
      cols_interes.append((c_idx, nombre_elem, True, unit))

  return cols_interes


def calcular_resultados(
    df_data,
    cols_interes,
    col_sample_idx,
    muestras_elegidas,
    blancos_seleccionados,
    df_params,
):
  """Calcula los promedios de los blancos y las concentraciones reales (% wt)."""
  # 1. Calcular promedio de blancos en ppb
  promedio_blancos_ppb = {}
  if blancos_seleccionados:
    col_serie = df_data.iloc[:, col_sample_idx].astype(str).str.strip()
    df_blancos = df_data[col_serie.isin(blancos_seleccionados)]

    for c_idx, _, es_conc, unidad in cols_interes:
      if es_conc:
        valores_ppb = []
        for _, fila in df_blancos.iterrows():
          v_str = (
              str(fila.iloc[c_idx])
              .replace("<", "")
              .replace(">", "")
              .replace(",", ".")
              .strip()
          )
          try:
            val_num = float(v_str)
            ppb = val_num * 1000.0 if unidad == "ppm" else val_num
            valores_ppb.append(ppb)
          except ValueError:
            pass

        if valores_ppb:
          promedio_blancos_ppb[c_idx] = sum(valores_ppb) / len(valores_ppb)

  # 2. Recalcular concentraciones (% wt)
  col_serie = df_data.iloc[:, col_sample_idx].astype(str).str.strip()
  df_filt = df_data[col_serie.isin(muestras_elegidas)].copy()

  params_dict = df_params.set_index("Sample Name").to_dict(orient="index")

  filas_export = []
  for _, fila in df_filt.iterrows():
    nombre_muestra = str(fila.iloc[col_sample_idx]).strip()
    p_m = params_dict.get(
        nombre_muestra, {"Masa (mg)": 15.0, "Volumen (mL)": 10.0}
    )

    masa_mg = float(p_m["Masa (mg)"])
    vol_ml = float(p_m["Volumen (mL)"])

    row_dict = {
        "Sample Name": nombre_muestra,
        "Masa (mg)": masa_mg,
        "Volumen (mL)": vol_ml,
    }

    for c_idx, titulo, es_conc, unidad in cols_interes:
      if es_conc:
        v_str = (
            str(fila.iloc[c_idx])
            .replace("<", "")
            .replace(">", "")
            .replace(",", ".")
            .strip()
        )
        try:
          val_num = float(v_str)
          ppb = val_num * 1000.0 if unidad == "ppm" else val_num
          blanco_ppb = promedio_blancos_ppb.get(c_idx, 0.0)
          ppb_corregido = max(0.0, ppb - blanco_ppb)

          pct = (ppb_corregido * vol_ml) / (masa_mg * 10000.0)
        except ValueError:
          pct = 0.0

        row_dict[titulo] = pct

    filas_export.append(row_dict)

  return pd.DataFrame(filas_export)


# --- INTERFAZ STREAMLIT ---

# Créditos en Barra Lateral (Sidebar)
with st.sidebar:
  st.header("ℹ️ Acerca de")
  st.markdown("""
    **Analizador ICP-MS - Concentración (% wt)**  
    * **Desarrollador:** Pedro J. Navarrete Segado  
    * **Institución:** Universidad de Jaén (UJA)  
    * **Contacto:** [pnsegado@ujaen.es](mailto:pnsegado@ujaen.es)  
    * **Licencia:** GNU General Public License v3.0 (GPL-3.0)  
    
    *Copyright (c) Pedro J. Navarrete Segado*
    """)
  st.divider()
  st.caption(
      "Web application to upload Agilent 7900 MassHunter files, input sample"
      " digestion data (mass & volume), calculate real concentrations in solid"
      " samples (ppm and %), and interactively visualize analytical results."
  )

# Encabezado Principal
st.title("🧪 Analizador ICP-MS - Concentración (% wt)")
st.caption(
    "Desarrollado por Pedro J. Navarrete Segado | Universidad de Jaén (UJA)"
)

# 1. Cargar Documento
uploaded_file = st.file_uploader(
    "1. Cargar Documento (Excel / CSV)", type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
  try:
    (
        df_raw,
        df_data,
        fila_encabezado,
        col_sample_idx,
        col_type_idx,
        col_date_idx,
    ) = procesar_archivo_raw(uploaded_file)
    cols_interes = preparar_mapeo_columnas(
        df_raw, fila_encabezado, col_sample_idx
    )

    # Identificar Muestras y Blancos sugeridos
    palabras_ignorar = [
        "blank",
        "blanco",
        "ppb",
        "calblk",
        "calstd",
        "blkvrfy",
        "qc",
        "driftchk",
        "cicspike",
        "isostd",
        "dilstd",
        "bkgnd",
        "fqblk",
    ]

    if col_type_idx is not None and fila_encabezado is not None:
      val_type = (
          df_data.iloc[:, col_type_idx]
          .astype(str)
          .str.strip()
          .str.lower()
      )
      df_muestras_raw = df_data[val_type == "sample"]
    else:
      df_muestras_raw = df_data

    todas_muestras = (
        df_muestras_raw.iloc[:, col_sample_idx].dropna().astype(str).tolist()
    )

    muestras_validas = []
    blancos_detectados = []

    for m in todas_muestras:
      m_clean = m.strip()
      m_lower = m_clean.lower()
      if m_clean and m_lower != "nan":
        if "blank" in m_lower or "blanco" in m_lower:
          blancos_detectados.append(m_clean)
        elif not any(p in m_lower for p in palabras_ignorar):
          muestras_validas.append(m_clean)

    # Quitar duplicados manteniendo orden
    muestras_validas = list(dict.fromkeys(muestras_validas))
    blancos_detectados = list(dict.fromkeys(blancos_detectados))

    st.divider()

    # 2 y 3. Selección de Muestras y Blancos
    col_sel1, col_sel2 = st.columns(2)

    with col_sel1:
      muestras_elegidas = st.multiselect(
          "2. Selecciona las muestras a analizar:",
          options=muestras_validas,
          default=muestras_validas,
      )

    with col_sel2:
      blancos_seleccionados = st.multiselect(
          "3. Selecciona los Blancos para restar:",
          options=blancos_detectados,
          default=blancos_detectados,
      )

    if not muestras_elegidas:
      st.warning("⚠️ Selecciona al menos una muestra para continuar.")
    else:
      st.divider()
      st.subheader("4. Parámetros de Digestión (Masa y Volumen)")
      st.info(
          "💡 **Pista:** Puedes editar los valores de Masa y Volumen directamente"
          " en la tabla a continuación:"
      )

      # DataFrame editable para Masa y Volumen
      df_params_init = pd.DataFrame({
          "Sample Name": muestras_elegidas,
          "Masa (mg)": [15.0] * len(muestras_elegidas),
          "Volumen (mL)": [10.0] * len(muestras_elegidas),
      })

      # st.data_editor permite modificar datos interactivamente
      df_params_edited = st.data_editor(
          df_params_init,
          num_rows="fixed",
          use_container_width=True,
          column_config={
              "Masa (mg)": st.column_config.NumberColumn(
                  min_value=0.001, format="%.2f"
              ),
              "Volumen (mL)": st.column_config.NumberColumn(
                  min_value=0.001, format="%.2f"
              ),
          },
      )

      # Calcular Resultados
      df_results = calcular_resultados(
          df_data,
          cols_interes,
          col_sample_idx,
          muestras_elegidas,
          blancos_seleccionados,
          df_params_edited,
      )

      st.divider()
      st.subheader("📊 Tabla de Resultados (% wt)")

      # Dar formato legible a la vista
      df_view = df_results.copy()
      cols_pct = [c for c in df_view.columns if "% wt" in c]

      for c in cols_pct:

        def fmt_val(val):
          if val >= 0.000001:
            return f"{val:.6f}%"
          elif val > 0:
            return f"{val:.3e}%"
          return "-"

        df_view[c] = df_view[c].apply(fmt_val)

      st.dataframe(df_view, use_container_width=True)

      # Botón de exportación a Excel
      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_results.to_excel(
            writer, index=False, sheet_name="Resultados_ICP-MS"
        )
      excel_bytes = output.getvalue()

      st.download_button(
          label="📥 Exportar Resultados a Excel",
          data=excel_bytes,
          file_name="Resultados_ICPMS_pct.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
      )

      # --- SECCIÓN GRÁFICA ---
      st.divider()
      st.subheader("📈 Gráfico de Concentraciones por Metal")

      cols_metales = [c for c in df_results.columns if "% wt" in c]

      if cols_metales:
        col_g1, col_g2 = st.columns([1, 3])

        with col_g1:
          usar_log = st.checkbox(
              "Usar escala logarítmica",
              value=False,
              help="Útil si hay diferencias de varios órdenes de magnitud entre metales.",
          )

          nombres_metales = [c.replace(" (% wt)", "") for c in cols_metales]
          metales_sel = st.multiselect(
              "Selecciona metales para graficar:",
              options=nombres_metales,
              default=nombres_metales[: min(5, len(nombres_metales))],
          )

        with col_g2:
          if metales_sel:
            cols_graficar = [f"{m} (% wt)" for m in metales_sel]
            df_plot = df_results.set_index("Sample Name")[cols_graficar]
            df_plot.columns = [
                c.replace(" (% wt)", "") for c in df_plot.columns
            ]

            fig, ax = plt.subplots(figsize=(10, 5), dpi=100)

            num_metales = len(metales_sel)
            cmap_dinamico = matplotlib.colormaps["turbo"].resampled(num_metales)

            df_plot.plot(
                kind="bar",
                ax=ax,
                width=0.7,
                edgecolor="black",
                linewidth=0.5,
                colormap=cmap_dinamico,
            )

            if usar_log:
              ax.set_yscale("log")
              ax.set_ylabel(
                  "Concentración (% wt) - Escala Log",
                  fontsize=10,
                  fontweight="bold",
              )
            else:
              ax.set_ylabel(
                  "Concentración (% wt)", fontsize=10, fontweight="bold"
              )

            ncols = 1 if num_metales <= 10 else (2 if num_metales <= 20 else 3)
            ax.legend(
                title="Elementos",
                bbox_to_anchor=(1.02, 1),
                loc="upper left",
                frameon=True,
                ncol=ncols,
                fontsize=8,
                title_fontsize=9,
            )

            ax.set_title(
                "Concentración Elemental en Muestras (% wt)",
                fontsize=12,
                fontweight="bold",
                pad=12,
            )
            ax.set_xlabel("Muestra", fontsize=10, fontweight="bold")
            ax.grid(axis="y", linestyle="--", alpha=0.6)
            plt.xticks(rotation=45, ha="right")
            fig.tight_layout()

            st.pyplot(fig)

            # Descarga de Imagen PNG
            img_buf = io.BytesIO()
            fig.savefig(img_buf, format="png", dpi=300, bbox_inches="tight")
            img_buf.seek(0)

            st.download_button(
                label="💾 Descargar Gráfico como PNG",
                data=img_buf,
                file_name="grafico_concentraciones_icpms.png",
                mime="image/png",
            )
            plt.close(fig)
          else:
            st.info(
                "Selecciona al menos un metal en el panel de la izquierda para"
                " visualizar el gráfico."
            )

  except Exception as e:
    st.error(f"❌ Error al procesar el archivo: {e}")
