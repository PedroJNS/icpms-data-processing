# 🧪 Analizador ICP-MS - Concentración en Muestras Sólidas

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/)
[![Python 3.9+](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://www.python.org/)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

Aplicación web desarrollada en **Streamlit** para el procesamiento automatizado y la visualización interactiva de datos de espectrometría de masas con plasma de acoplamiento inductivo (**ICP-MS**), optimizada para exportaciones de **Agilent 7900 MassHunter**.

---

## 🚀 Características Principales

* **📥 Carga Multiformato:** Soporte para archivos de datos en `.xlsx`, `.xls` y `.csv`.
* **🧠 Detección Inteligente:** Identificación automática de muestras, blancos de calibración/reactivos y filtrado automático de patrones e estándares internos (ISTD).
* **⚖️ Parámetros de Digestión Modificables:** Tabla interactiva para ajustar rápidamente la Masa de muestra ($mg$) y el Volumen de digestión ($mL$).
* **🔄 Conmutador de Unidades (% wt / ppm):** Botón *toggle* para alternar en tiempo real entre Porcentaje en Peso (`% wt`) y Partes por Millón (`ppm` / `mg/kg`).
* **🧮 Corrección por Blancos:** Resta automática del promedio de blancos seleccionados antes de calcular las concentraciones sólidas finales.
* **📈 Visualización Gráfica:** Generación de gráficos de barras interactivos con opción de escala logarítmica y paletas cromáticas dinámicas.
* **💾 Exportación Directa:** Descarga de tablas de resultados en Excel (`.xlsx`) y gráficos en calidad de publicación (`.png`).

---

## 🛠️ Instalación y Ejecución Local

Si deseas ejecutar la aplicación localmente en tu equipo:

1. **Clonar el repositorio:**
   ```bash
   git clone [https://github.com/TU-USUARIO/icpms-data-processing.git](https://github.com/TU-USUARIO/icpms-data-processing.git)
   cd icpms-data-processing
Instalar las dependencias:Bashpip install -r requirements.txt
Lanzar la aplicación:Bashstreamlit run app.py
📋 Estructura de requirements.txtPara desplegar en Streamlit Community Cloud, asegúrate de tener un archivo requirements.txt en la raíz del repositorio con el siguiente contenido:Plaintextstreamlit
pandas
openpyxl
matplotlib
🧪 Formulación Matemática UtilizadaCorrección de Blanco:$$\text{ppb}_{\text{corregido}} = \max(0, \text{ppb}_{\text{muestra}} - \overline{\text{ppb}}_{\text{blancos}})$$Cálculo en Partes por Millón (ppm / mg/kg):$$\text{ppm} = \frac{\text{ppb}_{\text{corregido}} \times \text{Volumen (mL)}}{\text{Masa (mg)}}$$Cálculo en Porcentaje en Peso (% wt):$$\% \text{ wt} = \frac{\text{ppm}}{10000} = \frac{\text{ppb}_{\text{corregido}} \times \text{Volumen (mL)}}{\text{Masa (mg)} \times 10000}$$👤 Autor y ContactoDesarrollador: Pedro J. Navarrete SegadoInstitución: Universidad de Jaén (UJA)Email de Contacto: pnsegado@ujaen.es📄 LicenciaEste proyecto está distribuido bajo la licencia GNU General Public License v3.0 (GPL-3.0). Puedes consultar el archivo LICENSE para más información.
