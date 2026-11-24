"""
Aplicacion Web Streamlit para Unificacion y Procesamiento de Comisiones NLG (National Life Group).
Permite cargar 1 o 2 archivos Excel (.xlsx), unifica los registros, clasifica comisiones
entre el Agente Principal y su Equipo, calcula estados (Activa / Chargeback) y genera:
1. El archivo PDF unificado con diseño y columnas identicas al PDF de referencia.
2. El bloque de texto preconstruido listo para copiar en un solo clic.
"""

import streamlit as st
import pandas as pd
import io
import fitz # PyMuPDF para preview del PDF

from processor import (
    extract_excel_data,
    unify_dataframes,
    process_commissions,
    generate_prebuilt_text,
    REQUIRED_COLUMNS
)
from pdf_generator import (
    build_unified_pdf,
    clean_filename
)

# Configuracion de pagina de Streamlit
st.set_page_config(
    page_title="Procesador de Comisiones NL",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS personalizados para una interfaz minimalista, moderna y ejecutiva
st.markdown("""
<style>
    /* Estilos generales */
    .main-header {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: #1E293B;
        font-weight: 700;
        font-size: 2.1rem;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        color: #64748B;
        font-size: 1.05rem;
        margin-bottom: 1.5rem;
    }
    .card-metric {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
    }
    .card-metric-title {
        font-size: 0.85rem;
        color: #64748B;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .card-metric-val {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0F172A;
        margin-top: 0.3rem;
    }
    .badge-activa {
        background-color: #DCFCE7;
        color: #166534;
        padding: 0.2rem 0.6rem;
        border-radius: 12px;
        font-weight: 600;
        font-size: 0.8rem;
    }
    .badge-chargeback {
        background-color: #FEE2E2;
        color: #991B1B;
        padding: 0.2rem 0.6rem;
        border-radius: 12px;
        font-weight: 600;
        font-size: 0.8rem;
    }
    /* Estilo del contenedor de texto preconstruido */
    .stCodeBlock {
        border-radius: 8px;
        border: 1px solid #CBD5E1;
    }
</style>
""", unsafe_allow_html=True)

# Encabezado
st.markdown('<div class="main-header">💼 Lector y Procesador de Comisiones NL</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-header">Unifica 1 o 2 archivos Excel de comisiones, clasifica por agente principal y equipo, y genera el PDF oficial junto con el texto preconstruido listo para copiar.</div>',
    unsafe_allow_html=True
)

# Inicializar variables en session_state
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'pdf_bytes' not in st.session_state:
    st.session_state.pdf_bytes = None
if 'pdf_filename' not in st.session_state:
    st.session_state.pdf_filename = None
if 'prebuilt_text' not in st.session_state:
    st.session_state.prebuilt_text = None
if 'preview_image' not in st.session_state:
    st.session_state.preview_image = None

# Seccion de carga de archivos
col_upload, col_info = st.columns([2, 1])

with col_upload:
    uploaded_files = st.file_uploader(
        "Carga 1 o 2 archivos de comisiones (.xlsx):",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Arrastra o selecciona uno o dos archivos de comisiones en formato Excel (.xlsx)"
    )

with col_info:
    st.info(
        "📌 **Instrucciones Rápidas:**\n"
        "- Carga **1 o 2 archivos `.xlsx`** simultáneamente.\n"
        "- El sistema detectará automáticamente el **Agente Principal** (`Exported By:`).\n"
        "- Haz clic en **'Procesar Comisiones'** para obtener el PDF y el texto formateado listo para copiar."
    )

# Validacion del numero de archivos
files_valid = False
if uploaded_files:
    if len(uploaded_files) > 2:
        st.warning(f"⚠️ Se cargaron {len(uploaded_files)} archivos. Solo se procesarán los primeros 2 archivos seleccionados.")
        active_files = uploaded_files[:2]
        files_valid = True
    elif len(uploaded_files) in [1, 2]:
        st.success(f"✅ {len(uploaded_files)} archivo(s) listo(s) para procesar.")
        active_files = uploaded_files
        files_valid = True
else:
    active_files = []

# Boton de accion
col_btn, col_empty = st.columns([1, 3])
with col_btn:
    btn_process = st.button("🚀 Procesar Comisiones", type="primary", disabled=not files_valid, use_container_width=True)

if btn_process and files_valid:
    with st.spinner("Procesando y validando archivos Excel..."):
        try:
            data_list = []
            for file in active_files:
                file_bytes = file.read()
                file.seek(0)
                df_extracted, exported_by, exported_on = extract_excel_data(file_bytes)
                data_list.append((df_extracted, exported_by, exported_on))

            # Unificacion de archivos
            unified_df, final_exported_by, final_exported_on = unify_dataframes(data_list)

            # Clasificacion y calculo de estados
            classified = process_commissions(unified_df, final_exported_by)

            # Generacion de texto preconstruido
            text_output = generate_prebuilt_text(classified)

            # Generacion de PDF unificado
            pdf_buf = build_unified_pdf(unified_df, final_exported_by, final_exported_on)
            pdf_bytes = pdf_buf.getvalue()
            pdf_name = clean_filename(final_exported_by)

            # Generar imagen de preview con PyMuPDF
            preview_img_bytes = None
            try:
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                if len(doc) > 0:
                    page = doc[0]
                    # Renderizar a 130 DPI para visualizacion nitida
                    pix = page.get_pixmap(dpi=130)
                    preview_img_bytes = pix.tobytes("png")
            except Exception as e_prev:
                preview_img_bytes = None

            # Guardar en estado de sesion
            st.session_state.processed_data = {
                'unified_df': unified_df,
                'exported_by': final_exported_by,
                'exported_on': final_exported_on,
                'classified': classified,
            }
            st.session_state.pdf_bytes = pdf_bytes
            st.session_state.pdf_filename = pdf_name
            st.session_state.prebuilt_text = text_output
            st.session_state.preview_image = preview_img_bytes

            st.toast("¡Comisiones procesadas exitosamente!", icon="🎉")

        except ValueError as e_val:
            st.error(f"❌ Error de validación: {str(e_val)}")
        except Exception as e:
            st.error(f"❌ Ocurrió un error inesperado al procesar los archivos: {str(e)}")

# Si hay resultados procesados en sesion, mostrarlos
if st.session_state.processed_data is not None:
    p_data = st.session_state.processed_data
    classified = p_data['classified']
    unified_df = p_data['unified_df']
    agent_by = p_data['exported_by']
    exported_on = p_data['exported_on']
    prebuilt_text = st.session_state.prebuilt_text
    pdf_bytes = st.session_state.pdf_bytes
    pdf_filename = st.session_state.pdf_filename

    st.markdown("---")

    # Informacion del Agente Principal
    col_ag1, col_ag2 = st.columns([3, 1])
    with col_ag1:
        st.markdown(f"### 👤 Agente Principal detectado: **{agent_by}**")
        st.caption(f"📅 Fecha(s) de exportación: {exported_on} | Total de filas de clientes: {len(unified_df)}")

    # Tarjetas de metricas
    nl_items = classified['principal']
    team_items = classified['team']

    nl_activas = sum(1 for x in nl_items if x['status'] == 'Activa')
    nl_chargebacks = sum(1 for x in nl_items if x['status'] == 'Chargeback')

    team_activas = sum(1 for x in team_items if x['status'] == 'Activa')
    team_chargebacks = sum(1 for x in team_items if x['status'] == 'Chargeback')

    total_gross = classified['raw_df']['Gross_Numeric'].sum()

    m1, m2, m3, m4, m5 = st.columns(5)
    with m1:
        st.markdown(f"""
        <div class="card-metric">
            <div class="card-metric-title">Registros Totales</div>
            <div class="card-metric-val">{len(unified_df)}</div>
        </div>
        """, unsafe_allow_html=True)
    with m2:
        st.markdown(f"""
        <div class="card-metric">
            <div class="card-metric-title">Clientes Únicos</div>
            <div class="card-metric-val">{len(nl_items) + len(team_items)}</div>
        </div>
        """, unsafe_allow_html=True)
    with m3:
        st.markdown(f"""
        <div class="card-metric">
            <div class="card-metric-title">Comisiones NL</div>
            <div class="card-metric-val" style="color: #0284C7;">{nl_activas} <span style="font-size:0.9rem; color:#64748B;">/ {nl_chargebacks} CB</span></div>
        </div>
        """, unsafe_allow_html=True)
    with m4:
        st.markdown(f"""
        <div class="card-metric">
            <div class="card-metric-title">Comisiones Equipo</div>
            <div class="card-metric-val" style="color: #6366F1;">{team_activas} <span style="font-size:0.9rem; color:#64748B;">/ {team_chargebacks} CB</span></div>
        </div>
        """, unsafe_allow_html=True)
    with m5:
        color_val = "#16A34A" if total_gross >= 0 else "#DC2626"
        st.markdown(f"""
        <div class="card-metric">
            <div class="card-metric-title">Monto Bruto Total</div>
            <div class="card-metric-val" style="color: {color_val};">${total_gross:,.2f}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Pestanas de visualizacion y descarga
    tab_text, tab_pdf, tab_summary, tab_data = st.tabs([
        "📋 Texto Preconstruido (Copiar)",
        "📄 Archivo PDF Unificado",
        "📊 Desglose de Clientes",
        "📁 Tabla Completa Unificada"
    ])

    # PESTAÑA 1: TEXTO PRECONSTRUIDO
    with tab_text:
        st.markdown("#### Bloque de Texto Listo para Copiar")
        st.markdown(
            "Utiliza el botón de copiado integrado en la esquina superior derecha del bloque, "
            "o el botón de acceso directo debajo:"
        )

        st.code(prebuilt_text, language=None)

        # Boton interactivo de copiado al portapapeles mediante componente HTML/JS
        copy_js = f"""
        <div style="margin-top: 10px;">
            <button id="btnCopy" onclick="copyToClipboard()" style="
                background-color: #0284C7;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 15px;
                font-weight: 600;
                border-radius: 6px;
                cursor: pointer;
                transition: background-color 0.2s ease;
            ">📋 Copiar Texto al Portapapeles</button>
            <span id="copyMsg" style="margin-left: 15px; color: #16A34A; font-weight: 600; display: none;">✅ ¡Copiado con éxito!</span>
        </div>
        <script>
        function copyToClipboard() {{
            const textToCopy = {repr(prebuilt_text)};
            navigator.clipboard.writeText(textToCopy).then(function() {{
                const msg = document.getElementById("copyMsg");
                msg.style.display = "inline";
                setTimeout(function() {{ msg.style.display = "none"; }}, 2500);
            }}).catch(function(err) {{
                alert('No se pudo copiar automáticamente. Por favor utiliza el botón de copiado en el bloque de código.');
            }});
        }}
        </script>
        """
        st.components.v1.html(copy_js, height=60)

    # PESTAÑA 2: ARCHIVO PDF UNIFICADO
    with tab_pdf:
        st.markdown("#### Archivo PDF Oficial Unificado")
        st.markdown(f"Nombre del archivo listo para descarga: `{pdf_filename}`")

        col_dl, col_blank = st.columns([1, 2])
        with col_dl:
            st.download_button(
                label=f"⬇️ Descargar PDF: {pdf_filename}",
                data=pdf_bytes,
                file_name=pdf_filename,
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )

        # Vista previa interactiva de la primera pagina del PDF
        if st.session_state.preview_image is not None:
            st.markdown("##### Vista Previa del Documento:")
            st.image(
                st.session_state.preview_image,
                caption="Vista previa del PDF generado (Página 1)",
                use_container_width=True
            )

    # PESTAÑA 3: DESGLOSE DE CLIENTES
    with tab_summary:
        st.markdown("#### Desglose Detallado por Cliente y Estado")
        
        c_nl, c_team = st.columns(2)

        with c_nl:
            st.markdown("##### 📌 Comision y Chargeback NL (Agente Principal)")
            if nl_items:
                nl_df = pd.DataFrame([{
                    'Asegurado': item['insured'],
                    'Estado': item['status'],
                    'Comisión Neta ($)': f"${item['total']:,.2f}",
                    'Pólizas': ", ".join(item['policies']),
                    'Registros': item['count']
                } for item in nl_items])
                st.dataframe(nl_df, use_container_width=True, hide_index=True)
            else:
                st.info("No se encontraron registros del agente principal.")

        with c_team:
            st.markdown("##### 👥 Comision y Chargeback de Equipo NL")
            if team_items:
                team_df_disp = pd.DataFrame([{
                    'Asegurado': item['insured'],
                    'Estado': item['status'],
                    'Comisión Neta ($)': f"${item['total']:,.2f}",
                    'Agente Asignado': item['agent'],
                    'Pólizas': ", ".join(item['policies']),
                    'Registros': item['count']
                } for item in team_items])
                st.dataframe(team_df_disp, use_container_width=True, hide_index=True)
            else:
                st.info("No se encontraron registros de equipo.")

    # PESTAÑA 4: TABLA COMPLETA UNIFICADA
    with tab_data:
        st.markdown("#### Registros Unificados de los Archivos")
        st.caption(f"Mostrando los {len(unified_df)} registros combinados con las 10 columnas oficiales:")
        st.dataframe(unified_df, use_container_width=True, hide_index=True)

# Pie de pagina de la aplicacion
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #94A3B8; font-size: 0.85rem;'>"
    "Desarrollado para procesamiento de reportes de comisiones NLG (National Life Group) con Streamlit y ReportLab."
    "</div>",
    unsafe_allow_html=True
)
