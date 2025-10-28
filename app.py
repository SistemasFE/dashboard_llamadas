#!/usr/bin/env python3
"""
Aplicación Streamlit para análisis de categorías de llamadas.

Esta aplicación permite subir archivos Excel con datos de llamadas,
analizar las categorías y generar reportes visuales interactivos.

Autor: Asistente IA
Fecha: 2025-01-02
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from collections import Counter
import tempfile
import os
from pathlib import Path
import sys
import logging

logger = logging.getLogger(__name__)

# Agregar el directorio scripts al path para importar módulos
sys.path.append(str(Path(__file__).parent.parent))

try:
    from categoria_analysis import ExcelCategoryAnalyzer
except ImportError:
    st.error("Error: No se pudo importar el módulo de análisis. Verifica la estructura del proyecto.")
    st.stop()

# Configuración de la página
st.set_page_config(
    page_title="Análisis de Categorías de Llamadas",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título principal
st.title("📊 Análisis de Categorías de Llamadas")
st.markdown("---")

# Función para crear gráficos de categorías
def crear_grafico_categorias(category_counter, title="Top Categorías"):
    """Crear gráfico de barras para categorías."""
    if not category_counter:
        return None

    # Convertir counter a DataFrame
    df_categories = pd.DataFrame(
        category_counter.most_common(20),  # Top 20
        columns=['Categoría', 'Frecuencia']
    )

    # Crear gráfico interactivo
    fig = px.bar(
        df_categories,
        x='Frecuencia',
        y='Categoría',
        orientation='h',
        title=title,
        labels={'Frecuencia': 'Número de Llamadas', 'Categoría': 'Categoría'},
        color='Frecuencia',
        color_continuous_scale='Blues'
    )

    fig.update_layout(
        height=max(400, len(df_categories) * 25),
        yaxis={'categoryorder': 'total ascending'}
    )

    return fig

# Función para crear gráfico de distribución
def crear_grafico_distribucion(counter, total_rows):
    """Crear gráfico de distribución de categorías por volumen."""
    if not counter:
        return None

    # Crear datos de distribución
    high_volume = []
    medium_volume = []
    low_volume = []

    for category, count in counter.items():
        percentage = (count / total_rows * 100) if total_rows > 0 else 0
        if percentage >= 10:
            high_volume.append((category, count, percentage))
        elif percentage >= 1:
            medium_volume.append((category, count, percentage))
        else:
            low_volume.append((category, count, percentage))

    # Crear datos para el gráfico
    segments = ['Alto Volumen (>10%)', 'Medio Volumen (1-10%)', 'Bajo Volumen (<1%)']
    counts = [len(high_volume), len(medium_volume), len(low_volume)]
    total_calls = [sum(count for _, count, _ in high_volume),
                   sum(count for _, count, _ in medium_volume),
                   sum(count for _, count, _ in low_volume)]

    fig = go.Figure(data=[
        go.Bar(
            x=segments,
            y=counts,
            name='Número de Categorías',
            marker_color='lightblue',
            yaxis='y1'
        ),
        go.Scatter(
            x=segments,
            y=[sum(count for _, count, _ in high_volume) / total_rows * 100 if total_rows > 0 else 0,
               sum(count for _, count, _ in medium_volume) / total_rows * 100 if total_rows > 0 else 0,
               sum(count for _, count, _ in low_volume) / total_rows * 100 if total_rows > 0 else 0],
            name='% del Total',
            mode='lines+markers',
            marker_color='darkblue',
            yaxis='y2'
        )
    ])

    fig.update_layout(
        title='Distribución de Categorías por Volumen',
        xaxis_title='Segmento',
        yaxis=dict(title='Número de Categorías', side='left'),
        yaxis2=dict(title='% del Total de Llamadas', side='right', overlaying='y'),
        legend=dict(x=0.5, y=1.1, xanchor='center'),
        height=400
    )

    return fig

# Función principal
def main():
    st.sidebar.header("⚙️ Configuración")

    # Selector de archivos
    uploaded_files = st.sidebar.file_uploader(
        "📁 Subir archivos Excel",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Selecciona uno o más archivos Excel con datos de llamadas"
    )

    # Configuración de fechas
    st.sidebar.subheader("📅 Filtro de Fechas (Opcional)")

    use_date_filter = st.sidebar.checkbox("Aplicar filtro de fechas")

    start_date = None
    end_date = None

    if use_date_filter:
        col1, col2 = st.sidebar.columns(2)

        with col1:
            start_date_input = st.date_input(
                "Fecha inicio",
                value=datetime.now() - timedelta(days=30),
                help="Fecha de inicio para filtrar los datos"
            )

        with col2:
            end_date_input = st.date_input(
                "Fecha fin",
                value=datetime.now(),
                help="Fecha de fin para filtrar los datos"
            )

        if start_date_input and end_date_input:
            if start_date_input > end_date_input:
                st.sidebar.error("❌ La fecha de inicio debe ser anterior a la fecha de fin")
            else:
                start_date = datetime.combine(start_date_input, datetime.min.time())
                end_date = datetime.combine(end_date_input, datetime.max.time())

    # Inicializar estado para resultados del análisis y configuraciones
    if 'analysis_results' not in st.session_state:
        st.session_state.analysis_results = None
    if 'date_filters' not in st.session_state:
        st.session_state.date_filters = {'start_date': None, 'end_date': None}

    # Procesar archivos cuando se suban
    if uploaded_files:
        if st.sidebar.button("🚀 Ejecutar Análisis", type="primary", use_container_width=True):

            with st.spinner("🔄 Procesando archivos..."):
                try:
                    # Crear analizador con filtros de fecha
                    # Crear directorio results si no existe
                    results_dir = Path(__file__).parent.parent / "results"
                    results_dir.mkdir(exist_ok=True)

                    analyzer = ExcelCategoryAnalyzer(
                        results_dir=str(results_dir),  # Usar directorio results
                        start_date=start_date,
                        end_date=end_date
                    )

                    # Procesar archivos subidos
                    temp_files = []

                    try:
                        for uploaded_file in uploaded_files:
                            # Guardar archivo temporalmente
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                                tmp_file.write(uploaded_file.getvalue())
                                temp_files.append(Path(tmp_file.name))

                        # Ejecutar análisis
                        total_counter, total_rows, detailed_analysis = analyzer.analyze_multiple_files(temp_files)

                        # Limpiar archivos temporales
                        for temp_file in temp_files:
                            os.unlink(temp_file)

                    except Exception as e:
                        # Limpiar archivos temporales en caso de error
                        for temp_file in temp_files:
                            try:
                                os.unlink(temp_file)
                            except:
                                pass
                        raise e

                    if not total_counter:
                        st.error("❌ No se pudieron extraer categorías de los archivos subidos.")
                        return

                    # Guardar resultados en el estado de sesión para mantenerlos tras la descarga
                    st.session_state.analysis_results = {
                        'total_counter': total_counter,
                        'total_rows': total_rows,
                        'detailed_analysis': detailed_analysis,
                        'files_processed': len(uploaded_files),
                        'generated_at': datetime.now().isoformat(),
                        'analyzer_params': {
                            'results_dir': str(results_dir),
                            'start_date': start_date,
                            'end_date': end_date
                        }
                    }

                    st.session_state.date_filters = {
                        'start_date': start_date,
                        'end_date': end_date
                    }

                    # Guardar Excel en memoria
                    st.session_state.excel_data = analyzer.generate_excel_report(
                        total_counter, total_rows, len(uploaded_files), detailed_analysis
                    )

                    # Guardar analizador para reutilizar sus métodos sin recalcular
                    st.session_state.last_analyzer = analyzer

                    # Mostrar notificación de éxito
                    st.success(f"✅ Análisis completado: {total_rows:,} filas procesadas")

                except Exception as e:
                    st.error(f"❌ Error durante el análisis: {e}")
                    logger.error(f"Error en el análisis: {e}")

    # Determinar si hay resultados que mostrar (de una ejecución actual o previa)
    analysis_data = st.session_state.get('analysis_results')

    if analysis_data:
        total_counter = analysis_data['total_counter']
        total_rows = analysis_data['total_rows']
        detailed_analysis = analysis_data['detailed_analysis']
        files_processed = analysis_data['files_processed']

        analyzer_params = analysis_data.get('analyzer_params', {})
        display_analyzer = st.session_state.get('last_analyzer')
        if display_analyzer is None:
            display_analyzer = ExcelCategoryAnalyzer(
                results_dir=analyzer_params.get('results_dir', str(Path(__file__).parent.parent / "results")),
                start_date=analyzer_params.get('start_date'),
                end_date=analyzer_params.get('end_date')
            )


        # Botón de descarga alineado a la derecha
        excel_data = st.session_state.get('excel_data')
        if excel_data:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"reporte_categorias_{timestamp}.xlsx"
            _, download_col = st.columns([3, 1])
            with download_col:
                st.download_button(
                    label="📥 Descargar Reporte Excel",
                    data=excel_data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
        # Crear pestañas para organizar la información
        tab1, tab2 = st.tabs(["📊 Resumen Ejecutivo", "📈 Gráficos"])

        with tab1:
            st.header("📊 Resumen Ejecutivo")

            # Métricas principales
            col1, col2 = st.columns(2)

            with col1:
                st.metric(
                    "📁 Archivos Procesados",
                    files_processed
                )

            with col2:
                st.metric(
                    "📞 Total de Llamadas",
                    f"{total_rows:,}"
                )

    

            # Ranking de categorías generales (alineado con el reporte Excel)
            st.subheader("🏆 Ranking de Categorías Generales")

            if total_counter:
                top_general = total_counter.most_common()
                df_top_general = pd.DataFrame([
                    {
                        "Ranking": idx,
                        "Categoría": category,
                        "Llamadas": count,
                        "% del Total": f"{(count / total_rows * 100) if total_rows > 0 else 0:.1f}%",
                        "Impacto Operativo": display_analyzer.get_business_impact_category(category, (count / total_rows * 100) if total_rows > 0 else 0)
                    }
                    for idx, (category, count) in enumerate(top_general, start=1)
                ])

                st.dataframe(
                    df_top_general,
                    use_container_width=True,
                    column_config={
                        "Ranking": st.column_config.NumberColumn("Ranking", format="%02d"),
                        "Categoría": st.column_config.TextColumn("Categoría", width="large"),
                        "Llamadas": st.column_config.NumberColumn("Llamadas", format="%d"),
                        "% del Total": st.column_config.TextColumn("% del Total"),
                        "Impacto Operativo": st.column_config.TextColumn("Impacto Operativo", width="medium")
                    }
                )
            else:
                st.info("No se encontraron categorías generales para mostrar.")

            # Análisis de subcategorías destacado (como en el Excel)
            st.subheader("🔍 Análisis de Subcategorías")
            subcategory_df = display_analyzer.generate_subcategory_analysis(detailed_analysis, total_rows)

            if subcategory_df is not None and not subcategory_df.empty:
                st.dataframe(
                    subcategory_df,
                    use_container_width=True,
                    column_config={
                        "Tipo": st.column_config.TextColumn("Tipo", width="medium"),
                        "Subcategoría": st.column_config.TextColumn("Subcategoría", width="large"),
                        "Frecuencia": st.column_config.TextColumn("Frecuencia"),
                        "% Total": st.column_config.TextColumn("% Total"),
                        "Prioridad": st.column_config.TextColumn("Prioridad", width="small")
                    }
                )
            else:
                st.info("No se encontraron subcategorías destacadas para mostrar.")

            # Desglose completo de rutas Categoria General -> Específica -> Subtipo
            st.subheader("🧭 Rutas Completas de Motivos")
            combined_detail = detailed_analysis.get('categoria_combinada_detalle') if detailed_analysis else None

            if combined_detail:
                df_routes = pd.DataFrame(combined_detail)

                # Agregar recuento por ruta y porcentaje
                route_counts = df_routes['ruta_completa'].value_counts().reset_index()
                route_counts.columns = ['ruta_completa', 'Frecuencia']
                route_counts['% del Total'] = route_counts['Frecuencia'] / total_rows * 100 if total_rows > 0 else 0

                # Extraer columnas separadas
                df_routes_display = route_counts.merge(
                    df_routes.drop_duplicates('ruta_completa')[['ruta_completa', 'categoria_general', 'categoria_especifica', 'subtipo']],
                    on='ruta_completa',
                    how='left'
                )

                df_routes_display['% del Total'] = df_routes_display['% del Total'].apply(lambda x: f"{x:.2f}%")

                st.dataframe(
                    df_routes_display,
                    use_container_width=True,
                    column_config={
                        'categoria_general': st.column_config.TextColumn('Categoría General', width='large'),
                        'categoria_especifica': st.column_config.TextColumn('Categoría Específica', width='large'),
                        'subtipo': st.column_config.TextColumn('Subtipo', width='large'),
                        'Frecuencia': st.column_config.NumberColumn('Frecuencia', format='%d'),
                        '% del Total': st.column_config.TextColumn('% del Total'),
                        'ruta_completa': st.column_config.TextColumn('Ruta Completa', width='large')
                    }
                )
            else:
                st.info("No se pudo generar el desglose completo de rutas para los motivos de llamada.")

            # Desglose de categorías por agente instalador
            st.subheader("👷‍♂️ Agentes Instaladores por Categoría")
            installer_detail = detailed_analysis.get('agente_instalador_detalle') if detailed_analysis else None

            if installer_detail:
                df_installers = pd.DataFrame(installer_detail)
                df_installers['% del Agente'] = df_installers['Porcentaje_Agente'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)

                st.dataframe(
                    df_installers,
                    use_container_width=True,
                    column_config={
                        'agente_instalador': st.column_config.TextColumn('Agente Instalador', width='large'),
                        'categoria_general': st.column_config.TextColumn('Categoría General', width='medium'),
                        'categoria_especifica': st.column_config.TextColumn('Categoría Específica', width='large'),
                        'subtipo': st.column_config.TextColumn('Subtipo', width='large'),
                        'Frecuencia': st.column_config.NumberColumn('Frecuencia', format='%d'),
                        '% del Agente': st.column_config.TextColumn('% del Agente'),
                        'ruta_completa': st.column_config.TextColumn('Ruta Completa', width='large')
                    }
                )
            else:
                st.info("No se pudo generar el análisis por agentes instaladores.")

        with tab2:
            st.header("📈 Visualizaciones")

            # Gráfico de categorías principales
            st.subheader("🏆 Ranking de Categorías Generales")
            fig_categories = crear_grafico_categorias(total_counter, "Top 20 Categorías Más Frecuentes")
            if fig_categories:
                st.plotly_chart(fig_categories, use_container_width=True)

            # Gráfico de distribución
            st.subheader("📊 Distribución de Llamadas por Volumen")
            fig_distribution = crear_grafico_distribucion(total_counter, total_rows)
            if fig_distribution:
                st.plotly_chart(fig_distribution, use_container_width=True)

            # Gráfico de subcategorías
            st.subheader("🔍 Top Subcategorías")
            subcategory_df = display_analyzer.generate_subcategory_analysis(detailed_analysis, total_rows)
            if subcategory_df is not None and not subcategory_df.empty:
                # Tomar top 15 subcategorías
                top_subcats = subcategory_df.head(15).copy()
                fig_subcats = px.bar(
                    top_subcats,
                    y='Subcategoría',
                    x='Frecuencia',
                    orientation='h',
                    title='Top 15 Subcategorías Más Frecuentes',
                    labels={'Frecuencia': 'Número de Llamadas', 'Subcategoría': 'Subcategoría'},
                    color='Frecuencia',
                    color_continuous_scale='Viridis',
                    text='Frecuencia'
                )
                fig_subcats.update_layout(
                    height=500,
                    yaxis={'categoryorder': 'total ascending'},
                    showlegend=False
                )
                fig_subcats.update_traces(texttemplate='%{text}', textposition='outside')
                st.plotly_chart(fig_subcats, use_container_width=True)

            # Gráfico de rutas completas
            st.subheader("🧭 Top Rutas Completas de Motivos")
            combined_detail = detailed_analysis.get('categoria_combinada_detalle') if detailed_analysis else None
            if combined_detail:
                df_routes = pd.DataFrame(combined_detail)
                route_counts = df_routes['ruta_completa'].value_counts().head(15).reset_index()
                route_counts.columns = ['ruta_completa', 'Frecuencia']
                
                fig_routes = px.bar(
                    route_counts,
                    y='ruta_completa',
                    x='Frecuencia',
                    orientation='h',
                    title='Top 15 Rutas Completas Más Frecuentes',
                    labels={'Frecuencia': 'Número de Llamadas', 'ruta_completa': 'Ruta Completa'},
                    color='Frecuencia',
                    color_continuous_scale='Blues',
                    text='Frecuencia'
                )
                fig_routes.update_layout(
                    height=600,
                    yaxis={'categoryorder': 'total ascending'},
                    showlegend=False
                )
                fig_routes.update_traces(texttemplate='%{text}', textposition='outside')
                st.plotly_chart(fig_routes, use_container_width=True)

            # Gráfico de agentes instaladores
            st.subheader("👷‍♂️ Top Agentes Instaladores por Frecuencia")
            installer_detail = detailed_analysis.get('agente_instalador_detalle') if detailed_analysis else None
            if installer_detail:
                df_installers = pd.DataFrame(installer_detail)
                # Agrupar por agente instalador y sumar frecuencias
                installers_grouped = df_installers.groupby('agente_instalador')['Frecuencia'].sum().reset_index()
                installers_grouped = installers_grouped.sort_values('Frecuencia', ascending=False).head(15)
                
                fig_installers = px.bar(
                    installers_grouped,
                    x='agente_instalador',
                    y='Frecuencia',
                    title='Top 15 Agentes Instaladores por Llamadas',
                    labels={'Frecuencia': 'Total de Llamadas', 'agente_instalador': 'Agente Instalador'},
                    color='Frecuencia',
                    color_continuous_scale='Oranges',
                    text='Frecuencia'
                )
                fig_installers.update_layout(
                    height=500,
                    xaxis={'categoryorder': 'total descending'},
                    showlegend=False
                )
                fig_installers.update_traces(texttemplate='%{text}', textposition='outside')
                fig_installers.update_xaxes(tickangle=-45)
                st.plotly_chart(fig_installers, use_container_width=True)
                
                # Gráfico de categorías por agente (top 10 agentes)
                st.subheader("📊 Distribución de Categorías por Agente Instalador")
                top_10_agents = installers_grouped.head(10)['agente_instalador'].tolist()
                df_agents_categories = df_installers[df_installers['agente_instalador'].isin(top_10_agents)]
                
                # Agrupar por agente y categoría general
                agents_cat_grouped = df_agents_categories.groupby(['agente_instalador', 'categoria_general'])['Frecuencia'].sum().reset_index()
                
                fig_agents_cat = px.bar(
                    agents_cat_grouped,
                    x='agente_instalador',
                    y='Frecuencia',
                    color='categoria_general',
                    title='Distribución de Categorías por Top 10 Agentes',
                    labels={'Frecuencia': 'Número de Llamadas', 'agente_instalador': 'Agente', 'categoria_general': 'Categoría'},
                    barmode='stack'
                )
                fig_agents_cat.update_layout(height=500)
                fig_agents_cat.update_xaxes(tickangle=-45)
                st.plotly_chart(fig_agents_cat, use_container_width=True)
    else:
        # Pantalla inicial cuando no hay archivos
        st.info("👆 Sube archivos Excel desde la barra lateral para comenzar el análisis.")

        # Información de ayuda
        with st.expander("ℹ️ Información de uso"):
            st.markdown("""
            ### Cómo usar esta aplicación:

            1. **📁 Subir archivos**: Usa la barra lateral para subir uno o más archivos Excel con datos de llamadas
            2. **📅 Filtros opcionales**: Puedes aplicar filtros de fecha para analizar períodos específicos
            3. **🚀 Ejecutar**: Haz clic en "Ejecutar Análisis" para procesar los datos
            4. **💾 Descargar**: Obtén el reporte completo en formato Excel

            ### Formatos de archivo soportados:
            - `.xlsx` (Excel moderno)
            - `.xls` (Excel antiguo)
            """)

if __name__ == "__main__":
    main()
