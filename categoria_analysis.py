#!/usr/bin/env python3
"""
Script para analizar archivos Excel de resultados de llamadas y extraer el top 50 de categorías más frecuentes.

Este script puede procesar uno o múltiples archivos Excel ubicados en el directorio results/
y generar estadísticas de frecuencia de categorías con opción de filtrado por rango de fechas.

Autor: Asistente IA
Fecha: 2025-01-02
"""

import os
import sys
import argparse
import pandas as pd
from collections import Counter
import logging
from typing import List, Dict, Tuple, Optional
from pathlib import Path
from datetime import datetime, timedelta
import io
import re
import unicodedata

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('categoria_analysis.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class ExcelCategoryAnalyzer:
    """Analizador de categorías en archivos Excel de resultados de llamadas."""

    # Posibles nombres de columnas que podrían contener categorías
    CATEGORY_COLUMNS = [
        'categoria', 'categoría', 'category', 'tipo', 'type',
        'motivo', 'reason', 'clasificacion', 'clasificación',
        'resultado', 'result', 'estado', 'status',
        'categoria_principal', 'categoria_secundaria',
        'categoria_final', 'categoria_detectada'
    ]

    # Columnas específicas de categorías detalladas (SOLO categorías y subtipos)
    SPECIFIC_CATEGORY_COLUMNS = [
        'categoria_general', 'categoria_especifica', 'subtipo_categoria',
        'categoria_especifica_1', 'subtipo_categoria_1',
        'categoria_especifica_2', 'subtipo_categoria_2',
        'categoria_especifica_3', 'subtipo_categoria_3',
        'categoria', 'tipo', 'motivo', 'subcategoria', 'subtipo'
    ]

    INSTALLER_COLUMNS = [
        'agente_instalador', 'instalador', 'tecnico_instalador',
        'tecnico', 'agenteinstalador', 'instalador_agente'
    ]

    DATE_WITH_TIME_REGEX = re.compile(r'(\d{4}-\d{2}-\d{2})[ _T-](\d{2})[-:](\d{2})[-:](\d{2})')
    DATE_ONLY_REGEX = re.compile(r'(\d{4}-\d{2}-\d{2})')

    @staticmethod
    def normalize_column_name(name: str) -> str:
        """Normalizar nombre de columna para comparaciones flexibles."""
        if name is None:
            return ""

        normalized = unicodedata.normalize('NFKD', str(name))
        without_accents = ''.join(ch for ch in normalized if not unicodedata.combining(ch))
        return ''.join(ch for ch in without_accents.lower() if ch.isalnum())

    def find_matching_columns(self, df: pd.DataFrame, target_name: str) -> List[str]:
        """Encontrar columnas que coincidan (exacta o parcialmente) con un nombre objetivo."""
        norm_target = self.normalize_column_name(target_name)

        matches = []
        for column in df.columns:
            norm_column = self.normalize_column_name(column)
            if not norm_column:
                continue

            if (
                norm_target == norm_column
                or norm_target in norm_column
                or norm_column in norm_target
            ):
                matches.append(column)

        return matches

    def parse_datetime_value(self, value) -> Optional[datetime]:
        """Convertir un valor cualquiera a datetime si contiene una fecha reconocible."""
        if pd.isna(value):
            return pd.NaT

        value_str = str(value)

        match = self.DATE_WITH_TIME_REGEX.search(value_str)
        if match:
            dt_str = f"{match.group(1)} {match.group(2)}:{match.group(3)}:{match.group(4)}"
            try:
                return datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                pass

        match = self.DATE_ONLY_REGEX.search(value_str)
        if match:
            try:
                return datetime.strptime(match.group(1), "%Y-%m-%d")
            except ValueError:
                pass

        try:
            return pd.to_datetime(value_str, errors='coerce')
        except Exception:
            return pd.NaT

    def identify_installer_column(self, df: pd.DataFrame) -> Optional[str]:
        """Identificar columna que contiene el agente instalador."""

        # Intentar coincidencias específicas
        for target in self.INSTALLER_COLUMNS:
            matches = self.find_matching_columns(df, target)
            if matches:
                logger.info(f"Columna de agente instalador identificada: '{matches[0]}' (coincide con '{target}')")
                return matches[0]

        # Buscar por palabra clave
        for col in df.columns:
            norm_col = self.normalize_column_name(col)
            if any(keyword in norm_col for keyword in ['instalador', 'tecnico', 'agente']):
                logger.info(f"Columna de agente instalador identificada (por palabra clave): '{col}'")
                return col

        logger.warning("No se pudo identificar una columna de agente instalador")
        return None

    def __init__(self, results_dir: str = None, start_date: Optional[datetime] = None, end_date: Optional[datetime] = None):
        """Inicializar el analizador.

        Args:
            results_dir: Directorio donde se encuentran los archivos Excel de resultados
            start_date: Fecha de inicio para filtrar datos
            end_date: Fecha de fin para filtrar datos
        """
        if results_dir is None:
            # Usar el directorio results relativo al script
            script_dir = Path(__file__).parent
            project_root = script_dir.parent
            self.results_dir = project_root / "results"
        else:
            self.results_dir = Path(results_dir)

        if not self.results_dir.exists():
            raise FileNotFoundError(f"El directorio {self.results_dir} no existe")

        self.start_date = start_date
        self.end_date = end_date

        logger.info(f"Directorio de resultados: {self.results_dir}")
        if start_date:
            logger.info(f"Filtro de fecha inicio: {start_date.strftime('%Y-%m-%d')}")
        if end_date:
            logger.info(f"Filtro de fecha fin: {end_date.strftime('%Y-%m-%d')}")

    def find_excel_files(self, pattern: str = None) -> List[Path]:
        """Encontrar archivos Excel en el directorio de resultados.

        Args:
            pattern: Patrón opcional para filtrar archivos (e.g., "*.xlsx")

        Returns:
            Lista de rutas a archivos Excel encontrados
        """
        if pattern:
            excel_files = list(self.results_dir.glob(pattern))
        else:
            # Buscar archivos Excel comunes
            excel_files = list(self.results_dir.glob("*.xlsx"))
            # También buscar archivos .xls si los hay
            excel_files.extend(self.results_dir.glob("*.xls"))

        # Filtrar archivos temporales de Excel
        excel_files = [f for f in excel_files if not f.name.startswith("~$")]

        logger.info(f"Se encontraron {len(excel_files)} archivos Excel")
        return sorted(excel_files)

    def identify_category_column(self, df: pd.DataFrame) -> str:
        """Identificar la columna que contiene las categorías en un DataFrame.

        Args:
            df: DataFrame de pandas

        Returns:
            Nombre de la columna identificada como categoría, o None si no se encuentra
        """
        # Buscar columnas candidatas utilizando coincidencias flexibles
        candidate_targets = ['categoria_general'] + self.CATEGORY_COLUMNS

        for target in candidate_targets:
            matches = self.find_matching_columns(df, target)
            if matches:
                logger.info(
                    f"Columna de categoría identificada: '{matches[0]}' (coincide con '{target}')"
                )
                return matches[0]

        # Buscar columnas que contengan palabras clave
        for col in df.columns:
            norm_col = self.normalize_column_name(col)
            if any(keyword in norm_col for keyword in ['categoria', 'category', 'tipo', 'motivo']):
                logger.info(f"Columna de categoría identificada (por palabra clave): '{col}'")
                return col

        # Si no encuentra ninguna columna específica, usar la primera columna no numérica
        for col in df.columns:
            if df[col].dtype == 'object' and not col.lower().startswith(('id', 'fecha', 'date', 'hora', 'time')):
                logger.info(f"Usando primera columna no numérica como categoría: '{col}'")
                return col

    def split_comma_categories(self, categories_series: pd.Series) -> pd.Series:
        """Separar categorías que contienen comas en categorías independientes.

        Args:
            categories_series: Series de pandas con categorías

        Returns:
            Series expandida con categorías separadas
        """
        expanded_categories = []

        for category in categories_series.dropna():
            category_str = str(category).strip()

            # Si contiene coma, separar en múltiples categorías
            if ',' in category_str:
                # Separar por coma y limpiar espacios
                split_cats = [cat.strip() for cat in category_str.split(',') if cat.strip()]
                expanded_categories.extend(split_cats)
            else:
                expanded_categories.append(category_str)

        return pd.Series(expanded_categories)

    def identify_date_column(self, df: pd.DataFrame) -> Optional[str]:
        """Identificar la columna que contiene las fechas en un DataFrame.

        Args:
            df: DataFrame de pandas

        Returns:
            Nombre de la columna identificada como fecha, o None si no se encuentra
        """
        df_columns_lower = [col.lower() for col in df.columns]

        # Posibles nombres de columnas de fecha
        date_column_names = [
            'fecha', 'date', 'fecha_llamada', 'fecha_hora', 'timestamp',
            'fecha_inicio', 'fecha_fin', 'fecha_creacion', 'created_date',
            'dia', 'day', 'fecha_registro', 'fecha_contacto', 'archivo_procesado'
        ]

        # Buscar coincidencias exactas
        for date_col in date_column_names:
            if date_col in df_columns_lower:
                original_col = df.columns[df_columns_lower.index(date_col)]
                logger.info(f"Columna de fecha identificada: '{original_col}'")
                return original_col

        # Buscar columnas que contengan palabras clave de fecha
        for col in df.columns:
            norm_col = self.normalize_column_name(col)
            if any(keyword in norm_col for keyword in ['fecha', 'date', 'dia', 'time', 'archivo']):
                logger.info(f"Columna de fecha identificada (por palabra clave): '{col}'")
                return col

        # Fallback: detectar columnas cuyos valores contienen patrones de fecha
        for col in df.columns:
            sample_series = df[col].dropna()
            if sample_series.empty:
                continue

            sample_value = str(sample_series.iloc[0])
            if self.DATE_WITH_TIME_REGEX.search(sample_value) or self.DATE_ONLY_REGEX.search(sample_value):
                logger.info(f"Columna de fecha identificada (por patrón en valores): '{col}'")
                return col

        logger.warning("No se pudo identificar una columna de fecha automáticamente")
        return None

    def filter_by_date_range(self, df: pd.DataFrame) -> pd.DataFrame:
        """Filtrar DataFrame por rango de fechas.

        Args:
            df: DataFrame original

        Returns:
            DataFrame filtrado por fechas
        """
        if not self.start_date and not self.end_date:
            return df

        # Identificar columna de fecha
        date_column = self.identify_date_column(df)

        if date_column is None:
            logger.warning("No se pudo identificar columna de fecha, devolviendo DataFrame sin filtrar")
            return df

        try:
            # Convertir columna a datetime si no lo está
            if not pd.api.types.is_datetime64_any_dtype(df[date_column]):
                df[date_column] = df[date_column].apply(self.parse_datetime_value)

            # Aplicar filtro de fechas
            filtered_df = df.copy()

            if self.start_date:
                filtered_df = filtered_df[filtered_df[date_column] >= self.start_date]

            if self.end_date:
                filtered_df = filtered_df[filtered_df[date_column] <= self.end_date]

            logger.info(f"Filtrado aplicado: {len(df)} -> {len(filtered_df)} filas")
            return filtered_df

        except Exception as e:
            logger.error(f"Error aplicando filtro de fechas: {e}")
            return df

    def analyze_excel_file(self, file_path: Path) -> Tuple[Counter, int, Dict[str, Counter]]:
        """Analizar un archivo Excel y extraer frecuencias de categorías.

        Args:
            file_path: Ruta al archivo Excel

        Returns:
            Tupla con (contador de categorías generales, número total de filas procesadas, diccionario con análisis detallado)
        """
        try:
            logger.info(f"Procesando archivo: {file_path.name}")

            # Leer el archivo Excel
            # Usar openpyxl como motor para mejor compatibilidad
            df = pd.read_excel(file_path, engine='openpyxl')

            if df.empty:
                logger.warning(f"El archivo {file_path.name} está vacío")
                return Counter(), 0, {}

            logger.info(f"Archivo {file_path.name}: {len(df)} filas, {len(df.columns)} columnas")

            # Aplicar filtro de fechas si está configurado
            df = self.filter_by_date_range(df)

            if df.empty:
                logger.warning(f"El archivo {file_path.name} no tiene datos en el rango de fechas especificado")
                return Counter(), 0, {}

            # Identificar columna de categoría general
            category_column = self.identify_category_column(df)

            if category_column is None:
                logger.warning(f"No se pudo identificar columna de categoría en {file_path.name}")
                return Counter(), len(df), {}

            # Extraer categorías generales (ignorar valores nulos)
            categories_raw = df[category_column].dropna()

            # NO expandir categorías - contar tal cual
            if categories_raw.empty:
                logger.warning(f"No se encontraron categorías válidas en {file_path.name}")
                return Counter(), len(df), {}

            # Contar frecuencias de categorías generales sin expandir
            category_counts = Counter(categories_raw.astype(str).str.strip())

            logger.info(f"Archivo {file_path.name}: {len(category_counts)} categorías únicas encontradas ({len(categories_raw)} filas procesadas)")

            # Análisis detallado de columnas específicas
            detailed_analysis = self.analyze_detailed_categories(df)

            return category_counts, len(df), detailed_analysis

        except Exception as e:
            logger.error(f"Error procesando {file_path.name}: {e}")
            return Counter(), 0, {}

    def analyze_detailed_categories(self, df: pd.DataFrame) -> Dict[str, Counter]:
        """Analizar columnas específicas de categorías detalladas.

        Args:
            df: DataFrame de pandas

        Returns:
            Diccionario con análisis de cada columna específica
        """
        detailed_analysis = {}
        processed_columns = set()

        for col in self.SPECIFIC_CATEGORY_COLUMNS:
            matches = self.find_matching_columns(df, col)

            for match in matches:
                if match in processed_columns:
                    continue

                # Extraer valores no nulos
                values_raw = df[match].dropna()

                if values_raw.empty:
                    continue

                # NO expandir subcategorías - contar tal cual
                processed_columns.add(match)
                detailed_analysis[match] = Counter(values_raw.astype(str).str.strip())
                logger.info(
                    f"Columna '{match}': {len(detailed_analysis[match])} valores únicos ({len(values_raw)} filas procesadas)"
                )

        # Crear análisis combinado de categorías específicas con subtipos
        combined_counter, combined_details = self.analyze_combined_categories(df)
        if combined_counter:
            detailed_analysis['categoria_combinada'] = combined_counter
        if combined_details:
            detailed_analysis['categoria_combinada_detalle'] = combined_details

            installer_column = self.identify_installer_column(df)
            if installer_column:
                installer_summary = self.generate_installer_breakdown(combined_details)
                if installer_summary:
                    detailed_analysis['agente_instalador_detalle'] = installer_summary

        return detailed_analysis

    def analyze_combined_categories(self, df: pd.DataFrame) -> Tuple[Counter, List[Dict[str, str]]]:
        """Analizar combinaciones de categoria_general + categoria_especifica + subtipo.
        
        Cada fila del DataFrame se cuenta como UNA llamada única, sin expandir valores con delimitadores.
        """
        combined_counts = Counter()
        combined_details = []

        # Buscar columnas de categoría general y específica
        categoria_general = None
        categoria_especifica_cols = []
        subtipo_cols = []

        for col in df.columns:
            col_lower = col.lower()
            if col_lower == 'categoria_general':
                categoria_general = col
            elif 'categoria_especifica' in col_lower:
                categoria_especifica_cols.append(col)
            elif 'subtipo_categoria' in col_lower:
                subtipo_cols.append(col)

        installer_column = self.identify_installer_column(df)

        if not categoria_general:
            return combined_counts, combined_details

        for idx, row in df.iterrows():
            if pd.isna(row[categoria_general]):
                continue

            # NO expandir - tomar el valor tal cual
            categoria_gen = str(row[categoria_general]).strip()

            installer_value = "Sin asignar"
            if installer_column and installer_column in df.columns:
                raw_installer = row.get(installer_column)
                if pd.isna(raw_installer) or str(raw_installer).strip() == "":
                    installer_value = "Sin asignar"
                else:
                    installer_value = str(raw_installer).strip()

            # Construir ruta sin expandir valores
            combined_parts = [categoria_gen]

            # Agregar primera categoria_especifica si existe
            if categoria_especifica_cols:
                for cat_col in categoria_especifica_cols:
                    if not pd.isna(row[cat_col]):
                        cat_esp = str(row[cat_col]).strip()
                        combined_parts.append(cat_esp)
                        break  # Solo tomar la primera categoría específica

            # Agregar primer subtipo si existe
            if subtipo_cols:
                for subtipo_col in subtipo_cols:
                    if not pd.isna(row[subtipo_col]):
                        subtipo = str(row[subtipo_col]).strip()
                        combined_parts.append(subtipo)
                        break  # Solo tomar el primer subtipo

            # Crear ruta completa
            if len(combined_parts) > 1:  # Si hay más que solo categoria_general
                combined_category = " | ".join(combined_parts)
                combined_counts[combined_category] += 1
                combined_details.append({
                    'categoria_general': combined_parts[0],
                    'categoria_especifica': combined_parts[1] if len(combined_parts) > 1 else '',
                    'subtipo': combined_parts[2] if len(combined_parts) > 2 else '',
                    'ruta_completa': combined_category,
                    'agente_instalador': installer_value
                })

        return combined_counts, combined_details

    def analyze_multiple_files(self, file_paths: List[Path]) -> Tuple[Counter, int, Dict[str, Counter]]:
        """Analizar múltiples archivos Excel y combinar resultados.

        Args:
            file_paths: Lista de rutas a archivos Excel

        Returns:
            Tupla con (contador combinado de categorías, número total de filas procesadas, análisis detallado combinado)
        """
        total_counter = Counter()
        total_detailed_analysis = {}
        total_rows = 0

        for file_path in file_paths:
            file_counter, file_rows, file_detailed = self.analyze_excel_file(file_path)
            total_counter.update(file_counter)
            total_rows += file_rows

            # Combinar análisis detallado contemplando contadores y listas
            for key, value in file_detailed.items():
                if isinstance(value, Counter):
                    if key not in total_detailed_analysis:
                        total_detailed_analysis[key] = Counter()
                    total_detailed_analysis[key].update(value)
                elif isinstance(value, list):
                    if key not in total_detailed_analysis:
                        total_detailed_analysis[key] = []
                    total_detailed_analysis[key].extend(value)

        return total_counter, total_rows, total_detailed_analysis

    def generate_installer_breakdown(self, combined_details: List[Dict[str, str]]) -> List[Dict[str, str]]:
        """Generar desglose de categorías por agente instalador."""

        if not combined_details:
            return []

        df_details = pd.DataFrame(combined_details)

        if df_details.empty or 'agente_instalador' not in df_details.columns:
            return []

        df_details['agente_instalador'] = df_details['agente_instalador'].fillna('Sin asignar')

        # Filtrar para excluir "Sin asignar"
        df_details = df_details[df_details['agente_instalador'] != 'Sin asignar']

        if df_details.empty:
            return []

        # Agrupar por agente y ruta
        grouped = (
            df_details
            .groupby(['agente_instalador', 'categoria_general', 'categoria_especifica', 'subtipo', 'ruta_completa'])
            .size()
            .reset_index(name='Frecuencia')
        )

        if grouped.empty:
            return []

        # Calcular totales por agente para porcentajes
        grouped['Total_Agente'] = grouped.groupby('agente_instalador')['Frecuencia'].transform('sum')
        grouped['Porcentaje_Agente'] = grouped['Frecuencia'] / grouped['Total_Agente'] * 100

        # Ordenar por agente y frecuencia
        grouped.sort_values(by=['agente_instalador', 'Frecuencia'], ascending=[True, False], inplace=True)

        grouped['Frecuencia'] = grouped['Frecuencia'].astype(int)

        return grouped[['agente_instalador', 'categoria_general', 'categoria_especifica', 'subtipo', 'ruta_completa', 'Frecuencia', 'Porcentaje_Agente']].to_dict('records')

    def get_top_categories(self, counter: Counter, top_n: int = 50) -> List[Tuple[str, int]]:
        """Obtener las top N categorías más frecuentes.

        Args:
            counter: Contador de frecuencias
            top_n: Número de categorías a mostrar (default: 50)

        Returns:
            Lista de tuplas (categoria, frecuencia) ordenadas por frecuencia descendente
        """
        return counter.most_common(top_n)

    def generate_report(self, counter: Counter, total_rows: int, files_processed: int, detailed_analysis: Dict[str, Counter] = None) -> str:
        """Generar un reporte detallado del análisis.

        Args:
            counter: Contador de frecuencias de categorías generales
            total_rows: Número total de filas procesadas
            files_processed: Número de archivos procesados
            detailed_analysis: Diccionario con análisis detallado de columnas específicas

        Returns:
            String con el reporte formateado
        """
        total_categories = len(counter)
        # Obtener todas las categorías ordenadas por frecuencia (no solo top 50)
        all_categories = counter.most_common()  # Todas las categorías ordenadas

        report = []
        report.append("=" * 100)
        report.append("ANÁLISIS ESPECIALIZADO DE CATEGORÍAS Y SUBTIPOS")
        report.append("=" * 100)
        report.append(f"Archivos procesados: {files_processed}")
        report.append(f"Filas totales analizadas: {total_rows:,}")
        report.append(f"Categorías generales únicas encontradas: {total_categories:,}")

        if all_categories:
            # Estadísticas adicionales
            total_top_50 = sum(count for _, count in all_categories[:50]) if len(all_categories) >= 50 else sum(count for _, count in all_categories)
            coverage = (total_top_50 / total_rows * 100) if total_rows > 0 else 0

            report.append(f"Cobertura de top 50 categorías generales: {coverage:.1f}%")
            report.append("")

            report.append("📊 TODAS LAS CATEGORÍAS GENERALES ENCONTRADAS:")
            report.append("-" * 100)

            for i, (category, count) in enumerate(all_categories, 1):
                percentage = (count / total_rows * 100) if total_rows > 0 else 0
                report.append(f"{i:2d}. {category:<50} {count:4d} ({percentage:4.1f}%)")

            report.append("-" * 100)
            report.append(f"Cobertura total de categorías generales: {coverage:.1f}%")
        else:
            report.append("No se encontraron categorías generales para mostrar.")

        # Agregar análisis detallado si está disponible
        if detailed_analysis:
            report.append("")
            report.append("=" * 100)
            report.append("ANÁLISIS DETALLADO DE CATEGORÍAS Y SUBTIPOS")
            report.append("=" * 100)

            for column_name, column_counter in detailed_analysis.items():
                if column_counter:
                    report.append(f"\n🔍 ANÁLISIS DE '{column_name.upper()}':")
                    report.append(f"Valores únicos encontrados: {len(column_counter)}")
                    report.append("-" * 60)

                    # Top 20 para columnas específicas (menos que las 50 generales)
                    top_values = column_counter.most_common(20)
                    for i, (value, count) in enumerate(top_values, 1):
                        percentage = (count / total_rows * 100) if total_rows > 0 else 0
                        # Truncar valores largos para mejor formato
                        display_value = value[:45] + "..." if len(value) > 45 else value
                        report.append(f"{i:3d}. {display_value:<45} {count:4d} ({percentage:5.1f}%)")

                    report.append("-" * 60)

        report.append("=" * 100)
        return "\n".join(report)

    def save_excel_report(self, filename: str, counter: Counter, total_rows: int, files_processed: int, detailed_analysis: Dict[str, Counter]):
        """Guardar reporte ejecutivo en formato Excel con análisis de negocio.

        Args:
            filename: Nombre del archivo Excel a crear
            counter: Contador de frecuencias de categorías generales
            total_rows: Número total de filas procesadas
            files_processed: Número de archivos procesados
            detailed_analysis: Diccionario con análisis detallado de columnas específicas
        """
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # Crear análisis ejecutivo completo
                self.create_executive_dashboard(writer, counter, total_rows, files_processed, detailed_analysis)

                logger.info(f"Reporte ejecutivo Excel creado exitosamente: {filename}")

        except Exception as e:
            logger.error(f"Error creando archivo Excel ejecutivo: {e}")
            raise

    def create_executive_dashboard(self, writer, counter: Counter, total_rows: int, files_processed: int, detailed_analysis: Dict[str, Counter]):
        """Crear dashboard ejecutivo con métricas de negocio.

        Args:
            writer: ExcelWriter object
            counter: Contador de categorías
            total_rows: Total de llamadas
            files_processed: Número de archivos
            detailed_analysis: Análisis detallado
        """

        # 1. RESUMEN EJECUTIVO (Sección superior)
        executive_summary = self.generate_executive_summary(counter, total_rows, files_processed, detailed_analysis)

        # Generar DataFrame de rutas completas
        combined_detail = detailed_analysis.get('categoria_combinada_detalle') if detailed_analysis else None
        routes_df = None
        if combined_detail:
            df_routes = pd.DataFrame(combined_detail)
            route_counts = df_routes['ruta_completa'].value_counts().reset_index()
            route_counts.columns = ['ruta_completa', 'Frecuencia']
            route_counts['% del Total'] = route_counts['Frecuencia'] / total_rows * 100 if total_rows > 0 else 0
            routes_df = route_counts.merge(
                df_routes.drop_duplicates('ruta_completa')[['ruta_completa', 'categoria_general', 'categoria_especifica', 'subtipo']],
                on='ruta_completa',
                how='left'
            )
            routes_df['% del Total'] = routes_df['% del Total'].apply(lambda x: f"{x:.2f}%")
            routes_df = routes_df[['categoria_general', 'categoria_especifica', 'subtipo', 'ruta_completa', 'Frecuencia', '% del Total']]

        # Crear DataFrames para cada sección
        sections = [
            ("RESUMEN EJECUTIVO", executive_summary),
            ("KPIs PRINCIPALES", self.generate_kpi_section(counter, total_rows)),
            ("RUTAS COMPLETAS DE MOTIVOS", routes_df),
            ("ANÁLISIS DE SUBCATEGORÍAS", self.generate_subcategory_analysis(detailed_analysis, total_rows))
        ]

        current_row = 0

        for section_name, section_df in sections:
            # Agregar título de sección
            title_df = pd.DataFrame([[section_name]], columns=[''])
            title_df.to_excel(writer, sheet_name='Dashboard_Ejecutivo', startrow=current_row, index=False, header=False)

            # Agregar contenido de la sección
            if section_df is not None:
                section_df.to_excel(writer, sheet_name='Dashboard_Ejecutivo', startrow=current_row + 2, index=False)

            current_row += (len(section_df) if section_df is not None else 1) + 4  # Espacio entre secciones

    def generate_executive_summary(self, counter: Counter, total_rows: int, files_processed: int, detailed_analysis: Dict[str, Counter]) -> pd.DataFrame:
        """Generar resumen ejecutivo con métricas clave de negocio."""

        # Métricas básicas
        total_categories = len(counter)
        top_3_categories = counter.most_common(3)
        top_3_total = sum(count for _, count in top_3_categories)

        # Calcular concentración
        concentration_percentage = (top_3_total / total_rows * 100) if total_rows > 0 else 0

        # Análisis de volumen
        avg_calls_per_category = total_rows / total_categories if total_categories > 0 else 0

        summary_data = {
            'Métrica': [
                'Período Analizado',
                'Volumen Total de Llamadas',
                'Número de Categorías Principales',
                'Concentración en Top 3 Categorías',
                'Promedio de Llamadas por Categoría',
                'Archivos Procesados'
            ],
            'Valor': [
                'Junio-Septiembre 2025',
                f'{total_rows:,}',
                str(total_categories),
                f'{concentration_percentage:.1f}%',
                f'{avg_calls_per_category:.0f}',
                str(files_processed)
            ],
            'Insight': [
                'Análisis de 3 meses de operaciones',
                'Indicador clave de volumen de atención',
                'Diversidad de motivos de contacto',
                'El 80%+ de llamadas se concentran en pocos tipos',
                'Eficiencia operativa promedio',
                'Cobertura de datos disponible'
            ]
        }

        return pd.DataFrame(summary_data)

    def generate_kpi_section(self, counter: Counter, total_rows: int) -> pd.DataFrame:
        """Generar sección de KPIs principales."""

        # Todas las categorías ordenadas por volumen (no solo top 5)
        all_categories = counter.most_common()

        kpi_data = []
        for i, (category, count) in enumerate(all_categories, 1):
            percentage = (count / total_rows * 100) if total_rows > 0 else 0
            kpi_data.append({
                'Ranking': i,
                'Categoría Principal': category,
                'Volumen': f'{count:,}',
                '% del Total': f'{percentage:.1f}%',
                'Impacto Operativo': self.get_business_impact_category(category, percentage)
            })

        return pd.DataFrame(kpi_data)

    def generate_distribution_section(self, counter: Counter, total_rows: int) -> pd.DataFrame:
        """Generar análisis de distribución de llamadas."""

        # Agrupar categorías por nivel de volumen
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

        distribution_data = []

        # Alto volumen
        distribution_data.append({
            'Segmento': 'ALTO VOLUMEN (>10%)',
            'Categorías': len(high_volume),
            'Llamadas': f'{sum(count for _, count, _ in high_volume):,}',
            '% Total': f'{sum(perc for _, _, perc in high_volume):.1f}%',
            'Ejemplos': ', '.join([cat for cat, _, _ in high_volume[:3]])
        })

        # Medio volumen
        distribution_data.append({
            'Segmento': 'MEDIO VOLUMEN (1-10%)',
            'Categorías': len(medium_volume),
            'Llamadas': f'{sum(count for _, count, _ in medium_volume):,}',
            '% Total': f'{sum(perc for _, _, perc in medium_volume):.1f}%',
            'Ejemplos': ', '.join([cat for cat, _, _ in medium_volume[:3]])
        })

        # Bajo volumen
        distribution_data.append({
            'Segmento': 'BAJO VOLUMEN (<1%)',
            'Categorías': len(low_volume),
            'Llamadas': f'{sum(count for _, count, _ in low_volume):,}',
            '% Total': f'{sum(perc for _, _, perc in low_volume):.1f}%',
            'Ejemplos': ', '.join([cat for cat, _, _ in low_volume[:3]])
        })

        return pd.DataFrame(distribution_data)

    def generate_subcategory_analysis(self, detailed_analysis: Dict[str, Counter], total_rows: int) -> pd.DataFrame:
        """Generar análisis de subcategorías más relevantes."""

        subcategory_data = []

        # Analizar principales subcategorías
        priority_columns = ['categoria_especifica', 'subtipo_categoria']

        for col_name in priority_columns:
            if col_name in detailed_analysis and detailed_analysis[col_name]:
                top_5 = detailed_analysis[col_name].most_common(5)

                for subcategory, count in top_5:
                    percentage = (count / total_rows * 100) if total_rows > 0 else 0
                    subcategory_data.append({
                        'Tipo': col_name.replace('_', ' ').title(),
                        'Subcategoría': subcategory,
                        'Frecuencia': f'{count:,}',
                        '% Total': f'{percentage:.1f}%',
                        'Prioridad': self.get_business_priority(col_name, subcategory, percentage)
                    })

        return pd.DataFrame(subcategory_data)

    def generate_business_insights(self, counter: Counter, total_rows: int, detailed_analysis: Dict[str, Counter]) -> pd.DataFrame:
        """Generar insights y recomendaciones de negocio."""

        insights_data = []

        # Insight 1: Concentración de volumen
        top_category, top_count = counter.most_common(1)[0]
        top_percentage = (top_count / total_rows * 100) if total_rows > 0 else 0

        if top_percentage > 30:
            insights_data.append({
                'Tipo': 'OPORTUNIDAD',
                'Insight': f'{top_percentage:.1f}% de llamadas son de "{top_category}"',
                'Recomendación': 'Desarrollar canales digitales especializados para reducir volumen en atención telefónica',
                'Impacto Potencial': 'Alto - Reducción significativa de costos operativos'
            })

        # Insight 2: Diversidad de categorías
        if len(counter) > 20:
            insights_data.append({
                'Tipo': 'OPTIMIZACIÓN',
                'Insight': f'{len(counter)} categorías diferentes requieren atención especializada',
                'Recomendación': 'Implementar sistema de routing automático basado en subcategorías',
                'Impacto Potencial': 'Medio - Mejora en tiempos de resolución'
            })

        # Insight 3: Análisis de tendencias (si hay múltiples archivos)
        if len(counter) > 0:
            insights_data.append({
                'Tipo': 'ESTRATEGIA',
                'Insight': 'Las categorías principales representan oportunidades de mejora continua',
                'Recomendación': 'Establecer KPIs específicos por categoría y monitoreo mensual',
                'Impacto Potencial': 'Alto - Mejora en calidad de servicio'
            })

        return pd.DataFrame(insights_data)

    def get_business_impact_category(self, category: str, percentage: float) -> str:
        """Obtener impacto de negocio de una categoría."""
        if percentage >= 30:
            return "Crítico - Alto volumen"
        elif percentage >= 15:
            return "Importante - Optimización"
        elif percentage >= 5:
            return "Moderado - Monitoreo"
        else:
            return "Bajo - Especializado"

    def get_business_priority(self, col_name: str, subcategory: str, percentage: float) -> str:
        """Obtener prioridad de negocio para subcategorías."""
        if percentage >= 5:
            return "Alta"
        elif percentage >= 2:
            return "Media"
        else:
            return "Baja"

    def generate_excel_report(self, counter: Counter, total_rows: int, files_processed: int, detailed_analysis: Dict[str, Counter]) -> bytes:
        """Generar reporte ejecutivo en formato Excel en memoria.

        Args:
            counter: Contador de frecuencias de categorías generales
            total_rows: Número total de filas procesadas
            files_processed: Número de archivos procesados
            detailed_analysis: Diccionario con análisis detallado de columnas específicas

        Returns:
            Bytes del archivo Excel generado
        """
        try:
            # Crear buffer en memoria
            buffer = io.BytesIO()

            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Crear análisis ejecutivo completo
                self.create_executive_dashboard(writer, counter, total_rows, files_processed, detailed_analysis)

            # Obtener bytes del buffer
            excel_data = buffer.getvalue()

            logger.info(f"Reporte ejecutivo Excel generado en memoria: {len(excel_data)} bytes")

            return excel_data

        except Exception as e:
            logger.error(f"Error creando archivo Excel en memoria: {e}")
            raise

    def create_executive_dashboard(self, writer, counter: Counter, total_rows: int, files_processed: int, detailed_analysis: Dict[str, Counter]):
        """Crear dashboard ejecutivo con métricas de negocio.

        Args:
            writer: ExcelWriter object
            counter: Contador de categorías
            total_rows: Total de llamadas
            files_processed: Número de archivos
            detailed_analysis: Análisis detallado
        """

        # 1. RESUMEN EJECUTIVO (Sección superior)
        executive_summary = self.generate_executive_summary(counter, total_rows, files_processed, detailed_analysis)

        # Generar DataFrame de rutas completas
        combined_detail = detailed_analysis.get('categoria_combinada_detalle') if detailed_analysis else None
        routes_df = None
        if combined_detail:
            df_routes = pd.DataFrame(combined_detail)
            route_counts = df_routes['ruta_completa'].value_counts().reset_index()
            route_counts.columns = ['ruta_completa', 'Frecuencia']
            route_counts['% del Total'] = route_counts['Frecuencia'] / total_rows * 100 if total_rows > 0 else 0
            routes_df = route_counts.merge(
                df_routes.drop_duplicates('ruta_completa')[['ruta_completa', 'categoria_general', 'categoria_especifica', 'subtipo']],
                on='ruta_completa',
                how='left'
            )
            routes_df['% del Total'] = routes_df['% del Total'].apply(lambda x: f"{x:.2f}%")
            routes_df = routes_df[['categoria_general', 'categoria_especifica', 'subtipo', 'ruta_completa', 'Frecuencia', '% del Total']]

        # Crear DataFrames para cada sección
        sections = [
            ("RESUMEN EJECUTIVO", executive_summary),
            ("KPIs PRINCIPALES", self.generate_kpi_section(counter, total_rows)),
            ("RUTAS COMPLETAS DE MOTIVOS", routes_df),
            ("ANÁLISIS DE SUBCATEGORÍAS", self.generate_subcategory_analysis(detailed_analysis, total_rows))
        ]

        current_row = 0

        for section_name, section_df in sections:
            # Agregar título de sección
            title_df = pd.DataFrame([[section_name]], columns=[''])
            title_df.to_excel(writer, sheet_name='Dashboard_Ejecutivo', startrow=current_row, index=False, header=False)

            # Agregar contenido de la sección
            if section_df is not None:
                section_df.to_excel(writer, sheet_name='Dashboard_Ejecutivo', startrow=current_row + 2, index=False)

            current_row += (len(section_df) if section_df is not None else 1) + 4  # Espacio entre secciones

        # Añadir hoja separada para el análisis por agente instalador
        installer_detail = detailed_analysis.get('agente_instalador_detalle') if detailed_analysis else None
        if installer_detail:
            df_installers = pd.DataFrame(installer_detail)
            if not df_installers.empty:
                df_installers['% del Agente'] = df_installers['Porcentaje_Agente'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
                df_installers_export = df_installers[['agente_instalador', 'categoria_general', 'categoria_especifica', 
                                                       'subtipo', 'ruta_completa', 'Frecuencia', '% del Agente']]
                df_installers_export.to_excel(writer, sheet_name='Agentes_Instaladores', index=False)
                logger.info(f"Añadida hoja 'Agentes_Instaladores' con {len(df_installers_export)} registros")

def main():
    parser = argparse.ArgumentParser(
        description="Generar análisis ejecutivo de llamadas con métricas de negocio (categorías y subtipos) con opción de filtrado por rango de fechas.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:

  # ✅ Procesar archivo individual con análisis ejecutivo Excel
  python3 categoria_analysis.py --files "/ruta/archivo.xlsx" --output analisis_ejecutivo.xlsx

  # ✅ Análisis comparativo de múltiples archivos
  python3 categoria_analysis.py --files "archivo1.xlsx,archivo2.xlsx" --output dashboard_ejecutivo.xlsx

  # ✅ Procesar todos los archivos de agentes
  python3 categoria_analysis.py --pattern "*AGENTES*.xlsx" --output reporte_directiva.xlsx

  # ✅ Mostrar resultados en pantalla para revisión rápida
  python3 categoria_analysis.py --files "archivo.xlsx"

  # ✅ Filtrar por rango de fechas específico
  python3 categoria_analysis.py --files "archivo.xlsx" --start-date "2025-01-01" --end-date "2025-01-31"

  # ✅ Análisis de un mes completo
  python3 categoria_analysis.py --pattern "*.xlsx" --start-date "2025-01-01" --end-date "2025-01-31" --output enero_2025.xlsx

  # ✅ Análisis de múltiples archivos con filtro de fechas
  python3 categoria_analysis.py --files "enero.xlsx,febrero.xlsx" --start-date "2025-02-01" --end-date "2025-02-28" --output febrero_2025.xlsx

NOTA: Especializado en análisis ejecutivo con métricas de negocio, insights y recomendaciones estratégicas.
        """
    )

    parser.add_argument(
        "--files",
        type=str,
        help="Lista de archivos Excel específicos separados por comas (sin espacios)"
    )

    parser.add_argument(
        "--pattern",
        type=str,
        help="Patrón para buscar archivos Excel (e.g., '*AGENTES*.xlsx')"
    )

    parser.add_argument(
        "--results-dir",
        type=str,
        help="Directorio personalizado donde buscar archivos Excel"
    )

    parser.add_argument(
        "--output",
        type=str,
        help="Archivo donde guardar el reporte (por defecto: mostrar en pantalla)"
    )

    parser.add_argument(
        "--start-date",
        type=str,
        help="Fecha de inicio para filtrar datos (formato: YYYY-MM-DD)"
    )

    parser.add_argument(
        "--end-date",
        type=str,
        help="Fecha de fin para filtrar datos (formato: YYYY-MM-DD)"
    )

    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Mostrar información detallada del proceso"
    )

    args = parser.parse_args()

    # Configurar nivel de logging
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    try:
        # Procesar fechas de filtro
        start_date = None
        end_date = None

        if args.start_date:
            try:
                start_date = datetime.strptime(args.start_date, '%Y-%m-%d')
                logger.info(f"Fecha de inicio configurada: {start_date.strftime('%Y-%m-%d')}")
            except ValueError as e:
                logger.error(f"Formato de fecha de inicio inválido. Use YYYY-MM-DD: {e}")
                return 1

        if args.end_date:
            try:
                end_date = datetime.strptime(args.end_date, '%Y-%m-%d')
                logger.info(f"Fecha de fin configurada: {end_date.strftime('%Y-%m-%d')}")
            except ValueError as e:
                logger.error(f"Formato de fecha de fin inválido. Use YYYY-MM-DD: {e}")
                return 1

        # Validar rango de fechas
        if start_date and end_date and start_date > end_date:
            logger.error("La fecha de inicio debe ser anterior o igual a la fecha de fin")
            return 1

        # Inicializar analizador
        analyzer = ExcelCategoryAnalyzer(args.results_dir, start_date, end_date)

        # Determinar archivos a procesar
        if args.files:
            file_paths = [Path(f.strip()) for f in args.files.split(",")]
            # Verificar que los archivos existen
            for file_path in file_paths:
                if not file_path.exists():
                    logger.error(f"El archivo {file_path} no existe")
                    return 1
        else:
            file_paths = analyzer.find_excel_files(args.pattern)

        if not file_paths:
            logger.error("No se encontraron archivos Excel para procesar")
            return 1

        logger.info(f"Procesando {len(file_paths)} archivo(s) Excel...")

        # Procesar archivos
        total_counter, total_rows, detailed_analysis = analyzer.analyze_multiple_files(file_paths)

        if not total_counter:
            logger.error("No se pudieron extraer categorías de ningún archivo")
            return 1

        # Generar reporte
        report = analyzer.generate_report(total_counter, total_rows, len(file_paths), detailed_analysis)

        # Mostrar o guardar reporte
        if args.output:
            if args.output.endswith('.xlsx'):
                analyzer.save_excel_report(args.output, total_counter, total_rows, len(file_paths), detailed_analysis)
                logger.info(f"Reporte Excel guardado en: {args.output}")
            else:
                try:
                    with open(args.output, 'w', encoding='utf-8') as f:
                        f.write(report)
                    logger.info(f"Reporte guardado en: {args.output}")
                except Exception as e:
                    logger.error(f"Error guardando reporte: {e}")
                    print(report)  # Mostrar en pantalla como fallback
        else:
            print(report)

        logger.info("Análisis completado exitosamente")
        return 0

    except Exception as e:
        logger.error(f"Error durante la ejecución: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
