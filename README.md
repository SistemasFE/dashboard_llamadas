# 📊 Sistema de Análisis de Categorías de Llamadas

Sistema completo de análisis y visualización de llamadas categorizadas, con generación de reportes ejecutivos en formato Excel y dashboards interactivos en Streamlit.

## 📋 Tabla de Contenidos

- [Descripción](#-descripción)
- [Características](#-características)
- [Requisitos](#-requisitos)
- [Instalación](#-instalación)
- [Uso](#-uso)
  - [Interfaz Web Streamlit](#interfaz-web-streamlit)
  - [Script de Línea de Comandos](#script-de-línea-de-comandos)
- [Estructura de Archivos](#-estructura-de-archivos)
- [Formato de Datos](#-formato-de-datos)
- [Reportes Generados](#-reportes-generados)
- [Ejemplos](#-ejemplos)
- [Solución de Problemas](#-solución-de-problemas)

---

## 🎯 Descripción

Este sistema permite analizar grandes volúmenes de datos de llamadas categorizadas, proporcionando:
- **Análisis estadístico** de categorías, subcategorías y subtipos
- **Visualizaciones interactivas** con gráficos dinámicos
- **Reportes ejecutivos** en formato Excel con múltiples hojas
- **Filtrado por fechas** para análisis de períodos específicos
- **Desglose por agente instalador** con métricas detalladas

## ✨ Características

### Análisis Completo
- ✅ Conteo preciso de llamadas (sin duplicaciones)
- ✅ Ranking de categorías generales con impacto operativo
- ✅ Análisis de subcategorías y subtipos
- ✅ Rutas completas de motivos (categoría general → específica → subtipo)
- ✅ Desglose por agente instalador (excluye "Sin asignar")

### Visualizaciones
- 📊 Gráficos de barras interactivos con Plotly
- 📈 Distribución de llamadas por volumen
- 🥧 Análisis de concentración de categorías
- 👷‍♂️ Gráficos de agentes instaladores
- 📉 Distribución de categorías por agente

### Reportes Excel
- 📄 Dashboard Ejecutivo con métricas clave
- 📋 KPIs principales
- 🗺️ Rutas completas de motivos
- 🔍 Análisis de subcategorías
- 👨‍💼 Hoja dedicada de agentes instaladores

### Filtros y Opciones
- 📅 Filtrado por rango de fechas
- 📁 Procesamiento de múltiples archivos
- 🔄 Análisis comparativo entre períodos
- 💾 Descarga de reportes con timestamp

---

## 🔧 Requisitos

### Software
- **Python**: 3.8 o superior
- **Sistema Operativo**: Linux, macOS, Windows (con WSL recomendado)

### Dependencias Python
```
streamlit >= 1.28.0
pandas >= 2.0.0
numpy >= 1.24.0
plotly >= 5.17.0
openpyxl >= 3.1.0
```

---

## 📦 Instalación

### Opción 1: Instalación Rápida

```bash
# Navegar al directorio
cd clasificador_llamadas/scripts/analisis

# Instalar dependencias
pip install -r requirements.txt
```

### Opción 2: Entorno Virtual (Recomendado)

```bash
# Crear entorno virtual
python3 -m venv venv-streamlit

# Activar entorno virtual
source venv-streamlit/bin/activate  # Linux/Mac
# o
venv-streamlit\Scripts\activate     # Windows

# Instalar dependencias
pip install -r requirements.txt
```

### Verificar Instalación

```bash
python3 -c "import streamlit, pandas, plotly, openpyxl; print('✅ Todas las dependencias instaladas correctamente')"
```

---

## 🚀 Uso

### Interfaz Web Streamlit

La forma más sencilla de usar el sistema es a través de la interfaz web interactiva.

#### Iniciar la Aplicación

```bash
streamlit run app.py
```

La aplicación se abrirá automáticamente en tu navegador en `http://localhost:8501`

#### Pasos de Uso

1. **📁 Subir Archivos**
   - Usa la barra lateral para subir uno o más archivos Excel (.xlsx o .xls)
   - Los archivos deben contener columnas de categorías (ver [Formato de Datos](#-formato-de-datos))

2. **📅 Configurar Filtros (Opcional)**
   - Selecciona un rango de fechas para filtrar los datos
   - Los filtros se aplican automáticamente al análisis

3. **🚀 Ejecutar Análisis**
   - Haz clic en "Ejecutar Análisis" en la barra lateral
   - El sistema procesará los archivos y generará todos los análisis

4. **📊 Explorar Resultados**
   - **Pestaña "Resumen Ejecutivo"**: Métricas principales y tablas detalladas
   - **Pestaña "Gráficos"**: Visualizaciones interactivas

5. **💾 Descargar Reporte**
   - Haz clic en "Descargar Reporte Excel" en la parte superior
   - El archivo incluye todas las métricas y análisis detallados

---

### Script de Línea de Comandos

Para análisis automatizados o procesamiento por lotes, puedes usar el script de línea de comandos.

#### Sintaxis Básica

```bash
python3 categoria_analysis.py [opciones]
```

#### Opciones Disponibles

| Opción | Descripción | Ejemplo |
|--------|-------------|---------|
| `--files` | Lista de archivos específicos (separados por comas) | `--files "archivo1.xlsx,archivo2.xlsx"` |
| `--pattern` | Patrón para buscar archivos | `--pattern "*AGENTES*.xlsx"` |
| `--results-dir` | Directorio donde buscar archivos | `--results-dir "/ruta/datos"` |
| `--output` | Archivo donde guardar el reporte Excel | `--output "reporte.xlsx"` |
| `--start-date` | Fecha de inicio (YYYY-MM-DD) | `--start-date "2025-01-01"` |
| `--end-date` | Fecha de fin (YYYY-MM-DD) | `--end-date "2025-01-31"` |
| `--verbose`, `-v` | Mostrar información detallada | `-v` |

#### Ejemplos de Uso

**1. Procesar un archivo y mostrar resultados en pantalla:**
```bash
python3 categoria_analysis.py --files "datos_llamadas.xlsx"
```

**2. Generar reporte Excel de múltiples archivos:**
```bash
python3 categoria_analysis.py \
  --files "enero.xlsx,febrero.xlsx,marzo.xlsx" \
  --output "reporte_q1_2025.xlsx"
```

**3. Analizar todos los archivos de un mes específico:**
```bash
python3 categoria_analysis.py \
  --pattern "*.xlsx" \
  --start-date "2025-01-01" \
  --end-date "2025-01-31" \
  --output "reporte_enero_2025.xlsx"
```

**4. Buscar archivos con patrón específico:**
```bash
python3 categoria_analysis.py \
  --pattern "*AGENTES*.xlsx" \
  --results-dir "/ruta/datos" \
  --output "reporte_agentes.xlsx" \
  --verbose
```

**5. Análisis comparativo entre períodos:**
```bash
# Período 1
python3 categoria_analysis.py \
  --files "datos.xlsx" \
  --start-date "2025-01-01" \
  --end-date "2025-01-31" \
  --output "enero_2025.xlsx"

# Período 2
python3 categoria_analysis.py \
  --files "datos.xlsx" \
  --start-date "2025-02-01" \
  --end-date "2025-02-28" \
  --output "febrero_2025.xlsx"
```

---

## 📂 Estructura de Archivos

```
analisis/
├── README.md                      # Este archivo
├── requirements.txt               # Dependencias Python
├── app.py                        # Aplicación web Streamlit
├── categoria_analysis.py         # Motor de análisis
├── __init__.py                   # Módulo Python
├── categoria_analysis.log        # Log de ejecuciones
├── results/                      # Directorio de resultados (auto-creado)
└── venv-streamlit/               # Entorno virtual (opcional)
```

### Descripción de Archivos Principales

#### `app.py`
Aplicación web interactiva con Streamlit. Proporciona interfaz gráfica para:
- Subir archivos Excel
- Configurar filtros
- Visualizar análisis en tiempo real
- Descargar reportes

#### `categoria_analysis.py`
Motor de análisis central. Contiene:
- Clase `ExcelCategoryAnalyzer`: Procesamiento y análisis de datos
- Identificación automática de columnas
- Generación de reportes Excel
- Cálculo de métricas y KPIs

#### `requirements.txt`
Lista de dependencias necesarias para ejecutar el sistema.

---

## 📄 Formato de Datos

### Columnas Requeridas

Los archivos Excel deben contener al menos **una columna de categoría**. El sistema identifica automáticamente las columnas usando estos patrones:

#### Columnas de Categoría (al menos una)
- `categoria_general` ⭐ (recomendado)
- `categoria`
- `category`
- `tipo`
- `type`

#### Columnas de Subcategoría (opcionales)
- `categoria_especifica` / `categoria_especifica_1` / `categoria_especifica_2` / `categoria_especifica_3`
- `subtipo_categoria` / `subtipo_categoria_1` / `subtipo_categoria_2` / `subtipo_categoria_3`
- `subcategoria`
- `motivo`

#### Columnas de Agente (opcionales)
- `agente_instalador` ⭐ (recomendado)
- `agente`
- `instalador`
- `tecnico`

#### Columnas de Fecha (opcionales, para filtrado)
- `fecha`
- `fecha_llamada`
- `date`
- Cualquier columna con tipo datetime

### Ejemplo de Estructura

```
| fecha       | categoria_general | categoria_especifica | subtipo_categoria | agente_instalador |
|-------------|-------------------|----------------------|-------------------|-------------------|
| 2025-01-15  | Información       | Contratación         | Estado Y Plazos   | Juan Pérez        |
| 2025-01-16  | Gestión           | Fallo Del Sistema    | Incidencia        | María García      |
| 2025-01-17  | Información       | Formación            | No Contemplado    | Juan Pérez        |
```

### Notas Importantes

1. **Valores con múltiples categorías**: Si una celda contiene múltiples valores separados por comas (ej: "Formación, Contratación"), el sistema los mantiene como un solo valor para evitar duplicar el conteo.

2. **Valores nulos**: Las filas con valores nulos en la columna de categoría principal se excluyen del análisis.

3. **Agentes sin asignar**: Los registros con agente instalador = "Sin asignar" se excluyen automáticamente del análisis de agentes.

---

## 📊 Reportes Generados

### Interfaz Streamlit

#### Pestaña "Resumen Ejecutivo"
- **Métricas Principales**
  - Total de archivos procesados
  - Total de llamadas analizadas

- **Ranking de Categorías Generales**
  - Tabla con ranking, frecuencia, porcentaje e impacto operativo

- **Análisis de Subcategorías**
  - Top subcategorías con frecuencia, porcentaje y prioridad

- **Rutas Completas de Motivos**
  - Desglose completo: Categoría General → Específica → Subtipo

- **Agentes Instaladores por Categoría**
  - Análisis detallado por agente con porcentajes

#### Pestaña "Gráficos"
- **Ranking de Categorías Generales**: Gráfico de barras Top 20
- **Distribución por Volumen**: Gráfico combinado
- **Top Subcategorías**: Gráfico de barras horizontal Top 15
- **Top Rutas Completas**: Gráfico de barras horizontal Top 15
- **Top Agentes Instaladores**: Gráfico de barras Top 15
- **Distribución por Agente**: Gráfico apilado Top 10 agentes

### Reporte Excel

El archivo Excel generado contiene las siguientes hojas:

#### Hoja 1: "Dashboard_Ejecutivo"

**Sección: RESUMEN EJECUTIVO**
- Período analizado
- Volumen total de llamadas
- Número de categorías únicas
- Top 3 categorías y su concentración
- Promedio de llamadas por categoría
- Insights de negocio

**Sección: KPIs PRINCIPALES**
- KPI 1: Categoría más frecuente
- KPI 2: Concentración top 3
- KPI 3: Diversidad de categorías
- Indicadores clave para la toma de decisiones

**Sección: RUTAS COMPLETAS DE MOTIVOS**
- Categoría General
- Categoría Específica
- Subtipo
- Ruta Completa
- Frecuencia
- % del Total

**Sección: ANÁLISIS DE SUBCATEGORÍAS**
- Tipo (categoría_especifica, subtipo_categoria, etc.)
- Subcategoría
- Frecuencia
- % Total
- Prioridad (Alta/Media/Baja)

#### Hoja 2: "Agentes_Instaladores"
- Agente Instalador
- Categoría General
- Categoría Específica
- Subtipo
- Ruta Completa
- Frecuencia
- % del Agente

---

## 💡 Ejemplos

### Ejemplo 1: Análisis Rápido de un Archivo

```bash
# Usando Streamlit (recomendado para exploración)
streamlit run app.py
# Luego subir el archivo desde la interfaz

# O usando línea de comandos
python3 categoria_analysis.py --files "llamadas_enero.xlsx"
```

### Ejemplo 2: Reporte Mensual Automatizado

```bash
#!/bin/bash
# Script para generar reporte mensual

FECHA_INICIO="2025-01-01"
FECHA_FIN="2025-01-31"
ARCHIVO_ENTRADA="datos_completos.xlsx"
ARCHIVO_SALIDA="reporte_enero_2025.xlsx"

python3 categoria_analysis.py \
  --files "$ARCHIVO_ENTRADA" \
  --start-date "$FECHA_INICIO" \
  --end-date "$FECHA_FIN" \
  --output "$ARCHIVO_SALIDA" \
  --verbose

echo "✅ Reporte generado: $ARCHIVO_SALIDA"
```

### Ejemplo 3: Análisis Comparativo

```python
# Script Python para análisis comparativo
import subprocess
from datetime import datetime

meses = [
    ("2025-01-01", "2025-01-31", "enero"),
    ("2025-02-01", "2025-02-28", "febrero"),
    ("2025-03-01", "2025-03-31", "marzo")
]

for inicio, fin, mes in meses:
    output = f"reporte_{mes}_2025.xlsx"
    cmd = [
        "python3", "categoria_analysis.py",
        "--files", "datos_completos.xlsx",
        "--start-date", inicio,
        "--end-date", fin,
        "--output", output
    ]
    subprocess.run(cmd, check=True)
    print(f"✅ Generado: {output}")
```

### Ejemplo 4: Procesar Múltiples Archivos por Agente

```bash
# Analizar archivos de diferentes agentes
python3 categoria_analysis.py \
  --files "agente1_enero.xlsx,agente2_enero.xlsx,agente3_enero.xlsx" \
  --output "consolidado_agentes_enero.xlsx"
```

---

## 🔍 Solución de Problemas

### Error: "No module named 'streamlit'"

**Causa**: Streamlit no está instalado.

**Solución**:
```bash
pip install streamlit
# o
pip install -r requirements.txt
```

### Error: "No se pudo identificar columna de categoría"

**Causa**: El archivo Excel no tiene una columna reconocible como categoría.

**Solución**:
- Asegúrate de que exista al menos una columna con nombres como: `categoria_general`, `categoria`, `tipo`, etc.
- O renombra tu columna de categorías a `categoria_general`

### Error: TypeError con 'width'

**Causa**: Versión incompatible de Streamlit.

**Solución**:
```bash
pip install --upgrade streamlit
```

### Los números no cuadran / Suma mayor al total

**Causa**: Ya está corregido en la versión actual. Si persiste, verifica que estás usando la última versión del código.

**Solución**: El sistema ahora cuenta cada fila del Excel como una llamada exactamente, sin expansiones ni duplicaciones.

### No aparecen los agentes instaladores

**Causa**: Los agentes tienen el valor "Sin asignar" o la columna no está presente.

**Solución**:
- Verifica que exista una columna `agente_instalador` en tus datos
- Los registros con "Sin asignar" se excluyen automáticamente del análisis de agentes

### La aplicación Streamlit no se abre

**Causa**: Puerto 8501 ya está en uso.

**Solución**:
```bash
# Usar otro puerto
streamlit run app.py --server.port 8502

# O detener el proceso en el puerto 8501
lsof -ti:8501 | xargs kill -9  # Linux/Mac
```

### Error al leer archivos Excel

**Causa**: Archivo corrupto o formato no compatible.

**Solución**:
- Verifica que el archivo sea .xlsx o .xls válido
- Abre el archivo en Excel y guárdalo nuevamente
- Verifica que no esté protegido con contraseña

---

## 📝 Logs

El sistema genera un archivo de log `categoria_analysis.log` que contiene:
- Timestamp de cada ejecución
- Archivos procesados
- Número de filas y columnas
- Columnas identificadas
- Errores y advertencias

### Ver logs recientes:

```bash
tail -n 50 categoria_analysis.log
```

---

## 🤝 Soporte

Para reportar problemas o solicitar nuevas funcionalidades:
1. Revisa primero la sección [Solución de Problemas](#-solución-de-problemas)
2. Verifica el archivo de log para más detalles
3. Contacta al equipo de desarrollo con:
   - Descripción del problema
   - Archivo de log relevante
   - Ejemplo de datos (si es posible)

---

## 📜 Licencia

Copyright © 2025 - Sistema de Análisis de Categorías de Llamadas

---

## 🔄 Historial de Cambios

### Versión Actual
- ✅ Eliminada expansión de valores con delimitadores
- ✅ Conteo exacto: 1 fila = 1 llamada
- ✅ Exclusión automática de "Sin asignar" en agentes
- ✅ Reportes Excel con hoja dedicada de agentes instaladores
- ✅ Visualizaciones mejoradas con Plotly
- ✅ Filtrado por rango de fechas
- ✅ Interfaz Streamlit optimizada

---

**Última actualización**: Octubre 2025
