import streamlit as st
import pandas as pd
import plotly.express as px
import io

# --- Configuración de la Página ---
st.set_page_config(
    page_title="SICET INGENIERÍA - Análisis de Datos",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Funciones de Utilidad ---

def format_currency(value):
    """Formatea un número como moneda COP sin decimales."""
    try:
        # Formatea como entero, usa '.' como separador de miles
        return f"$ {int(value):,}".replace(",", ".")
    except (ValueError, TypeError):
        return "$ 0"

@st.cache_data
def convert_df_to_csv(df):
    """Convierte un DataFrame a CSV para descarga."""
    output = io.BytesIO()
    # Usar utf-8-sig para asegurar compatibilidad con Excel
    df.to_csv(output, index=False, encoding='utf-8-sig')
    return output.getvalue()


# --- Funciones de Carga y Procesamiento de Datos ---

# Mapeo de hojas (basado en la lógica JS)
REQUIRED_SHEETS_MAPPING = {
    'hoja 1': 'INFORMACION',
    'hoja 3': 'COMENTARIOS',
    'hoja 4': 'NOMINA'
}
MIN_MONTHLY_SHEET = 7  # Hoja 7 en adelante (índice 6 en Python)

def find_column(df, patterns):
    """Encuentra la primera columna que coincida con una lista de patrones."""
    for col in df.columns:
        col_upper = str(col).upper()
        for pattern in patterns:
            if pattern in col_upper:
                return col
    return None

@st.cache_data
def load_data(uploaded_file):
    """
    Carga y procesa el archivo Excel, replicando la lógica de parseo de JS.
    """
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        normalized_sheet_names = [name.strip().upper() for name in sheet_names]
        
        # --- 1. Encontrar y Parsear Hojas Requeridas ---
        sheet_name_map = {}
        for key, name in REQUIRED_SHEETS_MAPPING.items():
            try:
                sheet_name_map[key] = sheet_names[normalized_sheet_names.index(name)]
            except ValueError:
                st.error(f"No se encontró la hoja requerida: '{name}' ({key}).")
                return None
        
        # Leer hojas principales (dtype=str para proteger Cédulas)
        df_empleados = pd.read_excel(xls, sheet_name=sheet_name_map['hoja 1'], dtype=str)
        df_comentarios = pd.read_excel(xls, sheet_name=sheet_name_map['hoja 3'], dtype=str)
        df_nomina = pd.read_excel(xls, sheet_name=sheet_name_map['hoja 4'], dtype=str)

        # --- 2. Parsear Horas Extras (Hojas 7+) ---
        df_horas_extras_list = []
        he_sheet_names = []
        for i, sheet_name in enumerate(sheet_names):
            normalized_name = normalized_sheet_names[i]
            # Lógica JS: No es una hoja requerida, contiene '2025', está en/después de la hoja 7
            if (normalized_name not in REQUIRED_SHEETS_MAPPING.values() and 
                "2025" in normalized_name and  
                i >= MIN_MONTHLY_SHEET - 1):
                
                df_month = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
                df_month['MES'] = sheet_name.strip()
                df_horas_extras_list.append(df_month)
                he_sheet_names.append(sheet_name.strip())
        
        # --- 3. Procesar y Limpiar df_empleados (Hoja 1) ---
        id_col_emp = find_column(df_empleados, ['CÉDULA', 'ID', 'NÚMERO DE CONTACTO'])
        name_col_emp = find_column(df_empleados, ['NOMBRE', 'TÉCNICO', 'EMPLEADO'])
        phone_col_emp = find_column(df_empleados, ['TELÉFONO', 'CONTACTO'])
        
        if not id_col_emp or not name_col_emp:
            st.error("No se pudieron encontrar 'Cédula' o 'Nombre' en la hoja 'INFORMACION'.")
            return None
        
        # Guardar todas las columnas (para el modal) pero renombrar las clave
        rename_map_emp = {id_col_emp: 'CEDULA', name_col_emp: 'NOMBRE'}
        if phone_col_emp:
            rename_map_emp[phone_col_emp] = 'TELEFONO'
        df_empleados = df_empleados.rename(columns=rename_map_emp)
        df_empleados = df_empleados[df_empleados['CEDULA'].notna() & (df_empleados['CEDULA'] != '')]
        
        # --- 4. Procesar df_comentarios (Hoja 3) ---
        id_col_com = find_column(df_comentarios, ['CÉDULA', 'ID'])
        comment_col = find_column(df_comentarios, ['COMENTARIOS', 'OBSERVACIONES'])
        
        df_comentarios = df_comentarios.rename(columns={id_col_com: 'CEDULA', comment_col: 'COMENTARIOS'})
        df_comentarios = df_comentarios[df_comentarios['CEDULA'].notna() & (df_comentarios['CEDULA'] != '')]
        # Seleccionar solo las columnas de interés
        df_comentarios = df_comentarios[['CEDULA', 'COMENTARIOS']].set_index('CEDULA')
        
        # --- 5. Procesar df_nomina (Hoja 4) ---
        id_col_nom = find_column(df_nomina, ['CÉDULA', 'ID'])
        df_nomina = df_nomina.rename(columns={id_col_nom: 'CEDULA'})
        df_nomina = df_nomina[df_nomina['CEDULA'].notna() & (df_nomina['CEDULA'] != '')]
        
        # Renombrar y convertir columnas numéricas
        rename_map_nom = {
            find_column(df_nomina, ['SALARIO BASE']): 'SALARIO_BASE',
            find_column(df_nomina, ['CONTRIBUCIONES EMPLEADOR', 'CONTRIBUCIONES DEL EMPLEADOR']): 'CONTRIBUCION_EMPR',
            find_column(df_nomina, ['CONTRIBUCIONES EMPLEADO', 'CONTRIBUCIONES DEL EMPLEADO']): 'CONTRIBUCION_EMPL',
            find_column(df_nomina, ['APORTE ARL']): 'APORTE_ARL',
            find_column(df_nomina, ['SALARIO REAL']): 'SALARIO_REAL',
            find_column(df_nomina, ['SALARIO BRUTO']): 'SALARIO_BRUTO',
            find_column(df_nomina, ['HORAS EXTRA']): 'HORAS_EXTRA_NOM',
            find_column(df_nomina, ['TOTAL A PAGAR AL EMPLEADO']): 'TOTAL_PAGAR_NOM'
        }
        
        df_nomina = df_nomina.rename(columns=rename_map_nom)
        
        num_cols = ['SALARIO_BASE', 'CONTRIBUCION_EMPR', 'CONTRIBUCION_EMPL', 'APORTE_ARL', 
                    'SALARIO_REAL', 'SALARIO_BRUTO', 'HORAS_EXTRA_NOM', 'TOTAL_PAGAR_NOM']
        
        for col in num_cols:
            if col in df_nomina.columns:
                # Solo convertir a numérico si la columna fue encontrada y renombrada
                df_nomina[col] = pd.to_numeric(df_nomina[col], errors='coerce').fillna(0)
            
        df_nomina = df_nomina.set_index('CEDULA')
        
        # --- 6. Combinar Datos (Replicar lógica JS) ---
        # Añadir 'TOTAL_PAGAR_NOM' y 'HORAS_EXTRA_NOM' de Nómina a Comentarios
        df_comentarios = df_comentarios.join(df_nomina[['HORAS_EXTRA_NOM', 'TOTAL_PAGAR_NOM']], how='left')
        df_comentarios['HORAS_EXTRA_NOM'].fillna(0, inplace=True)
        df_comentarios['TOTAL_PAGAR_NOM'].fillna(0, inplace=True)

        # --- 7. Procesar Horas Extras (Pandas Melt) ---
        df_he_processed = pd.DataFrame()
        if df_horas_extras_list:
            processed_dfs = []
            for df_month in df_horas_extras_list:
                id_col_he = find_column(df_month, ['CÉDULA', 'ID'])
                if not id_col_he:
                    continue # Omitir hoja si no tiene Cédula

                # Encontrar columnas de clasificación (basado en lógica JS)
                classification_cols = []
                for i, col in enumerate(df_month.columns):
                    if i > 1 and ('HORA EXTRA' in str(col).upper() or 'RECARGO' in str(col).upper()):
                        classification_cols.append(col)
                
                if not classification_cols:
                    continue # Omitir si no hay columnas de horas
                    
                # "Derretir" (melt) la tabla a formato largo
                df_melted = df_month.melt(
                    id_vars=[id_col_he, 'MES'],
                    value_vars=classification_cols,
                    var_name='CLASIFICACION',
                    value_name='HORAS'
                )
                
                df_melted = df_melted.rename(columns={id_col_he: 'CEDULA'})
                df_melted['HORAS'] = pd.to_numeric(df_melted['HORAS'], errors='coerce').fillna(0)
                df_melted = df_melted[df_melted['HORAS'] > 0] # Mantener solo registros con horas
                processed_dfs.append(df_melted)

            if processed_dfs:
                df_he_processed = pd.concat(processed_dfs, ignore_index=True)
                df_he_processed = df_he_processed[df_he_processed['CEDULA'].notna() & (df_he_processed['CEDULA'] != '')]

        return {
            "empleados": df_empleados.reset_index(drop=True), # Resetear índice para filtros
            "comentarios": df_comentarios.reset_index(),
            "nomina": df_nomina.reset_index(),
            "horas_extras": df_he_processed,
            "he_sheet_names": he_sheet_names
        }

    except Exception as e:
        st.error(f"Error crítico al procesar el archivo: {e}")
        return None

# --- Funciones de UI por Sección ---

def show_empleados(df_empleados, df_nomina_lookup):
    st.header("👥 Empleados")
    
    # 1. Búsqueda
    search_query = st.text_input("Buscar por Nombre o Cédula...", key="emp_search")
    
    df_filtered = df_empleados
    if search_query:
        query = search_query.lower()
        df_filtered = df_empleados[
            df_empleados['NOMBRE'].str.lower().str.contains(query) |
            df_empleados['CEDULA'].str.contains(query)
        ]
        
    if df_filtered.empty:
        st.warning("No se encontraron empleados con ese criterio.")
        return

    # 2. "Tarjetas" de Empleados (Grid de 3 columnas)
    cols = st.columns(3)
    col_idx = 0
    
    for _, emp in df_filtered.iterrows():
        cedula = emp['CEDULA']
        nombre = emp['NOMBRE']
        
        # Buscar datos de nómina
        nomina_data = df_nomina_lookup.loc[cedula] if cedula in df_nomina_lookup.index else pd.Series()
        
        # --- Verificación de seguridad ---
        salario_real = nomina_data.get('SALARIO_REAL', 0)
        telefono = emp.get('TELEFONO', 'N/A')
        
        with cols[col_idx]:
            with st.container(border=True):
                st.subheader(f"{nombre}")
                st.text(f"Cédula: {cedula}")
                st.text(f"Teléfono: {telefono}")
                st.metric(label="Salario Real", value=format_currency(salario_real))
                
                # --- INICIO DE LA CORRECCIÓN ---
                # 3. Reemplazar st.dialog con st.popover
                # Esto elimina el botón y el modal, y lo reemplaza con un popover más simple
                with st.popover("Ver Detalle Completo", use_container_width=True):
                    st.subheader(nombre)
                    # Mostrar *todos* los datos de la Hoja 1 para este empleado
                    detalle_df = emp.to_frame(name="Valor")
                    detalle_df.index.name = "Campo"
                    st.dataframe(detalle_df, use_container_width=True)
                # --- FIN DE LA CORRECCIÓN ---
        
        col_idx = (col_idx + 1) % 3

def show_comentarios(df_comentarios, df_empleados_master):
    st.header("💬 Comentarios y Observaciones")
    
    # Unir con nombres de empleados
    df_comentarios = df_comentarios.join(df_empleados_master[['NOMBRE', 'TELEFONO']], on='CEDULA', how='left')
    
    # 1. Métrica Total
    st.metric("Total General de Comentarios Registrados", len(df_comentarios))
    
    # 2. Filtro Dropdown
    employee_list = df_comentarios.sort_values('NOMBRE')['NOMBRE'].dropna().unique()
    options = ["Todos los empleados"] + list(employee_list)
    selected_employee = st.selectbox(
        "Seleccione un empleado para filtrar...",
        options=options
    )
    
    # 3. Filtrar datos
    if selected_employee == "Todos los empleados":
        df_filtered = df_comentarios
    else:
        df_filtered = df_comentarios[df_comentarios['NOMBRE'] == selected_employee]
        
    if df_filtered.empty:
        st.warning("No hay comentarios para el criterio seleccionado.")
        return

    # 4. "Tarjetas" de Comentarios
    for _, item in df_filtered.iterrows():
        with st.container(border=True, key=item['CEDULA']):
            st.subheader(f"{item.get('NOMBRE', 'N/A')} ({item['CEDULA']})")
            
            cols = st.columns(2)
            cols[0].text(f"Teléfono: {item.get('TELEFONO', 'N/A')}")
            # --- Verificación de seguridad ---
            cols[1].text(f"Total a Pagar: {format_currency(item.get('TOTAL_PAGAR_NOM', 0))}")
            
            st.divider()
            # Usar un área de texto deshabilitada para mostrar el comentario
            st.text_area("Comentarios:", value=item.get('COMENTARIOS', 'N/A'), disabled=True, height=100)

def show_horas_extras(df_he, df_empleados_master, he_sheet_names):
    st.header("⏳ Análisis de Horas Extras")

    if df_he.empty:
        st.warning("No se encontraron datos de Horas Extras para analizar.")
        return

    # 1. Unir nombres de empleados
    df_he = df_he.join(df_empleados_master['NOMBRE'], on='CEDULA', how='left')
    df_he['NOMBRE'] = df_he['NOMBRE'].fillna(df_he['CEDULA']) # Usar Cédula si no hay nombre
    
    # 2. Tarjetas de Resumen
    total_general_hours = df_he['HORAS'].sum()
    total_clasificaciones = df_he['CLASIFICACION'].nunique()
    total_meses = df_he['MES'].nunique()
    
    cols_metrics = st.columns(3)
    cols_metrics[0].metric("Total General de Horas Extras", f"{total_general_hours:.2f} hrs")
    cols_metrics[1].metric("Total Clasificaciones Únicas", total_clasificaciones)
    cols_metrics[2].metric("Total Meses Procesados", total_meses)
    
    # 3. Top 3 Empleados
    with st.expander("🏆 Top 3 Empleados con Más Horas Extras", expanded=True):
        df_top_emp = df_he.groupby('NOMBRE')['HORAS'].sum().nlargest(3).reset_index()
        df_top_emp.index = ['🥇', '🥈', '🥉'][:len(df_top_emp)] # Añadir emojis
        st.dataframe(
            df_top_emp.style.format({'HORAS': '{:.2f} hrs'}),
            use_container_width=True
        )

    st.divider()
    
    # 4. Gráficos
    st.subheader("Gráficos de Horas Extras")
    
    try:
        # Gráfico 1: Por Clasificación (Doughnut)
        df_chart1 = df_he.groupby('CLASIFICACION')['HORAS'].sum().reset_index()
        fig1 = px.pie(df_chart1, names='CLASIFICACION', values='HORAS', 
                      title="Horas Extras por Clasificación (Total)")
        
        # Gráfico 2: Por Empleado (Bar)
        df_chart2 = df_he.groupby('NOMBRE')['HORAS'].sum().reset_index()
        fig2 = px.bar(df_chart2, x='NOMBRE', y='HORAS', 
                      title="Horas Extras Totales por Empleado (Comparación)")
        
        # Gráfico 3: Por Mes (Line)
        df_chart3 = df_he.groupby(['MES', 'NOMBRE'])['HORAS'].sum().reset_index()
        fig3 = px.line(df_chart3, x='MES', y='HORAS', color='NOMBRE', 
                       title="Horas Extras por Mes y Empleado (Comparación Temporal)")
        
        # Mostrar gráficos
        cols_charts = st.columns(2)
        cols_charts[0].plotly_chart(fig1, use_container_width=True)
        cols_charts[1].plotly_chart(fig2, use_container_width=True)
        st.plotly_chart(fig3, use_container_width=True)
    
    except Exception as e:
        st.error(f"Error al generar gráficos de Horas Extras: {e}")

    st.divider()

    # 5. Tabla Detallada con Filtro
    st.subheader("📊 Detalle de Horas Extras por Empleado y Clasificación")
    
    month_options = ["Total General"] + he_sheet_names
    selected_month = st.selectbox("Filtrar por Mes:", options=month_options)
    
    df_table_data = df_he
    if selected_month != "Total General":
        df_table_data = df_he[df_he['MES'] == selected_month]
        
    if not df_table_data.empty:
        # Usar pivot_table es la forma "Pythonica" de crear esta tabla
        pivot = pd.pivot_table(
            df_table_data,
            values='HORAS',
            index='NOMBRE',
            columns='CLASIFICACION',
            aggfunc='sum',
            fill_value=0,
            margins=True,       # ¡Esto añade la fila y columna de Total automáticamente!
            margins_name="TOTAL GENERAL"
        )
        
        # Aplicar estilo para destacar los totales
        st.dataframe(
            pivot.style.format('{:.2f}')
                   .set_properties(**{'font-weight': 'bold'}, subset=pd.IndexSlice['TOTAL GENERAL', :])
                   .set_properties(**{'font-weight': 'bold'}, subset=pd.IndexSlice[:, 'TOTAL GENERAL']),
            use_container_width=True
        )
    else:
        st.info("No hay datos de horas extras para la selección actual.")

def show_nomina(df_nomina, df_empleados_master):
    st.header("💰 Análisis de Nómina")
    
    if df_nomina.empty:
        st.warning("No hay datos de Nómina para analizar.")
        return

    # 1. Tarjetas de Resumen
    # Verificar si las columnas existen antes de sumar. Si no, usar 0.
    total_real = df_nomina['SALARIO_REAL'].sum() if 'SALARIO_REAL' in df_nomina.columns else 0
    total_bruto = df_nomina['SALARIO_BRUTO'].sum() if 'SALARIO_BRUTO' in df_nomina.columns else 0
    total_empr = df_nomina['CONTRIBUCION_EMPR'].sum() if 'CONTRIBUCION_EMPR' in df_nomina.columns else 0
    total_empl = df_nomina['CONTRIBUCION_EMPL'].sum() if 'CONTRIBUCION_EMPL' in df_nomina.columns else 0
    
    cols_metrics = st.columns(4)
    cols_metrics[0].metric("Total Salario Real (Acumulado)", format_currency(total_real))
    cols_metrics[1].metric("Total Salario Bruto (Acumulado)", format_currency(total_bruto))
    cols_metrics[2].metric("Total Contribuciones Empleador", format_currency(total_empr))
    cols_metrics[3].metric("Total Contribuciones Empleado", format_currency(total_empl))

    st.divider()

    # 2. Tabla Detallada con Búsqueda
    st.subheader("Lista Detallada de Nómina")
    
    # Unir nombres de empleados
    df_nomina = df_nomina.join(df_empleados_master['NOMBRE'], on='CEDULA', how='left')
    df_nomina['NOMBRE'] = df_nomina['NOMBRE'].fillna(df_nomina['CEDULA'])
    
    search_query = st.text_input("Buscar empleado por Nombre o Cédula...", key="nom_search")
    
    df_filtered = df_nomina
    if search_query:
        query = search_query.lower()
        df_filtered = df_nomina[
            df_nomina['NOMBRE'].str.lower().str.contains(query) |
            df_nomina['CEDULA'].str.contains(query)
        ]
    
    # Preparar DataFrame para mostrar
    display_cols = [
        'CEDULA', 'NOMBRE', 'SALARIO_BASE', 'CONTRIBUCION_EMPR',
        'CONTRIBUCION_EMPL', 'APORTE_ARL', 'SALARIO_BRUTO', 'SALARIO_REAL'
    ]
    # Asegurarse de que las columnas existan
    df_display_safe = df_filtered.copy()
    for col in display_cols:
        if col not in df_display_safe.columns:
            df_display_safe[col] = 0
            
    df_display = df_display_safe[display_cols]
    
    # Aplicar formato de moneda
    format_dict = {
        'SALARIO_BASE': format_currency,
        'CONTRIBUCION_EMPR': format_currency,
        'CONTRIBUCION_EMPL': format_currency,
        'APORTE_ARL': format_currency,
        'SALARIO_BRUTO': format_currency,
        'SALARIO_REAL': format_currency,
    }
    st.dataframe(df_display.style.format(format_dict), use_container_width=True)

    # 3. Botón de Exportar
    csv_data = convert_df_to_csv(df_display)
    st.download_button(
        label="⬇️ Exportar Resumen de Nómina (CSV)",
        data=csv_data,
        file_name="Resumen_Nomina_SICET.csv",
        mime="text/csv",
    )
    
    st.divider()

    # 4. Gráficos
    st.subheader("Gráficos de Nómina")
    
    try:
        # Gráfico 1: Distribución Salario Real (Pie)
        if 'SALARIO_REAL' in df_nomina.columns and df_nomina['SALARIO_REAL'].sum() > 0:
            fig1 = px.pie(
                df_nomina[df_nomina['SALARIO_REAL'] > 0], 
                names='NOMBRE', 
                values='SALARIO_REAL',
                title="Distribución de Salario Real por Empleado"
            )
            
        else:
            fig1 = px.pie(title="Distribución de Salario Real por Empleado (Sin datos)")

        
        # Gráfico 2: Comparativa de Conceptos (Bar)
        # (Las variables total_empr, total_empl ya se calcularon de forma segura arriba)
        total_arl = df_nomina['APORTE_ARL'].sum() if 'APORTE_ARL' in df_nomina.columns else 0
        total_base = df_nomina['SALARIO_BASE'].sum() if 'SALARIO_BASE' in df_nomina.columns else 0

        df_chart2 = pd.DataFrame({
            'Concepto': ['Salario Base', 'Contrib. Empleador', 'Contrib. Empleado', 'Aporte ARL'],
            'Monto': [total_base, total_empr, total_empl, total_arl]
        })
        fig2 = px.bar(
            df_chart2, x='Concepto', y='Monto',
            title="Comparativa de Conceptos de Nómina (Totales)",
            color='Concepto'
        )
        
        cols_charts = st.columns(2)
        cols_charts[0].plotly_chart(fig1, use_container_width=True)
        cols_charts[1].plotly_chart(fig2, use_container_width=True)
    
    except Exception as e:
        st.error(f"Error al generar gráficos de Nómina: {e}")

# --- Lógica Principal de la Aplicación ---

def main():
    st.title("SICET INGENIERÍA - Análisis de Datos")
    
    # Carga la imagen desde la carpeta local 'assets'
    try:
        st.image("assets/logo-sicet-azul.png", width=250)
    except Exception as e:
        # Si no encuentra la imagen local, muestra una advertencia
        st.warning(f"No se pudo cargar el logo: {e}. (Asegúrate de tener la carpeta 'assets' con el logo)")

    # 1. Carga de Archivo
    uploaded_file = st.file_uploader("Cargar Archivo Excel", type=["xlsx", "xls"])

    if uploaded_file is None:
        st.info("Por favor, cargue un archivo Excel para comenzar el análisis.")
        st.stop()
        
    # 2. Cargar y cachear datos
    data = load_data(uploaded_file)
    
    if data is None:
        st.error("El archivo no pudo ser procesado. Verifique el formato y las hojas requeridas.")
        st.stop()

    # 3. Navegación en Sidebar
    st.sidebar.title("Navegación")
    section = st.sidebar.radio(
        "Seleccione una sección:",
        ("👥 Empleados", "💬 Comentarios", "⏳ Horas Extras", "💰 Nómina"),
        captions=["Info y búsqueda", "Observaciones", "Análisis HE", "Análisis de Pago"]
    )
    
    # --- Pre-procesar datos para UI (crear lookups) ---
    df_empleados_master = data['empleados'].set_index('CEDULA')
    df_nomina_lookup = data['nomina'].set_index('CEDULA')
    
    # 4. Enrutamiento de Secciones
    if section == "👥 Empleados":
        show_empleados(data['empleados'], df_nomina_lookup)
        
    elif section == "💬 Comentarios":
        show_comentarios(data['comentarios'], df_empleados_master)
        
    elif section == "⏳ Horas Extras":
        show_horas_extras(data['horas_extras'], df_empleados_master, data['he_sheet_names'])
        
    elif section == "💰 Nómina":
        show_nomina(data['nomina'], df_empleados_master)

if __name__ == "__main__":
    main()