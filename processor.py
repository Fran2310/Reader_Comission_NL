"""
Modulo de procesamiento de datos de comisiones (National Life Group - NLG).
Contiene la logica para:
1. Extraer datos y metadatos (Exported By, Exported On) desde archivos Excel (.xlsx).
2. Unificar uno o dos DataFrames de registros manteniendo las 10 columnas requeridas.
3. Clasificar registros entre el Agente Principal (NL) y el Equipo (Equipo NL).
4. Agrupar duplicados por Insured/Annuitant, sumar Gross Com. Paid y determinar el estado (Activa / Chargeback).
5. Generar el bloque de texto preconstruido listo para copiar.
"""

import re
import unicodedata
import io
import pandas as pd
import openpyxl

REQUIRED_COLUMNS = [
    'Agent #',
    'Agent Name',
    'Policy #',
    'Insured/Annuitant',
    'Process Date',
    'Premium',
    'Gross Com. Paid',
    'Com. Rate',
    'Part. %',
    'Payment Date'
]

def normalize_name(name: str) -> set:
    """
    Normaliza un nombre eliminando acentos, caracteres especiales,
    convirtiendo a minusculas y retornando un conjunto de palabras significativas.
    Permite comparar nombres con orden invertido (Apellidos, Nombres vs Nombres Apellidos).
    """
    if not name or pd.isna(name):
        return set()
    s = str(name).strip()
    s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8')
    tokens = re.findall(r'[a-zA-Z0-9]+', s.lower())
    return set(tokens)

def names_match(name1: str, name2: str) -> bool:
    """
    Compara dos nombres normalizados:
    - Retorna True si los conjuntos de palabras coinciden exactamente.
    - Retorna True si uno es subconjunto del otro (ej. iniciales o segundo apellido omitido)
      con al menos 2 palabras coincidentes.
    """
    s1 = normalize_name(name1)
    s2 = normalize_name(name2)
    if not s1 or not s2:
        return False
    if s1 == s2:
        return True
    # Subconjunto con al menos 2 palabras coincidentes
    if (s1.issubset(s2) or s2.issubset(s1)) and len(s1.intersection(s2)) >= 2:
        return True
    return False

def clean_currency(val) -> float:
    """
    Limpia y convierte un valor monetario o numerico a float.
    Maneja simbolos '$', comas, espacios y parentesis para numeros negativos:
    ($100.50) -> -100.50
    """
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return 0.0
    # Limpiar simbolos
    s = s.replace('$', '').replace(',', '').strip()
    # Manejar formato contable de negativos entre parentesis
    if s.startswith('(') and s.endswith(')'):
        try:
            return -float(s[1:-1].strip())
        except ValueError:
            return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0

def clean_percentage(val) -> float:
    """
    Limpia un porcentaje a valor decimal:
    0.15 -> 0.15
    '15.00%' -> 0.15
    '100%' -> 1.00
    """
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s.endswith('%'):
        try:
            return float(s[:-1].strip()) / 100.0
        except ValueError:
            return 0.0
    try:
        f = float(s)
        return f
    except ValueError:
        return 0.0

def extract_excel_data(file_source):
    """
    Lee un archivo Excel (.xlsx) ya sea como ruta en disco o BytesIO / UploadedFile.
    Extrae:
    - DataFrame con los registros de clientes limpios.
    - exported_by: Nombre del agente extraido de 'Exported By:'.
    - exported_on: Fecha de exportacion extraida de 'Exported On:'.
    """
    if isinstance(file_source, bytes):
        file_source = io.BytesIO(file_source)

    wb = openpyxl.load_workbook(file_source, data_only=True)
    ws = wb.active

    exported_by = None
    exported_on = None
    header_row_idx = None
    headers = []
    data_rows = []

    for r in range(1, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        
        # Escaneo de metadatos (Exported By / Exported On)
        for c_idx, val in enumerate(row_vals):
            if val is not None:
                sval = str(val).strip()
                if 'Exported By:' in sval:
                    m = re.search(r'Exported By:\s*(.*)', sval, re.IGNORECASE)
                    if m and m.group(1).strip():
                        exported_by = m.group(1).strip()
                    elif c_idx + 1 < len(row_vals) and row_vals[c_idx + 1]:
                        exported_by = str(row_vals[c_idx + 1]).strip()
                if 'Exported On:' in sval:
                    m = re.search(r'Exported On:\s*(.*)', sval, re.IGNORECASE)
                    if m and m.group(1).strip():
                        exported_on = m.group(1).strip()
                    elif c_idx + 1 < len(row_vals) and row_vals[c_idx + 1]:
                        exported_on = str(row_vals[c_idx + 1]).strip()

        # Deteccion de fila de encabezados
        if header_row_idx is None:
            normalized_vals = [str(v).strip() if v is not None else '' for v in row_vals]
            if 'Agent #' in normalized_vals or 'Policy #' in normalized_vals:
                header_row_idx = r
                headers = [str(v).strip() if v is not None else f'Col_{i}' for i, v in enumerate(row_vals)]
                continue

        # Lectura de filas de datos
        if header_row_idx is not None and r > header_row_idx:
            # Validar que sea fila de cliente
            v0 = row_vals[0]
            if v0 is not None:
                sv0 = str(v0).strip()
                # Omitir filas de pie de pagina o totales
                if sv0 and not sv0.startswith('Exported') and not sv0.startswith('Total') and sv0 != 'None':
                    # Tomar las columnas correspondientes a headers
                    row_slice = row_vals[:len(headers)]
                    # Comprobar que no sea fila completamente vacia
                    if any(c is not None and str(c).strip() != '' for c in row_slice):
                        data_rows.append(row_slice)

    if not headers:
        raise ValueError("No se encontro la fila de encabezados con las columnas requeridas ('Agent #', 'Policy #', etc.).")

    df = pd.DataFrame(data_rows, columns=headers)
    df.columns = [str(c).strip() for c in df.columns]

    # Validar columnas requeridas
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"El archivo no contiene todas las columnas requeridas. Faltan: {', '.join(missing)}")

    # Filtrar solo las 10 columnas requeridas en el orden oficial
    df = df[REQUIRED_COLUMNS].copy()

    return df, exported_by, exported_on

def unify_dataframes(data_list):
    """
    Combina uno o dos DataFrames en uno solo, unificando los metadatos.
    data_list: lista de tuplas (df, exported_by, exported_on)
    Retorna: (unified_df, final_exported_by, final_exported_on)
    """
    if not data_list:
        raise ValueError("No hay datos para unificar.")

    dfs = [item[0] for item in data_list]
    exported_by_list = [item[1] for item in data_list if item[1]]
    exported_on_list = [item[2] for item in data_list if item[2]]

    unified_df = pd.concat(dfs, ignore_index=True)

    # Limpiar posibles filas vacias residuales
    unified_df = unified_df.dropna(subset=['Insured/Annuitant', 'Gross Com. Paid'])
    unified_df = unified_df[unified_df['Insured/Annuitant'].astype(str).str.strip() != '']

    # Exported By: preferir el primero no nulo
    final_exported_by = exported_by_list[0] if exported_by_list else "Agente Principal"

    # Exported On: si son diferentes, unificarlos, si no, tomar el unico
    if exported_on_list:
        unique_dates = list(dict.fromkeys(exported_on_list))
        final_exported_on = " - ".join(unique_dates)
    else:
        final_exported_on = pd.Timestamp.now().strftime('%m/%d/%Y')

    return unified_df, final_exported_by, final_exported_on

def process_commissions(df: pd.DataFrame, principal_agent_name: str):
    """
    Aplica las reglas de negocio:
    1. Agrega columna Gross_Numeric limpia de Gross Com. Paid.
    2. Compara Agent Name contra principal_agent_name:
       - Principal (NL): Coincide con Exported By:
       - Equipo (Equipo NL): No coincide con Exported By:
    3. Agrupa por Insured/Annuitant en cada grupo.
    4. Suma Gross Com. Paid:
       - > 0 -> 'Activa'
       - <= 0 -> 'Chargeback'
    5. Ordena el equipo agrupado por Agente (orden alfabetico) y dentro del agente
       por el orden de aparicion del Insured/Annuitant (reproduciendo el formato exacto del reporte).

    Retorna diccionario:
    {
        'principal': [
            {'insured': str, 'status': str, 'total': float, 'agent': str, 'policies': list}, ...
        ],
        'team': [
            {'insured': str, 'status': str, 'total': float, 'agent': str, 'policies': list}, ...
        ],
        'raw_df': df_con_columnas_auxiliares
    }
    """
    work_df = df.copy()
    work_df['Gross_Numeric'] = work_df['Gross Com. Paid'].apply(clean_currency)
    work_df['Is_Principal'] = work_df['Agent Name'].apply(lambda x: names_match(x, principal_agent_name))

    # 1. Procesar Principal (NL)
    df_principal = work_df[work_df['Is_Principal']].copy()
    principal_results = []
    if not df_principal.empty:
        # Mantener orden de aparicion
        insured_seen = []
        for _, row in df_principal.iterrows():
            ins = str(row['Insured/Annuitant']).strip()
            if ins not in insured_seen:
                insured_seen.append(ins)
        
        for ins in insured_seen:
            sub = df_principal[df_principal['Insured/Annuitant'].astype(str).str.strip() == ins]
            total_sum = round(sub['Gross_Numeric'].sum(), 2)
            status = 'Activa' if total_sum > 0 else 'Chargeback'
            agent = sub['Agent Name'].iloc[0]
            policies = list(sub['Policy #'].unique())
            principal_results.append({
                'insured': ins,
                'status': status,
                'total': total_sum,
                'agent': agent,
                'policies': policies,
                'count': len(sub)
            })

    # 2. Procesar Equipo (Equipo NL)
    df_team = work_df[~work_df['Is_Principal']].copy()
    team_results = []
    if not df_team.empty:
        # Obtener agentes unicos del equipo ordenados alfabeticamente
        unique_team_agents = sorted(df_team['Agent Name'].dropna().unique())
        
        insured_team_seen = []
        for agent in unique_team_agents:
            agent_sub = df_team[df_team['Agent Name'] == agent]
            for _, row in agent_sub.iterrows():
                ins = str(row['Insured/Annuitant']).strip()
                if ins not in insured_team_seen:
                    insured_team_seen.append(ins)
                    # Sumar todas las apariciones de ese asegurado en todo el equipo
                    sub_ins = df_team[df_team['Insured/Annuitant'].astype(str).str.strip() == ins]
                    total_sum = round(sub_ins['Gross_Numeric'].sum(), 2)
                    status = 'Activa' if total_sum > 0 else 'Chargeback'
                    policies = list(sub_ins['Policy #'].unique())
                    team_results.append({
                        'insured': ins,
                        'status': status,
                        'total': total_sum,
                        'agent': agent,
                        'policies': policies,
                        'count': len(sub_ins)
                    })

    return {
        'principal': principal_results,
        'team': team_results,
        'raw_df': work_df
    }

def generate_prebuilt_text(classified_data: dict) -> str:
    """
    Genera el bloque de texto con el formato exacto requerido:

    Comision y Chargeback NL
    * [Insured/Annuitant]: Activa
    * [Insured/Annuitant]: Chargeback

    Comision y Chargeback de Equipo NL
    * [Insured/Annuitant]: Activa
    * [Insured/Annuitant]: Chargeback
    """
    lines = []
    
    lines.append("Comision y Chargeback NL")
    if classified_data['principal']:
        for item in classified_data['principal']:
            lines.append(f"* {item['insured']}: {item['status']}")
    else:
        lines.append("* (Sin registros para el agente principal)")

    lines.append("") # Linea en blanco de separacion

    lines.append("Comision y Chargeback de Equipo NL")
    if classified_data['team']:
        for item in classified_data['team']:
            lines.append(f"* {item['insured']}: {item['status']}")
    else:
        lines.append("* (Sin registros de equipo)")

    return "\n".join(lines)
