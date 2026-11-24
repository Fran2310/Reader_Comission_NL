"""
Modulo de generacion de reportes PDF unificados (National Life Group - NLG).
Replica el diseño, fuentes, paleta de colores y estructura exacta del PDF de referencia:
- Pagina de alta definicion (1748.57 x 2262.86 pt) basada en la exportacion nativa de Excel.
- Encabezado verde suave (#C6E0B4) con texto en negrita centrado.
- Filas alternadas en azul cielo (#83CCEB) y azul hielo (#C0E6F5).
- Cuadricula completa con lineas de 0.68 pt.
- Alineaciones contables: montos y porcentajes a la derecha, codigos y nombres a la izquierda.
- Formato contable para valores negativos: ($1,234.56).
- Pie de pagina con 'Exported On:' y 'Exported By:'.
- Soporte para fuentes Calibri / Calibri-Bold registradas dinamicamente con fallback a Helvetica.
"""

import os
import re
import io
import pandas as pd
from reportlab.lib.colors import HexColor
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Registro de fuentes de alta fidelidad
FONT_REGULAR = 'Helvetica'
FONT_BOLD = 'Helvetica-Bold'

for calibri_dir in ['C:/Windows/Fonts', '/usr/share/fonts/truetype', '/Library/Fonts']:
    reg_path = os.path.join(calibri_dir, 'calibri.ttf')
    bold_path = os.path.join(calibri_dir, 'calibrib.ttf')
    if os.path.exists(reg_path) and os.path.exists(bold_path):
        try:
            pdfmetrics.registerFont(TTFont('Calibri', reg_path))
            pdfmetrics.registerFont(TTFont('Calibri-Bold', bold_path))
            FONT_REGULAR = 'Calibri'
            FONT_BOLD = 'Calibri-Bold'
            break
        except Exception:
            pass

def clean_filename(agent_name: str) -> str:
    """
    Genera el nombre de archivo estandar:
    NLG_[nombre del agente]_PayableGrossComission.pdf
    Limpia caracteres invalidos en nombres de archivo (Windows/Linux/Mac).
    """
    safe_name = str(agent_name or 'Agente').strip()
    # Remover caracteres prohibidos en sistemas de archivos
    safe_name = re.sub(r'[<>:"/\\|?*]', '', safe_name)
    # Reemplazar multiples espacios por uno solo
    safe_name = re.sub(r'\s+', ' ', safe_name).strip()
    return f"NLG_{safe_name}_PayableGrossComission.pdf"

def format_currency_str(val) -> str:
    """Formatea numero a moneda con negativos contables: ($100.00)."""
    if val is None or pd.isna(val) or val == '':
        return '$0.00'
    try:
        if isinstance(val, str):
            s = val.replace('$', '').replace(',', '').strip()
            if s.startswith('(') and s.endswith(')'):
                f = -float(s[1:-1])
            else:
                f = float(s)
        else:
            f = float(val)
        if f < 0:
            return f"(${abs(f):,.2f})"
        return f"${f:,.2f}"
    except Exception:
        return str(val)

def format_percent_str(val) -> str:
    """Formatea numero a porcentaje: 0.15 -> 15.00%, 1 -> 100.00%."""
    if val is None or pd.isna(val) or val == '':
        return '0.00%'
    try:
        if isinstance(val, str):
            s = val.strip()
            if s.endswith('%'):
                return s
            f = float(s)
        else:
            f = float(val)
        if f <= 1.0 and f > 0:
            return f"{f * 100:.2f}%"
        return f"{f:.2f}%"
    except Exception:
        return str(val)

def format_date_str(val) -> str:
    """Formatea fecha a formato MM/DD/YYYY."""
    if val is None or pd.isna(val) or val == '':
        return ''
    if hasattr(val, 'strftime'):
        return val.strftime('%m/%d/%Y')
    return str(val).strip()

def build_unified_pdf(df: pd.DataFrame, exported_by: str, exported_on: str, output_target=None):
    """
    Genera el archivo PDF unificado basado en el dataframe y metadatos.
    output_target: Ruta str de archivo en disco o BytesIO.
    Si es None, crea y retorna un BytesIO con el PDF binario.
    """
    if output_target is None:
        target_buffer = io.BytesIO()
    elif isinstance(output_target, str):
        target_buffer = output_target
    else:
        target_buffer = output_target

    # Dimensiones exactas del PDF de referencia (exportacion de Excel a 200 DPI)
    PAGE_WIDTH = 1748.57
    PAGE_HEIGHT = 2262.86
    LEFT_MARGIN = 143.36
    RIGHT_MARGIN = PAGE_WIDTH - 1516.01  # ~232.56
    TOP_MARGIN = 169.8
    BOTTOM_MARGIN = 100.0

    doc = SimpleDocTemplate(
        target_buffer,
        pagesize=(PAGE_WIDTH, PAGE_HEIGHT),
        leftMargin=LEFT_MARGIN,
        rightMargin=RIGHT_MARGIN,
        topMargin=TOP_MARGIN,
        bottomMargin=BOTTOM_MARGIN
    )

    story = []

    # 1. Titulo del documento
    title_style = ParagraphStyle(
        'DocTitle',
        fontName=FONT_BOLD,
        fontSize=17.8,
        leading=22,
        textColor=HexColor('#000000'),
        spaceAfter=38
    )
    story.append(Paragraph('Payable Gross Commission', title_style))

    # 2. Estilos para celdas
    h_style = ParagraphStyle(
        'HStyle',
        fontName=FONT_BOLD,
        fontSize=11.0,
        leading=13,
        alignment=1, # Centrado
        textColor=HexColor('#000000')
    )
    cell_left = ParagraphStyle(
        'CLeft',
        fontName=FONT_REGULAR,
        fontSize=10.5,
        leading=12.5,
        alignment=0, # Izquierda
        textColor=HexColor('#000000')
    )
    cell_right = ParagraphStyle(
        'CRight',
        fontName=FONT_REGULAR,
        fontSize=10.5,
        leading=12.5,
        alignment=2, # Derecha
        textColor=HexColor('#000000')
    )

    # 3. Construccion de datos de la tabla
    headers = [
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

    table_data = []
    # Encabezados
    table_data.append([Paragraph(h, h_style) for h in headers])

    # Filas de datos
    for _, row in df.iterrows():
        table_data.append([
            Paragraph(str(row.get('Agent #') or '').strip(), cell_left),
            Paragraph(str(row.get('Agent Name') or '').strip(), cell_left),
            Paragraph(str(row.get('Policy #') or '').strip(), cell_left),
            Paragraph(str(row.get('Insured/Annuitant') or '').strip(), cell_left),
            Paragraph(format_date_str(row.get('Process Date')), cell_left),
            Paragraph(format_currency_str(row.get('Premium')), cell_right),
            Paragraph(format_currency_str(row.get('Gross Com. Paid')), cell_right),
            Paragraph(format_percent_str(row.get('Com. Rate')), cell_right),
            Paragraph(format_percent_str(row.get('Part. %')), cell_right),
            Paragraph(format_date_str(row.get('Payment Date')), cell_left),
        ])

    # 10 columnas con ancho exacto uniforme (137.26 pt cada una)
    col_widths = [137.265] * 10

    # Estilos de la tabla
    t_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), HexColor('#C6E0B4')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.68, HexColor('#000000')),
    ])

    # Alternar colores de filas
    for i in range(1, len(table_data)):
        bg = HexColor('#83CCEB') if i % 2 == 1 else HexColor('#C0E6F5')
        t_style.add('BACKGROUND', (0, i), (-1, i), bg)

    table = Table(table_data, colWidths=col_widths, style=t_style, repeatRows=1)
    story.append(table)

    # 4. Pie de pagina
    story.append(Spacer(1, 28))
    footer_style = ParagraphStyle(
        'FooterStyle',
        fontName=FONT_BOLD,
        fontSize=12.34,
        leading=16,
        textColor=HexColor('#000000'),
        spaceAfter=4
    )
    story.append(Paragraph(f'Exported On: {exported_on}', footer_style))
    story.append(Paragraph(f'Exported By: {exported_by}', footer_style))

    # Construir el documento
    doc.build(story)

    if output_target is None:
        target_buffer.seek(0)
        return target_buffer
    return output_target
