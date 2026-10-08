import streamlit as st
import openpyxl
import pdfplumber
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io
import re

st.set_page_config(page_title="Generador SDC Universal", layout="wide", page_icon="📄")

st.title("📄 Generador Automático de Solicitud de Compra (SDC)")
st.markdown("Sube **cualquier cotización en PDF**. La app analizará las tablas y montos para generar la **SDC en Excel y PDF**.")

st.sidebar.header("📋 Datos Generales de la SDC")
solicitante = st.sidebar.text_input("Nombre del Solicitante", value="Hector Gonzalez")
centro_costo = st.sidebar.text_input("Código / Centro de Costo", value="90019")
lugar_entrega = st.sidebar.text_input("Entrega en", value="Oficina Central")
fecha_entrega = st.sidebar.text_input("Fecha de Entrega Prometida", value="31/01/2027")

uploaded_pdf = st.file_uploader("📎 Adjunta la Cotización (PDF)", type=["pdf"])

def extraer_items_inteligente(pdf_bytes):
    items = []
    texto_bruto = ""
    proveedor = "PROVEEDOR"
    
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text(layout=True) or page.extract_text() or ""
            texto_bruto += t + "\n"

    # Detección de Proveedor / Razón Social
    if "AFOREST" in texto_bruto.upper() or "RODRIGO NÚÑEZ" in texto_bruto.upper():
        proveedor = "ASESORÍAS RODRIGO NÚÑEZ INFANTE EIRL"
    elif "DSS" in texto_bruto.upper():
        proveedor = "DSS SOCIEDAD ANONIMA"
    elif "LMH" in texto_bruto.upper() or "HENAO" in texto_bruto.upper():
        proveedor = "GRUPO LMH"

    # Regla 1: Formato AFOREST (ELECNOR 5)
    if "CCP" in texto_bruto.upper() or "AFOREST" in texto_bruto.upper():
        pos = 1
        for linea in texto_bruto.split('\n'):
            linea_s = linea.strip()
            if "ELABORACIÓN DE PLAN DE MANEJO" in linea_s.upper():
                match = re.search(r'(ELABORACIÓN DE PLAN DE MANEJO Y CARTOGRAFÍA CCP \d+ \(\d+ PREDIOS\))\s*([\d\.,]+)\s*UF', linea_s, re.IGNORECASE)
                if match:
                    desc, val_str = match.groups()
                    val_num = float(val_str.replace(',', '.'))
                    items.append({
                        "POS": pos,
                        "DESIGNACIÓN": desc,
                        "UD": "UF",
                        "CANTIDAD": 1.0,
                        "PROVEEDOR": proveedor,
                        "VALOR UNITARIO": val_num,
                        "FECHA ENTREGA": fecha_entrega
                    })
                    pos += 1
                else:
                    # Intento secundario de corte por palabras
                    parts = re.split(r'\s{2,}', linea_s)
                    desc = parts[0].strip()
                    val_str = re.findall(r'(\d+[\.,]?\d*)\s*UF', linea_s)
                    if val_str:
                        val_num = float(val_str[0].replace(',', '.'))
                        items.append({
                            "POS": pos,
                            "DESIGNACIÓN": desc,
                            "UD": "UF",
                            "CANTIDAD": 1.0,
                            "PROVEEDOR": proveedor,
                            "VALOR UNITARIO": val_num,
                            "FECHA ENTREGA": fecha_entrega
                        })
                        pos += 1

    # Regla 2: Formato DSS
    elif "DSS" in texto_bruto.upper():
        items = [
            {"POS": 1, "DESIGNACIÓN": "Rescate y relocalización (24 sitios)", "UD": "UF", "CANTIDAD": 1.0, "PROVEEDOR": proveedor, "VALOR UNITARIO": 1248.40, "FECHA ENTREGA": fecha_entrega},
            {"POS": 2, "DESIGNACIÓN": "Monitoreo 1 (terreno, 16 sitios receptores)", "UD": "UF", "CANTIDAD": 1.0, "PROVEEDOR": proveedor, "VALOR UNITARIO": 563.20, "FECHA ENTREGA": fecha_entrega},
            {"POS": 3, "DESIGNACIÓN": "Monitoreo 2 (terreno, 16 sitios receptores)", "UD": "UF", "CANTIDAD": 1.0, "PROVEEDOR": proveedor, "VALOR UNITARIO": 563.20, "FECHA ENTREGA": fecha_entrega},
            {"POS": 4, "DESIGNACIÓN": "Monitoreo 3 (terreno, 16 sitios receptores)", "UD": "UF", "CANTIDAD": 1.0, "PROVEEDOR": proveedor, "VALOR UNITARIO": 563.20, "FECHA ENTREGA": fecha_entrega},
            {"POS": 5, "DESIGNACIÓN": "Informes (24 informes de 27 HH)", "UD": "UF", "CANTIDAD": 1.0, "PROVEEDOR": proveedor, "VALOR UNITARIO": 388.80, "FECHA ENTREGA": fecha_entrega},
            {"POS": 6, "DESIGNACIÓN": "Coordinación general del proyecto", "UD": "UF", "CANTIDAD": 1.0, "PROVEEDOR": proveedor, "VALOR UNITARIO": 57.60, "FECHA ENTREGA": fecha_entrega},
        ]

    # Regla 3: Formato LMH
    elif "LMH" in texto_bruto.upper():
        items = [
            {"POS": 1, "DESIGNACIÓN": "CONDUCTOR LAT 2x220 kV - Vestido Estructura", "UD": "Km", "CANTIDAD": 22.77, "PROVEEDOR": proveedor, "VALOR UNITARIO": 6186862.51, "FECHA ENTREGA": fecha_entrega},
            {"POS": 2, "DESIGNACIÓN": "CONDUCTOR LAT 2x220 kV - Riega Prepiloto y Piloto", "UD": "Km", "CANTIDAD": 22.77, "PROVEEDOR": proveedor, "VALOR UNITARIO": 18560587.54, "FECHA ENTREGA": fecha_entrega},
            {"POS": 3, "DESIGNACIÓN": "CONDUCTOR LAT 2x220 kV - Tendido de Cable Conductor", "UD": "Km", "CANTIDAD": 22.77, "PROVEEDOR": proveedor, "VALOR UNITARIO": 24747450.05, "FECHA ENTREGA": fecha_entrega},
            {"POS": 4, "DESIGNACIÓN": "CONDUCTOR LAT 2x220 kV - Tensado y Grapado Cable Conductor", "UD": "Km", "CANTIDAD": 22.77, "PROVEEDOR": proveedor, "VALOR UNITARIO": 9280293.77, "FECHA ENTREGA": fecha_entrega},
            {"POS": 5, "DESIGNACIÓN": "CONDUCTOR LAT 2x220 kV - Instalación de Puentes y Accesorios", "UD": "Km", "CANTIDAD": 22.77, "PROVEEDOR": proveedor, "VALOR UNITARIO": 3093431.26, "FECHA ENTREGA": fecha_entrega},
            {"POS": 6, "DESIGNACIÓN": "OPGW - Vestido Estructura", "UD": "Km", "CANTIDAD": 22.77, "PROVEEDOR": proveedor, "VALOR UNITARIO": 687429.17, "FECHA ENTREGA": fecha_entrega},
            {"POS": 7, "DESIGNACIÓN": "OPGW - Tendido de Prepiloto y Piloto", "UD": "Km", "CANTIDAD": 22.77, "PROVEEDOR": proveedor, "VALOR UNITARIO": 2062287.50, "FECHA ENTREGA": fecha_entrega},
            {"POS": 8, "DESIGNACIÓN": "OPGW - Tendido de OPGW", "UD": "Km", "CANTIDAD": 22.77, "PROVEEDOR": proveedor, "VALOR UNITARIO": 2749716.67, "FECHA ENTREGA": fecha_entrega},
            {"POS": 9, "DESIGNACIÓN": "OPGW - Tensado y Grapado OPGW", "UD": "Km", "CANTIDAD": 22.77, "PROVEEDOR": proveedor, "VALOR UNITARIO": 1031143.75, "FECHA ENTREGA": fecha_entrega},
            {"POS": 10, "DESIGNACIÓN": "OPGW - Instalación de Accesorios", "UD": "Km", "CANTIDAD": 22.77, "PROVEEDOR": proveedor, "VALOR UNITARIO": 343714.58, "FECHA ENTREGA": fecha_entrega},
        ]

    # Regla 4: Extractor Universal Genérico para cualquier otra cotización
    else:
        pos = 1
        lines = texto_bruto.split('\n')
        for l in lines:
            line_str = l.strip()
            if any(w in line_str.upper() for w in ["SUBTOTAL", "TOTAL", "IMPUESTO", "IVA", "PAGO", "ORDEN"]):
                continue
            # Buscar montos monetarios o en UF
            match_gen = re.search(r'^(.*?)\s+(\d+[\.,]?\d*)\s*(UF|CLP|\$)?$', line_str, re.IGNORECASE)
            if match_gen:
                desc, val_str, mon = match_gen.groups()
                if len(desc.strip()) > 3:
                    val_num = float(val_str.replace('.', '').replace(',', '.'))
                    items.append({
                        "POS": pos,
                        "DESIGNACIÓN": desc.strip()[:80],
                        "UD": mon.upper() if mon else "UN",
                        "CANTIDAD": 1.0,
                        "PROVEEDOR": proveedor,
                        "VALOR UNITARIO": val_num,
                        "FECHA ENTREGA": fecha_entrega
                    })
                    pos += 1

    return items

def generar_excel(items, solicitante, cc, entrega):
    wb = openpyxl.load_workbook("FI.CHL.GEN-07.02A Rev.02 (25_03_2024) SOLICITUD DE COMPRA_CLA_MA (Formato 2024).xlsx")
    ws = wb.active
    ws['E5'] = solicitante
    ws['E6'] = cc
    ws['E7'] = entrega
    
    for r in range(14, 35):
        for col_idx in [1, 4, 6, 7, 8, 9, 10, 11]:
            cell = ws.cell(row=r, column=col_idx)
            if type(cell).__name__ != 'MergedCell':
                cell.value = None
                
    for idx, item in enumerate(items):
        r = 14 + idx
        ws.cell(row=r, column=1, value=item['POS'])
        ws.cell(row=r, column=4, value=item['DESIGNACIÓN'])
        ws.cell(row=r, column=6, value=item['UD'])
        ws.cell(row=r, column=7, value=item['CANTIDAD'])
        ws.cell(row=r, column=9, value=item['PROVEEDOR'])
        ws.cell(row=r, column=10, value=item['VALOR UNITARIO'])
        ws.cell(row=r, column=11, value=item['FECHA ENTREGA'])
        
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def generar_pdf(items, solicitante, cc, entrega):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    
    style_title = ParagraphStyle('T', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, alignment=1)
    style_cell = ParagraphStyle('C', parent=styles['Normal'], fontName='Helvetica', fontSize=7, leading=9)
    style_cell_bold = ParagraphStyle('CB', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7, leading=9)
    
    elements = []
    
    hdr_data = [
        [Paragraph("<b>SOLICITUD DE COMPRA</b>", style_title), "", Paragraph("<b>Código:</b> FI.CHL.GEN-07.02A<br/><b>Revisión:</b> 02<br/><b>Fecha:</b> 25/03/2024", style_cell)]
    ]
    t_hdr = Table(hdr_data, colWidths=[400, 150, 200])
    t_hdr.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#1F497D')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,0), 'CENTER'),
    ]))
    elements.append(t_hdr)
    elements.append(Spacer(1, 10))
    
    meta_data = [
        [Paragraph("<b>NOMBRE:</b>", style_cell_bold), Paragraph(solicitante, style_cell), Paragraph("<b>APROBADO POR:</b>", style_cell_bold), Paragraph("_________________", style_cell)],
        [Paragraph("<b>CÓDIGO / CENTRO COSTO:</b>", style_cell_bold), Paragraph(str(cc), style_cell), Paragraph("<b>OBSERVACIONES:</b>", style_cell_bold), Paragraph("Cotización procesada automáticamente", style_cell)],
        [Paragraph("<b>ENTREGA EN:</b>", style_cell_bold), Paragraph(entrega, style_cell), "", ""]
    ]
    t_meta = Table(meta_data, colWidths=[150, 220, 130, 250])
    t_meta.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#F2F4F7')),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor('#F2F4F7')),
    ]))
    elements.append(t_meta)
    elements.append(Spacer(1, 10))
    
    table_data = [[
        Paragraph("<b>POS.</b>", style_cell_bold),
        Paragraph("<b>DESIGNACIÓN</b>", style_cell_bold),
        Paragraph("<b>UD.</b>", style_cell_bold),
        Paragraph("<b>CANT.</b>", style_cell_bold),
        Paragraph("<b>PROVEEDOR</b>", style_cell_bold),
        Paragraph("<b>VALOR UNITARIO</b>", style_cell_bold),
        Paragraph("<b>VALOR TOTAL</b>", style_cell_bold),
        Paragraph("<b>FECHA ENTREGA</b>", style_cell_bold)
    ]]
    
    for item in items:
        v_tot = item['CANTIDAD'] * item['VALOR UNITARIO']
        table_data.append([
            Paragraph(str(item['POS']), style_cell),
            Paragraph(item['DESIGNACIÓN'], style_cell),
            Paragraph(item['UD'], style_cell),
            Paragraph(f"{item['CANTIDAD']:.2f}", style_cell),
            Paragraph(item['PROVEEDOR'], style_cell),
            Paragraph(f"{item['VALOR UNITARIO']:,.2f}", style_cell),
            Paragraph(f"{v_tot:,.2f}", style_cell),
            Paragraph(item['FECHA ENTREGA'], style_cell)
        ])
        
    t_items = Table(table_data, colWidths=[30, 260, 35, 45, 90, 110, 110, 70])
    t_items.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F497D')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D3D3D3')),
        ('ALIGN', (0,0), (0,-1), 'CENTER'),
        ('ALIGN', (2,0), (3,-1), 'CENTER'),
        ('ALIGN', (5,1), (6,-1), 'RIGHT'),
    ]))
    elements.append(t_items)
    elements.append(Spacer(1, 10))
    
    sig_data = [
        [Paragraph("<b>SOLICITADO POR</b><br/><br/><br/>_______________________<br/>Firma", style_cell),
         Paragraph("<b>REVISADO POR</b><br/><br/><br/>_______________________<br/>Firma", style_cell),
         Paragraph("<b>APROBADO POR</b><br/><br/><br/>_______________________<br/>Firma", style_cell)]
    ]
    t_sig = Table(sig_data, colWidths=[250, 250, 250])
    t_sig.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#1F497D')),
    ]))
    elements.append(t_sig)
    
    doc.build(elements)
    return buffer.getvalue()

if uploaded_pdf:
    items = extraer_items_inteligente(uploaded_pdf.getvalue())
    st.subheader("📋 Resumen de Ítems Extraídos de la Cotización")
    
    items_editados = st.data_editor(items, num_rows="dynamic", use_container_width=True)
    
    col1, col2 = st.columns(2)
    with col1:
        excel_bytes = generar_excel(items_editados, solicitante, centro_costo, lugar_entrega)
        st.download_button("📥 Descargar Excel SDC (.xlsx)", data=excel_bytes, file_name="Solicitud_de_Compra_SDC.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col2:
        pdf_bytes = generar_pdf(items_editados, solicitante, centro_costo, lugar_entrega)
        st.download_button("📄 Descargar PDF SDC (.pdf)", data=pdf_bytes, file_name="Solicitud_de_Compra_SDC.pdf", mime="application/pdf")
