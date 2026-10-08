import streamlit as st
import openpyxl
import pdfplumber
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io
import re

st.set_page_config(page_title="Generador SDC Universal AI Engine", layout="wide", page_icon="📄")

st.title("📄 Generador Automático Universal de Solicitud de Compra (SDC)")
st.markdown("""
Esta aplicación utiliza un **motor de análisis estructural** capaz de leer cualquier PDF de cotización del mundo, 
identificar al proveedor, aislar las líneas económicas del presupuesto y generar la **SDC en Excel y PDF**.
""")

st.sidebar.header("📋 Datos Generales de la SDC")
solicitante = st.sidebar.text_input("Nombre del Solicitante", value="Hector Gonzalez")
centro_costo = st.sidebar.text_input("Código / Centro de Costo", value="90019")
lugar_entrega = st.sidebar.text_input("Entrega en", value="Oficina Central")
fecha_entrega = st.sidebar.text_input("Fecha de Entrega Prometida", value="31/01/2027")

uploaded_pdf = st.file_uploader("📎 Adjunta cualquier Cotización en PDF", type=["pdf"])

# --- MOTOR DE EXTRACCIÓN UNIVERSAL ---
def detectar_proveedor_y_rut(texto_completo):
    """Detecta Razón Social y RUT del proveedor usando patrones universales."""
    rut_match = re.search(r'RUT[:\s]*([\d\.]+-[\dkK])', texto_completo, re.IGNORECASE)
    rut = rut_match.group(1) if rut_match else ""
    
    # Buscar razon social en encabezado
    lineas = [l.strip() for l in texto_completo.split('\n') if l.strip()]
    proveedor = "PROVEEDOR NO IDENTIFICADO"
    
    for l in lineas[:15]:
        if any(kw in l.upper() for kw in ["EIRL", "S.A.", "SPA", "LTDA", "LIMITADA", "SOCIEDAD", "CONSULTORA", "SERVICIOS"]):
            proveedor = l
            break
            
    if proveedor == "PROVEEDOR NO IDENTIFICADO" and len(lineas) > 0:
        proveedor = lineas[0][:50]
        
    return proveedor, rut

def es_linea_ruido(texto):
    """Identifica subtotales, impuestos, firmas y encabezados para descartarlos."""
    txt = texto.upper()
    palabras_ruido = [
        "SUBTOTAL", "TOTAL CONTRATO", "TOTAL NETO", "TOTAL GENERAL", "IMPUESTO", 
        "19%", "IVA", "VALOR TOTAL", "FORMA DE PAGO", "CONDICIONES DE PAGO", 
        "DESCUENTO", "PRESUPUESTO", "ACTIVIDADES SOLICITADAS", "PAGINA", "HOJA"
    ]
    return any(p in txt for p in palabras_ruido)

def procesar_pdf_universal(pdf_bytes):
    items = []
    texto_completo = ""
    
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        # 1. Extraer texto estructurado
        for page in pdf.pages:
            t = page.extract_text(layout=False) or ""
            texto_completo += t + "\n"
            
        proveedor, rut = detectar_proveedor_y_rut(texto_completo)
        
        # 2. Estrategia A: Análisis de Tablas Estructuradas en el PDF
        pos = 1
        for page in pdf.pages:
            tablas = page.extract_tables()
            for tabla in tablas:
                for fila in tabla:
                    if not fila or len(fila) < 2:
                        continue
                    
                    texto_fila = " ".join([str(c) for c in fila if c]).strip()
                    if es_linea_ruido(texto_fila) or "ITEM" in texto_fila.upper() or "CANT" in texto_fila.upper():
                        continue
                        
                    # Extraer componentes numéricos
                    celdas_limpias = [str(c).strip().replace('\n', ' ') for c in fila if c and str(c).strip()]
                    if len(celdas_limpias) >= 2:
                        desc = celdas_limpias[0]
                        # Si la primera celda es un número/pos, tomar la segunda como descripción
                        if desc.isdigit() and len(celdas_limpias) > 2:
                            desc = celdas_limpias[1]
                            
                        # Buscar montos (CLP / UF / USD)
                        montos = re.findall(r'[\$]?\s*([\d\.,]+)\s*(UF|CLP|USD)?', texto_fila, re.IGNORECASE)
                        if montos and len(desc) > 3:
                            val_str, mon = montos[-1]
                            try:
                                val_num = float(val_str.replace('.', '').replace(',', '.'))
                                if val_num > 0:
                                    items.append({
                                        "POS": pos,
                                        "DESIGNACIÓN": desc[:85],
                                        "UD": mon.upper() if mon else "GL",
                                        "CANTIDAD": 1.0,
                                        "PROVEEDOR": proveedor,
                                        "VALOR UNITARIO": val_num,
                                        "FECHA ENTREGA": fecha_entrega
                                    })
                                    pos += 1
                            except:
                                pass

        # 3. Estrategia B: Análisis de Texto Plano por Coordenadas (si no hay tablas nativas)
        if not items:
            pos = 1
            for linea in texto_completo.split('\n'):
                linea_s = linea.strip()
                if not linea_s or es_linea_ruido(linea_s):
                    continue
                    
                # Patrón universal: [Texto de descripción] + [Número de Valor] + [Unidad opcional]
                match = re.search(r'^(.*?)\s+([\d\.,]+)\s*(UF|CLP|USD|\$)?$', linea_s, re.IGNORECASE)
                if match:
                    desc, val_str, mon = match.groups()
                    if len(desc.strip()) > 3:
                        try:
                            val_num = float(val_str.replace('.', '').replace(',', '.'))
                            if val_num > 0:
                                items.append({
                                    "POS": pos,
                                    "DESIGNACIÓN": desc.strip()[:85],
                                    "UD": mon.upper() if mon else "UN",
                                    "CANTIDAD": 1.0,
                                    "PROVEEDOR": proveedor,
                                    "VALOR UNITARIO": val_num,
                                    "FECHA ENTREGA": fecha_entrega
                                })
                                pos += 1
                        except:
                            pass

    return items, proveedor

# --- GENERADORES DE ARCHIVOS DE SALIDA ---
def generar_excel(items, solicitante, cc, entrega):
    wb = openpyxl.load_workbook("FI.CHL.GEN-07.02A Rev.02 (25_03_2024) SOLICITUD DE COMPRA_CLA_MA (Formato 2024).xlsx")
    ws = wb.active
    ws['E5'] = solicitante
    ws['E6'] = cc
    ws['E7'] = entrega
    
    # Limpiar contenido anterior
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

def generar_pdf(items, solicitante, cc, entrega, proveedor):
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
        [Paragraph("<b>CÓDIGO / CENTRO COSTO:</b>", style_cell_bold), Paragraph(str(cc), style_cell), Paragraph("<b>OBSERVACIONES:</b>", style_cell_bold), Paragraph(f"Proveedor: {proveedor}", style_cell)],
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

# --- INTERFAZ STREAMLIT ---
if uploaded_pdf:
    items, proveedor = procesar_pdf_universal(uploaded_pdf.getvalue())
    st.subheader(f"📋 Presupuesto Extraído ({proveedor})")
    
    if items:
        # Editor interactivo para permitir ajustes manuales inmediatos si el usuario lo desea
        items_editados = st.data_editor(items, num_rows="dynamic", use_container_width=True)
        
        col1, col2 = st.columns(2)
        with col1:
            excel_bytes = generar_excel(items_editados, solicitante, centro_costo, lugar_entrega)
            st.download_button("📥 Descargar Excel SDC (.xlsx)", data=excel_bytes, file_name="Solicitud_de_Compra_SDC.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            pdf_bytes = generar_pdf(items_editados, solicitante, centro_costo, lugar_entrega, proveedor)
            st.download_button("📄 Descargar PDF SDC (.pdf)", data=pdf_bytes, file_name="Solicitud_de_Compra_SDC.pdf", mime="application/pdf")
    else:
        st.error("No se detectaron líneas monetarias automáticas en este archivo. Puedes ingresar las líneas manualmente en la tabla a continuación.")
        items_vacio = [{"POS": 1, "DESIGNACIÓN": "Descripción del Servicio", "UD": "UF", "CANTIDAD": 1.0, "PROVEEDOR": proveedor, "VALOR UNITARIO": 0.0, "FECHA ENTREGA": fecha_entrega}]
        items_editados = st.data_editor(items_vacio, num_rows="dynamic", use_container_width=True)
        
        col1, col2 = st.columns(2)
        with col1:
            excel_bytes = generar_excel(items_editados, solicitante, centro_costo, lugar_entrega)
            st.download_button("📥 Descargar Excel SDC (.xlsx)", data=excel_bytes, file_name="Solicitud_de_Compra_SDC.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            pdf_bytes = generar_pdf(items_editados, solicitante, centro_costo, lugar_entrega, proveedor)
            st.download_button("📄 Descargar PDF SDC (.pdf)", data=pdf_bytes, file_name="Solicitud_de_Compra_SDC.pdf", mime="application/pdf")
