import streamlit as st
import openpyxl
import pdfplumber
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io
import re

st.set_page_config(page_title="Generador de Solicitudes de Compra (SDC)", layout="wide", page_icon="📄")

st.title("📄 Generador Automático de Solicitud de Compra (SDC)")
st.markdown("Sube **cualquier cotización en PDF** para extraer sus ítems y generar la **Solicitud de Compra (SDC)** en Excel y PDF.")

st.sidebar.header("📋 Datos Generales")
solicitante = st.sidebar.text_input("Nombre del Solicitante", value="Hector Gonzalez")
centro_costo = st.sidebar.text_input("Código / Centro de Costo", value="90019")
lugar_entrega = st.sidebar.text_input("Entrega en", value="Oficina Central")
fecha_entrega = st.sidebar.text_input("Fecha de Entrega Prometida", value="31/01/2027")

uploaded_pdf = st.file_uploader("📎 Adjunta la Cotización (PDF)", type=["pdf"])

def extraer_items_de_pdf(pdf_bytes):
    items = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        pos = 1
        for page in pdf.pages:
            # 1. Intentar extraer tablas estructuradas del PDF
            tablas = page.extract_tables()
            for tabla in tablas:
                for fila in tabla:
                    if not fila or len(fila) < 3:
                        continue
                    # Filtrar encabezados comunes de tablas
                    texto_fila = " ".join([str(c) for c in fila if c]).upper()
                    if "DESCRIP" in texto_fila or "CANT" in texto_fila or "ITEM" in texto_fila or "PRECIO" in texto_fila:
                        continue
                    
                    # Intentar obtener descripción, cantidad y precio
                    textos = [str(c).strip().replace('\n', ' ') for c in fila if c]
                    if len(textos) >= 3:
                        desc = textos[0] if len(textos[0]) > 3 else textos[1]
                        # Buscar números en la fila
                        nums = [re.sub(r'[^\d,\.]', '', t) for t in textos if re.search(r'\d', t)]
                        val_unit = 0.0
                        cant = 1.0
                        if len(nums) >= 2:
                            try:
                                cant = float(nums[0].replace(',', '.'))
                                val_unit = float(nums[1].replace('.', '').replace(',', '.'))
                            except:
                                pass
                        elif len(nums) == 1:
                            try:
                                val_unit = float(nums[0].replace('.', '').replace(',', '.'))
                            except:
                                pass
                        
                        if desc and len(desc) > 2:
                            items.append({
                                "pos": pos,
                                "designacion": desc[:80],
                                "ud": "Unid",
                                "cantidad": cant,
                                "proveedor": "PROVEEDOR",
                                "val_unitario": val_unit,
                                "fecha_entrega": fecha_entrega
                            })
                            pos += 1
            
            # 2. Si no encontró tablas formales, buscar por líneas de texto
            if not items:
                texto = page.extract_text() or ""
                lineas = texto.split('\n')
                for l in lineas:
                    # Buscar líneas con montos
                    match = re.search(r'^(.*?)\s+(\d+[\.,]?\d*)\s+(Km|M|UN|GL|UF|UNID)?\s*[\$]?\s*([\d\.,]+)', l.strip(), re.IGNORECASE)
                    if match:
                        desc, cant, ud, val = match.groups()
                        if len(desc.strip()) > 3:
                            try:
                                v_u = float(val.replace('.', '').replace(',', '.'))
                                c_u = float(cant.replace(',', '.'))
                                items.append({
                                    "pos": pos,
                                    "designacion": desc.strip()[:80],
                                    "ud": ud if ud else "GL",
                                    "cantidad": c_u,
                                    "proveedor": "PROVEEDOR",
                                    "val_unitario": v_u,
                                    "fecha_entrega": fecha_entrega
                                })
                                pos += 1
                            except:
                                pass
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
        ws.cell(row=r, column=1, value=item['pos'])
        ws.cell(row=r, column=4, value=item['designacion'])
        ws.cell(row=r, column=6, value=item['ud'])
        ws.cell(row=r, column=7, value=item['cantidad'])
        ws.cell(row=r, column=9, value=item['proveedor'])
        ws.cell(row=r, column=10, value=item['val_unitario'])
        ws.cell(row=r, column=11, value=item['fecha_entrega'])
        
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
        v_tot = item['cantidad'] * item['val_unitario']
        table_data.append([
            Paragraph(str(item['pos']), style_cell),
            Paragraph(item['designacion'], style_cell),
            Paragraph(item['ud'], style_cell),
            Paragraph(f"{item['cantidad']:.2f}", style_cell),
            Paragraph(item['proveedor'], style_cell),
            Paragraph(f"$ {item['val_unitario']:,.2f}", style_cell),
            Paragraph(f"$ {v_tot:,.2f}", style_cell),
            Paragraph(item['fecha_entrega'], style_cell)
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
    st.info("🔄 Procesando PDF de la cotización...")
    items = extraer_items_de_pdf(uploaded_pdf.getvalue())
    
    if items:
        st.success(f"✅ Se extrajeron automáticamente {len(items)} ítems de la cotización.")
        
        # Permitir editar la tabla directamente en pantalla si el usuario desea corregir algo
        items_editados = st.data_editor(items, num_rows="dynamic", use_container_width=True)
        
        col1, col2 = st.columns(2)
        with col1:
            excel_bytes = generar_excel(items_editados, solicitante, centro_costo, lugar_entrega)
            st.download_button("📥 Descargar Excel SDC (.xlsx)", data=excel_bytes, file_name="Solicitud_de_Compra_SDC.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            pdf_bytes = generar_pdf(items_editados, solicitante, centro_costo, lugar_entrega)
            st.download_button("📄 Descargar PDF SDC (.pdf)", data=pdf_bytes, file_name="Solicitud_de_Compra_SDC.pdf", mime="application/pdf")
    else:
        st.warning("⚠️ No se pudieron extraer tablas automáticas de este PDF. Puedes agregar los ítems manualmente abajo.")
