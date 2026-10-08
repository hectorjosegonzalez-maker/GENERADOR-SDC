import streamlit as st
import openpyxl
import pdfplumber
import json
import io
import re
from openai import OpenAI  # O el cliente de la API de IA de tu preferencia
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Generador SDC con IA", layout="wide", page_icon="📄")

st.title("📄 Generador Universal de Solicitudes de Compra (SDC)")
st.markdown("Sube **cualquier cotización en PDF**. La IA analizará el documento, aislará la **tabla económica limpia** y generará la SDC.")

# Configuración de Clave API (Guardada en secretos de Streamlit)
api_key = st.sidebar.text_input("Llave API de IA (OpenAI / Gemini)", type="password")
solicitante = st.sidebar.text_input("Nombre del Solicitante", value="Hector Gonzalez")
centro_costo = st.sidebar.text_input("Código / Centro de Costo", value="90019")
lugar_entrega = st.sidebar.text_input("Entrega en", value="Oficina Central")
fecha_entrega = st.sidebar.text_input("Fecha de Entrega Prometida", value="31/01/2027")

uploaded_pdf = st.file_uploader("📎 Adjunta cualquier Cotización en PDF", type=["pdf"])

def analizar_pdf_con_ia(texto_pdf, api_key):
    """Envia el texto completo a la IA para extraer únicamente la tabla económica limpia."""
    client = OpenAI(api_key=api_key)
    
    prompt = f"""
    Eres un experto en compras corporativas. Revisa el texto de la siguiente cotización y extrae ÚNICAMENTE la tabla económica del presupuesto neto.
    
    REGLAS ESTRICTAS:
    1. Extrae solo las posiciones ejecutables de compra (descripción, unidad, cantidad, valor unitario y proveedor).
    2. IGNORA completamente filas de Subtotal, IVA, Impuestos, Totales Generales, Formas de pago y Descuentos.
    3. IGNORA tablas secundarias que no sean de cobro (listado de predios, coordinaciones, itinerarios, equipo de trabajo).
    4. Devuelve el resultado en formato JSON estricto con las siguientes llaves:
       - "proveedor": "Razón social del emisor"
       - "items": [
           {{"pos": 1, "designacion": "...", "ud": "UF/CLP", "cantidad": 1, "val_unitario": 0.0}}
         ]
    
    TEXTO DE LA COTIZACIÓN:
    {texto_pdf}
    """
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        response_format={"type": "json_object"},
        messages=[{"role": "user", "content": prompt}]
    )
    
    return json.loads(response.choices[0].message.content)

# Lógica de generación de Excel y PDF...
