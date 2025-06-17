import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
import io
import base64
# Librerías para código de barras
from barcode import Code128
from barcode.writer import ImageWriter
from PIL import Image

st.set_page_config(page_title="Generador de Etiquetas", layout="wide")

# CSS personalizado
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stButton > button {
        width: 100%;
        background-color: #4CAF50;
        color: white;
        padding: 0.5rem;
        border-radius: 5px;
    }
</style>
""", unsafe_allow_html=True)

# Título principal
st.title("🏷️ Generador de Etiquetas de Importación 76x25mm")
st.markdown("---")

# Información fija del importador (sin No. PARTE ya que irá en código de barras)
info_importador = [
    "**IMPORTADOR: **MOTORMAN DE BAJA CALIFORNIA SA DE CV",
    "MARISCAL SUCRE 6738 LA CIENEGA PONIENTE,",
    "TIJUANA, B.C. 22114 RFC: MBC210723RP9",
    "**DESCRIPCION:**",
    "**CONTENIDO:**",
    "**HECHO EN:**"
]

# Crear columnas para el layout
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("📝 Datos de la Etiqueta")
    
    # Campos de entrada
    descripcion = st.text_input("DESCRIPCIÓN:", value="ABANICO PARA RADIADOR CON MOTOR")
    
    # Contenido con número y unidad
    col_cont1, col_cont2 = st.columns([1, 3])
    with col_cont1:
        cantidad_contenido = st.number_input("CONTENIDO:", min_value=1, max_value=1000, value=1, step=1)
    with col_cont2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.text("PIEZA(S)")
    
    hecho_en = st.text_input("HECHO EN:", value="CHINA")
    numero_parte = st.text_input("No. PARTE (SKU):", value="12345-ABC")
    cantidad_etiquetas = st.number_input("CANTIDAD DE ETIQUETAS:", min_value=1, max_value=100, value=10, step=1)

with col2:
    st.subheader("👁️ Vista Previa")
    
    # Crear vista previa con HTML
    if cantidad_contenido == 1:
        contenido_text = f"{cantidad_contenido} PIEZA"
    else:
        contenido_text = f"{cantidad_contenido} PIEZAS"
    
    preview_html = f"""
    <div style="border: 2px solid black; padding: 10px; width: 228px; height: 85px; background-color: white; font-family: Arial; font-size: 7px; line-height: 1.2; position: relative;">
        <div><strong>IMPORTADOR:</strong> MOTORMAN DE BAJA CALIFORNIA SA DE CV</div>
        <div>MARISCAL SUCRE 6738 LA CIENEGA PONIENTE,</div>
        <div>TIJUANA, B.C. 22114 RFC: MBC210723RP9</div>
        <div><strong>DESCRIPCION:</strong> {descripcion.upper()}</div>
        <div><strong>CONTENIDO:</strong> {contenido_text}</div>
        <div><strong>HECHO EN:</strong> {hecho_en.upper()}</div>
        <div style="position: absolute; bottom: 5px; right: 5px; border: 1px solid black; padding: 2px; font-size: 6px;">
            <div style="background: repeating-linear-gradient(90deg, black, black 1px, white 1px, white 2px); height: 15px; width: 60px;"></div>
            <div style="text-align: center; font-size: 5px; margin-top: 2px;">{numero_parte.upper()}</div>
        </div>
    </div>
    """
    st.markdown(preview_html, unsafe_allow_html=True)

st.markdown("---")

# Funciones para generar archivos
def generar_pdf_etiquetas(datos):
    """Genera un PDF con las etiquetas y código de barras"""
    buffer = io.BytesIO()
    
    # Crear canvas
    c = canvas.Canvas(buffer, pagesize=(76*mm, 25*mm))
    
    # Para cada etiqueta
    for etiqueta in datos:
        # Determinar texto de contenido
        cantidad_contenido = etiqueta.get('cantidad_contenido', 1)
        if cantidad_contenido == 1:
            contenido = f"{cantidad_contenido} PIEZA"
        else:
            contenido = f"{cantidad_contenido} PIEZAS"
        
        # Convertir valores a mayúsculas
        descripcion = etiqueta['descripcion'].upper()
        hecho_en = etiqueta['hecho_en'].upper()
        numero_parte = etiqueta.get('numero_parte', '').upper()
        
        # Configurar fuente
        font_size = 6
        
        # Posiciones para cada línea
        y_positions = [22*mm, 19.5*mm, 17*mm, 14.5*mm, 12*mm, 9.5*mm]
        
        # Primera línea: IMPORTADOR en negrita, resto normal
        c.setFont("Helvetica-Bold", font_size)
        c.drawString(5*mm, y_positions[0], "IMPORTADOR: ")
        ancho_importador = c.stringWidth("IMPORTADOR: ", "Helvetica-Bold", font_size)
        c.setFont("Helvetica", font_size)
        c.drawString(5*mm + ancho_importador, y_positions[0], "MOTORMAN DE BAJA CALIFORNIA SA DE CV")
        
        # Segunda y tercera línea (normal)
        c.setFont("Helvetica", font_size)
        c.drawString(5*mm, y_positions[1], "MARISCAL SUCRE 6738 LA CIENEGA PONIENTE,")
        c.drawString(5*mm, y_positions[2], "TIJUANA, B.C. 22114 RFC: MBC210723RP9")
        
        # Cuarta línea: DESCRIPCION en negrita
        c.setFont("Helvetica-Bold", font_size)
        c.drawString(5*mm, y_positions[3], "DESCRIPCION: ")
        ancho_descripcion = c.stringWidth("DESCRIPCION: ", "Helvetica-Bold", font_size)
        c.setFont("Helvetica", font_size)
        c.drawString(5*mm + ancho_descripcion, y_positions[3], descripcion)
        
        # Quinta línea: CONTENIDO en negrita
        c.setFont("Helvetica-Bold", font_size)
        c.drawString(5*mm, y_positions[4], "CONTENIDO: ")
        ancho_contenido = c.stringWidth("CONTENIDO: ", "Helvetica-Bold", font_size)
        c.setFont("Helvetica", font_size)
        c.drawString(5*mm + ancho_contenido, y_positions[4], contenido)
        
        # Sexta línea: HECHO EN en negrita
        c.setFont("Helvetica-Bold", font_size)
        c.drawString(5*mm, y_positions[5], "HECHO EN: ")
        ancho_hecho = c.stringWidth("HECHO EN: ", "Helvetica-Bold", font_size)
        c.setFont("Helvetica", font_size)
        
        # IMPORTANTE: Verificar si el texto es muy largo y ajustarlo
        texto_hecho_en = hecho_en
        ancho_disponible = 45*mm  # Más espacio disponible (era 35mm)
        ancho_texto_actual = c.stringWidth(texto_hecho_en, "Helvetica", font_size)
        
        # Si el texto es muy largo, reducir el tamaño de fuente
        font_size_hecho = font_size
        while ancho_texto_actual > ancho_disponible and font_size_hecho > 4:
            font_size_hecho -= 0.5
            c.setFont("Helvetica", font_size_hecho)
            ancho_texto_actual = c.stringWidth(texto_hecho_en, "Helvetica", font_size_hecho)
        
        c.drawString(5*mm + ancho_hecho, y_positions[5], texto_hecho_en)
        
        # Código de barras
        if numero_parte:
            try:
                # Generar código de barras
                barcode_buffer = io.BytesIO()
                code = Code128(numero_parte, writer=ImageWriter())
                
                # Opciones para código de barras más compacto
                options = {
                    'module_width': 0.3,      # Más delgado (reducido de 0.2)
                    'module_height': 4,        # Menos alto (reducido de 5)
                    'font_size': 5,            # Texto más pequeño (reducido de 7)
                    'text_distance': 2,        # Menos separación (reducido de 3)
                    'quiet_zone': 2,           # Margen mínimo (reducido de 2)
                    'write_text': True         # Mostrar texto
                }
                
                code.write(barcode_buffer, options=options)
                barcode_buffer.seek(0)
                
                # Cargar imagen
                barcode_image = Image.open(barcode_buffer)
                
                # Posición y tamaño - AÚN MÁS PEQUEÑO Y MÁS A LA DERECHA
                barcode_width = 40*mm      # Más pequeño (era 30mm)
                barcode_height = 9*mm      # Menos alto (era 10mm)
                barcode_x = (76*mm - barcode_width) / 2  # CENTRADO HORIZONTALMENTE
                barcode_y = 0.5*mm         # Pegado al fondo
                
                # Dibujar imagen
                c.drawInlineImage(barcode_image, barcode_x, barcode_y, 
                                width=barcode_width, height=barcode_height)
            
            except Exception as e:
                print(f"Error con código de barras: {e}")
                # Si falla el código de barras, al menos mostrar el texto
                c.setFont("Helvetica", 5)
                c.drawString(barcode_x, barcode_y, f"[{numero_parte}]")
        
        c.showPage()
    
    c.save()
    buffer.seek(0)
    return buffer

def generar_excel_etiquetas(datos):
    """Genera un Excel con las etiquetas"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Etiquetas"
    
    # Estilos
    borde = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    alineacion = Alignment(wrap_text=True, vertical='center', horizontal='left')
    normal = Font(bold=False, size=8)
    
    # Configurar columnas
    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 25
    
    # Crear etiquetas
    fila_actual = 1
    col_actual = 1
    
    for etiqueta in datos:
        celda = ws.cell(row=fila_actual, column=col_actual)
        
        # Preparar contenido
        cantidad_contenido = etiqueta.get('cantidad_contenido', 1)
        if cantidad_contenido == 1:
            contenido = f"{cantidad_contenido} PIEZA"
        else:
            contenido = f"{cantidad_contenido} PIEZAS"
        
        descripcion = etiqueta['descripcion'].upper()
        hecho_en = etiqueta['hecho_en'].upper()
        numero_parte = etiqueta.get('numero_parte', '').upper()
        
        # Contenido de la etiqueta
        if numero_parte:
            contenido_etiqueta = f"""IMPORTADOR: MOTORMAN DE BAJA CALIFORNIA SA DE CV
MARISCAL SUCRE 6738 LA CIENEGA PONIENTE,
TIJUANA, B.C. 22114 RFC: MBC210723RP9
DESCRIPCION: {descripcion}
CONTENIDO: {contenido}
HECHO EN: {hecho_en}
[CÓDIGO DE BARRAS: {numero_parte}]"""
        else:
            contenido_etiqueta = f"""IMPORTADOR: MOTORMAN DE BAJA CALIFORNIA SA DE CV
MARISCAL SUCRE 6738 LA CIENEGA PONIENTE,
TIJUANA, B.C. 22114 RFC: MBC210723RP9
DESCRIPCION: {descripcion}
CONTENIDO: {contenido}
HECHO EN: {hecho_en}"""
        
        celda.value = contenido_etiqueta
        celda.font = normal
        celda.alignment = alineacion
        celda.border = borde
        
        ws.row_dimensions[fila_actual].height = 21
        
        # Avanzar posición
        col_actual += 1
        if col_actual > 3:
            col_actual = 1
            fila_actual += 1
    
    # Guardar en buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# Sección de acciones
st.subheader("🚀 Acciones")

col_btn1, col_btn2, col_btn3 = st.columns(3)

with col_btn1:
    if st.button("📄 Generar PDF", type="primary"):
        # Preparar datos
        datos = []
        for i in range(cantidad_etiquetas):
            datos.append({
                'descripcion': descripcion,
                'cantidad_contenido': cantidad_contenido,
                'hecho_en': hecho_en,
                'numero_parte': numero_parte
            })
        
        # Generar PDF
        pdf_buffer = generar_pdf_etiquetas(datos)
        
        # Botón de descarga
        st.download_button(
            label="⬇️ Descargar PDF",
            data=pdf_buffer,
            file_name="etiquetas_importacion.pdf",
            mime="application/pdf"
        )

with col_btn2:
    if st.button("📊 Generar Excel"):
        # Preparar datos
        datos = []
        for i in range(cantidad_etiquetas):
            datos.append({
                'descripcion': descripcion,
                'cantidad_contenido': cantidad_contenido,
                'hecho_en': hecho_en,
                'numero_parte': numero_parte
            })
        
        # Generar Excel
        excel_buffer = generar_excel_etiquetas(datos)
        
        # Botón de descarga
        st.download_button(
            label="⬇️ Descargar Excel",
            data=excel_buffer,
            file_name="etiquetas_importacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with col_btn3:
    # Plantilla de ejemplo
    if st.button("📋 Descargar Plantilla"):
        # Crear datos de ejemplo con numero_parte
        datos_ejemplo = pd.DataFrame([
            {"descripcion": "ABANICO PARA RADIADOR CON MOTOR", "cantidad_contenido": 1, "hecho_en": "CHINA", "numero_parte": "FAN-12345", "cantidad_etiquetas": 10},
            {"descripcion": "BOMBA DE AGUA", "cantidad_contenido": 2, "hecho_en": "JAPÓN", "numero_parte": "WP-67890", "cantidad_etiquetas": 5},
            {"descripcion": "FILTRO DE ACEITE", "cantidad_contenido": 5, "hecho_en": "MÉXICO", "numero_parte": "OF-11223", "cantidad_etiquetas": 20},
        ])
        
        # Convertir a Excel
        buffer = io.BytesIO()
        datos_ejemplo.to_excel(buffer, index=False)
        buffer.seek(0)
        
        st.download_button(
            label="⬇️ Descargar Plantilla Excel",
            data=buffer,
            file_name="plantilla_etiquetas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Sección de importación desde Excel
st.markdown("---")
st.subheader("📥 Importar desde Excel")

uploaded_file = st.file_uploader("Seleccione un archivo Excel", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        st.success(f"✅ Archivo cargado: {len(df)} productos encontrados")
        
        # Mostrar datos
        st.dataframe(df)
        
        # Selección de productos
        selected_indices = st.multiselect(
            "Seleccione los productos para generar etiquetas:",
            options=df.index,
            default=df.index.tolist(),
            format_func=lambda x: f"{df.loc[x, 'descripcion']} - {df.loc[x, 'cantidad_etiquetas']} etiquetas"
        )
        
        if selected_indices:
            col_imp1, col_imp2 = st.columns(2)
            
            with col_imp1:
                if st.button("📄 Generar PDF de Seleccionados"):
                    # CORRECCIÓN IMPORTANTE: No agrupar por descripción
                    etiquetas_datos = []
                    
                    # Procesar cada índice seleccionado INDIVIDUALMENTE
                    for idx in selected_indices:
                        row = df.loc[idx]
                        
                        # Obtener datos ESPECÍFICOS de ESTA fila
                        descripcion = str(row.get('descripcion', ''))
                        cantidad_contenido = int(row.get('cantidad_contenido', 1))
                        hecho_en = str(row.get('hecho_en', ''))
                        cantidad_etiquetas = int(row.get('cantidad_etiquetas', 1))
                        
                        # IMPORTANTE: Obtener el número de parte de ESTA fila específica
                        numero_parte = ""
                        if "numero_parte" in row and pd.notna(row["numero_parte"]):
                            numero_parte = str(row["numero_parte"])
                        elif "sku" in row and pd.notna(row["sku"]):
                            numero_parte = str(row["sku"])
                        elif "part_number" in row and pd.notna(row["part_number"]):
                            numero_parte = str(row["part_number"])
                        
                        # Generar las etiquetas para ESTA fila específica
                        # NO agrupar con otras filas
                        for _ in range(cantidad_etiquetas):
                            etiquetas_datos.append({
                                'descripcion': descripcion,
                                'cantidad_contenido': cantidad_contenido,
                                'hecho_en': hecho_en,
                                'numero_parte': numero_parte  # Número de parte ESPECÍFICO de esta fila
                            })
                    
                    # Generar PDF
                    pdf_buffer = generar_pdf_etiquetas(etiquetas_datos)
                    
                    st.download_button(
                        label=f"⬇️ Descargar PDF ({len(etiquetas_datos)} etiquetas)",
                        data=pdf_buffer,
                        file_name="etiquetas_seleccionadas.pdf",
                        mime="application/pdf"
                    )
            
            with col_imp2:
                if st.button("📊 Generar Excel de Seleccionados"):
                    # CORRECCIÓN IMPORTANTE: No agrupar por descripción
                    etiquetas_datos = []
                    
                    # Procesar cada índice seleccionado INDIVIDUALMENTE
                    for idx in selected_indices:
                        row = df.loc[idx]
                        
                        # Obtener datos ESPECÍFICOS de ESTA fila
                        descripcion = str(row.get('descripcion', ''))
                        cantidad_contenido = int(row.get('cantidad_contenido', 1))
                        hecho_en = str(row.get('hecho_en', ''))
                        cantidad_etiquetas = int(row.get('cantidad_etiquetas', 1))
                        
                        # IMPORTANTE: Obtener el número de parte de ESTA fila específica
                        numero_parte = ""
                        if "numero_parte" in row and pd.notna(row["numero_parte"]):
                            numero_parte = str(row["numero_parte"])
                        elif "sku" in row and pd.notna(row["sku"]):
                            numero_parte = str(row["sku"])
                        elif "part_number" in row and pd.notna(row["part_number"]):
                            numero_parte = str(row["part_number"])
                        
                        # Generar las etiquetas para ESTA fila específica
                        # NO agrupar con otras filas
                        for _ in range(cantidad_etiquetas):
                            etiquetas_datos.append({
                                'descripcion': descripcion,
                                'cantidad_contenido': cantidad_contenido,
                                'hecho_en': hecho_en,
                                'numero_parte': numero_parte  # Número de parte ESPECÍFICO de esta fila
                            })
                    
                    # Generar Excel
                    excel_buffer = generar_excel_etiquetas(etiquetas_datos)
                    
                    st.download_button(
                        label=f"⬇️ Descargar Excel ({len(etiquetas_datos)} etiquetas)",
                        data=excel_buffer,
                        file_name="etiquetas_seleccionadas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    
    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {str(e)}")

# Footer
st.markdown("---")
st.markdown("🏭 **Generador de Etiquetas de Importación** - MOTORMAN DE BAJA CALIFORNIA SA DE CV")