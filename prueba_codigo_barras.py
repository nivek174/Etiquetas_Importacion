import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from barcode import Code128
from barcode.writer import ImageWriter
import io
from PIL import Image, ImageTk
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import os
import tempfile

class PruebaEtiquetaConCodigoBarras:
    def __init__(self, root):
        self.root = root
        self.root.title("Prueba de Etiqueta 76x25mm con Código de Barras")
        self.root.geometry("800x600")
        
        # Información fija del importador
        self.info_importador = [
            "**IMPORTADOR: **MOTORMAN DE BAJA CALIFORNIA SA DE CV",
            "MARISCAL SUCRE 6738 LA CIENEGA PONIENTE,",
            "TIJUANA, B.C. 22114 RFC: MBC210723RP9",
            "**DESCRIPCION:**",
            "**CONTENIDO:**",
            "**HECHO EN:**"
        ]
        
        # Frame principal
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        ttk.Label(main_frame, text="Prueba de Etiqueta con Código de Barras", 
                 font=("Arial", 16, "bold")).pack(pady=10)
        
        # Frame para entrada de datos
        datos_frame = ttk.LabelFrame(main_frame, text="Datos de la Etiqueta", padding="10")
        datos_frame.pack(fill=tk.X, pady=10)
        
        # Campos de entrada
        ttk.Label(datos_frame, text="DESCRIPCIÓN:").grid(row=0, column=0, sticky="w", pady=5)
        self.descripcion_var = tk.StringVar(value="ABANICO PARA RADIADOR CON MOTOR")
        ttk.Entry(datos_frame, textvariable=self.descripcion_var, width=40).grid(row=0, column=1, pady=5, padx=5)
        
        ttk.Label(datos_frame, text="CONTENIDO:").grid(row=1, column=0, sticky="w", pady=5)
        self.cantidad_contenido_var = tk.IntVar(value=1)
        contenido_frame = ttk.Frame(datos_frame)
        contenido_frame.grid(row=1, column=1, sticky="w", pady=5, padx=5)
        ttk.Spinbox(contenido_frame, from_=1, to=1000, width=5, 
                   textvariable=self.cantidad_contenido_var).pack(side=tk.LEFT)
        ttk.Label(contenido_frame, text=" PIEZA(S)").pack(side=tk.LEFT)
        
        ttk.Label(datos_frame, text="HECHO EN:").grid(row=2, column=0, sticky="w", pady=5)
        self.hecho_en_var = tk.StringVar(value="CHINA")
        ttk.Entry(datos_frame, textvariable=self.hecho_en_var, width=40).grid(row=2, column=1, pady=5, padx=5)
        
        ttk.Label(datos_frame, text="No. PARTE:").grid(row=3, column=0, sticky="w", pady=5)
        self.numero_parte_var = tk.StringVar(value="12345-ABC")
        ttk.Entry(datos_frame, textvariable=self.numero_parte_var, width=40).grid(row=3, column=1, pady=5, padx=5)
        
        # Opciones de visualización
        opciones_frame = ttk.LabelFrame(main_frame, text="Opciones de Código de Barras", padding="10")
        opciones_frame.pack(fill=tk.X, pady=10)
        
        self.mostrar_codigo_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opciones_frame, text="Mostrar código de barras en la etiqueta", 
                       variable=self.mostrar_codigo_var, 
                       command=self.actualizar_vista_previa).pack()
        
        # Frame para vista previa
        preview_frame = ttk.LabelFrame(main_frame, text="Vista Previa", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Canvas para vista previa (escala 4x para mejor visualización)
        self.preview_canvas = tk.Canvas(preview_frame, bg="white", width=304, height=100)
        self.preview_canvas.pack()
        
        # Botones
        botones_frame = ttk.Frame(main_frame)
        botones_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(botones_frame, text="Actualizar Vista Previa", 
                  command=self.actualizar_vista_previa).pack(side=tk.LEFT, padx=5)
        ttk.Button(botones_frame, text="Generar PDF de Prueba", 
                  command=self.generar_pdf_prueba).pack(side=tk.LEFT, padx=5)
        ttk.Button(botones_frame, text="Ver Tamaño Real (PDF)", 
                  command=self.ver_tamano_real).pack(side=tk.LEFT, padx=5)
        
        # Actualizar vista previa inicial
        self.actualizar_vista_previa()
    
    def actualizar_vista_previa(self):
        """Actualiza la vista previa de la etiqueta"""
        # Limpiar canvas
        self.preview_canvas.delete("all")
        
        # Escala para vista previa (4x)
        scale = 4
        
        # Dibujar borde de la etiqueta
        self.preview_canvas.create_rectangle(2, 2, 302, 98, outline="black", width=2)
        
        # Obtener datos
        descripcion = self.descripcion_var.get().upper()
        cantidad_contenido = self.cantidad_contenido_var.get()
        contenido = f"{cantidad_contenido} PIEZA" if cantidad_contenido == 1 else f"{cantidad_contenido} PIEZAS"
        hecho_en = self.hecho_en_var.get().upper()
        numero_parte = self.numero_parte_var.get().upper()
        
        # Posiciones Y para el texto (escaladas)
        if self.mostrar_codigo_var.get() and numero_parte:
            # Con código de barras - texto más comprimido
            y_positions = [8, 16, 24, 32, 40, 48]
            font_size = 7
        else:
            # Sin código de barras - texto normal
            y_positions = [10, 20, 30, 40, 50, 60]
            font_size = 8
        
        # Dibujar textos
        textos = [
            ("IMPORTADOR: ", "MOTORMAN DE BAJA CALIFORNIA SA DE CV"),
            ("", "MARISCAL SUCRE 6738 LA CIENEGA PONIENTE,"),
            ("", "TIJUANA, B.C. 22114 RFC: MBC210723RP9"),
            ("DESCRIPCION: ", descripcion),
            ("CONTENIDO: ", contenido),
            ("HECHO EN: ", hecho_en)
        ]
        
        for i, (bold_text, normal_text) in enumerate(textos):
            y_pos = y_positions[i] * scale
            
            if bold_text:
                # Texto en negrita
                x_bold = 10
                self.preview_canvas.create_text(x_bold, y_pos, 
                                             text=bold_text, 
                                             font=("Arial", font_size, "bold"), 
                                             anchor="w")
                # Calcular ancho del texto en negrita
                x_normal = x_bold + len(bold_text) * 7
            else:
                x_normal = 10
            
            # Texto normal
            self.preview_canvas.create_text(x_normal, y_pos, 
                                         text=normal_text, 
                                         font=("Arial", font_size), 
                                         anchor="w")
        
        # Dibujar código de barras (simulado)
        if self.mostrar_codigo_var.get() and numero_parte:
            # Posición y tamaño del código de barras - MÁS PEQUEÑO
            barcode_y = 62 * scale  # Subido ligeramente (era 64)
            barcode_height = 22 * scale  # Ligeramente más alto para más separación
            barcode_width = 45 * scale   # Reducido
            barcode_x = 304 - barcode_width - 20  # Alineado a la derecha
            
            # Dibujar rectángulo de fondo
            self.preview_canvas.create_rectangle(barcode_x, barcode_y, 
                                               barcode_x + barcode_width, 
                                               barcode_y + barcode_height,
                                               fill="white", outline="")
            
            # Simular las barras (solo la parte superior)
            bar_height = 14 * scale
            bar_width = 1.5  # Barras más delgadas
            for i in range(0, int(barcode_width), int(bar_width * 2)):
                if i % 3 < 1.5:
                    self.preview_canvas.create_line(barcode_x + i, barcode_y + 2,
                                                  barcode_x + i, barcode_y + bar_height,
                                                  fill="black", width=bar_width)
            
            # Texto del número de parte debajo (más separado)
            text_x = barcode_x + barcode_width / 2
            self.preview_canvas.create_text(text_x, barcode_y + bar_height + 6,  # Más separación (era +4)
                                          text=numero_parte,
                                          font=("Arial", 6),
                                          anchor="center")
    
    def generar_pdf_prueba(self):
        """Genera un PDF de prueba con la etiqueta"""
        # Obtener datos
        descripcion = self.descripcion_var.get().upper()
        cantidad_contenido = self.cantidad_contenido_var.get()
        contenido = f"{cantidad_contenido} PIEZA" if cantidad_contenido == 1 else f"{cantidad_contenido} PIEZAS"
        hecho_en = self.hecho_en_var.get().upper()
        numero_parte = self.numero_parte_var.get().upper()
        
        # Crear archivo temporal
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_file.close()
        
        try:
            # Crear PDF
            c = canvas.Canvas(temp_file.name, pagesize=(76*mm, 25*mm))
            
            # Configurar fuentes
            font_size = 6
            font_size_small = 5
            
            # Posiciones Y para el texto
            if self.mostrar_codigo_var.get() and numero_parte:
                # Con código de barras
                y_positions = [22*mm, 19.5*mm, 17*mm, 14.5*mm, 12*mm, 9.5*mm]
            else:
                # Sin código de barras
                y_positions = [22*mm, 19.5*mm, 17*mm, 14*mm, 11*mm, 8*mm]
            
            # Dibujar textos
            # Línea 1: IMPORTADOR
            c.setFont("Helvetica-Bold", font_size)
            c.drawString(5*mm, y_positions[0], "IMPORTADOR: ")
            ancho_importador = c.stringWidth("IMPORTADOR: ", "Helvetica-Bold", font_size)
            c.setFont("Helvetica", font_size)
            c.drawString(5*mm + ancho_importador, y_positions[0], "MOTORMAN DE BAJA CALIFORNIA SA DE CV")
            
            # Líneas 2 y 3: Dirección
            c.setFont("Helvetica", font_size)
            c.drawString(5*mm, y_positions[1], "MARISCAL SUCRE 6738 LA CIENEGA PONIENTE,")
            c.drawString(5*mm, y_positions[2], "TIJUANA, B.C. 22114 RFC: MBC210723RP9")
            
            # Línea 4: DESCRIPCION
            c.setFont("Helvetica-Bold", font_size)
            c.drawString(5*mm, y_positions[3], "DESCRIPCION: ")
            ancho_desc = c.stringWidth("DESCRIPCION: ", "Helvetica-Bold", font_size)
            c.setFont("Helvetica", font_size)
            c.drawString(5*mm + ancho_desc, y_positions[3], descripcion)
            
            # Línea 5: CONTENIDO
            c.setFont("Helvetica-Bold", font_size)
            c.drawString(5*mm, y_positions[4], "CONTENIDO: ")
            ancho_cont = c.stringWidth("CONTENIDO: ", "Helvetica-Bold", font_size)
            c.setFont("Helvetica", font_size)
            c.drawString(5*mm + ancho_cont, y_positions[4], contenido)
            
            # Línea 6: HECHO EN
            c.setFont("Helvetica-Bold", font_size)
            c.drawString(5*mm, y_positions[5], "HECHO EN: ")
            ancho_hecho = c.stringWidth("HECHO EN: ", "Helvetica-Bold", font_size)
            c.setFont("Helvetica", font_size)
            c.drawString(5*mm + ancho_hecho, y_positions[5], hecho_en)
            
            # Código de barras
            if self.mostrar_codigo_var.get() and numero_parte:
                try:
                    # Generar código de barras
                    buffer = io.BytesIO()
                    code = Code128(numero_parte, writer=ImageWriter())
                    
                    # Opciones para código de barras más compacto
                    options = {
                        'module_width': 0.15,      # Barras más delgadas
                        'module_height': 3.5,      # Altura reducida de barras
                        'font_size': 6,            # Texto más pequeño
                        'text_distance': 2.5,      # Un poco más de espacio entre barras y texto
                        'quiet_zone': 1,           # Márgenes mínimos
                        'write_text': True         # Mostrar texto
                    }
                    
                    code.write(buffer, options=options)
                    buffer.seek(0)
                    
                    # Cargar imagen
                    barcode_image = Image.open(buffer)
                    
                    # Posición y tamaño - MÁS PEQUEÑO
                    barcode_width = 35*mm      # Reducido de 55mm a 35mm
                    barcode_height = 8*mm       # Reducido de 12mm a 8mm
                    barcode_x = 76*mm - barcode_width - 5*mm  # Alineado a la derecha
                    barcode_y = 2*mm            # Subido ligeramente de 1mm a 2mm
                    
                    # Dibujar imagen
                    c.drawInlineImage(barcode_image, barcode_x, barcode_y, 
                                    width=barcode_width, height=barcode_height)
                
                except Exception as e:
                    print(f"Error con código de barras: {e}")
            
            # Guardar PDF
            c.save()
            
            # Abrir PDF
            if os.name == 'nt':  # Windows
                os.startfile(temp_file.name)
            else:
                os.system(f"open '{temp_file.name}'" if os.name == 'posix' else f"xdg-open '{temp_file.name}'")
            
            messagebox.showinfo("Éxito", f"PDF generado: {temp_file.name}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar PDF: {str(e)}")
    
    def ver_tamano_real(self):
        """Genera un PDF con varias etiquetas para ver el tamaño real"""
        archivo = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            title="Guardar PDF con etiquetas"
        )
        
        if not archivo:
            return
        
        try:
            # Crear PDF tamaño carta
            from reportlab.lib.pagesizes import letter
            c = canvas.Canvas(archivo, pagesize=letter)
            
            # Información para dibujar
            descripcion = self.descripcion_var.get().upper()
            cantidad_contenido = self.cantidad_contenido_var.get()
            contenido = f"{cantidad_contenido} PIEZA" if cantidad_contenido == 1 else f"{cantidad_contenido} PIEZAS"
            hecho_en = self.hecho_en_var.get().upper()
            numero_parte = self.numero_parte_var.get().upper()
            
            # Dibujar múltiples etiquetas
            x_offset = 20*mm
            y_offset = 250*mm
            
            for fila in range(8):  # 8 filas
                for col in range(2):  # 2 columnas
                    x = x_offset + col * 85*mm
                    y = y_offset - fila * 30*mm
                    
                    # Dibujar rectángulo de la etiqueta
                    c.rect(x, y, 76*mm, 25*mm)
                    
                    # Dibujar contenido (simplificado)
                    c.setFont("Helvetica", 5)
                    c.drawString(x + 2*mm, y + 20*mm, "IMPORTADOR: MOTORMAN DE BAJA CALIFORNIA SA DE CV")
                    c.drawString(x + 2*mm, y + 17*mm, "DESCRIPCION: " + descripcion[:30])
                    c.drawString(x + 2*mm, y + 14*mm, "CONTENIDO: " + contenido)
                    c.drawString(x + 2*mm, y + 11*mm, "HECHO EN: " + hecho_en)
                    c.drawString(x + 2*mm, y + 8*mm, "No. PARTE: " + numero_parte)
                    
                    # Línea para código de barras
                    if self.mostrar_codigo_var.get():
                        c.drawString(x + 38*mm, y + 2*mm, "|||||||||||||||")
            
            c.save()
            
            messagebox.showinfo("Éxito", f"PDF guardado: {archivo}")
            
            # Abrir PDF
            if os.name == 'nt':
                os.startfile(archivo)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PruebaEtiquetaConCodigoBarras(root)
    root.mainloop()