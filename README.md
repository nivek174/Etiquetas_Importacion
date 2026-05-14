# Etiquetas de Importacion
> Generador de etiquetas de 76x25mm y 76x51mm para mercancia de importacion con codigo de barras

## Descripcion

Aplicacion para generar etiquetas de importacion con informacion del importador, descripcion del producto, contenido, pais de origen y codigo de barras Code128. Disponible en version web (Streamlit) y escritorio (Tkinter). Genera etiquetas en formato PDF listas para impresora de etiquetas y en Excel para impresion en hoja.

Cumple con los requisitos de etiquetado de aduanas mexicanas para Motorman de Baja California.

## Funcionalidades

- **Version web (Streamlit)**: Interfaz moderna en navegador con vista previa en tiempo real
- **Version escritorio (Tkinter)**: Aplicacion nativa de Windows con previsualizacion
- **Dos tamanios**: Etiquetas de 76x25mm y 76x51mm
- **Ingreso manual**: Formulario para etiquetas individuales
- **Importacion masiva**: Carga desde Excel con seleccion multiple de productos
- **Codigo de barras**: Code128 generado dinamicamente con python-barcode
- **Exportacion PDF**: Control preciso de posicionamiento con ReportLab
- **Exportacion Excel**: 3 etiquetas por fila con formato de impresion
- **Plantilla Excel**: Descarga de plantilla con formato esperado
- **Deteccion flexible**: Reconoce multiples nombres de columna (sku, codigo, numero_parte, etc.)

## Tecnologias

- **Python 3.12**
- **Streamlit** — Interfaz web
- **Tkinter** — Interfaz de escritorio
- **ReportLab** — Generacion de PDF
- **python-barcode + Pillow** — Codigos de barras Code128
- **Pandas + openpyxl** — Lectura/escritura Excel

## Instalacion

```bash
pip install -r requirements.txt
```

## Uso

### Version Streamlit (recomendada)

```bash
streamlit run appEtiquetas.py
```

### Version Tkinter

```bash
python Etiquetas_Imp_76X25.py                    # Etiquetas 76x25mm
python generador_Etiquetas_Importacion76x51       # Etiquetas 76x51mm
```

### Formato Excel de entrada

| descripcion | cantidad_contenido | hecho_en | numero_parte | cantidad_etiquetas |
|---|---|---|---|---|
| ABANICO PARA RADIADOR | 1 | CHINA | FAN-12345 | 10 |
| BOMBA DE AGUA | 2 | JAPON | WP-67890 | 5 |

## Estructura del Proyecto

```
ETIQUETAS_IMPORTACION/
├── appEtiquetas.py                            # App Streamlit (recomendada)
├── Etiquetas_Imp_76X25.py                     # App Tkinter 76x25mm
├── generador_Etiquetas_Importacion76x51       # App Tkinter 76x51mm
├── requirements.txt                           # Dependencias
└── .devcontainer/                             # Config VS Code Dev Container
```

## Autor

Kevin Salazar — Motorman de Baja California
