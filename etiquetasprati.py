import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from barcode import EAN13
from barcode.writer import ImageWriter
import os
import tempfile

# === CONFIGURACIÓN ===
ruta_Excel = r"C:\Users\Umco\OneDrive\Desktop\ETIQUETAS DEPRATI\de\datos_imprimir_etiquetas.xlsx"
Ruta_logo = r"C:\Users\Umco\OneDrive\Desktop\ETIQUETAS DEPRATI\de\logo_bn.bmp"
rutaSalida = r"C:\Users\Umco\OneDrive\Desktop\ETIQUETAS DEPRATI\de\etiquetas_final.pdf"

dpi = 300
width_cm, height_cm = 3, 6
width_px = int((width_cm / 2.54) * dpi)
height_px = int((height_cm / 2.54) * dpi)

# === FUENTES ===
font_path = "arial.ttf"
font = ImageFont.truetype(font_path, 22)
font_small = ImageFont.truetype(font_path, 18)
font_big = ImageFont.truetype(font_path, 26)
ean_font = ImageFont.truetype(font_path, 18)

# === CARGAR DATOS (forzando Cod3, Cod4, Cod6 como texto) ===
df = pd.read_excel(ruta_Excel, dtype={'Cod3': str, 'Cod4': str, 'Cod6': str})
etiquetas = []

for _, row in df.iterrows():
    cantidad = int(row['Cantidad'])

    for _ in range(cantidad):  # Repetir según cantidad
        img = Image.new("RGB", (width_px, height_px), "white")
        draw = ImageDraw.Draw(img)

        # === LOGO ===
        if os.path.exists(Ruta_logo):
            logo = Image.open(Ruta_logo).resize((int(width_px * 0.9), 80))
            img.paste(logo, ((width_px - logo.width) // 2, 5))

        y = 95
        draw.text((11, y), f"    {row['Cod1']}                 {row['Cod2']}", font=font, fill="black")
        y += 30
        cod3 = str(row['Cod3']).zfill(5)
        draw.text((11, y), f"                                     {cod3}", font=font, fill="black")
        y += 30
        cod4 = str(row['Cod4']).zfill(4)
        cod6 = str(row['Cod6']).zfill(5)
        draw.text((11, y), f"{cod4}                {cod6}     {row['Cod5']}", font=font, fill="black")
        y += 35

        # === CÓDIGO DE BARRAS VERTICAL ===
        ean_code = ''.join(filter(str.isdigit, str(row['EAN']).strip())).zfill(12)[:12]
        img_barra = EAN13(ean_code, writer=ImageWriter()).render(writer_options={"module_height": 30, "font_size": 10})
        img_barra = img_barra.crop((0, 0, img_barra.width, img_barra.height - 12))
        img_barra = img_barra.resize((250, 95))

        # Rotar código de barras
        img_barra = img_barra.rotate(90, expand=True)

        # EAN como texto a la derecha, uno por línea
        ean_text_img = Image.new("RGB", (20, img_barra.height), "white")


        # Combinar barra + texto
        combinado_barras = Image.new("RGB", (img_barra.width + ean_text_img.width, img_barra.height), "white")
        combinado_barras.paste(img_barra, (0, 0))
        combinado_barras.paste(ean_text_img, (img_barra.width, 0))

        barcode_x = (width_px - combinado_barras.width) // 2
        img.paste(combinado_barras, (barcode_x, y))
        y += combinado_barras.height + 10

        # === PRECIOS ===
        draw.text((11, y), "SUBTOTAL", font=font_small, fill="black")
        draw.text((width_px - 100, y), f"${str(row['Subtotal']).replace('$','')}", font=font_small, fill="black")
        y += 22
        draw.text((11, y), "IVA 15%", font=font_small, fill="black")
        draw.text((width_px - 100, y), f"${str(row['IVA']).replace('$','')}", font=font_small, fill="black")
        y += 22
        draw.text((11, y), "PRECIO FINAL", font=font_small, fill="black")
        draw.text((width_px - 110, y), f"${str(row['Total']).replace('$','')}", font=font_big, fill="black")

        # Rotar toda la etiqueta
        img_rotada = img.rotate(90, expand=True)
        etiquetas.append(img_rotada)

# === CREAR PDF (una etiqueta por hoja 3x6 cm) ===
from reportlab.lib.pagesizes import landscape
page_width = height_cm * cm
page_height = width_cm * cm

c = canvas.Canvas(rutaSalida, pagesize=(page_width, page_height))

for etiqueta in etiquetas:
    temp_img_path = tempfile.mktemp(suffix=".png")
    etiqueta.save(temp_img_path)
    c.drawImage(temp_img_path, 0, 0, width=page_width, height=page_height)
    c.showPage()

c.save()
print("Listo")
