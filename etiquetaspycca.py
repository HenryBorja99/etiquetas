def main ():
    import pandas as pd
    from PIL import Image, ImageDraw, ImageFont
    from barcode.codex import Code128
    from barcode.writer import ImageWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm
    from reportlab.lib.utils import ImageReader
    import os

    # CONFIGURACIÓN DE PLANTILLA 
    label_w_cm = 3.2
    label_h_cm = 2.5
    left_cm, right_cm = 0.1, 0.1
    top_cm, bottom_cm = 0.1, 0.1
    space_h_cm, space_v_cm = 0.3, 0.0
    cols, rows = 3, 1

    page_w_cm = left_cm + cols * label_w_cm + (cols - 1) * space_h_cm + right_cm
    page_h_cm = top_cm + rows * label_h_cm + (rows - 1) * space_v_cm + bottom_cm

    excel_path = r"C:\Users\Umco\OneDrive\Desktop\ETIQUETAS DEPRATI\PYCCA\etiquetas_pycca.xlsx"
    logo_path  = r"C:\Users\Umco\OneDrive\Desktop\ETIQUETAS DEPRATI\PYCCA\Logo_pycca.bmp"
    output_pdf = r"C:\Users\Umco\OneDrive\Desktop\ETIQUETAS DEPRATI\PYCCA\etiquetas_generadas.pdf"

    # FUENTES
    font_path = "arial.ttf"
    font_big    = ImageFont.truetype(font_path, 45)
    font_decimal= ImageFont.truetype(font_path, 30)
    font_dollar = ImageFont.truetype(font_path, 38)
    font_superior = ImageFont.truetype(font_path, 22)
    font_small  = ImageFont.truetype(font_path, 16)
    font_tiny   = ImageFont.truetype(font_path, 12)

    # CARGAR DATOS
    df = pd.read_excel(excel_path, sheet_name="Precios")
    df = df[df["Cantidad"] > 0]
    if df.empty:
        raise SystemExit("No hay productos con cantidad > 0 en el Excel.")

    # FUNCIÓN PARA CREAR UNA ETIQUETA
    def draw_label(row):
        codigo = str(row["Código"]).strip()
        descripcion = str(row["Descripción"]).strip()
        precio_entero = str(row["Precio_Entero"]).strip()
        precio_decimal = f"{int(row['Precio_Decimal']):02d}"  

        dpi = 204
        label_w_px = int((label_w_cm / 2.54) * dpi)
        label_h_px = int((label_h_cm / 2.54) * dpi)

        img = Image.new("RGB", (label_w_px, label_h_px), "white")
        draw = ImageDraw.Draw(img)

        if os.path.exists(logo_path):
            try:
                logo = Image.open(logo_path).convert("RGBA")
                target_h = 40
                ratio = target_h / logo.height
                logo = logo.resize((int(logo.width * ratio), target_h))
                img.paste(logo, (10, 5), logo)
            except Exception:
                pass

        code_w = draw.textbbox((0, 0), codigo, font=font_small)[2]
        draw.text((label_w_px - 28 - code_w, 8), codigo, font=font_superior, fill="black")

        pe_w = draw.textbbox((0, 0), precio_entero, font=font_big)[2]
        pd_w = draw.textbbox((0, 0), precio_decimal, font=font_decimal)[2]
        d_w  = draw.textbbox((0, 0), "$", font=font_dollar)[2]
        total_w = pe_w + pd_w + d_w
        x0 = (label_w_px - total_w) // 2
        y_price = 45
        draw.text((x0, y_price + 6), "$", font=font_dollar, fill="black")
        draw.text((x0 + d_w, y_price), precio_entero, font=font_big, fill="black")
        draw.text((x0 + d_w + pe_w, y_price + 2), precio_decimal, font=font_decimal, fill="black")

        max_desc_width = label_w_px - 20
        desc = descripcion
        while draw.textbbox((0, 0), desc, font=font_small)[2] > max_desc_width and len(desc) > 3:
            desc = desc[:-1]
        if desc != descripcion:
            desc = desc[:-1]
        desc_w = draw.textbbox((0, 0), desc, font=font_small)[2]
        y_desc = y_price + 42
        draw.text(((label_w_px - desc_w) // 2, y_desc), desc, font=font_small, fill="black")
    #Iva
        iva_text = "Incluido IVA"
        iva_w = draw.textbbox((0, 0), iva_text, font=font_tiny)[2]
        y_iva = y_desc + 22
        draw.text(((label_w_px - iva_w) // 2, y_iva), iva_text, font=font_tiny, fill="black")

    # Barras
        barcode_img = Code128(codigo, writer=ImageWriter()).render(
            writer_options={
                "dpi": 203,
                "module_width": 0.25,
                "module_height": 8,
                "quiet_zone": 3,
                "write_text": False
            }
        )
        bx = (label_w_px - barcode_img.width) // 2
        by = y_iva + 12
        img.paste(barcode_img, (bx, by))

        return img
    # FUNCIÓN PARA CREAR UNA ETIQUETA SEPARADORA
    def draw_separador(texto):
        dpi = 204
        label_w_px = int((label_w_cm / 2.54) * dpi)
        label_h_px = int((label_h_cm / 2.54) * dpi)
        img = Image.new("RGB", (label_w_px, label_h_px), "white")
        draw = ImageDraw.Draw(img)

        try:
            font = ImageFont.truetype(font_path, 40)
        except:
            font = ImageFont.load_default()

        bbox = draw.textbbox((0, 0), texto, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        x_text = (label_w_px - text_w) // 2
        y_text = (label_h_px - text_h) // 2
        draw.text((x_text, y_text), texto, font=font, fill="black")
        return img

    # CREAR PDF CON DISTRIBUCIÓN 3x1 (10.5 × 2.7 cm)
    page_w, page_h = page_w_cm * cm, page_h_cm * cm
    os.makedirs(os.path.dirname(output_pdf), exist_ok=True)
    c = canvas.Canvas(output_pdf, pagesize=(page_w, page_h))

    slot = 0
    for _, row in df.iterrows():
        for _ in range(int(row["Cantidad"])):
            im = draw_label(row)
            ir = ImageReader(im)

            col = slot % cols
            row_idx = slot // cols

            x = (left_cm + col * (label_w_cm + space_h_cm)) * cm
            y = page_h - (top_cm * cm) - (row_idx + 1) * label_h_cm * cm - row_idx * space_v_cm * cm

            c.drawImage(ir, x, y, width=label_w_cm * cm, height=label_h_cm * cm)

            slot += 1
            if slot == cols * rows:
                c.showPage()
                slot = 0
        # Insertar etiqueta separadora con el valor de la columna "Original"
        original_text = str(row["Original"]).strip()
        im_sep = draw_separador(original_text)
        ir_sep = original_text.zfill(4)

        col = slot % cols
        row_idx = slot // cols
        x = (left_cm + col * (label_w_cm + space_h_cm)) * cm
        y = page_h - (top_cm * cm) - (row_idx + 1) * label_h_cm * cm - row_idx * space_v_cm * cm

        c.setFont("Helvetica",12)
        c.drawCentredString(x + (label_w_cm * cm) / 2, y + (label_h_cm * cm) / 2 - 10, f"{ir_sep}")

        slot += 1
        if slot == cols * rows:
            c.showPage()
            slot = 0

        if slot > 0:
            c.showPage()
            slot = 0

    c.save()
    print(f"\n PDF generado correctamente: {output_pdf}")
main()
