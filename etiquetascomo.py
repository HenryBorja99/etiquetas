def main ():
    import pandas as pd
    from PIL import Image, ImageDraw, ImageFont
    from barcode import EAN13
    from barcode.writer import ImageWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm
    from reportlab.lib.utils import ImageReader
    import os
    import re

    # CONFIGURACIÓN DE PLANTILLA
    label_w_cm = 3.2
    label_h_cm = 2.5
    left_cm, right_cm = 0.1, 0.1
    top_cm, bottom_cm = 0.1, 0.1
    space_h_cm, space_v_cm = 0.3, 0.0
    cols, rows = 3, 1

    page_w_cm = left_cm + cols * label_w_cm + (cols - 1) * space_h_cm + right_cm
    page_h_cm = top_cm + rows * label_h_cm + (rows - 1) * space_v_cm + bottom_cm

    excel_path = r"C:\Users\Umco\OneDrive\Desktop\ETIQUETAS DEPRATI\COMOHOGAR\Copia de UMCO S.A.xlsx"
    output_pdf = r"C:\Users\Umco\OneDrive\Desktop\ETIQUETAS DEPRATI\COMOHOGAR\etiquetas_generadas.pdf"

    # FUENTES
    font_path = "arial.ttf"
    font_pvp = ImageFont.truetype(font_path, 23)
    font_afiliado = ImageFont.truetype(font_path, 26)
    font_small = ImageFont.truetype(font_path, 14)
    font_tiny = ImageFont.truetype(font_path, 13)
    font_tiny2 = ImageFont.truetype(font_path, 18)
    font_bold = ImageFont.truetype(font_path, 20)
    font_small2 = ImageFont.truetype(font_path, 20)
    # CARGAR DATOS
    def load_data(file_path):
        try:
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names
            df = None
            for sheet in sheet_names:
                temp_df = pd.read_excel(file_path, sheet_name=sheet)
                temp_df.columns = [col.strip() for col in temp_df.columns]
                if 'CANT2' in temp_df.columns:
                    df = temp_df[temp_df["CANT2"] > 0]
                    if df.empty:
                        temp_df['CANT2'] = 1
                        df = temp_df
                    break
            if df is None:
                df = pd.read_excel(file_path, sheet_name=sheet_names[0])
                df.columns = [col.strip() for col in df.columns]
                df['CANT2'] = 1
            return df
        except Exception as e:
            print(f"Error al cargar Excel: {e}")
            return None

    df = load_data(excel_path)
    if df is None:
        raise SystemExit("No se pudo cargar los datos del Excel.")

    required_columns = ["COD", "DESCRIPCION", "REFERENCIAS", "BARRAS", "PVP", "AFILIADO CONTADO", "CANT2"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise SystemExit(f"Faltan columnas: {', '.join(missing_columns)}")

    # FUNCIÓN PARA CREAR UNA ETIQUETA
    def draw_label(row):
        codigo = str(row["COD"]).strip()
        descripcion = str(row["DESCRIPCION"]).strip()
        referencias = str(row["REFERENCIAS"]).strip()
        barras_1 = str(row["BARRAS"]).strip()   # aquí ya vienen los 13 dígitos completos
        pvp = float(row["PVP"])
        afiliado = float(row["AFILIADO CONTADO"])

        pvp_str = f"${pvp:,.2f}"
        afiliado_str = f"${afiliado:,.2f}"

        dpi = 204
        label_w_px = int((label_w_cm / 2.54) * dpi)
        label_h_px = int((label_h_cm / 2.54) * dpi)

        img = Image.new("RGB", (label_w_px, label_h_px), "white")
        draw = ImageDraw.Draw(img)

        # Encabezado
        draw.text((10, 10), codigo, font=font_small, fill="black")
        draw.text((label_w_px - 120, 10), "#", font=font_small, fill="black")
        draw.text((label_w_px - 85, 10), referencias, font=font_small, fill="black")

        # Descripción en una o dos líneas
        max_desc_width = label_w_px - 20
        words = descripcion.split()
        line1, line2, temp_line = "", "", ""

        for word in words:
            test_line = temp_line + " " + word if temp_line else word
            if draw.textbbox((0, 0), test_line, font=font_small)[2] <= max_desc_width:
                temp_line = test_line
            else:
                line1 = temp_line.strip()
                line2 = " ".join(words[words.index(word):])
                break

        if not line1 and draw.textbbox((0, 0), descripcion, font=font_small)[2] <= max_desc_width:
            line1 = descripcion
            desc_y = 30
            draw.text((10, desc_y), line1, font=font_small, fill="black")
        else:
            if not line1:
                line1 = temp_line.strip()
            draw.text((10, 30), line1, font=font_small, fill="black")
            draw.text((10, 45), line2, font=font_small, fill="black")
            desc_y = 45

        # Precios
        draw.text((18, desc_y + 19), "P.V.P.", font=font_tiny2, fill="black")
        draw.text((15, desc_y + 40), pvp_str, font=font_pvp, fill="black")
        draw.text((118, desc_y + 19), "AFILIADO CONTADO", font=font_tiny, fill="black")
        draw.text((140, desc_y + 38), afiliado_str, font=font_afiliado, fill="black")

        # Código de barras EAN-13
        try:
            ean = EAN13(barras_1, writer=ImageWriter())
            barcode_img = ean.render(writer_options={
                "dpi": 203,
                "module_width": 0.25,
                "module_height": 7,
                "quiet_zone": 3,
                "font_size": 12,
                "text_distance": 1,
                "write_text": False
            })
            bx = (label_w_px - barcode_img.width) // 2
            by = 110
            img.paste(barcode_img, (bx, by))

            # Formatear el texto en bloques estilo EAN-13
            texto_formateado = f"{barras_1[0]}    {barras_1[1:7]}    {barras_1[7:13]}"

            # Fuente 
            font_ean = font_small2

            # Medir y centrar
            text_bbox = draw.textbbox((0, 0), texto_formateado, font=font_ean)
            text_width = text_bbox[2] - text_bbox[0]
            text_x = bx + (barcode_img.width - text_width) // 2
            text_y = by + barcode_img.height -8

            # Dibujar el texto formateado
            draw.text((text_x, text_y), texto_formateado, font=font_ean, fill="black")

        except Exception as e:
            draw.text((10, 100), "ERROR EAN13", font=font_small, fill="red")

        return img

    # FUNCIÓN PARA CREAR UNA ETIQUETA SEPARADORA
    def draw_separador(texto):
        dpi = 204
        label_w_px = int((label_w_cm / 2.54) * dpi)
        label_h_px = int((label_h_cm / 2.54) * dpi)
        img = Image.new("RGB", (label_w_px, label_h_px), "white")
        draw = ImageDraw.Draw(img)
        font = ImageFont.truetype(font_path, 40)
        bbox = draw.textbbox((0, 0), texto, font=font)
        x_text = (label_w_px - (bbox[2] - bbox[0])) // 2
        y_text = (label_h_px - (bbox[3] - bbox[1])) // 2
        draw.text((x_text, y_text), texto, font=font, fill="black")
        return img

    # CREAR PD
    page_w, page_h = page_w_cm * cm, page_h_cm * cm
    os.makedirs(os.path.dirname(output_pdf), exist_ok=True)
    c = canvas.Canvas(output_pdf, pagesize=(page_w, page_h))

    slot = 0
    for _, row in df.iterrows():
        cantidad = int(row["CANT2"])

        # Imprimir etiquetas del producto
        for _ in range(cantidad):
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

        # Insertar etiqueta separadora
        referencia_raw = str(row["REFERENCIAS"]).strip()
        referencia_formateada = referencia_raw.zfill(4)
        col = slot % cols
        row_idx = slot // cols
        x = (left_cm + col * (label_w_cm + space_h_cm)) * cm
        y = page_h - (top_cm * cm) - (row_idx + 1) * label_h_cm * cm - row_idx * space_v_cm * cm

        c.setFont("Helvetica", 17)
        c.drawCentredString(x + (label_w_cm * cm) / 2, y + (label_h_cm * cm) / 2 - 10, f"REF: {referencia_formateada}")

        slot += 1
        if slot == cols * rows:
            c.showPage()
            slot = 0

        # Forzar nueva página para el siguiente producto
        if slot > 0:
            c.showPage()
            slot = 0

    c.save()
    print(f"\nPDF generado correctamente: {output_pdf}")
main()
