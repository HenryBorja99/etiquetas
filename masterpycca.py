def main ():
    import pandas as pd
    from PIL import Image, ImageDraw, ImageFont
    from barcode.ean import EuropeanArticleNumber13
    from barcode.writer import ImageWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm
    from reportlab.lib.utils import ImageReader
    import os
    from textwrap import wrap

    # ================================
    # CONFIGURACIÓN GENERAL
    # ================================
    label_w_cm = 9.1
    label_h_cm = 5.1
    left_cm, right_cm = 0.1, 0.1
    top_cm, bottom_cm = 0.1, 0.1
    space_h_cm, space_v_cm = 0.3, 0.0
    cols, rows = 1, 1

    page_w_cm = left_cm + cols * label_w_cm + (cols - 1) * space_h_cm + right_cm
    page_h_cm = top_cm + rows * label_h_cm + (rows - 1) * space_v_cm + bottom_cm

    #Rutas
    excel_path = r"C:\Users\Umco\OneDrive\Desktop\ETIQUETAS DEPRATI\PYCCA\etiquetas_pycca.xlsx"
    logo_path  = r"C:\Users\Umco\OneDrive\Desktop\ETIQUETAS DEPRATI\PYCCA\Logo_pycca.bmp"
    output_pdf = r"C:\Users\Umco\OneDrive\Desktop\ETIQUETAS DEPRATI\PYCCA\master_generadas.pdf"

    # ================================
    # FUENTES
    # ================================
    # Asegúrate de que esta fuente esté disponible en tu sistema
    font_path = "arial.ttf"
    font_separador   = ImageFont.truetype(font_path, 85)
    font_medium  = ImageFont.truetype(font_path, 23)
    font_small   = ImageFont.truetype(font_path, 19)
    font_big_pycca     = ImageFont.truetype(font_path, 100)
    font_cantidad = ImageFont.truetype(font_path, 27)
    font_umco = ImageFont.truetype(font_path, 27)

    # ================================
    # CARGAR DATOS Y VALIDAR
    # ================================
    df = pd.read_excel(excel_path, sheet_name="master")
    df.columns = df.columns.str.strip()
    required_cols = [
        "CODIGO FABRICA:", "DESCRIPCION:",
        "CODIGO BARRA/EAN", "CODIGO PYCCA:", "Cantidad"
    ]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise SystemExit(f"Faltan columnas en el Excel: {missing}")
    df = df[df["Cantidad"] > 0]
    if df.empty:
        raise SystemExit("No hay productos con cantidad > 0 en el Excel.")

    # ================================
    # FUNCIÓN PARA DIBUJAR ETIQUETA
    # ================================
    def draw_label(row):
        dpi = 204
        label_w_px = int((label_w_cm / 2.54) * dpi)
        label_h_px = int((label_h_cm / 2.54) * dpi)
        img = Image.new("RGB", (label_w_px, label_h_px), "white")
        draw = ImageDraw.Draw(img)

        # Coordenadas y dimensiones
        margin = 10
        thick_line = 4
        
        # Marco exterior
        draw.rectangle([0, 0, label_w_px - 1, label_h_px - 1], outline="black", width=thick_line)
        
        # Líneas divisorias horizontales
        line1_y = 60
        line2_y = 260
        draw.line([(0, line1_y), (label_w_px, line1_y)], fill="black", width=thick_line)
        draw.line([(0, line2_y), (label_w_px, line2_y)], fill="black", width=thick_line)
        
        # Líneas divisorias verticales
        draw.line([(250, 0), (250, line1_y)], fill="black", width=thick_line)

        # Línea vertical principal
        main_vertical_line_x = 450
        draw.line([(main_vertical_line_x, 0), (main_vertical_line_x, label_h_px)], fill="black", width=thick_line)

        # Líneas horizontales internas en la sección derecha
        line3_y = line1_y + 50
        line4_y = line3_y + 90
        draw.line([(main_vertical_line_x, line3_y), (label_w_px, line3_y)], fill="black", width=thick_line)
        draw.line([(main_vertical_line_x, line4_y), (label_w_px, line4_y)], fill="black", width=thick_line)
        
        # Datos de la fila
        orden_compra = orden_compra_manual
        fabrica = str(row["CODIGO FABRICA:"]).strip()
        descripcion = str(row["DESCRIPCION:"]).strip()
        ean = str(row["CODIGO BARRA/EAN"]).strip()
        pycca = str(row["CODIGO PYCCA:"]).strip()
        partes = "1"
        unidades = str(row["Unidades"]).strip()
        cantidad = str(row["Cantidad"]).strip()

        # Validaciones
        if len(ean) != 13 or not ean.isdigit():
            raise ValueError(f"Código EAN inválido: {ean}")

        # Contenido de la etiqueta
        # Seccion superior izquierda: ORDEN DE COMPRA
        draw.text((margin, 8), "ORDEN DE COMPRA: ", font=font_small, fill="black")
        bbox_oc_num = draw.textbbox((0, 0), orden_compra, font=font_medium)
        text_oc_num_w = bbox_oc_num[2] - bbox_oc_num[0]
        draw.text((margin + (250 - 2 * margin - text_oc_num_w) // 2, 16 + (draw.textbbox((0, 0), "ORDEN DE COMPRA: ", font=font_small)[3] - draw.textbbox((0, 0), "ORDEN DE COMPRA: ", font=font_small)[1])), orden_compra, font=font_medium, fill="black")

        # Seccion superior central: UMCO S.A.
        bbox_umco = draw.textbbox((0, 0), "UMCO S. A.", font=font_umco)
        text_umco_w = bbox_umco[2] - bbox_umco[0]
        draw.text((250 + (main_vertical_line_x - 250 - text_umco_w) // 2, 10), "UMCO S. A.", font=font_umco, fill="black")

        # Seccion superior derecha: LOGO PYCCA
        if os.path.exists(logo_path):
            try:
                logo = Image.open(logo_path).convert("RGB")
                logo_w, logo_h = logo.size
                max_h = 50
                if logo_h > max_h:
                    scale = max_h / logo_h
                    logo = logo.resize((int(logo_w * scale), max_h))
                
                logo_w, logo_h = logo.size
                lx = main_vertical_line_x + (label_w_px - main_vertical_line_x - logo_w) // 2 - 15
                ly = (line1_y - logo_h) // 2
                img.paste(logo, (lx, ly))
            except Exception as e:
                print(f"Error al cargar o pegar el logo: {e}")

        #CODIGO PYCCA
        draw.text((margin, line1_y + 10), "CODIGO PYCCA:", font=font_small, fill="black")
        bbox_pycca = draw.textbbox((0, 0), pycca, font=font_big_pycca)
        text_pycca_w = bbox_pycca[2] - bbox_pycca[0]
        draw.text((margin + (350 -1 * margin - text_pycca_w) // 2, line1_y + 45), pycca, font=font_big_pycca, fill="black")

        # CODIGO FABRICA
        draw.text((main_vertical_line_x + 5, line1_y + 5), "CODIGO FABRICA:", font=font_small, fill="black")
        draw.text((main_vertical_line_x + 5, line1_y + 22), fabrica, font=font_medium, fill="black")
        
        #  DESCRIPCION
        draw.text((main_vertical_line_x + 5, line3_y + 5), "DESCRIPCION:", font=font_small, fill="black")
        lines = wrap(descripcion, width=17)
        y_text = line3_y + 25
        for line in lines:
            draw.text((main_vertical_line_x + 5, y_text), line, font=font_small, fill="black")
            y_text += 20
        
        #PARTES
        draw.text((main_vertical_line_x + 5, line4_y + 5), "PARTES:", font=font_small, fill="black")
        draw.text((main_vertical_line_x + 5, line4_y + 25), partes, font=font_medium, fill="black")

        # CODIGO DE BARRAS
        draw.text((margin, line2_y + 5), "CODIGO BARRA/EAN:", font=font_small, fill="black")
        barcode_img = EuropeanArticleNumber13(ean, writer=ImageWriter()).render({
            "dpi": 204,
            "module_width": 0.25,
            "module_height": 10,
            "quiet_zone": 3,
            "write_text": False,
        #  "font_size": 10,
        #  "text_distance": 6,
        # "font_path": font_path

        })
        bx = margin + (main_vertical_line_x - 2 * margin - barcode_img.width) // 2
        by = label_h_px - barcode_img.height - 20
        img.paste(barcode_img, (bx, by))
        
        # CANTIDAD
        draw.text((main_vertical_line_x + 5, line2_y + 5), "CANTIDAD:", font=font_small, fill="black")
        
        bbox_cant = draw.textbbox((0, 0), unidades, font=font_cantidad)
        text_cant_w = bbox_cant[2] - bbox_cant[0]
        x_cant = main_vertical_line_x + ((label_w_px - main_vertical_line_x - text_cant_w) // 2)
        draw.text((x_cant, line2_y + 25), unidades, font=font_cantidad, fill="black")

        bbox_unid = draw.textbbox((0, 0), "Unidades", font=font_cantidad)
        text_unid_w = bbox_unid[2] - bbox_unid[0]
        x_unid = main_vertical_line_x + ((label_w_px - main_vertical_line_x - text_unid_w) // 2)
        draw.text((x_unid, line2_y + 45), "Unidades", font=font_cantidad, fill="black")
        
        return img  

    # ================================
    # FUNCIÓN PARA SEPARADOR
    # ================================
    def draw_separador(texto):
        dpi = 204
        label_w_px = int((label_w_cm / 2.54) * dpi)
        label_h_px = int((label_h_cm / 2.54) * dpi)
        img = Image.new("RGB", (label_w_px, label_h_px), "white")
        draw = ImageDraw.Draw(img)
        draw.rectangle([0, label_h_px // 2 - 10, label_w_px, label_h_px // 2 + 10], fill="white")
        bbox = draw.textbbox((0, 0), texto, font=font_separador)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        x_text = (label_w_px - text_w) // 2
        y_text = (label_h_px - text_h) // 2
        draw.text((x_text, y_text), texto, font=font_separador, fill="black")
        return img

    # ================================
    # GENERAR PDF
    # ================================
    page_w, page_h = page_w_cm * cm, page_h_cm * cm
    os.makedirs(os.path.dirname(output_pdf) or '.', exist_ok=True)
    c = canvas.Canvas(output_pdf, pagesize=(page_w, page_h))

    orden_compra_manual = str(input("Digite el número de orden de compra:")).strip()

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
    #separador
        texto_sep = f"CODIGO {str(row['CODIGO FABRICA:']).strip()}"
        texto_formateado = texto_sep.zfill(4)

        col = slot % cols
        row_idx = slot // cols
        x = (left_cm + col * (label_w_cm + space_h_cm)) * cm
        y = page_h - (top_cm * cm) - (row_idx + 1) * label_h_cm * cm - row_idx * space_v_cm * cm

        c.setFont("Helvetica", 10)
        c.drawCentredString(x + (label_w_cm * cm) / 2, y + (label_h_cm * cm) / 2 - 10, texto_formateado)

        slot += 1
        if slot == cols * rows:
            c.showPage()
            slot = 0

    if slot > 0:
        c.showPage()
        slot = 0

    c.save()
    print(f"\nPDF generado correctamente: {output_pdf}")
main()
