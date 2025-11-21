import re
import os
from openpyxl import Workbook
from tkinter import Tk, filedialog

def procesar_archivo_txt(ruta_txt, ruta_xlsx):
    with open(ruta_txt, "r", encoding="utf-8") as f:
        contenido = f.read()

    bloques = re.findall(r"(8M\d{13}P1.*?)(?=8M\d{13}P1|$)", contenido, re.DOTALL)
    filas = []

    for bloque in bloques:
        match_ean = re.search(r"[A-Z]{2}EC(\d{13})", bloque)
        ean = match_ean.group(1) if match_ean else ""
        cod1 = bloque[12:15]
        cod6 = "02016"
        cod3 = cod4 = ""
        pos_umco = bloque.find("UMCO S.A.")
        if pos_umco != -1:
            pos_cod3 = pos_umco + len("UMCO S.A.") + 21
            cod3 = bloque[pos_cod3-1:pos_cod3 + 5]
            cod4 = bloque[pos_cod3 + 5:pos_cod3 + 9]

        cod5 = ""
        match_cod5 = re.search(r"([A-Z]{2})EC", bloque)
        if match_cod5:
            letras = match_cod5.group(1)
            cod5 = f"{letras}   {bloque[15:17]}"

        cod2 = "C                522"
        subtotal = iva = total = "$0.00"
        cantidad = 0
        if match_ean:
            pos = match_ean.end()
            try:
                subtotal = f"${int(bloque[pos:pos + 9]) / 100:.2f}"
                iva = f"${int(bloque[pos + 10:pos + 18]) / 100:.2f}"
                total = f"${int(bloque[pos + 20:pos + 27]) / 100:.2f}"
                cantidad = int(bloque[pos + 30:pos + 33]) + 3
            except:
                pass

        if ean and cod1 and cod3:
            filas.append([ean, cod1, cod2, cod3, cod4, cod5, cod6, cantidad, subtotal, iva, total])

    # === Guardar en Excel ===
    wb = Workbook()
    ws = wb.active
    ws.title = "deprati"
    encabezado = ['EAN', 'Cod1', 'Cod2', 'Cod3', 'Cod4', 'Cod5', 'Cod6', 'Cantidad', 'Subtotal', 'IVA', 'Total']
    ws.append(encabezado)
    for fila in filas:
        ws.append(fila)

    wb.save(ruta_xlsx)
    print("Archivo Excel generado:", ruta_xlsx)

# === USAR EXPLORADOR PARA SELECCIONAR ARCHIVO .TXT ===
Tk().withdraw()  # Oculta ventana raíz de Tkinter
archivo_txt = filedialog.askopenfilename(
    title="Selecciona el archivo .TXT",
    filetypes=[("Archivos de texto", "*.txt")]
)

if archivo_txt:
    nombre_sin_ext = "datos_imprimir_etiquetas"
    carpeta_destino = os.path.dirname(archivo_txt)
    archivo_xlsx = os.path.join(carpeta_destino, f"{nombre_sin_ext}.xlsx")

    procesar_archivo_txt(archivo_txt, archivo_xlsx)
else:
    print(" No se seleccionó ningún archivo.")
