from reportlab.graphics.barcode import eanbc
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPM

# Crear código EAN-13
ean = eanbc.Ean13BarcodeWidget('123456789012')

# Ajustar tamaño y diseño
ean.barHeight = 25  # altura de barras normales
ean.humanReadable = True  # muestra los números debajo

# Crear dibujo
d = Drawing(200, 100)
d.add(ean)

# Exportar como imagen PNG
renderPM.drawToFile(d, "ean13_profesional.png", fmt="PNG")
