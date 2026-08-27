# Imagen del pipeline de altas automáticas Thona:
# Flask (recibe la llamada de Odoo) + Playwright/Chromium (RPA) + LibreOffice (export a PDF)

FROM python:3.11-slim

WORKDIR /app

# LibreOffice: necesario para generar el PDF de la Solicitud de Movimiento
# (ver generar_solicitud_movimiento.py, usa `soffice --headless --convert-to pdf`)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Instala Chromium + todas las librerías de sistema que necesita para correr
# en un contenedor sin entorno gráfico (--with-deps se encarga de eso)
RUN python -m playwright install --with-deps chromium

# Código del pipeline y plantillas
COPY generar_layout_asegurados.py .
COPY generar_solicitud_movimiento.py .
COPY rpa_thona_alta.py .
COPY webhook_odoo.py .
COPY Layout_Asegurados_template.xlsx .
COPY Solicitud_Movimientos_template.xlsx .

# Variables de entorno esperadas en tiempo de ejecución (se configuran en
# EasyPanel, NUNCA se hardcodean aquí ni se suben al repo):
#   THONA_USER, THONA_PASS, ODOO_WEBHOOK_SECRET

EXPOSE 5001

CMD ["python", "webhook_odoo.py"]
