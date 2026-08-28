"""
Servicio Flask que recibe el webhook de Odoo cuando se confirma una venta
del producto Preventivo (o cualquier plan que use este proveedor Thona),
y orquesta el pipeline completo:

    1. Genera Layout_Asegurados.xlsx con los datos del formulario de compra.
    2. Genera Solicitud_de_Movimiento.xlsx + PDF con vigencia calculada.
    3. Corre el RPA para emitir el alta en el portal Thona.
    4. Escribe el folio de vuelta en la orden de venta de Odoo (vía XML-RPC/JSON-RPC).
    5. Si algo falla, notifica (ej. Slack/correo) para intervención manual —
       NUNCA debe fallar en silencio, porque hay dinero real (venta ya cobrada)
       de por medio.

Cómo se conecta desde Odoo:
    - Automation Rule sobre `sale.order`, disparada en "On Update" cuando
      `state` cambia a 'sale' (orden confirmada) y el producto pertenece
      a la categoría "Preventivo Thona".
    - La Automation Rule llama un webhook (server action tipo "Ejecutar código
      Python" con `requests.post(...)`, o un módulo custom con `ir.actions.server`)
      mandando el payload de abajo a este servicio.

Ejecutar:
    pip install flask --break-system-packages
    python3 webhook_odoo.py
"""

import os
import logging
from datetime import datetime
from flask import Flask, request, jsonify

from generar_layout_asegurados import generar_layout_asegurados
from generar_solicitud_movimiento import generar_solicitud_movimiento
from rpa_thona_alta import emitir_movimiento_thona

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
log = logging.getLogger("thona_pipeline")

TRABAJO_DIR = "/tmp/thona_pipeline"
os.makedirs(TRABAJO_DIR, exist_ok=True)

WEBHOOK_SECRET = os.environ.get("ODOO_WEBHOOK_SECRET", "")  # validar origen de la llamada


@app.route("/webhook/thona/alta", methods=["POST"])
def recibir_alta():
    payload = request.get_json(force=True)

    # --- Seguridad mínima: validar secreto compartido con Odoo ---
    if WEBHOOK_SECRET and request.headers.get("X-Webhook-Secret") != WEBHOOK_SECRET:
        return jsonify({"error": "unauthorized"}), 401

    try:
        orden_id = payload["order_id"]          # id de sale.order en Odoo, para trazabilidad
        asegurado = payload["asegurado"]         # dict con nombres, apellidos, fecha_nacimiento, genero, subgrupo
        contratante = payload.get("contratante", "U HEALTH INSURTECH")
        no_poliza = payload.get("no_poliza", "68873-00")
        rfc_contratante = payload.get("rfc_contratante", "UHI240702UR5")
        tipo_movimiento = payload.get("tipo_movimiento", "alta")  # "alta" o "baja"

        if tipo_movimiento == "baja":
            tipo_modificacion_thona = "Bajas de Asegurados"
            tipo_solicitud = "ENDOSO_D"
            tipo_movimiento_texto = "BAJA DE ASEGURADO"
            descripcion_forma = "DAR DE BAJA A ASEGURADOS"
        else:
            tipo_modificacion_thona = "Altas de Asegurados"
            tipo_solicitud = "ENDOSO_A"
            tipo_movimiento_texto = "ALTA DE ASEGURADO"
            descripcion_forma = "DAR DE ALTA A ASEGURADOS"

        base = f"{TRABAJO_DIR}/orden_{orden_id}_{tipo_movimiento}"

        # 1. Generar Layout de Asegurados
        layout_path = generar_layout_asegurados([asegurado], salida=f"{base}_layout.xlsx")

        # 2. Generar Solicitud de Movimiento (xlsx + pdf)
        datos_solicitud = {
            "no_poliza": no_poliza,
            "ramo": "ACCIDENTES PERSONALES",
            "contratante": contratante,
            "tipo_poliza": "GRUPO o COLECTIVO",
            "tipo_solicitud": tipo_solicitud,
            "tipo_movimiento": tipo_movimiento_texto,
            "observaciones": f"{tipo_movimiento_texto} automática — orden Odoo #{orden_id}",
        }
        xlsx_path, pdf_path = generar_solicitud_movimiento(datos_solicitud, salida_base=f"{base}_solicitud")

        # 3. Correr el RPA (Nivel 2). Si prefieres arrancar en Nivel 1,
        #    comenta este bloque y en vez de eso notifica a un humano con
        #    los 2 archivos ya listos para subir manualmente.
        folio = emitir_movimiento_thona(
            no_poliza=no_poliza,
            tipo_modificacion=tipo_modificacion_thona,
            conteo_polizas=1,
            descripcion=descripcion_forma,
            nombre_contratante=contratante,
            rfc_contratante=rfc_contratante,
            comentario_bitacora=f"{tipo_movimiento_texto} automática vía Odoo — orden #{orden_id}",
            layout_asegurados_path=layout_path,
            solicitud_path=pdf_path,
        )

        log.info(f"Orden {orden_id}: {tipo_movimiento} emitida en Thona, folio {folio}")

        # 4. TODO: escribir el folio de vuelta en Odoo (XML-RPC/JSON-RPC)
        # actualizar_folio_en_odoo(orden_id, folio)

        return jsonify({"ok": True, "folio": folio}), 200

    except Exception as e:
        log.exception(f"Fallo el pipeline para la orden {payload.get('order_id')}")
        # TODO: notificar a Slack/correo/Odoo activity para intervención manual
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, threaded=True)
