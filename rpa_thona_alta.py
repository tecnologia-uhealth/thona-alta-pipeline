"""
RPA con Playwright que automatiza el flujo del portal de agentes Thona
(login -> Mi Área de Trabajo -> Emisiones -> Modificación de Póliza ->
llenar formulario -> subir 2 archivos -> enviar trámite -> capturar folio).

✅ ESTADO: Selectores confirmados contra el portal real (agosto 2026),
capturados con `playwright codegen` durante un alta real exitosa
(folio E-MD-20260826-08217). El único punto sin verificar en múltiples
casos es la selección de fecha en el datepicker (celda "1" = día 1 del
mes que esté mostrando el calendario al abrirse) — confirmar que esto
se comporta igual si se corre en un día distinto del mes.

Requiere:
    pip install playwright
    python -m playwright install chromium   (o usar --channel=chrome si ya tienes Chrome)

Variables de entorno esperadas (NUNCA hardcodear credenciales en el código):
    THONA_USER, THONA_PASS
"""

import os
import re
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

THONA_URL = "https://thona-agentes.azurewebsites.net/"


def emitir_movimiento_thona(
    no_poliza: str,
    tipo_modificacion: str,        # texto exacto del link, ej. "Altas de Asegurados" o "Bajas de Asegurados"
    conteo_polizas: int,
    descripcion: str,
    nombre_contratante: str,
    rfc_contratante: str,
    comentario_bitacora: str,
    layout_asegurados_path: str,   # ruta al .xlsx generado (obligatorio, "Requerido")
    solicitud_path: str,           # ruta al .xlsm/.pdf generado (opcional, "Opcional")
    headless: bool = True,
    screenshot_dir: str = "/tmp/thona_rpa_debug",
) -> str:
    """
    Devuelve el folio generado por el portal (ej. "E-MD-20260826-08217").
    Lanza excepción si algo falla, para que el orquestador (webhook_odoo.py)
    lo capture, marque el registro en estado 'error' y reintente después.
    """
    user = os.environ["THONA_USER"]
    password = os.environ["THONA_PASS"]
    os.makedirs(screenshot_dir, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context()
        page = context.new_page()

        try:
            # --- 1. Login ---
            page.goto(THONA_URL, timeout=30000)
            page.get_by_role("textbox", name="Ingresa tu usuario").fill(user)
            page.get_by_role("button", name="Continuar...").click()
            page.get_by_role("button", name="Si...").click()
            page.get_by_role("textbox", name="Ingresa tu contraseña").fill(password)
            page.get_by_role("button", name="Continuar...").click()
            page.get_by_role("button", name="Aceptar").click()
            page.get_by_role("button", name="Continuar").click()

            # --- 2. Navegar: Mi Área de Trabajo -> Emisiones -> Modificación de Póliza ---
            # Esto abre una VENTANA EMERGENTE (popup) — es una pestaña/página nueva,
            # no la misma. Todo lo que sigue ocurre en `formulario`, no en `page`.
            page.get_by_role("link", name="Mi Área de Trabajo").click()
            page.get_by_role("button", name="Emisiones").click()
            with page.expect_popup() as popup_info:
                page.get_by_role("link", name="Modificación de Póliza").click()
            formulario = popup_info.value

            # --- 3. Tipo de trámite: "Modificación póliza" (no Póliza nueva / Renovación / Thona Elite) ---
            formulario.get_by_text(
                "Modificación póliza Modificación pólizaPóliza nuevaRenovaciónThona Elite"
            ).click()

            # --- 4. Llenar formulario inicial ---
            formulario.locator("#ctl00_ContentPlaceHolder1_txtNoPoliza").fill(no_poliza)

            # Dropdown "Tipo de Modificación" — el .nth(1) es frágil (depende de que
            # solo haya 2 dropdowns "Seleccione una opción..." en la página y este
            # sea el segundo). Si el portal cambia de diseño, esto es lo primero a revisar.
            formulario.get_by_role("button", name="Seleccione una opción...   ").nth(1).click()
            formulario.locator("a").filter(has_text=tipo_modificacion).click()

            formulario.locator("#ctl00_ContentPlaceHolder1_txtConteoPol").fill(str(conteo_polizas))

            # Selector de fecha de aplicación: abre el datepicker y elige el día 1
            # del mes que se esté mostrando (asume que el calendario abre en el mes actual).
            formulario.get_by_role("button").filter(has_text=re.compile(r"^$")).click()
            formulario.get_by_role("cell", name="1").nth(1).click()

            formulario.locator("#ctl00_ContentPlaceHolder1_txtDescMod").fill(descripcion)

            # Clic sin llenar — parte del flujo grabado real (posiblemente dispara
            # validación del formulario). Se deja igual que en la grabación exitosa.
            formulario.locator("div:nth-child(11) > div > .form-group").first.click()

            formulario.locator("#ctl00_ContentPlaceHolder1_txtContraNom").fill(nombre_contratante)
            formulario.locator("#ctl00_ContentPlaceHolder1_txtContraRFC").fill(rfc_contratante)
            formulario.locator("#ctl00_ContentPlaceHolder1_txtCtrlPA").click()
            formulario.locator("#ctl00_ContentPlaceHolder1_txtObs").fill(comentario_bitacora)

            formulario.get_by_role("link", name="Continuar   ").click()

            # --- 5. Subir Layout de Asegurados (casilla "Requerido") ---
            formulario.get_by_role("listitem").filter(
                has_text="1Requerido No"
            ).get_by_role("insertion").click()
            formulario.get_by_role("button", name="Choose File").set_input_files(layout_asegurados_path)

            # --- 6. Subir Solicitud de Movimientos (casilla "Opcional") ---
            if solicitud_path:
                formulario.get_by_role("listitem").filter(
                    has_text="3Opcional No adjuntadoOtro"
                ).get_by_role("insertion").click()
                formulario.get_by_role("button", name="Choose File").set_input_files(solicitud_path)

            formulario.get_by_role("link", name="Continuar   ").click()

            # --- 7. Enviar trámite ---
            formulario.get_by_role("link", name="Enviar trámite   ").click()
            formulario.get_by_role("button", name="Aceptar").click()
            formulario.get_by_role("link", name="Ir al trámite   ").click()

            # --- 8. Capturar folio de la pantalla "VISUALIZANDO TRÁMITE" ---
            # Formato observado: E-MD-20260826-08217
            formulario.wait_for_load_state("networkidle", timeout=20000)
            texto_pagina = formulario.locator("body").inner_text()
            match = re.search(r"[EB]-MD-\d{8}-\d{5}", texto_pagina)
            folio = match.group(0) if match else "FOLIO_NO_DETECTADO"

            formulario.screenshot(path=f"{screenshot_dir}/exito_folio_{folio}.png")
            return folio

        except (PWTimeout, Exception) as e:
            page.screenshot(path=f"{screenshot_dir}/error.png")
            raise RuntimeError(f"Fallo el RPA en el portal Thona: {e}") from e
        finally:
            browser.close()


if __name__ == "__main__":
    # Ejecución manual de prueba (requiere THONA_USER / THONA_PASS en el entorno)
    folio = emitir_movimiento_thona(
        no_poliza="68873-00",
        tipo_modificacion="Altas de Asegurados",
        conteo_polizas=1,
        descripcion="DAR DE ALTA A ASEGURADOS",
        nombre_contratante="U HEALTH INSURTECH SA DE CV",
        rfc_contratante="UHI240702UR5",
        comentario_bitacora="Alta automática vía Odoo",
        layout_asegurados_path="/tmp/Layout_Asegurados_TEST.xlsx",
        solicitud_path=None,
        headless=False,
    )
    print("Folio generado:", folio)
