#!/usr/bin/env python3
"""
Descarga el Plan Técnico Fundamental de Numeración del IFT usando Playwright.

El portal del IFT es una app JSF: el archivo no está en una URL directa, hay que
pulsar el botón de descarga. Por eso se automatiza con el mismo navegador que ya
usas para el scraper.

    python descargar_ift.py            # guarda ift_numeracion.csv
    python descargar_ift.py --ver      # con navegador visible, para ver qué pasa

Luego:
    python gmaps_scraper.py -n "Dentistas" -c "Mérida" --ift ift_numeracion.csv

Si el portal cambia y esto falla, la ruta manual es:
  1. Abrir https://sns.ift.org.mx:8081/sns-frontend/planes-numeracion/descarga-publica.xhtml
  2. Elegir "Numeración Geográfica" y descargar el CSV
  3. Guardarlo como ift_numeracion.csv junto al scraper
El parser acepta cualquier CSV con columnas ZONA, SERIE, NUMERACION_INICIAL,
NUMERACION_FINAL y MODALIDAD (o TIPO_SERVICIO), sin importar el orden.
"""

from __future__ import annotations

import argparse
import asyncio
import sys
import zipfile
from pathlib import Path

from playwright.async_api import async_playwright

PORTAL = "https://sns.ift.org.mx:8081/sns-frontend/planes-numeracion/descarga-publica.xhtml"

# Textos con los que suele estar etiquetado el botón/enlace de descarga.
ETIQUETAS = [
    "Numeración Geográfica",
    "Numeracion Geografica",
    "Descargar",
    "Exportar",
    "CSV",
]


async def descargar(destino: Path, headless: bool) -> int:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless, args=["--no-sandbox"])
        # El portal del IFT ha tenido certificados con cadena incompleta.
        ctx = await browser.new_context(ignore_https_errors=True, accept_downloads=True)
        page = await ctx.new_page()
        try:
            print(f"· Abriendo {PORTAL}")
            await page.goto(PORTAL, wait_until="domcontentloaded", timeout=60000)
            await page.wait_for_timeout(3000)

            for etiqueta in ETIQUETAS:
                objetivo = page.get_by_text(etiqueta, exact=False).first
                if not await objetivo.count():
                    continue
                print(f"· Intentando con «{etiqueta}»")
                try:
                    async with page.expect_download(timeout=90000) as dl_info:
                        await objetivo.click(timeout=10000)
                    dl = await dl_info.value
                    sugerido = dl.suggested_filename or "ift.csv"
                    crudo = destino.with_name("_ift_descarga_" + sugerido)
                    await dl.save_as(crudo)
                    print(f"· Descargado: {crudo.name} ({crudo.stat().st_size:,} bytes)")

                    if zipfile.is_zipfile(crudo):
                        with zipfile.ZipFile(crudo) as z:
                            csvs = [n for n in z.namelist() if n.lower().endswith(".csv")]
                            if not csvs:
                                print("! El ZIP no traía ningún CSV.")
                                return 1
                            with z.open(csvs[0]) as src:
                                destino.write_bytes(src.read())
                        crudo.unlink()
                    else:
                        crudo.replace(destino)

                    print(f"✓ Guardado en {destino.resolve()}")
                    return 0
                except Exception as exc:
                    print(f"  · no funcionó ({type(exc).__name__})")
                    continue

            print("! No encontré el botón de descarga. El portal debió cambiar.")
            print("  Descárgalo a mano; las instrucciones están arriba en este archivo.")
            captura = destino.with_name("ift_portal.png")
            await page.screenshot(path=str(captura), full_page=True)
            print(f"  Guardé una captura del portal en {captura} por si ayuda.")
            return 1
        finally:
            await ctx.close()
            await browser.close()


def main() -> int:
    ap = argparse.ArgumentParser(description="Descarga los rangos de numeración del IFT")
    ap.add_argument("-o", "--salida", default="ift_numeracion.csv")
    ap.add_argument("--ver", action="store_true", help="Navegador visible")
    a = ap.parse_args()
    try:
        return asyncio.run(descargar(Path(a.salida), headless=not a.ver))
    except Exception as exc:
        print(f"! Falló la descarga: {type(exc).__name__}: {exc}")
        print("  Descárgalo a mano desde el portal del IFT (ver encabezado del archivo).")
        return 1


if __name__ == "__main__":
    sys.exit(main())
