#!/usr/bin/env python3
"""
Scraper de Google Maps con Playwright.

Flujo:
    1. Abre Google Maps
    2. Busca "[Negocio] en [Ciudad, Estado]"
    3. Espera resultados
    4. Hace scroll por la lista completa
    5. Extrae las tarjetas de negocio
    6. Entra a cada ficha para los datos que no salen en la lista
    7. Guarda en SQLite conforme avanza (reanudable)
    8. Exporta a Excel

Uso básico:
    python gmaps_scraper.py --negocio "Dentistas" --ciudad "Guadalajara, Jalisco"

Ver README.md para todas las opciones.
"""

from __future__ import annotations

import argparse
import asyncio
import csv
import hashlib
import random
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urlparse

from playwright.async_api import (
    Error as PWError,
    TimeoutError as PWTimeout,
    async_playwright,
)

import phone_mx
from storage import Store, export_excel

# --------------------------------------------------------------------------
# Configuración
# --------------------------------------------------------------------------

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36"
)

FEED = 'div[role="feed"]'
CARD_LINK = 'div[role="feed"] a[href*="/maps/place/"]'

# Varios selectores por campo: Google rota clases con frecuencia, así que
# probamos en cascada y nos quedamos con el primero que responda.
SEL = {
    "nombre": ["h1.DUwDvf", 'h1[class*="DUwDvf"]', 'div[role="main"] h1'],
    "categoria": [
        "button.DkEaL",
        'button[jsaction*="category"]',
        'div.LBgpqf button[jsaction*="pane"]',
    ],
    "rating_box": ["div.F7nice", 'div[class*="F7nice"]'],
    "direccion": [
        'button[data-item-id="address"] div.Io6YTe',
        'button[data-item-id="address"]',
    ],
    "website": ['a[data-item-id="authority"]', 'a[aria-label^="Sitio web"]'],
    "plus_code": ['button[data-item-id="oloc"] div.Io6YTe'],
    "precio": ['span[aria-label*="Precio"]', 'span[aria-label*="Price"]'],
    "cerrado": ["div.fCEvvc", 'span[style*="color"] >> text=/Cerrado permanentemente/i'],
    "horarios_btn": [
        'div[jsaction*="openhours"]',
        'button[data-item-id="oh"]',
        'div[aria-label*="Horario"]',
    ],
}

EMAIL_RE = re.compile(
    r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,24}"
)
# Basura que aparece en el HTML pero no es el email del negocio. Tres familias:
# archivos y paquetes que el regex confunde con emails, dominios de plantilla
# que nadie reemplazó, y correos de la plataforma (no del negocio).
EMAIL_BLOCKLIST = re.compile(
    r"("
    # archivos, sprites y paquetes npm
    r"\.png$|\.jpg$|\.jpeg$|\.gif$|\.svg$|\.webp$|\.ico$|\.css$|\.js$|"
    r"@2x|@3x|u003|core-js|@sentry|@babel|@types|@fontawesome|@wordpress|@emotion|"
    # placeholders de plantilla que nadie cambió
    r"@dominio\.|@tudominio|@sudominio|@midominio|@example\.|@ejemplo\.|"
    r"@domain\.com|@yourdomain|@email\.com|@correo\.com|@empresa\.com|"
    r"^(usuario|nombre|correo|tucorreo|tuemail|youremail|email|test|demo)@|"
    # plataformas y agregadores: es el correo del portal, no del negocio
    r"@doctoralia\.|@sentry\.|sentry\.io|@wix|wixpress|@squarespace|@godaddy|"
    r"@shopify|@mailchimp|@jimdo|@weebly|@sitey|@webnode|@blogger\.|"
    r"@w3\.org|@schema\.org|@example\.org|noreply|no-reply|donotreply"
    r")",
    re.IGNORECASE,
)
WHATSAPP_RE = re.compile(r"(?:wa\.me/|api\.whatsapp\.com/send\?phone=)(\+?\d{8,15})")

CONSENT_TEXTS = [
    "Aceptar todo", "Rechazar todo", "Accept all", "Reject all",
    "Acepto", "Aceitar tudo", "Alle akzeptieren",
]


@dataclass
class Config:
    negocio: str = ""
    ciudad: str = ""
    batch: str | None = None
    max_resultados: int = 120
    headless: bool = True
    lento: float = 1.0
    db: str = "gmaps.db"
    salida: str = ""
    idioma: str = "es"
    pais: str = "MX"
    emails: bool = True
    refresh: bool = False
    ift: str | None = None
    solo_export: bool = False
    max_scrolls: int = 60
    timeout: int = 30000


@dataclass
class Card:
    url: str
    place_key: str
    nombre_lista: str = ""
    extras: dict = field(default_factory=dict)


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------

def log(msg: str) -> None:
    print(msg, flush=True)


async def pause(cfg: Config, base: float = 1.0) -> None:
    """Espera con jitter. Sin esto Maps empieza a servir páginas vacías."""
    await asyncio.sleep(random.uniform(base * 0.7, base * 1.4) * cfg.lento)


def place_key_from_url(url: str) -> str:
    """
    Identificador estable del lugar. Google embebe el feature id en la URL
    como !1s0x<hex>:0x<hex>; es lo más cercano a un place_id sin usar la API.
    """
    m = re.search(r"!1s(0x[0-9a-f]+:0x[0-9a-f]+)", url)
    if m:
        return m.group(1)
    m = re.search(r"[?&]cid=(\d+)", url)
    if m:
        return f"cid:{m.group(1)}"
    m = re.search(r"/maps/place/([^/@]+)", url)
    base = m.group(1) if m else url
    return "h:" + hashlib.sha1(base.encode("utf-8", "ignore")).hexdigest()[:20]


def parse_rating_block(text: str) -> tuple[float | None, int | None]:
    """
    Convierte el bloque de rating de Maps a números.
    Ejemplos: '4.7(1,234)'  '4,7 (89)'  '5.0(3)'  '4.3\n(2.1 mil)'
    """
    if not text:
        return None, None
    clean = text.replace(" ", " ").replace("\xa0", " ").strip()

    rating = None
    m = re.search(r"(\d+[.,]\d+|\d+)\s*(?:\(|estrella|star|$)", clean)
    if m:
        try:
            rating = float(m.group(1).replace(",", "."))
        except ValueError:
            rating = None
    if rating is not None and not (0 < rating <= 5):
        rating = None

    n = None
    m = re.search(r"\(([\d.,\s]+)\s*(mil|k)?\)", clean, re.IGNORECASE)
    if m:
        digits = re.sub(r"[^\d.,]", "", m.group(1))
        mult = 1000 if m.group(2) else 1
        try:
            if mult == 1000:
                n = int(float(digits.replace(",", ".")) * 1000)
            else:
                n = int(re.sub(r"[.,\s]", "", digits))
        except ValueError:
            n = None
    return rating, n


def coords_from_url(url: str) -> tuple[float | None, float | None]:
    m = re.search(r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)", url)
    if m:
        return float(m.group(1)), float(m.group(2))
    m = re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", url)
    if m:
        return float(m.group(1)), float(m.group(2))
    return None, None


def pick_best_email(emails: list[str], domain: str | None) -> str | None:
    if not emails:
        return None
    if domain:
        host = domain.lower().lstrip("www.")
        same = [e for e in emails if e.lower().endswith("@" + host) or host in e.lower()]
        if same:
            emails = same
    priority = ("contacto", "info", "hola", "ventas", "citas", "contact")
    for p in priority:
        for e in emails:
            if e.lower().startswith(p):
                return e
    return emails[0]


# --------------------------------------------------------------------------
# Extracción de la ficha
# --------------------------------------------------------------------------

async def first_text(page, keys: list[str], timeout: int = 1200) -> str:
    for sel in keys:
        try:
            loc = page.locator(sel).first
            if await loc.count():
                txt = (await loc.inner_text(timeout=timeout)).strip()
                if txt:
                    return txt
        except (PWTimeout, PWError):
            continue
    return ""


async def first_attr(page, keys: list[str], attr: str, timeout: int = 1200) -> str:
    for sel in keys:
        try:
            loc = page.locator(sel).first
            if await loc.count():
                val = await loc.get_attribute(attr, timeout=timeout)
                if val:
                    return val.strip()
        except (PWTimeout, PWError):
            continue
    return ""


async def extract_phone(page) -> tuple[str, str]:
    """
    Devuelve (telefono, etiqueta). El botón del teléfono lleva el número en
    data-item-id="phone:tel:+52...", que es más limpio que el texto visible.
    """
    try:
        loc = page.locator('button[data-item-id^="phone:tel:"]').first
        if await loc.count():
            item = await loc.get_attribute("data-item-id") or ""
            aria = await loc.get_attribute("aria-label") or ""
            num = item.split("phone:tel:")[-1]
            if num:
                return num, aria
    except (PWTimeout, PWError):
        pass
    txt = await first_text(page, ['button[aria-label^="Teléfono"]', 'button[aria-label^="Phone"]'])
    return (txt, txt)


async def extract_hours(page) -> str:
    """Horario semanal en una línea: 'Lun 9–18; Mar 9–18; ...'"""
    try:
        rows = page.locator('table[class*="eK4R0e"] tr, div[class*="t39EBf"] table tr')
        n = await rows.count()
        parts = []
        for i in range(min(n, 7)):
            txt = (await rows.nth(i).inner_text(timeout=800)).strip()
            txt = re.sub(r"\s*\n\s*", " ", txt)
            if txt:
                parts.append(txt)
        if parts:
            return "; ".join(parts)
    except (PWTimeout, PWError):
        pass
    aria = await first_attr(page, SEL["horarios_btn"], "aria-label")
    return re.sub(r"\s+", " ", aria).strip()


async def scrape_detail(page, card: Card, cfg: Config) -> dict:
    await page.goto(card.url, wait_until="domcontentloaded", timeout=cfg.timeout)
    try:
        await page.wait_for_selector(SEL["nombre"][0], timeout=cfg.timeout)
    except PWTimeout:
        pass
    await pause(cfg, 0.8)

    nombre = await first_text(page, SEL["nombre"]) or card.nombre_lista
    categoria = await first_text(page, SEL["categoria"])
    rating_txt = await first_text(page, SEL["rating_box"])
    rating, num_resenas = parse_rating_block(rating_txt)

    if rating is None:
        aria = await first_attr(page, ['span[role="img"][aria-label*="estrella"]',
                                       'span[role="img"][aria-label*="star"]'], "aria-label")
        r2, n2 = parse_rating_block(aria)
        rating = rating if rating is not None else r2
        num_resenas = num_resenas if num_resenas is not None else n2

    direccion = await first_text(page, SEL["direccion"])
    website = await first_attr(page, SEL["website"], "href")
    plus_code = await first_text(page, SEL["plus_code"])
    precio = await first_attr(page, SEL["precio"], "aria-label")
    horarios = await extract_hours(page)
    telefono_raw, tel_label = await extract_phone(page)

    cuerpo = await first_text(page, ['div[role="main"]'], timeout=2000)
    cerrado = bool(re.search(r"cerrado permanentemente|permanently closed", cuerpo, re.I))

    lat, lng = coords_from_url(page.url)

    return {
        "place_key": card.place_key,
        "query": cfg.negocio,
        "ubicacion": cfg.ciudad,
        "nombre": nombre,
        "categoria": categoria,
        "telefono_raw": telefono_raw,
        "telefono_label": tel_label,
        "website": website,
        "rating": rating,
        "num_resenas": num_resenas,
        "maps_url": page.url,
        "direccion": direccion,
        "lat": lat,
        "lng": lng,
        "plus_code": plus_code,
        "nivel_precio": precio,
        "horarios": horarios,
        "cerrado_permanente": int(cerrado),
    }


# --------------------------------------------------------------------------
# Email desde el sitio web (solo home)
# --------------------------------------------------------------------------

async def scrape_email(context, url: str, cfg: Config) -> tuple[str | None, list[str], str | None, str]:
    """
    Abre la home del sitio y saca emails + WhatsApp.
    Devuelve (email_principal, todos_los_emails, whatsapp, html_para_contexto).
    """
    if not url:
        return None, [], None, ""

    page = await context.new_page()
    # Bloqueamos recursos pesados: solo necesitamos el HTML.
    await page.route(
        "**/*",
        lambda route: asyncio.ensure_future(
            route.abort()
            if route.request.resource_type in {"image", "media", "font", "stylesheet"}
            else route.continue_()
        ),
    )
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=20000)
        await page.wait_for_timeout(1200)
        html = await page.content()

        mailtos = await page.eval_on_selector_all(
            'a[href^="mailto:"]',
            "els => els.map(e => e.getAttribute('href'))",
        )
        found: list[str] = []
        for href in mailtos or []:
            m = EMAIL_RE.search((href or "").replace("mailto:", ""))
            if m:
                found.append(m.group(0))

        for m in EMAIL_RE.finditer(html):
            found.append(m.group(0))

        seen, emails = set(), []
        for e in found:
            e = e.strip().rstrip(".").lower()
            if e in seen or EMAIL_BLOCKLIST.search(e) or len(e) > 80:
                continue
            seen.add(e)
            emails.append(e)

        wa = None
        mw = WHATSAPP_RE.search(html)
        if mw:
            wa = mw.group(1)

        host = urlparse(url).netloc.lower().replace("www.", "")
        return pick_best_email(emails, host), emails, wa, html[:200000]
    except (PWTimeout, PWError):
        return None, [], None, ""
    finally:
        await page.close()


# --------------------------------------------------------------------------
# Listado: consentimiento, búsqueda, scroll
# --------------------------------------------------------------------------

async def handle_consent(page, cfg: Config) -> None:
    if "consent." not in page.url and "/sorry/" not in page.url:
        return
    log("  · Pantalla de consentimiento detectada, aceptando…")
    for txt in CONSENT_TEXTS:
        try:
            btn = page.get_by_role("button", name=re.compile(txt, re.I)).first
            if await btn.count():
                await btn.click(timeout=4000)
                await page.wait_for_load_state("domcontentloaded", timeout=15000)
                await pause(cfg, 1.5)
                return
        except (PWTimeout, PWError):
            continue
    try:
        await page.locator('form button, form input[type="submit"]').first.click(timeout=4000)
        await pause(cfg, 1.5)
    except (PWTimeout, PWError):
        pass


async def open_search(page, cfg: Config) -> bool:
    query = f"{cfg.negocio} en {cfg.ciudad}"
    url = (
        "https://www.google.com/maps/search/"
        + query.replace(" ", "+")
        + f"?hl={cfg.idioma}&gl={cfg.pais.lower()}"
    )
    log(f"\n▶ Buscando: {query}")
    await page.goto(url, wait_until="domcontentloaded", timeout=cfg.timeout)
    await handle_consent(page, cfg)
    await pause(cfg, 2.0)

    try:
        await page.wait_for_selector(FEED, timeout=cfg.timeout)
        return True
    except PWTimeout:
        # Si la búsqueda es muy específica, Maps abre la ficha directa sin lista.
        if await page.locator(SEL["nombre"][0]).count():
            log("  · Un solo resultado (Maps abrió la ficha directamente).")
            return True
        log("  ! No cargó la lista de resultados. ¿Búsqueda sin resultados o bloqueo?")
        return False


async def scroll_feed(page, cfg: Config) -> list[Card]:
    feed = page.locator(FEED)
    if not await feed.count():
        # Ficha única
        url = page.url
        return [Card(url=url, place_key=place_key_from_url(url))]

    previo, estancado = 0, 0
    for i in range(cfg.max_scrolls):
        await feed.evaluate("el => el.scrollBy(0, el.scrollHeight)")
        await pause(cfg, 1.4)

        n = await page.locator(CARD_LINK).count()

        fin = await page.get_by_text(
            re.compile(r"llegado al final de la lista|reached the end of the list", re.I)
        ).count()
        log(f"  · scroll {i + 1}: {n} resultados")

        if fin:
            log("  · Fin de la lista.")
            break
        if n >= cfg.max_resultados:
            log(f"  · Alcanzado el máximo de {cfg.max_resultados}.")
            break
        if n == previo:
            estancado += 1
            if estancado >= 3:
                log("  · Sin resultados nuevos, deteniendo el scroll.")
                break
        else:
            estancado = 0
        previo = n

    # Recolectamos href + nombre + rating de la lista (rápido, sin abrir fichas).
    raw = await page.eval_on_selector_all(
        CARD_LINK,
        """els => els.map(a => {
            const cont = a.closest('div[jsaction]') || a.parentElement;
            const txt  = cont ? cont.innerText : '';
            return { href: a.href, label: a.getAttribute('aria-label') || '', txt };
        })""",
    )

    cards, vistos = [], set()
    for item in raw:
        href = item.get("href") or ""
        if "/maps/place/" not in href:
            continue
        key = place_key_from_url(href)
        if key in vistos:
            continue
        vistos.add(key)
        cards.append(Card(url=href, place_key=key, nombre_lista=item.get("label", "")))
        if len(cards) >= cfg.max_resultados:
            break
    return cards


# --------------------------------------------------------------------------
# Orquestación
# --------------------------------------------------------------------------

async def run_search(context, store: Store, cfg: Config) -> int:
    page = await context.new_page()
    nuevos = 0
    errores = 0
    try:
        if not await open_search(page, cfg):
            store.log_search(cfg.negocio, cfg.ciudad, 0, "sin_resultados")
            return 0

        cards = await scroll_feed(page, cfg)
        log(f"  ✓ {len(cards)} negocios en la lista.")

        ya = set() if cfg.refresh else store.all_keys()
        pendientes = [c for c in cards if c.place_key not in ya]
        if len(pendientes) < len(cards):
            log(f"  · {len(cards) - len(pendientes)} ya estaban en la base, se saltan.")

        for i, card in enumerate(pendientes, 1):
            etiqueta = card.nombre_lista or card.place_key
            # Un solo negocio raro NUNCA debe tumbar la corrida completa: si algo
            # explota aquí, se registra, se sigue con el siguiente, y lo ya
            # guardado permanece en la base.
            try:
                data = await scrape_detail(page, card, cfg)

                contexto = data.pop("telefono_label", "") or ""
                telefono_raw = data.pop("telefono_raw", "")

                email = wa = None
                emails_todos: list[str] = []
                if cfg.emails and data.get("website"):
                    try:
                        email, emails_todos, wa, html = await scrape_email(
                            context, data["website"], cfg
                        )
                        contexto = f"{contexto}\n{html[:20000]}"
                    except Exception as exc:  # el sitio del negocio es terreno hostil
                        log(f"       · sitio web falló ({type(exc).__name__}), sigo sin email")

                info = phone_mx.classify(telefono_raw, context=contexto, region=cfg.pais)
                data.update(info.as_dict())
                data["email"] = email
                data["emails_todos"] = emails_todos
                data["whatsapp"] = wa

                store.upsert(data)
                nuevos += 1
                log(
                    f"  [{i}/{len(pendientes)}] ✓ {(data.get('nombre') or '?')[:42]:<42} "
                    f"{(data.get('telefono') or '—'):<15} "
                    f"{(data.get('tipo_telefono') or '—'):<14} "
                    f"{(email or '—')}"
                )
            except Exception as exc:
                errores += 1
                log(f"  [{i}/{len(pendientes)}] ✗ {etiqueta[:42]}: {type(exc).__name__}: {exc}")
                if errores >= 15:
                    log("  ! Demasiados errores seguidos; corto esta búsqueda "
                        "(lo guardado se conserva). Prueba con --lento 2.")
                    break
                # La página pudo quedar en mal estado; la reciclamos.
                try:
                    await page.close()
                except Exception:
                    pass
                page = await context.new_page()

            await pause(cfg, 1.1)

        store.log_search(cfg.negocio, cfg.ciudad, nuevos, "ok" if not errores else f"ok ({errores} fallos)")
    finally:
        try:
            await page.close()
        except Exception:
            pass
    return nuevos


def load_batch(path: str) -> list[tuple[str, str]]:
    """CSV con columnas negocio,ciudad (con o sin encabezado)."""
    pares = []
    with open(path, newline="", encoding="utf-8-sig") as fh:
        for row in csv.reader(fh):
            if len(row) < 2:
                continue
            a, b = row[0].strip(), row[1].strip()
            if not a or not b:
                continue
            if a.lower() in {"negocio", "business", "query"}:
                continue
            pares.append((a, b))
    return pares


async def main_async(cfg: Config) -> None:
    if cfg.ift:
        if phone_mx.load_ift_ranges(cfg.ift):
            log(f"✓ Rangos IFT cargados ({phone_mx.ift_count():,} rangos) → fijo/celular exacto.")
        else:
            log(f"! No pude leer los rangos IFT en {cfg.ift}; sigo con heurísticas.")
    else:
        log("· Sin CSV del IFT: el tipo de teléfono será heurístico. Ver README.")

    store = Store(cfg.db)

    trabajos = load_batch(cfg.batch) if cfg.batch else [(cfg.negocio, cfg.ciudad)]
    if not trabajos or not trabajos[0][0]:
        log("! Nada que buscar. Usa --negocio y --ciudad, o --batch archivo.csv")
        return

    salida = cfg.salida or default_output(cfg, trabajos)
    total = 0
    try:
        total = await _scrapear(store, cfg, trabajos)
    finally:
        # El Excel se escribe SIEMPRE: si algo explota a mitad de camino, los
        # negocios ya guardados salen igual. Nunca se pierde una corrida.
        filas = store.fetch() if cfg.batch else store.fetch(cfg.negocio, cfg.ciudad)
        if filas:
            try:
                export_excel(filas, salida, sheet_name=(cfg.negocio or "Negocios")[:31])
                log(f"\n✓ {total} negocios nuevos en esta corrida · "
                    f"{store.count()} en total en la base.")
                log(f"✓ Excel: {Path(salida).resolve()}")
            except Exception as exc:
                log(f"\n! No pude escribir el Excel ({type(exc).__name__}: {exc}).")
                log(f"  Los datos están a salvo en {Path(cfg.db).resolve()}")
                log(f"  Recupéralos con:  python gmaps_scraper.py --solo-export -o resultados.xlsx")
        else:
            log("\n· No se guardó ningún negocio.")
        store.close()


async def _scrapear(store: Store, cfg: Config, trabajos: list[tuple[str, str]]) -> int:
    total = 0
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=cfg.headless,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ],
        )
        context = await browser.new_context(
            user_agent=USER_AGENT,
            locale="es-MX",
            timezone_id="America/Mexico_City",
            viewport={"width": 1440, "height": 900},
            geolocation=None,
        )
        context.set_default_timeout(cfg.timeout)
        # Quita la bandera webdriver que delata a Playwright.
        await context.add_init_script(
            "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});"
        )
        try:
            for negocio, ciudad in trabajos:
                cfg.negocio, cfg.ciudad = negocio, ciudad
                try:
                    total += await run_search(context, store, cfg)
                except Exception as exc:
                    log(f"  ! Error en '{negocio} en {ciudad}': {type(exc).__name__}: {exc}")
                    store.log_search(negocio, ciudad, 0, f"error: {type(exc).__name__}")
        finally:
            for cerrar in (context.close, browser.close):
                try:
                    await cerrar()
                except Exception:
                    pass
    return total


def default_output(cfg: Config, trabajos: list[tuple[str, str]]) -> str:
    if cfg.batch:
        base = Path(cfg.batch).stem
    else:
        base = f"{trabajos[0][0]}_{trabajos[0][1]}"
    slug = re.sub(r"[^\w\-]+", "_", base, flags=re.UNICODE).strip("_").lower()
    return f"{slug or 'gmaps'}.xlsx"


def parse_args(argv: list[str]) -> Config:
    p = argparse.ArgumentParser(
        description="Scraper de Google Maps con Playwright",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""ejemplos:
  python gmaps_scraper.py --negocio "Dentistas" --ciudad "Guadalajara, Jalisco"
  python gmaps_scraper.py --negocio "Gimnasios" --ciudad "CDMX" --max 200 --ver
  python gmaps_scraper.py --batch busquedas.csv --ift ift_numeracion.csv
  python gmaps_scraper.py --solo-export --salida todo.xlsx
""",
    )
    p.add_argument("--negocio", "-n", default="", help="Giro o tipo de negocio")
    p.add_argument("--ciudad", "-c", default="", help='Ciudad, Estado (ej. "Mérida, Yucatán")')
    p.add_argument("--batch", "-b", help="CSV con columnas negocio,ciudad")
    p.add_argument("--max", "-m", type=int, default=120, dest="max_resultados",
                   help="Máximo de negocios por búsqueda (default 120)")
    p.add_argument("--ver", action="store_true", help="Abre el navegador visible (debug)")
    p.add_argument("--lento", type=float, default=1.0,
                   help="Multiplicador de pausas; súbelo si Maps te empieza a bloquear")
    p.add_argument("--db", default="gmaps.db", help="Base SQLite para el checkpoint")
    p.add_argument("--salida", "-o", default="", help="Ruta del Excel de salida")
    p.add_argument("--sin-emails", action="store_true", help="No visitar sitios web")
    p.add_argument("--refresh", action="store_true",
                   help="Re-scrapear negocios ya guardados en vez de saltarlos")
    p.add_argument("--ift", help="CSV de numeración del IFT (fijo/celular exacto)")
    p.add_argument("--solo-export", action="store_true",
                   help="No scrapear: solo exportar lo que ya está en la base")
    p.add_argument("--idioma", default="es")
    p.add_argument("--pais", default="MX")
    a = p.parse_args(argv)

    return Config(
        negocio=a.negocio,
        ciudad=a.ciudad,
        batch=a.batch,
        max_resultados=a.max_resultados,
        headless=not a.ver,
        lento=a.lento,
        db=a.db,
        salida=a.salida,
        emails=not a.sin_emails,
        refresh=a.refresh,
        ift=a.ift,
        solo_export=a.solo_export,
        idioma=a.idioma,
        pais=a.pais,
    )


def main() -> None:
    cfg = parse_args(sys.argv[1:])
    if cfg.solo_export:
        store = Store(cfg.db)
        filas = store.fetch(cfg.negocio, cfg.ciudad) if (cfg.negocio and cfg.ciudad) else store.fetch()
        salida = cfg.salida or "gmaps_export.xlsx"
        export_excel(filas, salida)
        log(f"✓ {len(filas)} registros exportados a {salida}")
        store.close()
        return
    try:
        asyncio.run(main_async(cfg))
    except KeyboardInterrupt:
        log("\n· Interrumpido. Lo scrapeado quedó guardado; vuelve a correr para continuar.")


if __name__ == "__main__":
    main()
