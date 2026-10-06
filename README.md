# Scraper de Google Maps (Playwright)

Extrae negocios de Google Maps a Excel. Guarda en SQLite conforme avanza, así
que si se cae o lo interrumpes, retoma donde quedó.

## Instalación

```bash
pip install -r requirements.txt
playwright install chromium
```

## Uso

```bash
# Una búsqueda
python gmaps_scraper.py --negocio "Dentistas" --ciudad "Mérida, Yucatán"

# Ver el navegador mientras trabaja (útil la primera vez)
python gmaps_scraper.py -n "Gimnasios" -c "CDMX" --ver --max 30

# Varias búsquedas de corrido
python gmaps_scraper.py --batch busquedas.ejemplo.csv

# Solo exportar lo que ya está en la base, sin volver a scrapear
python gmaps_scraper.py --solo-export --salida clientes.xlsx
```

### Opciones

| Flag | Qué hace |
|---|---|
| `-n, --negocio` | Giro a buscar (`"Dentistas"`) |
| `-c, --ciudad` | `"Ciudad, Estado"` |
| `-b, --batch` | CSV con columnas `negocio,ciudad` |
| `-m, --max` | Máximo de negocios por búsqueda (default 120) |
| `-o, --salida` | Ruta del Excel; si no la das se genera del nombre de la búsqueda |
| `--ver` | Navegador visible en vez de headless |
| `--lento 2` | Duplica las pausas. Súbelo si Maps te empieza a servir páginas vacías |
| `--sin-emails` | No visita los sitios web (mucho más rápido) |
| `--refresh` | Vuelve a scrapear negocios ya guardados en vez de saltarlos |
| `--db` | Archivo SQLite del checkpoint (default `gmaps.db`) |
| `--ift` | CSV de numeración del IFT — ver abajo |
| `--solo-export` | No scrapea, solo exporta la base a Excel |

## Los tres archivos: cuál es cuál

| Archivo | Qué es |
|---|---|
| `busquedas.ejemplo.csv` | **Entrada.** La lista de qué buscar. No contiene resultados. |
| `gmaps.db` | **La fuente de verdad.** Base SQLite donde se guarda cada negocio en cuanto se termina de scrapear. Aquí nunca se pierde nada. |
| `*.xlsx` | **Salida.** Una foto del contenido de la base al terminar. Se puede regenerar cuantas veces quieras. |

El Excel toma su nombre de la búsqueda. Con `--batch busquedas.ejemplo.csv` el
archivo sale como `busquedas_ejemplo.xlsx` — mismo nombre que el CSV de entrada,
lo cual confunde. Usa `-o` para nombrarlo tú:

```bash
python gmaps_scraper.py --batch busquedas.ejemplo.csv -o dentistas_mexico.xlsx
```

### Recuperar los datos si algo salió mal

Nada se pierde: todo está en `gmaps.db`. Para volver a sacar el Excel:

```bash
python gmaps_scraper.py --solo-export -o resultados.xlsx
```

### Abrir el archivo `.db`

Es una base SQLite. Tres formas, de más fácil a más flexible:

```bash
python ver_db.py                    # resumen y estadísticas en consola
python ver_db.py --lista            # todos los negocios
python ver_db.py --csv datos.csv    # volcado a CSV
```

También puedes abrirlo con [DB Browser for SQLite](https://sqlitebrowser.org)
(gratis, interfaz gráfica), o desde Python con el módulo `sqlite3`.

### Si Excel te muestra `MÃ©rida` en vez de `Mérida`

Eso pasa **solo con archivos CSV**, no con los `.xlsx`. Excel asume la
codificación de Windows y no UTF-8 salvo que el archivo lleve un BOM.
El `busquedas.ejemplo.csv` y todo lo que genera `ver_db.py --csv` ya lo llevan.
Para un CSV ajeno que se vea así: Datos → Obtener datos → Desde texto/CSV →
elegir "UTF-8" en el diálogo de importación.

## Datos que extrae

Nombre · Categoría · Teléfono (crudo y en E.164) · **Fijo/Celular** · LADA ·
Website · Email · Rating · Nº de reseñas · URL de Google Maps · Dirección ·
Latitud/Longitud · Horarios · Nivel de precio · WhatsApp · Cerrado permanentemente

El Excel sale con encabezado congelado, autofiltro, hipervínculos en website /
email / Maps, y una hoja **Resumen** con conteos (cuántos tienen teléfono, sitio,
email, cuántos son celular vs fijo, rating promedio, y cuántos **no tienen sitio
web** — que suelen ser los mejores prospectos).

## Fijo vs celular en México: léelo antes de confiar en esa columna

Desde la homologación de la marcación de agosto de 2019, en México los números
fijos y móviles tienen exactamente el mismo formato de 10 dígitos, y desaparecieron
los prefijos `1`, `044` y `045`. **La información ya no está en el número.** Por eso
`libphonenumber` devuelve `FIXED_LINE_OR_MOBILE` para casi cualquier número mexicano,
y cualquier herramienta que te prometa distinguirlos solo con el número te está
adivinando.

El scraper resuelve esto en cascada:

| Nivel | Método | Confianza |
|---|---|---|
| 1 | Rangos oficiales del IFT | **alta** |
| 2 | Prefijo histórico `+52 1` / `044`, enlace de WhatsApp en el sitio, etiqueta "Cel." / "Conmutador" / "ext." | media |
| 3 | libphonenumber (resuelve 800 y poco más) | baja / `Indeterminado` |

La columna **Confianza tipo** te dice de dónde salió cada clasificación, para que
no trates igual un dato del IFT que una inferencia por WhatsApp.

### Cómo cargar los rangos del IFT

Si en tus resultados ves muchos "Indeterminado", esto es lo que lo arregla.

```bash
python descargar_ift.py                                          # baja el CSV
python gmaps_scraper.py -n "Dentistas" -c "Mérida" --ift ift_numeracion.csv
```

`descargar_ift.py` automatiza el portal del IFT con Playwright, porque el archivo
no está en una URL directa (es una app JSF, hay que pulsar el botón). Si el portal
cambia y el script falla, te deja una captura de pantalla y puedes bajarlo a mano
desde
[sns.ift.org.mx](https://sns.ift.org.mx:8081/sns-frontend/planes-numeracion/descarga-publica.xhtml)
— elige "Numeración Geográfica", guárdalo como `ift_numeracion.csv`.

El parser acepta cualquier CSV con columnas `ZONA`, `SERIE`, `NUMERACION_INICIAL`,
`NUMERACION_FINAL` y `MODALIDAD` (o `TIPO_SERVICIO`), en cualquier orden, y
normaliza `MPP`/`MPF` → Celular y `FPP`/`FIJO` → Fijo. Vuelve a bajarlo cada
pocos meses: el IFT asigna rangos nuevos constantemente.

El parser es flexible con los nombres de columna: acepta `ZONA`, `SERIE`,
`NUMERACION_INICIAL`, `NUMERACION_FINAL` y `MODALIDAD` (o `TIPO_SERVICIO`), y
normaliza `MPP` / `MPF` → Celular, `FPP` / `FIJO` → Fijo. Con el archivo cargado,
la columna pasa de "Indeterminado" a exacta para prácticamente todos los números.

## Emails

Configuraste "solo la home": el scraper abre el sitio del negocio, bloquea
imágenes y CSS (solo necesita el HTML), y saca los `mailto:` más los emails que
encuentre en el texto. Filtra basura típica (Sentry, Wix, `@2x.png`, dominios de
ejemplo) y prefiere el email del propio dominio del negocio, priorizando
`contacto@` > `info@` > `hola@` > `ventas@`.

Tasa de acierto realista: entre 30% y 50% de los negocios **que tienen sitio web**.
Si más adelante quieres subirla, la mejora obvia es probar `/contacto`, `/contact`
y `/about` cuando la home no da nada — está aislado en `scrape_email()`.

### Emails que se descartan a propósito

El filtro tira tres familias de basura que aparecen mucho en sitios de negocios
mexicanos: placeholders de plantilla que nadie reemplazó (`usuario@dominio.com`),
correos de la plataforma y no del negocio (`contacto-mx@doctoralia.com`,
Wix, Shopify, Squarespace), y falsos positivos del regex (`logo@2x.png`,
paquetes npm como `@babel`, DSNs de Sentry). Si ves un email raro colándose,
agrégalo a `EMAIL_BLOCKLIST` al inicio de `gmaps_scraper.py`.

## Reanudación

Cada negocio se escribe a SQLite en cuanto se termina. Si vuelves a correr la
misma búsqueda, los que ya están se saltan y verás
`· N ya estaban en la base, se saltan`. Con `--refresh` los vuelve a bajar.
La deduplicación es por *feature id* de Google (`0x...:0x...`), que es estable
aunque el negocio cambie de nombre, así que funciona también entre búsquedas
distintas que se traslapen.

## Pruebas

```bash
python tests/test_pipeline.py
```

Levanta un Google Maps simulado en localhost que reproduce el DOM real
(`role="feed"`, `h1.DUwDvf`, `div.F7nice`, `button[data-item-id="phone:tel:"]`,
`a[data-item-id="authority"]`, tabla de horarios) y verifica el flujo completo:
scroll, deduplicación, extracción de ficha, extracción de email, clasificación
telefónica, checkpoint, reanudación y el Excel de salida. No toca Google.

## Cuando algo se rompe

**Google cambia sus clases CSS cada pocos meses.** Casi todo lo que se rompa va a
ser un selector. Están todos en el diccionario `SEL` al inicio de
`gmaps_scraper.py`, con varias alternativas por campo. Para arreglarlo: corre con
`--ver`, inspecciona el elemento en el navegador y agrega el selector nuevo **al
inicio** de la lista correspondiente. Los `data-item-id` (dirección, teléfono,
website) son mucho más estables que las clases y por eso los uso donde se puede.

**Una corrida ya no se puede perder.** Cada negocio se escribe a la base al
terminarlo; un error en un negocio se registra y sigue con el siguiente (tras 15
fallos seguidos corta esa búsqueda, porque eso ya es rate limiting y no mala
suerte); y el Excel se escribe en un `finally`, así que sale aunque el proceso
truene a medio camino. Si aun así no aparece, `--solo-export` lo regenera.

**Si empiezan a salir fichas vacías o te aparece un captcha**, es rate limiting.
Sube `--lento 2` o `3`, baja `--max`, y córrelo en tandas en vez de mil negocios
seguidos. El scraper ya manda un user-agent real, quita la bandera `navigator.webdriver`
y mete pausas con jitter, pero no hay magia: la velocidad es lo que te delata.

## Nota legal

Scrapear Google Maps va contra los Términos de Servicio de Google, aunque los
datos de negocios sean públicos. Para uso comercial serio, la Places API es la
vía soportada (de pago, y no da emails). Úsalo con criterio y a tu propio riesgo.
Si vas a usar los emails para prospección en frío, revisa que cumplas la
LFPDPPP y los requisitos de consentimiento y baja.
