# Análisis de arquitectura — Scraper de Google Maps

> Documento hecho como parte de mi formación en Ingeniería en Software.
> Aplico los temas vistos en clase (SOLID, atributos de calidad,
> requerimientos y patrones de diseño) a este proyecto. Solo marco lo que
> encontré de verdad en el código; si algo no está, lo digo como
> **"No identificado"** en lugar de inventarlo.

---

## 1. ¿Qué hace el proyecto?

Es una herramienta de línea de comandos en **Python + Playwright** que busca
negocios en Google Maps (por ejemplo "Dentistas en Mérida"), recorre los
resultados y extrae datos de cada ficha: nombre, categoría, teléfono,
dirección, sitio web, horarios, calificación y coordenadas. Si el negocio
tiene sitio web, también intenta encontrar su correo y su WhatsApp.

Además **clasifica cada teléfono como fijo o celular** (en México ya no se
distinguen por el número, así que usa una cascada de métodos con un nivel de
confianza). Guarda todo en una base **SQLite** conforme avanza, para poder
retomar si se interrumpe, y al final exporta a **Excel**.

**Tecnologías:** Python 3, Playwright (navegador), SQLite, openpyxl (Excel),
phonenumbers (validar teléfonos).

---

## 2. Arquitectura sencilla

No usa un framework: son **módulos**, cada uno con un trabajo. `gmaps_scraper.py`
es el orquestador y los otros dos módulos lo apoyan.

```
  Usuario (terminal)
        │  python gmaps_scraper.py -n "Dentistas" -c "Mérida, Yucatán"
        ▼
 ┌──────────────────────────────────────────────┐
 │ gmaps_scraper.py   (orquestador + extracción)│
 │  1. abre Maps y busca                        │
 │  2. hace scroll → lista de tarjetas          │
 │  3. entra a cada ficha → extrae datos        │
 │  4. (opcional) visita su sitio → email/WhatsApp
 └───────┬───────────────────┬──────────────────┘
         │ usa               │ usa
         ▼                   ▼
 ┌───────────────┐   ┌─────────────────────┐
 │  phone_mx.py  │   │     storage.py      │
 │ clasifica tel │   │ Store → gmaps.db    │
 │ fijo / celular│   │ export_excel → .xlsx│
 └───────▲───────┘   └─────────▲───────────┘
         │ rangos IFT          │ lee
 ┌───────┴────────┐    ┌───────┴───────┐
 │descargar_ift.py│    │   ver_db.py   │
 │ (baja el CSV)  │    │ (consulta BD) │
 └────────────────┘    └───────────────┘

 Todo el acceso a internet pasa por Playwright (Chromium)
 → Google Maps y los sitios web de los negocios
```

### Flujo de una ejecución

1. `main()` lee las opciones de la terminal y arma un objeto `Config`.
2. `main_async()` carga (si se indicó) los rangos del IFT y abre la base con `Store`.
3. `_scrapear()` lanza Chromium y, por cada búsqueda, llama a `run_search()`.
4. `run_search()`: `open_search()` → `scroll_feed()` (hace scroll hasta el final
   o hasta el máximo) → filtra los negocios que ya están en la base.
5. Por cada negocio nuevo: `scrape_detail()` extrae la ficha →
   `scrape_email()` visita el sitio (si hay) → `phone_mx.classify()` clasifica
   el teléfono → `store.upsert()` guarda **ese negocio de inmediato**.
6. En un bloque `finally`, `export_excel()` genera el Excel **aunque algo haya
   fallado a la mitad**.

---

## 3. Principios SOLID

| Principio | Evaluación | Evidencia en el código |
|---|---|---|
| **S** — Responsabilidad única | **Parcial** | Bien a nivel de módulos: `phone_mx.py` solo trata teléfonos, `storage.py` solo persistencia y exportación, `ver_db.py` solo consulta. Dentro de `phone_mx.py`, la clase `IFTRanges` (leer/buscar rangos) está separada de la función `classify()`. **Mal:** `gmaps_scraper.py` (~770 líneas) mezcla selectores, extracción, navegación, email, orquestación y lectura de argumentos. |
| **O** — Abierto/Cerrado | **Parcial** | Para adaptarse a cambios de Google solo se agrega un selector a la lista del diccionario `SEL`, sin tocar la lógica; igual con `EMAIL_BLOCKLIST` para nuevos correos basura. `Store.upsert()` arma el `INSERT` según las columnas existentes. Pero agregar un nuevo nivel a la cascada de `classify()` obliga a editar la función. |
| **L** — Sustitución de Liskov | **No identificado** | No hay herencia propia (solo `@dataclass`), así que no hay subclases que sustituir. |
| **I** — Segregación de interfaces | **No identificado** | No hay interfaces ni clases abstractas. Es un script pequeño sin contratos explícitos. |
| **D** — Inversión de dependencias | **Parcial** | Bien: `run_search(context, store, cfg)` **recibe** el navegador, la base y la configuración por parámetro en vez de crearlos adentro; `Store(path)` acepta cualquier ruta (los tests usan una base temporal). Mal: `gmaps_scraper.py` depende directo del módulo `phone_mx` y de un estado global (`_IFT`), y no hay abstracciones intermedias. |

---

## 4. Atributos de calidad

### Pensando en el usuario
- **Confiabilidad:** cada negocio se guarda en cuanto termina; si se cae, se
  pierde a lo mucho uno. Un error en un negocio no tumba la corrida, y al
  acumular 15 fallos en una búsqueda la corta (probable bloqueo). El Excel se
  escribe en un `finally`.
- **Usabilidad:** opciones claras en español (`--ver`, `--max`, `--lento`,
  `--sin-emails`, `--solo-export`), README detallado y un Excel con filtros,
  hipervínculos y hoja **Resumen**.
- **Rendimiento (con compromiso):** pausas aleatorias para no ser bloqueado
  (`--lento` las multiplica); al visitar sitios web bloquea imágenes, fuentes
  y CSS porque solo necesita el HTML.
- **Exactitud y transparencia de datos:** cada tipo de teléfono incluye su
  **fuente y nivel de confianza** (alta/media/baja), en vez de aparentar certeza.

### Pensando en el programador
- **Testeabilidad:** es lo más fuerte del proyecto. `tests/test_pipeline.py`
  levanta un servidor local que simula Google Maps y verifica el flujo
  completo sin tocar Google. Además hay funciones puras fáciles de probar
  (`parse_rating_block`, `coords_from_url`, `pick_best_email`, `classify`).
- **Mantenibilidad / modificabilidad:** selectores centralizados en `SEL` con
  varias alternativas; el README explica qué hacer cuando Google cambia su diseño.
- **Robustez:** ante cambios de la página, los selectores se prueban en
  cascada y se prefieren atributos estables (`data-item-id`).
- **Portabilidad:** funciona donde corra Playwright; depende de tres librerías
  (`requirements.txt`).
- **Documentación:** docstrings en español y un README que explica incluso las
  decisiones (por qué fijo/celular no se puede saber solo por el número).

---

## 5. Requerimientos

### Funcionales (qué hace el sistema)
- **RF1.** Buscar negocios por giro y ciudad, de forma individual o por lote desde un CSV (`--batch`).
- **RF2.** Recorrer la lista de resultados haciendo scroll (`scroll_feed`).
- **RF3.** Extraer de cada ficha: nombre, categoría, calificación, reseñas, dirección, sitio web, teléfono, horarios, coordenadas, plus code, precio y si cerró permanentemente (`scrape_detail`).
- **RF4.** Extraer correo y WhatsApp desde el sitio web del negocio (`scrape_email`).
- **RF5.** Clasificar el teléfono como fijo, celular o indeterminado, con su confianza (`phone_mx.classify`).
- **RF6.** Guardar cada negocio en SQLite al terminarlo, sin duplicados (`Store.upsert`).
- **RF7.** Reanudar una búsqueda saltando los negocios ya guardados (`--refresh` para repetirlos).
- **RF8.** Exportar a Excel con hoja de resumen (`export_excel`, `--solo-export`).
- **RF9.** Consultar la base y volcarla a CSV (`ver_db.py`).
- **RF10.** Descargar los rangos de numeración del IFT (`descargar_ift.py`).

### No funcionales (cómo lo hace)
- **RNF1. Fiabilidad:** guardado incremental, errores aislados por negocio, límite de 15 fallos por búsqueda.
- **RNF2. Rendimiento controlado:** pausas con variación aleatoria y opción `--lento`.
- **RNF3. Robustez ante cambios externos:** selectores en cascada.
- **RNF4. Usabilidad:** interfaz de terminal en español y salida legible (Excel con filtros).
- **RNF5. Mantenibilidad y testeabilidad:** módulos separados y pruebas con un Maps simulado.
- **RNF6. Compatibilidad con Excel:** limpia caracteres que Excel rechaza, recorta textos largos, y antepone `'` a textos que empiezan con `=`, `+` o `@` para que no se interpreten como fórmulas (`_limpiar`).
- **RNF7. Legalidad y ética:** el README advierte que scrapear Maps va contra los Términos de Servicio de Google y menciona la LFPDPPP para el uso de correos.

---

## 6. Patrones de diseño

> Prefiero marcar pocos patrones y honestos que forzar los 23 de GoF.

### Creacionales
- **Singleton — aproximado:** en `phone_mx.py`, `_IFT = IFTRanges()` es una
  única instancia a nivel de módulo, compartida por todo el programa mediante
  `load_ift_ranges()` y `classify()`. Es el "singleton de Python" (un módulo se
  importa una sola vez), pero no impone la unicidad con la clase.
- **Factory, Builder, Prototype:** **No identificados.**

### Estructurales
- **Facade (Fachada):** `phone_mx.classify()` esconde toda la cascada (rangos
  IFT, pistas del contexto, libphonenumber) detrás de una sola llamada. `Store`
  también esconde `sqlite3` detrás de métodos simples (`upsert`, `fetch`, `count`).
- **Adapter — aproximado:** `IFTRanges.load()` con `_pick()` y
  `_normalize_modalidad()` convierte un CSV del IFT cuyos nombres de columna
  varían (`MPP`, `MPF`, `FPP`…) a un formato interno único (rangos ordenados).
- **Repository (no es de GoF):** `Store` centraliza el acceso a datos, separando
  la lógica del scraper de los detalles de SQLite.

### De comportamiento
- **Chain of Responsibility — aproximado:** `classify()` prueba tres niveles en
  orden (rangos IFT → pistas del contexto → libphonenumber); el primero que
  resuelve el caso responde y los demás no se ejecutan. Está hecho con `if`
  dentro de una función, no con clases encadenadas. Algo parecido hacen
  `first_text()` y `first_attr()` al probar selectores en orden.
- **Strategy, Observer, State, Template Method, Command:** **No identificados.**
- **Circuit breaker (no es de GoF), aproximado:** `run_search()` corta la
  búsqueda al llegar a 15 errores para no insistir contra un bloqueo.

---

## 7. Qué mejoraría (lo que aprendí al analizarlo)

1. Dividir `gmaps_scraper.py` en módulos (navegación, extracción de fichas, extracción de correo, línea de comandos) para cumplir mejor **S**.
2. Reemplazar el estado global `_IFT` por una instancia que se pase por parámetro (**D**) y facilite las pruebas.
3. Convertir la cascada de `classify()` en una lista de "reglas" para poder agregar niveles sin editar la función (**O**).
4. En `pick_best_email()`, `domain.lower().lstrip("www.")` quita *cualquier* letra `w` o punto del inicio, no solo el prefijo `www.`; usar `removeprefix("www.")`.
5. En `descargar_ift.py`, `ignore_https_errors=True` desactiva la verificación del certificado; conviene limitarlo y documentar por qué.
6. En `run_search()`, el contador `errores` nunca se reinicia tras un éxito, así que cuenta fallos acumulados y no "seguidos" como dice el README; reiniciarlo en cada éxito o ajustar el texto.
7. Quitar el párrafo repetido del README sobre las columnas del IFT y reemplazar en las pruebas los datos de contacto por ejemplos ficticios.
