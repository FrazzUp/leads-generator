"""
Clasificación de números telefónicos mexicanos: FIJO vs CELULAR.

CONTEXTO IMPORTANTE
-------------------
Desde la homologación de la marcación en México (agosto 2019), los números fijos
y móviles comparten el mismo formato de 10 dígitos y ya NO se distinguen por el
prefijo 1 / 044 / 045. Por eso libphonenumber devuelve FIXED_LINE_OR_MOBILE para
casi cualquier número mexicano: la información simplemente no está en el número.

La única fuente autoritativa es el Plan Técnico Fundamental de Numeración del
IFT, que publica los rangos asignados con su tipo de servicio (fijo / móvil).
Este módulo usa un esquema en cascada:

  1. Rangos IFT (exacto)      -> confianza ALTA
  2. Señales de la página     -> confianza MEDIA   (WhatsApp, "Cel.", +52 1 ...)
  3. libphonenumber           -> confianza BAJA / INDETERMINADO

Ver README para cómo descargar el CSV del IFT.
"""

from __future__ import annotations

import csv
import os
import re
from bisect import bisect_right
from dataclasses import dataclass

import phonenumbers
from phonenumbers import PhoneNumberType, number_type

DEFAULT_REGION = "MX"

FIJO = "Fijo"
CELULAR = "Celular"
INDETERMINADO = "Indeterminado"

ALTA = "alta"
MEDIA = "media"
BAJA = "baja"


@dataclass
class PhoneInfo:
    raw: str | None = None
    e164: str | None = None
    national: str | None = None
    lada: str | None = None
    tipo: str = INDETERMINADO
    confianza: str = BAJA
    fuente: str = ""
    valido: bool = False

    def as_dict(self) -> dict:
        if not (self.raw or "").strip():
            # Sin teléfono no hay nada que clasificar: columnas vacías, no
            # un "Indeterminado" que ensucie los filtros del Excel.
            return {
                "telefono": None, "telefono_e164": None, "lada": None,
                "tipo_telefono": None, "tipo_confianza": None, "tipo_fuente": None,
            }
        return {
            "telefono": self.national or self.raw or None,
            "telefono_e164": self.e164,
            "lada": self.lada,
            "tipo_telefono": self.tipo,
            "tipo_confianza": self.confianza,
            "tipo_fuente": self.fuente,
        }


# --------------------------------------------------------------------------
# 1. Rangos IFT
# --------------------------------------------------------------------------

class IFTRanges:
    """
    Índice de búsqueda sobre el archivo de numeración del IFT.

    El CSV del IFT trae (los nombres de columna varían entre publicaciones):
        ZONA, SERIE, NUMERACION_INICIAL, NUMERACION_FINAL, OCUPACION,
        MODALIDAD, RAZON_SOCIAL, ...

    MODALIDAD suele ser: 'FIJO', 'MOVIL', 'MPP' (móvil prepago),
    'MPF' (móvil pospago), 'FPP' (fijo). Normalizamos a FIJO / CELULAR.
    """

    def __init__(self) -> None:
        self._starts: list[int] = []
        self._ends: list[int] = []
        self._types: list[str] = []
        self.loaded = False
        self.count = 0

    @staticmethod
    def _normalize_modalidad(value: str) -> str | None:
        v = (value or "").strip().upper()
        if not v:
            return None
        if v.startswith("MPP") or v.startswith("MPF") or "MOV" in v or v == "M":
            return CELULAR
        if v.startswith("FPP") or "FIJ" in v or v == "F":
            return FIJO
        return None

    @staticmethod
    def _pick(header: list[str], *candidates: str) -> str | None:
        norm = {re.sub(r"[^A-Z]", "", h.upper()): h for h in header}
        for cand in candidates:
            key = re.sub(r"[^A-Z]", "", cand.upper())
            if key in norm:
                return norm[key]
        return None

    def load(self, path: str) -> bool:
        if not path or not os.path.exists(path):
            return False

        rows: list[tuple[int, int, str]] = []
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as fh:
            sample = fh.read(8192)
            fh.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel
            reader = csv.DictReader(fh, dialect=dialect)
            header = reader.fieldnames or []

            c_zona = self._pick(header, "ZONA", "CLAVE", "LADA")
            c_serie = self._pick(header, "SERIE")
            c_ini = self._pick(header, "NUMERACION_INICIAL", "NUMERACIONINICIAL", "INICIO")
            c_fin = self._pick(header, "NUMERACION_FINAL", "NUMERACIONFINAL", "FIN")
            c_mod = self._pick(header, "MODALIDAD", "TIPO_SERVICIO", "TIPOSERVICIO", "TIPO")

            if not (c_zona and c_serie and c_ini and c_fin and c_mod):
                return False

            for row in reader:
                tipo = self._normalize_modalidad(row.get(c_mod, ""))
                if not tipo:
                    continue
                try:
                    zona = re.sub(r"\D", "", row[c_zona] or "")
                    serie = re.sub(r"\D", "", row[c_serie] or "")
                    ini = int(re.sub(r"\D", "", row[c_ini] or "0"))
                    fin = int(re.sub(r"\D", "", row[c_fin] or "0"))
                except (KeyError, ValueError):
                    continue
                if not zona or not serie:
                    continue
                # El número nacional completo es ZONA + SERIE + últimos 4 dígitos.
                prefix = zona + serie
                width = 10 - len(prefix)
                if width < 0:
                    continue
                start = int(prefix + str(ini).zfill(width)[-width:]) if width else int(prefix)
                end = int(prefix + str(fin).zfill(width)[-width:]) if width else int(prefix)
                if end < start:
                    start, end = end, start
                rows.append((start, end, tipo))

        if not rows:
            return False

        rows.sort(key=lambda r: r[0])
        self._starts = [r[0] for r in rows]
        self._ends = [r[1] for r in rows]
        self._types = [r[2] for r in rows]
        self.count = len(rows)
        self.loaded = True
        return True

    def lookup(self, national_digits: str) -> str | None:
        if not self.loaded or len(national_digits) != 10:
            return None
        n = int(national_digits)
        idx = bisect_right(self._starts, n) - 1
        # Los rangos pueden solaparse; revisamos algunos hacia atrás.
        for i in range(idx, max(-1, idx - 6), -1):
            if i < 0:
                break
            if self._starts[i] <= n <= self._ends[i]:
                return self._types[i]
        return None


_IFT = IFTRanges()


def load_ift_ranges(path: str | None) -> bool:
    """Carga el CSV de numeración del IFT. Devuelve True si quedó disponible."""
    if not path:
        return False
    return _IFT.load(path)


def ift_available() -> bool:
    return _IFT.loaded


def ift_count() -> int:
    return _IFT.count


# --------------------------------------------------------------------------
# 2. Señales de contexto
# --------------------------------------------------------------------------

_CEL_HINTS = re.compile(
    r"(whats\s?app|wa\.me|api\.whatsapp|\bcel(?:ular)?\b|\bmóvil\b|\bmovil\b|\bmob(?:ile)?\b)",
    re.IGNORECASE,
)
_FIJO_HINTS = re.compile(
    r"(\btel(?:éfono|efono)?\s*fijo\b|\bconmutador\b|\bext\.?\s*\d|\bextensi[óo]n\b|\boficina\b)",
    re.IGNORECASE,
)

# LADAs de 2 dígitos en México (el resto son de 3).
_LADA2 = {"55", "56", "33", "81"}


def _split_lada(national: str) -> str:
    if len(national) != 10:
        return ""
    return national[:2] if national[:2] in _LADA2 else national[:3]


# --------------------------------------------------------------------------
# 3. API principal
# --------------------------------------------------------------------------

def classify(raw: str | None, context: str = "", region: str = DEFAULT_REGION) -> PhoneInfo:
    """
    Normaliza y clasifica un teléfono.

    `context` es texto libre asociado al negocio (HTML del sitio, etiqueta del
    botón en Maps, enlaces encontrados). Se usa solo como señal secundaria.
    """
    info = PhoneInfo(raw=raw)
    if not raw or not str(raw).strip():
        return info

    text = str(raw)

    # El prefijo histórico "+52 1 ..." / "044" / "045" todavía aparece mucho en
    # directorios y sitios web; cuando está, es evidencia fuerte de móvil.
    legacy_mobile = bool(re.search(r"(\+?52\s*1\s*\d)|(^0?4[45]\s*\d)", text))

    try:
        parsed = phonenumbers.parse(text, region)
    except phonenumbers.NumberParseException:
        parsed = None

    # "+52 1 33 ..." y "044/045" son 11 dígitos: no validan. Se reintenta sin el 1.
    if legacy_mobile and (parsed is None or not phonenumbers.is_valid_number(parsed)):
        stripped = re.sub(r"\D", "", text)
        stripped = re.sub(r"^(52)?1(?=\d{10}$)", r"\1", stripped)
        stripped = re.sub(r"^0?4[45](?=\d{10}$)", "", stripped)
        try:
            retry = phonenumbers.parse(stripped, region)
            if phonenumbers.is_valid_number(retry):
                parsed = retry
        except phonenumbers.NumberParseException:
            pass

    if parsed is None:
        return info

    if not phonenumbers.is_valid_number(parsed):
        # Guardamos lo que se pueda aunque no valide (ej. números con extensión rara).
        digits = re.sub(r"\D", "", text)[-10:]
        if len(digits) == 10:
            info.national = f"{digits[:_len_lada(digits)]} {digits[_len_lada(digits):]}"
            info.lada = _split_lada(digits)
        return info

    info.valido = True
    info.e164 = phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.E164)
    info.national = phonenumbers.format_number(
        parsed, phonenumbers.PhoneNumberFormat.NATIONAL
    )
    national_digits = str(parsed.national_number)
    info.lada = _split_lada(national_digits)

    # --- Nivel 1: rangos IFT -------------------------------------------------
    if parsed.country_code == 52:
        hit = _IFT.lookup(national_digits)
        if hit:
            info.tipo = hit
            info.confianza = ALTA
            info.fuente = "IFT"
            return info

    # --- Nivel 2: señales de contexto ---------------------------------------
    ctx = f"{text}\n{context or ''}"
    if legacy_mobile:
        info.tipo = CELULAR
        info.confianza = MEDIA
        info.fuente = "prefijo 1/044 histórico"
        return info
    if _CEL_HINTS.search(ctx):
        info.tipo = CELULAR
        info.confianza = MEDIA
        info.fuente = "WhatsApp/etiqueta 'cel'"
        return info
    if _FIJO_HINTS.search(ctx):
        info.tipo = FIJO
        info.confianza = MEDIA
        info.fuente = "etiqueta 'conmutador/ext'"
        return info

    # --- Nivel 3: libphonenumber --------------------------------------------
    t = number_type(parsed)
    if t == PhoneNumberType.MOBILE:
        info.tipo, info.confianza, info.fuente = CELULAR, ALTA, "libphonenumber"
    elif t == PhoneNumberType.FIXED_LINE:
        info.tipo, info.confianza, info.fuente = FIJO, ALTA, "libphonenumber"
    elif t == PhoneNumberType.TOLL_FREE:
        info.tipo, info.confianza, info.fuente = "Sin costo (800)", ALTA, "libphonenumber"
    else:
        info.tipo = INDETERMINADO
        info.confianza = BAJA
        info.fuente = (
            "MX unificó fijo/móvil en 2019; carga el CSV del IFT para resolverlo"
        )
    return info


def _len_lada(national: str) -> int:
    return 2 if national[:2] in _LADA2 else 3


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        ok = load_ift_ranges(sys.argv[1])
        print(f"IFT cargado: {ok} ({ift_count()} rangos)")
        samples = sys.argv[2:]
    else:
        samples = sys.argv[1:]

    for s in samples or ["+52 55 1234 5678", "+52 1 33 1234 5678", "998 123 4567", "800 123 4567"]:
        print(s, "->", classify(s).as_dict())
