#!/usr/bin/env python3
"""
Mira qué hay dentro de gmaps.db sin instalar nada.

    python ver_db.py                    # resumen
    python ver_db.py --lista            # todos los negocios, una línea cada uno
    python ver_db.py --buscar dental    # filtra por nombre
    python ver_db.py --csv salida.csv   # vuelca todo a CSV (UTF-8 con BOM, abre bien en Excel)

Para exportar a Excel usa el scraper:
    python gmaps_scraper.py --solo-export -o resultados.xlsx
"""

from __future__ import annotations

import argparse
import csv
import sqlite3
import sys
from pathlib import Path


def main() -> int:
    ap = argparse.ArgumentParser(description="Inspecciona la base del scraper")
    ap.add_argument("--db", default="gmaps.db")
    ap.add_argument("--lista", action="store_true")
    ap.add_argument("--buscar")
    ap.add_argument("--csv")
    a = ap.parse_args()

    ruta = Path(a.db)
    if not ruta.exists():
        print(f"! No existe {ruta.resolve()}")
        return 1

    con = sqlite3.connect(ruta)
    con.row_factory = sqlite3.Row

    total = con.execute("SELECT COUNT(*) FROM negocios").fetchone()[0]
    print(f"Base: {ruta.resolve()}  ({ruta.stat().st_size / 1024:.0f} KB)")
    print(f"Negocios guardados: {total}\n")

    print("Por búsqueda:")
    for r in con.execute(
        "SELECT query, ubicacion, COUNT(*) n FROM negocios "
        "GROUP BY query, ubicacion ORDER BY n DESC"
    ):
        print(f"  {r['n']:>5}  {r['query']} en {r['ubicacion']}")

    print("\nCobertura de datos:")
    for etiqueta, cond in [
        ("con teléfono", "telefono IS NOT NULL AND telefono != ''"),
        ("con sitio web", "website IS NOT NULL AND website != ''"),
        ("con email", "email IS NOT NULL AND email != ''"),
        ("SIN sitio web (prospectos)", "website IS NULL OR website = ''"),
    ]:
        n = con.execute(f"SELECT COUNT(*) FROM negocios WHERE {cond}").fetchone()[0]
        pct = f"{n / total * 100:.0f}%" if total else "—"
        print(f"  {n:>5}  ({pct:>4})  {etiqueta}")

    print("\nTipo de teléfono:")
    for r in con.execute(
        "SELECT COALESCE(tipo_telefono,'(sin teléfono)') t, COUNT(*) n "
        "FROM negocios GROUP BY t ORDER BY n DESC"
    ):
        print(f"  {r['n']:>5}  {r['t']}")

    if a.buscar or a.lista:
        sql = "SELECT nombre, telefono, tipo_telefono, email, rating, num_resenas FROM negocios"
        params = ()
        if a.buscar:
            sql += " WHERE nombre LIKE ?"
            params = (f"%{a.buscar}%",)
        sql += " ORDER BY num_resenas DESC"
        print()
        for r in con.execute(sql, params):
            print(
                f"  {(r['nombre'] or '')[:40]:<40} {(r['telefono'] or '—'):<15} "
                f"{(r['tipo_telefono'] or '—'):<14} {(r['email'] or '—'):<32} "
                f"{r['rating'] or '—'} ({r['num_resenas'] or 0})"
            )

    if a.csv:
        filas = con.execute("SELECT * FROM negocios").fetchall()
        # utf-8-sig = UTF-8 con BOM: es lo que hace que Excel respete los acentos.
        with open(a.csv, "w", newline="", encoding="utf-8-sig") as fh:
            if filas:
                w = csv.DictWriter(fh, fieldnames=filas[0].keys())
                w.writeheader()
                w.writerows(dict(r) for r in filas)
        print(f"\n✓ {len(filas)} filas escritas en {Path(a.csv).resolve()}")

    con.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
