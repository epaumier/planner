import calendar
import json
from datetime import date

PERSONS = ["Tedd", "Mathilde", "Alistair", "Yoko"]
MONTHS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
          "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
WEEKDAYS = ["L", "M", "Me", "J", "V", "S", "D"]

BLOCK = 9    # day, weekday, 4 persons, events, info, spacer
MARGIN = 1   # empty first column, like the original sheet
BAND = 34    # header row + 31 day rows + 2 blank rows
MONTHS_PER_BAND = 2


def year_sheet(year):
    cells = {}

    def put(r, c, v, **fmt):
        cell = {"v": v}
        cell.update(fmt)
        cells[f"{r},{c}"] = cell

    for band in range(6):
        top = band * BAND
        for j in range(MONTHS_PER_BAND):
            m = band * MONTHS_PER_BAND + j + 1
            if m > 12:
                break
            c0 = MARGIN + j * BLOCK
            put(top, c0, MONTHS[m - 1], b=True, bg="#fde047")
            for p, name in enumerate(PERSONS):
                put(top, c0 + 2 + p, name, b=True, bg="#fde047")
            put(top, c0 + 6, "Événements", b=True, bg="#fde047")
            put(top, c0 + 7, "Info supp", b=True, bg="#fde047")
            ndays = calendar.monthrange(year, m)[1]
            for d in range(1, 32):
                r = top + d
                if d <= ndays:
                    put(r, c0, str(d))
                    put(r, c0 + 1, WEEKDAYS[date(year, m, d).weekday()])
    return {
        "name": str(year),
        "rows": 6 * BAND,
        "cols": MARGIN + MONTHS_PER_BAND * BLOCK,
        "cells": cells,
        "zebraColor": "#f8fafc",
        "zebraPeriod": BAND,
    }


workbook = {"sheets": [year_sheet(2026), year_sheet(2027)]}

with open("seed-calendar.js", "w", encoding="utf-8") as f:
    f.write("window.SEED_CALENDAR = ")
    f.write(json.dumps(workbook, ensure_ascii=False))
    f.write(";\n")

print("sheets:", [(s["name"], s["rows"], s["cols"], len(s["cells"])) for s in workbook["sheets"]])
print("2026-01-01 weekday:", WEEKDAYS[date(2026, 1, 1).weekday()], "(expect J)")
