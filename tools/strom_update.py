#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
strom_update.py – Datenbasis (Excel) und Dashboard zum Stromverbrauch pflegen
============================================================================

Ordner "Auswertung und Ablesungen":
  Stromzaehlerwerte.xlsx     Datenbasis: Ablesungen, Auswertung, Monatswerte, Stammdaten, Info
  Dashboard/index.html       Dashboard (statisch, liest messwerte.js)
  Dashboard/messwerte.js     Daten fürs Dashboard – wird von diesem Skript aus der Excel erzeugt
  github-token.txt           GitHub-Token (fbw159)
  _claude/                   dieses Skript, WISSEN.md, backup/

Die Excel ist die Datenbasis. Das Skript liest die Eingaben (Ablesungen Spalten A–E und die
Stammdaten), ergänzt ggf. eine neue Ablesung und baut die Arbeitsmappe mit allen Formeln,
Formaten und dem Diagramm neu auf. Danach wird messwerte.js geschrieben und – mit --push –
das Dashboard auf GitHub Pages veröffentlicht.

Aufrufe am PC (Python 3.8+, openpyxl; LibreOffice optional für die Neuberechnung):
  python3 strom_update.py                     # Abgleich mit GitHub (Handy-Ablesungen übernehmen) + Excel neu
  python3 strom_update.py --neu 2026-12-01 60512,3 --quelle "Foto IMG_20261201_101500.jpg (10:15 Uhr)" --push
  python3 strom_update.py --init daten.json   # Erstaufbau aus JSON

Aufruf in der Cloud (Handy, PC aus) – arbeitet nur mit messwerte.js im geklonten Repo und pusht per git:
  python3 tools/strom_update.py --cloud REPO_ORDNER --token-datei TOKEN.txt --neu 2026-12-01 60512,3 --quelle "Foto (Handy)"

Abgleich: GitHub ist führend für index.html und dieses Skript (tools/strom_update.py), die Excel für alles
andere. Vor jedem Lauf am PC werden Ablesungen, die nur online stehen (vom Handy), in die Excel übernommen.
"""
from __future__ import annotations

import argparse
import base64
import datetime as dt
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path

try:
    import openpyxl
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.properties import PageSetupProperties
except ImportError:  # pragma: no cover
    sys.exit("openpyxl fehlt – bitte installieren: pip install openpyxl")

# --------------------------------------------------------------------------------------
# Konfiguration
# --------------------------------------------------------------------------------------
SKRIPT_DIR = Path(__file__).resolve().parent
BASE = SKRIPT_DIR.parent
EXCEL_NAME = "Stromzaehlerwerte.xlsx"
GH_OWNER = "fbw159"
GH_REPO = "stromverbrauch"
PAGES_URL = f"https://{GH_OWNER}.github.io/{GH_REPO}/"
REPO_BESCHREIBUNG = ("Stromverbrauch unserer Wohnung: Zählerstände, Verbrauch, Kosten nach Vertrag, "
                     "Abschlag-Check und Vergleich mit der alten Wohnung (Dashboard)")
MAX_ZEILEN = 2000            # Formelbereiche in der Excel reichen bis zu dieser Zeile
BACKUPS_BEHALTEN = 8

FONT = "Arial"
C_BRAND = "1F3A5F"
C_HDR_IN = "2A5C9A"
C_HDR_CALC = "4A4A4A"
C_IN_TXT = "0000FF"
C_MUTED = "7F7F7F"
C_LINK = "0563C1"
C_GRID = "D9D9D9"
FARBEN = {  # Dashboard-Farbe -> (Excel-Tönung, Excel-Vollfarbe)
    "blau": ("E6F0FB", "2A78D6"),
    "orange": ("FCEBE2", "EB6834"),
    "türkis": ("E2F5EE", "1BAF7A"),
    "gelb": ("FDF3D9", "EDA100"),
    "violett": ("ECEAFB", "6250D6"),
}

F_DATUM = "DD.MM.YYYY"
F_MONAT = "MMM YYYY"
F_KWH1 = "#,##0.0"
F_KWH0 = "#,##0"
F_KWHD = "0.00"
F_CT = "0.00"
F_EUR = '#,##0.00 "€"'
F_EUR_STRICH = '#,##0.00 "€";-#,##0.00 "€";"–"'
F_EUR_TAG = '0.0000 "€"'
F_PROZ = '+0.0%;-0.0%;0.0%'
F_PROZ0 = '0%'
F_SALDO = '+#,##0.00 "€";-#,##0.00 "€";0.00 "€"'

# Zeilen-/Spaltenlayout der Arbeitsmappe (vom Leser und vom Erzeuger gemeinsam genutzt)
ABL_HDR = 6          # Kopfzeile im Blatt "Ablesungen"
ABL_ERSTE = 7        # erste Datenzeile


# --------------------------------------------------------------------------------------
# Hilfsfunktionen
# --------------------------------------------------------------------------------------
def to_date(v) -> dt.date | None:
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Datum nicht lesbar: {v!r}")


def to_float(v) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if "," in s:                       # deutsches Format: 60.512,3
        s = s.replace(".", "").replace(",", ".")
    return float(s)


def de_zahl(x: float, nachkomma: int = 1) -> str:
    s = f"{x:,.{nachkomma}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def de_datum(d: dt.date) -> str:
    return d.strftime("%d.%m.%Y")


def add_months(d: dt.date, n: int) -> dt.date:
    m = d.month - 1 + n
    y = d.year + m // 12
    m = m % 12 + 1
    last = [31, 29 if (y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)) else 28,
            31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
    return dt.date(y, m, min(d.day, last))


def iso(d: dt.date | None) -> str | None:
    return d.isoformat() if d else None


def jetzt() -> dt.datetime:
    """Aktuelle Zeit in Deutschland (die Linux-VMs laufen auf UTC)."""
    try:
        from zoneinfo import ZoneInfo
        return dt.datetime.now(ZoneInfo("Europe/Berlin")).replace(tzinfo=None)
    except Exception:  # noqa: BLE001
        return dt.datetime.now()


def sha(b: bytes) -> str:
    import hashlib
    return hashlib.sha256(b).hexdigest()


# --------------------------------------------------------------------------------------
# Daten laden / prüfen
# --------------------------------------------------------------------------------------
def daten_aus_json(pfad: Path) -> dict:
    roh = json.loads(Path(pfad).read_text(encoding="utf-8"))
    d = {"wohnungen": roh["wohnungen"], "ablesungen": [], "vertraege": [], "vergleich": []}
    for a in roh["ablesungen"]:
        d["ablesungen"].append({"datum": to_date(a["datum"]), "stand": float(a["stand"]), "wohnung": a["wohnung"],
                                "quelle": a.get("quelle", ""), "bemerkung": a.get("bemerkung", "")})
    for key in ("vertraege", "vergleich"):
        for v in roh.get(key, []):
            v = dict(v)
            for f in ("ab", "bis", "preisgarantie_bis"):
                v[f] = to_date(v.get(f))
            d[key].append(v)
    return d


def _abschnitt(ws, titel: str):
    """Sucht im Blatt eine Abschnittsüberschrift in Spalte A und liefert (Kopfzeile, {Spaltenname: Spalte})."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().upper().startswith(titel):
            hdr = r + 1
            cols = {}
            for c in range(1, ws.max_column + 1):
                h = ws.cell(hdr, c).value
                if isinstance(h, str):
                    cols[h.split("\n")[0].strip()] = c
            return hdr, cols
    raise ValueError(f"Abschnitt {titel!r} im Blatt {ws.title!r} nicht gefunden")


def daten_aus_excel(pfad: Path) -> dict:
    wb = openpyxl.load_workbook(pfad, data_only=False)
    ws = wb["Ablesungen"]
    hdr = None
    for r in range(1, 30):
        if ws.cell(r, 1).value == "Datum":
            hdr = r
            break
    if hdr is None:
        raise ValueError("Kopfzeile 'Datum' im Blatt 'Ablesungen' nicht gefunden")
    abl = []
    r = hdr + 1
    while ws.cell(r, 1).value not in (None, ""):
        abl.append({"datum": to_date(ws.cell(r, 1).value), "stand": to_float(ws.cell(r, 2).value),
                    "wohnung": str(ws.cell(r, 3).value or "").strip(),
                    "quelle": str(ws.cell(r, 4).value or "").strip(),
                    "bemerkung": str(ws.cell(r, 5).value or "").strip()})
        r += 1

    st = wb["Stammdaten"]

    def lese_tarife(titel):
        h, cols = _abschnitt(st, titel)
        out = []
        r = h + 1
        while st.cell(r, 1).value not in (None, ""):
            g = lambda name: st.cell(r, cols[name]).value if name in cols else None  # noqa: E731
            out.append({
                "bezeichnung": str(g("Bezeichnung") or "").strip(),
                "anbieter": str(g("Anbieter") or "").strip(),
                "tarif": str(g("Tarif") or "").strip(),
                "wohnung": str(g("Wohnung") or "").strip(),
                "ab": to_date(g("gültig ab")),
                "bis": to_date(g("gültig bis")),
                "ap": to_float(g("Arbeitspreis")),
                "gp": to_float(g("Grundpreis lt. Vertrag")),
                "gp_einheit": str(g("Einheit") or "€/Jahr").strip(),
                "abschlag": to_float(g("Abschlag")),
                "abschlag_faellig": str(g("Abschlag fällig") or "").strip(),
                "preisgarantie_bis": to_date(g("Preisgarantie bis")),
                "laufzeit": str(g("Laufzeit & Kündigung") or "").strip(),
                "status": str(g("Status") or "").strip(),
                "quelle": str(g("Quelle") or "").strip(),
            })
            r += 1
        return out

    vertraege = lese_tarife("VERTRÄGE")
    vergleich = lese_tarife("VERGLEICHSTARIFE")
    h, cols = _abschnitt(st, "WOHNUNGEN")
    wohnungen = []
    r = h + 1
    while st.cell(r, 1).value not in (None, ""):
        g = lambda name: st.cell(r, cols[name]).value if name in cols else None  # noqa: E731
        k = to_float(g("Zählerkonstante"))
        wohnungen.append({
            "name": str(g("Wohnung") or "").strip(), "kurz": str(g("Kurzname") or "").strip(),
            "zeitraum": str(g("Zeitraum") or "").strip(), "personen": int(to_float(g("Personen")) or 0),
            "ausstattung": str(g("Heizung & Warmwasser") or "").strip(),
            "zaehlertyp": str(g("Zählertyp") or "").strip(), "zaehlernr": str(g("Zähler-Nr.") or "").strip(),
            "konstante": int(k) if k else None, "farbe": str(g("Dashboard-Farbe") or "blau").strip(),
            "hinweis": str(g("Hinweis") or "").strip()})
        r += 1
    return {"wohnungen": wohnungen, "ablesungen": abl, "vertraege": vertraege, "vergleich": vergleich}


def pruefen(d: dict) -> list[str]:
    """Sortiert und prüft die Daten. Wirft bei harten Fehlern, liefert Warnungen."""
    warn = []
    namen = [w["name"] for w in d["wohnungen"]]
    for a in d["ablesungen"]:
        if a["wohnung"] not in namen:
            raise ValueError(f"Unbekannte Wohnung {a['wohnung']!r} am {a['datum']} (erlaubt: {namen})")
        if a["stand"] is None or a["datum"] is None:
            raise ValueError(f"Ablesung unvollständig: {a}")
    # Wohnungen chronologisch, Ablesungen je Wohnung nach Datum
    erste = {}
    for a in d["ablesungen"]:
        erste[a["wohnung"]] = min(erste.get(a["wohnung"], a["datum"]), a["datum"])
    d["wohnungen"].sort(key=lambda w: erste.get(w["name"], dt.date.max))
    ordnung = {w["name"]: i for i, w in enumerate(d["wohnungen"])}
    d["ablesungen"].sort(key=lambda a: (ordnung[a["wohnung"]], a["datum"]))
    for i in range(1, len(d["ablesungen"])):
        p, a = d["ablesungen"][i - 1], d["ablesungen"][i]
        if a["datum"] < p["datum"]:
            raise ValueError("Ablesungen verschiedener Wohnungen überlappen zeitlich – die Formeln setzen "
                             "eine chronologische Liste voraus.")
        if a["wohnung"] == p["wohnung"] and a["stand"] < p["stand"]:
            raise ValueError(f"Zählerstand sinkt: {de_datum(p['datum'])} {p['stand']} -> "
                             f"{de_datum(a['datum'])} {a['stand']} ({a['wohnung']})")
        if a["wohnung"] == p["wohnung"]:
            tage = (a["datum"] - p["datum"]).days
            if tage > 0:
                pro_tag = (a["stand"] - p["stand"]) / tage
                if pro_tag > 30:
                    warn.append(f"Auffällig hoher Verbrauch {pro_tag:.1f} kWh/Tag bis {de_datum(a['datum'])}")
    d["vertraege"].sort(key=lambda v: v["ab"])
    for v in d["vertraege"] + d["vergleich"]:
        if v["ap"] is None or v["gp"] is None:
            raise ValueError(f"Tarif ohne Preise: {v['bezeichnung']}")
    return warn


def gp_jahr(v: dict) -> float:
    return v["gp"] * 12 if v.get("gp_einheit", "").strip() == "€/Monat" else v["gp"]


# --------------------------------------------------------------------------------------
# Excel erzeugen
# --------------------------------------------------------------------------------------
def _font(**kw):
    kw.setdefault("name", FONT)
    kw.setdefault("size", 10)
    return Font(**kw)


def _fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


THIN = Side(style="thin", color=C_GRID)
B_UNTEN = Border(bottom=THIN)


def _titel(ws, text, sub=None, breite_bis="H"):
    ws["A1"] = text
    ws["A1"].font = _font(size=16, bold=True, color=C_BRAND)
    ws.row_dimensions[1].height = 24
    if sub:
        ws["A2"] = sub
        ws["A2"].font = _font(size=9, italic=True, color=C_MUTED)


def _kopf(ws, row, col, text, fill=C_HDR_CALC, breite=None, wrap=True):
    c = ws.cell(row, col, text)
    c.font = _font(bold=True, color="FFFFFF")
    c.fill = _fill(fill)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    c.border = Border(left=Side(style="thin", color="FFFFFF"), right=Side(style="thin", color="FFFFFF"))
    if breite:
        ws.column_dimensions[get_column_letter(col)].width = breite
    return c


def _abschnittstitel(ws, row, text, col=1):
    c = ws.cell(row, col, text)
    c.font = _font(size=11, bold=True, color=C_BRAND)
    return c


def _zelle(ws, row, col, value, fmt=None, eingabe=False, bold=False, color=None, align=None, italic=False, size=10):
    c = ws.cell(row, col, value)
    c.font = _font(bold=bold, italic=italic, size=size,
                   color=(C_IN_TXT if eingabe else (color or "000000")))
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = Alignment(horizontal=align, vertical="center")
    return c


def excel_erzeugen(d: dict, pfad: Path):
    wb = openpyxl.Workbook()
    ws_a = wb.active
    ws_a.title = "Ablesungen"
    ws_e = wb.create_sheet("Auswertung")
    ws_m = wb.create_sheet("Monatswerte")
    ws_s = wb.create_sheet("Stammdaten")
    ws_i = wb.create_sheet("Info")
    for ws, farbe in ((ws_a, C_HDR_IN), (ws_e, C_BRAND), (ws_m, C_BRAND), (ws_s, C_HDR_CALC), (ws_i, C_HDR_CALC)):
        ws.sheet_properties.tabColor = farbe
        ws.sheet_view.showGridLines = False

    alle_wohnungen = d["wohnungen"]
    abl = d["ablesungen"]
    aktive = [w for w in alle_wohnungen if any(a["wohnung"] == w["name"] for a in abl)]
    wohnungen = aktive
    vertraege = d["vertraege"]
    vergleich = d["vergleich"]
    n_abl = len(abl)
    letzte_zeile = ABL_ERSTE + n_abl - 1

    # ---------------- Stammdaten ----------------
    ws = ws_s
    _titel(ws, "Stammdaten – Verträge, Vergleichstarife, Wohnungen",
           "Preise brutto inkl. MwSt. laut Vertragsunterlagen (Ordner Strom\\SWDU und Strom\\Grüüün). "
           "Blaue Werte = Eingaben; Grundpreis €/Jahr und €/Tag werden berechnet.")
    tarif_cols = [("Nr.", 5), ("Bezeichnung", 26), ("Anbieter", 30), ("Tarif", 34), ("Wohnung", 17),
                  ("gültig ab", 11), ("gültig bis", 11), ("Arbeitspreis\n[ct/kWh]", 12),
                  ("Grundpreis lt. Vertrag", 12), ("Einheit", 9), ("Grundpreis\n[€/Jahr]", 11),
                  ("Grundpreis\n[€/Tag]", 11), ("Abschlag\n[€/Monat]", 11), ("Abschlag fällig", 17),
                  ("Preisgarantie bis", 12), ("Laufzeit & Kündigung", 58), ("Status", 58), ("Quelle", 62)]

    def tarifblock(start_row, titel, liste):
        _abschnittstitel(ws, start_row, titel)
        hdr = start_row + 1
        for j, (name, w) in enumerate(tarif_cols, start=1):
            _kopf(ws, hdr, j, name, fill=C_HDR_IN if j not in (11, 12) else C_HDR_CALC, breite=w)
        ws.row_dimensions[hdr].height = 30
        for i, v in enumerate(liste):
            r = hdr + 1 + i
            _zelle(ws, r, 1, i + 1, align="center")
            _zelle(ws, r, 2, v["bezeichnung"], eingabe=True, bold=True)
            _zelle(ws, r, 3, v["anbieter"], eingabe=True)
            _zelle(ws, r, 4, v["tarif"], eingabe=True)
            _zelle(ws, r, 5, v["wohnung"], eingabe=True)
            _zelle(ws, r, 6, v["ab"], F_DATUM, eingabe=True, align="center")
            _zelle(ws, r, 7, v["bis"], F_DATUM, eingabe=True, align="center")
            _zelle(ws, r, 8, v["ap"], F_CT, eingabe=True)
            _zelle(ws, r, 9, v["gp"], "0.00", eingabe=True)
            _zelle(ws, r, 10, v["gp_einheit"], eingabe=True, align="center")
            _zelle(ws, r, 11, f'=IF(J{r}="€/Monat",I{r}*12,I{r})', F_EUR)
            _zelle(ws, r, 12, f"=K{r}/365", F_EUR_TAG)
            _zelle(ws, r, 13, v.get("abschlag"), F_EUR, eingabe=True)
            _zelle(ws, r, 14, v.get("abschlag_faellig") or None, eingabe=True)
            _zelle(ws, r, 15, v.get("preisgarantie_bis"), F_DATUM, eingabe=True, align="center")
            _zelle(ws, r, 16, v.get("laufzeit") or None, eingabe=True)
            _zelle(ws, r, 17, v.get("status") or None, eingabe=True)
            _zelle(ws, r, 18, v.get("quelle") or None, eingabe=True)
            for c in range(1, 19):
                ws.cell(r, c).border = B_UNTEN
        return hdr + 1, hdr + len(liste)

    v_erste, v_letzte = tarifblock(4, "VERTRÄGE (werden für die Kostenberechnung verwendet, nach „gültig ab“ sortiert)", vertraege)
    vg_start = v_letzte + 3
    vg_erste, vg_letzte = tarifblock(vg_start, "VERGLEICHSTARIFE (nicht abgeschlossen – nur für den Tarifvergleich)", vergleich)
    if not vergleich:
        vg_letzte = vg_erste - 1
    w_start = max(vg_letzte, vg_erste) + 3
    _abschnittstitel(ws, w_start, "WOHNUNGEN")
    w_hdr = w_start + 1
    w_cols = ["Nr.", "Wohnung", "Kurzname", "Zeitraum", "Personen", "Heizung & Warmwasser", "Zählertyp",
              "Zähler-Nr.", "Zählerkonstante\n[U/kWh]", "Dashboard-Farbe", "Hinweis"]
    for j, name in enumerate(w_cols, start=1):
        _kopf(ws, w_hdr, j, name, fill=C_HDR_IN)
    ws.row_dimensions[w_hdr].height = 30
    for i, w in enumerate(alle_wohnungen):
        r = w_hdr + 1 + i
        tint = FARBEN.get(w["farbe"], FARBEN["blau"])[0]
        _zelle(ws, r, 1, i + 1, align="center")
        _zelle(ws, r, 2, w["name"], eingabe=True, bold=True).fill = _fill(tint)
        _zelle(ws, r, 3, w["kurz"], eingabe=True)
        _zelle(ws, r, 4, w["zeitraum"], eingabe=True)
        _zelle(ws, r, 5, w["personen"], eingabe=True, align="center")
        _zelle(ws, r, 6, w["ausstattung"], eingabe=True)
        _zelle(ws, r, 7, w["zaehlertyp"] or None, eingabe=True)
        _zelle(ws, r, 8, w["zaehlernr"] or None, eingabe=True)
        _zelle(ws, r, 9, w["konstante"], eingabe=True, align="center")
        _zelle(ws, r, 10, w["farbe"], eingabe=True, align="center")
        _zelle(ws, r, 11, w["hinweis"] or None, eingabe=True)
        for c in range(1, 12):
            ws.cell(r, c).border = B_UNTEN
    w_erste, w_letzte = w_hdr + 1, w_hdr + len(alle_wohnungen)
    ws.freeze_panes = "C6"

    # Bereichsnamen (als Text für Formeln)
    S = "Stammdaten!"
    rng_ab = f"{S}$F${v_erste}:$F${v_letzte}"
    rng_bez = f"{S}$B${v_erste}:$B${v_letzte}"
    rng_ap = f"{S}$H${v_erste}:$H${v_letzte}"
    rng_gptag = f"{S}$L${v_erste}:$L${v_letzte}"
    MZ = MAX_ZEILEN
    A_DAT = f"Ablesungen!$A${ABL_ERSTE}:$A${MZ}"
    A_STD = f"Ablesungen!$B${ABL_ERSTE}:$B${MZ}"
    A_WHG = f"Ablesungen!$C${ABL_ERSTE}:$C${MZ}"

    # ---------------- Ablesungen ----------------
    ws = ws_a
    _titel(ws, "Stromzähler – Ablesungen",
           "Datenbasis fürs Strom-Dashboard. Eingaben nur in den Spalten A–E (blau); die Spalten F–N rechnen automatisch.")
    ws["A3"] = "Letzter Stand"
    ws["A3"].font = _font(bold=True, size=9)
    _zelle(ws, 3, 2, f"=INDEX($A${ABL_ERSTE}:$A${MZ},MATCH(9.99E+307,$A${ABL_ERSTE}:$A${MZ}))", F_DATUM, bold=True, align="left")
    _zelle(ws, 3, 3, f"=INDEX($B${ABL_ERSTE}:$B${MZ},MATCH(9.99E+307,$A${ABL_ERSTE}:$A${MZ}))", F_KWH1 + ' "kWh"', bold=True, align="left")
    ws["D3"] = "Dashboard:"
    ws["D3"].font = _font(bold=True)
    ws["D3"].alignment = Alignment(horizontal="right")
    ws["E3"] = PAGES_URL
    ws["E3"].hyperlink = PAGES_URL
    ws["E3"].font = _font(color=C_LINK, underline="single")
    # Legende Wohnungen (aktuelle zuerst)
    ws["A4"] = "Wohnungen:"
    ws["A4"].font = _font(bold=True)
    plaetze = [("B", "C"), ("D", None), ("E", None)]
    for (c1, c2), w in zip(plaetze, list(reversed(aktive))):
        tint = FARBEN.get(w["farbe"], FARBEN["blau"])[0]
        c = ws[f"{c1}4"]
        c.value = f"{w['name']} ({w['kurz']})"
        c.font = _font(size=9, bold=True)
        c.fill = _fill(tint)
        c.alignment = Alignment(horizontal="center")
        if c2:
            ws[f"{c2}4"].fill = _fill(tint)
            ws.merge_cells(f"{c1}4:{c2}4")
    # Gruppen-Kopf
    ws.merge_cells("A5:E5")
    ws["A5"] = "EINGABE"
    ws["A5"].font = _font(bold=True, color=C_HDR_IN, size=9)
    ws["A5"].fill = _fill("DCE8F7")
    ws["A5"].alignment = Alignment(horizontal="center")
    ws.merge_cells("F5:N5")
    ws["F5"] = "BERECHNET (Formeln – nicht überschreiben)"
    ws["F5"].font = _font(bold=True, color="595959", size=9)
    ws["F5"].fill = _fill("EDEDED")
    ws["F5"].alignment = Alignment(horizontal="center")
    abl_cols = [("Datum", 12, C_HDR_IN), ("Zählerstand\n[kWh]", 14, C_HDR_IN), ("Wohnung", 19, C_HDR_IN),
                ("Quelle", 56, C_HDR_IN), ("Bemerkung", 60, C_HDR_IN),
                ("Tage", 7, C_HDR_CALC), ("Verbrauch\n[kWh]", 11, C_HDR_CALC), ("Ø kWh\npro Tag", 9, C_HDR_CALC),
                ("Hochrechnung\n[kWh/Jahr]", 15, C_HDR_CALC), ("Tarif", 25, C_HDR_CALC),
                ("Arbeitspreis\n[ct/kWh]", 13, C_HDR_CALC), ("Grundpreis\n[€/Tag]", 12, C_HDR_CALC),
                ("Kosten\nZeitraum [€]", 13, C_HDR_CALC), ("Ø Kosten\n[€/Monat]", 12, C_HDR_CALC)]
    for j, (name, w, fill) in enumerate(abl_cols, start=1):
        _kopf(ws, ABL_HDR, j, name, fill=fill, breite=w)
    ws.row_dimensions[ABL_HDR].height = 30
    for i, a in enumerate(abl):
        r = ABL_ERSTE + i
        p = r - 1
        _zelle(ws, r, 1, a["datum"], F_DATUM, eingabe=True, align="center")
        _zelle(ws, r, 2, a["stand"], F_KWH1, eingabe=True)
        _zelle(ws, r, 3, a["wohnung"], eingabe=True)
        _zelle(ws, r, 4, a["quelle"] or None, eingabe=True)
        _zelle(ws, r, 5, a["bemerkung"] or None, eingabe=True)
        mitte = f"$A{p}+F{r}/2"
        _zelle(ws, r, 6, f'=IF($C{r}=$C{p},$A{r}-$A{p},"")', "0", align="center")
        _zelle(ws, r, 7, f'=IF($C{r}=$C{p},$B{r}-$B{p},"")', F_KWH1)
        _zelle(ws, r, 8, f'=IF(N(F{r})>0,G{r}/F{r},"")', F_KWHD)
        _zelle(ws, r, 9, f'=IF(H{r}="","",H{r}*365)', F_KWH0)
        _zelle(ws, r, 10, f'=IF(G{r}="","",IFERROR(INDEX({rng_bez},MATCH({mitte},{rng_ab},1)),""))')
        _zelle(ws, r, 11, f'=IF(J{r}="","",INDEX({rng_ap},MATCH({mitte},{rng_ab},1)))', F_CT)
        _zelle(ws, r, 12, f'=IF(J{r}="","",INDEX({rng_gptag},MATCH({mitte},{rng_ab},1)))', F_EUR_TAG)
        _zelle(ws, r, 13, f'=IF(J{r}="","",G{r}*K{r}/100+F{r}*L{r})', F_EUR)
        _zelle(ws, r, 14, f'=IF(M{r}="","",M{r}/F{r}*365/12)', F_EUR)
        for c in range(1, 15):
            ws.cell(r, c).border = B_UNTEN
    # bedingte Formatierung: Tönung je Wohnung + Trennlinie beim Wohnungswechsel
    cf_bereich = f"A{ABL_ERSTE}:N{MZ}"
    for w in wohnungen:
        tint = FARBEN.get(w["farbe"], FARBEN["blau"])[0]
        ws.conditional_formatting.add(cf_bereich, FormulaRule(formula=[f'$C{ABL_ERSTE}="{w["name"]}"'], fill=_fill(tint)))
    ws.conditional_formatting.add(
        f"A{ABL_ERSTE + 1}:N{MZ}",
        FormulaRule(formula=[f'AND($C{ABL_ERSTE + 1}<>"",$C{ABL_ERSTE}<>"",$C{ABL_ERSTE + 1}<>$C{ABL_ERSTE})'],
                    border=Border(top=Side(style="thin", color="404040"))))
    dv = DataValidation(type="list", formula1=f"=Stammdaten!$B${w_erste}:$B${w_letzte}", allow_blank=True,
                        showErrorMessage=True, errorTitle="Wohnung", error="Bitte eine Wohnung aus den Stammdaten wählen.")
    ws.add_data_validation(dv)
    dv.add(f"C{ABL_ERSTE}:C{MZ}")
    ws.freeze_panes = f"B{ABL_ERSTE}"
    ws.auto_filter.ref = f"A{ABL_HDR}:N{letzte_zeile}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    # ---------------- Monatswerte ----------------
    ws = ws_m
    _titel(ws, "Monatswerte",
           "Zählerstand an Monatsanfang/-ende linear zwischen den benachbarten Ablesungen interpoliert. "
           "Teilmonat = nur der Zeitraum, der durch Ablesungen abgedeckt ist. Neue Monate ergänzt das Update-Skript.")
    m_cols = [("Monat", 11), ("Jahr", 7), ("Monat-\nNr.", 7), ("Wohnung", 19), ("Von", 11), ("Bis", 11),
              ("Tage", 7), ("Stand Von\n[kWh]", 12), ("Stand Bis\n[kWh]", 12), ("Verbrauch\n[kWh]", 11),
              ("Ø kWh\npro Tag", 9), ("Voller Monat\n[kWh]", 13), ("Tarif", 25), ("Kosten\n[€]", 11),
              ("Hinweis", 28), ("Zeile\nVon", 8), ("Zeile\nBis", 8)]
    M_HDR = 4
    for j, (name, w) in enumerate(m_cols, start=1):
        _kopf(ws, M_HDR, j, name, fill=C_BRAND if j < 16 else "A6A6A6", breite=w)
    ws.row_dimensions[M_HDR].height = 30

    def erste_dat(zelle):
        return f"INDEX({A_DAT},MATCH({zelle},{A_WHG},0))"

    def letzte_dat(zelle):
        return f"INDEX({A_DAT},MATCH({zelle},{A_WHG},0)+COUNTIF({A_WHG},{zelle})-1)"

    def interpol(dat, idx):
        i0 = f"INDEX({A_DAT},{idx})"
        s0 = f"INDEX({A_STD},{idx})"
        i1 = f"INDEX({A_DAT},{idx}+1)"
        s1 = f"INDEX({A_STD},{idx}+1)"
        return f"=IF({dat}={i0},{s0},{s0}+({s1}-{s0})*({dat}-{i0})/({i1}-{i0}))"

    monate = []   # (Monatserster, Wohnung)
    for w in wohnungen:
        dats = [a["datum"] for a in abl if a["wohnung"] == w["name"]]
        if len(dats) < 2:
            continue
        m = dt.date(dats[0].year, dats[0].month, 1)
        while m < dats[-1]:
            monate.append((m, w["name"]))
            m = add_months(m, 1)
    for i, (m, wname) in enumerate(monate):
        r = M_HDR + 1 + i
        _zelle(ws, r, 1, m, F_MONAT, align="center")
        _zelle(ws, r, 2, f"=YEAR(A{r})", "0", align="center")
        _zelle(ws, r, 3, f"=MONTH(A{r})", "0", align="center")
        _zelle(ws, r, 4, wname)
        _zelle(ws, r, 5, f"=MAX(A{r},{erste_dat(f'D{r}')})", F_DATUM, align="center")
        _zelle(ws, r, 6, f"=MIN(EDATE(A{r},1),{letzte_dat(f'D{r}')})", F_DATUM, align="center")
        _zelle(ws, r, 7, f"=MAX(0,F{r}-E{r})", "0", align="center")
        _zelle(ws, r, 16, f"=MATCH(E{r},{A_DAT},1)", "0", color=C_MUTED, align="center")
        _zelle(ws, r, 17, f"=MATCH(F{r},{A_DAT},1)", "0", color=C_MUTED, align="center")
        _zelle(ws, r, 8, interpol(f"E{r}", f"P{r}"), F_KWH1)
        _zelle(ws, r, 9, interpol(f"F{r}", f"Q{r}"), F_KWH1)
        _zelle(ws, r, 10, f"=I{r}-H{r}", F_KWH1)
        _zelle(ws, r, 11, f'=IF(G{r}>0,J{r}/G{r},"")', F_KWHD)
        _zelle(ws, r, 12, f'=IF(K{r}="","",K{r}*DAY(EOMONTH(A{r},0)))', F_KWH0)
        mitte = f"E{r}+G{r}/2"
        _zelle(ws, r, 13, f'=IF(G{r}=0,"",IFERROR(INDEX({rng_bez},MATCH({mitte},{rng_ab},1)),""))')
        _zelle(ws, r, 14, f'=IF(M{r}="","",J{r}*INDEX({rng_ap},MATCH({mitte},{rng_ab},1))/100'
                           f'+G{r}*INDEX({rng_gptag},MATCH({mitte},{rng_ab},1)))', F_EUR)
        _zelle(ws, r, 15, f'=IF(G{r}<DAY(EOMONTH(A{r},0)),"Teilmonat ("&G{r}&" Tage erfasst)","")', color=C_MUTED)
        tint = FARBEN.get(next(w["farbe"] for w in wohnungen if w["name"] == wname), FARBEN["blau"])[0]
        for c in range(1, 18):
            ws.cell(r, c).border = B_UNTEN
            if c <= 15:
                ws.cell(r, c).fill = _fill(tint)
    m_erste, m_letzte = M_HDR + 1, M_HDR + max(1, len(monate))
    ws.freeze_panes = f"B{M_HDR + 1}"
    if monate:
        ws.auto_filter.ref = f"A{M_HDR}:O{m_letzte}"
    MW = lambda col: f"Monatswerte!${col}${m_erste}:${col}${m_letzte}"  # noqa: E731

    # ---------------- Auswertung ----------------
    ws = ws_e
    _titel(ws, "Auswertung",
           "Alle Werte sind Formeln auf Basis der Blätter „Ablesungen“, „Monatswerte“ und „Stammdaten“.")
    for col, w in zip("ABCDEFGHIJKLMNO", (38, 20, 14, 13, 13, 12, 14, 13, 13, 14, 11, 13, 9, 9, 9)):
        ws.column_dimensions[col].width = w
    # Block 1: Kennzahlen je Wohnung
    r0 = 4
    _abschnittstitel(ws, r0, "KENNZAHLEN JE WOHNUNG")
    k_cols = ["Wohnung", "Erste\nAblesung", "Letzte\nAblesung", "Tage", "Verbrauch\n[kWh]", "Ø kWh\npro Tag",
              "Hochrechnung\n[kWh/Jahr]", "Anzahl\nAblesungen", "Kosten erfasst\n[€]", "Ø Kosten\n[€/Monat]",
              "Zeile\nerste", "Zeile\nletzte"]
    for j, name in enumerate(k_cols, start=1):
        _kopf(ws, r0 + 1, j, name, fill=C_BRAND if j <= 10 else "A6A6A6")
    ws.row_dimensions[r0 + 1].height = 30
    k_row = {}
    for i, w in enumerate(wohnungen):
        r = r0 + 2 + i
        k_row[w["name"]] = r
        tint = FARBEN.get(w["farbe"], FARBEN["blau"])[0]
        _zelle(ws, r, 1, w["name"], bold=True).fill = _fill(tint)
        _zelle(ws, r, 11, f"=MATCH(A{r},{A_WHG},0)", "0", color=C_MUTED, align="center")
        _zelle(ws, r, 12, f"=K{r}+COUNTIF({A_WHG},A{r})-1", "0", color=C_MUTED, align="center")
        _zelle(ws, r, 2, f"=INDEX({A_DAT},K{r})", F_DATUM, align="center")
        _zelle(ws, r, 3, f"=INDEX({A_DAT},L{r})", F_DATUM, align="center")
        _zelle(ws, r, 4, f"=C{r}-B{r}", "0", align="center")
        _zelle(ws, r, 5, f"=INDEX({A_STD},L{r})-INDEX({A_STD},K{r})", F_KWH1)
        _zelle(ws, r, 6, f'=IF(D{r}>0,E{r}/D{r},"")', F_KWHD, bold=True)
        _zelle(ws, r, 7, f'=IF(F{r}="","",F{r}*365)', F_KWH0, bold=True)
        _zelle(ws, r, 8, f"=COUNTIF({A_WHG},A{r})", "0", align="center")
        _zelle(ws, r, 9, f"=SUMIFS(Ablesungen!$M${ABL_ERSTE}:$M${MZ},{A_WHG},A{r})", F_EUR_STRICH)
        _zelle(ws, r, 10, f'=IF(I{r}>0,I{r}/SUMIFS(Ablesungen!$F${ABL_ERSTE}:$F${MZ},{A_WHG},A{r},'
                          f'Ablesungen!$M${ABL_ERSTE}:$M${MZ},">0")*365/12,"")', F_EUR)
        for c in range(1, 13):
            ws.cell(r, c).border = B_UNTEN
    akt = wohnungen[-1]["name"]
    r_akt = k_row[akt]
    r = r0 + 2 + len(wohnungen) + 1

    # Block 2: Vergleich
    if len(wohnungen) >= 2:
        alt = wohnungen[-2]["name"]
        r_alt = k_row[alt]
        _abschnittstitel(ws, r, f"VERGLEICH: {akt} ↔ {alt}")
        r += 1
        _zelle(ws, r, 1, f"Ø kWh pro Tag – {alt}")
        _zelle(ws, r, 2, f"=F{r_alt}", F_KWHD)
        r += 1
        _zelle(ws, r, 1, f"Ø kWh pro Tag – {akt}")
        _zelle(ws, r, 2, f"=F{r_akt}", F_KWHD)
        r += 1
        _zelle(ws, r, 1, "Veränderung (gesamter Zeitraum)", bold=True)
        _zelle(ws, r, 2, f'=IF(AND(N(B{r - 2})>0,N(B{r - 1})>0),B{r - 1}/B{r - 2}-1,"")', F_PROZ, bold=True)
        r += 2
        _zelle(ws, r, 1, "Saisonvergleich je Kalendermonat (aus „Monatswerte“)", bold=True, color=C_BRAND)
        _zelle(ws, r, 2, f"alt = {alt} · neu = {akt}", italic=True, color=C_MUTED, size=9)
        r += 1
        s_cols = ["Monat", "Tage\nalt", "kWh\nalt", "Ø kWh/Tag\nalt",
                  "Tage\nneu", "kWh\nneu", "Ø kWh/Tag\nneu", "Veränderung"]
        for j, name in enumerate(s_cols, start=1):
            _kopf(ws, r, j, name, fill=C_BRAND)
        ws.row_dimensions[r].height = 30
        s_erste = r + 1
        monatsnamen = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September",
                       "Oktober", "November", "Dezember"]
        for mnr in range(1, 13):
            rr = s_erste + mnr - 1
            _zelle(ws, rr, 1, monatsnamen[mnr - 1])
            for off, (wn, rw) in enumerate(((alt, r_alt), (akt, r_akt))):
                ct, ck, cd = 2 + off * 3, 3 + off * 3, 4 + off * 3
                lt, lk = get_column_letter(ct), get_column_letter(ck)
                _zelle(ws, rr, ct, f"=SUMIFS({MW('G')},{MW('D')},$A${rw},{MW('C')},{mnr})", "0", align="center")
                _zelle(ws, rr, ck, f"=SUMIFS({MW('J')},{MW('D')},$A${rw},{MW('C')},{mnr})", F_KWH0)
                _zelle(ws, rr, cd, f'=IF({lt}{rr}>0,{lk}{rr}/{lt}{rr},"")', F_KWHD)
            _zelle(ws, rr, 8, f'=IF(AND(N(D{rr})>0,N(G{rr})>0),G{rr}/D{rr}-1,"")', F_PROZ)
            for c in range(1, 9):
                ws.cell(rr, c).border = B_UNTEN
        s_letzte = s_erste + 11
        rr = s_letzte + 1
        _zelle(ws, rr, 1, "Monate mit Daten in beiden Wohnungen", bold=True)
        bed = f"(B{s_erste}:B{s_letzte}>0)*(E{s_erste}:E{s_letzte}>0)"
        _zelle(ws, rr, 4, f'=IFERROR(SUMPRODUCT({bed}*C{s_erste}:C{s_letzte})/SUMPRODUCT({bed}*B{s_erste}:B{s_letzte}),"")', F_KWHD, bold=True)
        _zelle(ws, rr, 7, f'=IFERROR(SUMPRODUCT({bed}*F{s_erste}:F{s_letzte})/SUMPRODUCT({bed}*E{s_erste}:E{s_letzte}),"")', F_KWHD, bold=True)
        _zelle(ws, rr, 8, f'=IF(AND(N(D{rr})>0,N(G{rr})>0),G{rr}/D{rr}-1,"")', F_PROZ, bold=True)
        for c in range(1, 9):
            ws.cell(rr, c).border = Border(top=Side(style="thin", color="404040"))
        r = rr + 3

    # Block 3: Jahreswerte + Diagramm
    _abschnittstitel(ws, r, "JAHRESWERTE (Kalenderjahre, linear interpoliert)")
    r += 1
    j_cols = ["Jahr", "Wohnung", "Von", "Bis", "Tage\nerfasst", "Stand Von\n[kWh]", "Stand Bis\n[kWh]",
              "Verbrauch\n[kWh]", "Ø kWh\npro Tag", "Hochrechnung\n[kWh/Jahr]", "Abdeckung", "Kosten\nerfasst [€]",
              "Zeile\nVon", "Zeile\nBis"]
    for j, name in enumerate(j_cols, start=1):
        _kopf(ws, r, j, name, fill=C_BRAND if j <= 12 else "A6A6A6")
    ws.row_dimensions[r].height = 30
    jahre = []
    for w in wohnungen:
        dats = [a["datum"] for a in abl if a["wohnung"] == w["name"]]
        if len(dats) < 2:
            continue
        for y in range(dats[0].year, dats[-1].year + 1):
            if dt.date(y + 1, 1, 1) > dats[0] and dt.date(y, 1, 1) < dats[-1]:
                jahre.append((y, w["name"]))
    j_erste = r + 1
    j_zeile = {}
    for i, (y, wname) in enumerate(jahre):
        rr = j_erste + i
        j_zeile[(y, wname)] = rr
        wohn = next(w for w in wohnungen if w["name"] == wname)
        tint = FARBEN.get(wohn["farbe"], FARBEN["blau"])[0]
        _zelle(ws, rr, 1, y, "0", align="center")
        _zelle(ws, rr, 2, wname)
        _zelle(ws, rr, 3, f"=MAX(DATE(A{rr},1,1),{erste_dat(f'B{rr}')})", F_DATUM, align="center")
        _zelle(ws, rr, 4, f"=MIN(DATE(A{rr}+1,1,1),{letzte_dat(f'B{rr}')})", F_DATUM, align="center")
        _zelle(ws, rr, 5, f"=MAX(0,D{rr}-C{rr})", "0", align="center")
        _zelle(ws, rr, 13, f"=MATCH(C{rr},{A_DAT},1)", "0", color=C_MUTED, align="center")
        _zelle(ws, rr, 14, f"=MATCH(D{rr},{A_DAT},1)", "0", color=C_MUTED, align="center")
        _zelle(ws, rr, 6, interpol(f"C{rr}", f"M{rr}"), F_KWH1)
        _zelle(ws, rr, 7, interpol(f"D{rr}", f"N{rr}"), F_KWH1)
        _zelle(ws, rr, 8, f"=G{rr}-F{rr}", F_KWH0)
        _zelle(ws, rr, 9, f'=IF(E{rr}>0,H{rr}/E{rr},"")', F_KWHD)
        _zelle(ws, rr, 10, f'=IF(I{rr}="","",I{rr}*365)', F_KWH0, bold=True)
        _zelle(ws, rr, 11, f"=E{rr}/(DATE(A{rr}+1,1,1)-DATE(A{rr},1,1))", F_PROZ0, align="center")
        _zelle(ws, rr, 12, f"=SUMIFS({MW('N')},{MW('D')},B{rr},{MW('B')},A{rr})", F_EUR_STRICH)
        for c in range(1, 15):
            ws.cell(rr, c).border = B_UNTEN
            if c <= 2:
                ws.cell(rr, c).fill = _fill(tint)
    j_letzte = j_erste + len(jahre) - 1
    r = j_letzte + 1
    _zelle(ws, r, 1, "Hochrechnung = Ø kWh pro Tag × 365. Teiljahre (Abdeckung < 100 %) sind hochgerechnet.",
           italic=True, color=C_MUTED, size=9)

    # Diagrammdaten (rechts neben den Jahreswerten) + Säulendiagramm
    dd_col = 16  # Spalte P
    ws.cell(j_erste - 1, dd_col, "Diagrammdaten").font = _font(bold=True, size=9, color=C_MUTED)
    ws.cell(j_erste - 1, dd_col).alignment = Alignment(wrap_text=True)
    jahr_liste = sorted({y for y, _ in jahre})
    for j, w in enumerate(wohnungen, start=1):
        c = ws.cell(j_erste - 1, dd_col + j, w["name"])
        c.font = _font(bold=True, size=9, color=C_MUTED)
        c.alignment = Alignment(wrap_text=True)
        ws.column_dimensions[get_column_letter(dd_col + j)].width = 13
    ws.column_dimensions[get_column_letter(dd_col)].width = 11
    for i, y in enumerate(jahr_liste):
        rr = j_erste + i
        ws.cell(rr, dd_col, str(y)).font = _font(size=9, color=C_MUTED)
        for j, w in enumerate(wohnungen, start=1):
            if (y, w["name"]) in j_zeile:
                c = ws.cell(rr, dd_col + j, f"=J{j_zeile[(y, w['name'])]}")
                c.number_format = F_KWH0
                c.font = _font(size=9, color=C_MUTED)
    if jahr_liste:
        ch = BarChart()
        ch.type = "col"
        ch.grouping = "clustered"
        ch.overlap = 100
        ch.gapWidth = 70
        ch.title = "Jahresverbrauch (hochgerechnet auf 365 Tage)"
        ch.y_axis.title = "kWh/Jahr"
        ch.y_axis.numFmt = "#,##0"
        ch.y_axis.majorGridlines.spPr = None
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.legend.position = "b"
        ch.height = 7.5
        ch.width = 17
        dref = Reference(ws, min_col=dd_col + 1, max_col=dd_col + len(wohnungen), min_row=j_erste - 1,
                         max_row=j_erste + len(jahr_liste) - 1)
        cref = Reference(ws, min_col=dd_col, min_row=j_erste, max_row=j_erste + len(jahr_liste) - 1)
        ch.add_data(dref, titles_from_data=True)
        ch.set_categories(cref)
        for s, w in zip(ch.series, wohnungen):
            voll = FARBEN.get(w["farbe"], FARBEN["blau"])[1]
            s.graphicalProperties.solidFill = voll
            s.graphicalProperties.line.solidFill = voll
        ch.dataLabels = DataLabelList()
        ch.dataLabels.showVal = True
        ch.dataLabels.showCatName = False
        ch.dataLabels.showSerName = False
        ch.dataLabels.showLegendKey = False
        ch.dataLabels.showPercent = False
        ch.dataLabels.numFmt = "#,##0"
        ws.add_chart(ch, f"A{r + 2}")
        r += 2 + 16

    # Block 4: Abrechnungsjahr aktueller Vertrag
    r += 1
    nv = len(vertraege)
    vz = v_erste + nv - 1   # Zeile des aktuellen Vertrags in den Stammdaten
    _abschnittstitel(ws, r, "ABRECHNUNGSJAHR – AKTUELLER VERTRAG (Prognose)")
    r += 1
    zeilen = []

    def zeile(label, formel, fmt=None, hinweis=None, bold=False, saldo=False, links=False):
        nonlocal r
        _zelle(ws, r, 1, label, bold=bold)
        c = _zelle(ws, r, 2, formel, fmt, bold=bold, align="left" if links else "right")
        if saldo:
            _zelle(ws, r, 3, f'=IF(B{r}<0,"Nachzahlung",IF(B{r}>0,"Guthaben",""))', bold=True)
        if hinweis:
            _zelle(ws, r, 4, hinweis, italic=True, color=C_MUTED, size=9)
        for cc in range(1, 4):
            ws.cell(r, cc).border = B_UNTEN
        zeilen.append(r)
        r += 1
        return c.coordinate

    z_vertrag = zeile("Vertrag", f"=Stammdaten!B{vz}", None, None, links=True)
    z_beginn = zeile("Lieferbeginn", f"=Stammdaten!F{vz}", F_DATUM)
    z_ende = zeile("Abrechnungsjahr endet", f"=EDATE({z_beginn},12)-1", F_DATUM, "12 Monate ab Lieferbeginn")
    # Anfangsstand: Ablesung am Vortag des Lieferbeginns (Wechselablesung), sonst interpoliert
    z_anf = zeile("Anfangsstand [kWh]", interpol(f"({z_beginn}-1)", f"MATCH({z_beginn}-1,{A_DAT},1)"), F_KWH1,
                  "Zählerstand am Vortag des Lieferbeginns (Wechselablesung)")
    z_ldat = zeile("Letzte Ablesung", f"=C{r_akt}", F_DATUM)
    z_lstd = zeile("Zählerstand letzte Ablesung [kWh]", f"=INDEX({A_STD},L{r_akt})", F_KWH1)
    z_kwh = zeile("Verbrauch seit Lieferbeginn [kWh]", f"={z_lstd}-{z_anf}", F_KWH1)
    z_tage = zeile("Tage seit Lieferbeginn", f"={z_ldat}-({z_beginn}-1)", "0")
    zeile("Ø kWh pro Tag seit Lieferbeginn", f'=IF({z_tage}>0,{z_kwh}/{z_tage},"")', F_KWHD)
    z_ap = f"Stammdaten!H{vz}"
    z_gpj = f"Stammdaten!K{vz}"
    z_abs = f"Stammdaten!M{vz}"
    z_kost = zeile("Kosten seit Lieferbeginn [€]", f"={z_kwh}*{z_ap}/100+{z_tage}*{z_gpj}/365", F_EUR)
    z_absb = zeile("Abschläge anteilig [€]", f"={z_tage}/365*12*{z_abs}", F_EUR, "anteilig für die erfassten Tage")
    zeile("Saldo bisher", f"={z_absb}-{z_kost}", F_SALDO, bold=True, saldo=True)
    z_basis = zeile("Prognose-Basis: Ø kWh pro Tag", f"=F{r_akt}", F_KWHD,
                    f"Ø seit erster Ablesung in der {akt}")
    z_rest = zeile("Resttage bis Abrechnungsende", f"=MAX(0,{z_ende}-{z_ldat})", "0")
    z_pkwh = zeile("Prognose Jahresverbrauch [kWh]", f"={z_kwh}+{z_basis}*{z_rest}", F_KWH0, bold=True)
    z_pkost = zeile("Prognose Jahreskosten [€]", f"={z_gpj}+{z_pkwh}*{z_ap}/100", F_EUR, bold=True)
    z_absj = zeile("Abschläge im Abrechnungsjahr [€]", f"=12*{z_abs}", F_EUR)
    zeile("Erwarteter Saldo am Abrechnungsende", f"={z_absj}-{z_pkost}", F_SALDO, bold=True, saldo=True)
    zeile("Passender Abschlag [€/Monat]", f"=ROUNDUP({z_pkost}/12,0)", F_EUR, "aufgerundet auf volle Euro", bold=True)

    # Block 5: Tarifvergleich
    r += 1
    _abschnittstitel(ws, r, "TARIFVERGLEICH BEI PROGNOSE-JAHRESVERBRAUCH")
    r += 1
    for j, name in enumerate(["Tarif", "Arbeitspreis\n[ct/kWh]", "Grundpreis\n[€/Jahr]", "Jahreskosten\n[€]",
                              "Differenz zum\naktuellen Tarif"], start=1):
        _kopf(ws, r, j, name, fill=C_BRAND)
    ws.row_dimensions[r].height = 30
    r += 1
    t_zeilen = [(v_erste + i, "") for i in range(nv)] + [(vg_erste + i, " (Vergleich)") for i in range(len(vergleich))]
    t_akt_row = None
    for sz, zusatz in t_zeilen:
        _zelle(ws, r, 1, f'=Stammdaten!B{sz}&"{zusatz}"', bold=(sz == vz))
        _zelle(ws, r, 2, f"=Stammdaten!H{sz}", F_CT)
        _zelle(ws, r, 3, f"=Stammdaten!K{sz}", F_EUR)
        _zelle(ws, r, 4, f"=C{r}+{z_pkwh}*B{r}/100", F_EUR, bold=(sz == vz))
        if sz == vz:
            t_akt_row = r
        for c in range(1, 6):
            ws.cell(r, c).border = B_UNTEN
        r += 1
    for rr in range(r - len(t_zeilen), r):
        _zelle(ws, rr, 5, f"=D{rr}-$D${t_akt_row}", '+#,##0.00 "€";-#,##0.00 "€";"–"')
    ws.freeze_panes = "B4"

    # ---------------- Info ----------------
    ws = ws_i
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 120
    ws["B1"] = "Strom – Datenbasis & Auswertung"
    ws["B1"].font = _font(size=16, bold=True, color=C_BRAND)
    texte = [
        ("Dashboard", None),
        (PAGES_URL, "link"),
        ("Neue Ablesung", None),
        ("1. Foto vom Zähler in den Ordner „Auswertung und Ablesungen“ legen und Claude Bescheid geben – Claude trägt den Wert ein, "
         "aktualisiert Excel und Dashboard und verschiebt das Foto nach „zu_loeschen“.", ""),
        ("2. Oder selbst: im Blatt „Ablesungen“ unter der letzten Zeile eine neue Zeile anfügen und nur die Spalten A–E ausfüllen "
         "(Formeln F–N aus der Zeile darüber herunterkopieren). Das Online-Dashboard ändert sich erst nach dem Update-Skript.", ""),
        ("Farben", None),
        ("Blaue Schrift = Eingabe · schwarze Schrift = Formel · graue Schrift = Hilfsspalte · "
         "Zeilen orange hinterlegt = Robert-Koch-Straße (alte Wohnung) · blau hinterlegt = Hufelandstraße.", ""),
        ("Berechnung", None),
        ("Verbrauch je Ableseintervall = Differenz der Zählerstände; Ø kWh pro Tag = Verbrauch / Tage; Hochrechnung = Ø × 365.", ""),
        ("Monats- und Jahreswerte: Zählerstand an den Grenzen linear zwischen den benachbarten Ablesungen interpoliert "
         "(lange Ableseintervalle ergeben geglättete Werte).", ""),
        ("Kosten = Verbrauch × Arbeitspreis + Tage × Grundpreis/365. Maßgeblich ist der Vertrag, der zur Mitte des Intervalls gilt "
         "(Stammdaten → VERTRÄGE, nach „gültig ab“ sortiert).", ""),
        ("Prognose Abrechnungsjahr = gemessener Verbrauch seit Lieferbeginn + Ø kWh/Tag der aktuellen Wohnung × Resttage. "
         "Saldo = Abschläge (12 × Abschlag) − Prognosekosten.", ""),
        ("Bekannte Lücken", None),
        ("Hufelandstraße: Einzugszählerstand (Mai 2025) fehlt – steht in der SWDU-Jahresverbrauchsabrechnung "
         "(Link in Strom\\SWDU\\06.05.2026 1128.txt) bzw. der SWDU-Schlussrechnung. PDF in den Ordner legen, dann wird er ergänzt.", ""),
        ("Robert-Koch-Straße: keine Ablesungen nach dem 01.09.2024. Der Wert vom 03.03.2022 ist identisch mit dem Vortag (vermutlich doppelt erfasst).", ""),
        ("Zähler Hufelandstraße", None),
        ("Siemens 7CA54 Ferraris-Drehstromzähler, Zähler-Nr. 94012994, 75 Umdrehungen pro kWh (Foto vom 01.11.2026).", ""),
        ("Pflege", None),
        ("Diese Datei wird vom Skript _claude/strom_update.py aus den Eingaben neu aufgebaut (Formate, Formeln, Diagramm). "
         "Eigene Zusatzblätter bitte in einer separaten Datei anlegen. Sicherungen: _claude/backup.", ""),
    ]
    rr = 3
    for text, art in texte:
        if art is None:
            if rr > 3:
                rr += 1
            c = ws.cell(rr, 2, text)
            c.font = _font(size=11, bold=True, color=C_BRAND)
        elif art == "link":
            c = ws.cell(rr, 2, text)
            c.hyperlink = text
            c.font = _font(color=C_LINK, underline="single")
        else:
            c = ws.cell(rr, 2, text)
            c.font = _font()
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[rr].height = 27 if len(text) > 125 else 15
        rr += 1
    ws.cell(rr + 1, 2, f"Stand der Datei: {jetzt().strftime('%d.%m.%Y %H:%M')}").font = _font(size=9, italic=True, color=C_MUTED)

    for blatt in wb.worksheets:
        blatt.page_setup.orientation = "landscape"
        blatt.page_setup.paperSize = blatt.PAPERSIZE_A4
        blatt.page_setup.fitToWidth = 1
        blatt.page_setup.fitToHeight = 0
        blatt.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        blatt.print_options.horizontalCentered = True
        blatt.page_margins.left = blatt.page_margins.right = 0.4
    ws_e.page_setup.orientation = "portrait"
    ws_e.print_area = f"A1:L{ws_e.max_row}"
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    pfad = Path(pfad)
    tmp = pfad.with_name(pfad.stem + ".tmp.xlsx")
    wb.save(tmp)
    tmp.replace(pfad)


# --------------------------------------------------------------------------------------
# LibreOffice-Neuberechnung (damit auch Vorschauen Werte zeigen)
# --------------------------------------------------------------------------------------
def neuberechnen(pfad: Path) -> str:
    """Rechnet alle Formeln mit LibreOffice durch und übernimmt das Ergebnis, wenn es fehlerfrei ist.
    So stehen auch in Vorschauen (OneDrive, Explorer) Werte. Ohne LibreOffice rechnet Excel beim Öffnen."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return "übersprungen (LibreOffice nicht gefunden – Excel rechnet beim Öffnen)"
    env = dict(os.environ, SAL_USE_VCLPLUGIN="svp")
    with tempfile.TemporaryDirectory() as td:
        profil = Path(td, "profil")
        makro_dir = profil / "user" / "basic" / "Standard"
        try:
            subprocess.run([soffice, f"-env:UserInstallation={profil.as_uri()}", "--headless",
                            "--terminate_after_init"], capture_output=True, timeout=120, env=env)
        except Exception as e:  # noqa: BLE001
            return f"übersprungen ({e})"
        if not makro_dir.exists():
            return "übersprungen (kein LibreOffice-Profil)"
        (makro_dir / "Module1.xba").write_text(
            '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
            '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">\n'
            "Sub RecalculateAndSave()\n  ThisComponent.calculateAll()\n  ThisComponent.store()\n  ThisComponent.close(True)\nEnd Sub\n"
            "</script:module>", encoding="utf-8")
        arbeits = Path(td, pfad.name)
        shutil.copy2(pfad, arbeits)
        vorher = arbeits.stat().st_mtime_ns, arbeits.stat().st_size
        try:
            subprocess.run([soffice, f"-env:UserInstallation={profil.as_uri()}", "--headless", "--norestore",
                            "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
                            str(arbeits)], capture_output=True, timeout=180, env=env)
        except subprocess.TimeoutExpired:
            return "übersprungen (Zeitüberschreitung)"
        if (arbeits.stat().st_mtime_ns, arbeits.stat().st_size) == vorher:
            return "übersprungen (LibreOffice hat nichts gespeichert)"
        wb = openpyxl.load_workbook(arbeits, data_only=True)
        fehler, formeln = [], 0
        codes = ("#DIV/0!", "#N/A", "#NAME?", "#REF!", "#VALUE!", "#NUM!", "#NULL!")
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and (c.value in codes or c.value.startswith("Err:")):
                        fehler.append(f"{ws.title}!{c.coordinate}={c.value}")
        if fehler:
            return "NICHT übernommen – Formelfehler: " + ", ".join(fehler[:20])
        tmp = pfad.with_name(pfad.stem + ".tmp.xlsx")
        shutil.copy2(arbeits, tmp)
        tmp.replace(pfad)
        return "ok – Formeln berechnet, keine Fehler"


# --------------------------------------------------------------------------------------
# Dashboard-Daten
# --------------------------------------------------------------------------------------
def messwerte_js(d: dict) -> str:
    def tarif(v):
        return {"bezeichnung": v["bezeichnung"], "anbieter": v["anbieter"], "tarif": v["tarif"], "wohnung": v["wohnung"],
                "ab": iso(v["ab"]), "bis": iso(v["bis"]), "ap": v["ap"], "gpJahr": round(gp_jahr(v), 4),
                "gpAngabe": v["gp"], "gpEinheit": v["gp_einheit"], "abschlag": v.get("abschlag"),
                "abschlagFaellig": v.get("abschlag_faellig") or "", "preisgarantieBis": iso(v.get("preisgarantie_bis")),
                "laufzeit": v.get("laufzeit") or "", "status": v.get("status") or "", "quelle": v.get("quelle") or ""}
    daten = {
        "erzeugt": jetzt().strftime("%Y-%m-%d %H:%M"),
        "quelle": EXCEL_NAME,
        "wohnungen": [{"name": w["name"], "kurz": w["kurz"], "zeitraum": w["zeitraum"], "personen": w["personen"],
                       "ausstattung": w["ausstattung"], "zaehlertyp": w["zaehlertyp"], "konstante": w["konstante"],
                       "farbe": w["farbe"], "hinweis": w["hinweis"]} for w in d["wohnungen"]],
        "ablesungen": [{"datum": iso(a["datum"]), "stand": a["stand"], "wohnung": a["wohnung"], "quelle": a["quelle"],
                        "bemerkung": a["bemerkung"]} for a in d["ablesungen"]],
        "vertraege": [tarif(v) for v in d["vertraege"]],
        "vergleich": [tarif(v) for v in d["vergleich"]],
        "offenLokal": sorted(d.get("offen_lokal", [])),
    }
    kopf = ("/*\n * Strom-Dashboard – Ablesungen, Verträge, Wohnungen\n"
            f" * Automatisch erzeugt aus „{EXCEL_NAME}“ am {jetzt().strftime('%d.%m.%Y %H:%M')}.\n"
            " * Bitte nicht von Hand ändern: die Excel pflegen und _claude/strom_update.py ausführen.\n */\n")
    return kopf + "window.STROM = " + json.dumps(daten, ensure_ascii=False, indent=1) + ";\n"


README = f"""# Stromverbrauch

Einfaches Dashboard zum Stromverbrauch unserer Wohnung (Hufelandstraße, seit Mai 2025) mit den Werten der alten Wohnung
(Robert-Koch-Straße, 2021–2024) zum Vergleich: Zählerstand eintragen, geschätzter Jahresverbrauch, zwei Diagramme, Historie.

**Live:** {PAGES_URL} (Eintragen nur mit dem persönlichen Link)

- `index.html` – Dashboard
- `messwerte.js` – Ablesungen; wird aus der lokalen Excel-Datenbasis erzeugt. Im Dashboard oder per Handy eingetragene
  Ablesungen stehen in `offenLokal`, bis sie in die Excel übernommen sind.
- `relay/worker.js` – Cloudflare Worker, der Eingaben aus dem Dashboard speichert
- `tools/strom_update.py` – Update-Skript (Excel, messwerte.js, Veröffentlichung, Handy-Weg)

Berechnung: Verbrauch pro Tag = Differenz zweier Zählerstände / Tage; geschätzter Jahresverbrauch = Ø pro Tag × 365;
Jahreswerte linear zwischen den Ablesungen interpoliert, angefangene Jahre hochgerechnet.
"""


# --------------------------------------------------------------------------------------
# GitHub
# --------------------------------------------------------------------------------------
class GitHub:
    def __init__(self, token: str):
        self.token = token

    def api(self, method, path, data=None):
        req = urllib.request.Request(
            "https://api.github.com" + path, method=method,
            data=None if data is None else json.dumps(data).encode("utf-8"),
            headers={"Authorization": f"Bearer {self.token}", "Accept": "application/vnd.github+json",
                     "X-GitHub-Api-Version": "2022-11-28", "User-Agent": "strom-update"})
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                body = r.read()
                return r.status, (json.loads(body) if body else None)
        except urllib.error.HTTPError as e:
            body = e.read()
            try:
                return e.code, json.loads(body)
            except Exception:  # noqa: BLE001
                return e.code, {"message": body.decode("utf-8", "replace")}

    def ok(self, res, was):
        s, j = res
        if s >= 300:
            raise RuntimeError(f"GitHub-Fehler bei {was}: {s} {j.get('message') if isinstance(j, dict) else j}")
        return j

    def repo_sicherstellen(self):
        s, j = self.api("GET", f"/repos/{GH_OWNER}/{GH_REPO}")
        if s == 404:
            self.ok(self.api("POST", "/user/repos", {
                "name": GH_REPO, "description": REPO_BESCHREIBUNG, "homepage": PAGES_URL, "private": False,
                "auto_init": True, "has_issues": False, "has_wiki": False, "has_projects": False}), "Repo anlegen")
            for _ in range(10):
                time.sleep(2)
                s, j = self.api("GET", f"/repos/{GH_OWNER}/{GH_REPO}/git/ref/heads/main")
                if s == 200:
                    break
            print(f"Repo angelegt: https://github.com/{GH_OWNER}/{GH_REPO}")
        elif s >= 300:
            raise RuntimeError(f"Repo nicht erreichbar: {s} {j}")

    def pages_sicherstellen(self):
        s, j = self.api("GET", f"/repos/{GH_OWNER}/{GH_REPO}/pages")
        if s == 404:
            self.ok(self.api("POST", f"/repos/{GH_OWNER}/{GH_REPO}/pages",
                             {"source": {"branch": "main", "path": "/"}}), "GitHub Pages aktivieren")
            print("GitHub Pages aktiviert")

    def dateien_pushen(self, dateien: dict, nachricht: str) -> str | None:
        ref = self.ok(self.api("GET", f"/repos/{GH_OWNER}/{GH_REPO}/git/ref/heads/main"), "Ref lesen")
        basis = ref["object"]["sha"]
        commit = self.ok(self.api("GET", f"/repos/{GH_OWNER}/{GH_REPO}/git/commits/{basis}"), "Commit lesen")
        items = []
        for pfad, inhalt in dateien.items():
            blob = self.ok(self.api("POST", f"/repos/{GH_OWNER}/{GH_REPO}/git/blobs",
                                    {"content": base64.b64encode(inhalt).decode(), "encoding": "base64"}), f"Blob {pfad}")
            items.append({"path": pfad, "mode": "100644", "type": "blob", "sha": blob["sha"]})
        baum = self.ok(self.api("POST", f"/repos/{GH_OWNER}/{GH_REPO}/git/trees",
                                {"base_tree": commit["tree"]["sha"], "tree": items}), "Tree")
        if baum["sha"] == commit["tree"]["sha"]:
            return None
        neu = self.ok(self.api("POST", f"/repos/{GH_OWNER}/{GH_REPO}/git/commits",
                               {"message": nachricht, "tree": baum["sha"], "parents": [basis]}), "Commit")
        self.ok(self.api("PATCH", f"/repos/{GH_OWNER}/{GH_REPO}/git/refs/heads/main", {"sha": neu["sha"]}), "Ref setzen")
        return neu["sha"]


def veroeffentlichen(base: Path, nachricht: str):
    token_datei = base / "github-token.txt"
    if not token_datei.exists():
        raise SystemExit(f"Token fehlt: {token_datei}")
    gh = GitHub(token_datei.read_text(encoding="utf-8").strip())
    gh.repo_sicherstellen()
    dash = base / "Dashboard"
    dateien = {
        "index.html": (dash / "index.html").read_bytes(),
        "messwerte.js": (dash / "messwerte.js").read_bytes(),
        "README.md": README.encode("utf-8"),
        ".nojekyll": b"",
    }
    skript = base / "_claude" / "strom_update.py"
    if skript.exists():
        dateien["tools/strom_update.py"] = skript.read_bytes()
    commit = gh.dateien_pushen(dateien, nachricht)
    gh.pages_sicherstellen()
    stand = sync_stand(base)
    for rpfad in SYNC_DATEIEN:
        if rpfad in dateien:
            stand[rpfad] = sha(dateien[rpfad])
    sync_merken(base, stand)
    if commit:
        print(f"Gepusht: {commit[:7]} – {nachricht}")
    else:
        print("GitHub: keine Änderungen")
    print(f"Dashboard: {PAGES_URL}")



# --------------------------------------------------------------------------------------
# Handy/Cloud-Weg und Abgleich
# --------------------------------------------------------------------------------------
def daten_aus_messwerte(text: str) -> dict:
    """Liest messwerte.js (window.STROM = {...};) zurück in die interne Datenstruktur."""
    i = text.index("window.STROM = ") + len("window.STROM = ")
    j = text.rindex(";")
    s_ = json.loads(text[i:j])
    d = {"wohnungen": [], "ablesungen": [], "vertraege": [], "vergleich": [],
         "offen_lokal": list(s_.get("offenLokal", []))}
    for w in s_["wohnungen"]:
        d["wohnungen"].append({"name": w["name"], "kurz": w.get("kurz", ""), "zeitraum": w.get("zeitraum", ""),
                               "personen": w.get("personen", 0), "ausstattung": w.get("ausstattung", ""),
                               "zaehlertyp": w.get("zaehlertyp", ""), "zaehlernr": "", "konstante": w.get("konstante"),
                               "farbe": w.get("farbe", "blau"), "hinweis": w.get("hinweis", "")})
    for a in s_["ablesungen"]:
        d["ablesungen"].append({"datum": to_date(a["datum"]), "stand": float(a["stand"]), "wohnung": a["wohnung"],
                                "quelle": a.get("quelle", ""), "bemerkung": a.get("bemerkung", "")})
    for key in ("vertraege", "vergleich"):
        for v in s_.get(key, []):
            d[key].append({"bezeichnung": v["bezeichnung"], "anbieter": v["anbieter"], "tarif": v["tarif"],
                           "wohnung": v["wohnung"], "ab": to_date(v["ab"]), "bis": to_date(v["bis"]), "ap": v["ap"],
                           "gp": v["gpAngabe"], "gp_einheit": v["gpEinheit"], "abschlag": v.get("abschlag"),
                           "abschlag_faellig": v.get("abschlagFaellig", ""),
                           "preisgarantie_bis": to_date(v.get("preisgarantieBis")), "laufzeit": v.get("laufzeit", ""),
                           "status": v.get("status", ""), "quelle": v.get("quelle", "")})
    return d


def neue_ablesung(d: dict, datum_txt: str, stand_txt: str, quelle: str, bemerkung: str, wohnung: str | None,
                  ersetzen: bool) -> dict:
    datum, stand = to_date(datum_txt), to_float(stand_txt)
    wohnung = wohnung or d["ablesungen"][-1]["wohnung"]
    vorhanden = [a for a in d["ablesungen"] if a["wohnung"] == wohnung and a["datum"] == datum]
    if vorhanden and not ersetzen:
        raise SystemExit(f"Für {de_datum(datum)} gibt es schon eine Ablesung ({vorhanden[0]['stand']}). "
                         "Mit --ersetzen überschreiben.")
    d["ablesungen"] = [a for a in d["ablesungen"] if not (a["wohnung"] == wohnung and a["datum"] == datum)]
    neu = {"datum": datum, "stand": stand, "wohnung": wohnung, "quelle": quelle, "bemerkung": bemerkung}
    d["ablesungen"].append(neu)
    return neu


def kurzbericht(d: dict):
    akt = d["ablesungen"][-1]["wohnung"]
    a_akt = [a for a in d["ablesungen"] if a["wohnung"] == akt]
    if len(a_akt) >= 2:
        tage = (a_akt[-1]["datum"] - a_akt[0]["datum"]).days
        kwh = a_akt[-1]["stand"] - a_akt[0]["stand"]
        print(f"{akt}: {de_zahl(kwh)} kWh in {tage} Tagen = {de_zahl(kwh / tage, 2)} kWh/Tag "
              f"(≈ {de_zahl(kwh / tage * 365, 0)} kWh/Jahr)")
        p, lz = a_akt[-2], a_akt[-1]
        t2 = (lz["datum"] - p["datum"]).days
        if t2 > 0:
            print(f"Letztes Intervall {de_datum(p['datum'])}–{de_datum(lz['datum'])}: "
                  f"{de_zahl(lz['stand'] - p['stand'])} kWh, {de_zahl((lz['stand'] - p['stand']) / t2, 2)} kWh/Tag")


def git(repo: Path, *args, token: str | None = None) -> str:
    cmd = ["git", "-C", str(repo)]
    if token:
        auth = base64.b64encode(f"x-access-token:{token}".encode()).decode()
        cmd += ["-c", "credential.helper=", "-c", f"http.extraHeader=Authorization: Basic {auth}"]
    r = subprocess.run(cmd + list(args), capture_output=True, text=True, timeout=120,
                       env=dict(os.environ, GIT_TERMINAL_PROMPT="0"))
    if r.returncode != 0:
        fehler = (r.stderr or r.stdout).replace(token or "\0", "***")
        raise RuntimeError(f"git {' '.join(args)} fehlgeschlagen: {fehler.strip()}")
    return r.stdout


def cloud_main(args):
    """Handy-Weg: Ablesung direkt in messwerte.js des geklonten Repos eintragen und pushen (ohne Excel)."""
    repo = Path(args.cloud).resolve()
    token = Path(args.token_datei).read_text(encoding="utf-8").strip()
    git(repo, "pull", "--ff-only", "-q")
    datei = repo / "messwerte.js"
    d = daten_aus_messwerte(datei.read_text(encoding="utf-8"))
    pruefen(d)
    if not args.neu:
        raise SystemExit("--cloud braucht --neu DATUM STAND")
    neu = neue_ablesung(d, args.neu[0], args.neu[1], args.quelle, args.bemerkung, args.wohnung, args.ersetzen)
    for w in pruefen(d):
        print("WARNUNG:", w)
    d["offen_lokal"] = sorted(set(d.get("offen_lokal", [])) | {neu["datum"].isoformat()})
    datei.write_text(messwerte_js(d), encoding="utf-8")
    kurzbericht(d)
    msg = args.nachricht or f"Ablesung {de_datum(neu['datum'])}: {de_zahl(neu['stand'])} kWh (per Handy)"
    git(repo, "add", "messwerte.js")
    git(repo, "-c", "user.name=fbw159", "-c", "user.email=fbw159@users.noreply.github.com", "commit", "-q", "-m", msg)
    git(repo, "push", "-q", "origin", "HEAD:main", token=token)
    print(f"Gepusht: {msg}")
    print(f"Dashboard: {PAGES_URL}")
    print(f"Noch nicht in der Excel: {', '.join(d['offen_lokal'])}")


def online_datei(base: Path, pfad: str) -> bytes | None:
    token = (base / "github-token.txt").read_text(encoding="utf-8").strip()
    req = urllib.request.Request(f"https://api.github.com/repos/{GH_OWNER}/{GH_REPO}/contents/{pfad}",
                                 headers={"Authorization": f"Bearer {token}", "Accept": "application/vnd.github.raw",
                                          "User-Agent": "strom-update"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.read()
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise


SYNC_DATEIEN = {  # Repo-Pfad -> lokaler Pfad (relativ zum Basisordner)
    "index.html": "Dashboard/index.html",
    "tools/strom_update.py": "_claude/strom_update.py",
}


def sync_stand(base: Path) -> dict:
    f = base / "_claude" / "sync.json"
    try:
        return json.loads(f.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return {}


def sync_merken(base: Path, stand: dict):
    (base / "_claude" / "sync.json").write_text(json.dumps(stand, indent=1, ensure_ascii=False), encoding="utf-8")


def abgleich(base: Path, d: dict) -> list[dict]:
    """Übernimmt Ablesungen, die nur online stehen, und aktualisiert index.html/Skript aus dem Repo."""
    online = online_datei(base, "messwerte.js")
    neu = []
    if online:
        od = daten_aus_messwerte(online.decode("utf-8"))
        namen = {w["name"] for w in d["wohnungen"]}
        vorhanden = {(a["wohnung"], a["datum"]) for a in d["ablesungen"]}
        for a in od["ablesungen"]:
            if (a["wohnung"], a["datum"]) not in vorhanden and a["wohnung"] in namen:
                d["ablesungen"].append(a)
                neu.append(a)
                print(f"Vom Handy übernommen: {de_datum(a['datum'])} {de_zahl(a['stand'])} kWh ({a['quelle']})")
    stand = sync_stand(base)
    for rpfad, lpfad in SYNC_DATEIEN.items():
        ob = online_datei(base, rpfad)
        lokal = base / lpfad
        if ob is None:
            continue
        lb = lokal.read_bytes() if lokal.exists() else None
        if lb == ob:
            stand[rpfad] = sha(ob)
            continue
        basis = stand.get(rpfad)
        lokal_geaendert = lb is not None and basis is not None and sha(lb) != basis
        online_geaendert = basis is None or sha(ob) != basis
        if lokal_geaendert and online_geaendert:
            print(f"WARNUNG: {lpfad} lokal und online geändert – lokale Fassung bleibt, online wird überschrieben.")
            continue
        if online_geaendert:
            if lb is not None:
                bdir = base / "_claude" / "backup"
                bdir.mkdir(parents=True, exist_ok=True)
                (bdir / f"{Path(lpfad).stem}_{jetzt().strftime('%Y%m%d-%H%M%S')}{Path(lpfad).suffix}").write_bytes(lb)
            lokal.parent.mkdir(parents=True, exist_ok=True)
            lokal.write_bytes(ob)
            stand[rpfad] = sha(ob)
            print(f"Aus dem Repo aktualisiert: {lpfad}")
    sync_merken(base, stand)
    return neu

# --------------------------------------------------------------------------------------
# Hauptprogramm
# --------------------------------------------------------------------------------------
def backup(pfad: Path, base: Path):
    if not pfad.exists():
        return
    bdir = base / "_claude" / "backup"
    bdir.mkdir(exist_ok=True)
    ziel = bdir / f"{pfad.stem}_{jetzt().strftime('%Y%m%d-%H%M%S')}{pfad.suffix}"
    shutil.copy2(pfad, ziel)
    alte = sorted(bdir.glob(f"{pfad.stem}_*{pfad.suffix}"))
    for f in alte[:-BACKUPS_BEHALTEN]:
        try:
            f.unlink()
        except OSError:
            pass


def main(argv=None):
    ap = argparse.ArgumentParser(description="Strom-Datenbasis und Dashboard aktualisieren")
    ap.add_argument("--neu", nargs=2, metavar=("DATUM", "STAND"), help="neue Ablesung, z. B. 2026-12-01 60512,3")
    ap.add_argument("--quelle", default="", help="Quelle der neuen Ablesung (z. B. Fotoname)")
    ap.add_argument("--bemerkung", default="")
    ap.add_argument("--wohnung", default=None, help="Standard: Wohnung der letzten Ablesung")
    ap.add_argument("--ersetzen", action="store_true", help="vorhandene Ablesung am selben Tag ersetzen")
    ap.add_argument("--init", metavar="JSON", help="Erstaufbau aus JSON statt aus der Excel")
    ap.add_argument("--push", action="store_true", help="Dashboard danach auf GitHub veröffentlichen")
    ap.add_argument("--nur-push", action="store_true", help="wie --push")
    ap.add_argument("--nachricht", default=None, help="Commit-Nachricht")
    ap.add_argument("--ohne-recalc", action="store_true", help="LibreOffice-Neuberechnung überspringen")
    ap.add_argument("--ohne-abgleich", action="store_true", help="keine Handy-Ablesungen/Dateien von GitHub holen")
    ap.add_argument("--ordner", default=None, help="Basisordner (Standard: Ordner über _claude)")
    ap.add_argument("--cloud", metavar="REPO_ORDNER", help="Handy-Weg: nur messwerte.js im Repo ändern und pushen")
    ap.add_argument("--token-datei", default=None, help="Token-Datei für --cloud")
    args = ap.parse_args(argv)

    if args.cloud:
        cloud_main(args)
        return

    base = Path(args.ordner).resolve() if args.ordner else BASE
    excel = base / EXCEL_NAME
    dash = base / "Dashboard"
    push = args.push or args.nur_push

    if (base / f"~${EXCEL_NAME}").exists():
        raise SystemExit("Die Excel ist gerade geöffnet (~$-Sperrdatei vorhanden) – bitte schließen und erneut starten.")

    d = daten_aus_json(Path(args.init)) if args.init else daten_aus_excel(excel)
    pruefen(d)
    abgleich_ok = False
    if not args.ohne_abgleich and not args.init:
        try:
            uebernommen = abgleich(base, d)
            abgleich_ok = True
            if not uebernommen:
                print("Abgleich: keine neuen Handy-Ablesungen")
        except Exception as e:  # noqa: BLE001
            print(f"WARNUNG: Abgleich mit GitHub nicht möglich ({e})")
    d["offen_lokal"] = []
    neu = None
    if args.neu:
        neu = neue_ablesung(d, args.neu[0], args.neu[1], args.quelle, args.bemerkung, args.wohnung, args.ersetzen)
    for w in pruefen(d):
        print("WARNUNG:", w)

    backup(excel, base)
    excel_erzeugen(d, excel)
    print(f"Excel geschrieben: {excel}")
    if not args.ohne_recalc:
        print("Neuberechnung:", neuberechnen(excel))

    dash.mkdir(exist_ok=True)
    (dash / "messwerte.js").write_text(messwerte_js(d), encoding="utf-8")
    print(f"messwerte.js geschrieben ({len(d['ablesungen'])} Ablesungen)")
    kurzbericht(d)

    if push:
        if not abgleich_ok and not args.ohne_abgleich and not args.init:
            raise SystemExit("Nicht veröffentlicht: ohne Abgleich könnten Handy-Ablesungen online überschrieben werden.")
        if args.nachricht:
            msg = args.nachricht
        elif neu:
            msg = f"Ablesung {de_datum(neu['datum'])}: {de_zahl(neu['stand'])} kWh"
        else:
            msg = "Dashboard aktualisieren"
        veroeffentlichen(base, msg)


if __name__ == "__main__":
    main()
