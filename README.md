# Stromverbrauch

Einfaches Dashboard zum Stromverbrauch unserer Wohnung (Hufelandstraße, seit Mai 2025) mit den Werten der alten Wohnung
(Robert-Koch-Straße, 2021–2024) zum Vergleich: Zählerstand eintragen, geschätzter Jahresverbrauch, zwei Diagramme, Historie.

**Live:** https://fbw159.github.io/stromverbrauch/ (Eintragen nur mit dem persönlichen Link)

- `index.html` – Dashboard
- `messwerte.js` – Ablesungen; wird aus der lokalen Excel-Datenbasis erzeugt. Im Dashboard oder per Handy eingetragene
  Ablesungen stehen in `offenLokal`, bis sie in die Excel übernommen sind.
- `relay/worker.js` – Cloudflare Worker, der Eingaben aus dem Dashboard speichert
- `tools/strom_update.py` – Update-Skript (Excel, messwerte.js, Veröffentlichung, Handy-Weg)

Berechnung: Verbrauch pro Tag = Differenz zweier Zählerstände / Tage; geschätzter Jahresverbrauch = Ø pro Tag × 365;
Jahreswerte linear zwischen den Ablesungen interpoliert, angefangene Jahre hochgerechnet.
