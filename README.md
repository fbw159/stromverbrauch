# Stromverbrauch

Dashboard zum Stromverbrauch unserer Wohnung (Hufelandstraße, seit Mai 2025) mit Vergleich zur alten Wohnung
(Robert-Koch-Straße, Ablesungen 2021–2024). Beide Wohnungen: 2 Personen, Fernwärme, Warmwasser per Durchlauferhitzer.

**Live:** https://fbw159.github.io/stromverbrauch/

- `index.html` – Dashboard (Kennzahlen, Verlauf, Saison- und Jahresvergleich, Kosten & Abschlag-Check,
  Spielwiese mit Tarif-, Dusch- und Geräte-Rechner sowie Zählerscheiben-Stoppuhr, Tabellen)
- `messwerte.js` – Ablesungen, Verträge und Wohnungsdaten; wird automatisch aus der lokalen Excel-Datenbasis erzeugt

Berechnung:

- Verbrauch je Ableseintervall = Differenz der Zählerstände, Ø kWh/Tag = Verbrauch / Tage, Hochrechnung = Ø × 365
- Monats- und Jahreswerte: Zählerstände linear zwischen den Ablesungen interpoliert
- Kosten = Verbrauch × Arbeitspreis + Tage × Grundpreis/365 (Preise brutto laut Vertrag)
- Prognose Abrechnungsjahr = gemessener Verbrauch seit Lieferbeginn + Ø kWh/Tag × Resttage
