/*
 * Strom-Dashboard – Ablesungen, Verträge, Wohnungen
 * Stand 2026-11-06 09:50 (zuletzt per Dashboard-Eingabe geändert).
 * Datenbasis ist die Excel auf dem PC; online eingetragene Ablesungen stehen in offenLokal, bis sie übernommen sind.
 */
window.STROM = {
 "erzeugt": "2026-11-06 09:50",
 "quelle": "Stromzaehlerwerte.xlsx",
 "wohnungen": [
  {
   "name": "Robert-Koch-Straße",
   "kurz": "Alte Wohnung",
   "zeitraum": "bis 05/2025 (Ablesungen 08/2021 – 09/2024)",
   "personen": 2,
   "ausstattung": "Fernwärme, Warmwasser per elektrischem Durchlauferhitzer",
   "zaehlertyp": "",
   "konstante": null,
   "farbe": "orange",
   "hinweis": "Werte aus „Zählerstände_Robert-Koch_Str..xlsx“; gleiche Geräte und Wohnsituation (2 Personen) wie heute – daher vergleichbar"
  },
  {
   "name": "Hufelandstraße",
   "kurz": "Aktuelle Wohnung",
   "zeitraum": "seit 05/2025 (Ablesungen ab 04/2026)",
   "personen": 2,
   "ausstattung": "Fernwärme, Warmwasser per elektrischem Durchlauferhitzer",
   "zaehlertyp": "Siemens 7CA54 Ferraris-Drehstromzähler",
   "konstante": 75,
   "farbe": "blau",
   "hinweis": "Einzug Mai 2025; Einzugszählerstand fehlt noch (steht in der SWDU-Jahresabrechnung)"
  }
 ],
 "ablesungen": [
  {
   "datum": "2021-08-04",
   "stand": 18764,
   "wohnung": "Robert-Koch-Straße",
   "quelle": "Excel „Zählerstände_Robert-Koch_Str.“",
   "bemerkung": ""
  },
  {
   "datum": "2021-12-31",
   "stand": 19766,
   "wohnung": "Robert-Koch-Straße",
   "quelle": "Excel „Zählerstände_Robert-Koch_Str.“",
   "bemerkung": ""
  },
  {
   "datum": "2022-03-02",
   "stand": 20174,
   "wohnung": "Robert-Koch-Straße",
   "quelle": "Excel „Zählerstände_Robert-Koch_Str.“",
   "bemerkung": ""
  },
  {
   "datum": "2022-03-03",
   "stand": 20174,
   "wohnung": "Robert-Koch-Straße",
   "quelle": "Excel „Zählerstände_Robert-Koch_Str.“",
   "bemerkung": "gleicher Stand wie am Vortag – vermutlich doppelt erfasst"
  },
  {
   "datum": "2022-03-31",
   "stand": 20398,
   "wohnung": "Robert-Koch-Straße",
   "quelle": "Excel „Zählerstände_Robert-Koch_Str.“",
   "bemerkung": ""
  },
  {
   "datum": "2022-06-30",
   "stand": 21100,
   "wohnung": "Robert-Koch-Straße",
   "quelle": "Excel „Zählerstände_Robert-Koch_Str.“",
   "bemerkung": ""
  },
  {
   "datum": "2022-12-31",
   "stand": 22521,
   "wohnung": "Robert-Koch-Straße",
   "quelle": "Excel „Zählerstände_Robert-Koch_Str.“",
   "bemerkung": ""
  },
  {
   "datum": "2023-03-03",
   "stand": 23000,
   "wohnung": "Robert-Koch-Straße",
   "quelle": "Excel „Zählerstände_Robert-Koch_Str.“",
   "bemerkung": ""
  },
  {
   "datum": "2023-07-21",
   "stand": 23994,
   "wohnung": "Robert-Koch-Straße",
   "quelle": "Excel „Zählerstände_Robert-Koch_Str.“",
   "bemerkung": ""
  },
  {
   "datum": "2023-08-31",
   "stand": 24398,
   "wohnung": "Robert-Koch-Straße",
   "quelle": "Excel „Zählerstände_Robert-Koch_Str.“",
   "bemerkung": ""
  },
  {
   "datum": "2024-09-01",
   "stand": 27106,
   "wohnung": "Robert-Koch-Straße",
   "quelle": "Excel „Zählerstände_Robert-Koch_Str.“",
   "bemerkung": "letzter Wert der alten Wohnung"
  },
  {
   "datum": "2026-04-24",
   "stand": 58603,
   "wohnung": "Hufelandstraße",
   "quelle": "SWDU – Bestätigung Zählerstanderfassung",
   "bemerkung": "eigene Ablesung, an SWDU gemeldet (zur Kündigung)"
  },
  {
   "datum": "2026-05-31",
   "stand": 58900.7,
   "wohnung": "Hufelandstraße",
   "quelle": "Wechselablesung Netzbetreiber (Portal: 58.901 kWh am 31.05.2026)",
   "bemerkung": "Lieferantenwechsel: Ende SWDU 31.05. = Anfangsstand grüüün 01.06.2026"
  },
  {
   "datum": "2026-11-01",
   "stand": 60282.5,
   "wohnung": "Hufelandstraße",
   "quelle": "Foto IMG_20261101_131144.jpg (13:11 Uhr)",
   "bemerkung": ""
  },
  {
   "datum": "2026-11-06",
   "stand": 60300,
   "wohnung": "Hufelandstraße",
   "quelle": "Eingabe im Dashboard",
   "bemerkung": ""
  }
 ],
 "vertraege": [
  {
   "bezeichnung": "SWDU PartnerStrom Online",
   "anbieter": "Stadtwerke Duisburg AG (SWDU)",
   "tarif": "PartnerStrom Online (später „PartnerStrom Regional 12“)",
   "wohnung": "Hufelandstraße",
   "ab": "2025-05-06",
   "bis": "2026-05-31",
   "ap": 28.4,
   "gpJahr": 113.45,
   "gpAngabe": 113.45,
   "gpEinheit": "€/Jahr",
   "abschlag": 76,
   "abschlagFaellig": "monatlich zum 20.",
   "preisgarantieBis": "2026-05-31",
   "laufzeit": "12 Monate, danach unbefristet; Kündigungsfrist 1 Monat",
   "status": "beendet – Sonderkündigung zum 31.05.2026 nach angekündigter Preiserhöhung („zu teuer“)",
   "quelle": "SWDU: Eingangsbestätigung 08.05.2025, Preisanpassung 16.04.2026, Kündigungsbestätigung 24.04.2026"
  },
  {
   "bezeichnung": "grüüün Natüüürlich10",
   "anbieter": "grüüün (Münch Energie GmbH & Co. KG)",
   "tarif": "Natüüürlich10",
   "wohnung": "Hufelandstraße",
   "ab": "2026-06-01",
   "bis": null,
   "ap": 24.82,
   "gpJahr": 92.88,
   "gpAngabe": 7.74,
   "gpEinheit": "€/Monat",
   "abschlag": 60,
   "abschlagFaellig": "monatlich zum 1.",
   "preisgarantieBis": "2027-11-30",
   "laufzeit": "12 Monate, kündbar zum 31.05.2027 (Frist 1 Monat); Preisgarantie 18 Monate, danach 102 Monate Energiepreis-Garantie",
   "status": "aktiv seit 01.06.2026",
   "quelle": "grüüün: Auftragsbestätigung 25.04.2026, Vertragsbestätigung mit Abschlag 12.05.2026"
  }
 ],
 "vergleich": [
  {
   "bezeichnung": "SWDU Preis ab 01.06.2026",
   "anbieter": "Stadtwerke Duisburg AG (SWDU)",
   "tarif": "PartnerStrom Regional 12 – angekündigter Preis ab 01.06.2026",
   "wohnung": "Hufelandstraße",
   "ab": "2026-06-01",
   "bis": null,
   "ap": 30.96,
   "gpJahr": 87.11,
   "gpAngabe": 87.11,
   "gpEinheit": "€/Jahr",
   "abschlag": null,
   "abschlagFaellig": "",
   "preisgarantieBis": null,
   "laufzeit": "",
   "status": "nicht wirksam geworden (gekündigt) – zeigt, was ohne Wechsel fällig gewesen wäre",
   "quelle": "SWDU: Preisanpassung durch Kostenänderung 16.04.2026"
  }
 ],
 "offenLokal": [
  "2026-11-06"
 ]
};
