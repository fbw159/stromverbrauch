# Worker „stromverbrauch“

`https://stromverbrauch.fbw159.workers.dev` nimmt Zählerstände aus dem Dashboard entgegen und schreibt sie in
`messwerte.js` dieses Repos (GitHub-Contents-API, jede Eingabe = ein Commit, Konflikte per `sha` mit Wiederholung).

- `GET /status` → `{ ok, letzte, daten }` (prüft den Zugangscode, liefert die aktuellen Daten)
- `POST /ablesung` `{ datum, stand, ersetzen? }` → trägt eine Ablesung für die aktuelle Wohnung ein
- `DELETE /ablesung` `{ datum }` → löscht eine online eingetragene Ablesung, solange sie nicht in der Excel steht
- Zugang: Header `x-key` mit dem Code aus dem persönlichen Link (`#k=…`); im Code steht nur dessen SHA-256.
- Secret: `GITHUB_TOKEN` (Schreibrecht Contents auf dieses Repo). Code: `worker.js` (ES-Modul, `compatibility_date` 2026-10-01).
- Online eingetragene Ablesungen stehen in `offenLokal`, bis das Skript am PC sie in die Excel übernommen hat.

Prüfen: `GET /status` ohne Header → 401; mit gültigem `x-key` → 200.
