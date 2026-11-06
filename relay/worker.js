/*
 * Cloudflare Worker „stromverbrauch“ – speichert Zählerstände aus dem Dashboard in messwerte.js
 * (Repo fbw159/stromverbrauch) über die GitHub-Contents-API. Jede Eingabe = ein Commit.
 *
 *   GET    /status    -> { ok, letzte, daten }              Zugangscode prüfen, aktuelle Daten holen
 *   POST   /ablesung  { datum, stand, ersetzen? }           Ablesung eintragen
 *   DELETE /ablesung  { datum }                             online eingetragene Ablesung löschen
 *                                                           (nur solange sie noch nicht in der Excel ist)
 * Zugang: Header „x-key“; hier steht nur der SHA-256 des Codes. Secret: GITHUB_TOKEN.
 */
const REPO = 'fbw159/stromverbrauch';
const DATEI = 'messwerte.js';
const CODE_SHA256 = '2d2a1d337c3a2333825e37d9233da126373cd7630471baf3788a1ef04f051267';
const ORIGINS = ['https://fbw159.github.io', 'null'];
const MAX_PRO_TAG = 60; // kWh – darüber ist die Eingabe sicher ein Tippfehler

function cors(req) {
  const o = req.headers.get('Origin');
  return {
    'Access-Control-Allow-Origin': ORIGINS.includes(o) ? o : ORIGINS[0],
    'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, x-key',
    'Access-Control-Max-Age': '86400',
    'Vary': 'Origin'
  };
}
function antwort(req, status, obj) {
  return new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json; charset=utf-8', ...cors(req) } });
}
async function sha256(text) {
  const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join('');
}
function b64decode(b64) {
  const bin = atob(b64.replace(/\n/g, ''));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}
function b64encode(text) {
  const bytes = new TextEncoder().encode(text);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}
function jetztBerlin() {
  const p = new Intl.DateTimeFormat('sv-SE', { timeZone: 'Europe/Berlin', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false }).format(new Date());
  return p.replace(',', '').slice(0, 16); // „JJJJ-MM-TT HH:MM“
}
const tagNr = iso => Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) / 864e5;
const deDatum = iso => iso.slice(8, 10) + '.' + iso.slice(5, 7) + '.' + iso.slice(0, 4);
const deZahl = (x, d) => x.toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d });

async function gh(env, method, pfad, body) {
  const r = await fetch('https://api.github.com/repos/' + REPO + pfad, {
    method,
    headers: { 'Authorization': 'Bearer ' + env.GITHUB_TOKEN, 'Accept': 'application/vnd.github+json', 'User-Agent': 'stromverbrauch-worker', 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined
  });
  return { status: r.status, json: await r.json().catch(() => ({})) };
}
async function laden(env) {
  const r = await gh(env, 'GET', '/contents/' + DATEI + '?ref=main');
  if (r.status !== 200) throw new Error('GitHub ' + r.status);
  const text = b64decode(r.json.content);
  const i = text.indexOf('window.STROM = '), j = text.lastIndexOf(';');
  return { sha: r.json.sha, daten: JSON.parse(text.slice(i + 15, j)) };
}
function serialisieren(daten) {
  return '/*\n * Strom-Dashboard – Ablesungen, Verträge, Wohnungen\n' +
    ' * Stand ' + daten.erzeugt + ' (zuletzt per Dashboard-Eingabe geändert).\n' +
    ' * Datenbasis ist die Excel auf dem PC; online eingetragene Ablesungen stehen in offenLokal, bis sie übernommen sind.\n */\n' +
    'window.STROM = ' + JSON.stringify(daten, null, 1) + ';\n';
}
function sortieren(daten) {
  const erste = {};
  daten.ablesungen.forEach(a => { if (!(a.wohnung in erste) || a.datum < erste[a.wohnung]) erste[a.wohnung] = a.datum; });
  daten.ablesungen.sort((a, b) => (erste[a.wohnung] < erste[b.wohnung] ? -1 : erste[a.wohnung] > erste[b.wohnung] ? 1 : a.datum < b.datum ? -1 : a.datum > b.datum ? 1 : 0));
}
function aktuelleWohnung(daten) {
  let w = null, d = '';
  daten.ablesungen.forEach(a => { if (a.datum >= d) { d = a.datum; w = a.wohnung; } });
  return w;
}
function letzte(daten) {
  const w = aktuelleWohnung(daten);
  const l = daten.ablesungen.filter(a => a.wohnung === w).sort((a, b) => (a.datum < b.datum ? -1 : 1)).pop();
  return l ? { datum: l.datum, stand: l.stand, wohnung: w } : null;
}

// Änderung anwenden; liefert { fehler, status } oder { nachricht }
function eintragen(daten, eingabe) {
  const datum = String(eingabe.datum || ''), stand = Number(eingabe.stand);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(datum) || isNaN(tagNr(datum))) return { status: 400, fehler: 'Datum fehlt oder ist ungültig.' };
  const heute = jetztBerlin().slice(0, 10);
  if (datum > heute) return { status: 400, fehler: 'Das Datum liegt in der Zukunft.' };
  if (!isFinite(stand) || stand <= 0 || stand > 9999999) return { status: 400, fehler: 'Zählerstand ist keine gültige Zahl.' };
  const wohnung = aktuelleWohnung(daten);
  const liste = daten.ablesungen.filter(a => a.wohnung === wohnung).sort((a, b) => (a.datum < b.datum ? -1 : 1));
  if (datum < liste[0].datum) return { status: 400, fehler: 'Das Datum liegt vor der ersten Ablesung dieser Wohnung.' };
  const gleich = liste.find(a => a.datum === datum);
  if (gleich && !eingabe.ersetzen) return { status: 409, fehler: 'Für den ' + deDatum(datum) + ' gibt es schon ' + deZahl(gleich.stand, 1) + ' kWh.', vorhanden: gleich.stand };
  if (gleich && !(daten.offenLokal || []).includes(datum)) return { status: 409, fehler: 'Diese Ablesung steht schon in der Excel – bitte dort ändern.' };
  const vor = liste.filter(a => a.datum < datum).pop(), nach = liste.find(a => a.datum > datum);
  if (vor && stand < vor.stand) return { status: 422, fehler: 'Kleiner als der Stand vom ' + deDatum(vor.datum) + ' (' + deZahl(vor.stand, 1) + ' kWh).' };
  if (nach && stand > nach.stand) return { status: 422, fehler: 'Größer als der Stand vom ' + deDatum(nach.datum) + ' (' + deZahl(nach.stand, 1) + ' kWh).' };
  if (vor) {
    const tage = tagNr(datum) - tagNr(vor.datum);
    if (tage > 0 && (stand - vor.stand) / tage > MAX_PRO_TAG) return { status: 422, fehler: 'Das wären ' + deZahl((stand - vor.stand) / tage, 1) + ' kWh pro Tag – bitte prüfen.' };
  }
  daten.ablesungen = daten.ablesungen.filter(a => !(a.wohnung === wohnung && a.datum === datum));
  daten.ablesungen.push({ datum, stand: Math.round(stand * 10) / 10, wohnung, quelle: 'Eingabe im Dashboard', bemerkung: '' });
  sortieren(daten);
  daten.offenLokal = [...new Set([...(daten.offenLokal || []), datum])].sort();
  return { nachricht: 'Ablesung ' + deDatum(datum) + ': ' + deZahl(Math.round(stand * 10) / 10, 1) + ' kWh (Dashboard)' };
}
function loeschen(daten, eingabe) {
  const datum = String(eingabe.datum || '');
  const offen = daten.offenLokal || [];
  if (!offen.includes(datum)) return { status: 409, fehler: 'Nur online eingetragene Ablesungen, die noch nicht in der Excel stehen, können hier gelöscht werden.' };
  const wohnung = aktuelleWohnung(daten);
  const vorher = daten.ablesungen.length;
  daten.ablesungen = daten.ablesungen.filter(a => !(a.wohnung === wohnung && a.datum === datum));
  if (daten.ablesungen.length === vorher) return { status: 404, fehler: 'Ablesung nicht gefunden.' };
  daten.offenLokal = offen.filter(d => d !== datum);
  return { nachricht: 'Ablesung ' + deDatum(datum) + ' gelöscht (Dashboard)' };
}

export default {
  async fetch(req, env) {
    if (req.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors(req) });
    const pfad = new URL(req.url).pathname.replace(/\/+$/, '');
    const code = req.headers.get('x-key') || '';
    if (!code || (await sha256(code)) !== CODE_SHA256) return antwort(req, 401, { fehler: 'Zugangscode fehlt oder ist falsch.' });
    try {
      if (req.method === 'GET' && pfad === '/status') {
        const { daten } = await laden(env);
        return antwort(req, 200, { ok: true, letzte: letzte(daten), daten });
      }
      if (pfad === '/ablesung' && (req.method === 'POST' || req.method === 'DELETE')) {
        const eingabe = await req.json().catch(() => ({}));
        for (let versuch = 0; versuch < 3; versuch++) {
          const { sha, daten } = await laden(env);
          const erg = req.method === 'POST' ? eintragen(daten, eingabe) : loeschen(daten, eingabe);
          if (erg.fehler) return antwort(req, erg.status, erg);
          daten.erzeugt = jetztBerlin();
          const r = await gh(env, 'PUT', '/contents/' + DATEI, { message: erg.nachricht, content: b64encode(serialisieren(daten)), sha, branch: 'main' });
          if (r.status === 200 || r.status === 201) return antwort(req, 200, { ok: true, daten, commit: r.json.commit && r.json.commit.sha });
          if (r.status !== 409 && r.status !== 422) return antwort(req, 502, { fehler: 'GitHub ' + r.status + ': ' + (r.json.message || '') });
        }
        return antwort(req, 409, { fehler: 'Gleichzeitige Änderung – bitte noch einmal speichern.' });
      }
      return antwort(req, 404, { fehler: 'Unbekannte Adresse.' });
    } catch (e) {
      return antwort(req, 502, { fehler: 'Speichern nicht möglich: ' + e.message });
    }
  }
};
