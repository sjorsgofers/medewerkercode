// generate.js — leest het Excel-bestand en maakt een versleuteld data.js
// Gebruik:  node generate.js "pad/naar/in dienst.xlsx"
//
// Elk medewerker-record wordt versleuteld met een sleutel die is afgeleid
// van postcode + huisnummer. Alleen wie het juiste adres invoert kan zijn
// eigen record ontsleutelen. De ruwe namen/codes staan dus NIET leesbaar
// in de gepubliceerde site.

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const webcrypto = require('crypto').webcrypto;
const subtle = webcrypto.subtle;
const getRandomValues = (arr) => webcrypto.getRandomValues(arr);

const PBKDF2_ITER = 200000; // moet gelijk zijn aan index.html

// --- credential normaliseren (identiek aan de browser) ---
function normCredential(postcode, huisnummer) {
  const pc = String(postcode).replace(/\s+/g, '').toUpperCase();
  const nrMatch = String(huisnummer).match(/\d+/);
  const nr = nrMatch ? nrMatch[0] : '';
  return pc + '|' + nr;
}

async function deriveKey(credential, salt) {
  const enc = new TextEncoder();
  const base = await subtle.importKey('raw', enc.encode(credential), 'PBKDF2', false, ['deriveKey']);
  return subtle.deriveKey(
    { name: 'PBKDF2', salt, iterations: PBKDF2_ITER, hash: 'SHA-256' },
    base,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
}

function b64(bytes) {
  return Buffer.from(bytes).toString('base64');
}

(async () => {
  const file = process.argv[2];
  if (!file) {
    console.error('Geef het pad naar het Excel-bestand op.\n  node generate.js "in dienst.xlsx"');
    process.exit(1);
  }

  const wb = XLSX.readFile(file);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  const globalSalt = getRandomValues(new Uint8Array(16));
  const records = [];
  const skipped = [];

  for (const r of rows.slice(1)) {
    const code = r[0];
    const naam = r[1];
    const adres = r[2];
    if (!code || !naam || !adres) continue;

    const text = String(adres);
    // postcode: 4 cijfers + 2 letters
    const pcMatch = text.match(/(\d{4})\s*([A-Za-z]{2})/);
    // straat + huisnummer: eerste regel
    const firstLine = text.split(/\r?\n/)[0].trim();
    const nrMatch = firstLine.match(/(\d+)\s*([a-zA-Z])?\s*$/);

    if (!pcMatch || !nrMatch) {
      skipped.push(`${code} ${naam} (adres niet te parsen: "${firstLine}")`);
      continue;
    }

    const postcode = pcMatch[1] + pcMatch[2];
    const huisnummer = nrMatch[1];
    const credential = normCredential(postcode, huisnummer);

    const iv = getRandomValues(new Uint8Array(12));
    const key = await deriveKey(credential, globalSalt);
    const payload = JSON.stringify({ naam: String(naam).trim(), code: String(code).trim() });
    const ct = new Uint8Array(
      await subtle.encrypt({ name: 'AES-GCM', iv }, key, new TextEncoder().encode(payload))
    );

    records.push({ iv: b64(iv), ct: b64(ct) });
  }

  // schud de volgorde zodat positie niets verraadt
  for (let i = records.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [records[i], records[j]] = [records[j], records[i]];
  }

  const data = {
    salt: b64(globalSalt),
    iter: PBKDF2_ITER,
    records,
  };

  const out = 'window.MEDEWERKER_DATA = ' + JSON.stringify(data) + ';\n';
  fs.writeFileSync(path.join(__dirname, 'data.js'), out, 'utf8');

  console.log(`Klaar: ${records.length} medewerkers versleuteld -> data.js`);
  if (skipped.length) {
    console.log('\nOvergeslagen (handmatig nakijken):');
    skipped.forEach((s) => console.log('  - ' + s));
  }
})();
