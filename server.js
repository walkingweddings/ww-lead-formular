const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const DASHBOARD_PIN = '1234';
const LEADS_FILE = path.join(__dirname, 'leads.json');
const RESEND_API_KEY = process.env.RESEND_API_KEY || '';
const FROM_EMAIL = process.env.FROM_EMAIL || 'Walking Weddings <onboarding@resend.dev>';

// In-memory leads (persisted to file)
let leads = [];
try { leads = JSON.parse(fs.readFileSync(LEADS_FILE, 'utf8')); } catch {}

function saveLeads() {
  fs.writeFileSync(LEADS_FILE, JSON.stringify(leads, null, 2));
}

function esc(str) {
  return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

async function sendEmail({ to, subject, html, replyTo }) {
  if (!RESEND_API_KEY) {
    console.log('RESEND_API_KEY not set - email skipped');
    return false;
  }
  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${RESEND_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ from: FROM_EMAIL, to: [to], subject, html, reply_to: replyTo || undefined })
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.message || JSON.stringify(data));
  return true;
}

// MIME types
const MIME = {
  '.html': 'text/html',
  '.css': 'text/css',
  '.js': 'application/javascript',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.json': 'application/json'
};

function serveStatic(req, res) {
  let filePath = req.url === '/' ? '/index.html' : req.url.split('?')[0];
  filePath = path.join(__dirname, 'public', decodeURIComponent(filePath));

  const ext = path.extname(filePath);
  const contentType = MIME[ext] || 'application/octet-stream';

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found');
      return;
    }
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(data);
  });
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try { resolve(JSON.parse(body)); }
      catch (e) { reject(e); }
    });
  });
}

async function sendTeamEmail(lead) {
  const dateStr = lead.noDateYet ? 'Noch kein fixes Datum' : (lead.weddingDates?.join(', ') || 'Nicht angegeben');
  const locStr = lead.noLocationYet ? 'Noch keine Location' : (lead.locations?.join(', ') || 'Nicht angegeben');

  return sendEmail({
    to: 'contact@walkingweddings.com',
    subject: `Neuer Messe-Lead: ${lead.name}`,
    html: `
      <div style="text-align:center;margin-bottom:20px;">
        <img src="https://ww-lead-formular-production.up.railway.app/assets/ww_logoColor_wtagline%20(1).svg" alt="Walking Weddings" style="width:180px;height:auto;" />
      </div>
      <h2 style="font-family:Georgia,serif;color:#393e3f;">Neuer Lead - HochzeitStil Messe</h2>
      <table style="border-collapse:collapse;width:100%;max-width:500px;font-family:Arial,sans-serif;">
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Name</td><td style="padding:8px;border:1px solid #d4c4a8;">${esc(lead.name)}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Telefon</td><td style="padding:8px;border:1px solid #d4c4a8;">${esc(lead.phone)}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">E-Mail</td><td style="padding:8px;border:1px solid #d4c4a8;">${esc(lead.email)}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Hochzeitsdatum</td><td style="padding:8px;border:1px solid #d4c4a8;">${esc(dateStr)}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Location</td><td style="padding:8px;border:1px solid #d4c4a8;">${esc(locStr)}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Interesse</td><td style="padding:8px;border:1px solid #d4c4a8;">${esc(lead.interests?.join(', ') || '-')}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Stunden</td><td style="padding:8px;border:1px solid #d4c4a8;">${esc(lead.hours || '-')}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Zusatzprodukte</td><td style="padding:8px;border:1px solid #d4c4a8;">${esc(lead.extras?.join(', ') || '-')}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Budget</td><td style="padding:8px;border:1px solid #d4c4a8;">${esc(lead.budget || '-')}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Anmerkungen</td><td style="padding:8px;border:1px solid #d4c4a8;">${esc(lead.remarks || '-')}</td></tr>
      </table>
    `
  });
}

async function sendCoupleEmail(lead) {
  const dateStr = lead.noDateYet ? 'Noch kein fixes Datum' : (lead.weddingDates?.join(', ') || '-');
  const locStr = lead.noLocationYet ? 'Noch keine Location' : (lead.locations?.join(', ') || '-');

  return sendEmail({
    to: lead.email,
    replyTo: 'contact@walkingweddings.com',
    subject: 'Danke fuer euer Interesse - Walking Weddings',
    html: `
      <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#393e3f;">
        <div style="text-align:center;padding:30px 0;background:#393e3f;">
          <img src="https://ww-lead-formular-production.up.railway.app/assets/ww_logoWhite_wtagline%20(1).svg" alt="Walking Weddings" style="width:220px;height:auto;margin:0 auto;" />
        </div>
        <div style="padding:30px 20px;">
          <p>Liebe/r ${esc(lead.name)},</p>
          <p>vielen Dank fuer euer Interesse an Walking Weddings auf der HochzeitStil Messe!</p>
          <p>Wir freuen uns sehr, dass wir euch kennenlernen durften. Unser Team wird sich in Kuerze bei euch melden, um ein unverbindliches Kennenlerngespraech zu vereinbaren.</p>

          <div style="background:#f5f0e8;border:1px solid #d4c4a8;padding:20px;margin:24px 0;">
            <p style="font-family:Georgia,serif;font-size:16px;margin:0 0 12px;letter-spacing:2px;text-transform:uppercase;">Eure Angaben</p>
            <table style="border-collapse:collapse;width:100%;font-size:14px;">
              <tr><td style="padding:6px 8px;font-weight:bold;vertical-align:top;width:140px;">Datum</td><td style="padding:6px 8px;">${esc(dateStr)}</td></tr>
              <tr><td style="padding:6px 8px;font-weight:bold;vertical-align:top;">Location</td><td style="padding:6px 8px;">${esc(locStr)}</td></tr>
              <tr><td style="padding:6px 8px;font-weight:bold;vertical-align:top;">Interesse</td><td style="padding:6px 8px;">${esc(lead.interests?.join(', ') || '-')}</td></tr>
              ${lead.hours ? `<tr><td style="padding:6px 8px;font-weight:bold;vertical-align:top;">Stunden</td><td style="padding:6px 8px;">${esc(lead.hours)}</td></tr>` : ''}
              ${lead.extras?.length ? `<tr><td style="padding:6px 8px;font-weight:bold;vertical-align:top;">Zusatzprodukte</td><td style="padding:6px 8px;">${esc(lead.extras.join(', '))}</td></tr>` : ''}
              ${lead.budget ? `<tr><td style="padding:6px 8px;font-weight:bold;vertical-align:top;">Budget</td><td style="padding:6px 8px;">${esc(lead.budget)}</td></tr>` : ''}
              ${lead.remarks ? `<tr><td style="padding:6px 8px;font-weight:bold;vertical-align:top;">Anmerkungen</td><td style="padding:6px 8px;">${esc(lead.remarks)}</td></tr>` : ''}
            </table>
          </div>

          <p>In der Zwischenzeit haben wir etwas Besonderes fuer euch vorbereitet — unseren <strong>Hochzeitsguide</strong> mit Tipps, Inspiration und allem, was ihr fuer eure Planung braucht:</p>
          <p style="text-align:center;margin:24px 0;">
            <a href="https://ww-lead-formular-production.up.railway.app/hochzeitsguide.html" style="display:inline-block;padding:14px 32px;background:#d4c4a8;color:#393e3f;text-decoration:none;letter-spacing:2px;font-size:14px;text-transform:uppercase;font-weight:bold;">Euer Hochzeitsguide</a>
          </p>
          <p>Schaut euch auch gerne auf unserer Website um oder folgt uns auf Instagram:</p>
          <p style="text-align:center;margin:24px 0;">
            <a href="https://www.walkingweddings.com" style="display:inline-block;padding:12px 28px;background:#393e3f;color:#fff;text-decoration:none;letter-spacing:2px;font-size:13px;text-transform:uppercase;">Unsere Website</a>
          </p>
          <p style="text-align:center;">
            <a href="https://www.instagram.com/walkingweddings" style="color:#747c7d;text-decoration:none;">@walkingweddings auf Instagram</a>
          </p>
          <p>Wir freuen uns auf eure Hochzeit!</p>
          <div style="margin-top:30px;padding-top:20px;border-top:1px solid #d4c4a8;">
            <p style="margin:0;font-weight:bold;">Walking Weddings OG</p>
            <p style="margin:4px 0;color:#747c7d;">Kiran: 0660 4822420</p>
            <p style="margin:4px 0;color:#747c7d;">Ian: 0660 6357799</p>
            <p style="margin:4px 0;color:#747c7d;">contact@walkingweddings.com</p>
            <p style="margin:4px 0;color:#747c7d;">www.walkingweddings.com</p>
          </div>
        </div>
      </div>
    `
  });
}

// Dashboard HTML
function renderDashboard() {
  const stars = n => '★'.repeat(n) + '☆'.repeat(5 - n);
  const rows = leads.slice().reverse().map((l, i) => {
    const dateStr = l.noDateYet ? 'Kein Datum' : (l.weddingDates?.join(', ') || '-');
    const locStr = l.noLocationYet ? 'Keine Location' : (l.locations?.join(', ') || '-');
    const emailStatus = l.emailSent ? '&#10003;' : '&#10007;';
    return `
      <tr>
        <td>${esc(l.name)}</td>
        <td>${esc(l.phone)}</td>
        <td>${esc(l.email)}</td>
        <td>${esc(dateStr)}</td>
        <td>${esc(locStr)}</td>
        <td>${esc(l.interests?.join(', ') || '-')}</td>
        <td>${esc(l.hours || '-')}</td>
        <td>${esc(l.extras?.join(', ') || '-')}</td>
        <td>${esc(l.budget || '-')}</td>
        <td>${esc(l.remarks || '-')}</td>
        <td style="color:#d4a017;">${stars(l.rating || 0)}</td>
        <td>${esc(l.notes || '-')}</td>
        <td style="text-align:center;">${emailStatus}</td>
        <td style="font-size:11px;color:#747c7d;">${l.timestamp ? new Date(l.timestamp).toLocaleString('de-AT') : '-'}</td>
      </tr>`;
  }).join('');

  const total = leads.length;
  const avgRating = total ? (leads.reduce((s, l) => s + (l.rating || 0), 0) / total).toFixed(1) : '0';
  const hot = leads.filter(l => (l.rating || 0) >= 4).length;

  return `<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>WW Dashboard - HochzeitStil Leads</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f0e8; color: #393e3f; }
  .header { background: #393e3f; color: #fff; padding: 20px 30px; display: flex; justify-content: space-between; align-items: center; }
  .header h1 { font-family: Georgia, serif; font-weight: 400; letter-spacing: 4px; font-size: 20px; }
  .header .sub { color: #cbd4d4; font-size: 11px; letter-spacing: 2px; }
  .stats { display: flex; gap: 20px; padding: 20px 30px; }
  .stat-card { background: #fff; border: 1px solid #d4c4a8; padding: 16px 24px; flex: 1; text-align: center; }
  .stat-card .num { font-size: 32px; font-weight: 700; color: #393e3f; }
  .stat-card .label { font-size: 12px; color: #747c7d; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }
  .table-wrap { padding: 0 30px 30px; overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; background: #fff; font-size: 13px; }
  th { background: #393e3f; color: #fff; padding: 10px 8px; text-align: left; font-weight: 600; font-size: 11px; text-transform: uppercase; letter-spacing: 1px; white-space: nowrap; }
  td { padding: 10px 8px; border-bottom: 1px solid #e8e0d0; vertical-align: top; }
  tr:hover td { background: #faf7f2; }
  .actions { padding: 10px 30px; display: flex; gap: 10px; }
  .btn { padding: 8px 20px; background: #393e3f; color: #fff; border: none; cursor: pointer; font-size: 12px; letter-spacing: 1px; text-transform: uppercase; text-decoration: none; }
  .btn:hover { background: #555; }
  .empty { text-align: center; padding: 60px; color: #747c7d; }
</style>
</head>
<body>
  <div class="header">
    <div>
      <h1>WALKING WEDDINGS</h1>
      <span class="sub">HOCHZEITSTIL MESSE - LEAD DASHBOARD</span>
    </div>
    <div style="text-align:right;">
      <span class="sub">Stand #42 &bull; Arena Nova</span>
    </div>
  </div>
  <div class="stats">
    <div class="stat-card"><div class="num">${total}</div><div class="label">Leads gesamt</div></div>
    <div class="stat-card"><div class="num">${avgRating}</div><div class="label">&#216; Bewertung</div></div>
    <div class="stat-card"><div class="num">${hot}</div><div class="label">Hot Leads (4-5&#9733;)</div></div>
  </div>
  <div class="actions">
    <a class="btn" href="/api/leads/xlsx">Excel Export</a>
    <a class="btn" href="/api/leads/csv">CSV Export</a>
    <button class="btn" onclick="location.reload()">Aktualisieren</button>
  </div>
  <div class="table-wrap">
    ${total === 0 ? '<div class="empty">Noch keine Leads erfasst.</div>' : `
    <table>
      <thead>
        <tr>
          <th>Name</th><th>Telefon</th><th>E-Mail</th><th>Datum</th><th>Location</th><th>Interesse</th>
          <th>Stunden</th><th>Extras</th><th>Budget</th><th>Anmerkungen</th>
          <th>Bewertung</th><th>Notizen</th><th>Mail</th><th>Zeitpunkt</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>`}
  </div>
</body>
</html>`;
}

function generateCSV() {
  const header = 'Name,Telefon,Email,Hochzeitsdatum,Location,Interesse,Stunden,Extras,Budget,Anmerkungen,Bewertung,Notizen,Email gesendet,Zeitpunkt\n';
  const csvEsc = (v) => '"' + String(v || '').replace(/"/g, '""') + '"';
  const rows = leads.map(l => {
    const dateStr = l.noDateYet ? 'Kein Datum' : (l.weddingDates?.join('; ') || '-');
    const locStr = l.noLocationYet ? 'Keine Location' : (l.locations?.join('; ') || '-');
    return [l.name, l.phone, l.email, dateStr, locStr, l.interests?.join('; '), l.hours, l.extras?.join('; '), l.budget, l.remarks, l.rating, l.notes, l.emailSent ? 'Ja' : 'Nein', l.timestamp || ''].map(csvEsc).join(',');
  }).join('\n');
  return header + rows;
}

// --- Minimal zero-dep XLSX writer ---
// XLSX is a ZIP of XML parts. We build a STORED (uncompressed) ZIP by hand,
// which lets us avoid adding any npm dependencies.

const CRC32_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC32_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function zipStored(files) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const f of files) {
    const nameBuf = Buffer.from(f.name, 'utf8');
    const dataBuf = Buffer.isBuffer(f.data) ? f.data : Buffer.from(f.data, 'utf8');
    const crc = crc32(dataBuf);
    const size = dataBuf.length;

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(size, 18);
    local.writeUInt32LE(size, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    localParts.push(local, nameBuf, dataBuf);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(size, 20);
    central.writeUInt32LE(size, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, nameBuf);

    offset += local.length + nameBuf.length + dataBuf.length;
  }
  const cdStart = offset;
  const cdSize = centralParts.reduce((s, p) => s + p.length, 0);

  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(cdSize, 12);
  eocd.writeUInt32LE(cdStart, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...localParts, ...centralParts, eocd]);
}

function xmlEsc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
    // strip control chars that are illegal in XML 1.0
    .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');
}

function colLetter(n) {
  // 1 -> A, 26 -> Z, 27 -> AA
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function sheetRow(rowIdx, cells) {
  const parts = cells.map((val, i) => {
    const ref = colLetter(i + 1) + rowIdx;
    return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(val)}</t></is></c>`;
  }).join('');
  return `<row r="${rowIdx}">${parts}</row>`;
}

function generateXLSX() {
  const headers = ['Name','Telefon','Email','Hochzeitsdatum','Location','Interesse','Stunden','Extras','Budget','Anmerkungen','Bewertung','Notizen','Email gesendet','Zeitpunkt'];
  const dataRows = leads.map(l => {
    const dateStr = l.noDateYet ? 'Kein Datum' : (l.weddingDates?.join('; ') || '');
    const locStr = l.noLocationYet ? 'Keine Location' : (l.locations?.join('; ') || '');
    return [
      l.name || '',
      l.phone || '',
      l.email || '',
      dateStr,
      locStr,
      l.interests?.join('; ') || '',
      l.hours || '',
      l.extras?.join('; ') || '',
      l.budget || '',
      l.remarks || '',
      l.rating != null ? String(l.rating) : '',
      l.notes || '',
      l.emailSent ? 'Ja' : 'Nein',
      l.timestamp ? new Date(l.timestamp).toLocaleString('de-AT') : ''
    ];
  });

  const allRows = [headers, ...dataRows];
  const sheetRowsXml = allRows.map((cells, i) => sheetRow(i + 1, cells)).join('');
  const lastCol = colLetter(headers.length);
  const dimension = `A1:${lastCol}${allRows.length}`;

  const sheetXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<dimension ref="${dimension}"/>` +
    `<sheetData>${sheetRowsXml}</sheetData>` +
    `</worksheet>`;

  const workbookXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<sheets><sheet name="Leads" sheetId="1" r:id="rId1"/></sheets>` +
    `</workbook>`;

  const workbookRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
    `</Relationships>`;

  const rootRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
    `</Relationships>`;

  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
    `</Types>`;

  return zipStored([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    { name: 'xl/workbook.xml', data: workbookXml },
    { name: 'xl/_rels/workbook.xml.rels', data: workbookRels },
    { name: 'xl/worksheets/sheet1.xml', data: sheetXml }
  ]);
}

const server = http.createServer(async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const url = new URL(req.url, `http://${req.headers.host}`);

  // Dashboard (PIN protected via query param)
  if (req.method === 'GET' && url.pathname === '/dashboard') {
    if (url.searchParams.get('pin') !== DASHBOARD_PIN) {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(`<!DOCTYPE html>
<html lang="de"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>WW Dashboard - Login</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #393e3f; color: #fff; display: flex; justify-content: center; align-items: center; min-height: 100vh; }
  .login { text-align: center; }
  .login h1 { font-family: Georgia, serif; font-weight: 400; letter-spacing: 5px; font-size: 22px; margin-bottom: 8px; }
  .login .sub { color: #cbd4d4; font-size: 11px; letter-spacing: 2px; display: block; margin-bottom: 30px; }
  .login input { padding: 12px 20px; font-size: 18px; letter-spacing: 8px; text-align: center; border: 1px solid #d4c4a8; background: transparent; color: #fff; width: 180px; }
  .login input::placeholder { color: #747c7d; letter-spacing: 4px; font-size: 14px; }
  .login button { display: block; margin: 16px auto 0; padding: 10px 30px; background: #d4c4a8; color: #393e3f; border: none; cursor: pointer; font-size: 12px; letter-spacing: 2px; text-transform: uppercase; }
  .err { color: #e8a0a0; font-size: 12px; margin-top: 12px; }
</style>
</head><body>
<div class="login">
  <h1>WALKING WEDDINGS</h1>
  <span class="sub">LEAD DASHBOARD</span>
  <form method="GET" action="/dashboard">
    <input type="password" name="pin" placeholder="PIN" maxlength="4" autofocus>
    <button type="submit">Anmelden</button>
  </form>
  ${url.searchParams.has('pin') ? '<p class="err">Falscher PIN</p>' : ''}
</div>
</body></html>`);
      return;
    }
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(renderDashboard());
    return;
  }

  // CSV Export
  if (req.method === 'GET' && url.pathname === '/api/leads/csv') {
    res.writeHead(200, {
      'Content-Type': 'text/csv; charset=utf-8',
      'Content-Disposition': 'attachment; filename="hochzeitstil_leads.csv"'
    });
    res.end('\ufeff' + generateCSV()); // BOM for Excel
    return;
  }

  // XLSX Export (echte Excel-Datei)
  if (req.method === 'GET' && url.pathname === '/api/leads/xlsx') {
    try {
      const buf = generateXLSX();
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="hochzeitstil_leads.xlsx"',
        'Content-Length': buf.length
      });
      res.end(buf);
    } catch (err) {
      console.error('XLSX export failed:', err);
      res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('XLSX export failed: ' + err.message);
    }
    return;
  }

  // API: Get leads (for dashboard AJAX if needed)
  if (req.method === 'GET' && url.pathname === '/api/leads') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(leads));
    return;
  }

  // API: Submit lead + send emails
  if (req.method === 'POST' && url.pathname === '/api/send-emails') {
    try {
      const lead = await parseBody(req);
      lead.timestamp = new Date().toISOString();
      lead.emailSent = false;

      // Store lead
      leads.push(lead);
      saveLeads();
      console.log(`Lead saved: ${lead.name} (total: ${leads.length})`);

      const results = { team: false, couple: false };

      try {
        await sendTeamEmail(lead);
        results.team = true;
        console.log(`Team email sent for: ${lead.name}`);
      } catch (err) {
        console.error('Team email failed:', err.message);
      }

      if (lead.email) {
        try {
          await sendCoupleEmail(lead);
          results.couple = true;
          console.log(`Couple email sent to: ${lead.email}`);
        } catch (err) {
          console.error('Couple email failed:', err.message);
        }
      }

      // Update email status
      if (results.team || results.couple) {
        lead.emailSent = true;
        saveLeads();
      }

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true, results }));
    } catch (err) {
      console.error('Email error:', err);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: err.message }));
    }
    return;
  }

  // API: Update lead rating/notes (from internal rating screen)
  if (req.method === 'POST' && url.pathname === '/api/update-lead') {
    try {
      const data = await parseBody(req);
      // Find last lead matching name+email and update rating/notes
      for (let i = leads.length - 1; i >= 0; i--) {
        if (leads[i].name === data.name && leads[i].email === data.email) {
          leads[i].rating = data.rating || 0;
          leads[i].notes = data.notes || '';
          saveLeads();
          break;
        }
      }
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true }));
    } catch (err) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false }));
    }
    return;
  }

  // Static files
  serveStatic(req, res);
});

server.listen(PORT, () => {
  console.log(`Walking Weddings Lead Server running on port ${PORT}`);
});
