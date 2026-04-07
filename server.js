const http = require('http');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

const PORT = process.env.PORT || 3000;

const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: 'contact@walkingweddings.com',
    pass: process.env.GMAIL_APP_PASSWORD
  }
});

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
  let filePath = req.url === '/' ? '/index.html' : req.url;
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

  await transporter.sendMail({
    from: '"Walking Weddings Lead" <contact@walkingweddings.com>',
    to: 'contact@walkingweddings.com',
    subject: `Neuer Messe-Lead: ${lead.name}`,
    html: `
      <h2 style="font-family:Georgia,serif;color:#393e3f;">Neuer Lead - HochzeitStil Messe</h2>
      <table style="border-collapse:collapse;width:100%;max-width:500px;font-family:Arial,sans-serif;">
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Name</td><td style="padding:8px;border:1px solid #d4c4a8;">${lead.name}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Telefon</td><td style="padding:8px;border:1px solid #d4c4a8;">${lead.phone}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">E-Mail</td><td style="padding:8px;border:1px solid #d4c4a8;">${lead.email}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Hochzeitsdatum</td><td style="padding:8px;border:1px solid #d4c4a8;">${dateStr}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Interesse</td><td style="padding:8px;border:1px solid #d4c4a8;">${lead.interests?.join(', ') || '-'}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Stunden</td><td style="padding:8px;border:1px solid #d4c4a8;">${lead.hours || '-'}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Zusatzprodukte</td><td style="padding:8px;border:1px solid #d4c4a8;">${lead.extras?.join(', ') || '-'}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Budget</td><td style="padding:8px;border:1px solid #d4c4a8;">${lead.budget || '-'}</td></tr>
        <tr><td style="padding:8px;border:1px solid #d4c4a8;font-weight:bold;background:#f5f0e8;">Anmerkungen</td><td style="padding:8px;border:1px solid #d4c4a8;">${lead.remarks || '-'}</td></tr>
      </table>
    `
  });
}

async function sendCoupleEmail(lead) {
  await transporter.sendMail({
    from: '"Walking Weddings" <contact@walkingweddings.com>',
    replyTo: 'contact@walkingweddings.com',
    to: lead.email,
    subject: 'Danke fuer euer Interesse - Walking Weddings',
    html: `
      <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#393e3f;">
        <div style="text-align:center;padding:30px 0;background:#393e3f;">
          <h1 style="color:#fff;font-family:Georgia,serif;font-weight:400;letter-spacing:5px;font-size:24px;margin:0;">WALKING WEDDINGS</h1>
          <p style="color:#cbd4d4;font-size:11px;letter-spacing:3px;margin:8px 0 0;">FILM AND PHOTOGRAPHY</p>
        </div>
        <div style="padding:30px 20px;">
          <p>Liebe/r ${lead.name},</p>
          <p>vielen Dank fuer euer Interesse an Walking Weddings auf der HochzeitStil Messe!</p>
          <p>Wir freuen uns sehr, dass wir euch kennenlernen durften. Unser Team wird sich in Kuerze bei euch melden, um ein unverbindliches Kennenlerngespraech zu vereinbaren.</p>
          <p>In der Zwischenzeit koennt ihr euch gerne auf unserer Website umsehen oder uns auf Instagram folgen:</p>
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

const server = http.createServer(async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  if (req.method === 'POST' && req.url === '/api/send-emails') {
    try {
      const lead = await parseBody(req);
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

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true, results }));
    } catch (err) {
      console.error('Email error:', err);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: err.message }));
    }
    return;
  }

  // Static files
  serveStatic(req, res);
});

server.listen(PORT, () => {
  console.log(`Walking Weddings Lead Server running on port ${PORT}`);
});
