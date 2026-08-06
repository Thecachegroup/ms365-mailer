'use strict';

// ms365-mailer — MCP HTTP endpoint for Vercel
// Credentials via Vercel environment variables — never hardcode them.

const https = require('https');
const querystring = require('querystring');

const TENANT_ID     = process.env.GRAPH_TENANT_ID;
const CLIENT_ID     = process.env.GRAPH_CLIENT_ID;
const CLIENT_SECRET = process.env.GRAPH_CLIENT_SECRET;
const SENDER_EMAIL  = process.env.SENDER_EMAIL  || 'andrew.hurnard@thecachegroup.com.au';
const SENDER_NAME   = process.env.SENDER_NAME   || 'Andrew Hurnard';
const SENDER_TITLE  = process.env.SENDER_TITLE  || 'Director';
const SENDER_COMPANY= process.env.SENDER_COMPANY|| 'The Cache Group';
const SENDER_PHONE  = process.env.SENDER_PHONE  || '0417 037 451';
const SIGN_OFF      = process.env.SIGN_OFF      || 'Regards';
const MCP_SECRET    = process.env.MCP_SHARED_SECRET || '';

// Graph refuses fileAttachment payloads above ~3 MB on a simple sendMail.
const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;

const LOGO_B64 = 'iVBORw0KGgoAAAANSUhEUgAAATgAAABGCAYAAABYIIhvAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAAZdEVYdFNvZnR3YXJlAEFkb2JlIEltYWdlUmVhZHlxyWU8AAADImlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNS4wLWMwNjAgNjEuMTM0Nzc3LCAyMDEwLzAyLzEyLTE3OjMyOjAwICAgICAgICAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdFJlZj0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlUmVmIyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ1M1IE1hY2ludG9zaCIgeG1wTU06SW5zdGFuY2VJRD0ieG1wLmlpZDo0MzU3MTc5OEREQkUxMUUzQkY4NkYzREU2RTJFRDcwOCIgeG1wTU06RG9jdW1lbnRJRD0ieG1wLmRpZDo0MzU3MTc5OUREQkUxMUUzQkY4NkYzREU2RTJFRDcwOCI+IDx4bXBNTTpEZXJpdmVkRnJvbSBzdFJlZjppbnN0YW5jZUlEPSJ4bXAuaWlkOjQzNTcxNzk2RERCRTExRTNCRjg2RjNERTZFMkVENzA4IiBzdFJlZjpkb2N1bWVudElEPSJ4bXAuZGlkOjQzNTcxNzk3RERCRTExRTNCRjg2RjNERTZFMkVENzA4Ii8+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwv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';

// ── Signature ─────────────────────────────────────────────────────────────────

function buildHtmlBody(bodyText) {
  const bodyHtml = bodyText
    .split('\n')
    .map(l => `<p style="margin:0 0 4px 0;">${l.trim() || '&nbsp;'}</p>`)
    .join('');

  const sig = `
<br>
<p style="margin:0 0 4px 0;font-family:Arial,sans-serif;font-size:13px;">${SIGN_OFF}</p>
<br>
<span style="font-family:Arial,sans-serif;font-size:13px;line-height:1.6;">
  <strong style="color:#1a1a1a;">${SENDER_NAME}</strong><br>
  <strong style="color:#7B1E3C;">${SENDER_TITLE}</strong><br>
  <strong style="color:#7B1E3C;">${SENDER_COMPANY}</strong><br>
  <strong style="color:#7B1E3C;">${SENDER_PHONE}</strong>
</span>
<br>
<img src="data:image/png;base64,${LOGO_B64}" alt="${SENDER_COMPANY}" style="max-width:260px;height:auto;display:block;margin-top:10px;" />`;

  return `<html><body>${bodyHtml}${sig}</body></html>`;
}

// ── Graph helpers ─────────────────────────────────────────────────────────────

function httpsPost(hostname, path, headers, body) {
  return new Promise((resolve, reject) => {
    const buf = Buffer.from(typeof body === 'string' ? body : JSON.stringify(body));
    const req = https.request(
      { hostname, path, method: 'POST', headers: { ...headers, 'Content-Length': buf.length } },
      (res) => { let d = ''; res.on('data', c => d += c); res.on('end', () => resolve({ status: res.statusCode, body: d })); }
    );
    req.on('error', reject);
    req.write(buf);
    req.end();
  });
}

// GET returning raw bytes, following redirects.
// Graph answers /content with a 302 to a pre-signed storage URL — the auth
// header is deliberately dropped on the hop so credentials never leave Graph.
function httpsGetBuffer(url, headers, depth) {
  depth = depth || 0;
  return new Promise((resolve, reject) => {
    if (depth > 5) return reject(new Error('Too many redirects fetching attachment'));
    let u;
    try { u = new URL(url); } catch (e) { return reject(new Error(`Bad URL: ${url}`)); }
    const req = https.request(
      { hostname: u.hostname, path: u.pathname + u.search, method: 'GET', headers },
      (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          res.resume();
          return resolve(httpsGetBuffer(res.headers.location, {}, depth + 1));
        }
        if (res.statusCode !== 200) {
          let d = '';
          res.on('data', c => d += c);
          return res.on('end', () => reject(new Error(`Graph GET ${res.statusCode}: ${String(d).slice(0, 300)}`)));
        }
        const chunks = [];
        res.on('data', c => chunks.push(c));
        res.on('end', () => resolve(Buffer.concat(chunks)));
      }
    );
    req.on('error', reject);
    req.end();
  });
}

async function getToken() {
  const body = querystring.stringify({
    grant_type: 'client_credentials', client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET, scope: 'https://graph.microsoft.com/.default'
  });
  const res = await httpsPost('login.microsoftonline.com', `/${TENANT_ID}/oauth2/v2.0/token`,
    { 'Content-Type': 'application/x-www-form-urlencoded' }, body);
  const parsed = JSON.parse(res.body);
  if (!parsed.access_token) throw new Error(parsed.error_description || 'Token failed');
  return parsed.access_token;
}

async function graphSendMail(token, message) {
  const res = await httpsPost('graph.microsoft.com',
    `/v1.0/users/${encodeURIComponent(SENDER_EMAIL)}/sendMail`,
    { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    JSON.stringify({ message, saveToSentItems: true }));
  if (res.status !== 202) throw new Error(`Graph API ${res.status}: ${res.body}`);
}

// ── OneDrive attachments ──────────────────────────────────────────────────────

const MIME_BY_EXT = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  doc:  'application/msword',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xls:  'application/vnd.ms-excel',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  pdf:  'application/pdf',
  csv:  'text/csv',
  txt:  'text/plain',
  png:  'image/png',
  jpg:  'image/jpeg',
  jpeg: 'image/jpeg',
  zip:  'application/zip',
  msg:  'application/vnd.ms-outlook'
};

function mimeFor(filename) {
  const ext = String(filename).split('.').pop().toLowerCase();
  return MIME_BY_EXT[ext] || 'application/octet-stream';
}

// Encode a OneDrive-relative path for Graph path addressing.
// "CONTRACTOR AGREEMENTS/Devinia Liddelow/Brief.docx" -> percent-encoded segments
function encodeDrivePath(p) {
  return String(p)
    .replace(/\\/g, '/')      // tolerate Windows-style separators
    .replace(/^\/+/, '')      // drop any leading slash
    .split('/')
    .filter(Boolean)
    .map(encodeURIComponent)
    .join('/');
}

async function fetchOneDriveAttachment(token, path) {
  const encoded = encodeDrivePath(path);
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(SENDER_EMAIL)}`
            + `/drive/root:/${encoded}:/content`;
  const buf = await httpsGetBuffer(url, { 'Authorization': `Bearer ${token}` });

  if (buf.length > MAX_ATTACHMENT_BYTES) {
    throw new Error(
      `Attachment "${path}" is ${(buf.length / 1024 / 1024).toFixed(1)} MB — `
      + `Graph refuses attachments over 3 MB on a direct send.`
    );
  }

  const name = String(path).replace(/\\/g, '/').split('/').filter(Boolean).pop();
  return {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name,
    contentType: mimeFor(name),
    contentBytes: buf.toString('base64'),
    _bytes: buf.length      // stripped before sending; used for the result message
  };
}

// ── Tools ─────────────────────────────────────────────────────────────────────

const TOOLS = [{
  name: 'send_email',
  description: `Send an email from ${SENDER_EMAIL}. Shows a preview unless confirm is true. Appends TCG signature automatically. `
    + `To attach a file, prefer attach_from_onedrive — pass the file's path relative to the OneDrive root `
    + `(e.g. "CONTRACTOR AGREEMENTS/Devinia Liddelow/Consultancy Brief Devinia Liddelow 19022027.docx") and the server `
    + `fetches it from OneDrive itself. Use the attachments parameter only for files that do not exist in OneDrive, `
    + `and only when small — it requires the whole file base64-encoded inline, which is unreliable above a few KB.`,
  inputSchema: {
    type: 'object',
    required: ['to', 'subject', 'body'],
    properties: {
      to:      { type: 'string', description: 'Recipient email. Comma-separate for multiple.' },
      subject: { type: 'string', description: 'Subject line' },
      body:    { type: 'string', description: 'Plain-text body (signature appended automatically)' },
      cc:      { type: 'string', description: 'CC address (optional)' },
      confirm: { type: 'boolean', description: 'false = preview only (default), true = send', default: false },
      attach_from_onedrive: {
        type: 'array',
        description: 'PREFERRED. Paths of files to attach, relative to the OneDrive root of '
          + SENDER_EMAIL + '. Forward or back slashes both work. The server fetches each file from '
          + 'OneDrive via Graph — no file content passes through the conversation, so there is no '
          + 'size or transcription limit beyond Graph\'s 3 MB attachment ceiling.',
        items: { type: 'string' }
      },
      attachments: {
        type: 'array',
        description: 'Fallback for files not in OneDrive. Requires the full file base64-encoded inline — '
          + 'only practical for small files. Prefer attach_from_onedrive.',
        items: {
          type: 'object',
          required: ['name', 'contentType', 'contentBytes'],
          properties: {
            name:         { type: 'string' },
            contentType:  { type: 'string' },
            contentBytes: { type: 'string', description: 'Base64-encoded file content' }
          }
        }
      }
    }
  }
}];

async function callSendEmail(args) {
  const { to, subject, body, cc, confirm, attachments, attach_from_onedrive } = args;

  const drivePaths = Array.isArray(attach_from_onedrive) ? attach_from_onedrive.filter(Boolean) : [];
  const inlineAtts = Array.isArray(attachments) ? attachments : [];

  const attachmentSummary = [
    ...drivePaths.map(p => `  - ${p}  (from OneDrive)`),
    ...inlineAtts.map(a => `  - ${a.name}  (inline)`)
  ].join('\n');

  const preview = `FROM:    ${SENDER_EMAIL}\nTO:      ${to}\n`
    + (cc ? `CC:      ${cc}\n` : '')
    + `SUBJECT: ${subject}\n`
    + (attachmentSummary ? `ATTACHED:\n${attachmentSummary}\n` : '')
    + `\n${body}\n\n[Signature appended]`;

  if (!confirm) {
    return { preview: true, text: `PREVIEW (not sent):\n\n${preview}\n\nCall again with confirm: true to send.` };
  }

  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Missing Graph credentials — check Vercel environment variables.');
  }

  const token = await getToken();

  const toRecipients = to.split(',').map(a => ({ emailAddress: { address: a.trim() } }));
  const message = {
    subject,
    body: { contentType: 'HTML', content: buildHtmlBody(body) },
    toRecipients,
    from: { emailAddress: { address: SENDER_EMAIL, name: SENDER_NAME } }
  };
  if (cc) message.ccRecipients = cc.split(',').map(a => ({ emailAddress: { address: a.trim() } }));

  const built = [];

  // Fetched server-side from OneDrive — the preferred path.
  for (const p of drivePaths) {
    built.push(await fetchOneDriveAttachment(token, p));
  }

  // Legacy inline base64 — kept for files that are not in OneDrive.
  for (const a of inlineAtts) {
    built.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: a.name, contentType: a.contentType, contentBytes: a.contentBytes,
      _bytes: Buffer.from(a.contentBytes, 'base64').length
    });
  }

  let sentNote = '';
  if (built.length > 0) {
    const total = built.reduce((n, a) => n + (a._bytes || 0), 0);
    if (total > MAX_ATTACHMENT_BYTES) {
      throw new Error(`Attachments total ${(total / 1024 / 1024).toFixed(1)} MB — over Graph's 3 MB limit for a direct send.`);
    }
    message.attachments = built.map(({ _bytes, ...rest }) => rest);
    sentNote = ' with ' + built.map(a => `${a.name} (${a._bytes.toLocaleString()} bytes)`).join(', ');
  }

  await graphSendMail(token, message);
  return { preview: false, text: `✓ Email sent to ${to}${sentNote}` };
}

// ── MCP router ────────────────────────────────────────────────────────────────

async function handleMcp(rpc) {
  const { method, params, id } = rpc;
  if (method === 'initialize') return { jsonrpc: '2.0', id, result: { protocolVersion: '2024-11-05', capabilities: { tools: {} }, serverInfo: { name: 'ms365-mailer', version: '1.1.0' } } };
  if (method === 'notifications/initialized') return null;
  if (method === 'ping') return { jsonrpc: '2.0', id, result: {} };
  if (method === 'tools/list') return { jsonrpc: '2.0', id, result: { tools: TOOLS } };
  if (method === 'tools/call') {
    const { name, arguments: args } = params;
    if (name === 'send_email') {
      try {
        const result = await callSendEmail(args);
        return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: result.text }] } };
      } catch (err) {
        return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: `✗ ${err.message}` }], isError: true } };
      }
    }
    return { jsonrpc: '2.0', id, error: { code: -32601, message: `Unknown tool: ${name}` } };
  }
  return { jsonrpc: '2.0', id, error: { code: -32601, message: `Unknown method: ${method}` } };
}

// ── Vercel handler ────────────────────────────────────────────────────────────

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method === 'GET') return res.json({ status: 'ok', server: 'ms365-mailer', sender: SENDER_EMAIL });
  // Secret check — reject requests not targeting /mcp/{SECRET}
  if (MCP_SECRET) {
    const reqPath = (req.url || '').split('?')[0];
    if (reqPath !== '/mcp/' + MCP_SECRET && reqPath !== '/' + MCP_SECRET) {
      return res.status(404).json({ error: 'Not found' });
    }
  }
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });
  try {
    const rpc = req.body;
    if (Array.isArray(rpc)) {
      const results = (await Promise.all(rpc.map(handleMcp))).filter(Boolean);
      return res.json(results);
    }
    const result = await handleMcp(rpc);
    if (result === null) return res.status(204).end();
    return res.json(result);
  } catch (err) {
    return res.status(500).json({ jsonrpc: '2.0', error: { code: -32603, message: err.message } });
  }
};
