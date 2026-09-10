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

// ── Where attachments are read from ──────────────────────────────────────────
// Attachments used to come from whichever mailbox the deployment sends as:
// every profile's `drive` resolved to SENDER_EMAIL, so Matt's deployment looked
// on Matt's OneDrive and Nuria's on Nuria's. The files are on neither. That is
// why "attach the document" worked for one person and failed for the other two.
//
// ATTACH_DRIVE_ID points every deployment at ONE shared library — the TCG
// SharePoint team site — addressed by driveId, so the sender identity and the
// file location are finally separate things. Leave it unset and behaviour is
// exactly as before: each profile reads its own `drive` owner's OneDrive.
//
// ATTACH_ROOT is an optional folder inside that library that every relative
// path is resolved under, so callers keep passing "Consultancy Brief X.docx"
// rather than a library-specific prefix.
const ATTACH_DRIVE_ID = process.env.ATTACH_DRIVE_ID || '';
const ATTACH_ROOT     = process.env.ATTACH_ROOT     || '';

// ── Optional extra senders ────────────────────────────────────────────────────
// Both default OFF. A deployment sends as its own SENDER_EMAIL and nothing else
// unless explicitly opted in here.
//
// Payroll used to be unconditional, so EVERY deployment could ask to send as
// payroll@ and the only thing refusing it was that mailbox being absent from
// the app's Exchange access policy. That left a tenant-side policy in charge of
// a code-side decision: one policy edit, or one new app registration created
// without one, and a recruiter's connector can send payslips. Fail closed here
// instead, and let the access policy be the second gate rather than the only
// one.
//
// Defaulting OFF means a deployment that needs payroll must say so. Set
// ENABLE_PAYROLL_SENDER=true on that project BEFORE deploying this change, or
// its payroll sends stop working.
const ENABLE_CAREERS  = String(process.env.ENABLE_CAREERS_SENDER || '').trim().toLowerCase() === 'true';
const ENABLE_PAYROLL  = String(process.env.ENABLE_PAYROLL_SENDER || '').trim().toLowerCase() === 'true';
const CAREERS_EMAIL   = 'careers@thecachegroup.com.au';
const PAYROLL_ADDRESS = 'payroll@thecachegroup.com.au';
const PAYROLL_MAILBOX = 'payrollmb@thecachegroup.com.au';

// ── Sender profiles ─────────────────────────────────────────────────────────
// Which mailboxes this server may send as, and the signature each one carries.
//
// The app registration behind this server holds APPLICATION Mail.Send, which
// Graph describes in the portal as "send mail as any user". It is not scoped:
// Graph will send as anyone in the tenant if asked. This allowlist is the only
// thing standing between a typo and an email that appears to come from someone
// else. An address that is not a key here is refused before a token is even
// requested. Adding a key genuinely widens what this server can impersonate —
// treat it as a security decision, not configuration.
//
// `mailbox` is what goes in the Graph URL and must resolve to a real mailbox.
// `address` is what recipients see in the From line, and must be an address
// that mailbox actually owns, or Graph rewrites it to the primary and the
// whole exercise is pointless. For payroll the two differ on purpose: the UPN
// is still payrollmb@ (changing the primary SMTP does not move a UPN) while
// the primary — and so the From line — is now payroll@.
//
// A profile field left empty is omitted from the signature rather than
// rendered blank. Payroll carries no personal name, title or mobile.

const SENDERS = {
  [SENDER_EMAIL.toLowerCase()]: {
    mailbox: SENDER_EMAIL,
    address: SENDER_EMAIL,
    name:    SENDER_NAME,
    title:   SENDER_TITLE,
    company: SENDER_COMPANY,
    phone:   SENDER_PHONE,
    signOff: SIGN_OFF,
    drive:   SENDER_EMAIL,
    // IDENTITY, not address. The impersonation guard below refuses a body
    // carrying another sender's name or mobile. careers@ deliberately carries
    // THIS person's name and mobile, so without a shared identity key each
    // profile becomes the "other" of the other, and an ordinary email that
    // mentions the sender's own mobile is refused as impersonation of himself.
    // Profiles that put the same person in the signature share one identity.
    identity: SENDER_EMAIL.toLowerCase()
  }
};

// ── careers@ ─────────────────────────────────────────────────────────────────
// The shared candidate inbox. Signed by whoever's deployment this is, so a
// candidate can see who wrote to them while their reply still lands in the
// shared inbox rather than one person's. Same identity as the owner above.
//
// The guard is not paranoia: a deployment whose SENDER_EMAIL is careers@ would
// otherwise have its own profile silently overwritten by this one.
if (ENABLE_CAREERS && SENDER_EMAIL.toLowerCase() !== CAREERS_EMAIL) {
  SENDERS[CAREERS_EMAIL] = {
    mailbox: CAREERS_EMAIL,
    address: CAREERS_EMAIL,
    name:    SENDER_NAME,
    title:   SENDER_TITLE,
    company: SENDER_COMPANY,
    phone:   SENDER_PHONE,
    signOff: SIGN_OFF,
    // Attachments come from the deployment owner's OneDrive, not the shared
    // mailbox — careers@ has no drive of its own.
    drive:   SENDER_EMAIL,
    identity: SENDER_EMAIL.toLowerCase()
  };
}

// ── payroll@ ─────────────────────────────────────────────────────────────────
if (ENABLE_PAYROLL) {
  SENDERS[PAYROLL_ADDRESS] = {
    mailbox: PAYROLL_MAILBOX,
    address: PAYROLL_ADDRESS,
    name:    'The Payroll Team',
    // Other forms of the same name a body might end with. Used ONLY by the
    // trailing sign-off stripper, never by the impersonation guard — they
    // identify THIS sender, so they must never cause a refusal. Without them,
    // a body ending "Regards / Payroll" (which payroll-copilot may well
    // compose) survives the stripper and the recipient sees the sign-off twice.
    altNames: ['Payroll', 'Payroll Team'],
    title:   '',
    company: SENDER_COMPANY,
    phone:   '',
    signOff: SIGN_OFF,
    // Attachments still come from Andrew's OneDrive. The payroll mailbox has
    // no drive of its own, and every path callers pass — AI Working Folder,
    // CONTRACTOR AGREEMENTS — lives on his. Pointing this at the sender would
    // break every attachment on a payroll send.
    drive:   SENDER_EMAIL,
    // Its own identity: payroll@ carries no personal name or mobile, and a
    // body signed by a person is genuinely wrong on a payroll send.
    identity: 'payroll'
  };

  // Same mailbox reachable by its other address, so a caller who says
  // payrollmb@ gets the same profile rather than a refusal.
  SENDERS[PAYROLL_MAILBOX] = SENDERS[PAYROLL_ADDRESS];
}

// Set SENDER_EMAIL to an address that is also a built-in key and the built-in
// wins silently: Andrew's profile disappears, every default send goes out as
// Payroll, and `drive` points at a mailbox with no OneDrive so every
// attachment 404s — without a word of complaint. Refuse to start instead.
//
// Checked against the literal key list, and AFTER the alias assignment. Both
// details matter and the first attempt at this guard got both wrong: comparing
// SENDERS[key].address to SENDER_EMAIL misses an EXACT collision, because the
// surviving profile's address is the very value that collided; and a guard
// sitting above the alias line cannot see the payrollmb@ collision at all,
// because that key does not exist until the line below has run.
const BUILTIN_SENDER_KEYS = [
  'payroll@thecachegroup.com.au',
  'payrollmb@thecachegroup.com.au'
];
if (BUILTIN_SENDER_KEYS.includes(SENDER_EMAIL.toLowerCase())) {
  throw new Error(
    `SENDER_EMAIL "${SENDER_EMAIL}" collides with a built-in sender profile. `
    + 'Change SENDER_EMAIL, or remove the conflicting profile from SENDERS.'
  );
}

function allowedSenders() {
  const seen = [];
  for (const p of Object.values(SENDERS)) {
    if (!seen.includes(p.address)) seen.push(p.address);
  }
  return seen;
}

function resolveSender(from) {
  if (from === undefined || from === null) {
    return SENDERS[SENDER_EMAIL.toLowerCase()];
  }
  // Strict on type. String([]) is '' and so is String(['']) and String([null]),
  // so a bare String(from).trim() === '' check quietly resolves an
  // array-shaped `from` to the DEFAULT sender: a caller that malformed its
  // from while meaning payroll would send as Andrew — full name, title and
  // mobile — with no error at all. Refuse the shape rather than guess at the
  // intent behind it.
  if (typeof from !== 'string') {
    throw new Error('`from` must be a string email address.');
  }
  const key = from.trim().toLowerCase();
  if (key === '') return SENDERS[SENDER_EMAIL.toLowerCase()];
  // hasOwnProperty, not a bare lookup. SENDERS is a plain object literal, so
  // it inherits Object.prototype: a bare SENDERS[key] returns a truthy
  // FUNCTION for "tostring", "constructor", "valueof" and friends, which
  // would sail through a `if (!profile)` check and then produce a preview
  // reading "FROM: undefined" instead of a refusal.
  const profile = Object.prototype.hasOwnProperty.call(SENDERS, key)
    ? SENDERS[key]
    : null;
  if (!profile) {
    throw new Error(
      `Refusing to send as "${from}". Accepted values: `
      + Object.keys(SENDERS).join(', ')
      + '. Those resolve to the From lines: ' + allowedSenders().join(', ') + '.'
    );
  }
  return profile;
}

// Graph refuses fileAttachment payloads above ~3 MB on a simple sendMail.
const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;

// TCG logo, 312 x 70 px. Sent as an inline cid: attachment, never as a data:
// URI — Outlook and Gmail both strip data: image sources, which is why the
// logo silently vanished from every send before v1.2.0.
const LOGO_CID    = 'tcglogo';
const LOGO_NAME   = 'tcg-logo.png';
const LOGO_WIDTH  = 260;   // displayed width
const LOGO_HEIGHT = 58;    // 312:70 scaled to 260 — both attributes set so Outlook does not stretch it
const LOGO_B64 = 'iVBORw0KGgoAAAANSUhEUgAAATgAAABGCAYAAABYIIhvAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAAZdEVYdFNvZnR3YXJlAEFkb2JlIEltYWdlUmVhZHlxyWU8AAADImlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNS4wLWMwNjAgNjEuMTM0Nzc3LCAyMDEwLzAyLzEyLTE3OjMyOjAwICAgICAgICAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdFJlZj0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlUmVmIyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ1M1IE1hY2ludG9zaCIgeG1wTU06SW5zdGFuY2VJRD0ieG1wLmlpZDo0MzU3MTc5OEREQkUxMUUzQkY4NkYzREU2RTJFRDcwOCIgeG1wTU06RG9jdW1lbnRJRD0ieG1wLmRpZDo0MzU3MTc5OUREQkUxMUUzQkY4NkYzREU2RTJFRDcwOCI+IDx4bXBNTTpEZXJpdmVkRnJvbSBzdFJlZjppbnN0YW5jZUlEPSJ4bXAuaWlkOjQzNTcxNzk2RERCRTExRTNCRjg2RjNERTZFMkVENzA4IiBzdFJlZjpkb2N1bWVudElEPSJ4bXAuZGlkOjQzNTcxNzk3RERCRTExRTNCRjg2RjNERTZFMkVENzA4Ii8+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwv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';

// ── Sign-off guard ────────────────────────────────────────────────────────────
// The signature block below already opens with SIGN_OFF ("Regards") and carries
// the sender's name. If a body also ends with a sign-off, the recipient sees it
// twice. Trim a trailing sign-off — and a trailing bare sender name — off the
// end of the body before anything is built from it. Trailing only: a "Thanks"
// in the middle of a sentence is left alone.

const SIGN_OFF_LINES = [
  'regards', 'kind regards', 'warm regards', 'best regards', 'kindest regards',
  'best', 'best wishes', 'all the best',
  'thanks', 'thank you', 'thanks again', 'many thanks', 'thanks so much',
  'cheers', 'sincerely', 'yours sincerely', 'yours faithfully'
];

function isSignOffLine(line) {
  const s = line.trim().toLowerCase().replace(/[,.!]+$/, '');
  return SIGN_OFF_LINES.includes(s);
}

// EVERY name this server can send under, not just the active one.
//
// The stripper has to recognise "Andrew Hurnard" at the foot of a body that is
// going out as Payroll — that is precisely the case it exists for. Binding it
// to the active profile's name silently disables it exactly when `from` is
// set, and because the trim loop breaks on the first line it does not
// recognise, an unmatched name also shields the "Regards" above it. The result
// is a payslip from payroll@ signed by Andrew, with a doubled sign-off.
function knownSenderNames() {
  const names = [SENDER_NAME];
  for (const p of Object.values(SENDERS)) {
    for (const n of [p.name, ...(Array.isArray(p.altNames) ? p.altNames : [])]) {
      if (n && !names.includes(n)) names.push(n);
    }
  }
  return names.filter(Boolean);
}

function isSenderNameLine(line, names) {
  const s = line.trim().toLowerCase().replace(/[,.!]+$/, '');
  if (!s) return false;
  const list = Array.isArray(names) ? names : [names];
  return list.some(n => {
    const full = String(n || '').trim().toLowerCase();
    if (!full) return false;
    const first = full.split(/\s+/)[0];
    // An article is not a name. Without this, "The Payroll Team" teaches the
    // stripper that a trailing line reading "The" is a signature, and it eats
    // the last line of the body.
    if (['the', 'a', 'an'].includes(first)) return s === full;
    return s === full || s === first;
  });
}

function stripTrailingSignOff(bodyText, names) {
  if (typeof bodyText !== 'string') return bodyText;

  const lines = bodyText.split('\n');
  let removed = 0;

  // Never chew through more than a sign-off plus a name.
  while (removed < 2) {
    let i = lines.length - 1;
    while (i >= 0 && lines[i].trim() === '') i--;
    if (i < 0) break;

    if (isSenderNameLine(lines[i], names) || isSignOffLine(lines[i])) {
      lines.length = i;
      removed++;
      continue;
    }
    break;
  }

  if (removed === 0) return bodyText;

  while (lines.length && lines[lines.length - 1].trim() === '') lines.pop();

  const out = lines.join('\n');
  // A body that was nothing but a sign-off is left as the caller wrote it.
  return out.trim() === '' ? bodyText : out;
}

// ── Signature ─────────────────────────────────────────────────────────────────
// Every line carries its own font-family, size and colour on a block-level
// element. Nothing relies on inheritance from a wrapping <span> or <body>:
// receiving clients rewrite the outer HTML when they quote a reply, and any
// styling that lived on a wrapper is lost at that point. That is what made the
// signature look different in every reply before v1.2.0.

const FONT  = 'font-family:Arial,Helvetica,sans-serif;font-size:13px;line-height:1.5;';
const INK   = '#1a1a1a';
const BRAND = '#932B46';   // TCG brand burgundy, RGB(147,43,70)

function buildHtmlBody(bodyText, profile) {
  // Blank lines are dropped rather than rendered as empty paragraphs: each
  // paragraph already carries its own bottom margin, so an extra <p>&nbsp;</p>
  // just doubles the gap. This is what produced the run of blank lines above
  // the sign-off in every send before v1.2.0.
  // Body text is interpolated into HTML, and the schema documents it as PLAIN
  // text. Without escaping, "<see attached>" silently vanishes in the client
  // and an anchor tag becomes a live link — inside an email that genuinely
  // originates from the address contractors are told to trust for payslips
  // and bank details. That needs no compromise of this server: payroll-copilot
  // composes bodies from content swept out of the payroll mailbox, which is
  // contractor-supplied.
  //
  // Deliberately only & < > — the three characters that can change structure.
  // Escaping quotes or anything else would start altering how ordinary text
  // renders; these three cannot. "Smith & Sons" becomes "Smith &amp; Sons" in
  // the source and still reads "Smith & Sons" on screen.
  const esc = s => String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  const bodyHtml = bodyText
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean)
    .map(l => `<p style="margin:0 0 10px 0;${FONT}color:${INK};">${esc(l)}</p>`)
    .join('');

  // An empty profile field is omitted entirely rather than rendered as an
  // empty <p>. Payroll has no personal name, title or mobile, and a run of
  // empty paragraphs shows up as a visible gap above the logo.
  //
  // Every detail line carries margin:0 and the spacing before the logo lives
  // on the logo's own paragraph. Putting the gap on the phone line — as this
  // did before — loses it whenever the phone is blank.
  // Escaped like the body. These fields are constants and env vars today, not
  // caller-reachable, so this is not closing a hole — it is stopping the two
  // halves of one function disagreeing about whether interpolated text is
  // escaped, which is how the next person introduces one.
  const line = (text, colour) =>
    (text && String(text).trim())
      ? `<p style="margin:0;${FONT}color:${colour};"><strong>${esc(text)}</strong></p>`
      : '';

  const sig =
      `<p style="margin:18px 0 12px 0;${FONT}color:${INK};">${esc(profile.signOff)}</p>`
    + line(profile.name,    INK)
    + line(profile.title,   BRAND)
    + line(profile.company, BRAND)
    + line(profile.phone,   BRAND)
    + `<p style="margin:12px 0 0 0;"><img src="cid:${LOGO_CID}" alt="${profile.company}"`
    + ` width="${LOGO_WIDTH}" height="${LOGO_HEIGHT}"`
    + ` style="width:${LOGO_WIDTH}px;height:${LOGO_HEIGHT}px;display:block;border:0;outline:none;text-decoration:none;" /></p>`;

  return `<html><body style="margin:0;padding:0;">${bodyHtml}${sig}</body></html>`;
}

function logoAttachment() {
  const bytes = Buffer.from(LOGO_B64, 'base64').length;
  return {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: LOGO_NAME,
    contentType: 'image/png',
    contentBytes: LOGO_B64,
    isInline: true,
    contentId: LOGO_CID,
    _bytes: bytes,
    _inline: true
  };
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

async function graphSendMail(token, message, mailbox) {
  if (!mailbox) throw new Error('graphSendMail called without a mailbox');
  const res = await httpsPost('graph.microsoft.com',
    `/v1.0/users/${encodeURIComponent(mailbox)}/sendMail`,
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
    .replace(/\\/g, '/')     // tolerate Windows-style separators
    .replace(/^\/+/, '')     // drop any leading slash
    .split('/')
    .filter(Boolean)
    .map(encodeURIComponent)
    .join('/');
}

// The drive every attachment is read from. A shared library wins over the
// sender's own OneDrive whenever ATTACH_DRIVE_ID is set.
function driveBase(profile) {
  if (ATTACH_DRIVE_ID) {
    return `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(ATTACH_DRIVE_ID)}`;
  }
  if (!profile || !profile.drive) {
    throw new Error('No attachment drive configured — set ATTACH_DRIVE_ID.');
  }
  return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(profile.drive)}/drive`;
}

function withAttachRoot(p) {
  const rel = String(p).replace(/\\/g, '/').replace(/^\/+/, '');
  if (!ATTACH_ROOT) return rel;
  return `${String(ATTACH_ROOT).replace(/\\/g, '/').replace(/^\/+|\/+$/g, '')}/${rel}`;
}

function isNotFound(err) {
  return /Graph GET 404/.test(String(err && err.message));
}

async function graphGetJson(url, token) {
  const buf = await httpsGetBuffer(url, { 'Authorization': `Bearer ${token}` });
  return JSON.parse(buf.toString('utf8'));
}

// Resolve one caller-supplied path to a real item BEFORE anything is sent.
//
// Two things were wrong before. The caller had to know the exact full path from
// the drive root, because the path was percent-encoded and handed straight to
// Graph — a near miss came back as a bare `itemNotFound` that did not even name
// what it had asked for. And nothing was resolved until the send, so the
// preview happily listed a file that did not exist and only the real send
// failed. Both are fixed here: an exact path is tried first, a plain filename
// falls back to a search of the library, and every caller resolves at preview
// time.
async function resolveOneDriveItem(token, rawPath, profile) {
  const base = driveBase(profile);
  const rel = encodeDrivePath(withAttachRoot(rawPath));
  const wanted = String(rawPath).replace(/\\/g, '/').split('/').filter(Boolean).pop() || '';

  if (rel) {
    try {
      const meta = await graphGetJson(`${base}/root:/${rel}`, token);
      if (meta && meta.id && meta.file) {
        return { base, id: meta.id, name: meta.name, size: meta.size || 0, how: 'path' };
      }
      if (meta && meta.id && meta.folder) {
        throw new Error(`"${rawPath}" is a folder, not a file.`);
      }
    } catch (err) {
      // Only a genuine miss falls through to search. An auth or throttling
      // failure must surface as itself rather than turning into "not found",
      // which is how a permissions problem gets diagnosed as a typo.
      if (!isNotFound(err)) throw err;
    }
  }

  if (!wanted) throw new Error('Empty attachment path.');

  const hits = await graphGetJson(
    `${base}/root/search(q='${encodeURIComponent(wanted.replace(/'/g, "''"))}')`
    + `?$select=id,name,size,file,parentReference&$top=50`, token);

  const files = (hits && Array.isArray(hits.value) ? hits.value : [])
    .filter(v => v && v.file && String(v.name).toLowerCase() === wanted.toLowerCase());

  if (files.length === 1) {
    const f = files[0];
    const folder = f.parentReference && f.parentReference.path
      ? String(f.parentReference.path).replace(/^\/drive(s)?\/[^/]+\/root:?/, '').replace(/^\/+/, '')
      : '';
    return { base, id: f.id, name: f.name, size: f.size || 0, how: 'search', folder };
  }

  if (files.length > 1) {
    const list = files.slice(0, 10).map(f => {
      const path = f.parentReference && f.parentReference.path ? f.parentReference.path : '';
      return `  - ${path}/${f.name}`;
    }).join('\n');
    throw new Error(
      `"${wanted}" matches ${files.length} files in the attachment library. `
      + `Pass the full path instead:\n${list}`
    );
  }

  throw new Error(
    `Attachment not found: "${rawPath}". Looked for the exact path`
    + (ATTACH_ROOT ? ` under "${ATTACH_ROOT}"` : '')
    + `, then searched the attachment library for a file named "${wanted}". `
    + `Check the name, or list the folder first.`
  );
}

async function fetchOneDriveAttachment(token, item) {
  const buf = await httpsGetBuffer(`${item.base}/items/${item.id}/content`,
                                   { 'Authorization': `Bearer ${token}` });

  if (buf.length > MAX_ATTACHMENT_BYTES) {
    throw new Error(
      `Attachment "${item.name}" is ${(buf.length / 1024 / 1024).toFixed(1)} MB — `
      + `Graph refuses attachments over 3 MB on a direct send.`
    );
  }

  return {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: item.name,
    contentType: mimeFor(item.name),
    contentBytes: buf.toString('base64'),
    _bytes: buf.length        // stripped before sending; used for the result message
  };
}

// ── Tools ─────────────────────────────────────────────────────────────────────

const TOOLS = [{
  name: 'send_email',
  description: `Send an email. Defaults to ${SENDER_EMAIL}; pass from to send as another allowed mailbox `
    + `(currently ${allowedSenders().join(' or ')}). Anything else is refused. `
    + `Payroll notices go from payroll@thecachegroup.com.au — contractors know that address. `
    + `Shows a preview unless confirm is true. Appends the matching TCG signature automatically — `
    + `the payroll signature carries no personal name, title or mobile. `
    + `To attach a file, prefer attach_from_onedrive — pass the file's path relative to the OneDrive root `
    + `(e.g. "CONTRACTOR AGREEMENTS/Devinia Liddelow/Consultancy Brief Devinia Liddelow 19022027.docx") and the server `
    + `fetches it from OneDrive itself. Use the attachments parameter only for files that do not exist in OneDrive, `
    + `and only when small — it requires the whole file base64-encoded inline, which is unreliable above a few KB.`,
  inputSchema: {
    type: 'object',
    required: ['to', 'subject', 'body'],
    properties: {
      from: {
        type: 'string',
        description: 'Mailbox to send as. Omit for ' + SENDER_EMAIL + '. '
          + 'Allowed: ' + allowedSenders().join(', ') + '. Any other address is refused.'
      },
      to: { type: 'string', description: 'Recipient email. Comma-separate for multiple.' },
      subject: { type: 'string', description: 'Subject line' },
      body: { type: 'string', description: 'Plain-text body (signature appended automatically)' },
      cc: { type: 'string', description: 'CC address (optional)' },
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
            name: { type: 'string' },
            contentType: { type: 'string' },
            contentBytes: { type: 'string', description: 'Base64-encoded file content' }
          }
        }
      }
    }
  }
}];

async function callSendEmail(args) {
  const { from, to, subject, body, cc, confirm, attachments, attach_from_onedrive } = args;

  // Resolved before anything else, so an address that is not on the allowlist
  // is refused at the door — in preview as well as on send. A preview that
  // shows a FROM the server would go on to refuse is worse than no preview.
  const profile = resolveSender(from);

  // Validated before the preview branch. Without this, a missing `to` gives a
  // clean-looking preview reading "TO: undefined" and then throws deep inside
  // the send on `to.split`. Pre-existing on main; cheap to close here.
  if (typeof to !== 'string' || !to.trim()) {
    throw new Error('`to` is required and must be a non-empty string.');
  }
  if (typeof body !== 'string' || !body.trim()) {
    throw new Error('`body` is required and must be a non-empty string.');
  }
  if (typeof subject !== 'string' || !subject.trim()) {
    throw new Error('`subject` is required and must be a non-empty string.');
  }

  // Parsed here, not at send time. `to: ","` has a truthy trim, so the check
  // above passes it and Graph would be handed two empty recipients.
  const toList = to.split(',').map(a => a.trim()).filter(Boolean);
  if (!toList.length) {
    throw new Error('`to` contained no usable addresses.');
  }

  // Preview and send must both use the same text, so clean it once, here.
  const cleanBody = stripTrailingSignOff(body, knownSenderNames());

  // Belt and braces over the stripper. That only trims a sign-off and a name
  // off the END of a body — a name above a phone number, or anywhere mid-body,
  // survives it untouched. Anything that identifies a DIFFERENT sender is
  // refused rather than posted to a contractor.
  //
  // THE DISCRIMINATOR IS LINE SHAPE, NOT POSITION. A name or an address is
  // impersonation when it STANDS ALONE on a line — that is a signature. It is
  // ordinary content when it sits inside a sentence: "approved by Andrew
  // Hurnard", "send your timesheets to payroll@thecachegroup.com.au". Those
  // are among the most common sentences this system writes, and a guard that
  // refuses them gets switched off within a fortnight, after which it protects
  // nothing at all.
  //
  // An earlier attempt scoped by POSITION instead — the last six lines — and
  // was wrong in both directions. Payroll emails are short, so the window
  // usually spanned the whole body and scoped nothing; and appending a
  // six-line disclaimer footer beneath a genuine pasted signature pushed that
  // signature out of the window, defeating the check entirely.
  //
  // A PHONE NUMBER is matched anywhere in the body instead. It never appears
  // innocently, and since every Australian mobile is 04XX XXX XXX the last
  // nine digits cannot collide between two different mobiles — the check can
  // only ever match a genuine appearance of that number.
  //
  // Job titles and single-word names are never checked: "Director" and
  // "payroll" are ordinary English.
  const normWs   = s => String(s).toLowerCase().replace(/\s+/g, ' ').trim();
  const digitsOf = s => String(s).replace(/\D/g, '');

  const bodyDigits = digitsOf(cleanBody);

  // Lines whose entire content is one identifier, allowing for the decoration
  // a pasted signature carries: "Andrew Hurnard", "-- Andrew Hurnard",
  // "Andrew Hurnard |".
  //
  // Only genuine signature delimiters are stripped. RFC 3676 makes "-- " the
  // delimiter; a single "- " is a BULLET. Stripping both turned a line of a
  // list of outstanding timesheets — "- Andrew Hurnard" — into a refusal.
  const stripDecor = l => l
    .replace(/^(?:--+|[–—*|]+)\s*/, '')
    .replace(/[,.!|\s]+$/, '');

  // "Andrew Hurnard <andrew.hurnard@thecachegroup.com.au>" is ONE line
  // carrying TWO identifiers, and on its own matches neither. That is exactly
  // how Outlook pastes a contact, which makes it the likeliest shape a real
  // signature arrives in — and it defeated an earlier version of this check
  // completely. Expand such a line into itself plus both halves.
  const expandPair = l => {
    const m = l.match(/^(.*?)\s*<\s*([^<>]+?)\s*>$/);
    return m ? [l, m[1], m[2]] : [l];
  };

  const standaloneLines = cleanBody.split('\n')
    .flatMap(l => expandPair(stripDecor(normWs(l))))
    .filter(Boolean);

  for (const other of new Set(Object.values(SENDERS))) {
    // Identity, not object equality. careers@ is a DIFFERENT profile carrying
    // the SAME person's name and mobile, so `other === profile` treats each as
    // impersonating the other: Matt sending from matt@ with his own mobile in
    // the body would be refused for "impersonating" his own careers@ profile.
    // A profile with no identity falls back to its address so an older or
    // hand-added profile still gets compared rather than silently skipped.
    const otherId   = other.identity   || String(other.address || '').toLowerCase();
    const profileId = profile.identity || String(profile.address || '').toLowerCase();
    if (otherId === profileId) continue;

    // Last nine digits: survives +61 vs 0, spaces, hyphens and run-together.
    if (other.phone) {
      const tail9 = digitsOf(other.phone).slice(-9);
      if (tail9.length === 9 && bodyDigits.includes(tail9)) {
        throw new Error(
          `Body contains the phone number "${other.phone}", which belongs to a `
          + `different sender than ${profile.address}. Refusing to send.`
        );
      }
    }

    // Indexed, not value-compared. `field === other.name` exempts by VALUE, so
    // a profile whose address happened to equal its own single-word name would
    // silently skip the ADDRESS check as well. Position says what was meant.
    for (const [i, field] of [other.name, other.address].entries()) {
      if (!field) continue;
      // Index 0 is the name. A single-word name is ordinary English and is
      // never checked; the address at index 1 is never exempt.
      if (i === 0 && !/\s/.test(field)) continue;
      if (standaloneLines.includes(normWs(field))) {
        throw new Error(
          `Body contains "${field}" on a line of its own, which reads as a `
          + `signature for a different sender than ${profile.address}. `
          + 'Refusing to send.'
        );
      }
    }
  }

  const drivePaths = Array.isArray(attach_from_onedrive) ? attach_from_onedrive.filter(Boolean) : [];
  const inlineAtts = Array.isArray(attachments) ? attachments : [];

  const attachmentSummary = [
    ...drivePaths.map(p => `  - ${p} (from OneDrive)`),
    ...inlineAtts.map(a => `  - ${a.name} (inline)`)
  ].join('\n');

  // The preview is the ONLY thing standing between a draft and a live send,
  // and the question being approved is "does this look like it came from
  // payroll". Saying "[Signature appended]" hides the one part that answers
  // it. Show the real block, and disclose the mailbox when it differs from
  // the From line — payroll@ posts through payrollmb@ and a reader should not
  // have to know that to understand what they are approving.
  const sigPreview = [profile.signOff, profile.name, profile.title, profile.company, profile.phone]
    .filter(v => v && String(v).trim())
    .join('\n');

  const preview = `FROM: ${profile.address}\n`
    + (profile.mailbox.toLowerCase() !== profile.address.toLowerCase()
        ? `VIA MAILBOX: ${profile.mailbox}\n` : '')
    + `TO: ${to}\n`
    + (cc ? `CC: ${cc}\n` : '')
    + `SUBJECT: ${subject}\n`
    + (attachmentSummary ? `ATTACHED:\n${attachmentSummary}\n` : '')
    + `\n${cleanBody}\n\n${sigPreview}\n[TCG logo]`;

  if (!confirm) {
    return { preview: true, text: `PREVIEW (not sent):\n\n${preview}\n\nCall again with confirm: true to send.` };
  }

  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Missing Graph credentials — check Vercel environment variables.');
  }

  const token = await getToken();

  const toRecipients = toList.map(a => ({ emailAddress: { address: a } }));
  const message = {
    subject,
    body: { contentType: 'HTML', content: buildHtmlBody(cleanBody, profile) },
    toRecipients,
    from: { emailAddress: { address: profile.address, name: profile.name } }
  };
  // Same parse as `to`. Without the filter, a trailing comma hands Graph an
  // empty recipient — the exact bug just fixed two lines above.
  if (cc) {
    const ccList = String(cc).split(',').map(a => a.trim()).filter(Boolean);
    if (ccList.length) {
      message.ccRecipients = ccList.map(a => ({ emailAddress: { address: a } }));
    }
  }

  // The signature logo always rides along as an inline cid: part. It is hidden
  // from the attachment list the recipient sees, and from the result message.
  const built = [logoAttachment()];

  // Fetched server-side from OneDrive — the preferred path.
  for (const p of drivePaths) {
    built.push(await fetchOneDriveAttachment(token, p, profile.drive));
  }

  // Legacy inline base64 — kept for files that are not in OneDrive.
  for (const a of inlineAtts) {
    built.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: a.name, contentType: a.contentType, contentBytes: a.contentBytes,
      _bytes: Buffer.from(a.contentBytes, 'base64').length
    });
  }

  const total = built.reduce((n, a) => n + (a._bytes || 0), 0);
  if (total > MAX_ATTACHMENT_BYTES) {
    throw new Error(`Attachments total ${(total / 1024 / 1024).toFixed(1)} MB — over Graph's 3 MB limit for a direct send.`);
  }

  message.attachments = built.map(({ _bytes, _inline, ...rest }) => rest);

  const visible = built.filter(a => !a._inline);
  const sentNote = visible.length
    ? ' with ' + visible.map(a => `${a.name} (${a._bytes.toLocaleString()} bytes)`).join(', ')
    : '';

  await graphSendMail(token, message, profile.mailbox);
  // Graph returns 202 for "accepted", not "delivered as addressed". Exchange
  // normalises From to the mailbox's primary SMTP, so if payroll@ ever stops
  // being payrollmb@'s primary, sends keep returning 202 and recipients
  // quietly see payrollmb@ instead. Report what was actually confirmed rather
  // than asserting a From line nothing verified.
  return {
    preview: false,
    text: `✓ Accepted by Graph for mailbox ${profile.mailbox}, From set to ${profile.address}, to ${to}${sentNote}`
  };
}

// ── MCP router ────────────────────────────────────────────────────────────────

async function handleMcp(rpc) {
  const { method, params, id } = rpc;
  if (method === 'initialize') return { jsonrpc: '2.0', id, result: { protocolVersion: '2024-11-05', capabilities: { tools: {} }, serverInfo: { name: 'ms365-mailer', version: '1.3.0' } } };
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
  // Reports the misconfiguration rather than a cheerful ok. The fail-closed
  // 500 below sits under this branch, so without this an uptime monitor
  // pointed at GET would report green on a server where every send 500s.
  //
  // The sender address is dropped from the payload: it stopped describing what
  // this server does the moment there were two senders, and an endpoint
  // reachable without authentication need not volunteer which identity it is
  // bound to.
  if (req.method === 'GET') {
    return res.status(MCP_SECRET ? 200 : 500).json({
      status: MCP_SECRET ? 'ok' : 'misconfigured: MCP_SHARED_SECRET is not set',
      server: 'ms365-mailer'
    });
  }
  // Fail CLOSED. `if (MCP_SECRET)` meant that an unset environment variable
  // left this endpoint with no authentication at all, while
  // Access-Control-Allow-Origin is '*'. That was already wrong. With a second
  // sender added it becomes a phishing primitive: anyone who found the URL
  // could send mail appearing to come from the address contractors are told
  // to trust for payslips and bank details. This is the same fail-open shape
  // that was fixed in cats-mcp-server on 07/09/2026.
  if (!MCP_SECRET) {
    return res.status(500).json({ error: 'Server misconfigured: MCP_SHARED_SECRET is not set' });
  }
  const reqPath = (req.url || '').split('?')[0];
  if (reqPath !== '/mcp/' + MCP_SECRET && reqPath !== '/' + MCP_SECRET) {
    return res.status(404).json({ error: 'Not found' });
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
