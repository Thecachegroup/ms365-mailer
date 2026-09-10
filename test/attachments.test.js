'use strict';
// Attachment resolution tests. No network: https is stubbed by test/fake_https.js.
//
// These cover the three things that made "attach the document" fail for Matt
// and Nuria — the wrong drive, no search, and a preview that verified nothing.
const fs = require('fs');
const path = require('path');
const Module = require('module');

const TARGET = path.join(__dirname, '.mcp_test_target.js');
fs.writeFileSync(TARGET,
  fs.readFileSync(path.join(__dirname, '..', 'api', 'mcp.js'), 'utf8')
  + '\nmodule.exports._t = { resolveOneDriveItem, driveBase, withAttachRoot, encodeDrivePath, callSendEmail };\n');

const origLoad = Module._load;
Module._load = function (req) {
  if (req === 'https') return require(path.join(__dirname, 'fake_https.js'));
  return origLoad.apply(this, arguments);
};

let ROUTES = {};
global.__ROUTES = () => ROUTES;

process.env.GRAPH_TENANT_ID = 't';
process.env.GRAPH_CLIENT_ID = 'c';
process.env.GRAPH_CLIENT_SECRET = 's';
process.env.ATTACH_DRIVE_ID = 'b!SHARED';

delete require.cache[TARGET];
const T = require(TARGET)._t;

let pass = 0, fail = 0;
function ok(n, c, extra) {
  if (c) { pass++; console.log('  PASS', n); }
  else { fail++; console.log('  FAIL', n, extra || ''); }
}
async function throws(n, fn, re) {
  try { await fn(); fail++; console.log('  FAIL', n, '(did not throw)'); }
  catch (e) {
    if (re.test(e.message)) { pass++; console.log('  PASS', n); }
    else { fail++; console.log('  FAIL', n, 'msg:', e.message); }
  }
}

(async () => {
  console.log('1. exact path on the shared drive');
  ROUTES = { get: { '/v1.0/drives/b!SHARED/root:/AI%20Working%20Folder/x.docx':
    { status: 200, body: JSON.stringify({ id: 'ID1', name: 'x.docx', size: 2048, file: {} }) } } };
  let r = await T.resolveOneDriveItem('tok', 'AI Working Folder/x.docx', { drive: 'a@b.c' });
  ok('reads the shared library, not the sender drive',
     r.base === 'https://graph.microsoft.com/v1.0/drives/b!SHARED', r.base);
  ok('resolved by path', r.how === 'path' && r.id === 'ID1' && r.size === 2048);

  console.log('2. a bare filename falls back to search');
  ROUTES = { get: {
    '/v1.0/drives/b!SHARED/root:/brief.docx': { status: 404, body: '{"error":{"code":"itemNotFound"}}' },
    SEARCH: { status: 200, body: JSON.stringify({ value: [
      { id: 'ID2', name: 'brief.docx', size: 1024, file: {},
        parentReference: { path: '/drives/b!SHARED/root:/CONTRACTOR AGREEMENTS/Devinia Liddelow' } }] }) } } };
  r = await T.resolveOneDriveItem('tok', 'brief.docx', { drive: 'a@b.c' });
  ok('found by name', r.how === 'search' && r.id === 'ID2', JSON.stringify(r));
  ok('says which folder it came from', /Devinia Liddelow/.test(r.folder || ''), r.folder);

  console.log('3. an ambiguous name is refused, never guessed');
  ROUTES = { get: {
    '/v1.0/drives/b!SHARED/root:/brief.docx': { status: 404, body: 'x' },
    SEARCH: { status: 200, body: JSON.stringify({ value: [
      { id: 'a', name: 'brief.docx', size: 1, file: {}, parentReference: { path: '/drives/x/root:/One' } },
      { id: 'b', name: 'brief.docx', size: 1, file: {}, parentReference: { path: '/drives/x/root:/Two' } }] }) } } };
  await throws('two matches refused', () => T.resolveOneDriveItem('tok', 'brief.docx', { drive: 'a@b.c' }),
               /matches 2 files/);

  console.log('4. a miss names what it looked for');
  ROUTES = { get: {
    '/v1.0/drives/b!SHARED/root:/nope.docx': { status: 404, body: 'x' },
    SEARCH: { status: 200, body: JSON.stringify({ value: [] }) } } };
  await throws('not-found message is useful', () => T.resolveOneDriveItem('tok', 'nope.docx', { drive: 'a@b.c' }),
               /Attachment not found: "nope\.docx"/);

  console.log('5. a 401 is not silently turned into "not found"');
  let searched = false;
  ROUTES = { onSearch: () => { searched = true; }, get: {
    '/v1.0/drives/b!SHARED/root:/x.docx': { status: 401, body: '{"error":"unauthorized"}' },
    SEARCH: { status: 200, body: JSON.stringify({ value: [{ id: 'z', name: 'x.docx', file: {}, size: 1 }] }) } } };
  await throws('401 surfaces as itself', () => T.resolveOneDriveItem('tok', 'x.docx', { drive: 'a@b.c' }),
               /Graph GET 401/);
  ok('a permissions failure never falls through to search', searched === false);

  console.log('6. the preview reports the real file');
  ROUTES = { post: { status: 200, body: JSON.stringify({ access_token: 'tok' }) }, get: {
    '/v1.0/drives/b!SHARED/root:/pack.pdf': { status: 404, body: 'x' },
    SEARCH: { status: 200, body: JSON.stringify({ value: [
      { id: 'P1', name: 'pack.pdf', size: 51200, file: {},
        parentReference: { path: '/drives/b!SHARED/root:/Candidate Submissions' } }] }) } } };
  let res = await T.callSendEmail({ to: 'x@y.com', subject: 'S', body: 'Hello there', attach_from_onedrive: ['pack.pdf'] });
  ok('preview returned', res.preview === true);
  ok('preview shows name, size and where it was found',
     /pack\.pdf \(50 KB\) — found by name in Candidate Submissions/.test(res.text),
     res.text && res.text.split('\n').filter(l => l.includes('pack')).join('|'));

  console.log('7. the preview fails on a file that is not there');
  ROUTES = { post: { status: 200, body: JSON.stringify({ access_token: 'tok' }) }, get: {
    '/v1.0/drives/b!SHARED/root:/ghost.docx': { status: 404, body: 'x' },
    SEARCH: { status: 200, body: JSON.stringify({ value: [] }) } } };
  await throws('no more phantom attachments in a preview',
    () => T.callSendEmail({ to: 'x@y.com', subject: 'S', body: 'Hello there', attach_from_onedrive: ['ghost.docx'] }),
    /Attachment not found/);

  console.log('8. without ATTACH_DRIVE_ID nothing changes');
  delete require.cache[TARGET];
  process.env.ATTACH_DRIVE_ID = '';
  const T2 = require(TARGET)._t;
  ROUTES = { get: { '/v1.0/users/matt%40thecachegroup.com.au/drive/root:/a.docx':
    { status: 200, body: JSON.stringify({ id: 'O1', name: 'a.docx', size: 5, file: {} }) } } };
  const r2 = await T2.resolveOneDriveItem('tok', 'a.docx', { drive: 'matt@thecachegroup.com.au' });
  ok('falls back to the sender drive exactly as before',
     r2.id === 'O1' && /users\/matt/.test(r2.base), r2.base);

  try { fs.unlinkSync(TARGET); } catch (e) {}
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})();
