'use strict';
// Minimal https stub. Routes are set per test through global.__ROUTES().
const { EventEmitter } = require('events');

function routeFor(p) {
  const R = global.__ROUTES();
  if (/root\/search/.test(p)) {
    if (R.onSearch) R.onSearch();
    return (R.get || {}).SEARCH;
  }
  return (R.get || {})[p];
}

exports.request = function (opts, cb) {
  const req = new EventEmitter();
  req.write = () => {};
  req.end = function () {
    setImmediate(() => {
      const R = global.__ROUTES();
      let r = opts.method === 'POST' ? (R.post || { status: 200, body: '{}' }) : routeFor(opts.path);
      if (!r) r = { status: 404, body: JSON.stringify({ error: { code: 'itemNotFound', routed: opts.path } }) };
      const res = new EventEmitter();
      res.statusCode = r.status;
      res.headers = {};
      res.resume = () => {};
      cb(res);
      setImmediate(() => { res.emit('data', Buffer.from(r.body)); res.emit('end'); });
    });
  };
  return req;
};
