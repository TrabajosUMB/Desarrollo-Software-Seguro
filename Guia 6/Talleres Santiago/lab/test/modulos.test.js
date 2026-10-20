/**
 * Prueba de humo de los 4 módulos: cada versión vulnerable cae en el ataque y
 * cada versión segura lo rechaza. Ejecutar con `npm test` (cierre antes el laboratorio:
 * usa el puerto interno 4019 de A01).
 */
const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const { spawn } = require('node:child_process');
const path = require('node:path');

const PORT = 3990;
const BASE = `http://127.0.0.1:${PORT}`;
let server;

const json = async (url, opts) => {
  const r = await fetch(BASE + url, opts);
  return { status: r.status, body: await r.json() };
};
const post = (url, data, headers = {}) =>
  json(url, { method: 'POST', headers: { 'content-type': 'application/json', ...headers }, body: JSON.stringify(data) });
const bearer = (t) => ({ headers: { authorization: 'Bearer ' + t } });

before(async () => {
  server = spawn(process.execPath, ['server.js'], {
    cwd: path.join(__dirname, '..'),
    env: { ...process.env, PORT: String(PORT) },
    stdio: 'ignore',
  });
  for (let i = 0; i < 50; i++) {
    try { await fetch(BASE + '/api/modules'); return; } catch { await new Promise((r) => setTimeout(r, 100)); }
  }
  throw new Error('el servidor no arrancó');
});

after(() => { server.kill(); });

test('se cargan los cuatro módulos', async () => {
  const { body } = await json('/api/modules');
  assert.deepEqual(body.map((m) => m.id), ['A01', 'A02', 'A05', 'A07']);
});

test('A01: la versión vulnerable alcanza el servicio interno y la segura rechaza la URL', async () => {
  const v = await json('/a01/vuln/preview?url=' + encodeURIComponent('http://127.0.0.1:4019/'));
  assert.equal(v.status, 200);
  assert.match(v.body.respuestaServidor, /inventario-interno/);
  const fuera = await json('/a01/vuln/preview?url=' + encodeURIComponent('http://example.com/'));
  assert.equal(fuera.status, 400); // guardia del laboratorio
  const s = await json('/a01/safe/preview?id=' + encodeURIComponent('http://127.0.0.1:4019/'));
  assert.equal(s.status, 403);
  assert.equal((await json('/a01/safe/preview?id=logo')).status, 200);
});

test('A02: credenciales de fábrica y panel abierto vs. sesión real', async () => {
  assert.equal((await json('/a02/vuln/admin-panel')).status, 200);
  assert.equal((await post('/a02/vuln/login', { user: 'admin', pass: 'admin' })).status, 200);
  assert.equal((await post('/a02/safe/login', { user: 'admin', pass: 'admin' })).status, 401);
  assert.equal((await post('/a02/safe/login', { user: 'operador', pass: 'cualquiera123' })).status, 401);
  const ok = await post('/a02/safe/login', { user: 'operador', pass: 'unaClaveLarga123' });
  assert.equal(ok.status, 200);
  assert.equal(Object.hasOwn(ok.body, 'hashAlmacenadoDemo'), false);
  assert.equal((await json('/a02/safe/admin-panel')).status, 401);
  assert.equal((await json('/a02/safe/admin-panel', { headers: { 'x-session': 'valida' } })).status, 401);
  assert.equal((await json('/a02/safe/admin-panel', { headers: { 'x-session': ok.body.sesion } })).status, 200);
});

test('A05: la vulnerable devuelve el HTML tal cual y la segura lo escapa', async () => {
  const xss = '<img src=x onerror=alert(1)>';
  await post('/a05/vuln/reset', {}); await post('/a05/safe/reset', {});
  await post('/a05/vuln/comments', { text: xss }); await post('/a05/safe/comments', { text: xss });
  assert.deepEqual((await json('/a05/vuln/comments')).body.comments, [xss]);
  assert.deepEqual((await json('/a05/safe/comments')).body.comments, ['&lt;img src=x onerror=alert(1)&gt;']);
});

test('A07: la vulnerable acepta un token alterado; la segura lo rechaza sin caerse', async () => {
  const { body: { token } } = await post('/a07/login', { username: 'invitado' });
  const [h, p, s] = token.split('.');
  const payload = JSON.parse(Buffer.from(p, 'base64url').toString());
  const falso = [h, Buffer.from(JSON.stringify({ ...payload, rol: 'admin' })).toString('base64url'), s].join('.');
  const v = await json('/a07/vuln/perfil', bearer(falso));
  assert.equal(v.body.payload.rol, 'admin');
  assert.equal((await json('/a07/safe/perfil', bearer(falso))).status, 401);
  assert.equal((await json('/a07/safe/perfil', bearer(token))).status, 200);
  assert.equal((await json('/a07/safe/perfil', bearer(`${h}.${p}.abc`))).status, 401); // firma de otro largo
  assert.equal((await json('/a07/safe/perfil', bearer('a.b.c'))).status, 401);
});
