/**
 * Laboratorio OWASP Top 10:2025 - A01 (SSRF), A02, A05 (XSS), A07 (JWT)
 * USO EXCLUSIVO EDUCATIVO. Solo escucha en 127.0.0.1. No contiene secretos reales.
 *
 * Contrato de cada modulo (modules/aXX.js):
 *   exports.meta  = { id: 'A01', title: '...', short: '...', fragment: 'a01.html' }
 *   exports.mount = (app, ctx) => { ... }   // rutas bajo /aXX/vuln/* y /aXX/safe/*
 *   ctx = { port, express }
 */
const express = require('express');
const fs = require('fs');
const path = require('path');

const PORT = Number(process.env.PORT || 3000);
const app = express();
app.disable('x-powered-by');
app.use(express.json());
app.use(express.urlencoded({ extended: false }));

const loaded = [];
const dir = path.join(__dirname, 'modules');
for (const f of fs.readdirSync(dir).filter((n) => /^a\d\d\.js$/.test(n)).sort()) {
  try {
    const m = require(path.join(dir, f));
    m.mount(app, { port: PORT, express });
    loaded.push(m.meta);
    console.log('[lab] modulo cargado:', m.meta.id, '-', m.meta.title);
  } catch (e) {
    console.error('[lab] no se pudo cargar', f, '->', e.message);
  }
}

app.get('/api/modules', (_req, res) => res.json(loaded));
app.use(express.static(path.join(__dirname, 'public')));

app.listen(PORT, '127.0.0.1', () => console.log(`[lab] http://127.0.0.1:${PORT}`));
