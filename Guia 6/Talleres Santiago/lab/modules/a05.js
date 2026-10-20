/**
 * A05:2025 - Injection (enfoque XSS almacenado)
 * Laboratorio educativo local. Un libro de comentarios que en la version
 * vulnerable refleja HTML sin escapar, y en la version segura lo escapa.
 * El "payload" de ejemplo solo muestra una alerta en el propio navegador
 * de quien lo prueba; no se envía a ningún tercero.
 */
exports.meta = {
  id: 'A05',
  title: 'Injection: XSS almacenado',
  short: 'Comentarios sin escapar vs. escapados',
  fragment: 'a05.html',
};

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

let vulnComments = [];
let safeComments = [];

exports.mount = (app) => {
  // ---- Versión vulnerable: el comentario se guarda y se devuelve tal cual ----
  app.get('/a05/vuln/comments', (_req, res) => res.json({ comments: vulnComments }));
  app.post('/a05/vuln/comments', (req, res) => {
    const text = String((req.body || {}).text || '').slice(0, 500);
    vulnComments.push(text); // sin sanitizar ni escapar
    res.json({ ok: true, totalComentarios: vulnComments.length });
  });
  app.post('/a05/vuln/reset', (_req, res) => { vulnComments = []; res.json({ ok: true }); });

  // ---- Versión segura: se guarda el texto, pero se escapa antes de mostrarlo ----
  app.get('/a05/safe/comments', (_req, res) =>
    res.json({ comments: safeComments.map(escapeHtml) }));
  app.post('/a05/safe/comments', (req, res) => {
    const text = String((req.body || {}).text || '').slice(0, 500);
    safeComments.push(text);
    res.json({ ok: true, totalComentarios: safeComments.length });
  });
  app.post('/a05/safe/reset', (_req, res) => { safeComments = []; res.json({ ok: true }); });
};
