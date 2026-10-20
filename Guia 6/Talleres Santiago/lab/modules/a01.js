/**
 * A01:2025 - Broken Access Control (enfoque SSRF)
 * Laboratorio educativo local. La guardia de alcance (LAB_HOSTS) limita las
 * peticiones a este mismo equipo, nunca a Internet ni a redes ajenas.
 */
const http = require('http');

exports.meta = {
  id: 'A01',
  title: 'Broken Access Control: SSRF',
  short: 'SSRF contra un servicio interno',
  fragment: 'a01.html',
};

// "Servicio interno" simulado: solo escucha en 127.0.0.1.
const INTERNAL_PORT = 4019;
function startInternalService() {
  http.createServer((_req, res) => {
    res.writeHead(200, { 'content-type': 'application/json' });
    res.end(JSON.stringify({
      servicio: 'inventario-interno',
      nota: 'Este endpoint nunca debería ser alcanzable desde fuera del servidor.',
      claveApiSimulada: 'INV-DEMO-0000-NO-REAL',
    }));
  }).listen(INTERNAL_PORT, '127.0.0.1');
}

// Guardia del laboratorio (no forma parte de la vulnerabilidad): evita que
// este formulario sirva para tocar equipos ajenos al propio computador.
const LAB_HOSTS = new Set(['127.0.0.1', 'localhost']);

exports.mount = (app) => {
  startInternalService();

  // ---- Versión vulnerable: la app descarga "la URL que el usuario pida" ----
  app.get('/a01/vuln/preview', async (req, res) => {
    const target = req.query.url;
    if (!target) return res.status(400).json({ error: 'Falta el parametro url' });
    try {
      const u = new URL(target);
      if (u.protocol !== 'http:' || !LAB_HOSTS.has(u.hostname)) {
        return res.status(400).json({ error: 'Fuera del alcance del laboratorio: solo http hacia 127.0.0.1 o localhost.' });
      }
      const body = await new Promise((resolve, reject) => {
        http.get({ hostname: u.hostname, port: u.port || 80, path: u.pathname + u.search, timeout: 1500 },
          (r) => { let d = ''; r.on('data', (c) => (d += c)); r.on('end', () => resolve(d)); })
          .on('error', reject).on('timeout', function () { this.destroy(); reject(new Error('timeout')); });
      });
      res.json({ solicitado: target, respuestaServidor: body });
    } catch (e) {
      res.status(502).json({ error: 'No se pudo obtener la URL', detalle: e.message });
    }
  });

  // ---- Versión segura: solo se permite un identificador fijo, nunca una URL libre ----
  app.get('/a01/safe/preview', (req, res) => {
    const id = req.query.id;
    if (id !== 'logo') {
      return res.status(403).json({ error: 'Identificador no permitido. Use una lista de valores fijos, nunca una URL arbitraria.' });
    }
    res.json({ solicitado: id, nota: 'La version segura jamas abre una URL que venga del usuario; solo referencias internas controladas.' });
  });
};
