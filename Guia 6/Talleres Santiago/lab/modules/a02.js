/**
 * A02:2025 - Security Misconfiguration
 * Laboratorio educativo local. Simula un panel de administración expuesto
 * por una mala configuración (ruta no protegida + credenciales de fábrica)
 * y mensajes de error que revelan detalles internos.
 */
const crypto = require('crypto');

exports.meta = {
  id: 'A02',
  title: 'Security Misconfiguration',
  short: 'Panel admin expuesto y errores verbosos',
  fragment: 'a02.html',
};

// Credenciales "de fábrica" que nunca se cambiaron: ejemplo clásico de A02.
const FACTORY_CREDS = { user: 'admin', pass: 'admin' };

// Versión segura: la contraseña no se guarda en claro sino como hash scrypt con sal.
// (Laboratorio: la clave de demostración vive aquí; en un sistema real iría en la base de datos.)
const SAFE_USER = 'operador';
const SALT = Buffer.from('sal-de-demostracion-del-laboratorio');
const STORED_HASH = crypto.scryptSync('unaClaveLarga123', SALT, 32);
const sessions = new Set();

function checkPassword(pass) {
  const candidate = crypto.scryptSync(String(pass), SALT, 32);
  return crypto.timingSafeEqual(candidate, STORED_HASH);
}

exports.mount = (app) => {
  // ---- Versión vulnerable: el panel no pide sesión, solo compara credenciales
  // de fábrica en cada petición, y el error de servidor expone la pila interna.
  app.get('/a02/vuln/admin-panel', (_req, res) => {
    res.json({
      panel: 'admin',
      nota: 'Esta ruta no exige sesión ni rol: cualquiera que la conozca entra.',
      configuracionExpuesta: {
        version: 'app-demo 1.0.0',
        entorno: 'produccion',
        credencialesFabrica: FACTORY_CREDS, // nunca deberían verse ni existir ya
      },
    });
  });

  app.post('/a02/vuln/login', (req, res) => {
    const { user, pass } = req.body || {};
    if (user === FACTORY_CREDS.user && pass === FACTORY_CREDS.pass) {
      return res.json({ ok: true, rol: 'administrador' });
    }
    // El mensaje de error vulnerable revela estructura interna.
    res.status(500).json({
      ok: false,
      error: 'ReferenceError: tabla "usuarios_admin" en modulo auth.js linea 42',
      stackSimulado: 'at validarCredenciales (auth.js:42:11)\n at /a02/vuln/login (server.js:88:5)',
    });
  });

  // ---- Versión segura: sin credenciales de fábrica, sesión real y error genérico ----
  app.post('/a02/safe/login', (req, res) => {
    const { user, pass } = req.body || {};
    if (user === SAFE_USER && checkPassword(pass)) {
      const sesion = crypto.randomBytes(16).toString('hex');
      sessions.add(sesion);
      return res.json({ ok: true, rol: 'operador', sesion });
    }
    res.status(401).json({ ok: false, error: 'Credenciales inválidas.' });
  });

  app.get('/a02/safe/admin-panel', (req, res) => {
    if (!sessions.has(req.header('x-session'))) {
      return res.status(401).json({ error: 'Se requiere sesión válida. Sin sesión no hay acceso al panel.' });
    }
    res.json({ panel: 'admin', nota: 'Acceso con sesión verificada; sin datos sensibles en la respuesta.' });
  });
};
