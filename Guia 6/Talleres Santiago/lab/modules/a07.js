/**
 * A07:2025 - Authentication Failures (enfoque JWT con firma no verificada)
 * Laboratorio educativo local. Emite tokens firmados a mano (HMAC simplificado,
 * NO una librería JWT real) para poder mostrar, de forma controlada, qué pasa
 * cuando el servidor no valida la firma antes de confiar en el contenido.
 */
const crypto = require('crypto');

exports.meta = {
  id: 'A07',
  title: 'Authentication Failures: JWT',
  short: 'Token sin verificar la firma vs. verificado',
  fragment: 'a07.html',
};

const SECRET = 'clave-demo-solo-para-este-laboratorio'; // nunca una clave real
function b64url(obj) { return Buffer.from(JSON.stringify(obj)).toString('base64url'); }
function sign(payload) {
  const header = b64url({ alg: 'HS256', typ: 'JWT' });
  const body = b64url(payload);
  const sig = crypto.createHmac('sha256', SECRET).update(header + '.' + body).digest('base64url');
  return `${header}.${body}.${sig}`;
}
function decodePayload(token) {
  const parts = String(token).split('.');
  if (parts.length !== 3) throw new Error('Formato de token invalido');
  return JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
}
function verify(token) {
  try {
    const parts = String(token).split('.');
    if (parts.length !== 3) return null;
    const expected = crypto.createHmac('sha256', SECRET).update(parts[0] + '.' + parts[1]).digest();
    const given = Buffer.from(parts[2], 'base64url');
    // timingSafeEqual exige el mismo largo: una firma de otro largo es inválida, no un error.
    if (given.length !== expected.length || !crypto.timingSafeEqual(expected, given)) return null;
    return JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
  } catch {
    return null;
  }
}

exports.mount = (app) => {
  // Emite un token normal para el usuario "invitado" (rol bajo).
  app.post('/a07/login', (req, res) => {
    const username = (req.body || {}).username || 'invitado';
    const token = sign({ sub: username, rol: 'invitado', iat: Date.now() });
    res.json({ token });
  });

  // ---- Versión vulnerable: decodifica el payload y confía en él SIN comprobar la firma ----
  app.get('/a07/vuln/perfil', (req, res) => {
    const token = (req.header('authorization') || '').replace(/^Bearer\s+/i, '');
    if (!token) return res.status(401).json({ error: 'Falta token' });
    try {
      const payload = decodePayload(token); // <- no se llama a verify()
      res.json({ ok: true, payload, nota: 'Esta ruta nunca comprobó la firma del token.' });
    } catch (e) {
      res.status(400).json({ error: 'Token malformado', detalle: e.message });
    }
  });

  // ---- Versión segura: exige que la firma sea válida antes de confiar en el payload ----
  app.get('/a07/safe/perfil', (req, res) => {
    const token = (req.header('authorization') || '').replace(/^Bearer\s+/i, '');
    if (!token) return res.status(401).json({ error: 'Falta token' });
    const payload = verify(token);
    if (!payload) return res.status(401).json({ error: 'Firma inválida: el token fue alterado o no corresponde a este servidor.' });
    res.json({ ok: true, payload });
  });
};
