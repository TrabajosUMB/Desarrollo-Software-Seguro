/**
 * Recorre el laboratorio con clics de ratón reales en Edge, comprueba cada respuesta y su veredicto
 * y guarda la evidencia del Word en docs/evidencia. Requiere el laboratorio corriendo (npm start).
 *   node docs/herramientas/capturar_evidencia.mjs [urlBase]
 */
import { abrir } from './navegador.mjs';
import { mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import assert from 'node:assert/strict';

const BASE = process.argv[2] || 'http://127.0.0.1:3000/';
const SALIDA = join(dirname(fileURLToPath(import.meta.url)), '..', 'evidencia');
mkdirSync(SALIDA, { recursive: true });

const P = (n) => `.compare > .panel:nth-child(${n})`;       // panel n del escenario (1 vulnerable, 2 segura)
const BTN = (n) => `${P(n)} .btn`;
const texto = (id) => `document.getElementById(${JSON.stringify(id)}).textContent`;
const veredicto = (id) => `(document.getElementById(${JSON.stringify(id)}).closest('.out').querySelector('[data-verdict]') || {}).textContent || ''`;
const chip = (id) => `document.getElementById(${JSON.stringify(id)}).closest('.out').querySelector('[data-chip]').textContent`;

// ventana alta (2400 px) para que nada se desplace: el recorte de CDP usa coordenadas del documento
const b = await abrir({ ancho: 1440, alto: 2400, escala: 1.5, puerto: 9342 });
const ok = (cond, msg) => { assert.ok(cond, msg); console.log('  ok  ' + msg); };
const incluye = async (expr, esperado, msg) => {
  const v = String(await b.evaluar(expr));
  ok(v.includes(esperado), `${msg} (${esperado})`);
};

try {
  console.log('A01');
  await b.ir(BASE + '#A01');
  await b.clic(BTN(1), 'Pedir vista previa');
  await incluye(texto('a01-vuln-out'), 'inventario-interno', 'la vulnerable devuelve el servicio interno');
  await incluye(chip('a01-vuln-out'), '200', 'código HTTP de la vulnerable');
  await incluye(veredicto('a01-vuln-out'), 'El ataque funcionó', 'veredicto de la vulnerable');
  await b.captura(join(SALIDA, 'a01_vuln_ssrf.png'), [P(1)]);
  await b.clic(BTN(2), 'Probar con la URL interna');
  await incluye(texto('a01-safe-out'), 'Identificador no permitido', 'la segura rechaza la URL');
  await incluye(chip('a01-safe-out'), '403', 'código HTTP de la segura');
  await incluye(veredicto('a01-safe-out'), 'El ataque fue bloqueado', 'veredicto de la segura');
  await b.captura(join(SALIDA, 'a01_comparativo.png'), ['.compare']);

  console.log('A02');
  await b.ir(BASE + '#A02');
  await b.clic(BTN(1), 'Abrir panel (sin sesión)');
  await b.clic(BTN(1), 'Entrar con admin/admin');
  await incluye(texto('a02-vuln-panel'), 'credencialesFabrica', 'el panel vulnerable abre sin sesión');
  await incluye(texto('a02-vuln-login'), 'administrador', 'admin/admin entra como administrador');
  await incluye(veredicto('a02-vuln-login'), 'El ataque funcionó', 'veredicto del inicio de sesión vulnerable');
  await b.captura(join(SALIDA, 'a02_vuln_login.png'), [P(1)]);
  ok(await b.evaluar(`document.querySelector('${P(2)} [data-session="1"]').disabled`), 'abrir con sesión está deshabilitado antes de entrar');
  await b.clic(BTN(2), 'Abrir panel (sin sesión)');
  await b.clic(BTN(2), 'Probar admin/admin');
  await incluye(texto('a02-safe-panel'), 'Se requiere sesión', 'el panel seguro exige sesión');
  await incluye(texto('a02-safe-login'), 'Credenciales inválidas', 'admin/admin es rechazado');
  await incluye(veredicto('a02-safe-login'), 'El ataque fue bloqueado', 'veredicto de la segura');
  await b.captura(join(SALIDA, 'a02_seguro_rechaza.png'), [P(2)]);
  await b.clic(BTN(2), 'Entrar con credencial válida');
  await b.clic(BTN(2), 'Abrir panel (con sesión)');
  await incluye(texto('a02-safe-login'), '"sesion"', 'la credencial válida entrega una sesión');
  await incluye(texto('a02-safe-panel'), 'Acceso con sesión verificada', 'con sesión el panel responde');
  await incluye(veredicto('a02-safe-panel'), 'Uso legítimo permitido', 'el uso legítimo se marca como permitido');
  await b.captura(join(SALIDA, 'a02_seguro_con_sesion.png'), [P(2)]);

  console.log('A05');
  await b.ir(BASE + '#A05');
  await b.esperar(500);
  ok(b.dialogos.length === 0, 'al entrar no se dispara ningún alert de ensayos anteriores');
  await b.clic(BTN(1), 'Publicar etiqueta HTML');
  await b.clic(BTN(2), 'Publicar etiqueta HTML');
  await incluye(`document.getElementById('a05-vuln-list').textContent`, 'XSS ejecutado', 'la vulnerable ejecuta la etiqueta y deja su marca');
  await incluye(`document.getElementById('a05-safe-list').textContent`, '<img', 'la segura muestra la etiqueta como texto');
  ok(!(await b.evaluar(`!!document.querySelector('#a05-safe-list img')`)), 'la segura no crea ninguna etiqueta');
  ok(b.dialogos.some((d) => d.includes('alert: 1')), 'el alert(1) del ataque se dispara (cuadro recibido: ' + b.dialogos.join(' | ') + ')');
  await incluye(veredicto('a05-vuln-list'), 'El ataque funcionó', 'veredicto de la vulnerable');
  await incluye(veredicto('a05-safe-list'), 'El ataque fue bloqueado', 'veredicto de la segura');
  await b.captura(join(SALIDA, 'a05_comparativo.png'), ['.compare']);

  console.log('A07');
  await b.ir(BASE + '#A07');
  ok(await b.evaluar(`document.querySelector('[data-act="forge"]').disabled`), 'alterar rol está deshabilitado antes de tener token');
  await b.clic('.panel--wide .btn', 'Obtener token');
  await b.clic('.panel--wide .btn', 'Alterar rol a admin');
  await incluye(`document.getElementById('a07-change').textContent`, 'invitado → admin', 'el cambio de rol se muestra en grande');
  await b.clic(BTN(1), 'Consultar perfil (vulnerable)');
  await b.clic(BTN(2), 'Consultar perfil (seguro)');
  await incluye(texto('a07-vuln-out'), '"rol": "admin"', 'la vulnerable acepta el rol admin falsificado');
  await incluye(veredicto('a07-vuln-out'), 'El ataque funcionó', 'veredicto de la vulnerable');
  await incluye(texto('a07-safe-out'), 'Firma inválida', 'la segura rechaza la firma');
  await incluye(chip('a07-safe-out'), '401', 'código HTTP de la segura');
  await incluye(veredicto('a07-safe-out'), 'El ataque fue bloqueado', 'veredicto de la segura');
  await b.captura(join(SALIDA, 'a07_comparativo.png'), ['.panel--wide', P(1), P(2)]);

  ok(b.errores.length === 0, 'sin errores en consola: ' + (b.errores.join(' | ') || '0'));
  console.log('\nEvidencia guardada en', SALIDA);
} finally {
  await b.cerrar();
}
