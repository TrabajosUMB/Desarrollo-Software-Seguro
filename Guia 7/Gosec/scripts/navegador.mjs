/**
 * Control mínimo de Edge real por CDP (sin paquetes). Node 24 trae WebSocket y fetch globales.
 * Lo usan las capturas de evidencia y la inspección de diseño; todo ocurre en el equipo local.
 */
import { spawn, execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export async function abrir({ puerto = 9340, ancho = 1366, alto = 768, escala = 1, movimientoReducido = true } = {}) {
  const perfil = mkdtempSync(join(tmpdir(), 'edge-lab-'));
  const proc = spawn(EDGE, [
    '--headless=new', `--remote-debugging-port=${puerto}`, `--user-data-dir=${perfil}`,
    '--no-first-run', '--no-default-browser-check', '--disable-extensions', 'about:blank',
  ], { stdio: 'ignore' });

  let objetivo;
  for (let i = 0; i < 80 && !objetivo; i++) {
    try {
      const lista = await (await fetch(`http://127.0.0.1:${puerto}/json`)).json();
      objetivo = lista.find((t) => t.type === 'page');
    } catch { /* Edge aun arranca */ }
    if (!objetivo) await sleep(150);
  }
  if (!objetivo) { proc.kill(); throw new Error('Edge no abrió el puerto de depuración'); }

  const ws = new WebSocket(objetivo.webSocketDebuggerUrl);
  await new Promise((ok, fail) => { ws.onopen = ok; ws.onerror = () => fail(new Error('WebSocket CDP')); });

  let id = 0;
  const pendientes = new Map();
  const errores = [];
  const dialogos = [];
  ws.onmessage = (ev) => {
    const m = JSON.parse(ev.data);
    if (m.method === 'Page.javascriptDialogOpening') {
      // un alert() bloquea la pagina hasta que se atiende; se registra y se acepta
      dialogos.push(`${m.params.type}: ${m.params.message}`);
      cdp('Page.handleJavaScriptDialog', { accept: true }).catch(() => {});
    } else if (m.id && pendientes.has(m.id)) {
      const { ok, fail } = pendientes.get(m.id);
      pendientes.delete(m.id);
      m.error ? fail(new Error(m.error.message)) : ok(m.result);
    } else if (m.method === 'Runtime.exceptionThrown') {
      errores.push(m.params.exceptionDetails.exception?.description || m.params.exceptionDetails.text);
    } else if (m.method === 'Runtime.consoleAPICalled' && m.params.type === 'error') {
      errores.push('console.error: ' + m.params.args.map((a) => a.value ?? a.description).join(' '));
    }
  };
  const cdp = (method, params = {}) => new Promise((ok, fail) => {
    const n = ++id;
    pendientes.set(n, { ok, fail });
    ws.send(JSON.stringify({ id: n, method, params }));
  });

  await cdp('Page.enable');
  await cdp('Runtime.enable');
  await cdp('Emulation.setDeviceMetricsOverride', { width: ancho, height: alto, deviceScaleFactor: escala, mobile: ancho < 600 });
  // con movimiento reducido el desplazamiento es instantaneo: capturas deterministas
  if (movimientoReducido) await cdp('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-reduced-motion', value: 'reduce' }] });

  const api = {
    errores,
    dialogos,
    cdp,
    async ir(url) { await cdp('Page.navigate', { url }); await sleep(700); },
    async evaluar(expresion) {
      const r = await cdp('Runtime.evaluate', { expression: expresion, awaitPromise: true, returnByValue: true });
      if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description || 'error al evaluar');
      return r.result.value;
    },
    async esperar(ms) { await sleep(ms); },
    /** Clic con eventos de ratón reales sobre el primer elemento que coincide con el selector y el texto. */
    async clic(selector, texto) {
      const buscar = `((s, t) => [...document.querySelectorAll(s)].filter((e) => !t || e.textContent.includes(t))[0])(${JSON.stringify(selector)}, ${JSON.stringify(texto || '')})`;
      for (let intento = 1; intento <= 5; intento++) {
        const existe = await api.evaluar(`!!${buscar}`);
        if (!existe) throw new Error(`no se encontró ${selector} ${texto || ''}`);
        await api.evaluar(`${buscar}.scrollIntoView({ block: 'center' })`);
        // esperar a que el desplazamiento termine antes de medir
        let previo = -1;
        for (let i = 0; i < 20; i++) {
          const y = await api.evaluar('scrollY');
          if (y === previo) break;
          previo = y;
          await sleep(60);
        }
        const punto = await api.evaluar(`(() => { const e = ${buscar}; const b = e.getBoundingClientRect(); const x = b.x + b.width / 2, y = b.y + b.height / 2; const arriba = document.elementFromPoint(x, y); return { x, y, acierta: !!arriba && (e === arriba || e.contains(arriba)) }; })()`);
        if (!punto.acierta) { await sleep(150); continue; }   // algo se movió o tapa el botón: reintentar
        for (const tipo of ['mouseMoved', 'mousePressed', 'mouseReleased']) {
          await cdp('Input.dispatchMouseEvent', { type: tipo, x: punto.x, y: punto.y, button: 'left', clickCount: 1 });
        }
        await sleep(450);
        return;
      }
      throw new Error(`el punto de clic no coincide con ${selector} ${texto || ''}`);
    },
    async tecla(key, code) {
      for (const type of ['keyDown', 'keyUp']) await cdp('Input.dispatchKeyEvent', { type, key, code: code || key });
      await sleep(350);
    },
    /**
     * Guarda una captura PNG. Con `selectores` recorta a la union de esas cajas, sin salir de la
     * vista (la region debe caber en el alto de la ventana: no se usa captura "mas alla" porque
     * altera el estado de la pagina y deja clics perdidos).
     */
    async captura(ruta, selectores) {
      let clip;
      if (selectores) {
        // el clip de CDP usa coordenadas del documento: solo coincide con las de la vista si no hay desplazamiento
        if ((await api.evaluar('scrollY')) !== 0) { await api.evaluar('scrollTo(0, 0)'); await sleep(150); }
        const cajas = await api.evaluar(`${JSON.stringify(selectores)}.map((s) => { const b = document.querySelector(s).getBoundingClientRect(); return { x: b.x, y: b.y, w: b.width, h: b.height }; })`);
        const x = Math.min(...cajas.map((c) => c.x)), y = Math.min(...cajas.map((c) => c.y));
        const x2 = Math.max(...cajas.map((c) => c.x + c.w)), y2 = Math.max(...cajas.map((c) => c.y + c.h));
        const m = 8;
        const alto = await api.evaluar('innerHeight');
        if (y - m < 0 || y2 + m > alto) throw new Error(`la region ${selectores.join(' + ')} no cabe en la ventana (${Math.round(y2 - y)} px de alto)`);
        clip = { x: Math.max(0, x - m), y: y - m, width: x2 - x + 2 * m, height: y2 - y + 2 * m, scale: 1 };
      }
      // el puntero se aparta para que ningun boton salga en estado hover
      await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: 1, y: 1 });
      await sleep(150);
      const r = await cdp('Page.captureScreenshot', { format: 'png', ...(clip ? { clip } : {}) });
      writeFileSync(ruta, Buffer.from(r.data, 'base64'));
    },
    async cerrar() {
      try { ws.close(); } catch { /* ya cerrado */ }
      try { execFileSync('taskkill', ['/PID', String(proc.pid), '/T', '/F'], { stdio: 'ignore' }); } catch { /* ya termino */ }
      // Edge lanza procesos hijos que ya no cuelgan de `proc`: se cierran por su perfil temporal
      const filtro = `$_.CommandLine -like '*${perfil}*'`;
      try {
        execFileSync('powershell', ['-NoProfile', '-Command',
          `Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { ${filtro} } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }`], { stdio: 'ignore' });
      } catch { /* nada que cerrar */ }
    },
  };
  return api;
}
