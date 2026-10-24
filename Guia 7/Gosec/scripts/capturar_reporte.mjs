/**
 * Abre el informe HTML que genera gosec en Edge real y guarda capturas para la presentacion.
 * Requiere internet (el informe carga sus estilos desde un CDN) y Windows con Microsoft Edge.
 *   node scripts/capturar_reporte.mjs
 */
import { abrir } from './navegador.mjs';
import { mkdirSync } from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { dirname, join } from 'node:path';

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..');
const informe = pathToFileURL(join(raiz, 'resultados', 'govwa.html')).href;
const salida = join(raiz, 'presentacion', 'img');
mkdirSync(salida, { recursive: true });

const b = await abrir({ ancho: 1440, alto: 1000, escala: 1, puerto: 9344 });
try {
  await b.ir(informe);
  await b.esperar(3000);                      // babel compila el informe en el navegador
  await b.captura(join(salida, 'reporte_html_inicio.png'));
  const alto = await b.evaluar('document.documentElement.scrollHeight');
  console.log('alto total del informe:', alto, 'px');
  console.log('titulos de la pagina:', JSON.stringify(await b.evaluar(
    `[...document.querySelectorAll('h1,h2,h3,.title,.subtitle')].slice(0,12).map(e => e.textContent.trim().slice(0,70))`)));
  console.log('errores de consola:', b.errores.length ? b.errores : 0);
} finally {
  await b.cerrar();
}
