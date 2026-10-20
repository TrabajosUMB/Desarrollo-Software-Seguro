# Laboratorio OWASP Top 10:2025 (talleres)

- Arranque: `npm install && npm start` -> http://127.0.0.1:3000 (variable `PORT` opcional).
- Pruebas: `npm test` comprueba que cada módulo vulnerable cae en el ataque y cada módulo seguro lo rechaza. Cerrar antes el laboratorio, porque A01 usa el puerto interno 4019.
- Módulos en `modules/aXX.js`, interfaz en `public/fragments/aXX.html`. El contrato de cada módulo está en `server.js`.
- Navegación: la URL (`#A01`, `#A02`, `#A05`, `#A07`) decide qué módulo se ve; los scripts de cada panel se ejecutan aislados y reciben `(root, lab)` (ver `public/index.html`).
- Diseño: todos los valores salen de los tokens de `:root` en `public/index.html`. Dos columnas (vulnerable | segura) desde 1200 px de ancho; en pantallas menores se apilan y el riel pasa a franja horizontal.
- Evidencia: con el laboratorio corriendo, `node ../docs/herramientas/capturar_evidencia.mjs` recorre los cuatro talleres con clics reales en Edge, comprueba cada respuesta y regenera las capturas del Word.
