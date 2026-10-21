# Laboratorio OWASP Top 10:2025

Servidor local (Node.js y Express) con cuatro talleres de vulnerabilidades. Cada módulo ofrece el mismo flujo en dos versiones, una **vulnerable** y una **segura**, para atacar una y ver cómo la otra lo bloquea. Cada respuesta muestra su código HTTP y un veredicto.

| Módulo | Tema |
|---|---|
| A01 | Broken Access Control: SSRF contra un servicio interno |
| A02 | Security Misconfiguration: panel abierto, cuenta de fábrica y errores verbosos |
| A05 | Injection: XSS almacenado |
| A07 | Authentication Failures: token cuya firma no se verifica |

## Arranque

Requiere Node.js 18 o superior (probado con Node 24).

```bash
npm install
npm start
```

Abrir http://127.0.0.1:3000. Las pestañas se pueden abrir directo con la URL (`#A01`, `#A02`, `#A05`, `#A07`) y con las flechas del teclado.

## Pruebas

```bash
npm test
```

Comprueba que cada módulo vulnerable cae en su ataque y cada módulo seguro lo rechaza, incluidos los casos límite (firma de otro largo, token malformado, sesión falsa). Cierra antes el laboratorio si lo tienes corriendo.

## Puertos y variables

| Qué | Valor |
|---|---|
| Puerto del laboratorio | `3000`; se cambia con la variable `PORT` |
| Dirección de escucha | `127.0.0.1`, fija a propósito: el laboratorio no se expone a la red |
| Servicio interno de A01 | `4019`, fijo; simula un sistema que no debería ser alcanzable desde el formulario |

En PowerShell: `$env:PORT=3001; npm start`.

## Estructura

```
server.js            Carga los módulos y sirve la interfaz
modules/aXX.js       La lógica de cada taller (rutas /aXX/vuln/* y /aXX/safe/*)
public/index.html    Pestañas, estilos con tokens y el ayudante compartido `lab`
public/fragments/    Un panel por módulo (HTML y su script)
test/                Pruebas automáticas (node --test)
```

## Cómo agregar un módulo

1. Crear `modules/aXX.js` que exporte `meta` y `mount`:

   ```js
   exports.meta = { id: 'A06', title: 'Insecure Design: ...', short: '...', fragment: 'a06.html' };
   exports.mount = (app, ctx) => { /* rutas en /a06/vuln/* y /a06/safe/* */ };
   ```

2. Crear `public/fragments/a06.html` con el mismo esquema de los demás paneles: dos `section.panel` dentro de `.compare`, cada una con `data-kind="vuln"` o `"safe"` y su bloque `.out` de respuesta. `.compare` lleva `data-rows` con el número de filas de cada panel (4 o 5): así las respuestas de ambos lados quedan alineadas.
3. El script del panel recibe `(root, lab)`. `lab.run(boton, salida, url, opciones)` ejecuta la petición y pinta el código HTTP y el veredicto; marca con `data-attack="ataque"` los botones que lanzan el ataque.
4. Reiniciar el servidor: el módulo aparece solo como una pestaña nueva.

## Evidencia del informe

Con el laboratorio corriendo y en Windows con Microsoft Edge:

```bash
node ../docs/herramientas/capturar_evidencia.mjs
```

Recorre los cuatro talleres con clics reales, comprueba cada respuesta y cada veredicto, y regenera las capturas del informe en `../docs/evidencia`. Si una respuesta no es la esperada, falla.

## Problemas comunes

| Síntoma | Causa y solución |
|---|---|
| `EADDRINUSE` al arrancar | Hay otro laboratorio abierto. Ciérralo, o usa otro `PORT`. El puerto interno 4019 también debe estar libre. |
| Pantalla sin pestañas | El servidor no está corriendo: revisa `npm start` y recarga con Ctrl+F5. |
| Las columnas se ven apiladas | La ventana mide menos de 960 px de ancho: maximízala. |

## Uso responsable

Es vulnerable a propósito y solo sirve para aprender. Todas sus credenciales y claves (`admin / admin`, la clave de firma del token, la contraseña de demostración del operador) son de ejemplo. No lo publiques en Internet tal como está: el libro de comentarios de A05 es compartido y ejecuta el HTML que recibe.
