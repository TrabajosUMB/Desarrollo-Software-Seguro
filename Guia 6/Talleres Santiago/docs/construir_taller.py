# -*- coding: utf-8 -*-
"""Construye el Word de los 4 talleres (A01, A02, A05, A07) con estilos UMB reales."""
import sys
sys.path.insert(0, r"C:\Users\cuent\Documents\Universidad\Katherin\2026-02\Sistemas Transaccionales\herramientas_doc\build")
sys.path.insert(0, r"C:\Users\cuent\.claude\skills\umb-documento\references")
sys.stdout.reconfigure(encoding="utf-8")

import pathlib
from constructor import Documento
from corregir_docx import crear_estilo_codigo, cabecera_azul, agregar_referencia

EV = pathlib.Path(__file__).parent / "evidencia"
TITULO = "Guía 6: Laboratorio de vulnerabilidades OWASP Top 10:2025 (A01, A02, A05, A07)"
AUTORES = ("Cobos, Dilan Steven", "Rodríguez Ángel, Santiago", "Alarcón Ospina, Víctor Alfonso")


def portada(doc):
    doc.fijar_propiedades(TITULO, AUTORES)
    doc.portada_logo(3.0)
    doc.portada_linea()
    doc.portada_linea(TITULO, negrita=True)
    doc.portada_linea()
    for linea in (*AUTORES, "Universidad Manuela Beltrán", "Ingeniería de Software", "Facultad de Ingeniería",
                  "Desarrollo de Software Seguro", "Bogotá, Colombia", "2026"):
        doc.portada_linea(linea)
    doc.salto_pagina()


def listas(doc):
    doc.rotulo_centrado("Tabla de contenido")
    doc.campo_toc('TOC \\o "1-3" \\h \\z \\u')
    doc.rotulo_centrado("Lista de tablas", salto_antes=True)
    doc.campo_toc('TOC \\h \\z \\t "TITULO TABLA,1"')
    doc.rotulo_centrado("Lista de figuras", salto_antes=True)
    doc.campo_toc('TOC \\h \\z \\t "TITULO FIGURA,1"')
    doc.salto_pagina()


def introduccion(doc):
    doc.titulo(1, "Introducción", salto_antes=True)
    doc.parrafo(
        "Este documento corresponde a la Guía 6 de Desarrollo de Software Seguro. El grupo seleccionó cuatro "
        "vulnerabilidades del OWASP Top 10:2025 que no se repiten con los talleres ya entregados por otros grupos "
        "del curso: A01:2025 Broken Access Control, con enfoque en Server-Side Request Forgery (SSRF); "
        "A02:2025 Security Misconfiguration; A05:2025 Injection, con enfoque en Cross-Site Scripting (XSS) "
        "almacenado; y A07:2025 Authentication Failures, con enfoque en un token de sesión (JWT) cuya firma no "
        "se verifica.")
    doc.parrafo(
        "Para los cuatro talleres se construyó un laboratorio propio en Node.js y Express que corre únicamente "
        "en el equipo local (127.0.0.1), sin depender de un servicio externo. Cada módulo expone dos versiones "
        "del mismo flujo: una <b>vulnerable</b>, que reproduce el defecto descrito por OWASP, y una <b>segura</b>, "
        "que aplica la mitigación correspondiente sobre el mismo flujo. Esto permite comparar en vivo, ante el "
        "docente, el antes y el después de cada corrección.")
    doc.parrafo(
        "Todas las capturas de este documento corresponden a ejecuciones reales del laboratorio, realizadas el "
        "17 de octubre de 2026 en el navegador Microsoft Edge, con clics reales sobre cada botón y contra el "
        "servidor local. Cada respuesta muestra su código HTTP y un veredicto: «El ataque funcionó» en la versión "
        "vulnerable o «El ataque fue bloqueado» en la segura. El código fuente completo del "
        "laboratorio (carpeta <b>lab/</b>) se entrega junto con este documento.")


# --------------------------------------------------------------------------- A01
def taller_a01(doc):
    doc.titulo(1, "Taller 1. A01:2025 Broken Access Control (SSRF)", salto_antes=True)
    doc.titulo(2, "Nombre de la actividad")
    doc.parrafo("La imagen que pidió el servidor: explotando un Server-Side Request Forgery.")
    doc.titulo(2, "Vulnerabilidad asociada")
    doc.parrafo(
        "A01:2025 Broken Access Control. OWASP mantiene esta categoría en el puesto #1 del ranking: el 100 % de "
        "las aplicaciones probadas presentó alguna forma de control de acceso roto. Entre los CWE señalados "
        "explícitamente está <b>CWE-918 Server-Side Request Forgery (SSRF)</b>, que es el enfoque de este taller "
        "(OWASP Top 10:2025, A01).")
    doc.titulo(2, "Descripción del entorno")
    doc.parrafo(
        "Aplicación Node.js/Express (módulo <b>modules/a01.js</b> del laboratorio) con una función de \"vista "
        "previa de imagen\": el usuario escribe una URL y el propio servidor hace esa petición por detrás, para "
        "luego mostrar la respuesta. Junto al servidor principal corre un segundo servicio HTTP en el puerto "
        "4019, enlazado solo a 127.0.0.1, que simula un sistema interno (un \"inventario\") que nunca debería ser "
        "alcanzable desde un formulario público. Como medida de contención del propio laboratorio (y no como parte "
        "de la vulnerabilidad), el módulo solo permite peticiones http hacia 127.0.0.1 o localhost, de modo que "
        "el formulario no pueda usarse contra equipos ajenos.")
    doc.titulo(2, "Objetivo del reto")
    doc.parrafo(
        "Demostrar que, cuando el servidor no valida qué URL puede pedir, un usuario puede hacer que ese servidor "
        "contacte un servicio interno que no está expuesto al público, y leer su respuesta. Luego comparar con una "
        "versión que solo acepta identificadores de una lista fija (allowlist), nunca una URL libre.")
    doc.titulo(2, "Instrucciones")
    for i, paso in enumerate((
        "Ejecutar <b>npm install && npm start</b> dentro de la carpeta <b>lab/</b> y abrir "
        "http://localhost:3000 en el navegador.",
        "Abrir la pestaña <b>A01</b>. En el campo de la versión vulnerable ya aparece precargada la URL "
        "<b>http://127.0.0.1:4019/</b>, que es la del servicio interno.",
        "Hacer clic en <b>Pedir vista previa</b>. Observar que el servidor responde 200 OK con el contenido del "
        "servicio interno (incluida una \"clave de API\" simulada), aunque nadie debería poder llegar hasta allí "
        "desde este formulario.",
        "En la versión segura, hacer clic en <b>Probar con la URL interna</b>, que escribe esa misma URL en el campo "
        "de identificador y la envía. Verificar que la respuesta es 403 Prohibido: la versión segura solo acepta "
        "un valor de una lista fija (por ejemplo, <b>logo</b>), nunca una URL arbitraria.",
    ), start=1):
        doc.parrafo(f"Paso {i}. {paso}")
    doc.figura("a01", "Explotación de SSRF: la vista previa vulnerable revela el servicio interno",
               EV / "a01_vuln_ssrf.png")
    doc.figura("a01b", "Comparación: la versión vulnerable responde con los datos internos; la segura rechaza la URL",
               EV / "a01_comparativo.png")
    doc.titulo(2, "Resultados esperados")
    doc.parrafo(
        "La versión vulnerable devuelve el JSON del servicio interno (incluida la nota «Este endpoint nunca "
        "debería ser alcanzable desde fuera del servidor»). La versión segura devuelve el error «Identificador no "
        "permitido» para la misma URL.")
    doc.titulo(2, "Criterios de éxito")
    doc.parrafo(
        "El estudiante demuestra que el servidor vulnerable contacta un servicio interno a partir de una URL "
        "controlada por el usuario, y que la versión segura rechaza esa misma URL por no estar en la lista "
        "permitida.")
    doc.titulo(2, "Medidas de mitigación")
    doc.parrafo("Código de la versión segura (módulo <b>modules/a01.js</b>):")
    p = doc.d.add_paragraph(style="Codigo")
    p.add_run(
        "app.get('/a01/safe/preview', (req, res) => {\n"
        "  const id = req.query.id;\n"
        "  if (id !== 'logo') {\n"
        "    return res.status(403).json({ error: 'Identificador no permitido...' });\n"
        "  }\n"
        "  res.json({ solicitado: id, nota: 'Nunca se abre una URL que venga del usuario.' });\n"
        "});")
    doc.parrafo(
        "Nunca aceptar una URL completa escrita por el usuario para que el servidor la visite. Validar contra "
        "una lista fija de destinos permitidos (allowlist), nunca una lista de bloqueo. Cuando se necesite acceder "
        "a un recurso remoto, hacerlo por un identificador interno que el servidor resuelve, no por la URL "
        "literal. Segmentar la red para que los servicios internos no sean alcanzables ni siquiera si el filtro "
        "falla.")
    doc.titulo(2, "Preguntas de análisis")
    doc.parrafo(
        "1. ¿Por qué una lista de bloqueo (bloquear \"127.0.0.1\", \"localhost\", etc.) es más fácil de evadir que "
        "una lista de permitidos? 2. Si el servicio interno del taller fuera, en una empresa real, el panel de "
        "administración de la nube (por ejemplo, el servicio de metadatos de un proveedor cloud), ¿qué información "
        "sensible podría exponer un SSRF como este?")


# --------------------------------------------------------------------------- A02
def taller_a02(doc):
    doc.titulo(1, "Taller 2. A02:2025 Security Misconfiguration", salto_antes=True)
    doc.titulo(2, "Nombre de la actividad")
    doc.parrafo("admin / admin: el panel que nadie terminó de configurar.")
    doc.titulo(2, "Vulnerabilidad asociada")
    doc.parrafo(
        "A02:2025 Security Misconfiguration. Subió del puesto #5 al #2; el 100 % de las aplicaciones probadas "
        "presentó alguna forma de mala configuración. OWASP describe explícitamente dos síntomas que este taller "
        "reproduce: <b>cuentas por defecto con su contraseña sin cambiar</b> y <b>mensajes de error que revelan "
        "trazas de pila o detalles internos</b> (OWASP Top 10:2025, A02).")
    doc.titulo(2, "Descripción del entorno")
    doc.parrafo(
        "Módulo <b>modules/a02.js</b> del laboratorio: un panel de administración simulado. La versión vulnerable "
        "no exige sesión para leer la configuración del panel, conserva credenciales de fábrica (<b>admin / "
        "admin</b>) y, si el inicio de sesión falla, devuelve un error con una traza interna simulada. La versión "
        "segura no tiene credenciales de fábrica: guarda la contraseña como un hash scrypt con sal, emite un "
        "identificador de sesión aleatorio al iniciar sesión, exige ese identificador para leer el panel y "
        "responde con un mensaje genérico cuando las credenciales fallan.")
    doc.titulo(2, "Objetivo del reto")
    doc.parrafo(
        "Demostrar que una ruta administrativa sin control de sesión, combinada con credenciales de fábrica nunca "
        "cambiadas, permite leer la configuración interna de la aplicación y entrar como administrador sin "
        "explotar ningún error de programación, solo una mala configuración.")
    doc.titulo(2, "Instrucciones")
    for i, paso in enumerate((
        "Con el laboratorio corriendo, abrir la pestaña <b>A02</b>.",
        "En la versión vulnerable, hacer clic en <b>Abrir panel (sin sesión)</b>. Observar que el panel responde "
        "sin pedir ninguna credencial, y expone la configuración interna, incluidas las credenciales de fábrica.",
        "Hacer clic en <b>Entrar con admin/admin</b>. Verificar que el inicio de sesión es exitoso y el rol "
        "devuelto es <b>administrador</b>.",
        "Hacer clic en <b>Entrar con clave incorrecta</b> y observar que el error revela el nombre de una tabla y "
        "una traza de archivo y línea, información que nunca debería llegar al usuario.",
        "En la versión segura, hacer clic en <b>Abrir panel (sin sesión)</b> y en <b>Probar admin/admin</b>. "
        "Verificar que el panel exige una sesión y que las credenciales de fábrica son rechazadas con un "
        "mensaje genérico, sin ningún dato interno.",
        "Hacer clic en <b>Entrar con credencial válida</b>: el servidor devuelve un identificador de sesión. "
        "Luego hacer clic en <b>Abrir panel (con sesión)</b> y verificar que ahora el panel responde, sin "
        "ningún dato sensible.",
    ), start=1):
        doc.parrafo(f"Paso {i}. {paso}")
    doc.figura("a02", "Versión vulnerable: el panel abre sin sesión y admin/admin entra como administrador",
               EV / "a02_vuln_login.png")
    doc.figura("a02b", "Versión segura: el panel exige sesión y admin/admin es rechazado",
               EV / "a02_seguro_rechaza.png")
    doc.figura("a02c", "Versión segura: con la credencial válida se obtiene una sesión y el panel responde",
               EV / "a02_seguro_con_sesion.png")
    doc.titulo(2, "Resultados esperados")
    doc.parrafo(
        "La versión vulnerable entrega el panel sin sesión y acepta admin/admin con rol administrador. La versión "
        "segura devuelve «Se requiere sesión válida» para el panel sin sesión y «Credenciales inválidas» para "
        "admin/admin, sin ningún dato interno; con la credencial válida entrega una sesión y el panel responde.")
    doc.titulo(2, "Criterios de éxito")
    doc.parrafo(
        "El estudiante demuestra el acceso al panel sin sesión, el inicio de sesión con credenciales de fábrica, y "
        "la traza interna expuesta en el error; y muestra que la versión segura bloquea las tres cosas.")
    doc.titulo(2, "Medidas de mitigación")
    doc.parrafo("Código de la versión segura (módulo <b>modules/a02.js</b>):")
    p = doc.d.add_paragraph(style="Codigo")
    p.add_run(
        "const STORED_HASH = crypto.scryptSync('unaClaveLarga123', SALT, 32); // nunca en claro\n"
        "app.post('/a02/safe/login', (req, res) => {\n"
        "  const { user, pass } = req.body || {};\n"
        "  if (user === SAFE_USER && checkPassword(pass)) {\n"
        "    const sesion = crypto.randomBytes(16).toString('hex');\n"
        "    sessions.add(sesion);\n"
        "    return res.json({ ok: true, rol: 'operador', sesion });\n"
        "  }\n"
        "  res.status(401).json({ ok: false, error: 'Credenciales inválidas.' });\n"
        "});\n"
        "app.get('/a02/safe/admin-panel', (req, res) => {\n"
        "  if (!sessions.has(req.header('x-session'))) return res.status(401).json({ error: 'Se requiere sesión.' });\n"
        "  res.json({ panel: 'admin', nota: 'Sin datos sensibles.' });\n"
        "});")
    doc.parrafo(
        "Eliminar toda credencial de fábrica antes de desplegar, en especial para cuentas de administrador. "
        "Proteger cada ruta administrativa con verificación de sesión y de rol, no solo con su nombre poco "
        "conocido. Configurar manejo centralizado de errores que devuelva mensajes genéricos al cliente y registre "
        "el detalle real solo en el servidor. Usar el mismo proceso de endurecimiento (hardening) y las mismas "
        "credenciales por entorno distintas para desarrollo, pruebas y producción.")
    doc.titulo(2, "Preguntas de análisis")
    doc.parrafo(
        "1. ¿Por qué \"ocultar\" la ruta del panel (no enlazarla desde el menú) no es una mitigación real? "
        "2. ¿Qué información de la traza de error vulnerable usaría un atacante para planear su siguiente paso?")


# --------------------------------------------------------------------------- A05
def taller_a05(doc):
    doc.titulo(1, "Taller 3. A05:2025 Injection (XSS almacenado)", salto_antes=True)
    doc.titulo(2, "Nombre de la actividad")
    doc.parrafo("Un comentario que no era solo texto: XSS almacenado en un libro de comentarios.")
    doc.titulo(2, "Vulnerabilidad asociada")
    doc.parrafo(
        "A05:2025 Injection. El 100 % de las aplicaciones probadas se evaluó para alguna forma de inyección, y "
        "esta categoría tiene el mayor número de CVE de todo el Top 10:2025. OWASP identifica dentro de ella a "
        "<b>CWE-79 Improper Neutralization of Input During Web Page Generation (Cross-Site Scripting)</b>, con "
        "más de 30 000 CVE asociados, que es el enfoque de este taller (OWASP Top 10:2025, A05).")
    doc.titulo(2, "Descripción del entorno")
    doc.parrafo(
        "Módulo <b>modules/a05.js</b> del laboratorio: un libro de comentarios muy simple. La versión vulnerable "
        "guarda el comentario y lo devuelve tal cual; la interfaz lo inserta en la página con "
        "<b>innerHTML</b>, por lo que cualquier etiqueta HTML dentro del comentario se ejecuta en el navegador de "
        "quien lo lee. La versión segura escapa los caracteres <b>< > & \" '</b> antes de devolver el "
        "comentario, así que el navegador lo muestra como texto, nunca como código.")
    doc.titulo(2, "Objetivo del reto")
    doc.parrafo(
        "Demostrar que un comentario con una etiqueta <b><img></b> con un manejador <b>onerror</b> se "
        "ejecuta como HTML real en la versión vulnerable (el navegador intenta cargar la imagen y dispara el "
        "evento), y que la misma entrada queda como texto inofensivo en la versión segura.")
    doc.titulo(2, "Instrucciones")
    for i, paso in enumerate((
        "Con el laboratorio corriendo, abrir la pestaña <b>A05</b>.",
        "En la versión vulnerable, hacer clic en <b>Publicar etiqueta HTML</b>. El laboratorio publica un comentario "
        "con la etiqueta <b><img src=x onerror=...></b>, que lleva un manejador <b>onerror</b> con un "
        "<b>alert(1)</b> y una marca visible.",
        "Observar que el navegador muestra un cuadro de alerta con el número 1 y que, en el libro de comentarios, "
        "la etiqueta fue reemplazada por el texto <b>XSS ejecutado</b>: interpretó el comentario como una etiqueta "
        "<b><img></b> real y ejecutó su manejador de error. El veredicto es «El ataque funcionó».",
        "En la versión segura, hacer clic en <b>Publicar etiqueta HTML</b>. Observar que esta vez se ve el texto "
        "exacto de la etiqueta, como texto plano, sin alerta, y que el veredicto es «El ataque fue bloqueado».",
    ), start=1):
        doc.parrafo(f"Paso {i}. {paso}")
    doc.figura("a05", "Comparación: la etiqueta se ejecuta como HTML en la versión vulnerable (aparece XSS "
               "ejecutado) y queda como texto en la segura", EV / "a05_comparativo.png")
    doc.titulo(2, "Resultados esperados")
    doc.parrafo(
        "La versión vulnerable devuelve el comentario sin escapar y el navegador lo renderiza como HTML (salta el "
        "cuadro de alerta y aparece el texto XSS ejecutado, prueba de que la etiqueta <b><img></b> se creó de "
        "verdad). La versión "
        "segura devuelve el comentario con las entidades <b>&lt; &gt; &quot;</b>, que el navegador "
        "muestra como el texto literal de la etiqueta.")
    doc.titulo(2, "Criterios de éxito")
    doc.parrafo(
        "El estudiante logra que el navegador interprete una etiqueta HTML inyectada en la versión vulnerable, y "
        "muestra que la versión segura la neutraliza, mostrándola solo como texto.")
    doc.titulo(2, "Medidas de mitigación")
    doc.parrafo("Código de la versión segura (módulo <b>modules/a05.js</b>):")
    p = doc.d.add_paragraph(style="Codigo")
    p.add_run(
        "function escapeHtml(s) {\n"
        "  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')\n"
        "    .replace(/>/g, '&gt;').replace(/\"/g, '&quot;').replace(/'/g, '&#39;');\n"
        "}\n"
        "app.get('/a05/safe/comments', (_req, res) =>\n"
        "  res.json({ comments: safeComments.map(escapeHtml) }));")
    doc.parrafo(
        "Escapar siempre la salida que se inserta en HTML, en el punto donde se renderiza, no solo al guardar el "
        "dato. Evitar <b>innerHTML</b> con datos de usuario; preferir <b>textContent</b> o un motor de plantillas "
        "que escape por defecto (la mayoría de los frameworks modernos lo hacen). Añadir una Content Security "
        "Policy (CSP) como segunda capa de defensa frente a scripts inyectados.")
    doc.titulo(2, "Preguntas de análisis")
    doc.parrafo(
        "1. ¿Por qué escapar la entrada al guardarla no basta, y hay que escapar también al mostrarla? "
        "2. Si el comentario vulnerable lo pudiera leer otro usuario (no solo quien lo escribió), ¿qué podría "
        "hacer un script malicioso en su sesión (robar la cookie, actuar en su nombre)?")


# --------------------------------------------------------------------------- A07
def taller_a07(doc):
    doc.titulo(1, "Taller 4. A07:2025 Authentication Failures (JWT sin verificar)", salto_antes=True)
    doc.titulo(2, "Nombre de la actividad")
    doc.parrafo("De invitado a administrador: falsificando un token sin tocar la clave secreta.")
    doc.titulo(2, "Vulnerabilidad asociada")
    doc.parrafo(
        "A07:2025 Authentication Failures. OWASP describe aquí los fallos en la verificación de la identidad del "
        "usuario, incluyendo tokens de sesión que no se validan correctamente. El defecto de este taller "
        "corresponde a <b>CWE-287 Improper Authentication</b>: el servidor acepta como válida una afirmación de "
        "identidad (el contenido del token) sin comprobar su firma (OWASP Top 10:2025, A07).")
    doc.titulo(2, "Descripción del entorno")
    doc.parrafo(
        "Módulo <b>modules/a07.js</b> del laboratorio: emite tokens con el mismo formato que un JWT "
        "(encabezado.carga.firma, codificados en Base64url y firmados con HMAC-SHA256). La ruta de perfil "
        "vulnerable decodifica la carga del token y confía en su contenido sin llamar nunca a la función que "
        "verifica la firma. La ruta de perfil segura recalcula la firma y compara con "
        "<b>crypto.timingSafeEqual</b> antes de confiar en el contenido.")
    doc.titulo(2, "Objetivo del reto")
    doc.parrafo(
        "Obtener un token válido como usuario \"invitado\", modificar su carga para declararse \"admin\" sin "
        "volver a firmarlo (conservando la firma original, ahora inválida para el nuevo contenido), y comprobar "
        "que la ruta vulnerable acepta ese cambio mientras la ruta segura lo rechaza.")
    doc.titulo(2, "Instrucciones")
    for i, paso in enumerate((
        "Con el laboratorio corriendo, abrir la pestaña <b>A07</b> y hacer clic en <b>Obtener token de "
        "\"invitado\"</b>. El token aparece dividido en tres partes separadas por puntos y se copia automáticamente "
        "en los dos cuadros de texto de abajo.",
        "Hacer clic en <b>Alterar rol a admin</b>. El laboratorio decodifica en Base64url la parte central del token "
        "(la carga), cambia <b>\"rol\":\"invitado\"</b> por <b>\"rol\":\"admin\"</b>, la vuelve a codificar y deja el "
        "encabezado y la firma exactamente como estaban. Bajo cada cuadro se muestra el contenido ya decodificado.",
        "Con el token alterado ya cargado en el cuadro de la versión vulnerable, hacer clic en <b>Consultar perfil "
        "(vulnerable)</b>. Observar que la respuesta es 200 OK e incluye <b>\"rol\": \"admin\"</b>, aunque el token "
        "real nunca fue firmado con ese contenido.",
        "Con el mismo token en el cuadro de la versión segura, hacer clic en <b>Consultar perfil (seguro)</b>. "
        "Observar que la respuesta es 401 No autorizado, con un error de firma inválida.",
    ), start=1):
        doc.parrafo(f"Paso {i}. {paso}")
    doc.figura("a07", "Preparación del ataque y comparación: el perfil vulnerable acepta el rol admin falsificado; "
               "el seguro lo rechaza por firma inválida", EV / "a07_comparativo.png")
    doc.titulo(2, "Resultados esperados")
    doc.parrafo(
        "La ruta vulnerable devuelve <b>{\"ok\": true, \"payload\": {\"rol\": \"admin\", ...}}</b> para el token "
        "alterado. La ruta segura devuelve <b>{\"error\": \"Firma inválida: el token fue alterado o no "
        "corresponde a este servidor.\"}</b> para ese mismo token.")
    doc.titulo(2, "Criterios de éxito")
    doc.parrafo(
        "El estudiante logra que la ruta vulnerable confíe en un rol que el servidor nunca firmó, y demuestra que "
        "la ruta segura detecta la alteración y la rechaza.")
    doc.titulo(2, "Medidas de mitigación")
    doc.parrafo("Código de la versión segura (módulo <b>modules/a07.js</b>):")
    p = doc.d.add_paragraph(style="Codigo")
    p.add_run(
        "function verify(token) {\n"
        "  try {\n"
        "    const [h, b, s] = String(token).split('.');\n"
        "    const esperada = crypto.createHmac('sha256', SECRET).update(h + '.' + b).digest();\n"
        "    const dada = Buffer.from(s, 'base64url');\n"
        "    if (dada.length !== esperada.length || !crypto.timingSafeEqual(esperada, dada)) return null;\n"
        "    return JSON.parse(Buffer.from(b, 'base64url').toString('utf8'));\n"
        "  } catch { return null; }\n"
        "}\n"
        "app.get('/a07/safe/perfil', (req, res) => {\n"
        "  const payload = verify(token);\n"
        "  if (!payload) return res.status(401).json({ error: 'Firma invalida.' });\n"
        "  res.json({ ok: true, payload });\n"
        "});")
    doc.parrafo(
        "Verificar siempre la firma del token con una librería de JWT madura antes de leer su contenido; nunca "
        "decodificar y confiar en la carga sin ese paso. Usar <b>timingSafeEqual</b> (comparación en tiempo "
        "constante) para comparar firmas, nunca <b>===</b>, y evitar que el servidor acepte el algoritmo "
        "\"none\". Mantener la clave de firma fuera del código fuente, en una variable de entorno o un gestor de "
        "secretos, y usar tokens de vida corta.")
    doc.titulo(2, "Preguntas de análisis")
    doc.parrafo(
        "1. ¿Por qué firmar el token no sirve de nada si el servidor nunca comprueba esa firma al leerlo? "
        "2. Si un atacante no conoce la clave secreta del servidor, ¿por qué de todos modos logra que el servidor "
        "acepte un rol distinto al que firmó originalmente?")


def cuadro_resumen(doc):
    doc.titulo(1, "Resumen de los cuatro talleres")
    doc.tabla(
        "resumen", "Vulnerabilidades trabajadas en la Guía 6",
        ["Taller", "Categoría OWASP Top 10:2025", "CWE principal", "Módulo del laboratorio"],
        [
            ["1", "A01 Broken Access Control (SSRF)", "CWE-918", "modules/a01.js"],
            ["2", "A02 Security Misconfiguration", "CWE-16", "modules/a02.js"],
            ["3", "A05 Injection (XSS almacenado)", "CWE-79", "modules/a05.js"],
            ["4", "A07 Authentication Failures (JWT)", "CWE-287", "modules/a07.js"],
        ],
        anchos=[0.7, 3.2, 1.3, 2.0], alinear=["c", "l", "c", "l"], tam=10,
    )
    cabecera_azul(doc.d.tables[-1])


def anexo(doc):
    doc.titulo(1, "Anexo: cómo ejecutar el laboratorio")
    doc.parrafo(
        "Requisitos: Node.js 18 o superior. Dentro de la carpeta <b>lab/</b> de este entregable, ejecutar:")
    p = doc.d.add_paragraph(style="Codigo")
    p.add_run("npm install\nnpm start")
    doc.parrafo(
        "El servidor queda disponible en <b>http://127.0.0.1:3000</b> (o el puerto que indique la variable de "
        "entorno <b>PORT</b>). La interfaz tiene una pestaña por cada módulo (A01, A02, A05, A07); cada una "
        "incluye un bloque <b>vulnerable</b> y uno <b>segura</b> para repetir en vivo lo mostrado en este "
        "documento.")


def referencias(doc):
    doc.titulo(1, "Referencias")
    for partes in (
        [("OWASP. (2025a). ", False), ("A01:2025 Broken Access Control. ", True),
         ("OWASP Top 10:2025. https://top10.owasp.org/2025/A01_2025-Broken_Access_Control/", False)],
        [("OWASP. (2025b). ", False), ("A02:2025 Security Misconfiguration. ", True),
         ("OWASP Top 10:2025. https://top10.owasp.org/2025/A02_2025-Security_Misconfiguration/", False)],
        [("OWASP. (2025c). ", False), ("A05:2025 Injection. ", True),
         ("OWASP Top 10:2025. https://top10.owasp.org/2025/A05_2025-Injection/", False)],
        [("OWASP. (2025d). ", False), ("A07:2025 Authentication Failures. ", True),
         ("OWASP Top 10:2025. https://top10.owasp.org/2025/A07_2025-Authentication_Failures/", False)],
    ):
        agregar_referencia(doc.d, partes)


def construir(doc):
    portada(doc)
    listas(doc)
    introduccion(doc)
    taller_a01(doc)
    taller_a02(doc)
    taller_a05(doc)
    taller_a07(doc)
    cuadro_resumen(doc)
    anexo(doc)
    referencias(doc)


if __name__ == "__main__":
    ruta = pathlib.Path(__file__).parent / "Guia6_Talleres_A01_A02_A05_A07.docx"
    previo = Documento(estricto=False)
    crear_estilo_codigo(previo.d)
    construir(previo)
    doc = Documento(etiquetas=previo.etiquetas)
    crear_estilo_codigo(doc.d)
    construir(doc)
    doc.guardar(ruta)
    print("guardado:", ruta)
