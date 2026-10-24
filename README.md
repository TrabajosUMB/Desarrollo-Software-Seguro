# Desarrollo de Software Seguro

Trabajos de la asignatura Desarrollo de Software Seguro, Ingeniería de Software, Universidad Manuela Beltrán (periodo 2026-02).

![El mismo ataque contra la versión vulnerable y la segura del taller A07](Guia%206/Talleres%20Santiago/docs/evidencia/a07_comparativo.png)

El trabajo principal es un **laboratorio local de OWASP Top 10:2025**: cuatro talleres donde cada vulnerabilidad se puede atacar y, justo al lado, ver cómo la defensa la bloquea. Cada respuesta del servidor muestra su código HTTP y un veredicto ("El ataque funcionó" o "El ataque fue bloqueado").

## Probarlo en un minuto

Requiere [Node.js](https://nodejs.org) 18 o superior (probado con Node 24).

```bash
git clone https://github.com/TrabajosUMB/Desarrollo-Software-Seguro.git
cd "Desarrollo-Software-Seguro/Guia 6/Talleres Santiago/lab"
npm install
npm start
```

Abrir http://127.0.0.1:3000. Todo corre en el propio equipo, sin internet.

¿No quieres instalar nada? El informe con capturas reales de cada ataque está en [PDF](Guia%206/Talleres%20Santiago/docs/Guia6_Talleres_A01_A02_A05_A07.pdf).

## Los cuatro talleres

| Taller | Categoría OWASP 2025 | El ataque | La defensa |
|---|---|---|---|
| 1 | A01 Broken Access Control (CWE-918) | SSRF: el servidor visita cualquier URL y llega a un servicio interno. | Aceptar solo identificadores de una lista fija, nunca una URL. |
| 2 | A02 Security Misconfiguration (CWE-16) | Panel sin sesión, cuenta `admin / admin` y errores que revelan datos internos. | Sesión real, clave guardada como hash con sal y errores genéricos. |
| 3 | A05 Injection (CWE-79) | XSS almacenado: un comentario con HTML se ejecuta en el navegador. | Escapar la salida antes de mostrarla. |
| 4 | A07 Authentication Failures (CWE-287) | Un token con el rol alterado se acepta sin verificar la firma. | Recalcular y comparar la firma antes de confiar en el contenido. |

Cómo leer la pantalla: a la izquierda va la versión **vulnerable** y a la derecha la **segura**. Los botones rellenos lanzan el ataque; los de contorno son uso legítimo.

## Qué hay en el repositorio

| Carpeta | Contenido |
|---|---|
| [`Guia 7`](Guia%207/Gosec) | Análisis estático (SAST) con gosec sobre GoVWA y código propio: [README](Guia%207/Gosec/README.md), scripts, resultados reales y [presentación](Guia%207/Gosec/presentacion/Gosec_SAST_Guia7.pdf). |
| [`Guia 6`](Guia%206) | El laboratorio OWASP, su informe y la presentación de A06 Insecure Design. |
| [`Guia 4`](Guia%204) | Casos de abuso del proyecto Gasto Transparente (ODS 16): informe en [Word](Guia%204/CasosDeAbusoGastoTransparenteODS16.docx) y [PDF](Guia%204/CasosDeAbusoGastoTransparenteODS16.pdf), [presentación](Guia%204/Presentacion%20Casos%20de%20Abuso.pdf) y los scripts que los generan. |
| [`Guia 3`](Guia%203) | Cyber Kill Chain aplicado al sistema IRIS: [informe](Guia%203/Guia3_KillChain_SistemaIRIS.docx) y [presentación](Guia%203/Cyber_Kill_Chain.pptx). |
| [`LabPractico`](LabPractico) | Laboratorio práctico de casos de abuso y red team sobre Gasto Transparente, en tres partes más la fase de red team. |

### Dentro de `Guia 6/Talleres Santiago`

```
lab/                  El laboratorio (Node.js y Express) con sus pruebas
  modules/            Un módulo por taller: a01.js, a02.js, a05.js, a07.js
  public/             La interfaz: un panel por módulo en fragments/
  test/               Pruebas automáticas de los cuatro módulos
docs/
  Guia6_Talleres_A01_A02_A05_A07.pdf   Informe de los talleres (también en .docx)
  Guion_demo_en_vivo.md                Guion para presentar el laboratorio
  evidencia/                           Capturas reales de cada ataque
  herramientas/                        Scripts que regeneran la evidencia
```

Más detalle del laboratorio (puertos, cómo agregar un módulo, problemas comunes) en su [README](Guia%206/Talleres%20Santiago/lab/README.md). La presentación de A06 está en [`Guia 6/OWASP A06 Insecure Design`](Guia%206/OWASP%20A06%20Insecure%20Design).

## Comprobar que todo funciona

```bash
cd "Guia 6/Talleres Santiago/lab"
npm test
```

Verifica que cada módulo vulnerable cae en su ataque y cada módulo seguro lo rechaza. Cierra antes el laboratorio si lo tienes corriendo: el taller 1 usa el puerto interno 4019.

## Uso responsable

El laboratorio es vulnerable a propósito y solo sirve para aprender. Escucha únicamente en `127.0.0.1`, y todas sus credenciales y claves son de demostración. No lo publiques en Internet tal como está: el libro de comentarios del taller 3 es compartido y ejecuta el HTML que recibe.
