# Desarrollo-Software-Seguro

Trabajos de la asignatura Desarrollo de Software Seguro. Universidad Manuela Beltrán, periodo 2026-02.

## Contenido

| Carpeta | Qué hay |
|---|---|
| `Guia 3` | Cyber Kill Chain aplicado al sistema IRIS: informe (`.docx`) y presentación (`.pptx`). |
| `Guia 4` | Casos de abuso del proyecto Gasto Transparente (ODS 16): informe en Word y PDF, presentación y los scripts que los generan. |
| `Guia 6` | Talleres de vulnerabilidades OWASP Top 10:2025 (ver abajo) y la presentación de A06 Insecure Design. |
| `LabPractico` | Laboratorio práctico de casos de abuso y red team sobre Gasto Transparente, en sus tres partes. |

## Guía 6: laboratorio OWASP Top 10:2025

Cuatro talleres con una versión vulnerable y una segura de cada flujo:

| Taller | Categoría | Qué demuestra |
|---|---|---|
| 1 | A01 Broken Access Control | SSRF: un servidor que visita cualquier URL llega a un servicio interno. |
| 2 | A02 Security Misconfiguration | Panel abierto, cuenta de fábrica y errores que revelan datos internos. |
| 3 | A05 Injection | XSS almacenado: un comentario que el navegador ejecuta como código. |
| 4 | A07 Authentication Failures | Un token firmado cuyo contenido se acepta sin verificar la firma. |

Dónde está cada cosa, dentro de `Guia 6/Talleres Santiago`:

- `lab/`: el laboratorio (Node.js y Express), con sus pruebas.
- `docs/`: el informe de los talleres en Word y PDF, la evidencia (capturas reales), el guion de la demo en vivo y las herramientas para regenerar las capturas.

### Cómo ejecutar el laboratorio

Requiere Node.js 18 o superior.

```bash
cd "Guia 6/Talleres Santiago/lab"
npm install
npm start
```

Abrir http://127.0.0.1:3000. Para comprobar que cada módulo responde como se espera: `npm test`.

### Advertencia

El laboratorio es vulnerable a propósito. Solo escucha en `127.0.0.1` y todas sus credenciales y claves son de demostración. No debe publicarse en Internet tal como está: el libro de comentarios de A05 es compartido y ejecuta el HTML que recibe.
