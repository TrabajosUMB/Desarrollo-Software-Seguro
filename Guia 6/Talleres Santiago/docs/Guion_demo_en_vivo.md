# Guion de la demo en vivo: 4 talleres OWASP Top 10:2025

Duración sugerida: 12 a 15 minutos (3 por taller + 1 de cierre). Todo corre en local, sin internet.

## Antes de entrar a clase (2 minutos)

1. Cerrar cualquier `node` viejo: `Get-Process node | Stop-Process -Force` (PowerShell). Si queda uno corriendo, el puerto 4019 (servicio interno de A01) aparece ocupado.
2. En la carpeta `lab/`: `npm install` (solo la primera vez) y `npm start`.
3. Abrir `http://127.0.0.1:3000` en Chrome o Edge. Pantalla completa (F11) para proyectar: a 1366 píxeles de ancho o más se ven las dos columnas, vulnerable a la izquierda y segura a la derecha.
4. Si el puerto 3000 está ocupado: `$env:PORT=3001; npm start` y abrir ese puerto.
5. Opcional, antes de `npm start`: `npm test` confirma en 2 segundos que los 4 módulos responden como se espera.
6. Las pestañas también se abren directo con la URL (`#A01`, `#A02`, `#A05`, `#A07`) y con las flechas del teclado.

Cada respuesta muestra primero un veredicto en una franja grande ("El ataque funcionó" en rojo, "El ataque fue bloqueado" en verde, "Uso legítimo permitido" en gris) y debajo su código HTTP (200 OK, 401 No autorizado, 403 Prohibido). Léanlo en voz alta: es lo que el público debe llevarse. Los botones rellenos lanzan el ataque; los de contorno son uso legítimo o utilidades.

## Orden y qué decir

### Taller 1: A01 Broken Access Control (SSRF), 3 min
- Decir: "El servidor visita la URL que yo le pida. Eso lo convierte en mi mensajero hacia la red interna."
- Pestaña A01, lado vulnerable: clic en **Pedir vista previa** (ya trae `http://127.0.0.1:4019/`).
- Se ve: `200 OK` y el JSON del servicio interno, con una clave de API simulada.
- Lado seguro: clic en **Probar con la URL interna**. Se ve `403 Prohibido` y "Identificador no permitido".
- Cierre: "La mitigación es una lista de permitidos, no una lista de bloqueo."

### Taller 2: A02 Security Misconfiguration, 3 min
- Decir: "No hay ningún bug de código aquí: solo cosas que nadie configuró."
- Pestaña A02, lado vulnerable: **Abrir panel (sin sesión)** muestra la configuración y las credenciales de fábrica.
- Clic en **Entrar con admin/admin**: entra como administrador.
- Clic en **Entrar con clave incorrecta**: el error revela tabla y archivo interno.
- Lado seguro, en este orden (arriba del panel se lee "Sesión: ninguna" y luego "Sesión: activa"):
  1. **Abrir panel (sin sesión)**: `401`, "Se requiere sesión válida".
  2. **Probar admin/admin**: `401`, "Credenciales inválidas", sin ningún dato interno.
  3. **Entrar con credencial válida**: `200` y un identificador de sesión.
  4. **Abrir panel (con sesión)**: ahora sí responde, sin datos sensibles.
- Si preguntan por la contraseña: se guarda como hash scrypt con sal, nunca en claro.

### Taller 3: A05 Injection (XSS almacenado), 3 min
- Decir: "Un comentario que el navegador toma como código."
- Pestaña A05. Primero, en cualquiera de los dos lados, **Publicar comentario** con el texto normal que trae el campo: se ve igual en ambos.
- Lado vulnerable: **Publicar etiqueta HTML**. Salta un cuadro con el número 1 (aceptar) y, en el libro, la etiqueta se convierte en el texto "XSS ejecutado". Veredicto rojo: "El ataque funcionó: el navegador ejecutó el HTML". Lo comprobé con Edge real: el cuadro se dispara.
- Lado seguro: **Publicar etiqueta HTML**. Se ve el texto literal de la etiqueta y no salta nada. Veredicto verde: "El ataque fue bloqueado".
- Al entrar a la pestaña el libro siempre arranca vacío, así que nunca salta un cuadro de ensayos anteriores. **Reiniciar comentarios** lo limpia en cualquier momento.
- Cierre: "Se escapa al mostrar, no solo al guardar."

### Taller 4: A07 Authentication Failures (JWT), 3 min
- Decir: "El token está firmado, pero el servidor no mira la firma."
- Pestaña A07, **Paso 1**: **Obtener token de "invitado"**. Aparece el token y, debajo de cada cuadro, su contenido decodificado con `"rol":"invitado"`.
- **Paso 2**: **Alterar rol a admin** (se habilita solo cuando ya hay token). El contenido decodificado pasa a `"rol":"admin"` y la firma queda igual.
- **Consultar perfil (vulnerable)**: `200 OK` y `"rol": "admin"`.
- **Consultar perfil (seguro)**: `401`, "Firma inválida".
- Cierre: "Firmar no sirve si no se verifica."
- Plan B si el botón del paso 2 no responde: en la consola (F12) se puede pegar el fragmento de abajo; deja el token alterado en los dos cuadros.

```js
const t=document.getElementById('a07-token').textContent;
const [h,,s]=t.split('.');
const p=JSON.parse(atob(t.split('.')[1].replace(/-/g,'+').replace(/_/g,'/')));p.rol='admin';
const b=btoa(JSON.stringify(p)).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
const f=h+'.'+b+'.'+s;
document.getElementById('a07-vuln-token').value=f;
document.getElementById('a07-safe-token').value=f;
```

## Si algo falla

| Síntoma | Solución |
|---|---|
| "Port 4019 already in use" o A01 sin respuesta | Un `node` viejo sigue vivo: `Get-Process node | Stop-Process -Force` y `npm start` otra vez |
| Pantalla en blanco o sin pestañas | Revisar que `npm start` siga corriendo y recargar con Ctrl+F5 |
| El comentario del taller 3 no se limpia | Botón **Reiniciar comentarios** |
| Los tokens del taller 4 dejaron de coincidir | Volver a pulsar **Obtener token** y repetir el paso 2 |
| Las dos columnas se ven apiladas | La ventana mide menos de 1200 píxeles: maximizarla o usar F11 |

## Respaldo

Si la demo no corre, todo está documentado con capturas reales en `docs/Guia6_Talleres_A01_A02_A05_A07.pdf`. Las capturas se regeneran con `node docs/herramientas/capturar_evidencia.mjs` (con el laboratorio corriendo), que además comprueba cada respuesta.
