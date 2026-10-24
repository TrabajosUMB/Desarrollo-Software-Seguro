# Guía 7 · Análisis estático de seguridad (SAST) con gosec

Trabajo de **Desarrollo de Software Seguro** (Universidad Manuela Beltrán). Se instaló y configuró
[gosec](https://github.com/securego/gosec) v2.30.0, se analizó una aplicación vulnerable escrita en Go
y código propio, y se documentaron los hallazgos, su severidad, la mitigación y la experiencia de uso.

**Autores:** Cobos, Dilan Steven · Rodríguez Ángel, Santiago · Alarcón Ospina, Víctor Alfonso

> Todos los números de este documento salen de ejecuciones reales (carpeta `resultados/`).
> Lo que no se pudo medir se dice como tal.

## Qué es gosec

Escáner SAST para código Go: revisa el código fuente sin ejecutarlo. Analiza el árbol sintáctico (AST) y
la representación SSA, trae 62 reglas (G101 a G710, cada una con su CWE) y, desde la serie G70x, análisis de
flujo de datos que sigue una entrada del usuario hasta una función peligrosa. Código abierto (Apache 2.0).
Requiere Go 1.25 o superior. Su limitación principal: **solo analiza código Go**.

## Entorno y resultados de referencia

| Dato | Valor |
|---|---|
| Sistema | Windows 11, sin Docker, sin cuenta y sin administrador |
| Go | 1.27.2 |
| gosec | v2.30.0 (compilado desde el código fuente con `go install`, 1 min 50 s) |
| Objetivo vulnerable | [GoVWA](https://github.com/0c34/govwa), commit `4058f79f31eeb4a36d8f1e64bba1f0c899646e6f` |
| Tamaño analizado | 20 archivos, 1.569 líneas |
| Duración del escaneo | unos 5 segundos |

Un detalle práctico: al instalar desde el código fuente, `gosec -version` responde `dev`. La versión real se
lee con `go version -m gosec.exe`.

## Cómo reproducirlo

Requisitos: Go 1.25+, Git, Python (`py` en Windows) y un shell Bash (Git Bash).

```bash
# 1. instalar gosec
go install github.com/securego/gosec/v2/cmd/gosec@latest

# 2. escanear GoVWA (clona el commit fijado dentro de objetivos/, que no se sube al repositorio)
bash scripts/escanear.sh

# 3. escanear el código propio (versión vulnerable y versión corregida) y compararlas
bash scripts/escanear_ejemplo.sh

# 4. experimento con la regla G101 (credenciales escritas en el código)
bash scripts/experimento_g101.sh
```

Para leer los resultados:

```bash
py scripts/analizar.py                  # resumen por severidad, confianza, regla y archivo
py scripts/analizar.py G701             # además imprime el código marcado por esa regla
py scripts/analizar.py --csv            # escribe resultados/hallazgos.csv
```

`scripts/limpiar_rutas.py` se ejecuta al final de cada escaneo y reemplaza las rutas absolutas del equipo,
porque el repositorio es público.

## Hallazgos sobre GoVWA: 35 en total

| Severidad | Cantidad | Confianza | Cantidad |
|---|---|---|---|
| Alta | 1 | Alta | 28 |
| Media | 16 | Baja | 7 |
| Baja | 18 | | |

| Regla | CWE | Severidad / confianza | N | Qué es | Mitigación |
|---|---|---|---|---|---|
| G701 | CWE-89 | Alta / Alta | 1 | Inyección SQL (el dato del usuario llega a `DB.Query`) | Consulta parametrizada (`?`) y validar el tipo del identificador |
| G201 | CWE-89 | Media / Alta | 1 | SQL armado con `fmt.Sprintf` | La misma: el dato viaja aparte de la instrucción |
| G203 | CWE-79 | Media / Baja | 6 | `template.HTML` con texto del usuario (XSS) | Pasar texto plano a `html/template`; si hace falta HTML, sanitizar y añadir CSP |
| G401 | CWE-328 | Media / Alta | 3 | MD5 para contraseñas | bcrypt o argon2id, con sal |
| G501 | CWE-327 | Media / Alta | 3 | Importa `crypto/md5` | Igual que G401 |
| G124 | CWE-614 | Media / Alta | 2 | Cookie sin `Secure`, `HttpOnly` o `SameSite` | Activar las tres banderas |
| G112 | CWE-400 | Media / Baja | 1 | Servidor sin `ReadHeaderTimeout` (Slowloris) | `http.Server` con tiempos límite |
| G706 | CWE-117 | Baja / Alta | 11 | Inyección en logs | Registro estructurado (`log/slog`) y limpiar saltos de línea |
| G104 | CWE-703 | Baja / Alta | 7 | Error sin manejar | Comprobar `err` y responder de forma controlada |

La **confianza** indica qué tan seguro está gosec de que sea un problema real: lo de confianza baja
(7 hallazgos) se revisa a mano. En G203, por ejemplo, el hallazgo es real, pero gosec no podía saberlo.

Control: GoVWA trae junto a la consulta vulnerable una versión segura (`SafeQueryGetData`, con `DB.Prepare`);
gosec **no** la marcó.

## Antes y después con código propio

`ejemplo-propio/vulnerable/` reúne los errores típicos de Go y `ejemplo-propio/corregido/` los corrige.

| | Hallazgos |
|---|---|
| Código vulnerable | 17 (6 altos, 9 medios, 2 bajos; 15 reglas distintas) |
| Código corregido | 0 activos y 2 suprimidos con `#nosec` |

Los 2 suprimidos (G204 y G702) son un falso positivo: el comando solo se ejecuta tras validar el destino
con una lista de permitidos, pero gosec no entiende esa validación. Se suprimen con una justificación escrita
(`// #nosec G204 G702 -- ...`) y gosec lo registra (`-track-suppressions`), de modo que queda auditable.

## Lo que gosec no vio

| Caso | Qué pasó |
|---|---|
| IDOR (acceso a datos de otro usuario) | Es un fallo de lógica. En `idor.go` marcó el MD5 y errores sin manejar, pero no el acceso indebido |
| Secretos en `config.json` | La contraseña de la base de datos y la clave de sesión están en un JSON; gosec solo analiza `.go` |
| Credenciales (G101) | Depende del nombre de la variable y de la entropía del valor. De 4 variables con credenciales marcó 1; `password = "admin123"` aparece solo con `ignore_entropy`, y las variables llamadas `clave…` nunca (`resultados/experimento_g101.txt`) |
| Falsos positivos | Tras validar con lista de permitidos, G204 y G702 siguen avisando |

Conclusión: SAST se complementa con DAST, SCA y revisión manual.

## Experiencia de uso

| Criterio | Valoración |
|---|---|
| Instalación | Fácil: un comando; unos 2 minutos porque compila. Requiere Go; no necesita Docker |
| Facilidad de uso | Muy fácil: `gosec ./...` y listo |
| Resultados | Claros: regla, CWE, severidad, confianza, línea y código. Con falsos positivos y negativos |
| Interfaz | Terminal e informe HTML con filtros por severidad y confianza; también JSON, SARIF y JUnit para CI |

Recomendación: usarlo en CI con la Action oficial (`securego/gosec`), atender primero severidad alta con
confianza alta y complementarlo con DAST y revisión manual.

## Contenido de la carpeta

```
Gosec/
├── README.md
├── scripts/            escanear.sh, escanear_ejemplo.sh, experimento_g101.sh,
│                       analizar.py, comparar.py, limpiar_rutas.py,
│                       capturar_reporte.mjs, navegador.mjs, exportar_presentacion.ps1
├── ejemplo-propio/     vulnerable/ y corregido/ (módulo Go propio)
├── experimentos/g101/  experimento sobre la regla G101
├── resultados/         salidas reales: txt, json, sarif y html de GoVWA; ejemplo propio; entorno
├── presentacion/       Gosec_SAST_Guia7.pptx y .pdf (11 diapositivas, 10 minutos) y su constructor
└── objetivos/          GoVWA clonado por escanear.sh (código de terceros: no se redistribuye)
```

`presentacion/construir_presentacion.py` genera la presentación a partir de los JSON reales y comprueba con
`assert` que las cifras coinciden; las notas del orador de cada diapositiva llevan los tiempos.
`scripts/capturar_reporte.mjs` necesita Microsoft Edge e internet (el informe HTML carga sus estilos
desde un CDN).
