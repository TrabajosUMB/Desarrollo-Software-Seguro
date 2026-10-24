# -*- coding: utf-8 -*-
"""Quita de los resultados las rutas absolutas de este equipo (el repositorio es publico).

Uso: py limpiar_rutas.py <carpeta_resultados> <carpeta_Gosec>
Reemplaza, en cualquiera de sus escrituras (con \\, \\\\ o /), la ruta absoluta de la carpeta Gosec por "Gosec".
"""
import pathlib
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

# Gosec + (\ o \\ o /) + objetivos|ejemplo-propio + segmentos de ruta, hasta un espacio, ':' o comillas
PATRON_RUTA = re.compile(r'Gosec((?:\\{1,2}|/)(?:objetivos|ejemplo-propio)(?:(?:\\{1,2}|/)[^\\/"\s:\]\[]+)*)')

resultados = pathlib.Path(sys.argv[1])
raiz = pathlib.Path(sys.argv[2]).resolve()

barra_inversa = str(raiz)                       # C:\Users\...\Gosec
doble = barra_inversa.replace("\\", "\\\\")     # como aparece escapada en JSON
normal = barra_inversa.replace("\\", "/")       # C:/Users/.../Gosec
git_bash = "/" + normal[0].lower() + normal[2:]  # /c/Users/.../Gosec
url_archivo = "file:///" + normal               # SARIF

total = 0
for archivo in sorted(resultados.glob("*")):
    if archivo.suffix not in {".txt", ".json", ".sarif", ".html", ".md"}:
        continue
    texto = archivo.read_text(encoding="utf-8", errors="replace")
    nuevo = texto
    for viejo in (url_archivo, doble, barra_inversa, normal, git_bash):
        nuevo = nuevo.replace(viejo, "Gosec")
    # tras reemplazar quedan separadores mixtos (Gosec\objetivos\govwa\x.go, o con \\ en JSON): se unifican a /
    nuevo = PATRON_RUTA.sub(lambda m: "Gosec" + re.sub(r"\\{1,2}", "/", m.group(1)), nuevo)
    if nuevo != texto:
        archivo.write_text(nuevo, encoding="utf-8")
        total += 1
print(f"rutas locales limpiadas en {total} archivos")
restos = [a.name for a in resultados.glob("*") if a.is_file() and "Users" in a.read_text(encoding="utf-8", errors="replace") and "cuent" in a.read_text(encoding="utf-8", errors="replace")]
print("archivos que aun mencionan el usuario del equipo:", restos or "ninguno")
