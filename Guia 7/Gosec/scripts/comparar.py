# -*- coding: utf-8 -*-
"""Compara el escaneo de ejemplo-propio/vulnerable contra ejemplo-propio/corregido.

Lee resultados/ejemplo_vulnerable.json y resultados/ejemplo_corregido.json (los genera escanear_ejemplo.sh).
"""
import collections
import json
import pathlib
import sys

sys.stdout.reconfigure(encoding="utf-8")

RESULTADOS = pathlib.Path(__file__).resolve().parent.parent / "resultados"


def cargar(nombre):
    datos = json.loads((RESULTADOS / nombre).read_text(encoding="utf-8"))
    return datos["Issues"], datos["Stats"]


def corto(i):
    return f"{i['rule_id']} {i['severity'][0]}/{i['confidence'][0]}"


vulnerables, stats_v = cargar("ejemplo_vulnerable.json")
todos_c, stats_c = cargar("ejemplo_corregido.json")
# Con -track-suppressions gosec lista tambien lo suprimido con #nosec, marcado en "suppressions"
corregidos = [i for i in todos_c if not i.get("suppressions")]
suprimidos = [i for i in todos_c if i.get("suppressions")]

print(f"VULNERABLE: {stats_v['lines']} lineas, {len(vulnerables)} hallazgos")
print(f"CORREGIDO : {stats_c['lines']} lineas, {len(corregidos)} hallazgos activos, {len(suprimidos)} suprimido(s) con #nosec")
print()
print("severidad en vulnerable:", dict(collections.Counter(i["severity"] for i in vulnerables)))
print("severidad en corregido :", dict(collections.Counter(i["severity"] for i in corregidos)) or "sin hallazgos activos")
print()
print(f"{'regla':<6} {'sever.':<7} {'conf.':<6} {'CWE':<8} {'linea':<6} descripcion")
for i in sorted(vulnerables, key=lambda x: (x["rule_id"], int(x["line"].split("-")[0]))):
    print(f"{i['rule_id']:<6} {i['severity']:<7} {i['confidence']:<6} CWE-{i['cwe']['id']:<4} {i['line']:<6} {i['details'][:70]}")
print()
print("Hallazgos activos que siguen en la version corregida:")
if not corregidos:
    print("  ninguno")
for i in corregidos:
    print(f"  {corto(i)}  linea {i['line']}  {i['details'][:80]}")
print("Suprimidos con #nosec (deben llevar justificacion):")
for i in suprimidos:
    print(f"  {corto(i)}  linea {i['line']}  -- {i['suppressions'][0].get('justification', '(sin justificacion)')}")
print()
reglas_v = {i["rule_id"] for i in vulnerables}
reglas_c = {i["rule_id"] for i in corregidos}
print("reglas distintas detectadas antes :", len(reglas_v), sorted(reglas_v))
print("reglas que desaparecen tras corregir:", sorted(reglas_v - reglas_c))
print("hallazgos activos tras corregir   :", len(corregidos))
