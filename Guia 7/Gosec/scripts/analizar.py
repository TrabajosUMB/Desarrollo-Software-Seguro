# -*- coding: utf-8 -*-
"""Resume los resultados de gosec (resultados/govwa.json).

Uso:
  py scripts/analizar.py                  resumen por severidad, confianza, regla y archivo
  py scripts/analizar.py G701 G401        ademas, el codigo marcado de esas reglas
  py scripts/analizar.py --csv            escribe resultados/hallazgos.csv
  py scripts/analizar.py --archivo sqli/function.go   hallazgos de un archivo
"""
import collections
import csv
import json
import pathlib
import sys

sys.stdout.reconfigure(encoding="utf-8")

RAIZ = pathlib.Path(__file__).resolve().parent.parent
RESULTADOS = RAIZ / "resultados"
ORDEN = {"HIGH": 0, "MEDIUM": 1, "LOW": 2}


def archivo_corto(ruta: str) -> str:
    return ruta.replace("\\", "/").split("govwa/")[-1]


def cargar():
    datos = json.loads((RESULTADOS / "govwa.json").read_text(encoding="utf-8"))
    return datos["Issues"], datos["Stats"]


def resumen(issues, stats):
    print(f"archivos: {stats['files']} | lineas: {stats['lines']} | hallazgos: {len(issues)} | #nosec: {stats['nosec']}")
    print("severidad :", dict(collections.Counter(i["severity"] for i in issues)))
    print("confianza :", dict(collections.Counter(i["confidence"] for i in issues)))
    cruce = collections.Counter((i["severity"], i["confidence"]) for i in issues)
    print("severidad x confianza:", {f"{s}/{c}": n for (s, c), n in sorted(cruce.items(), key=lambda kv: (ORDEN[kv[0][0]], kv[0][1]))})
    print()
    print(f"{'regla':<6} {'severidad':<9} {'confianza':<9} {'CWE':<8} {'n':>2}  descripcion")
    por_regla = collections.defaultdict(list)
    for i in issues:
        por_regla[i["rule_id"]].append(i)
    for regla, lista in sorted(por_regla.items(), key=lambda kv: (ORDEN[kv[1][0]["severity"]], kv[0])):
        i = lista[0]
        print(f"{regla:<6} {i['severity']:<9} {i['confidence']:<9} CWE-{i['cwe']['id']:<4} {len(lista):>2}  {i['details'][:78]}")
    print()
    print("hallazgos por archivo:")
    for archivo, n in collections.Counter(archivo_corto(i["file"]) for i in issues).most_common():
        print(f"  {n:>2}  {archivo}")


def detalle(issues, reglas):
    for regla in reglas:
        for i in issues:
            if i["rule_id"] == regla:
                print(f"=== {regla} {i['severity']}/{i['confidence']} CWE-{i['cwe']['id']}  {archivo_corto(i['file'])}:{i['line']}")
                print(i["code"].rstrip()[:450])
                print()


def por_archivo(issues, fragmento):
    for i in sorted(issues, key=lambda x: int(x["line"].split("-")[0])):
        if fragmento in archivo_corto(i["file"]):
            print(f"{i['rule_id']:<5} {i['severity']:<7} {i['confidence']:<6} linea {i['line']:<7} {i['details'][:70]}")


def exportar_csv(issues):
    destino = RESULTADOS / "hallazgos.csv"
    with destino.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["regla", "severidad", "confianza", "cwe", "archivo", "linea", "descripcion"])
        for i in sorted(issues, key=lambda x: (ORDEN[x["severity"]], x["rule_id"], archivo_corto(x["file"]), int(x["line"].split("-")[0]))):
            w.writerow([i["rule_id"], i["severity"], i["confidence"], f"CWE-{i['cwe']['id']}",
                        archivo_corto(i["file"]), i["line"], i["details"]])
    print("escrito", destino)


if __name__ == "__main__":
    issues, stats = cargar()
    args = sys.argv[1:]
    if "--csv" in args:
        exportar_csv(issues)
    elif args and args[0] == "--archivo":
        por_archivo(issues, args[1])
    elif args:
        detalle(issues, args)
    else:
        resumen(issues, stats)
