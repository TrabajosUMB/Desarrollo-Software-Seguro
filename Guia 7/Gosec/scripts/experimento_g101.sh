#!/usr/bin/env bash
# Experimento: ¿que contraseñas escritas en el codigo detecta la regla G101 de gosec?
# Uso (desde Git Bash, en la carpeta Gosec):  bash scripts/experimento_g101.sh
set -euo pipefail

RAIZ="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
GOSEC="${GOSEC:-$(go env GOBIN)/gosec.exe}"
[ -x "$GOSEC" ] || GOSEC="gosec"
SALIDA="$RAIZ/resultados/experimento_g101.txt"

cd "$RAIZ/experimentos/g101"
{
  echo "Archivo analizado: experimentos/g101/main.go"
  echo
  nl -ba main.go | sed -n '5,11p'
  echo
  echo "=== 1) Configuracion por defecto ==="
  "$GOSEC" -no-fail -color=false -terse -include=G101 ./... 2>&1 | grep -E "G101|Issues" | sed -E 's#\[[^]]*main\.go:#[main.go:#'
  echo
  echo "=== 2) Con ignore_entropy=true (config_sin_entropia.json) ==="
  "$GOSEC" -no-fail -color=false -terse -include=G101 -conf=config_sin_entropia.json ./... 2>&1 | grep -E "G101|Issues" | sed -E 's#\[[^]]*main\.go:#[main.go:#'
} | tee "$SALIDA"
