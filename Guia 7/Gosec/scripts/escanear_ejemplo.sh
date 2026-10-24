#!/usr/bin/env bash
# Escanea el codigo propio (ejemplo-propio) antes y despues de corregirlo.
# Uso (desde Git Bash, en la carpeta Gosec):  bash scripts/escanear_ejemplo.sh
set -euo pipefail

RAIZ="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
RES="$RAIZ/resultados"
GOSEC="${GOSEC:-$(go env GOBIN)/gosec.exe}"
[ -x "$GOSEC" ] || GOSEC="gosec"
mkdir -p "$RES"

cd "$RAIZ"
( cd ejemplo-propio && go mod download && go vet ./... )   # gosec necesita que el codigo compile

for version in vulnerable corregido; do
  "$GOSEC" -no-fail -color=false -track-suppressions -fmt=text -out="$RES/ejemplo_$version.txt"  "./ejemplo-propio/$version/..."
  "$GOSEC" -no-fail              -track-suppressions -fmt=json -out="$RES/ejemplo_$version.json" "./ejemplo-propio/$version/..."
done

py -I "$RAIZ/scripts/limpiar_rutas.py" "$RES" "$RAIZ"
py -I "$RAIZ/scripts/comparar.py"
