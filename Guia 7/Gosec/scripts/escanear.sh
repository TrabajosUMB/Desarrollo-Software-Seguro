#!/usr/bin/env bash
# Escanea GoVWA con gosec y guarda los resultados en ../resultados.
# Requisitos: Go y gosec (go install github.com/securego/gosec/v2/cmd/gosec@latest).
# Uso (desde Git Bash, en la carpeta Gosec):  bash scripts/escanear.sh
set -euo pipefail

RAIZ="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OBJ="$RAIZ/objetivos/govwa"
RES="$RAIZ/resultados"
GOSEC="${GOSEC:-$(go env GOBIN)/gosec.exe}"
[ -x "$GOSEC" ] || GOSEC="gosec"

if [ ! -d "$OBJ" ]; then
  mkdir -p "$RAIZ/objetivos"
  git clone https://github.com/0c34/govwa.git "$OBJ"
  git -C "$OBJ" checkout 4058f79f31eeb4a36d8f1e64bba1f0c899646e6f   # commit analizado
fi
mkdir -p "$RES"

# Dependencias del proyecto analizado (solo descarga, no ejecuta su codigo)
( cd "$OBJ" && go mod download )

cd "$RAIZ"
# Cada formato es una ejecucion real de gosec; -no-fail evita que el codigo de salida 1 (hay hallazgos) corte el script.
"$GOSEC" -no-fail -color=false -fmt=text    -out="$RES/govwa.txt"   ./objetivos/govwa/...
"$GOSEC" -no-fail            -fmt=json    -out="$RES/govwa.json"  ./objetivos/govwa/...
"$GOSEC" -no-fail            -fmt=sarif   -out="$RES/govwa.sarif" ./objetivos/govwa/...
"$GOSEC" -no-fail            -fmt=html    -out="$RES/govwa.html"  ./objetivos/govwa/...
# Solo lo grave: filtro por severidad y confianza altas
"$GOSEC" -no-fail -color=false -severity=high -confidence=high -fmt=text -out="$RES/govwa_solo_graves.txt" ./objetivos/govwa/...

# Entorno exacto, para poder repetir el analisis
{
  echo "fecha:        $(date '+%Y-%m-%d %H:%M')"
  echo "gosec:        $(go version -m "$GOSEC" | awk '$1=="mod"{print $3" "$4}' | head -1 | cut -c1-60)"
  echo "go:           $(go version)"
  echo "objetivo:     https://github.com/0c34/govwa  commit $(git -C "$OBJ" rev-parse HEAD)"
  echo "comando base: gosec -no-fail -fmt=<formato> -out=<archivo> ./objetivos/govwa/..."
} > "$RES/entorno.txt"

# El repositorio es publico: se quitan las rutas absolutas del equipo de todos los resultados.
py -I "$RAIZ/scripts/limpiar_rutas.py" "$RES" "$RAIZ"
echo "Resultados en $RES"
