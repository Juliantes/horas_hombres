#!/bin/sh
# Zip liviano para llevar a Windows y ejecutar build.ps1 alli.
set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$ROOT/dist/HorasHombre-codigo-win.zip"
mkdir -p "$ROOT/dist"
rm -f "$OUT"
cd "$ROOT"
zip -j "$OUT" \
  HorasHombre.bat \
  build.ps1 \
  horas_hombre.py \
  sap_import.py \
  requirements-portable.txt \
  PAQUETE-WINDOWS.txt
echo "Creado: $OUT"