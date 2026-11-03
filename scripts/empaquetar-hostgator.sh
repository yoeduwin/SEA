#!/usr/bin/env bash
# Genera el paquete para subir a ejecutivambiental.com por el Administrador
# de archivos de cPanel:
#
#   clientes/  → SEAPD.html  (https://ejecutivambiental.com/clientes)
#   asesores/  → PAIC.html   (https://ejecutivambiental.com/asesores)
#
# Cada carpeta lleva su index.html, recaptcha.js y .htaccess, para que sea
# independiente. Uso:  bash scripts/empaquetar-hostgator.sh [carpeta-salida]
set -euo pipefail

RAIZ="$(cd "$(dirname "$0")/.." && pwd)"
SALIDA="${1:-$RAIZ/dist}"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

armar() {   # $1 carpeta destino, $2 html de origen
  mkdir -p "$TMP/$1"
  cp "$RAIZ/$2"                  "$TMP/$1/index.html"
  cp "$RAIZ/recaptcha.js"        "$TMP/$1/recaptcha.js"
  cp "$RAIZ/hostgator/.htaccess" "$TMP/$1/.htaccess"
}

armar clientes SEAPD.html
armar asesores PAIC.html

mkdir -p "$SALIDA"
ZIP="$SALIDA/sea-portales-hostgator.zip"
rm -f "$ZIP"
(cd "$TMP" && zip -qr "$ZIP" clientes asesores)

echo "Paquete listo: $ZIP"
unzip -l "$ZIP"
