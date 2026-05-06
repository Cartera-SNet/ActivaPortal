#!/bin/bash
# ============================================================
#  Activa IT — Descargador de Cartas Glosa Previsora
#  Script de instalación y ejecución
# ============================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$SCRIPT_DIR/venv"
APP="$SCRIPT_DIR/app.py"

echo ""
echo "============================================================"
echo "  🏥 Activa IT — Descargador de Cartas Glosa Previsora"
echo "============================================================"
echo ""

# ── 1. Crear entorno virtual si no existe ────────────────────
if [ ! -d "$VENV_DIR" ]; then
  echo "📦 Creando entorno virtual Python..."
  python3 -m venv "$VENV_DIR"
fi

# ── 2. Activar entorno virtual ───────────────────────────────
source "$VENV_DIR/bin/activate"

# ── 3. Instalar dependencias ─────────────────────────────────
echo "📥 Instalando dependencias (Flask, Playwright)..."
pip install --quiet flask playwright 2>&1 | grep -v "^$" || true

# ── 4. Instalar navegador Chromium para Playwright ───────────
echo "🌐 Verificando Playwright Chromium..."
python3 -m playwright install chromium 2>&1 | tail -3

# ── 5. Crear carpeta de descargas ────────────────────────────
mkdir -p "$SCRIPT_DIR/downloads"

# ── 6. Abrir el navegador automáticamente (macOS / Linux) ────
echo ""
echo "🚀 Iniciando servidor en http://localhost:5050"
echo "   (presiona Ctrl+C para detener)"
echo ""

# Abrir navegador en segundo plano
sleep 1.5 && {
  if command -v xdg-open &>/dev/null; then
    xdg-open http://localhost:5050 &
  elif command -v open &>/dev/null; then
    open http://localhost:5050 &
  fi
} &

# ── 7. Ejecutar la app ───────────────────────────────────────
python3 "$APP"
