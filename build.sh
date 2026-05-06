#!/bin/bash
# Railway Build Script
# Instala dependencias y prepara el entorno para Playwright

set -e

echo "📦 Instalando dependencias de Python..."
pip install --no-cache-dir -r requirements.txt

echo "🌐 Instalando navegador Chromium para Playwright..."
python -m playwright install chromium --with-deps

echo "✅ Build completado"