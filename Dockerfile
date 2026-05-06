FROM python:3.11-slim

# Instalar dependencias del sistema necesarias para Playwright/Chromium
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    curl \
    ca-certificates \
    fonts-liberation \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdbus-1-3 \
    libdrm2 \
    libgbm1 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libx11-xcb1 \
    libxcomposite1 \
    libxdamage1 \
    libxrandr2 \
    xdg-utils \
    && rm -rf /var/lib/apt/lists/*

# Establecer directorio de trabajo
WORKDIR /app

# Copiar archivos de la aplicación
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Instalar navegador Chromium para Playwright
RUN python -m playwright install chromium --with-deps

# Copiar el resto de la aplicación
COPY . .

# Crear directorio de descargas
RUN mkdir -p downloads

# Exponer puerto (Railway asigna el suyo via PORT env var)
EXPOSE 5050

# Comando de inicio
CMD ["python", "app.py"]