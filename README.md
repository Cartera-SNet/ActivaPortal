# 🏥 Activa IT — Descargador de Cartas Glosa Previsora

Aplicación local que automatiza la descarga de los archivos **Envios_D** y **ActaDevolucion**
desde el portal Activa IT para las facturas de Previsora SOAT.

---

## ▶️ Cómo usar

### macOS / Linux
```bash
chmod +x iniciar.sh
./iniciar.sh
```

### Windows
Doble clic en `iniciar.bat`

---

Luego abre tu navegador en **http://localhost:5050**

---

## 📋 Qué descarga

| Estado de factura | Archivo que descarga |
|---|---|
| Auditada: Pendiente de recibir Información | `Envios_D` |
| En radicación: Devolución de entrada | `ActaDevolucion` |

---

## 📂 Archivos descargados

Se guardan en `downloads/[periodo]/` dentro de la misma carpeta de la app.
Puedes cambiar la ruta desde la interfaz.

---

## Requisitos

- Python 3.8+
- Conexión a internet
- El script instala todo automáticamente en un entorno virtual
