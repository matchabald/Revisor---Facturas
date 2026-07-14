# 🧾 Revisor de Facturas de Transporte

App para verificar facturas de transporte contra tarifas registradas en base de datos.

---

## 📁 Archivos del proyecto

```
revisor-facturas/
├── INVOICE RUNNER.py              ← Script principal
├── DB.xlsx             ← Base de datos de tarifas (tu archivo)
├── requirements.txt    ← Librerías necesarias
├── README.md           ← Este archivo
└── revisiones.xlsx     ← Se crea automáticamente al guardar
```

---

## 🚀 Cómo correrlo localmente

### 1. Instalá Python
Descargá Python 3.11+ desde [python.org](https://python.org/downloads)  
⚠️ Marcá la casilla **"Add Python to PATH"** durante la instalación.

### 2. Instalá las librerías
Abrí el CMD (símbolo del sistema) y escribí:
```bash
pip install streamlit pandas openpyxl
```

### 3. Corré la app
Navegá a la carpeta del proyecto y ejecutá:
```bash
streamlit run app.py
```
Se abre automáticamente en tu navegador en `http://localhost:8501`

---

## ☁️ Cómo subir a GitHub y desplegarlo gratis (Streamlit Cloud)

### Desplegar en Streamlit Cloud
1. Entrá a [share.streamlit.io](https://share.streamlit.io)
2. Iniciá sesión con tu cuenta de GitHub
3. Hacé clic en **"New app"**
4. Seleccioná tu repositorio `revisor-facturas`
5. En **"Main file path"** escribí: `app.py`
6. Hacé clic en **"Deploy"**

---

## ⚙️ Personalización

Abrí `Invoice Runner` y edite estas líneas al principio, si es necesario:

```python
TASA_CAMBIO = 36.6243   # tasa USD → Córdobas modicaficable 
DB_FILE = "DB.xlsx"   # Nombre de archivo de tarifas
```

---

## 📌 Notas importantes

- El archivo `revisiones.xlsx` se crea automáticamente la primera vez que guardás una revisión.
- Si usás Streamlit Cloud, el historial **no persiste** entre sesiones (limitación del servidor gratuito). Para eso, en el futuro podemos conectar Google Sheets.
- La tolerancia de verificación es **±$0.01**.
