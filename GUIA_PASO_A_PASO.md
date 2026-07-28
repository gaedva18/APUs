# 🚀 GUÍA PASO A PASO - Extractor de APUs PDF a Excel

## 📋 Tabla de Contenidos
1. [Preparación Inicial](#1-preparación-inicial)
2. [Instalación de Dependencias](#2-instalación-de-dependencias)
3. [Descarga de Archivos](#3-descarga-de-archivos)
4. [Estructura de Carpetas](#4-estructura-de-carpetas)
5. [Configuración de Scripts](#5-configuración-de-scripts)
6. [Ejecución del Programa](#6-ejecución-del-programa)
7. [Verificación de Resultados](#7-verificación-de-resultados)

---

## 1️⃣ PREPARACIÓN INICIAL

### ✅ Requisitos Previos

Necesitas tener instalados:
- **Python 3.7+** → [Descargar aquí](https://www.python.org/downloads/)
- **Git** (opcional) → [Descargar aquí](https://git-scm.com/)
- **Un editor de texto** (VSCode, Sublime, Notepad++, etc.)

### ✅ Verificar que Python está instalado

**En Windows (Cmd o PowerShell):**
```bash
python --version
```

**En Linux/Mac (Terminal):**
```bash
python3 --version
```

**Resultado esperado:**
```
Python 3.9.0  (o superior)
```

Si no aparece, descarga Python desde [python.org](https://www.python.org/downloads/)

---

## 2️⃣ INSTALACIÓN DE DEPENDENCIAS

### 📦 Opción A: Instalación Rápida (Recomendado)

#### Paso 1: Crear archivo `requirements.txt`

Crea un archivo llamado `requirements.txt` en tu carpeta de trabajo con este contenido:

```
pdfplumber==0.9.0
openpyxl==3.10.2
pandas==1.5.3
```

#### Paso 2: Instalar todas las dependencias

**En Windows:**
```bash
pip install -r requirements.txt
```

**En Linux/Mac:**
```bash
pip3 install -r requirements.txt
```

**Resultado esperado:**
```
Successfully installed pdfplumber openpyxl pandas
```

---

### 📦 Opción B: Instalación Manual (Si necesitas OCR)

**Para PDF básicos:**
```bash
pip install pdfplumber openpyxl pandas
```

**Para PDFs escaneados (con OCR):**
```bash
pip install pdfplumber pytesseract pdf2image openpyxl pandas pillow
```

Luego instala **Tesseract OCR**:

**Windows:**
1. Descarga: [Tesseract Windows Installer](https://github.com/UB-Mannheim/tesseract/wiki)
2. Ejecuta el instalador
3. Anota la ruta de instalación (ejemplo: `C:\Program Files\Tesseract-OCR`)

**Linux:**
```bash
sudo apt-get update
sudo apt-get install tesseract-ocr
```

**Mac:**
```bash
brew install tesseract
```

---

## 3️⃣ DESCARGA DE ARCHIVOS

### 📥 Paso 1: Clonar o Descargar el Repositorio

**Opción A: Usando Git (Recomendado)**
```bash
git clone https://github.com/gaedva18/APUs.git
cd APUs
```

**Opción B: Descargar ZIP Manual**
1. Ve a: https://github.com/gaedva18/APUs
2. Click en verde "Code" → "Download ZIP"
3. Extrae el archivo ZIP en tu carpeta de trabajo

### 📥 Paso 2: Verifica que tienes estos archivos

```
carpeta_APUs/
├── 003_opt.pdf                 ✅ (debe estar)
├── 004.pdf                     ✅ (debe estar)
├── extract_pdf_to_excel.py     ✅ (descargado de GitHub)
├── extract_pdf_ocr_advanced.py ✅ (descargado de GitHub)
├── formato.xlsx                ✅ (plantilla)
├── README.md                   ✅ (instrucciones)
├── requirements.txt            ✅ (dependencias)
└── GUIA_PASO_A_PASO.md        ✅ (este archivo)
```

---

## 4️⃣ ESTRUCTURA DE CARPETAS

### 📁 Estructura Recomendada

```
📁 Mi Proyecto/
│
├── 📁 PDFs/                    ← Coloca aquí tus PDFs
│   ├── 003_opt.pdf
│   └── 004.pdf
│
├── 📁 Scripts/                 ← Aquí van los scripts Python
│   ├── extract_pdf_to_excel.py
│   └── extract_pdf_ocr_advanced.py
│
├── 📁 Resultados/              ← Aquí se guardarán los Excel
│   └── (se crea automáticamente)
│
├── requirements.txt            ← Dependencias
├── config.py                   ← Configuración (opcional)
└── README.md
```

### ✅ OPCIÓN MÁS SIMPLE (Recomendado para principiantes)

Si prefieres una estructura más simple:

```
📁 APUs/
├── 003_opt.pdf
├── 004.pdf
├── extract_pdf_to_excel.py
├── requirements.txt
└── formato.xlsx
```

**Todos los archivos en la misma carpeta** ✅

---

## 5️⃣ CONFIGURACIÓN DE SCRIPTS

### 📝 Script 1: `extract_pdf_to_excel.py` (BÁSICO)

Este es el script más simple. **No necesita cambios** si:
- Tus PDFs se llaman `003_opt.pdf` y `004.pdf`
- Están en la misma carpeta

**Si quieres personalizarlo:**

Abre el archivo con un editor de texto y busca estas líneas (alrededor de línea 200):

```python
def main():
    """Función principal"""
    print("=" * 60)
    print("EXTRACTOR DE APUs - PDF a Excel")
    print("=" * 60)
    
    # 👇 AQUÍ: Cambia los nombres de los PDFs si es necesario
    pdf_files = ['003_opt.pdf', '004.pdf']
    
    # 👇 AQUÍ: Cambia el nombre del archivo de salida
    output_file = 'APUs_Extraidos.xlsx'
```

**Ejemplo de personalización:**

```python
# Si tus PDFs tienen otros nombres:
pdf_files = ['presupuesto_001.pdf', 'presupuesto_002.pdf']

# Si quieres otro nombre para el Excel:
output_file = 'Mi_Presupuesto_2026.xlsx'
```

### 📝 Script 2: `extract_pdf_ocr_advanced.py` (AVANZADO)

Este script tiene OCR. Personaliza de igual forma:

```python
def main():
    print("=" * 70)
    print("EXTRACTOR AVANZADO DE APUs - PDF a Excel con OCR")
    print("=" * 70)
    
    # Cambia estos archivos
    pdf_files = ['003_opt.pdf', '004.pdf']
    
    # Cambia el nombre de salida
    output_file = 'APUs_Extraidos_Avanzado.xlsx'
```

**Además, si instalaste Tesseract en Windows en otra ruta:**

Busca esta línea (alrededor de línea 40):

```python
# NO NECESITA CAMBIOS si instalaste en: C:\Program Files\Tesseract-OCR
# Si instalaste en otra ruta, descomenta y cambia:
# import pytesseract
# pytesseract.pytesseract.pytesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

---

## 6️⃣ EJECUCIÓN DEL PROGRAMA

### ✅ Paso 1: Abre una Terminal/Consola

**Windows:**
1. Presiona `Windows + R`
2. Escribe: `cmd`
3. Presiona Enter

**Mac:**
1. Presiona `Command + Espacio`
2. Escribe: `terminal`
3. Presiona Enter

**Linux:**
- Abre Terminal normalmente o presiona `Ctrl + Alt + T`

### ✅ Paso 2: Navega a tu carpeta

**Windows:**
```bash
cd C:\ruta\a\tu\carpeta
```

Ejemplo:
```bash
cd C:\Users\MiUsuario\Documentos\APUs
```

**Mac/Linux:**
```bash
cd /ruta/a/tu/carpeta
```

Ejemplo:
```bash
cd ~/Documentos/APUs
```

### ✅ Paso 3: Verifica que estés en la carpeta correcta

```bash
dir          # Windows
ls           # Mac/Linux
```

Deberías ver:
```
003_opt.pdf
004.pdf
extract_pdf_to_excel.py
requirements.txt
...
```

### ✅ Paso 4: Ejecuta el script

**Para script BÁSICO:**
```bash
python extract_pdf_to_excel.py
```

**Para script AVANZADO (con OCR):**
```bash
python extract_pdf_ocr_advanced.py
```

### 📊 Resultado en Terminal

```
============================================================
EXTRACTOR DE APUs - PDF a Excel
============================================================
📄 Procesando: 003_opt.pdf
   Total de páginas: 49
   ✓ Página 1: 1 tabla(s) encontrada(s)
   ✓ Página 2: 1 tabla(s) encontrada(s)

✅ Total de registros extraídos: 234

📄 Procesando: 004.pdf
   Total de páginas: 10
   ✓ Página 1: 1 tabla(s) encontrada(s)

✅ Total de registros extraídos: 45

✅ Archivo exportado: APUs_Extraidos.xlsx

============================================================
✅ PROCESO COMPLETADO EXITOSAMENTE
📁 Archivo guardado: APUs_Extraidos.xlsx
============================================================
```

---

## 7️⃣ VERIFICACIÓN DE RESULTADOS

### 📂 Paso 1: Busca el archivo generado

En tu carpeta de trabajo, verás:

```
📁 APUs/
├── 003_opt.pdf
├── 004.pdf
├── extract_pdf_to_excel.py
└── 📊 APUs_Extraidos.xlsx  ← NUEVO ✅
```

### 📊 Paso 2: Abre el Excel

Haz doble click en `APUs_Extraidos.xlsx`

### ✅ Paso 3: Verifica la estructura

**Deberías ver hojas así:**

```
├── 📋 Hoja "003 opt"
│   ├── Columna: Página
│   ├── Columna: (datos de la tabla 1)
│   ├── Columna: (datos de la tabla 2)
│   └── Fila 1: Encabezados (AZULES)
│   └── Fila 2+: Datos
│
├── 📋 Hoja "004"
│   ├── Columna: Página
│   ├── Columna: (datos de la tabla 1)
│   └── Fila 1: Encabezados (AZULES)
│   └── Fila 2+: Datos
│
└── 📋 Hoja "Consolidado" (solo versión avanzada)
    ├── Columna: Fuente
    ├── Columna: Página
    ├── Columna: (datos combinados)
    └── Fila 1: Encabezados (AZULES)
```

### 🎨 Características que DEBERÍAS VER

✅ **Encabezados en AZUL** con letras BLANCAS  
✅ **Bordes** alrededor de todas las celdas  
✅ **Ancho automático** de columnas  
✅ **Texto centrado** en encabezados  
✅ **Números con formato** (decimales)  

---

## 🐛 SOLUCIÓN DE PROBLEMAS

### ❌ Error: "ModuleNotFoundError: No module named 'pdfplumber'"

**Solución:**
```bash
pip install pdfplumber --upgrade
```

---

### ❌ Error: "No such file or directory: '003_opt.pdf'"

**Solución:**
1. Verifica que el PDF está en la misma carpeta que el script
2. Verifica que el nombre es exactamente `003_opt.pdf` (mayúsculas/minúsculas importan)
3. Comprueba con:
   ```bash
   dir          # Windows
   ls           # Mac/Linux
   ```

---

### ❌ Error: "Tesseract is not installed"

**Solución:**
1. Instala Tesseract según tu SO (ver sección 2)
2. En Windows, después de instalar, reinicia la terminal
3. Verifica con:
   ```bash
   tesseract --version
   ```

---

### ❌ El script se ejecuta pero no extrae datos

**Solución (intenta en orden):**

1. **Verifica que Python ve el PDF:**
   ```python
   import os
   print(os.listdir())  # Debería mostrar tu PDF
   ```

2. **Prueba el script avanzado:**
   ```bash
   python extract_pdf_ocr_advanced.py
   ```

3. **Verifica que el PDF no está corrupto:**
   - Intenta abrirlo manualmente en Adobe Reader

4. **Revisa los permisos del archivo:**
   ```bash
   chmod +r 003_opt.pdf  # Mac/Linux
   ```

---

## 💡 TIPS Y TRUCOS

### 🔄 Procesar archivos automáticamente

**Windows (Batch):**

Crea un archivo `procesar.bat`:

```batch
@echo off
python extract_pdf_to_excel.py
pause
```

Haz doble click para ejecutar.

**Mac/Linux (Bash):**

Crea un archivo `procesar.sh`:

```bash
#!/bin/bash
python3 extract_pdf_to_excel.py
```

Luego:
```bash
chmod +x procesar.sh
./procesar.sh
```

---

### 📋 Procesar múltiples carpetas

**Windows:**
```bash
for /d %f in (*) do cd %f && python extract_pdf_to_excel.py && cd ..
```

**Mac/Linux:**
```bash
for folder in */; do (cd "$folder" && python3 extract_pdf_to_excel.py); done
```

---

### 🎨 Personalizar colores del Excel

En el script, busca:

```python
header_fill = PatternFill(start_color="4472C4", ...)  # Azul actual
```

**Códigos de color (formato hexadecimal):**
- Azul: `4472C4`
- Verde: `70AD47`
- Rojo: `C55A11`
- Gris: `595959`
- Negro: `000000`
- Blanco: `FFFFFF`

Cambia `"4472C4"` por tu color preferido.

---

## ✅ CHECKLIST FINAL

Antes de ejecutar, verifica:

- [ ] Python está instalado (`python --version`)
- [ ] Dependencias instaladas (`pip list | grep pdfplumber`)
- [ ] PDFs están en la carpeta correcta
- [ ] Script Python está en la carpeta correcta
- [ ] Abriste la terminal en la carpeta correcta (`dir` o `ls` muestra los archivos)
- [ ] Ejecutaste el comando correcto
- [ ] Excel se generó sin errores

---

## 📞 SOPORTE

Si tienes problemas:

1. **Lee el error completo** en la terminal
2. **Copia el error** y busca en Google
3. **Abre un Issue** en GitHub: https://github.com/gaedva18/APUs/issues
4. **Incluye:**
   - El error completo
   - Tu versión de Python (`python --version`)
   - Tu sistema operativo (Windows/Mac/Linux)
   - El nombre de tu PDF

---

## 📚 PRÓXIMOS PASOS

Una vez que domines esto:

- [ ] Personaliza los colores del Excel
- [ ] Automatiza con un script en lotes (batch/bash)
- [ ] Integra con Google Drive o Dropbox
- [ ] Crea una interfaz gráfica (GUI) con Tkinter
- [ ] Deploy en la nube (Heroku, AWS, etc.)

---

**¡Listo! 🎉 Ya puedes extraer tus APUs de PDFs a Excel**

Última actualización: 2026-07-28  
Versión: 1.0
