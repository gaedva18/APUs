# 📊 Extractor de APUs - PDF a Excel

Script profesional para **extraer datos de APUs desde PDFs** y exportarlos a Excel con formato automático.

---

## 🎯 Características

✅ **Extracción inteligente de datos**
- Detecta y extrae tablas automáticamente
- Soporta PDFs nativos y escaneados (OCR)
- Procesa texto estructurado

✅ **Formato profesional en Excel**
- Encabezados con colores corporativos
- Bordes y alineación automática
- Ancho de columnas ajustable
- Múltiples hojas (una por PDF)

✅ **Dos opciones de procesamiento**
1. **Script básico** - Rápido y ligero
2. **Script avanzado** - Con OCR para PDFs escaneados

---

## 🚀 Instalación Rápida

### Paso 1: Instalar dependencias

```bash
# Instalación básica
pip install pdfplumber openpyxl pandas

# Instalación avanzada (con OCR)
pip install pdfplumber pytesseract pdf2image openpyxl pandas pillow
```

### Paso 2: Instalar Tesseract (solo para OCR avanzado)

**En Windows:**
```bash
# Descargar desde: https://github.com/UB-Mannheim/tesseract/wiki
# O usar chocolatey:
choco install tesseract
```

**En Linux (Ubuntu/Debian):**
```bash
sudo apt-get install tesseract-ocr
```

**En macOS:**
```bash
brew install tesseract
```

### Paso 3: Preparar archivos

Coloca los siguientes archivos en la misma carpeta:
```
📁 tu_carpeta/
├── 003_opt.pdf
├── 004.pdf
├── extract_pdf_to_excel.py          (o extract_pdf_ocr_advanced.py)
└── formato.xlsx                     (plantilla - opcional)
```

---

## 📝 Uso

### Opción 1: Script Básico (Recomendado)

```bash
python extract_pdf_to_excel.py
```

**Salida:**
- ✅ `APUs_Extraidos.xlsx` 

**Ventajas:**
- ⚡ Muy rápido
- 📦 Pocas dependencias
- 🎯 Ideal para PDFs nativos

---

### Opción 2: Script Avanzado con OCR

```bash
python extract_pdf_ocr_advanced.py
```

**Salida:**
- ✅ `APUs_Extraidos_Avanzado.xlsx` (con hoja consolidada)

**Ventajas:**
- 🔍 Soporta PDFs escaneados
- 📊 Hoja consolidada automática
- 🎨 Formato más avanzado

---

## 📊 Estructura del Archivo de Salida

### Hojas en el Excel:

```
📄 APUs_Extraidos.xlsx
├── 📋 003 opt
│   ├── Página | Columna_1 | Columna_2 | ...
│   ├── 1     | Dato_A    | Dato_B    | ...
│   └── 2     | Dato_C    | Dato_D    | ...
│
├── 📋 004
│   ├── Página | Columna_1 | Columna_2 | ...
│   ├── 1     | Dato_E    | Dato_F    | ...
│   └── 2     | Dato_G    | Dato_H    | ...
│
└── 📋 Consolidado (solo versión avanzada)
    ├── Fuente | Página | Columna_1 | Columna_2 | ...
    ├── 003 opt| 1      | Dato_A    | Dato_B    | ...
    └── 004    | 1      | Dato_E    | Dato_F    | ...
```

---

## 🔧 Configuración Avanzada

### Personalizar nombres de hojas

Edita el script y cambia:

```python
pdf_files = ['003_opt.pdf', '004.pdf']
output_file = 'APUs_Extraidos.xlsx'
```

### Cambiar estilos de Excel

En la clase `ExcelExporter`, modifica:

```python
# Color de encabezado
header_fill = PatternFill(start_color="4472C4", ...)  # Azul

# Fuente de encabezado
header_font = Font(bold=True, color="FFFFFF", size=11)

# Ancho máximo de columnas
adjusted_width = min(max_length + 2, 50)  # Máximo 50 caracteres
```

---

## 🐛 Solución de Problemas

### Error: "ModuleNotFoundError: No module named 'pdfplumber'"

```bash
pip install pdfplumber --upgrade
```

### Error: "Tesseract is not installed"

Instala Tesseract según tu sistema operativo (ver sección Instalación).

### El PDF no extrae datos

1. **Verifica que el archivo existe:**
   ```bash
   ls -la *.pdf  # Linux/Mac
   dir *.pdf     # Windows
   ```

2. **Prueba el script avanzado:**
   ```bash
   python extract_pdf_ocr_advanced.py
   ```

3. **Revisa los permisos:**
   ```bash
   chmod +r *.pdf  # Linux/Mac
   ```

---

## 📈 Ejemplos de Salida

### Datos Extraídos de Tabla

| Página | Código   | Descripción           | Cantidad | Unitario | Total      |
|--------|----------|----------------------|----------|----------|------------|
| 1      | APU-001  | Material A            | 10       | 100.00   | 1,000.00   |
| 1      | APU-002  | Mano de obra          | 5        | 200.00   | 1,000.00   |
| 2      | APU-003  | Equipos               | 2        | 500.00   | 1,000.00   |

---

## 🎨 Formato del Excel

```
┌─────────────────────────────────────────┐
│ ENCABEZADO (Azul, Negrita, Blanco)      │
├─────────────────────────────────────────┤
│ Datos con bordes                        │
│ Texto ajustado automáticamente          │
│ Números formateados con decimales       │
└─────────────────────────────────────────┘
```

---

## 📚 Documentación de Clases

### `APUExtractor`
```python
extractor = APUExtractor('003_opt.pdf')
data = extractor.extract()  # Retorna lista de diccionarios
```

### `ExcelExporter`
```python
exporter = ExcelExporter('formato.xlsx', 'salida.xlsx')
exporter.export({'Hoja1': df1, 'Hoja2': df2})
```

### `AdvancedAPUExtractor` (versión con OCR)
```python
extractor = AdvancedAPUExtractor('004.pdf')
data = extractor.extract()  # Con OCR si es necesario
```

---

## 🤝 Contribuir

Para mejoras o reportar bugs:
1. Abre un Issue en GitHub
2. Proporciona detalles del error
3. Incluye un ejemplo del PDF problemático

---

## 📄 Licencia

MIT License - Libre para uso comercial y personal

---

## ⚡ Tips de Productividad

### Procesar múltiples carpetas

```bash
# Crea un script en lotes
for folder in */; do
    cd "$folder"
    python extract_pdf_to_excel.py
    cd ..
done
```

### Automatizar con cron (Linux/Mac)

```bash
# Edita crontab
crontab -e

# Agrega esta línea (diariamente a las 8 AM)
0 8 * * * cd /ruta/a/carpeta && python extract_pdf_to_excel.py
```

### Automatizar con Task Scheduler (Windows)

1. Abre Task Scheduler
2. Crea tarea → Acciones → Nueva
3. Programa: `python.exe`
4. Argumentos: `C:\ruta\extract_pdf_to_excel.py`

---

## 📞 Soporte

- 📧 Email: costospresupuesto.myq@gmail.com
- 💬 GitHub Issues: [Reportar problema](https://github.com/gaedva18/APUs/issues)
- 📖 Wiki: [Más información](https://github.com/gaedva18/APUs/wiki)

---

**Última actualización:** 2026-07-28  
**Versión:** 2.1.0  
**Estado:** ✅ Producción
