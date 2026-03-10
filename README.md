**DocuMapper** es una herramienta que **automatiza la generación de documentos Word personalizados** usando datos de Excel.

**El problema que resuelve:**
- Antes: Crear 100 documentos Word = 2-3 horas de trabajo manual
- Ahora: Crear 100 documentos Word = 30 segundos aproximadamente

---

## ARQUITECTURA

```
┌─────────────────────────────────────────────────────────┐
│                   USUARIO (Frontend)                     │
│              (Navegador web - index.html)                │
└──────────────────────┬──────────────────────────────────┘
                       │ (HTTP Requests)
                       ▼
┌─────────────────────────────────────────────────────────┐
│              SERVIDOR (Backend - Node.js)                │
│                                                          │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  │
│  │   Excel      │  │   Word       │  │   Mapping    │  │
│  │ Controller   │  │ Controller   │  │ Controller   │  │
│  │              │  │              │  │              │  │
│  │ • Parse      │  │ • Extract    │  │ • Generate   │  │
│  │ • Format     │  │ • Validate   │  │ • Replace    │  │
│  │ • Return     │  │ • Return     │  │ • Zip        │  │
│  └──────────────┘  └──────────────┘  └──────────────┘  │
└─────────────────────────────────────────────────────────┘
                       │ (JSON)
                       ▼
┌─────────────────────────────────────────────────────────┐
│              ALMACENAMIENTO (Temporal)                   │
│                                                          │
│  uploads/     downloads/     (En memoria)               │
│  (Archivos)   (Resultados)   (Plantillas)              │
└─────────────────────────────────────────────────────────┘
```

##  FLUJO DE DATOS

### **1. Usuario carga un Excel**
```
usuario.xlsx
     ↓
[excelInput] → [uploadExcel()] → API: /api/excel/parse
     ↓
Servidor: ExcelJS lee el archivo
     ↓
Extrae: Headers (columnas) + Data (filas)
     ↓
Responde JSON: { headers: [...], data: [...] }
     ↓
Muestra en Paso 2 ✓
```

### **2. Usuario carga una Plantilla Word**
```
plantilla.docx
     ↓
[wordInput] → [uploadWord()] → API: /api/word/extract-variables
     ↓
Servidor: PizZip + Docxtemplater leen el XML
     ↓
Busca patrones: {{variable}} en el documento
     ↓
Responde JSON: { placeholders: [...], templateId: "uuid" }
     ↓
Muestra en Paso 3 ✓
```

### **3: Usuario mapea columnas ↔ Placeholders**
```
Excel Columns        →    Word Placeholders
─────────────────        ──────────────────
nombre              →    {{nombre}}
empresa             →    {{empresa}}
fecha_contrato      →    {{fecha}}
     ↓
Guarda el mapeo en variable: mappings = { nombre: "nombre", ... }
     ↓
Muestra en Paso 4 ✓
```

### **4: Generación de Documentos (El CORE)**
```
Para CADA fila del Excel:
  → Tomar valores: { nombre: "Juan", empresa: "XYZ", fecha: "17/12/25" }
  → Reemplazar en plantilla:
    {{nombre}} → Juan
    {{empresa}} → XYZ
    {{fecha}} → 17/12/25
  → Generar documento: documento_Juan_1.docx
  → Guardar en /downloads/

Resultado: documento_Juan_1.docx, documento_Maria_2.docx, ...
     ↓
Muestra en Paso 5 ✓
```

### **5: Descarga en ZIP**
```
Todos los documentos → archiver.zip() → documentos.zip
     ↓
Usuario descarga un solo archivo en lugar de 100
```

---

## COMPONENTES TÉCNICOS

### **Frontend (Cliente)**
```
index.html
├── UI interactiva (5 pasos)
├── Validación de archivos
├── Gestión de estado (variables globales)
└── Comunicación API (fetch)

app.js
├── uploadExcel()      → POST /api/excel/parse
├── uploadWord()       → POST /api/word/extract-variables
├── showPreview()      → Muestra datos de muestra
├── generateAll()      → POST /api/mapping/generate
├── downloadAll()      → POST /api/mapping/download-all-docx
└── Gestión de UI (show/hide steps)

styles.css
├── Responsive design
├── Gradientes y animaciones
├── Mobile-first approach
```

### **Backend (Servidor)**

#### **excelController.js**
```javascript
parseExcel()
├── Lee archivo XLSX con ExcelJS
├── Extrae headers (fila 1)
├── Extrae datos (filas 2+)
├── Convierte fechas a DD/MM/YY
├── Convierte objetos a strings
└── Retorna JSON { headers, data }
```

#### **wordController.js**
```javascript
extractVariables()
├── Lee archivo DOCX
├── Busca patrones {{variable}} en XML
├── Detecta varios formatos (espacios variables)
├── Limpia y deduplica
└── Retorna JSON { placeholders, templateId }

getStoredTemplate()
├── Obtiene ruta del archivo guardado (en memoria)

clearTemplate()
├── Elimina archivo temporal
```

#### **mappingController.js**
```javascript
previewDocument()
├── Mapea datos de muestra
├── Genera documento de prueba
├── Devuelve descargable

generateDocuments()
├── Para CADA fila:
│  ├── Mapear datos
│  └── Reemplazar placeholders
│     ├── Intenta docxtemplater.render()
│     ├── Si falla: regex replacement manual
│  └── Guardar .docx
├── Retorna lista de archivos

downloadAllDocx()
├── Recibe lista de filenames
├── Crea ZIP con archiver
├── Devuelve archivo descargable

downloadAllPdf()
├── Igual que Docx
├── Solo cambia nombre a .pdf
```

#### **server.js**
```javascript
Middlewares:
├── CORS (permitir requests del frontend)
├── express.json (parsear JSON)
├── multer (manejar uploads)
└── express.static (servir archivos públicos)

Rutas:
├── POST /api/excel/parse
├── POST /api/word/extract-variables
├── POST /api/mapping/generate
├── POST /api/mapping/preview
├── POST /api/mapping/download-all-docx
├── POST /api/mapping/download-all-pdf
└── GET /api/download/:filename
```

---

## FLUJO DE ARCHIVOS

```
Usuario
  ├─ Excel: subido → /uploads/ → parseado → almacenado en memoria
  ├─ Word: subido → /uploads/ → leído → almacenado en memoria  
  └─ Generados: creados → /downloads/ → descargables → eliminados

Ciclo de vida:
1. Upload
2. Procesamiento en servidor
3. Almacenamiento temporal
4. Descarga
5. Limpieza automática
```

---

## TECNOLOGÍAS INVOLUCRADAS

### **Bibliotecas clave:**

| Librería | Función | Por qué |
|----------|---------|--------|
| **ExcelJS** | Leer Excel | Estándar en Node.js para XLSX |
| **docxtemplater** | Procesar Word | Inyector de variables en DOCX |
| **PizZip** | Manipular ZIP interno | Word es ZIP con XML adentro |
| **archiver** | Crear ZIPs | Empaquetar múltiples archivos |
| **multer** | Subir archivos | Middleware estándar Express |
| **Express** | Servidor HTTP | Framework web más popular |
| **Node.js** | Runtime API | JavaScript en servidor |

---

## CONCEPTOS CLAVE PARA EXPLICAR

1. **Template:** Es la plantilla Word con "espacios en blanco" ({{nombre}})
2. **Placeholder:** El "espacio en blanco" que será reemplazado
3. **Mapeo:** Decirle al programa "la columna 'nombre' del Excel va aquí"
4. **Lote:** Procesar muchos documentos a la vez
5. **XML manipulation:** Word es un ZIP que contiene XML
6. **Regex replacement:** Búsqueda y reemplazo con patrones
7. **Streaming:** Los ZIPs se generan en tiempo real (no en RAM)
8. **Async operations:** Manejo de uploads/downloads sin bloqueos

---
##  VENTAJAS ARQUITECTÓNICAS

**Separación de responsabilidades**
- Cada controlador hace UNA cosa
- Fácil de mantener y debuggear

**Stateless (sin estado)**
- Cada request es independiente
- Escalable horizontalmente

**Procesamiento robusto**
- Fallback si docxtemplater falla
- Limpieza automática de temporales

**Frontend-Backend desacoplado**
- Fácil cambiar UI sin tocar servidor
- Fácil cambiar API sin tocar frontend

---

## DocuMapper

### **Para clientes:**
> "Carga tu Excel con datos, tu plantilla Word con placeholders, 
> mapea las columnas, y nosotros generamos automáticamente todos 
> los documentos. En segundos lo que antes tomaba horas."

### **Para gestores:**
> "Reduce errores manuales, ahorra tiempo de operación, 
> escalable a cualquier volumen, bajo costo de mantenimiento."

### **Para desarrolladores:**
> "Arquitectura modular con Express + Node.js, 
> procesamiento ZIP interno con regex fallback, 
> APIs RESTful limpias, fácil de extender."

---

## Preguntas frecuentes!

**P: "¿Qué ocurre si tengo 10,000 filas?"**
R: Tomará ~1-2 minutos. El servidor procesa una por una sin problema.

**P: "¿Puedo cambiar el formato de las fechas?"**
R: Sí, edito el código `formatValue()` en excelController.js

**P: "¿Funciona con Google Drive?"**
R: Leerías en local, subirías el Excel. Integración con Drive requeriría cambios.

**P: "¿Qué pasa si se cae el servidor?"**
R: Los archivos temporales se limpian. Usuario debe reintentar.

**P: "¿Puedo agregar más procesamiento?"**
R: Sí, es fácil. La arquitectura modular permite extensiones.
