# 📘 Manual de Usuario - docuMapper

## Tabla de Contenidos
1. [Introducción](#introducción)
2. [Requisitos del Sistema](#requisitos-del-sistema)
3. [Instalación y Configuración](#instalación-y-configuración)
4. [Interfaz Principal](#interfaz-principal)
5. [Paso 1: Cargar Excel](#paso-1-cargar-excel)
6. [Paso 2: Cargar Plantilla Word](#paso-2-cargar-plantilla-word)
7. [Paso 3: Mapear Campos](#paso-3-mapear-campos)
8. [Paso 4: Configurar Generación](#paso-4-configurar-generación)
9. [Paso 5: Descargar Documentos](#paso-5-descargar-documentos)
10. [Preguntas Frecuentes](#preguntas-frecuentes)
11. [Solución de Problemas](#solución-de-problemas)
12. [Consejos y Mejores Prácticas](#consejos-y-mejores-prácticas)

---

## Introducción

**docuMapper** es una herramienta que automatiza la generación masiva de documentos Word basada en datos de Excel. 

### ¿Para qué sirve?

- 📊 **Automatizar reportes**: Genera reportes individuales para cada fila de Excel
- 💼 **Documentos personalizados**: Cada documento se llena con los datos correspondientes
- ⚡ **Ahorra tiempo**: En lugar de escribir 100 documentos uno por uno
- 🎨 **Conserva formato**: Los estilos y formato de la plantilla Word se mantienen
- 📈 **Exporta múltiples formatos**: DOCX y PDF

### Ejemplo práctico:

Tienes un Excel con 50 vendedores y quieres generar 50 certificados de venta (uno por vendedor) usando una plantilla Word. En lugar de llenar cada uno manualmente, docuMapper:
1. Lee los datos del Excel
2. Lee la plantilla Word con espacios como `{nombre}`, `{monto}`, etc.
3. Reemplaza automáticamente cada campo
4. Genera 50 documentos en segundos

---

## Requisitos del Sistema

### Mínimos:
- **Windows 7** o superior (o macOS/Linux con Node.js)
- **2 GB RAM**
- **500 MB espacio libre** en disco

### Software requerido:
- **Node.js 16+** ([Descargar](https://nodejs.org/))
- **navegador web moderno** (Chrome, Firefox, Edge)

### Archivos necesarios:
- ✅ Archivo Excel (.xlsx o .xls)
- ✅ Plantilla Word (.docx)

---

## Instalación y Configuración

### Opción 1: Instalación desde terminal

1. **Abre PowerShell** o Terminal en la carpeta del proyecto:
   ```powershell
   cd C:\Users\TuUsuario\Dev\docuMapper
   ```

2. **Instala dependencias**:
   ```powershell
   npm install
   ```

3. **Inicia el servidor**:
   ```powershell
   npm start
   ```

4. **Abre tu navegador**:
   - Ve a: `http://localhost:3000`
   - Verás la interfaz de carga

### Opción 2: Instalación desde Visual Studio Code

1. Abre VS Code
2. Abre la carpeta `docuMapper`
3. Abre la terminal integrada (Ctrl + `)
4. Ejecuta: `npm install && npm start`

---

## Interfaz Principal

```
┌─────────────────────────────────────────┐
│          docuMapper - Paso 1/5           │
│                                         │
│  ☑ Paso 1: Cargar Excel                │
│  ○ Paso 2: Cargar Plantilla Word       │
│  ○ Paso 3: Mapear Campos               │
│  ○ Paso 4: Configurar Generación       │
│  ○ Paso 5: Descargar Documentos        │
│                                         │
│  [Interfaz del paso actual]             │
│                                         │
│  [← Anterior]              [Siguiente →] │
└─────────────────────────────────────────┘
```

---

## Paso 1: Cargar Excel

### ¿Qué es este paso?
Cargas el archivo Excel que contiene los datos que se van a usar en los documentos.

### Requisitos del Excel:
✅ **Formato**: .xlsx o .xls
✅ **Estructura**: 
- Primera fila = nombres de columnas (headers)
- Filas siguientes = datos

### Ejemplo de Excel correcto:

| Nombre | Monto | Fecha | Porcentaje |
|--------|-------|-------|-----------|
| Juan Pérez | 50000 | 2026-03-27 | 0.15 |
| María Gómez | 75000 | 2026-03-27 | 0.20 |
| Carlos López | 100000 | 2026-03-27 | 0.25 |

### Cómo subir el Excel:

**Opción 1: Drag & Drop**
1. Haz clic en el área gris de carga
2. Arrastra tu archivo Excel hacia allá
3. Suéltalo

**Opción 2: Seleccionar manualmente**
1. Haz clic en el área gris
2. Se abre un navegador de archivos
3. Selecciona tu archivo Excel

### ¿Qué detecta docuMapper?
- ✅ Nombres de columnas
- ✅ Número de filas
- ✅ **Formatos de celdas** (Moneda, Porcentaje, Miles)
- ✅ Tipos de datos

### Ejemplo de detección de formatos:

Si en tu Excel tienes:
- Columna "Monto" con formato **Moneda ($)** → docuMapper lo detecta
- Columna "Comisión" con formato **Porcentaje (%)** → docuMapper lo detecta
- Columna "Cantidad" con formato **Miles con separador** → docuMapper lo detecta

**Resultado después de cargar**:
```
✓ 3 filas cargadas correctamente (4 columnas)
↓
Avanza al Paso 2
```

---

## Paso 2: Cargar Plantilla Word

### ¿Qué es este paso?
Cargas el archivo Word (.docx) que será la plantilla para generar todos los documentos.

### Requisitos de la Plantilla:
✅ **Formato**: .docx (Word 2007+)
✅ **Placeholders**: Variables en formato `{nombre_variable}`

### Cómo crear placeholders en Word:

1. Abre tu documento en Microsoft Word
2. Escribe normalmente el contenido
3. Donde quieras un valor dinámico, escribe entre llaves:
   ```
   Certifico que {nombre} ha completado el curso
   y obtuvo una calificación de {calificacion}
   ```

4. Guarda como `.docx`

### Ejemplo de plantilla:

```
=====================================
CERTIFICADO DE VENTA
=====================================

Fecha: {fecha}

Vendedor: {nombre}
Empresa: {empresa}
Región: {region}

Monto Total: {monto}
Comisión: {comision}

Porcentaje de Venta: {porcentaje}

Autorizado por: _________________
```

### Características de los placeholders:

- ✅ Pueden estar en cualquier parte del documento
- ✅ Múltiples en el mismo párrafo: `{nombre} vendió {cantidad} unidades por ${monto}`
- ✅ Insensibles a mayúsculas/minúsculas: `{Nombre}`, `{nombre}`, `{NOMBRE}` funcionan igual
- ✅ Permiten espacios: `{ nombre }` es válido
- ✅ Caracteres especiales: soportan acentos y ñ

### Cómo subir la Plantilla:

1. Haz clic en el área de carga
2. Selecciona tu archivo `.docx`
3. docuMapper extraerá automáticamente los placeholders

**Resultado después de cargar**:
```
✓ 4 placeholders encontrados:
  - {nombre}
  - {monto}
  - {comision}
  - {porcentaje}
↓
Avanza al Paso 3
```

---

## Paso 3: Mapear Campos

### ¿Qué es este paso?
**Mapeo** = conectar cada placeholder de Word con su columna correspondiente en Excel.

### Interfaz de Mapeo:

```
┌─────────────────────────────────────────────────┐
│ Placeholder: {nombre}  │  Columna Excel: [Nombre ▼] │  Formato: [Sin formato ▼] │
│ Placeholder: {monto}   │  Columna Excel: [Monto ▼]  │  Formato: [Moneda ($) ▼]  │
│ Placeholder: {comision}│  Columna Excel: [Monto ▼]  │  Formato: [% ▼]          │
│ Placeholder: {fecha}   │  Columna Excel: [Fecha ▼]  │  Formato: [Sin formato ▼] │
└─────────────────────────────────────────────────┘
Mapeos completados: 4/4 ✓
```

### Cómo mapear:

**Para cada placeholder:**

1. **Selecciona la columna Excel** que contiene los datos:
   - Haz clic en `[Selecciona columna Excel]`
   - Se despliega una lista con los nombres de columnas
   - Elige la que corresponda

2. **Selecciona el formato** (aparece automáticamente):
   - Si la columna es numérica, aparecen opciones de formato
   - Elige lo que necesites (moneda, porcentaje, etc.)

### Opciones de Formato:

| Formato | Ejemplo | Cuándo usar |
|---------|---------|-----------|
| **Sin formato** | `1500.5` | Números simples, IDs, códigos |
| **Moneda ($)** | `$ 1.500,50` | Dinero, precios, montos |
| **Con separador** | `1.500,50` | Cantidades grandes sin símbolo |
| **Porcentaje (%)** | `25,00%` | Tasas, comisiones, descuentos |

### Detección Automática:

✨ **docuMapper detecta automáticamente:**
- Si una columna tiene formato de Moneda en Excel → sugiere "Moneda ($)"
- Si tiene Porcentaje → sugiere "Porcentaje (%)"
- Si tiene Miles → sugiere "Con separador"

**Solo debes confirmar o cambiar si lo necesitas.**

### Vista Previa:

Antes de continuar, puedes ver cómo quedará:

1. Haz clic en **"Vista Previa"**
2. Se abre una ventana mostrando los datos del primer registro con los formatos aplicados
3. Verifica que todo se vea correcto

### Ejemplo de Vista Previa:

```
Certifico que Juan Pérez
obtuvo una calificación de 95,50%

Monto: $ 50.000,00
Comisión: $ 7.500,00
```

**Validaciones:**
- ❌ No puedes avanzar si falta mapear algún placeholder
- ❌ No puedes dejar campos en blanco

---

## Paso 4: Configurar Generación

### ¿Qué es este paso?
Configuras opciones finales antes de generar los documentos masivamente.

### Opciones disponibles:

#### 1. **Prefijo del nombre de archivo**
```
Ejemplo: "Informe_supervision"
Resultado: Informe_supervision_MAR_Juan_Pérez.docx
```

Usa esto para agrupar documentos por tipo (Reportes, Certificados, Facturas, etc.)

#### 2. **Mes**
```
Opciones: ENE, FEB, MAR, ABR, MAY, JUN, JUL, AGO, SEP, OCT, NOV, DIC
```

Se incluye en el nombre del archivo para identificar por período.

#### 3. **Columna para nombre**
```
Selecciona: ¿Cuál columna usa para el nombre individual?
Ejemplo: Si seleccionas "Nombre" → Informe_supervision_MAR_Juan_Pérez.docx
```

Esto personaliza cada archivo con el nombre del registro (vendedor, cliente, etc.)

### Patrón de nombres:

```
[Prefijo]_[Mes]_[Valor de Columna].docx

Ejemplos:
- Certificado_MAR_Juan_Pérez.docx
- Reportes_MAR_María_López.docx
- Facturas_ABR_Tech_Solutions.docx
```

### Resumen antes de generar:

```
📊 Vas a generar:
   • 3 documentos
   • Basados en: plantilla.docx
   • Con 4 campos mapeados
   • Formato de nombres: Certificado_MAR_[nombre]

[Generar Documentos]
```

---

## Paso 5: Descargar Documentos

### ¿Qué sucede aquí?
Se muestran los documentos generados y tienes opciones para descargarlos.

### Estados de Generación:

```
✓ 3 documentos generados exitosamente
  - Certificado_MAR_Juan_Pérez.docx    [Descargar]
  - Certificado_MAR_María_López.docx   [Descargar]
  - Certificado_MAR_Carlos_López.docx  [Descargar]
```

### Opciones de Descarga:

#### 1. **Descargar individual**
- Haz clic en `[Descargar]` junto a cada documento
- Se descarga un archivo a tu carpeta de Descargas

#### 2. **Zip de DOCX**
- Se comprimen todos los documentos en un archivo .zip
- Necesario para descargarlos todos a la vez

#### 3. **Zip de PDF**
- Convierte cada DOCX a PDF
- Luego los comprime en un .zip
- Útil si necesitas distribuir en formato PDF

### Flujo de Descarga:

```
[Generar DOCX]
    ↓
[Vista de documentos generados]
    ↓
[Selecciona opción de descarga]
    ├─ Descargar individual
    ├─ Descargar ZIP DOCX
    └─ Convertir a PDF y descargar ZIP
```

### Tiempo de descarga:

- **DOCX individual**: < 5 segundos
- **ZIP de 10 DOCX**: < 10 segundos
- **Conversión a PDF**: 10-30 segundos por documento

### Ubicación de descargas:

Los archivos se descargan a tu carpeta estándar de descargas:
- Windows: `C:\Users\[TuUsuario]\Downloads\`
- macOS: `/Users/[TuUsuario]/Downloads/`
- Linux: `~/Downloads/`

### Limpieza automática:

⏰ Los archivos en el servidor se borran automáticamente después de **15 minutos** si no los descargas.

---

## Preguntas Frecuentes

### P: ¿Puedo generar PDF directamente?
**R:** No, primero genera DOCX, luego puedes convertir a PDF en el Paso 5. Los DOCX se preservan para esta opción.

### P: ¿Cuántos documentos puedo generar?
**R:** Teóricamente ilimitado, pero se recomienda:
- Bueno: 10-100 documentos
- Aceptable: 100-500 documentos
- Considerar dividir: 500+ documentos

### P: ¿Qué pasa si tengo un placeholder que no mapeo?
**R:** El sistema no te deja avanzar. Debes mapear todos los placeholders de la plantilla Word.

### P: ¿Puedo cambiar la plantilla Word a mitad del proceso?
**R:** Sí, regresa a Paso 2 y carga una nueva plantilla. Se reiniciará el mapeo correspondiente.

### P: ¿Qué pasa si una celda Excel está vacía?
**R:** El placeholder correspondiente quedará en blanco en el documento.

### P: ¿Soporta Excel con múltiples hojas?
**R:** Solo usa la **primera hoja** del libro. Si tus datos están en otra hoja, copia a la primera antes de cargar.

### P: ¿Puedo usar columnas con fórmulas?
**R:** Sí, Excel calcula la fórmula y docuMapper usa el resultado. Se detectan correctamente los formatos.

### P: ¿Qué formatos de Excel se detectan?
**R:**
- ✅ Moneda ($, €, £, ¥)
- ✅ Porcentaje (%)
- ✅ Miles (#,##0)
- ✅ Fechas
- ✅ Números normales

### P: ¿Puedo usar caracteres especiales en placeholders?
**R:** No, solo letras, números, guiones bajos y acentos: `{nombre_completo}`, `{monto_venta}`

### P: ¿Cuánto espacio ocupa cada documento?
**R:** Depende de la plantilla, típicamente 50-200 KB por DOCX.

---

## Solución de Problemas

### ❌ Problema: "No se cargó ningún archivo"

**Causas:**
- El archivo no es Excel válido
- Formato no soportado (.csv, .xls antiguo)

**Solución:**
1. Abre el archivo en Excel
2. Guarda como "Excel Workbook (.xlsx)"
3. Intenta cargar de nuevo

---

### ❌ Problema: "El Excel no tiene encabezados"

**Causas:**
- La primera fila está vacía
- Los datos comienzan en una fila diferente

**Solución:**
```
❌ Incorrecto:
   [Fila 1: vacía]
   [Fila 2: nombres]
   [Fila 3: datos]

✅ Correcto:
   [Fila 1: nombres]
   [Fila 2: datos]
   [Fila 3: datos]
```

---

### ❌ Problema: "No se encontraron placeholders en Word"

**Causas:**
- El archivo no es Word válido (.docx)
- Los placeholders tienen formato incorrecto

**Solución:**

Verifica los placeholders en tu plantilla:
```
❌ Incorrecto:
   - $nombre$
   - [nombre]
   - <nombre>
   - nombre

✅ Correcto:
   - {nombre}
   - {nombre_completo}
   - { nombre }  (con espacios es OK)
```

---

### ❌ Problema: "No se puede encontrar la columna en Excel"

**Causas:**
- El nombre de la columna en Excel tiene espacios o caracteres diferentes
- Mayúsculas/minúsculas diferentes

**Solución:**

Verifica que el nombre en Excel coincida exactamente (incluye espacios):

```
Si tu columna se llama "Nombre Completo" en Excel,
selecciona exactamente "Nombre Completo" en el mapeo
(no "Nombre", ni "nombre_completo")
```

---

### ❌ Problema: Los formatos de Excel no aparecen en Paso 3

**Causas:**
- La columna no tiene datos en las primeras 5 filas
- El formato no fue guardado correctamente

**Solución:**
1. Asegúrate de que la columna tenga valores en al menos las primeras 5 filas
2. Guarda el Excel y recarga

---

### ❌ Problema: Los documentos generados están vacíos

**Causas:**
- Los placeholders en Word no coinciden exactamente
- Formato incorrecto de placeholders

**Solución:**
1. Revisa los nombres en el mapeo (Paso 3)
2. Verifica que en Word estén escritos exactamente igual
3. Elimina espacios innecesarios

---

### ❌ Problema: "Error al generar documentos"

**Causas:**
- Falta espacio en disco
- Error de conexión con el servidor
- Archivo Word corrupto

**Solución:**
1. Reinicia el servidor (Ctrl+C en terminal, luego `npm start`)
2. Intenta con números menores de documentos
3. Verifica que la plantilla Word se abre correctamente

---

### ❌ Problema: Los PDFs no se descargan

**Causas:**
- API de ConvertAPI no está configurada
- Los DOCX fueron eliminados antes de convertir

**Solución:**
1. Genera nuevamente los DOCX
2. Descarga los PDF inmediatamente después
3. Los archivos se borran después de 15 minutos

---

## Consejos y Mejores Prácticas

### 📋 Preparación de Excel

**1. Estructura clara:**
```
✅ BUENO:
   | Nombre | Monto | Fecha |
   |--------|-------|-------|
   | Juan   | 1000  | fecha |

❌ MALO:
   | NOMBRE | monto | FECHA |
   |--------|-------|-------|
   |        | 1000  | fecha |  (Primera fila parcial)
```

**2. Formatos de datos:**
- Moneda: Aplica formato `$` en Excel
- Porcentaje: Aplica formato `%` en Excel
- Fechas: Usa formato `DD/MM/YYYY`

**3. Evita:**
- ❌ Espacios en blanco al inicio/final de celdas
- ❌ Saltos de línea dentro de celdas
- ❌ Caracteres especiales (excepto acentos, ñ)

---

### 📄 Creación de Plantillas Word

**1. Estructura clara:**
```
✅ BUENO:
   Certifico que {nombre} ha vendido {cantidad} unidades
   por un monto de {monto}

❌ MALO:
   cert... que {nombre}, quien vendio {cantidad} unidades
   por monto de {monto}  (espacios/acentos inconsistentes)
```

**2. Nombres de placeholders:**
- Usa nombres descriptivos: `{nombre_cliente}`, `{monto_total}`
- Evita: `{x}`, `{dato1}`, `{var}`
- Mantén consistencia: `{nombre}` no `{Nombre}`

**3. Formato y estilos:**
- Los estilos Word se conservan
- Los placeholders heredan el formato del texto
- Usa etiquetas de título, párrafo, tabla normalmente

**4. Placeholders en tablas:**
```
✅ Funciona:
   | Concepto      | Valor      |
   |-------|-----------|
   | Total | {monto}   |

❌ Evita:
   | Concepto      | Valor      |
   |-------|-----------|
   | Tot{al} | {mont}o   |  (fragmentado)
```

---

### ⚙️ Optimización

**Para generar rápidamente:**
1. Mantén Excel con datos limpios
2. Evita archivos muy grandes (>10MB)
3. Genera en lotes pequeños (50-100) si es posible
4. Ten suficiente espacio en disco

**Para mejor calidad:**
1. Prueba primero con 1-2 documentos
2. Revisa el formato en el PDF antes de mass-producir
3. Mantén backups de tu plantilla

---

### 🔄 Flujo Recomendado

```
1. ✓ Prepara Excel (datos limpios)
   ↓
2. ✓ Crea plantilla Word (placeholders claros)
   ↓
3. ✓ Prueba con 1 fila (Paso 1-5)
   ↓
4. ✓ Verifica el documento generado
   ↓
5. ✓ Si todo OK, regresa a Paso 1 y carga completo
   ↓
6. ✓ Genera masivamente
```

---

### 📊 Límites Recomendados

| Cantidad | Tiempo Estimado | Recomendación |
|----------|-----------------|---------------|
| 1-10 | < 1 min | Prueba |
| 10-50 | 1-3 min | Pequeño lote |
| 50-100 | 3-5 min | Lote normal |
| 100-500 | 5-15 min | Lote grande |
| 500+ | 15+ min | Considerar dividir |

---

### 🎯 Casos de Uso Comunes

#### Certificados de Capacitación
```
Excel: | Nombres | Fecha | Curso | Calificación |
Word:  Certifico que {nombre} completó {curso}
       con calificación de {calificacion} el {fecha}
```

#### Facturas Masivas
```
Excel: | Cliente | Monto | IVA | Total |
Word:  Factura a {cliente}
       Subtotal: {monto}
       IVA: {iva}
       Total: {total}
```

#### Órdenes de Compra
```
Excel: | Proveedor | Cantidad | Producto | Precio |
Word:  Orden de Compra
       Proveedor: {proveedor}
       Producto: {producto}
       Cantidad: {cantidad}
       Total: {precio}
```

#### Reportes de Ventas
```
Excel: | Vendedor | Mes | Ventas | Comisión | Meta |
Word:  Reporte de {vendedor} - {mes}
       Ventas: {ventas}
       Comisión: {comision}
       Meta alcanzada: {meta_pct}%
```

---

## Contacto y Soporte

Si experimentas problemas no cubiertos en este manual:

1. **Revisa los logs** en la terminal donde corre `npm start`
2. **Verifica la configuración** de archivos Excel/Word
3. **Reinicia el servidor**: Ctrl+C y `npm start` nuevamente
4. **Limpia navegador**: Ctrl+Shift+Delete (caché)

---

**Versión**: 1.0  
**Última actualización**: Marzo 2026  
**Licencia**: Uso interno
