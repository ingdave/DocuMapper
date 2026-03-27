import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import fs from 'fs';
import path from 'path';
import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import archiver from 'archiver';
import { wordController } from './wordController.js';
import { exec } from 'child_process';
import { promisify } from 'util';
import ConvertApi from 'convertapi';

const execAsync = promisify(exec);

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const downloadsDir = path.join(__dirname, '../../downloads');



// Función para convertir valores a strings
function stringifyValue(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  if (typeof value === 'string') {
    return value;
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  if (value instanceof Date) {
    const day = String(value.getUTCDate()).padStart(2, '0');
    const month = String(value.getUTCMonth() + 1).padStart(2, '0');
    const year = String(value.getUTCFullYear()).slice(-2);
    return `${day}/${month}/${year}`;
  }
  // Para objetos complejos
  if (typeof value === 'object') {
    if (value.text) return String(value.text);
    if (value.result !== undefined) return String(value.result);
    return JSON.stringify(value);
  }
  return String(value);
}

// Función para formatear números con separador de miles (formato colombiano: . miles, , decimales)
function formatWithThousands(num) {
  const hasDecimals = num % 1 !== 0;
  const decimals = hasDecimals ? 2 : 0;
  const parts = Math.abs(num).toFixed(decimals).split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const formatted = parts.length > 1 ? parts[0] + ',' + parts[1] : parts[0];
  return num < 0 ? '-' + formatted : formatted;
}

// Función para formatear valores numéricos según opción del usuario
function formatNumberValue(value, formatType) {
  if (!formatType || formatType === 'raw') {
    return stringifyValue(value);
  }

  const num = typeof value === 'number' ? value : parseFloat(String(value).replace(/[^0-9.\-]/g, ''));
  if (isNaN(num)) {
    return stringifyValue(value);
  }

  switch (formatType) {
    case 'currency':
      return '$ ' + formatWithThousands(num);
    case 'thousands':
      return formatWithThousands(num);
    case 'percentage':
      // Multiplicar por 100 si el valor es decimal (ej: 0.25 -> 25%)
      const percentage = num < 1 && num > -1 ? num * 100 : num;
      return percentage.toFixed(2).replace('.', ',') + '%';
    default:
      return stringifyValue(value);
  }
}

// Meses en español abreviados
const MESES_ES = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'];

function getCurrentMonthSpanish() {
  return MESES_ES[new Date().getMonth()];
}

function buildFilename(row, filenameColumn, index, month, initialName) {
  const mes = month || getCurrentMonthSpanish();
  const prefix = initialName ? initialName.trim() : 'Informe_supervision';
  let contratista = '';
  
  if (filenameColumn && row[filenameColumn]) {
    contratista = String(row[filenameColumn]).trim();
  } else {
    // Fallback: usar el primer valor de la fila
    contratista = String(Object.values(row)[0] || `doc_${index}`).trim();
  }
  
  // Sanitizar: reemplazar espacios con _ y quitar caracteres especiales
  const sanitized = contratista
    .replace(/\s+/g, '_')
    .replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ_]/g, '')
    .substring(0, 50);
  
  return `${prefix}_${mes}_${sanitized}.docx`;
}

export const mappingController = {


  generateDocuments: async (req, res) => {
    try {
      const { templateId, mappings, excelData, filenameColumn, month, initialName, formatOptions } = req.body;

      if (!templateId || !mappings || !excelData) {
        return res.status(400).json({ error: 'Campos obligatorios faltantes' });
      }

      const templatePath = wordController.getStoredTemplate(templateId);
      if (!templatePath || !fs.existsSync(templatePath)) {
        return res.status(400).json({ error: 'Plantilla no encontrada o expirada' });
      }

      const generatedFiles = [];
      let successCount = 0;
      let errorCount = 0;

      // Leer la plantilla UNA SOLA VEZ antes del loop (optimización I/O)
      const templateBuffer = fs.readFileSync(templatePath);
      const generationStart = Date.now();

      // Pre-compilar expresiones regulares para el fallback (optimización CPU)
      const placeholderRegexes = {};
      for (const placeholder of Object.keys(mappings)) {
        placeholderRegexes[placeholder] = [
          new RegExp(`\\{\\s*${placeholder.trim()}\\s*\\}`, 'gi'),
          new RegExp(`\\{\\s*${placeholder.trim().replace(/\\s+/g, '\\\\s+')}\\s*\\}`, 'gi'),
          new RegExp(`\\{${placeholder.trim()}\\}`, 'gi')
        ];
      }

      // Generar un documento por cada fila de datos
      for (let i = 0; i < excelData.length; i++) {
        const rowStart = Date.now();
        try {
          const row = excelData[i];
          
          // Mapear datos con placeholders - convertir a strings
          const mappedData = {};
          for (const [placeholder, excelColumn] of Object.entries(mappings)) {
            const format = formatOptions?.[placeholder] || 'raw';
            mappedData[placeholder] = formatNumberValue(row[excelColumn], format);
          }

          // Instanciar PizZip desde el buffer en memoria (sin I/O de disco por fila)
          const zip = new PizZip(templateBuffer);
          
          let buffer;
          let filename;
          let filepath;
          
          try {
            const doc = new Docxtemplater(zip, { 
              linebreaks: true,
              nullGetter() {
                return "";
              }
            });
            doc.setData(mappedData);
            doc.render();
            buffer = doc.getZip().generate({ type: 'nodebuffer' });
            
            // Nombre con formato PREFIJO_MES_contratista
            filename = buildFilename(row, filenameColumn, i + 1, month, initialName);
            filepath = path.join(downloadsDir, filename);
            
            await fs.promises.writeFile(filepath, buffer);
            generatedFiles.push({ filename, index: i + 1 });
            successCount++;
            console.log(`✓ Doc ${i + 1}/${excelData.length} generado en ${Date.now() - rowStart}ms`);
            
          } catch (docError) {
            console.warn(`Fila ${i + 1} - Docxtemplater falló, usando alternativa:`, docError.message);
            
            // Fallback: Reemplazar manualmente en el XML
            try {
              const docXml = zip.file('word/document.xml');
              let xmlContent = docXml.asText();
              let replacementsDone = {};
              
              // Reemplazar cada placeholder
              for (const [placeholder, value] of Object.entries(mappedData)) {
                const stringValue = String(value).trim();
                let foundMatch = false;
                
                const patterns = placeholderRegexes[placeholder] || [];
                
                for (const regex of patterns) {
                  const before = xmlContent;
                  xmlContent = xmlContent.replace(regex, stringValue);
                  if (xmlContent !== before) {
                    foundMatch = true;
                    replacementsDone[placeholder] = true;
                    console.log(`✓ Reemplazado {{${placeholder}}} con "${stringValue}"`);
                    break;
                  }
                }
                
                if (!foundMatch) {
                  console.warn(`No se pudo encontrar {{${placeholder}}} en el XML`);
                }
              }
              
              // Limpieza final: eliminar cualquier placeholder {VARIABLE} que haya quedado sin mapear
              const beforeCleanup = xmlContent;
              xmlContent = xmlContent.replace(/\{\s*[a-zA-Z0-9_áéíóúñÁÉÍÓÚÑ]+\s*\}/g, "");
              if (xmlContent !== beforeCleanup) {
                console.log("✓ Limpieza final: Se eliminaron placeholders huérfanos del XML");
              }
              
              zip.file('word/document.xml', xmlContent);
              buffer = zip.generate({ type: 'nodebuffer' });
              
              console.log(`Fila ${i + 1} resumen: ${Object.keys(replacementsDone).length}/${Object.keys(mappedData).length} placeholders reemplazados`);
              
              filename = buildFilename(row, filenameColumn, i + 1, month, initialName);
              filepath = path.join(downloadsDir, filename);
              
              await fs.promises.writeFile(filepath, buffer);
              generatedFiles.push({ filename, index: i + 1, usedFallback: true });
              successCount++;
              console.log(`✓ Documento ${i + 1} generado con método alternativo`);
              
            } catch (fallbackErr) {
              console.error(`Error en alternativa para fila ${i + 1}:`, fallbackErr.message);
              generatedFiles.push({ 
                filename: null, 
                error: fallbackErr.message, 
                index: i + 1 
              });
              errorCount++;
            }
          }
        } catch (err) {
          console.error(`Error crítico generando documento ${i + 1}:`, err);
          generatedFiles.push({ 
            filename: null, 
            error: err.message, 
            index: i + 1 
          });
          errorCount++;
        }
      }

      // Limpiar plantilla si se generaron todos exitosamente
      if (errorCount === 0) {
        wordController.clearTemplate(templateId);
      }

      const validFiles = generatedFiles
        .filter(f => f.filename && !f.error)
        .map(f => f.filename);

      res.json({
        success: true,
        totalGenerated: successCount,
        totalAttempted: excelData.length,
        totalErrors: errorCount,
        files: generatedFiles.map(f => ({
          filename: f.filename,
          downloadUrl: f.filename ? `/api/download/${f.filename}` : null,
          error: f.error,
          index: f.index
        })),
        allFilenames: validFiles,
        downloadZipDocxUrl: validFiles.length > 0 ? '/api/mapping/download-all-docx' : null
      });

    } catch (error) {
      console.error('Document generation error:', error);
      res.status(500).json({ error: error.message });
    }
  },

  downloadAllDocx: async (req, res) => {
    try {
      const { filenames } = req.body;

      if (!filenames || !Array.isArray(filenames) || filenames.length === 0) {
        return res.status(400).json({ error: 'No se especificaron archivos' });
      }

      // Validar que todos los archivos existan y sean .docx
      const validFiles = filenames.filter(f => {
        const filepath = path.join(downloadsDir, f);
        return fs.existsSync(filepath) && f.endsWith('.docx');
      });

      if (validFiles.length === 0) {
        return res.status(400).json({ error: 'No se encontraron archivos .docx válidos' });
      }

      const zipFilename = `documentos_${Date.now()}.zip`;
      const zipPath = path.join(downloadsDir, zipFilename);

      const output = fs.createWriteStream(zipPath);
      const archive = archiver('zip', { zlib: { level: 1 } });

      output.on('close', () => {
        console.log(`✓ ZIP creado: ${zipFilename} (${archive.pointer()} bytes)`);
        res.json({
          success: true,
          downloadUrl: `/api/download/${zipFilename}`,
          message: `${validFiles.length} documents packaged`,
          zipSize: archive.pointer()
        });
      });

      archive.on('error', (err) => {
        console.error('Error del archivo:', err);
        res.status(500).json({ error: 'Error creando ZIP: ' + err.message });
      });

      archive.pipe(output);

      // Agregar archivos al ZIP
      validFiles.forEach(filename => {
        const filepath = path.join(downloadsDir, filename);
        archive.file(filepath, { name: filename });
      });

      archive.finalize();

    } catch (error) {
      console.error('Error de descarga:', error);
      res.status(500).json({ error: error.message });
    }
  },

  downloadAllPdf: async (req, res) => {
    try {
      const { filenames } = req.body;

      if (!filenames || !Array.isArray(filenames) || filenames.length === 0) {
        return res.status(400).json({ error: 'No se especificaron archivos' });
      }

      // Validar archivos DOCX existentes
      const validDocx = filenames.filter(f => {
        const filepath = path.join(downloadsDir, f);
        return fs.existsSync(filepath) && f.endsWith('.docx');
      });

      if (validDocx.length === 0) {
        return res.status(400).json({ error: 'No se encontraron archivos DOCX para convertir' });
      }

      const pdfFiles = [];
      const secret = process.env.CONVERT_API_SECRET || 'YOUR_SECRET_HERE';
      const convertapi = ConvertApi(secret);
      
      console.log(`Iniciando conversión a PDF para ${validDocx.length} archivos en paralelo:`, validDocx);
      
      if (secret === 'YOUR_SECRET_HERE' || !secret) {
        return res.status(500).json({ error: 'API Secret de ConvertAPI no configurado.' });
      }

      // Convertir archivos DOCX a PDF en paralelo usando Promise.all
      const conversionPromises = validDocx.map(async (docx) => {
        const inputPath = path.join(downloadsDir, docx);
        const pdfFilename = docx.replace('.docx', '.pdf');
        const outputPath = path.join(downloadsDir, pdfFilename);

        try {
          console.log(`Enviando a ConvertAPI: ${docx}...`);

          const result = await convertapi.convert('pdf', { File: inputPath }, 'docx');
          console.log(`✓ Recibido de ConvertAPI: ${docx}. Guardando...`);
          
          // Guardar archivos
          const files = await result.saveFiles(downloadsDir);
          
          if (fs.existsSync(outputPath)) {
            console.log(`✓ Validado: ${pdfFilename} existe.`);
            return pdfFilename;
          } else {
            console.warn(`⚠ Advertencia: El archivo ${outputPath} no existe después de saveFiles.`);
            // Intentar buscar si se guardó con otro nombre
            const savedFile = files[0];
            if (savedFile && fs.existsSync(savedFile)) {
               const actualName = path.basename(savedFile);
               if (actualName !== pdfFilename) {
                 console.log(`Cambiando nombre de ${actualName} a ${pdfFilename}`);
                 await fs.promises.rename(savedFile, outputPath);
                 return pdfFilename;
               }
            }
            return null;
          }
        } catch (convErr) {
          console.error(`✘ Error convirtiendo ${docx}:`, convErr.message);
          return null;
        }
      });

      const conversionResults = await Promise.all(conversionPromises);
      
      // Filtrar resultados válidos (no nulos)
      for (const res of conversionResults) {
        if (res) pdfFiles.push(res);
      }

      if (pdfFiles.length === 0) {
        return res.status(500).json({ 
          error: 'No se pudo convertir ningún archivo a PDF. Verifica tu conexión o el Secret Key de ConvertAPI.' 
        });
      }

      const zipFilename = `documentos_pdf_${Date.now()}.zip`;
      const zipPath = path.join(downloadsDir, zipFilename);

      const output = fs.createWriteStream(zipPath);
      const archive = archiver('zip', { zlib: { level: 1 } });

      output.on('close', () => {
        console.log(`✓ ZIP PDF (Nube) creado: ${zipFilename}`);
        res.json({
          success: true,
          downloadUrl: `/api/download/${zipFilename}`,
          message: `${pdfFiles.length} PDF documents packaged via Cloud`
        });
      });

      archive.pipe(output);
      pdfFiles.forEach(f => {
        archive.file(path.join(downloadsDir, f), { name: f });
      });
      archive.finalize();

    } catch (error) {
      console.error('Error generando PDF:', error);
      res.status(500).json({ error: error.message });
    }
  }
};
