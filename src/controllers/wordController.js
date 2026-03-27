import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Almacenamiento temporal de plantillas
let templateStorage = {};

export const wordController = {
  extractVariables: async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({ error: 'No se cargó ningún archivo' });
      }

      if (!req.file.originalname.toLowerCase().endsWith('.docx')) {
        if (req.file.path) fs.unlink(req.file.path, () => {});
        return res.status(400).json({ error: 'Formato inválido. Por favor carga un archivo .docx' });
      }

      const content = fs.readFileSync(req.file.path);
      let placeholders = [];

      // Primero intenta extracción raw (sin validación de docxtemplater)
      try {
        const zip = new PizZip(content);
        const docXml = zip.file('word/document.xml');
        if (docXml) {
          const xmlContent = docXml.asText();

          // Word puede fragmentar el texto en múltiples <w:r> runs dentro del XML.
          // Para capturar placeholders como {aseguradoranropoliza} correctamente,
          // eliminamos todas las etiquetas XML y buscamos en el texto plano resultante.
          const plainText = xmlContent.replace(/<[^>]+>/g, '');

          // Regex: solo letras, números, guiones bajos y letras acentuadas (sin espacios internos)
          const regex = /\{\s*([a-zA-Z0-9_áéíóúñÁÉÍÓÚÑ]+)\s*\}/g;
          const matches = plainText.match(regex) || [];
          placeholders = [...new Set(matches.map(m => m.replace(/[{}]/g, '').trim()))].filter(p => p.length > 0);

          console.log('✓ Variables extraídas:', placeholders);
        }
      } catch (xmlErr) {
        console.warn('Error extrayendo XML, intentando búsqueda en buffer:', xmlErr.message);

        // Fallback: busca en el buffer completo
        try {
          const bufferStr = content.toString('latin1');
          // Eliminar posibles tags XML residuales del buffer
          const plainBuffer = bufferStr.replace(/<[^>]+>/g, '');
          const regex = /\{([a-zA-Z0-9_áéíóúñÁÉÍÓÚÑ]+)\}/g;
          const matches = plainBuffer.match(regex) || [];
          placeholders = [...new Set(matches.map(m => m.replace(/[{}]/g, '').trim()))].filter(p => p.length > 0);

          console.log('✓ Variables extraídas (en buffer):', placeholders);
        } catch (bufErr) {
          console.error('Todos los métodos de extracción fallaron:', bufErr.message);
          placeholders = [];
        }
      }

      const templateId = `template_${Date.now()}`;
      templateStorage[templateId] = req.file.path;

      return res.json({
        success: true,
        placeholders: placeholders.length > 0 ? placeholders : [],
        totalPlaceholders: placeholders.length,
        filename: req.file.originalname,
        templateId
      });

    } catch (error) {
      console.error('Error en Word:', error.message);
      if (req.file && req.file.path) {
        fs.unlink(req.file.path, () => {});
      }
      return res.status(400).json({ error: 'Error procesando Word: ' + error.message });
    }
  },

  getStoredTemplate: (templateId) => {
    return templateStorage[templateId];
  },

  clearTemplate: (templateId) => {
    if (templateStorage[templateId]) {
      try {
        fs.unlinkSync(templateStorage[templateId]);
      } catch (err) {
        console.error('Error eliminando plantilla:', err);
      }
      delete templateStorage[templateId];
    }
  }
};
