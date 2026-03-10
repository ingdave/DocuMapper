import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { ExcelJS } from 'exceljs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/**
 * Utilidades para procesamiento de documentos
 */
export const docUtils = {
  /**
   * Genera un documento Word con los datos proporcionados
   */
  generateDocument: async (templateBuffer, mappedData) => {
    try {
      const zip = new PizZip(templateBuffer);
      const doc = new Docxtemplater(zip, { linebreaks: true });
      
      doc.setData(mappedData);
      doc.render();
      
      return doc.getZip().generate({ type: 'nodebuffer' });
    } catch (error) {
      throw new Error(`Error generando documento: ${error.message}`);
    }
  },

  /**
   * Mapea datos de Excel con placeholders
   */
  mapDataToPlaceholders: (rowData, mappings) => {
    const mappedData = {};
    
    for (const [placeholder, columnName] of Object.entries(mappings)) {
      mappedData[placeholder] = rowData[columnName] || '';
    }
    
    return mappedData;
  },

  /**
   * Valida que todos los placeholders estén mapeados
   */
  validateMappings: (placeholders, mappings) => {
    const mappedPlaceholders = new Set(Object.keys(mappings));
    const allMapped = placeholders.every(p => mappedPlaceholders.has(p) && mappings[p]);
    
    const unmapped = placeholders.filter(p => !mappings[p]);
    
    return {
      isValid: allMapped,
      unmapped,
      mapped: placeholders.filter(p => mappings[p]).length,
      total: placeholders.length
    };
  },

  /**
   * Genera un nombre de archivo único
   */
  generateFileName: (rowData, prefix = 'documento') => {
    const firstValue = Object.values(rowData)[0] || 'documento';
    const sanitized = String(firstValue).replace(/[^a-zA-Z0-9]/g, '_');
    const timestamp = Date.now();
    
    return `${prefix}_${sanitized}_${timestamp}.docx`;
  }
};
