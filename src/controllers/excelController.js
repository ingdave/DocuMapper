import ExcelJS from 'exceljs';
import fs from 'fs';

// Función para formatear valores según su tipo
function formatValue(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  
  // Si es una fecha - usar UTC para evitar problemas de zona horaria
  if (value instanceof Date) {
    const day = String(value.getUTCDate()).padStart(2, '0');
    const month = String(value.getUTCMonth() + 1).padStart(2, '0');
    const year = String(value.getUTCFullYear()).slice(-2);
    return `${day}/${month}/${year}`;
  }
  
  // Si es un objeto (que no sea Date), intenta JSON.stringify
  if (typeof value === 'object') {
    // Si tiene propiedad 'text', úsala
    if (value.text) return String(value.text);
    // Si tiene propiedad 'result', procesarla recursivamente para preservar tipo numérico
    if (value.result !== undefined) return formatValue(value.result);
    return String(value);
  }
  
  // Para números, mantener como número para permitir formateo posterior
  if (typeof value === 'number') return value;
  
  return String(value);
}

// Detectar tipo de formato desde el numFmt de Excel
function detectFormatType(numFmt) {
  if (!numFmt || numFmt === 'General') return null;
  const fmt = String(numFmt);
  // Porcentaje: contiene %
  if (/%/.test(fmt)) {
    return 'percentage';
  }
  // Moneda: símbolos $, €, £, ¥ o formatos contables con _-$
  if (/[$€£¥]/.test(fmt) || /_-.*\$/.test(fmt) || /"COP"|"USD"|"EUR"/.test(fmt)) {
    return 'currency';
  }
  // Formato contable (empieza con _- y tiene #)
  if (/^_/.test(fmt) && /#/.test(fmt)) {
    return 'currency';
  }
  // Separador de miles sin símbolo de moneda: #,##0
  if (/#,##0/.test(fmt)) {
    return 'thousands';
  }
  return null;
}

export const excelController = {
  parseExcel: async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({ error: 'No se cargó ningún archivo' });
      }

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(req.file.path);
      
      const worksheet = workbook.getWorksheet(1);
      if (!worksheet) {
        return res.status(400).json({ error: 'No se encontró hoja de cálculo' });
      }

      // Obtener headers CON su número de columna real (eachCell salta celdas vacías)
      const headerMap = [];
      worksheet.getRow(1).eachCell((cell, colNumber) => {
        if (cell.value) {
          headerMap.push({ name: cell.value.toString(), colNumber });
        }
      });

      const headers = headerMap.map(h => h.name);

      // Detectar formatos de columna desde las primeras filas de datos
      const columnFormats = {};
      const maxScanRows = Math.min(worksheet.rowCount, 6); // escanear hasta 5 filas de datos
      for (let rowNum = 2; rowNum <= maxScanRows; rowNum++) {
        const row = worksheet.getRow(rowNum);
        headerMap.forEach(({ name, colNumber }) => {
          if (columnFormats[name]) return; // ya detectado
          const cell = row.getCell(colNumber);
          if (cell.value !== null && cell.value !== undefined && cell.value !== '') {
            const numFmt = cell.numFmt || (cell.style && cell.style.numFmt);
            const detectedFormat = detectFormatType(numFmt);
            if (detectedFormat) {
              columnFormats[name] = detectedFormat;
            }
          }
        });
      }

      // Obtener datos usando los números de columna reales
      const data = [];
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header
        
        const rowData = {};
        headerMap.forEach(({ name, colNumber }) => {
          const cellValue = row.getCell(colNumber).value;
          rowData[name] = formatValue(cellValue);
        });
        
        if (Object.values(rowData).some(val => val !== '')) {
          data.push(rowData);
        }
      });

      res.json({
        success: true,
        headers,
        data,
        columnFormats,
        totalRows: data.length,
        filename: req.file.originalname
      });

      // Limpiar archivo temporal después de parsear
      fs.unlink(req.file.path, (err) => {
        if (err) console.error('Error eliminando excel temporal:', err);
      });

    } catch (error) {
      console.error('Error al parsear Excel:', error);
      res.status(500).json({ error: error.message });
    }
  }
};
