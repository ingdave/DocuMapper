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
    // Si tiene propiedad 'result', úsala
    if (value.result !== undefined) return String(value.result);
    return String(value);
  }
  
  // Para números y strings, devolver como está
  return String(value);
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

      // Obtener headers (primera fila)
      const headers = [];
      worksheet.getRow(1).eachCell((cell, colNumber) => {
        if (cell.value) {
          headers.push(cell.value.toString());
        }
      });

      // Obtener datos
      const data = [];
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header
        
        const rowData = {};
        headers.forEach((header, index) => {
          const cellValue = row.getCell(index + 1).value;
          rowData[header] = formatValue(cellValue);
        });
        
        if (Object.values(rowData).some(val => val !== '')) {
          data.push(rowData);
        }
      });

      res.json({
        success: true,
        headers,
        data,
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
