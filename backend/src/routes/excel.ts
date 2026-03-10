import express, { Router, Request, Response } from 'express';
import { ExcelData } from '../types';

const router = Router();

// Helper para procesar buffer Excel
function processExcelBuffer(buffer: Buffer): ExcelData {
  // Simulación - se implementará con exceljs
  return {
    columns: ['Nombre', 'Email', 'Teléfono', 'Empresa'],
    rows: [
      { 'Nombre': 'Juan Pérez', 'Email': 'juan@example.com', 'Teléfono': '+34 600000001', 'Empresa': 'Acme Corp' }
    ]
  };
}

// POST /api/excel/parse
router.post('/parse', express.raw({ type: 'application/octet-stream', limit: '50mb' }), 
  (req: Request, res: Response) => {
    try {
      const buffer = req.body as Buffer;
      const excelData = processExcelBuffer(buffer);
      res.json(excelData);
    } catch (error: any) {
      res.status(400).json({ error: error.message || 'Error procesando Excel' });
    }
  }
);

export { router as excelRouter };
