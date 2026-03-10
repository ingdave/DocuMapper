import express, { Router, Request, Response } from 'express';
import { WordVariable } from '../types';

const router = Router();

// Helper para extraer variables del Word
function extractWordVariables(buffer: Buffer): WordVariable[] {
  // Simulación - se refinaría con docxtemplater
  return [
    { name: 'nombre_cliente', placeholder: '{{nombre_cliente}}' },
    { name: 'correo', placeholder: '{{correo}}' },
    { name: 'telefono', placeholder: '{{telefono}}' },
    { name: 'empresa', placeholder: '{{empresa}}' }
  ];
}

// POST /api/word/extract-variables
router.post('/extract-variables', express.raw({ type: 'application/octet-stream', limit: '50mb' }), 
  (req: Request, res: Response) => {
    try {
      const buffer = req.body as Buffer;
      const variables = extractWordVariables(buffer);
      res.json({ variables });
    } catch (error: any) {
      res.status(400).json({ error: error.message || 'Error extrayendo variables' });
    }
  }
);

export { router as wordRouter };
