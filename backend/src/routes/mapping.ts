import express, { Router, Request, Response } from 'express';
import { GenerateDocumentRequest, MappingConfig } from '../types';

const router = Router();

// POST /api/mapping/validate
router.post('/validate', (req: Request, res: Response) => {
  try {
    const { mappingConfig, columns, variables } = req.body;
    
    // Validar que todas las columnas Excel estén mapeadas
    const unmappedColumns = (columns as string[]).filter((col: string) => !mappingConfig.mappings[col]);
    const unmappedVariables = (variables as any[]).filter((v: any) => 
      !Object.values(mappingConfig.mappings).includes(v.placeholder)
    );

    res.json({
      isValid: unmappedColumns.length === 0,
      unmappedColumns,
      unmappedVariables
    });
  } catch (error: any) {
    res.status(400).json({ error: error.message });
  }
});

// POST /api/mapping/generate
router.post('/generate', express.raw({ type: 'application/octet-stream', limit: '50mb' }), 
  (req: Request, res: Response) => {
    try {
      const { mappingConfig, rows, wordTemplateBase64 } = JSON.parse(req.body.toString());
      
      // Implementación de generación de documento
      // Esto se refinaría con docxtemplater
      
      res.json({ 
        message: 'Documento generado exitosamente',
        generated: true 
      });
    } catch (error: any) {
      res.status(400).json({ error: error.message });
    }
  }
);

// POST /api/mapping/save
router.post('/save', (req: Request, res: Response) => {
  try {
    const mappingConfig: MappingConfig = req.body;
    // Aquí se guardaría en base de datos o localStorage en frontend
    res.json({ success: true, timestamp: mappingConfig.timestamp });
  } catch (error: any) {
    res.status(400).json({ error: error.message });
  }
});

export { router as mappingRouter };
