import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import { excelController } from './controllers/excelController.js';
import { wordController } from './controllers/wordController.js';
import { mappingController } from './controllers/mappingController.js';

const app = express();
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));
app.use(express.static(path.join(__dirname, '../public')));

// Directorio de subidas y descargas
const uploadsDir = path.join(__dirname, '../uploads');
const downloadsDir = path.join(__dirname, '../downloads');

if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });
if (!fs.existsSync(downloadsDir)) fs.mkdirSync(downloadsDir, { recursive: true });

// Configuración de multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({ storage });

// Rutas
app.post('/api/excel/parse', upload.single('file'), excelController.parseExcel);
app.post('/api/word/extract-variables', upload.single('file'), wordController.extractVariables);
app.post('/api/mapping/generate', mappingController.generateDocuments);
app.post('/api/mapping/download-all-docx', mappingController.downloadAllDocx);
app.post('/api/mapping/download-all-pdf', mappingController.downloadAllPdf);

// Ruta de descarga
app.get('/api/download/:filename', (req, res) => {
  const filename = req.params.filename;
  const filepath = path.join(downloadsDir, filename);
  
  if (fs.existsSync(filepath)) {
    res.download(filepath, () => {
      fs.unlink(filepath, (err) => {
        if (err) console.error('Error eliminando archivo:', err);
      });
    });
  } else {
    res.status(404).json({ error: 'Archivo no encontrado' });
  }
});

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.listen(PORT, () => {
  console.log(`Servidor ejecutándose en http://localhost:${PORT}`);
});
