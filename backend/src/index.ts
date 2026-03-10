import express, { Express, Request, Response } from 'express';
import cors from 'cors';
import path from 'path';
import { excelRouter } from './routes/excel';
import { wordRouter } from './routes/word';
import { mappingRouter } from './routes/mapping';

const app: Express = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// Routes
app.use('/api/excel', excelRouter);
app.use('/api/word', wordRouter);
app.use('/api/mapping', mappingRouter);

// Health check
app.get('/api/health', (req: Request, res: Response) => {
  res.json({ status: 'OK' });
});

// Error handling
app.use((err: any, req: Request, res: Response, next: any) => {
  console.error(err.stack);
  res.status(500).json({ error: err.message || 'Error interno del servidor' });
});

app.listen(PORT, () => {
  console.log(`✓ Servidor ejecutándose en http://localhost:${PORT}`);
});
