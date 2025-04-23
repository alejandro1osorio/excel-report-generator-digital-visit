import dotenv from 'dotenv';
import mongoose from 'mongoose';
import connectDB from './config/db.js';
import { generateExcelReport } from './services/excelService.js';

dotenv.config();

(async () => {
  try {
    await connectDB();
    console.log('Conectado a MongoDB');

    await generateExcelReport();
    console.log('Reporte generado en ./reports/report.xlsx');

    process.exit();
  } catch (error) {
    console.error('Error al generar el reporte:', error);
    process.exit(1);
  }
})();