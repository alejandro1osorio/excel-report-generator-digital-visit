import ExcelJS from 'exceljs';
import { fetchData } from './fetchDataService.js';
import { formatDate } from '../utils/formatters.js';
import fs from 'fs';

export const generateExcelReport = async () => {
  const { usuarios, visitas, progresos, quizzes, respuestas, solicitudes } = await fetchData();

  const workbook = new ExcelJS.Workbook();
  const resumenSheet = workbook.addWorksheet('Resumen');
  const solicitudesSheet = workbook.addWorksheet('Solicitudes');

  const mainSheet = [];
  const solicitudesData = [];

  let allHeaders = new Set([
    'Territorio', 'Código Colgate', 'Cedula', 'Nombre completo', 'Dirección', 'Ciudad', 'Email',
    'Quiz Completado', 'Solicitud', 'Visita Completada', 'Fecha Completada'
  ]);

  usuarios.forEach(user => {
    const progreso = progresos.find(p => p.usuarios_id.toString() === user._id.toString());
    const visita = visitas.find(v => progreso && v._id.toString() === progreso.visitas_digitales_id.toString());
    const respuestaQuiz = respuestas.find(r => r.usuarios_id.toString() === user._id.toString());
    const quiz = visita ? quizzes.find(q => q._id.toString() === visita.quizzes_id.toString()) : null;
    const solicitud = solicitudes.find(s => s.usuarios_id.toString() === user._id.toString());

    const correo = typeof user.email === 'string' ? user.email.trim() : '';

    const fila = {
      'Territorio': user.territorio,
      'Código Colgate': user.codigo_colgate,
      'Cedula': user.cedula,
      'Nombre completo': `${user.nombres} ${user.apellidos}`,
      'Dirección': user.direccion,
      'Ciudad': user.ciudad,
      'Email': correo,
    };

    const estilosFila = {};

    if (visita) {
      visita.contenido.videos.forEach((video, i) => {
        const key = `Video ${i + 1} - ${video.titulo}`;
        const progresoVideo = progreso?.progreso.find(pv => pv.titulo_video === video.titulo);
        const estado = progresoVideo?.completado ? 'Visto' : 'No';
        fila[key] = estado;
        if (estado === 'Visto') estilosFila[key] = 'green';
        allHeaders.add(key);
      });
    }

    if (quiz) {
      quiz.preguntas.forEach((preg, i) => {
        const key = `Pregunta ${i + 1}`;
        const respuesta = respuestaQuiz?.respuestas[i] || '';
        fila[key] = respuesta;
        allHeaders.add(key);

        if (respuesta && respuesta !== preg.respuesta_correcta) {
          estilosFila[key] = 'red';
        }
      });
    }

    const quizStatus = progreso?.quiz_completado ? 'Si' : 'No';
    fila['Quiz Completado'] = quizStatus;
    if (quizStatus === 'Si') estilosFila['Quiz Completado'] = 'green';

    const solicitudStatus = solicitud ? 'Realizada' : 'No';
    fila['Solicitud'] = solicitudStatus;
    if (solicitudStatus === 'Realizada') estilosFila['Solicitud'] = 'green';

    const visitaStatus = progreso?.completado ? 'Si' : 'No';
    fila['Visita Completada'] = visitaStatus;
    if (visitaStatus === 'Si') estilosFila['Visita Completada'] = 'green';

    fila['Fecha Completada'] = formatDate(progreso?.fecha_completado);

    mainSheet.push({ fila, estilosFila });

    if (solicitud) {
      solicitudesData.push({
        'Código Colgate': user.codigo_colgate,
        'Cedula': user.cedula,
        'Nombre completo': `${user.nombres} ${user.apellidos}`,
        'Especialidad': solicitud.datos_envio?.nombre || '',
        'Dirección': solicitud.datos_envio?.direccion || '',
        'Ciudad': solicitud.datos_envio?.ciudad || '',
        'Teléfono': solicitud.datos_envio?.telefono || '',
        'Email': correo,
        'Fecha Solicitud': formatDate(solicitud.fecha_solicitud)
      });
    }
  });

  const headers = Array.from(allHeaders);
  const customWidths = {
    'Territorio': 13.57,
    'Código Colgate': 18.14,
    'Cedula': 15.71,
    'Nombre completo': 37.29,
    'Dirección': 24.14,
    'Ciudad': 23.43,
    'Email': 34.86,
    'Quiz Completado': 18.00,
    'Solicitud': 15.14,
    'Visita Completada': 18.43,
    'Fecha Completada': 19.43,
    'Video 1 - INTRODUCCIÓN': 24.43,
    'Video 2 - HIPERSENSIBILIDAD': 27.14,
    'Video 3 - TECNOLOGÍA PRO ARGIN': 31.86,
    'Video 4 - EFECTIVIDAD CREMA DENTAL COLGATE SENSITIVE': 31.14,
    'Video 5 - TARJETA COLGATE PASS': 30.29,
    'Pregunta 1': 32.86,
    'Pregunta 2': 35.29,
    'Pregunta 3': 109.14
  };

  resumenSheet.columns = headers.map(h => ({ header: h, key: h, width: customWidths[h] || 25 }));

  const headerRow = resumenSheet.getRow(1);
  headerRow.height = 42;
  headerRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DDEBF7' } };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });

  mainSheet.forEach(({ fila, estilosFila }) => {
    const row = resumenSheet.addRow(headers.map(h => fila[h] || ''));
    row.eachCell((cell, colNumber) => {
      const colKey = headers[colNumber - 1];
      if (estilosFila[colKey]) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: estilosFila[colKey] === 'green' ? 'C6EFCE' : 'FFCCCC' }
        };
      }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
  });

  if (solicitudesData.length) {
    const solicitudHeaders = Object.keys(solicitudesData[0]);
    solicitudesSheet.columns = solicitudHeaders.map(h => ({ header: h, key: h, width: 25 }));

    const solicitudHeaderRow = solicitudesSheet.getRow(1);
    solicitudHeaderRow.height = 42;
    solicitudHeaderRow.eachCell(cell => {
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DDEBF7' } };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    solicitudesData.forEach(fila => {
      const row = solicitudesSheet.addRow(solicitudHeaders.map(h => fila[h]));
      row.eachCell(cell => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });
  }

  if (!fs.existsSync('./reports')) fs.mkdirSync('./reports');
  await workbook.xlsx.writeFile('./reports/report.xlsx');
  console.log('✅ Excel generado con ExcelJS con videos, respuestas y solicitudes completas con anchos personalizados');
};
