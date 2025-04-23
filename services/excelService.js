import * as XLSX from 'xlsx';
import { fetchData } from './fetchDataService.js';
import { formatDate } from '../utils/formatters.js';
import fs from 'fs';

export const generateExcelReport = async () => {
  const { usuarios, visitas, progresos, quizzes, respuestas, solicitudes } = await fetchData();

  const mainSheet = [];
  const solicitudesSheet = [];
  const styleData = [];

  usuarios.forEach(user => {
    const progreso = progresos.find(p => p.usuarios_id.toString() === user._id.toString());
    const visita = visitas.find(v => progreso && v._id.toString() === progreso.visitas_digitales_id.toString());
    const respuestaQuiz = respuestas.find(r => r.usuarios_id.toString() === user._id.toString());
    const quiz = visita ? quizzes.find(q => q._id.toString() === visita.quizzes_id.toString()) : null;
    const solicitud = solicitudes.find(s => s.usuarios_id.toString() === user._id.toString());

    const correo = typeof user.email === 'string' ? user.email.trim() : '';
    console.log('📧 Correo de usuario:', user.email, '→ Procesado:', correo);

    const fila = {
      'Territorio': user.territorio,
      'Código Colgate': user.codigo_colgate,
      'Cedula': user.cedula,
      'Nombre completo': `${user.nombres} ${user.apellidos}`,
      'Dirección': user.direccion,
      'Ciudad': user.ciudad,
      'Email': user.email1 || user.email || '',
    };

    const estilosFila = {};

    if (visita) {
      visita.contenido.videos.forEach((video, i) => {
        const progresoVideo = progreso?.progreso.find(pv => pv.titulo_video === video.titulo);
        fila[`Video ${i + 1} - ${video.titulo}`] = progresoVideo?.completado ? 'Visto' : 'No';
      });
    }

    if (quiz) {
      quiz.preguntas.forEach((preg, i) => {
        const respuesta = respuestaQuiz?.respuestas[i] || '';
        const key = `Pregunta ${i + 1}`;
        fila[key] = respuesta;

        if (respuesta && respuesta !== preg.respuesta_correcta) {
          estilosFila[key] = { fill: { fgColor: { rgb: "FFCCCC" } } };
        }
      });
    }

    fila['Quiz Completado'] = progreso?.quiz_completado ? 'Si' : 'No';
    fila['Solicitud'] = solicitud ? 'Realizada' : 'No';
    fila['Visita Completada'] = progreso?.completado ? 'Si' : 'No';
    fila['Fecha Completada'] = formatDate(progreso?.fecha_completado);

    mainSheet.push(fila);
    styleData.push(estilosFila);

    if (solicitud) {
      solicitudesSheet.push({
        'Código Colgate': user.codigo_colgate,
        'Cedula': user.cedula,
        'Nombre completo': `${user.nombres} ${user.apellidos}`,
        'Especialidad': solicitud.datos_envio?.nombre || '',
        'Dirección': solicitud.datos_envio?.direccion || '',
        'Ciudad': solicitud.datos_envio?.ciudad || '',
        'Teléfono': solicitud.datos_envio?.telefono || '',
        'Email': user.email1 || user.email || '',
        'Fecha Solicitud': formatDate(solicitud.fecha_solicitud)
      });
    }
  });

  const wb = XLSX.utils.book_new();
  const resumenSheet = XLSX.utils.json_to_sheet(mainSheet, { cellStyles: true });

  const headerKeys = Object.keys(mainSheet[0] || {});
  const rowHeight = 42;
  const columnWidths = [
    12.38, 15.88, 12.75, 32.88, 11.63, 18.13, 31.25, 14.75, 10.38, 16.5, 18.13,
    23.38, 26.00, 30.88, 53.75, 31.5, 34.13, 24.75, 101.00
  ];

  headerKeys.forEach((key, colIdx) => {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c: colIdx });
    const cell = resumenSheet[cellRef] || {};
    cell.s = {
      font: { bold: true },
      alignment: { horizontal: 'center', vertical: 'center' },
      fill: { fgColor: { rgb: 'DDEBF7' } }
    };
    resumenSheet[cellRef] = cell;
  });

  resumenSheet['!rows'] = [{ hpt: rowHeight }];
  resumenSheet['!cols'] = columnWidths.map(w => ({ wch: w }));

  wb.SheetNames.push('Resumen');
  wb.SheetNames.push('Solicitudes');
  wb.Sheets['Resumen'] = resumenSheet;
  wb.Sheets['Solicitudes'] = XLSX.utils.json_to_sheet(solicitudesSheet);

  if (!fs.existsSync('./reports')) fs.mkdirSync('./reports');
  XLSX.writeFile(wb, './reports/report.xlsx');
};