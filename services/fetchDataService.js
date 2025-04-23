import User from '../models/User.js';
import VisitaDigital from '../models/VisitaDigital.js';
import ProgresoVisita from '../models/ProgresoVisita.js';
import Quiz from '../models/Quiz.js';
import RespuestaQuiz from '../models/RespuestaQuiz.js';
import Solicitud from '../models/Solicitud.js';

export const fetchData = async () => {
  const usuarios = await User.find();
  const visitas = await VisitaDigital.find();
  const progresos = await ProgresoVisita.find();
  const quizzes = await Quiz.find();
  const respuestas = await RespuestaQuiz.find();
  const solicitudes = await Solicitud.find();

  return { usuarios, visitas, progresos, quizzes, respuestas, solicitudes };
};
