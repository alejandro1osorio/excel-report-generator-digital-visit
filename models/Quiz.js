import mongoose from 'mongoose';

const quizSchema = new mongoose.Schema({
  titulo: String,
  preguntas: [
    {
      texto_pregunta: String,
      opciones: [String],
      respuesta_correcta: String,
    },
  ],
});

export default mongoose.model('Quiz', quizSchema);