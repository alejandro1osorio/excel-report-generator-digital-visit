import mongoose from 'mongoose';

const respuestaQuizSchema = new mongoose.Schema({
  usuarios_id: mongoose.Schema.Types.ObjectId,
  visitas_digitales_id: mongoose.Schema.Types.ObjectId,
  respuestas: [String],
  fecha_respuesta: Date,
});

export default mongoose.model('RespuestaQuiz', respuestaQuizSchema);