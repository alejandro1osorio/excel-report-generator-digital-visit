import mongoose from 'mongoose';

const videoSchema = new mongoose.Schema({
  titulo: String,
  descripcion: String,
  url_video: String,
});

const visitaDigitalSchema = new mongoose.Schema({
  titulo: String,
  contenido: {
    videos: [videoSchema],
  },
  quizzes_id: mongoose.Schema.Types.ObjectId,
});

export default mongoose.model('VisitaDigital', visitaDigitalSchema);