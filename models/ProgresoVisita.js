import mongoose from 'mongoose';

const progresoSchema = new mongoose.Schema({
  usuarios_id: mongoose.Schema.Types.ObjectId,
  visitas_digitales_id: mongoose.Schema.Types.ObjectId,
  completado: Boolean,
  quiz_completado: Boolean,
  progreso: [
    {
      titulo_video: String,
      completado: Boolean,
      fecha_visto: Date,
    },
  ],
  fecha_completado: Date,
});

export default mongoose.model('ProgresoVisita', progresoSchema);