import mongoose from 'mongoose';

const solicitudSchema = new mongoose.Schema({
  usuarios_id: mongoose.Schema.Types.ObjectId,
  visitas_digitales_id: mongoose.Schema.Types.ObjectId,
  datos_envio: {
    nombre: String, // representa la especialidad
    direccion: String,
    ciudad: String,
    telefono: String,
  },
  productos_recibidos: String,
  estado_solicitud: String,
  fecha_solicitud: Date,
});

export default mongoose.model('Solicitud', solicitudSchema);