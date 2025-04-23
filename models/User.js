import mongoose from 'mongoose';

const userSchema = new mongoose.Schema({
  nombres: String,
  apellidos: String,
  cedula: String,
  email: String,
  direccion: String,
  ciudad: String,
  codigo_colgate: String,
  role: String,
  territorio: String,
});

export default mongoose.model('User', userSchema, 'usuarios');