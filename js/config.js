/*
 * config.js — Configuración de la página (no de un edificio: los datos del edificio van en el Excel).
 *
 * ┌──────────────────────────────────────────────────────────────────────────────┐
 * │ GOOGLE_CLIENT_ID: ID de cliente OAuth (tipo "Aplicación web") de Google Cloud.│
 * │ Origen de JavaScript autorizado: https://msakiya.github.io                    │
 * │ Permiso (scope) único: https://www.googleapis.com/auth/drive.file             │
 * │ No es una contraseña: es público por diseño (va en la página).               │
 * │ Vacío = el botón "Conectar Google Drive" avisa que falta configurarlo.       │
 * └──────────────────────────────────────────────────────────────────────────────┘
 */
window.RECIBOS_CONFIG = {
  GOOGLE_CLIENT_ID: '',   // ← pega aquí el ID, ej. '1234567890-abcdefg.apps.googleusercontent.com'
  DRIVE_SCOPE: 'https://www.googleapis.com/auth/drive.file'
};
