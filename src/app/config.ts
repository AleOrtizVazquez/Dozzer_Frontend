// Configuración de direcciones de DOZZER.

// ¿Se abrió DOZZER desde la computadora o la red de la oficina (pruebas)?
const host = window.location.hostname;
const ES_LOCAL = host === 'localhost' || /^(127\.|10\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.)/.test(host);

// Backend publicado en Render (internet).
const API_PRODUCCION = 'https://dozzer-api.onrender.com/api';

// En pruebas, la API se busca en el mismo equipo (puerto 8080); publicado, se usa la de Render.
export const API_URL = ES_LOCAL
  ? `${window.location.protocol}//${host}:8080/api`
  : API_PRODUCCION;

// Dirección de la ficha del QR (el archivo src/etiqueta.html). Al escanear una etiqueta se abre
// esta página con los datos del bien, que viajan dentro del QR.
//
// IMPORTANTE: esta dirección queda IMPRESA en cada etiqueta. Se toma de la página desde la que se
// imprimen: imprime las etiquetas finales desde la versión publicada en internet
// (ej. https://dozzer.pages.dev), nunca desde la computadora de pruebas.
// Si se deja vacía (''), el QR lleva los datos como texto simple, sin página.
export const QR_FICHA_URL = `${window.location.origin}/etiqueta.html`;
