// Configuración de direcciones de DOZZER.

// La API se busca en el mismo equipo desde el que se abrió DOZZER (computadora o celular en la red).
export const API_URL = `${window.location.protocol}//${window.location.hostname}:8080/api`;

// Dirección pública de la ficha del QR (el archivo src/etiqueta.html publicado en internet).
// Al escanear una etiqueta se abre esta página con los datos del bien, que viajan dentro del QR.
//
// IMPORTANTE: esta dirección queda IMPRESA en cada etiqueta. Cuando la ficha esté publicada,
// cámbiala aquí por la definitiva (ej. 'https://dozzer.mx/etiqueta.html') ANTES de imprimir
// las etiquetas finales. Los QR impresos con otra dirección dejarán de mostrar la ficha.
//
// Valor actual: dirección de PRUEBA en la red de la oficina (solo funciona en el mismo Wi-Fi).
// Si se deja vacía (''), el QR lleva los datos como texto simple, sin página.
export const QR_FICHA_URL = 'http://192.168.1.72:4200/etiqueta.html';
