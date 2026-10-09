function numeroALetraColumna(numero) {
  let letra = "";
  let n = numero;
  while (n > 0) {
    const resto = (n - 1) % 26;
    letra = String.fromCharCode(65 + resto) + letra;
    n = Math.floor((n - 1) / 26);
  }
  return letra;
}

function letraAColumna(letra) {
  let n = 0;
  for (const ch of letra.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

function headerContiene(worksheet, columna, palabra) {
  const valor = worksheet.getCell(1, columna).value;
  return typeof valor === "string" && valor.trim().toLowerCase().includes(palabra.toLowerCase());
}

// Busca, de izquierda a derecha, la primera columna cuyo titulo (fila 1)
// contenga la palabra buscada (sin importar mayusculas/minusculas). Devuelve
// el numero de columna, o null si ninguna coincide.
function encontrarColumnaPorTitulo(worksheet, palabra, columnaMaxima) {
  for (let c = 1; c <= columnaMaxima; c++) {
    if (headerContiene(worksheet, c, palabra)) return c;
  }
  return null;
}

// Primera columna (de izquierda a derecha) que no tiene NINGUN contenido en
// toda la hoja (ni titulo ni datos en ninguna fila). Es el limite "columnaMaxima"
// + 1 que hay que usar para no quedarse corto buscando headers.
function primeraColumnaVacia(worksheet) {
  let ultimaConContenido = 0;
  worksheet.eachRow({ includeEmpty: false }, (fila) => {
    fila.eachCell({ includeEmpty: false }, (celda, numeroColumna) => {
      ultimaConContenido = Math.max(ultimaConContenido, numeroColumna);
    });
  });
  return ultimaConContenido + 1;
}

module.exports = { numeroALetraColumna, letraAColumna, encontrarColumnaPorTitulo, primeraColumnaVacia };
