const { numeroALetraColumna } = require("./excel-columnas");

const COLUMNA_Q = 17; // Q, fija: donde se pega el precio en el excel A
const COLUMNA_CODIGO_A = 2; // B, fija: codigo de producto en el excel A
const COLUMNA_CODIGO_B = 1; // A, fija: codigo de producto en el excel B
const COLUMNA_PRECIO_B = 4; // D, fija: precio en el excel B

function ultimaFilaConDatos(worksheet) {
  let ultima = 0;
  worksheet.eachRow({ includeEmpty: false }, (fila, numeroFila) => {
    ultima = Math.max(ultima, numeroFila);
  });
  return ultima;
}

function valorCrudo(valor) {
  if (valor === null || valor === undefined) return null;
  if (typeof valor === "object") {
    if ("result" in valor) return valor.result;
    if (Array.isArray(valor.richText)) return valor.richText.map((p) => p.text).join("");
    if (valor.text !== undefined) return valor.text;
    return null;
  }
  return valor;
}

// Para comparar codigos de producto sin que importe si uno quedo como texto
// y el otro como numero, ni espacios/mayusculas de mas.
function normalizarCodigo(valor) {
  return String(valor).trim().toUpperCase();
}

// Solo LEE los dos workbooks. Arma el plan para: 1) correr una posicion a la
// derecha todo lo que haya en el excel A desde la columna Q en adelante, y
// 2) pegar en la columna Q (ya vacia) el precio de cada producto, buscado
// por CODIGO (no por fila): codigo en excel B = columna A, precio en excel
// B = columna D, codigo en excel A = columna B.
function calcularPlanInsertarColumna(workbookA, workbookB) {
  const hojaA = workbookA.worksheets[0];
  const hojaB = workbookB.worksheets[0];

  const filaMaximaA = ultimaFilaConDatos(hojaA);
  const filaMaximaB = ultimaFilaConDatos(hojaB);

  // 1) excel B: armar el mapa codigo -> precio
  const precioPorCodigo = new Map();
  for (let f = 1; f <= filaMaximaB; f++) {
    const codigo = valorCrudo(hojaB.getCell(f, COLUMNA_CODIGO_B).value);
    const precio = valorCrudo(hojaB.getCell(f, COLUMNA_PRECIO_B).value);
    if (codigo === null || precio === null) continue;
    precioPorCodigo.set(normalizarCodigo(codigo), precio);
  }

  if (precioPorCodigo.size === 0) {
    throw new Error(
      `No se encontró ningún código (columna ${numeroALetraColumna(COLUMNA_CODIGO_B)}) con precio (columna ${numeroALetraColumna(COLUMNA_PRECIO_B)}) en el excel B.`
    );
  }

  // 2) excel A: para cada fila con codigo, buscar su precio en el mapa
  const valoresQ = [];
  const codigosSinPrecio = [];
  for (let f = 1; f <= filaMaximaA; f++) {
    const codigo = valorCrudo(hojaA.getCell(f, COLUMNA_CODIGO_A).value);
    if (codigo === null) continue;

    const precio = precioPorCodigo.get(normalizarCodigo(codigo));
    if (precio === undefined) {
      codigosSinPrecio.push({ fila: f, codigo });
      continue;
    }
    valoresQ.push({ fila: f, valor: precio });
  }

  return {
    columnaQ: numeroALetraColumna(COLUMNA_Q),
    columnaCodigoA: numeroALetraColumna(COLUMNA_CODIGO_A),
    columnaCodigoB: numeroALetraColumna(COLUMNA_CODIGO_B),
    columnaPrecioB: numeroALetraColumna(COLUMNA_PRECIO_B),
    filaMaximaA,
    filaMaximaB,
    cantidadValores: valoresQ.length,
    cantidadSinPrecio: codigosSinPrecio.length,
    codigosSinPrecio,
    valoresQ,
  };
}

module.exports = { calcularPlanInsertarColumna, COLUMNA_Q };
