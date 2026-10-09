const { numeroALetraColumna } = require("./excel-columnas");

const COLUMNA_Q = 17; // Q, fija
const COLUMNA_D_B = 4; // D, fija

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

// Solo LEE los dos workbooks. Arma el plan para: 1) correr una posicion a la
// derecha todo lo que haya en el excel A desde la columna Q en adelante, y
// 2) pegar en la columna Q (ya vacia) los valores de la columna D del excel B,
// fila por fila (misma fila en A y B).
function calcularPlanInsertarColumna(workbookA, workbookB) {
  const hojaA = workbookA.worksheets[0];
  const hojaB = workbookB.worksheets[0];

  const filaMaximaA = ultimaFilaConDatos(hojaA);
  const filaMaximaB = ultimaFilaConDatos(hojaB);

  let filaMaximaD = 0;
  for (let f = 1; f <= filaMaximaB; f++) {
    if (valorCrudo(hojaB.getCell(f, COLUMNA_D_B).value) !== null) filaMaximaD = f;
  }

  if (filaMaximaD === 0) {
    throw new Error(`La columna ${numeroALetraColumna(COLUMNA_D_B)} del excel B no tiene ningún valor para copiar.`);
  }

  const valoresQ = [];
  for (let f = 1; f <= filaMaximaD; f++) {
    const valor = valorCrudo(hojaB.getCell(f, COLUMNA_D_B).value);
    if (valor === null) continue;
    valoresQ.push({ fila: f, valor });
  }

  return {
    columnaQ: numeroALetraColumna(COLUMNA_Q),
    columnaDB: numeroALetraColumna(COLUMNA_D_B),
    filaMaximaA,
    filaMaximaD,
    cantidadValores: valoresQ.length,
    valoresQ,
  };
}

module.exports = { calcularPlanInsertarColumna, COLUMNA_Q };
