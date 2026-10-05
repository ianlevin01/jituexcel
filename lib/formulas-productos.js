const TEXTO_MARCADOR = "1:amount of goods";
const PRIMERA_COLUMNA_DATOS = 6; // F = CTNS
const ULTIMA_COLUMNA_DATOS_EXISTENTE = 13; // M = T.KGS
const PRIMERA_FILA_PRODUCTOS = 2;

function celdaVacia(valor) {
  return valor === null || valor === undefined || valor === "";
}

function ultimaFilaConDatos(worksheet) {
  let ultima = 0;
  worksheet.eachRow({ includeEmpty: false }, (fila, numeroFila) => {
    ultima = Math.max(ultima, numeroFila);
  });
  return ultima;
}

function encontrarFilaMarcador(worksheet, filaMaxima) {
  for (let f = 1; f <= filaMaxima; f++) {
    const valor = worksheet.getCell(f, 1).value;
    if (typeof valor === "string" && valor.trim().toLowerCase() === TEXTO_MARCADOR) {
      return f;
    }
  }
  return null;
}

function encontrarUltimaFilaProducto(worksheet, filaMarcador) {
  let ultima = PRIMERA_FILA_PRODUCTOS - 1;
  for (let f = PRIMERA_FILA_PRODUCTOS; f < filaMarcador; f++) {
    let tieneDatos = false;
    for (let c = PRIMERA_COLUMNA_DATOS; c <= ULTIMA_COLUMNA_DATOS_EXISTENTE; c++) {
      if (!celdaVacia(worksheet.getCell(f, c).value)) {
        tieneDatos = true;
        break;
      }
    }
    if (!tieneDatos) break;
    ultima = f;
  }
  return ultima;
}

// Solo LEE el workbook (nunca lo modifica ni lo vuelve a guardar) y arma el plan
// de que formula va en cada celda. La escritura real se hace aparte, quirurgicamente
// sobre el XML original (ver lib/xlsx-patch.js), para no perder nada que ExcelJS no
// sepa preservar (imagenes WPS, metadata, etc).
function calcularPlan(workbook) {
  const worksheet = workbook.worksheets[0];

  const filaMaxima = ultimaFilaConDatos(worksheet);
  const filaMarcador = encontrarFilaMarcador(worksheet, filaMaxima);

  if (!filaMarcador) {
    throw new Error('No se encontró la celda con el texto "1:Amount of Goods" en la columna A.');
  }

  const valorJ = worksheet.getCell(filaMarcador, 2).value;
  if (typeof valorJ !== "number") {
    throw new Error(
      `La celda B${filaMarcador} (al lado de "1:Amount of Goods") no tiene un número válido.`
    );
  }

  const ultimaFilaProducto = encontrarUltimaFilaProducto(worksheet, filaMarcador);
  if (ultimaFilaProducto < PRIMERA_FILA_PRODUCTOS) {
    throw new Error("No se encontraron filas de productos con datos (columnas F a M) antes de la fila del marcador.");
  }

  // Tres filas nuevas, inmediatamente debajo de la ultima fila con CUALQUIER
  // contenido en toda la hoja (incluyendo todo lo que haya debajo del marcador).
  const filaJ67 = filaMaxima + 1;
  const fila10000 = filaMaxima + 2;
  const filaX = filaMaxima + 3;
  const refX = `$A$${filaX}`;

  const celdasProducto = [];
  for (let f = PRIMERA_FILA_PRODUCTOS; f <= ultimaFilaProducto; f++) {
    celdasProducto.push({
      fila: f,
      celdas: {
        14: `66/J${f}*H${f}*G${f}/6.7`, // N
        15: `H${f}/6.7*(1+${refX})`, // O
        16: `O${f}*1.2`, // P
        17: `P${f}*G${f}*F${f}`, // Q
        18: `P${f}*G${f}-O${f}*G${f}`, // R
        19: `66/J${f}*R${f}`, // S
        20: `P${f}*6.7/H${f}`, // T
        21: `P${f}/O${f}`, // U
      },
    });
  }

  const celdasPie = [
    { fila: filaJ67, columna: 1, formula: `B${filaMarcador}/6.7` },
    { fila: fila10000, columna: 1, valor: 10000 },
    { fila: filaX, columna: 1, formula: `A${fila10000}/A${filaJ67}` },
  ];

  return {
    filaMarcador,
    valorJ,
    primeraFilaProducto: PRIMERA_FILA_PRODUCTOS,
    ultimaFilaProducto,
    cantidadFilasProducto: ultimaFilaProducto - PRIMERA_FILA_PRODUCTOS + 1,
    filaJ67,
    fila10000,
    filaX,
    celdasProducto,
    celdasPie,
  };
}

module.exports = { calcularPlan };
