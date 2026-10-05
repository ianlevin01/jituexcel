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

function aplicarFormulas(workbook) {
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

  worksheet.getCell(filaJ67, 1).value = { formula: `B${filaMarcador}/6.7` };
  worksheet.getCell(fila10000, 1).value = 10000;
  worksheet.getCell(filaX, 1).value = { formula: `A${fila10000}/A${filaJ67}` };

  const refX = `$A$${filaX}`;

  for (let f = PRIMERA_FILA_PRODUCTOS; f <= ultimaFilaProducto; f++) {
    worksheet.getCell(f, 14).value = { formula: `66/J${f}*H${f}*G${f}/6.7` }; // N
    worksheet.getCell(f, 15).value = { formula: `H${f}/6.7*(1+${refX})` }; // O
    worksheet.getCell(f, 16).value = { formula: `O${f}*1.2` }; // P
    worksheet.getCell(f, 17).value = { formula: `P${f}*G${f}*F${f}` }; // Q
    worksheet.getCell(f, 18).value = { formula: `P${f}*G${f}-O${f}*G${f}` }; // R
    worksheet.getCell(f, 19).value = { formula: `66/J${f}*R${f}` }; // S
    worksheet.getCell(f, 20).value = { formula: `P${f}*6.7/H${f}` }; // T
    worksheet.getCell(f, 21).value = { formula: `P${f}/O${f}` }; // U
  }

  return {
    filaMarcador,
    valorJ,
    primeraFilaProducto: PRIMERA_FILA_PRODUCTOS,
    ultimaFilaProducto,
    cantidadFilasProducto: ultimaFilaProducto - PRIMERA_FILA_PRODUCTOS + 1,
    filaJ67,
    fila10000,
    filaX,
  };
}

module.exports = { aplicarFormulas };
