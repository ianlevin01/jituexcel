const { numeroALetraColumna, encontrarColumnaPorTitulo } = require("./excel-columnas");

const PRIMERA_COLUMNA_DATOS = 6; // F = CTNS
const ULTIMA_COLUMNA_DATOS_EXISTENTE = 13; // M = T.KGS
const PRIMERA_FILA_PRODUCTOS = 2;
const COLUMNA_AMOUNT_POR_DEFECTO = 9; // I
const PALABRA_AMOUNT = "amount";

function celdaVacia(valor) {
  return valor === null || valor === undefined || valor === "";
}

function valorNumerico(valor) {
  if (typeof valor === "number") return valor;
  if (valor && typeof valor === "object" && typeof valor.result === "number") return valor.result;
  return null;
}

function ultimaFilaConDatos(worksheet) {
  let ultima = 0;
  worksheet.eachRow({ includeEmpty: false }, (fila, numeroFila) => {
    ultima = Math.max(ultima, numeroFila);
  });
  return ultima;
}

// Busca la columna de "Amount of Goods" en este orden, para que nunca se rompa
// aunque el excel no venga con el formato esperado:
//   1) la columna I, si su titulo contiene "amount" (en cualquier mayuscula/minuscula)
//   2) cualquier otra columna cuyo titulo contenga "amount"
//   3) si ninguna columna dice "amount", se usa igual la columna I
function encontrarColumnaAmount(worksheet) {
  const valorColumnaI = worksheet.getCell(1, COLUMNA_AMOUNT_POR_DEFECTO).value;
  if (typeof valorColumnaI === "string" && valorColumnaI.trim().toLowerCase().includes(PALABRA_AMOUNT)) {
    return COLUMNA_AMOUNT_POR_DEFECTO;
  }

  const ultimaColumna = Math.max(worksheet.columnCount || 0, worksheet.actualColumnCount || 0, ULTIMA_COLUMNA_DATOS_EXISTENTE);
  const columnaConAmount = encontrarColumnaPorTitulo(worksheet, PALABRA_AMOUNT, ultimaColumna);
  if (columnaConAmount) return columnaConAmount;

  return COLUMNA_AMOUNT_POR_DEFECTO;
}

// En la columna encontrada, el valor de "Amount of Goods" es el de la ULTIMA
// fila que tenga ahi un numero (sea un valor cargado a mano o el resultado de
// una formula), buscando en toda la hoja.
function encontrarValorAmountOfGoods(worksheet, filaMaxima) {
  const columna = encontrarColumnaAmount(worksheet);

  let filaEncontrada = null;
  let valorEncontrado = null;
  for (let f = PRIMERA_FILA_PRODUCTOS; f <= filaMaxima; f++) {
    const numero = valorNumerico(worksheet.getCell(f, columna).value);
    if (numero !== null) {
      filaEncontrada = f;
      valorEncontrado = numero;
    }
  }

  if (filaEncontrada === null) {
    throw new Error(
      `No se encontró ningún valor numérico en la columna ${numeroALetraColumna(columna)} para usar como "Amount of Goods".`
    );
  }

  return { fila: filaEncontrada, columna, valor: valorEncontrado };
}

function encontrarUltimaFilaProducto(worksheet, filaLimiteExclusiva) {
  let ultima = PRIMERA_FILA_PRODUCTOS - 1;
  for (let f = PRIMERA_FILA_PRODUCTOS; f < filaLimiteExclusiva; f++) {
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

  const { fila: filaAmount, columna: columnaAmount, valor: valorJ } = encontrarValorAmountOfGoods(worksheet, filaMaxima);

  const ultimaFilaProducto = encontrarUltimaFilaProducto(worksheet, filaMaxima + 1);
  if (ultimaFilaProducto < PRIMERA_FILA_PRODUCTOS) {
    throw new Error("No se encontraron filas de productos con datos (columnas F a M).");
  }

  // Tres filas nuevas, inmediatamente debajo de la ultima fila con CUALQUIER
  // contenido en toda la hoja, en la primera columna (A).
  const filaJ67 = filaMaxima + 1;
  const fila10000 = filaMaxima + 2;
  const filaX = filaMaxima + 3;
  const refX = `$A$${filaX}`;
  const refAmount = `${numeroALetraColumna(columnaAmount)}${filaAmount}`;

  const celdasProducto = [];
  for (let f = PRIMERA_FILA_PRODUCTOS; f <= ultimaFilaProducto; f++) {
    celdasProducto.push({
      fila: f,
      celdas: {
        14: { formula: `66/J${f}*H${f}*G${f}/6.7` }, // N
        15: { formula: `H${f}/6.7*(1+${refX})` }, // O
        16: { formula: `O${f}*1.2` }, // P
        17: { formula: `P${f}*G${f}*F${f}` }, // Q
        18: { formula: `P${f}*G${f}-O${f}*G${f}` }, // R
        19: { formula: `66/J${f}*R${f}` }, // S
        20: { formula: `P${f}*6.7/H${f}` }, // T
        21: { formula: `P${f}/O${f}` }, // U
      },
    });
  }

  const celdasPie = [
    { fila: filaJ67, columna: 1, formula: `${refAmount}/6.7` },
    { fila: fila10000, columna: 1, valor: 10000 },
    { fila: filaX, columna: 1, formula: `A${fila10000}/A${filaJ67}` },
  ];

  return {
    filaAmount,
    columnaAmount: numeroALetraColumna(columnaAmount),
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
