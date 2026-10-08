const { numeroALetraColumna, encontrarColumnaPorTitulo, primeraColumnaVacia } = require("./excel-columnas");

const PRIMERA_FILA_PRODUCTOS = 2;
const PALABRA_PRICE = "price";

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

function anchoDeHoja(worksheet) {
  return Math.max(worksheet.columnCount || 0, worksheet.actualColumnCount || 0, 50);
}

// Ultima fila "de producto" en A: arranca en la fila 2 y para en la primera
// fila que no tenga NADA en ninguna de las columnas ya existentes (todo lo
// que esta antes de la primera columna vacia).
function encontrarUltimaFilaProductoA(worksheet, columnaQ, filaLimiteExclusiva) {
  let ultima = PRIMERA_FILA_PRODUCTOS - 1;
  for (let f = PRIMERA_FILA_PRODUCTOS; f < filaLimiteExclusiva; f++) {
    let tieneDatos = false;
    for (let c = 1; c < columnaQ; c++) {
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

// Solo LEE los dos workbooks (nunca los modifica). Arma el plan de que valor/formula
// va en cada celda nueva del excel A. La escritura real se hace aparte, quirurgicamente
// sobre el XML original del excel A (ver lib/xlsx-patch.js).
function calcularPlanAB(workbookA, workbookB) {
  const hojaA = workbookA.worksheets[0];
  const hojaB = workbookB.worksheets[0];

  const columnaQ = primeraColumnaVacia(hojaA);
  const columnaR = columnaQ + 1;
  const columnaS = columnaQ + 2;

  const columnaPriceA = encontrarColumnaPorTitulo(hojaA, PALABRA_PRICE, columnaQ - 1);
  if (!columnaPriceA) {
    throw new Error('No se encontró ninguna columna con título "Price" en el excel A.');
  }

  const columnaPriceB = encontrarColumnaPorTitulo(hojaB, PALABRA_PRICE, anchoDeHoja(hojaB));
  if (!columnaPriceB) {
    throw new Error('No se encontró ninguna columna con título "Price" en el excel B.');
  }

  const filaMaximaA = ultimaFilaConDatos(hojaA);
  const ultimaFilaProducto = encontrarUltimaFilaProductoA(hojaA, columnaQ, filaMaximaA + 1);
  if (ultimaFilaProducto < PRIMERA_FILA_PRODUCTOS) {
    throw new Error("No se encontraron filas de productos con datos en el excel A.");
  }

  const letraQ = numeroALetraColumna(columnaQ);
  const letraR = numeroALetraColumna(columnaR);
  const letraPriceA = numeroALetraColumna(columnaPriceA);

  const celdasProducto = [];
  let filasOmitidas = 0;

  for (let f = PRIMERA_FILA_PRODUCTOS; f <= ultimaFilaProducto; f++) {
    const precioB = valorNumerico(hojaB.getCell(f, columnaPriceB).value);
    if (precioB === null) {
      filasOmitidas++;
      continue;
    }
    celdasProducto.push({
      fila: f,
      celdas: {
        [columnaQ]: { valor: precioB },
        [columnaR]: { formula: `${letraQ}${f}/6.7` },
        [columnaS]: { formula: `${letraPriceA}${f}/${letraR}${f}` },
      },
    });
  }

  if (celdasProducto.length === 0) {
    throw new Error('Ninguna fila tiene un precio numérico en la columna "Price" del excel B para copiar.');
  }

  return {
    columnaQ: letraQ,
    columnaR: letraR,
    columnaS: numeroALetraColumna(columnaS),
    columnaPriceA: letraPriceA,
    columnaPriceB: numeroALetraColumna(columnaPriceB),
    primeraFilaProducto: PRIMERA_FILA_PRODUCTOS,
    ultimaFilaProducto,
    cantidadFilasProcesadas: celdasProducto.length,
    filasOmitidas,
    celdasProducto,
  };
}

module.exports = { calcularPlanAB };
