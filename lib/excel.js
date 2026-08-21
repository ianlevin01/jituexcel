const ExcelJS = require("exceljs");

function normalizarEncabezado(valor) {
  if (typeof valor !== "string") return null;
  return valor.trim().toLowerCase().replace(/\.+$/, "");
}

function encontrarColumna(filaEncabezados, nombresAceptados) {
  let columna = null;
  filaEncabezados.eachCell((celda, numeroColumna) => {
    const valor = normalizarEncabezado(celda.value);
    if (valor && nombresAceptados.includes(valor)) columna = numeroColumna;
  });
  return columna;
}

function encontrarColumnaPrecios(fila) {
  return encontrarColumna(fila, ["precio", "precios"]);
}

function encontrarColumnaItemNo(fila) {
  return encontrarColumna(fila, ["item no", "item number"]);
}

function normalizarItemNo(valor) {
  if (valor && typeof valor === "object" && "result" in valor) valor = valor.result;
  return String(valor).trim().toLowerCase();
}

function tieneColumnasRequeridas(hoja) {
  const filaEncabezados = hoja.getRow(1);
  return Boolean(encontrarColumnaItemNo(filaEncabezados) && encontrarColumnaPrecios(filaEncabezados));
}

function workbookTieneColumnasRequeridas(workbook) {
  return workbook.worksheets.some((hoja) => tieneColumnasRequeridas(hoja));
}

function repreciarPorItemNo(workbookBase, workbookAModificar, porcentaje) {
  const factor = 1 + porcentaje / 100;
  let hojasActualizadas = 0;
  let filasOmitidas = 0;

  workbookAModificar.eachSheet((hojaAModificar) => {
    const colItemNoMod = encontrarColumnaItemNo(hojaAModificar.getRow(1));
    const colPreciosMod = encontrarColumnaPrecios(hojaAModificar.getRow(1));
    if (!colItemNoMod || !colPreciosMod) return;

    let hojaBase = workbookBase.getWorksheet(hojaAModificar.name);
    if (!hojaBase && workbookBase.worksheets.length === 1) {
      hojaBase = workbookBase.worksheets[0];
    }
    if (!hojaBase) return;

    const colItemNoBase = encontrarColumnaItemNo(hojaBase.getRow(1));
    const colPreciosBase = encontrarColumnaPrecios(hojaBase.getRow(1));
    if (!colItemNoBase || !colPreciosBase) return;

    const mapaPrecios = new Map();
    for (let f = 2; f <= hojaBase.rowCount; f++) {
      const itemNo = hojaBase.getRow(f).getCell(colItemNoBase).value;
      if (itemNo === null || itemNo === undefined || itemNo === "") continue;
      mapaPrecios.set(normalizarItemNo(itemNo), hojaBase.getRow(f).getCell(colPreciosBase).value);
    }

    for (let f = 2; f <= hojaAModificar.rowCount; f++) {
      const fila = hojaAModificar.getRow(f);
      const itemNo = fila.getCell(colItemNoMod).value;
      if (itemNo === null || itemNo === undefined || itemNo === "") continue;

      const precioBase = mapaPrecios.get(normalizarItemNo(itemNo));
      if (typeof precioBase !== "number") {
        filasOmitidas++;
        continue;
      }
      fila.getCell(colPreciosMod).value = precioBase * factor;
    }

    hojasActualizadas++;
  });

  return { hojasActualizadas, filasOmitidas };
}

async function repreciarDesdeBuffers(bufferBase, bufferAModificar, porcentaje) {
  const workbookBase = new ExcelJS.Workbook();
  await workbookBase.xlsx.load(bufferBase);

  const workbookAModificar = new ExcelJS.Workbook();
  await workbookAModificar.xlsx.load(bufferAModificar);

  const resultado = repreciarPorItemNo(workbookBase, workbookAModificar, porcentaje);
  const buffer = await workbookAModificar.xlsx.writeBuffer();

  return { buffer, ...resultado };
}

module.exports = {
  repreciarPorItemNo,
  repreciarDesdeBuffers,
  encontrarColumnaPrecios,
  encontrarColumnaItemNo,
  workbookTieneColumnasRequeridas,
};
