const ExcelJS = require("exceljs");

const SIN_IMAGEN = "__sin_imagen__";

function extraerImagenesPorFila(workbook, worksheet) {
  const porFila = new Map();
  worksheet.getImages().forEach((img) => {
    const fila = img.range.tl.row + 1;
    if (porFila.has(fila)) return;
    const media = workbook.model.media[img.imageId];
    if (media && media.buffer) {
      porFila.set(fila, {
        extension: media.extension,
        buffer: media.buffer,
        anchor: img.range,
      });
    }
  });
  return porFila;
}

function copiarFila(filaOrigen, filaDestino) {
  filaDestino.height = filaOrigen.height;
  filaOrigen.eachCell({ includeEmpty: true }, (celda, numeroColumna) => {
    const destCelda = filaDestino.getCell(numeroColumna);
    destCelda.value = celda.value;
    destCelda.style = celda.style;
  });
}

async function reordenarPorCategoria(workbook, clasificar) {
  const filaEncabezado = 1;
  const hojaOrigen = workbook.worksheets[0];
  const imagenesPorFila = extraerImagenesPorFila(workbook, hojaOrigen);

  const filasConImagen = [];
  for (let f = filaEncabezado + 1; f <= hojaOrigen.rowCount; f++) {
    if (imagenesPorFila.has(f)) filasConImagen.push(f);
  }

  if (filasConImagen.length === 0) {
    throw new Error("No se encontraron imágenes en el excel para poder agrupar.");
  }

  const imagenesParaClasificar = filasConImagen.map((f) => imagenesPorFila.get(f));
  const categorias = await clasificar(imagenesParaClasificar);

  const categoriaPorFila = new Map();
  filasConImagen.forEach((f, i) => categoriaPorFila.set(f, categorias[i]));

  const ordenGrupos = [];
  const filasPorGrupo = new Map();
  for (let f = filaEncabezado + 1; f <= hojaOrigen.rowCount; f++) {
    const categoria = categoriaPorFila.get(f) || SIN_IMAGEN;
    if (!filasPorGrupo.has(categoria)) {
      filasPorGrupo.set(categoria, []);
      ordenGrupos.push(categoria);
    }
    filasPorGrupo.get(categoria).push(f);
  }

  const nuevoOrdenFilas = ordenGrupos.flatMap((cat) => filasPorGrupo.get(cat));

  const nuevoWorkbook = new ExcelJS.Workbook();
  const nuevaHoja = nuevoWorkbook.addWorksheet(hojaOrigen.name);
  nuevaHoja.columns = hojaOrigen.columns.map((c) => ({ width: c.width }));

  copiarFila(hojaOrigen.getRow(filaEncabezado), nuevaHoja.getRow(filaEncabezado));

  nuevoOrdenFilas.forEach((filaOrigenNum, indice) => {
    const filaDestinoNum = filaEncabezado + 1 + indice;
    copiarFila(hojaOrigen.getRow(filaOrigenNum), nuevaHoja.getRow(filaDestinoNum));

    const img = imagenesPorFila.get(filaOrigenNum);
    if (img) {
      const nuevoImageId = nuevoWorkbook.addImage({ buffer: img.buffer, extension: img.extension });
      const ext = img.anchor.ext || { width: 80, height: 80 };
      nuevaHoja.addImage(nuevoImageId, {
        tl: { col: img.anchor.tl.col, row: filaDestinoNum - 1 },
        ext,
      });
    }
  });

  for (let i = 1; i < workbook.worksheets.length; i++) {
    const hojaExtra = workbook.worksheets[i];
    const nuevaHojaExtra = nuevoWorkbook.addWorksheet(hojaExtra.name);
    hojaExtra.eachRow({ includeEmpty: true }, (fila, numeroFila) => {
      copiarFila(fila, nuevaHojaExtra.getRow(numeroFila));
    });
  }

  const resumen = ordenGrupos.map((cat) => ({
    categoria: cat === SIN_IMAGEN ? "(sin imagen)" : cat,
    cantidad: filasPorGrupo.get(cat).length,
  }));

  return { workbook: nuevoWorkbook, resumen };
}

module.exports = { reordenarPorCategoria };
