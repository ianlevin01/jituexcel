const ExcelJS = require("exceljs");

const SIN_IMAGEN = "__sin_imagen__";
const EMU_POR_PIXEL = 9525;
const TAMANIO_POR_DEFECTO = { width: 80, height: 80 };

function calcularExt(range) {
  if (range.ext && range.ext.width && range.ext.height) return range.ext;

  const tl = range.tl;
  const br = range.br;
  if (!tl || !br) return TAMANIO_POR_DEFECTO;

  const nativeColOffTl = tl.nativeColOff || 0;
  const nativeRowOffTl = tl.nativeRowOff || 0;
  const nativeColOffBr = br.nativeColOff || 0;
  const nativeRowOffBr = br.nativeRowOff || 0;
  const filasDeDiferencia = Math.floor(br.row) - Math.floor(tl.row);
  const columnasDeDiferencia = Math.floor(br.col) - Math.floor(tl.col);

  if (filasDeDiferencia > 0 || columnasDeDiferencia > 0) {
    return TAMANIO_POR_DEFECTO;
  }

  const width = (nativeColOffBr - nativeColOffTl) / EMU_POR_PIXEL;
  const height = (nativeRowOffBr - nativeRowOffTl) / EMU_POR_PIXEL;

  if (!(width > 0) || !(height > 0)) return TAMANIO_POR_DEFECTO;
  return { width, height };
}

function ultimaFilaConDatos(worksheet) {
  let ultima = 0;
  worksheet.eachRow({ includeEmpty: false }, (fila, numeroFila) => {
    ultima = Math.max(ultima, numeroFila);
  });
  return ultima;
}

function extraerImagenesPorFila(workbook, worksheet, filaMaxima) {
  const porFila = new Map();
  let descartadasPorHuerfanas = 0;
  worksheet.getImages().forEach((img) => {
    const fila = Math.floor(img.range.tl.row) + 1;
    if (fila > filaMaxima) {
      descartadasPorHuerfanas++;
      return;
    }
    if (porFila.has(fila)) return;
    const media = workbook.model.media[img.imageId];
    if (media && media.buffer) {
      porFila.set(fila, {
        extension: media.extension,
        buffer: media.buffer,
        colAncla: img.range.tl.col,
        ext: calcularExt(img.range),
      });
    }
  });
  return { porFila, descartadasPorHuerfanas };
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

  // hay que calcular esto ANTES de tocar los anclajes de las imagenes: leer la posicion
  // de un anclaje (range.tl.row) tiene el efecto secundario de "inflar" rowCount en ExcelJS.
  const filaMaxima = ultimaFilaConDatos(hojaOrigen);

  const { porFila: imagenesPorFila, descartadasPorHuerfanas } = extraerImagenesPorFila(
    workbook,
    hojaOrigen,
    filaMaxima
  );

  const filasConImagen = [];
  for (let f = filaEncabezado + 1; f <= filaMaxima; f++) {
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
  for (let f = filaEncabezado + 1; f <= filaMaxima; f++) {
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
      nuevaHoja.addImage(nuevoImageId, {
        tl: { col: img.colAncla, row: filaDestinoNum - 1 },
        ext: img.ext,
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

  if (descartadasPorHuerfanas > 0) {
    resumen.push({
      categoria: `(descartadas: ${descartadasPorHuerfanas} imagen(es) ancladas fuera del rango de datos)`,
      cantidad: descartadasPorHuerfanas,
    });
  }

  return { workbook: nuevoWorkbook, resumen };
}

module.exports = { reordenarPorCategoria };
