const SIN_IMAGEN = "__sin_imagen__";
const EMU_POR_PIXEL = 9525;
const TAMANIO_POR_DEFECTO = { width: 80, height: 80 };
const MINIMO_CELDAS_CON_VALOR = 2;

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
    if (!media || !media.buffer) return;
    porFila.set(fila, {
      colAncla: img.range.tl.col,
      ext: calcularExt(img.range),
      buffer: media.buffer,
      extension: media.extension,
    });
  });
  return { porFila, descartadasPorHuerfanas };
}

function capturarFila(fila) {
  const celdas = [];
  let celdasConValor = 0;
  fila.eachCell({ includeEmpty: true }, (celda, numeroColumna) => {
    celdas.push({ numeroColumna, value: celda.value, style: celda.style });
    if (celda.value !== null && celda.value !== undefined && celda.value !== "") celdasConValor++;
  });
  return { height: fila.height, celdas, celdasConValor };
}

function limpiarFila(fila) {
  fila.eachCell({ includeEmpty: true }, (celda) => {
    celda.value = null;
  });
}

function aplicarSnapshot(fila, snapshot) {
  fila.height = snapshot.height;
  snapshot.celdas.forEach(({ numeroColumna, value, style }) => {
    const celda = fila.getCell(numeroColumna);
    celda.value = value;
    celda.style = style;
  });
}

async function reordenarPorCategoria(workbook, clasificar) {
  const filaEncabezado = 1;
  const worksheet = workbook.worksheets[0];

  // hay que calcular esto ANTES de tocar los anclajes de las imagenes: leer la posicion
  // de un anclaje (range.tl.row) tiene el efecto secundario de "inflar" rowCount en ExcelJS.
  const filaMaxima = ultimaFilaConDatos(worksheet);

  const { porFila: imagenesPorFila, descartadasPorHuerfanas } = extraerImagenesPorFila(workbook, worksheet, filaMaxima);

  // Snapshot de TODAS las filas antes de tocar nada (vamos a reescribir sobre la misma hoja).
  const snapshots = new Map();
  for (let f = filaEncabezado + 1; f <= filaMaxima; f++) {
    snapshots.set(f, capturarFila(worksheet.getRow(f)));
  }

  // Solo se consideran filas "reales": con imagen Y con informacion de verdad
  // (no una fila vacia con solo un precio suelto).
  let descartadasPorSinInfo = 0;
  const filasConImagen = [];
  for (let f = filaEncabezado + 1; f <= filaMaxima; f++) {
    if (!imagenesPorFila.has(f)) continue;
    if (snapshots.get(f).celdasConValor < MINIMO_CELDAS_CON_VALOR) {
      descartadasPorSinInfo++;
      continue;
    }
    filasConImagen.push(f);
  }

  if (filasConImagen.length === 0) {
    throw new Error("No se encontraron filas con imagen y datos para poder agrupar.");
  }

  const imagenesParaClasificar = filasConImagen.map((f) => imagenesPorFila.get(f));
  const categorias = await clasificar(imagenesParaClasificar);

  const categoriaPorFila = new Map();
  filasConImagen.forEach((f, i) => categoriaPorFila.set(f, categorias[i]));

  const ordenGrupos = [];
  const filasPorGrupo = new Map();
  filasConImagen.forEach((f) => {
    const categoria = categoriaPorFila.get(f) || SIN_IMAGEN;
    if (!filasPorGrupo.has(categoria)) {
      filasPorGrupo.set(categoria, []);
      ordenGrupos.push(categoria);
    }
    filasPorGrupo.get(categoria).push(f);
  });

  const nuevoOrdenFilas = ordenGrupos.flatMap((cat) => filasPorGrupo.get(cat));

  // Limpiar todas las anclas de imagenes actuales; se vuelven a incrustar (copia fresca)
  // en su nueva posicion. No se puede reusar el mismo imageId en dos anclas distintas:
  // ExcelJS mezcla el contenido cuando la misma imagen esta anclada en mas de un lugar.
  worksheet._media = [];

  nuevoOrdenFilas.forEach((filaOrigenNum, indice) => {
    const filaDestinoNum = filaEncabezado + 1 + indice;
    aplicarSnapshot(worksheet.getRow(filaDestinoNum), snapshots.get(filaOrigenNum));

    const img = imagenesPorFila.get(filaOrigenNum);
    const nuevoImageId = workbook.addImage({ buffer: img.buffer, extension: img.extension });
    worksheet.addImage(nuevoImageId, {
      tl: { col: img.colAncla, row: filaDestinoNum - 1 },
      ext: img.ext,
    });
  });

  // Las filas descartadas (o las que sobraron al achicarse la hoja) quedan con basura
  // vieja mas alla de la ultima fila valida: hay que vaciarlas.
  for (let f = filaEncabezado + 1 + nuevoOrdenFilas.length; f <= filaMaxima; f++) {
    limpiarFila(worksheet.getRow(f));
  }

  const resumen = ordenGrupos.map((cat) => ({
    categoria: cat === SIN_IMAGEN ? "(sin imagen)" : cat,
    cantidad: filasPorGrupo.get(cat).length,
  }));

  if (descartadasPorSinInfo > 0) {
    resumen.push({
      categoria: `(descartadas: ${descartadasPorSinInfo} fila(s) con imagen pero sin información)`,
      cantidad: descartadasPorSinInfo,
    });
  }

  if (descartadasPorHuerfanas > 0) {
    resumen.push({
      categoria: `(descartadas: ${descartadasPorHuerfanas} imagen(es) ancladas fuera del rango de datos)`,
      cantidad: descartadasPorHuerfanas,
    });
  }

  return { workbook, resumen };
}

module.exports = { reordenarPorCategoria };
