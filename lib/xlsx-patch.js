const JSZip = require("jszip");

const LETRAS_COLUMNA = { 1: "A", 14: "N", 15: "O", 16: "P", 17: "Q", 18: "R", 19: "S", 20: "T", 21: "U" };

function letraColumna(numeroColumna) {
  const letra = LETRAS_COLUMNA[numeroColumna];
  if (!letra) throw new Error(`Columna ${numeroColumna} no soportada por el parche.`);
  return letra;
}

function escaparXml(texto) {
  return String(texto).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Encuentra el archivo xl/worksheets/sheetN.xml que corresponde a la PRIMERA hoja
// del libro (en el orden en que aparece en xl/workbook.xml), igual que
// workbook.worksheets[0] de ExcelJS.
async function resolverPrimeraHojaXml(zip) {
  const workbookXml = await zip.file("xl/workbook.xml").async("string");
  const primerSheet = /<sheets>\s*<sheet[^>]*\/>/.exec(workbookXml);
  if (!primerSheet) throw new Error("No se pudo encontrar la primera hoja en workbook.xml.");
  const rIdMatch = /r:id="([^"]+)"/.exec(primerSheet[0]);
  if (!rIdMatch) throw new Error("No se pudo determinar el r:id de la primera hoja.");
  const rId = rIdMatch[1];

  const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const relRegex = new RegExp(`<Relationship[^>]*Id="${rId}"[^>]*Target="([^"]+)"`);
  const relMatch = relRegex.exec(relsXml);
  if (!relMatch) throw new Error(`No se encontró la relación ${rId} en workbook.xml.rels.`);

  let target = relMatch[1];
  if (!target.startsWith("worksheets/")) target = `worksheets/${target}`;
  return `xl/${target}`;
}

function extraerFila(sheetXml, numeroFila) {
  const regex = new RegExp(`<row r="${numeroFila}"[^>]*>([\\s\\S]*?)</row>`);
  const match = regex.exec(sheetXml);
  if (!match) return null;
  return { coincidenciaCompleta: match[0], contenido: match[1], indice: match.index };
}

// Saca (si existe) el atributo de estilo "s" de una celda existente en esa posicion,
// para no perder el color/formato que ya tenia la columna (ej: el amarillo/verde
// de las columnas N/O en la plantilla).
function extraerEstiloDeCelda(contenidoFila, refCelda) {
  const regex = new RegExp(`<c r="${refCelda}"[^>]*?\\ss="(\\d+)"[^>]*?(?:/>|>[\\s\\S]*?</c>)`);
  const match = regex.exec(contenidoFila);
  return match ? match[1] : null;
}

function quitarCeldaExistente(contenidoFila, refCelda) {
  const regex = new RegExp(`<c r="${refCelda}"[^>]*?(?:/>|>[\\s\\S]*?</c>)`);
  return contenidoFila.replace(regex, "");
}

function construirCeldaFormula(refCelda, formula, estilo) {
  const sAttr = estilo ? ` s="${estilo}"` : "";
  return `<c r="${refCelda}"${sAttr}><f>${escaparXml(formula)}</f></c>`;
}

function construirCeldaValor(refCelda, valor, estilo) {
  const sAttr = estilo ? ` s="${estilo}"` : "";
  return `<c r="${refCelda}"${sAttr}><v>${valor}</v></c>`;
}

function aplicarCeldasAFila(sheetXml, numeroFila, celdas) {
  const fila = extraerFila(sheetXml, numeroFila);
  if (!fila) {
    throw new Error(`No se encontró la fila ${numeroFila} en la hoja (se esperaba que ya existiera con datos).`);
  }

  let contenido = fila.contenido;
  let nuevasCeldas = "";

  Object.entries(celdas).forEach(([columna, formula]) => {
    const letra = letraColumna(Number(columna));
    const refCelda = `${letra}${numeroFila}`;
    const estilo = extraerEstiloDeCelda(contenido, refCelda);
    contenido = quitarCeldaExistente(contenido, refCelda);
    nuevasCeldas += construirCeldaFormula(refCelda, formula, estilo);
  });

  const filaNueva = `<row r="${numeroFila}"${fila.coincidenciaCompleta.match(/<row r="\d+"([^>]*)>/)[1]}>${contenido}${nuevasCeldas}</row>`;
  return sheetXml.slice(0, fila.indice) + filaNueva + sheetXml.slice(fila.indice + fila.coincidenciaCompleta.length);
}

function agregarFilasNuevas(sheetXml, filasPie) {
  const porFila = new Map();
  filasPie.forEach((c) => {
    if (!porFila.has(c.fila)) porFila.set(c.fila, []);
    porFila.get(c.fila).push(c);
  });

  let filasXml = "";
  porFila.forEach((celdas, numeroFila) => {
    let celdasXml = "";
    celdas.forEach((c) => {
      const letra = letraColumna(c.columna);
      const refCelda = `${letra}${numeroFila}`;
      celdasXml +=
        c.formula !== undefined
          ? construirCeldaFormula(refCelda, c.formula, null)
          : construirCeldaValor(refCelda, c.valor, null);
    });
    filasXml += `<row r="${numeroFila}">${celdasXml}</row>`;
  });

  const idx = sheetXml.indexOf("</sheetData>");
  if (idx === -1) throw new Error("No se encontró </sheetData> en la hoja.");
  return sheetXml.slice(0, idx) + filasXml + sheetXml.slice(idx);
}

function actualizarDimension(sheetXml, ultimaFila, ultimaColumnaLetra) {
  const regex = /<dimension ref="([^"]+)"\/>/;
  const match = regex.exec(sheetXml);
  if (!match) return sheetXml;
  const inicio = match[1].split(":")[0];
  return sheetXml.replace(regex, `<dimension ref="${inicio}:${ultimaColumnaLetra}${ultimaFila}"/>`);
}

// Si agregamos formulas nuevas, calcChain.xml queda desactualizado/incompleto.
// Hay que sacarlo del todo (Excel recalcula sin problema sin el cache), y tambien
// sacar las referencias a el en Content_Types y en workbook.xml.rels: dejar una
// referencia a un archivo que ya no existe es exactamente el tipo de inconsistencia
// que hace que Excel diga "el libro necesito reparaciones".
async function quitarCalcChain(zip) {
  if (!zip.file("xl/calcChain.xml")) return;
  zip.remove("xl/calcChain.xml");

  const contentTypesPath = "[Content_Types].xml";
  let contentTypes = await zip.file(contentTypesPath).async("string");
  contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, "");
  zip.file(contentTypesPath, contentTypes);

  const relsPath = "xl/_rels/workbook.xml.rels";
  let rels = await zip.file(relsPath).async("string");
  rels = rels.replace(/<Relationship[^>]*Target="calcChain\.xml"[^>]*\/>/, "");
  zip.file(relsPath, rels);
}

// Las formulas que agregamos no tienen un valor "cacheado" (<v>): Excel/WPS
// las calcula recien cuando recalcula la hoja. Si el libro tiene el calculo
// automatico desactivado (o si la ausencia de calcChain.xml hace que no se
// dispare un recalculo), esas celdas se ven en blanco hasta que alguien
// aprieta F9 a mano. Forzamos fullCalcOnLoad="1" para que siempre se
// recalcule todo apenas se abre el archivo, sin depender de esa configuracion.
async function forzarRecalculoAlAbrir(zip) {
  const path = "xl/workbook.xml";
  let workbookXml = await zip.file(path).async("string");

  if (/<calcPr\b[^>]*\/>/.test(workbookXml)) {
    workbookXml = workbookXml.replace(/<calcPr\b([^>]*)\/>/, (coincidencia, atributos) => {
      let nuevosAtributos = atributos.replace(/\sfullCalcOnLoad="[^"]*"/, "");
      nuevosAtributos = nuevosAtributos.replace(/\scalcMode="[^"]*"/, "");
      return `<calcPr${nuevosAtributos} calcMode="auto" fullCalcOnLoad="1"/>`;
    });
  } else {
    // no habia <calcPr>, lo agregamos (va despues de </sheets> o </definedNames>)
    workbookXml = workbookXml.replace(
      /(<\/definedNames>|<\/sheets>)/,
      `$1<calcPr calcMode="auto" fullCalcOnLoad="1"/>`
    );
  }

  zip.file(path, workbookXml);
}

async function aplicarFormulasAlXlsx(bufferOriginal, plan) {
  const zip = await JSZip.loadAsync(bufferOriginal);
  const rutaHoja = await resolverPrimeraHojaXml(zip);

  const archivoHoja = zip.file(rutaHoja);
  if (!archivoHoja) throw new Error(`No se encontró ${rutaHoja} dentro del archivo.`);

  let sheetXml = await archivoHoja.async("string");

  plan.celdasProducto.forEach(({ fila, celdas }) => {
    sheetXml = aplicarCeldasAFila(sheetXml, fila, celdas);
  });

  sheetXml = agregarFilasNuevas(sheetXml, plan.celdasPie);
  sheetXml = actualizarDimension(sheetXml, plan.filaX, "U");

  zip.file(rutaHoja, sheetXml);
  await quitarCalcChain(zip);
  await forzarRecalculoAlAbrir(zip);

  return zip.generateAsync({ type: "nodebuffer" });
}

module.exports = { aplicarFormulasAlXlsx };
