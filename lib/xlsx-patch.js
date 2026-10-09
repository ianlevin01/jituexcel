const JSZip = require("jszip");
const { numeroALetraColumna, letraAColumna } = require("./excel-columnas");

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

// Una fila puede venir "autocerrada" (<row r="9" ht="31.2" .../>, sin
// celdas, solo con formato de alto) cuando esta vacia pero ya tiene estilo
// propio en la plantilla. Hay que reconocerla igual que una fila "abierta"
// (<row r="9" ...>celdas</row>): si no, el codigo la trata como si no
// existiera y mas adelante agrega una fila NUEVA con el mismo numero,
// duplicando el r="9" en la hoja. Se prueba primero la forma autocerrada
// (inequivoca, termina en "/>") para no confundirla con una fila abierta
// cuyo cierre real esta mas lejos en el documento.
function extraerFila(sheetXml, numeroFila) {
  const regexCerrada = new RegExp(`<row r="${numeroFila}"((?:\\s[^>]*?)?)/>`);
  const matchCerrada = regexCerrada.exec(sheetXml);
  if (matchCerrada) {
    return { coincidenciaCompleta: matchCerrada[0], atributos: matchCerrada[1], contenido: "", indice: matchCerrada.index };
  }

  const regexAbierta = new RegExp(`<row r="${numeroFila}"((?:\\s[^>]*?)?)>([\\s\\S]*?)</row>`);
  const matchAbierta = regexAbierta.exec(sheetXml);
  if (!matchAbierta) return null;
  return { coincidenciaCompleta: matchAbierta[0], atributos: matchAbierta[1], contenido: matchAbierta[2], indice: matchAbierta.index };
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

// Numeros van como celda numerica normal; cualquier otra cosa (texto) va como
// "inline string", que no necesita tocar sharedStrings.xml para nada.
function construirCeldaValor(refCelda, valor, estilo) {
  const sAttr = estilo ? ` s="${estilo}"` : "";
  if (typeof valor === "number" && Number.isFinite(valor)) {
    return `<c r="${refCelda}"${sAttr}><v>${valor}</v></c>`;
  }
  const texto = escaparXml(String(valor));
  return `<c r="${refCelda}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${texto}</t></is></c>`;
}

function construirCelda(refCelda, definicion, estilo) {
  return definicion.formula !== undefined
    ? construirCeldaFormula(refCelda, definicion.formula, estilo)
    : construirCeldaValor(refCelda, definicion.valor, estilo);
}

// Los <c r="..."> dentro de una fila tienen que quedar en orden ascendente
// de columna: si no, Excel detecta el archivo como danado al abrirlo (y al
// "reparar" puede llegar a descartar la celda que esta fuera de orden). No
// alcanza con reemplazar in-place cuando la celda ya existia ahi (la
// posicion ya es correcta), pero cuando es una celda NUEVA en una columna
// que antes estaba vacia (por ejemplo, porque un corrimiento la dejo libre)
// hay que insertarla en el lugar que le corresponde, no al final.
function insertarCeldaOrdenada(contenidoFila, celdaXml, numeroColumna) {
  const regex = /<c r="([A-Za-z]+)\d+"/g;
  let match;
  let posicionInsercion = contenidoFila.length;
  while ((match = regex.exec(contenidoFila))) {
    if (letraAColumna(match[1]) > numeroColumna) {
      posicionInsercion = match.index;
      break;
    }
  }
  return contenidoFila.slice(0, posicionInsercion) + celdaXml + contenidoFila.slice(posicionInsercion);
}

// celdas: { [numeroColumna]: {formula: "..."} | {valor: numero} }
function aplicarCeldasAFila(sheetXml, numeroFila, celdas) {
  const fila = extraerFila(sheetXml, numeroFila);
  if (!fila) {
    throw new Error(`No se encontró la fila ${numeroFila} en la hoja (se esperaba que ya existiera con datos).`);
  }

  let contenido = fila.contenido;

  Object.entries(celdas)
    .sort(([a], [b]) => Number(a) - Number(b))
    .forEach(([columna, definicion]) => {
      const columnaNum = Number(columna);
      const letra = numeroALetraColumna(columnaNum);
      const refCelda = `${letra}${numeroFila}`;
      const estilo = extraerEstiloDeCelda(contenido, refCelda);
      contenido = quitarCeldaExistente(contenido, refCelda);
      const celdaXml = construirCelda(refCelda, definicion, estilo);
      contenido = insertarCeldaOrdenada(contenido, celdaXml, columnaNum);
    });

  const filaNueva = `<row r="${numeroFila}"${fila.atributos}>${contenido}</row>`;
  return sheetXml.slice(0, fila.indice) + filaNueva + sheetXml.slice(fila.indice + fila.coincidenciaCompleta.length);
}

// Igual que insertarCeldaOrdenada pero para filas enteras: una fila nueva
// tiene que quedar en orden ascendente de numero de fila respecto a las que
// YA estan en la hoja (si no, mismo problema que con las celdas: Excel
// detecta el archivo como danado). Si no hay ninguna fila posterior, se
// agrega justo antes de </sheetData>.
function insertarFilaOrdenada(sheetXml, filaXml, numeroFila) {
  const regex = /<row r="(\d+)"/g;
  let match;
  let posicionInsercion = null;
  while ((match = regex.exec(sheetXml))) {
    if (Number(match[1]) > numeroFila) {
      posicionInsercion = match.index;
      break;
    }
  }
  if (posicionInsercion === null) {
    const idx = sheetXml.indexOf("</sheetData>");
    if (idx === -1) throw new Error("No se encontró </sheetData> en la hoja.");
    posicionInsercion = idx;
  }
  return sheetXml.slice(0, posicionInsercion) + filaXml + sheetXml.slice(posicionInsercion);
}

// filasPie: [{fila, columna, formula} | {fila, columna, valor}, ...]
// Solo para filas que NO existen todavia en la hoja (ver extraerFila: una
// fila "autocerrada" sin celdas SI existe, y tiene que ir por
// aplicarCeldasAFila, no por aca, para no duplicar el numero de fila).
function agregarFilasNuevas(sheetXml, filasPie) {
  const porFila = new Map();
  filasPie.forEach((c) => {
    if (!porFila.has(c.fila)) porFila.set(c.fila, []);
    porFila.get(c.fila).push(c);
  });

  let resultado = sheetXml;
  [...porFila.entries()]
    .sort(([a], [b]) => a - b)
    .forEach(([numeroFila, celdas]) => {
      let celdasXml = "";
      celdas
        .slice()
        .sort((a, b) => a.columna - b.columna)
        .forEach((c) => {
          const letra = numeroALetraColumna(c.columna);
          const refCelda = `${letra}${numeroFila}`;
          celdasXml += construirCelda(refCelda, c, null);
        });
      const filaXml = `<row r="${numeroFila}">${celdasXml}</row>`;
      resultado = insertarFilaOrdenada(resultado, filaXml, numeroFila);
    });

  return resultado;
}

function actualizarDimension(sheetXml, ultimaFila, ultimaColumnaLetra) {
  const regex = /<dimension ref="([^"]+)"\/>/;
  const match = regex.exec(sheetXml);
  if (!match) return sheetXml;
  const inicio = match[1].split(":")[0];
  return sheetXml.replace(regex, `<dimension ref="${inicio}:${ultimaColumnaLetra}${ultimaFila}"/>`);
}

// Recalcula la dimension real escaneando todas las celdas que efectivamente
// quedaron en el XML, en vez de calcularlo "a mano": mas simple y a prueba
// de errores despues de correr columnas, donde la extension cambia por fila.
function recalcularDimensionReal(sheetXml) {
  const regex = /<c r="([A-Za-z]+)(\d+)"/g;
  let maxColumna = 0;
  let maxFila = 0;
  let match;
  while ((match = regex.exec(sheetXml))) {
    maxColumna = Math.max(maxColumna, letraAColumna(match[1]));
    maxFila = Math.max(maxFila, parseInt(match[2], 10));
  }
  if (maxColumna === 0 || maxFila === 0) return sheetXml;
  return actualizarDimension(sheetXml, maxFila, numeroALetraColumna(maxColumna));
}

// Corre una posicion a la derecha todas las celdas de esta fila que esten en
// la columna "columnaDesde" o mas a la derecha (ej: Q, R, S... -> R, S, T...).
// Se extraen todas primero y se reemplaza cada referencia dentro del string
// (son unicas por fila), asi que no hay riesgo de que una celda recien movida
// pise a otra que todavia no se movio.
function correrColumnasHaciaLaDerecha(contenidoFila, columnaDesde, numeroFila) {
  const regex = /<c r="([A-Za-z]+)(\d+)"/g;
  const coincidencias = [];
  let match;
  while ((match = regex.exec(contenidoFila))) {
    const columna = letraAColumna(match[1]);
    if (columna >= columnaDesde) {
      coincidencias.push({ refOriginal: match[0], columna });
    }
  }

  // Importante: procesar de la columna MAS ALTA a la MAS BAJA. Si se hiciera
  // al reves (izquierda a derecha), el texto recien renombrado (ej Q3->R3)
  // coincidiria con el nombre que la celda original siguiente todavia tiene
  // (la R3 de verdad, sin tocar), y el "replace" (que busca la primera
  // coincidencia en el string) terminaria renombrando la celda equivocada.
  // De derecha a izquierda, el nombre destino de cada paso ya quedo libre
  // (la celda que lo tenia se renombro en el paso anterior), asi que no hay
  // ninguna coincidencia ambigua posible.
  coincidencias.sort((a, b) => b.columna - a.columna);

  let resultado = contenidoFila;
  coincidencias.forEach(({ refOriginal, columna }) => {
    const nuevaLetra = numeroALetraColumna(columna + 1);
    resultado = resultado.replace(refOriginal, `<c r="${nuevaLetra}${numeroFila}"`);
  });

  return resultado;
}

// Dentro del texto de una formula (o de un atributo ref="Q4:Q8" de formula
// compartida), corre una posicion a la derecha cualquier referencia a celda
// que este en "columnaDesde" o mas a la derecha. No toca nombres de funcion
// (siempre seguidos de "(", nunca son referencia) ni identificadores de los
// que esta referencia es solo una parte (ej: no toca el "A1" de "DATA1").
function correrReferenciasEnFormula(texto, columnaDesde) {
  return texto.replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g, (coincidenciaCompleta, dolarCol, letras, dolarFila, digitos, indice, cadena) => {
    const anterior = cadena[indice - 1];
    if (anterior && /[A-Za-z0-9_]/.test(anterior)) return coincidenciaCompleta;
    const siguiente = cadena[indice + coincidenciaCompleta.length];
    if (siguiente === "(") return coincidenciaCompleta; // nombre de funcion (ej LOG10(...))

    const columna = letraAColumna(letras);
    if (columna < columnaDesde) return coincidenciaCompleta;
    const nuevaLetra = numeroALetraColumna(columna + 1);
    return `${dolarCol}${nuevaLetra}${dolarFila}${digitos}`;
  });
}

// Recorre TODAS las etiquetas <f> de la hoja (no solo las de las filas que se
// mueven): una formula en una columna que no se mueve puede referenciar una
// celda de una columna que si se mueve (ej: la columna S con "66/J4*R4"
// referencia R, que se corre a S al insertar en Q). Hay que actualizar tanto
// el texto de la formula como el atributo "ref" de las formulas compartidas
// (el rango declarado en la celda que define el grupo, ej ref="N4:N8").
function correrReferenciasDeFormulasEnHoja(sheetXml, columnaDesde) {
  return sheetXml.replace(/<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/g, (coincidenciaCompleta, atributos, contenido) => {
    const nuevosAtributos = atributos.replace(/ref="([^"]+)"/, (m, rango) => `ref="${correrReferenciasEnFormula(rango, columnaDesde)}"`);
    if (contenido === undefined) return `<f${nuevosAtributos}/>`;
    const nuevoContenido = correrReferenciasEnFormula(contenido, columnaDesde);
    return `<f${nuevosAtributos}>${nuevoContenido}</f>`;
  });
}

// Columna maxima que tiene CONTENIDO REAL en la hoja (escaneando las celdas
// que efectivamente existen), antes de correr nada. Sirve para distinguir los
// <col> de ancho/estilo que describen columnas reales de la plantilla (que
// hay que correr junto con sus celdas) de los <col> "catch-all" que Excel usa
// para el ancho por defecto del resto de la hoja hasta la columna 16384
// (esos NO hay que tocarlos: correrlos se saldria del limite de columnas).
function columnaMaximaReal(sheetXml) {
  const regex = /<c r="([A-Za-z]+)\d+"/g;
  let maxColumna = 0;
  let match;
  while ((match = regex.exec(sheetXml))) {
    maxColumna = Math.max(maxColumna, letraAColumna(match[1]));
  }
  return maxColumna;
}

// Corre (si corresponde) los <col min="X" max="Y" .../> que definen ancho/
// estilo por columna, y los <mergeCell ref="A1:M1"/>. En los <col>, solo se
// corren los rangos que quedan COMPLETAMENTE dentro del contenido real de la
// hoja (ver columnaMaximaReal); en los merge, cada extremo se corre de forma
// independiente (como una referencia de formula), asi que un merge que
// "cruza" la columna insertada se extiende para absorberla, igual que haria
// Excel al insertar una columna real.
function correrColumnasEnColsYMerges(sheetXml, columnaDesde, maxColumnaReal) {
  let resultado = sheetXml.replace(/<col min="(\d+)" max="(\d+)"([^>]*)\/>/g, (coincidencia, min, max, resto) => {
    const minNum = Number(min);
    const maxNum = Number(max);
    if (minNum >= columnaDesde && maxNum <= maxColumnaReal) {
      return `<col min="${minNum + 1}" max="${maxNum + 1}"${resto}/>`;
    }
    // El "catch-all" que arranca justo donde termina el contenido real (ej.
    // min=23 cuando maxColumnaReal=22) tiene que arrancar una columna mas
    // adelante para no superponerse con el rango de arriba, que ya se corrio
    // (si no, quedarian dos <col> distintos describiendo la misma columna).
    // El "max" de ese catch-all NO se toca: es un limite lejano (puede ser
    // 16384, el maximo de Excel) y correrlo lo mandaria fuera de rango.
    if (minNum === maxColumnaReal + 1 && minNum >= columnaDesde) {
      return `<col min="${minNum + 1}" max="${max}"${resto}/>`;
    }
    return coincidencia;
  });

  resultado = resultado.replace(/<mergeCell ref="([^"]+)"\/>/g, (coincidencia, ref) => {
    return `<mergeCell ref="${correrReferenciasEnFormula(ref, columnaDesde)}"/>`;
  });

  return resultado;
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

async function cargarHojaPrincipal(bufferOriginal) {
  const zip = await JSZip.loadAsync(bufferOriginal);
  const rutaHoja = await resolverPrimeraHojaXml(zip);
  const archivoHoja = zip.file(rutaHoja);
  if (!archivoHoja) throw new Error(`No se encontró ${rutaHoja} dentro del archivo.`);
  const sheetXml = await archivoHoja.async("string");
  return { zip, rutaHoja, sheetXml };
}

async function cerrarYGenerar(zip, rutaHoja, sheetXml) {
  zip.file(rutaHoja, sheetXml);
  await quitarCalcChain(zip);
  await forzarRecalculoAlAbrir(zip);
  return zip.generateAsync({ type: "nodebuffer" });
}

// Usado por la funcionalidad de "Formulas": celdas por fila de producto (N-U)
// mas 3 filas nuevas al pie (celdasPie) con el calculo de "Amount of Goods".
async function aplicarFormulasAlXlsx(bufferOriginal, plan) {
  let { zip, rutaHoja, sheetXml } = await cargarHojaPrincipal(bufferOriginal);

  plan.celdasProducto.forEach(({ fila, celdas }) => {
    sheetXml = aplicarCeldasAFila(sheetXml, fila, celdas);
  });

  sheetXml = agregarFilasNuevas(sheetXml, plan.celdasPie);
  sheetXml = actualizarDimension(sheetXml, plan.filaX, "U");

  return cerrarYGenerar(zip, rutaHoja, sheetXml);
}

// Usado por la funcionalidad de "Excel A/B": en el excel A, corre una
// posicion a la derecha todo lo que haya desde la columna Q en adelante
// (en cada fila que tenga algo ahi), y pega en la columna Q (ya vacia) los
// valores de plan.valoresQ (vienen de la columna D del excel B).
async function insertarColumnaYPegarValores(bufferOriginal, plan) {
  let { zip, rutaHoja, sheetXml } = await cargarHojaPrincipal(bufferOriginal);

  const columnaQNum = letraAColumna(plan.columnaQ);
  const maxColumnaReal = columnaMaximaReal(sheetXml);

  // 0) actualizar las formulas que ya existan en el excel A: si alguna
  // referencia una celda de la columna Q en adelante (texto de formula o
  // "ref" de formula compartida), esa referencia tiene que apuntar a la
  // nueva posicion antes de mover ninguna celda. Tambien se corren los
  // anchos/estilos de columna y los merge de celdas, por la misma razon.
  sheetXml = correrReferenciasDeFormulasEnHoja(sheetXml, columnaQNum);
  sheetXml = correrColumnasEnColsYMerges(sheetXml, columnaQNum, maxColumnaReal);

  // 1) correr columnas, fila por fila, solo donde la fila existe y tiene datos
  for (let numeroFila = 1; numeroFila <= plan.filaMaximaA; numeroFila++) {
    const fila = extraerFila(sheetXml, numeroFila);
    if (!fila) continue;
    const contenidoCorrido = correrColumnasHaciaLaDerecha(fila.contenido, columnaQNum, numeroFila);
    if (contenidoCorrido === fila.contenido) continue;
    const filaNueva = `<row r="${numeroFila}"${fila.atributos}>${contenidoCorrido}</row>`;
    sheetXml = sheetXml.slice(0, fila.indice) + filaNueva + sheetXml.slice(fila.indice + fila.coincidenciaCompleta.length);
  }

  // 2) pegar los valores en la columna Q: si la fila ya existe (muy probable,
  // ya que recien se corrieron sus columnas) se inserta ahi; si no existe
  // (el excel B tiene mas filas que el A), se agrega una fila nueva.
  const paraFilasExistentes = [];
  const paraFilasNuevas = [];
  plan.valoresQ.forEach(({ fila, valor }) => {
    if (extraerFila(sheetXml, fila)) {
      paraFilasExistentes.push({ fila, valor });
    } else {
      paraFilasNuevas.push({ fila, columna: columnaQNum, valor });
    }
  });

  paraFilasExistentes.forEach(({ fila, valor }) => {
    sheetXml = aplicarCeldasAFila(sheetXml, fila, { [columnaQNum]: { valor } });
  });

  if (paraFilasNuevas.length > 0) {
    sheetXml = agregarFilasNuevas(sheetXml, paraFilasNuevas);
  }

  sheetXml = recalcularDimensionReal(sheetXml);

  return cerrarYGenerar(zip, rutaHoja, sheetXml);
}

module.exports = { aplicarFormulasAlXlsx, insertarColumnaYPegarValores };
