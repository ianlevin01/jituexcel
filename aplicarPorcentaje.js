const fs = require("fs");
const path = require("path");
const readline = require("readline");
const ExcelJS = require("exceljs");

const NOMBRES_COLUMNA_ACEPTADOS = ["precio", "precios"];
const ARCHIVO_BASE = "base.xlsx";
const ARCHIVO_A_MODIFICAR = "actualizados.xlsx";

function pedirPorcentaje() {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve, reject) => {
    rl.question("¿Qué porcentaje querés aplicarle a la columna Precio? ", (respuesta) => {
      rl.close();
      const porcentaje = Number(respuesta.replace(",", "."));
      if (Number.isNaN(porcentaje)) {
        reject(new Error(`"${respuesta}" no es un número válido.`));
        return;
      }
      resolve(porcentaje);
    });
  });
}

function encontrarColumnaPrecios(filaEncabezados) {
  let columna = null;
  filaEncabezados.eachCell((celda, numeroColumna) => {
    if (typeof celda.value === "string" && NOMBRES_COLUMNA_ACEPTADOS.includes(celda.value.trim().toLowerCase())) {
      columna = numeroColumna;
    }
  });
  return columna;
}

function aplicarPorcentaje(workbookBase, workbookAModificar, porcentaje) {
  const factor = 1 + porcentaje / 100;
  let hojasActualizadas = 0;
  let filasOmitidas = 0;

  workbookAModificar.eachSheet((hojaAModificar, idHoja) => {
    const columnaPreciosAModificar = encontrarColumnaPrecios(hojaAModificar.getRow(1));
    if (!columnaPreciosAModificar) return;

    let hojaBase = workbookBase.getWorksheet(hojaAModificar.name);
    if (!hojaBase && workbookBase.worksheets.length === 1) {
      hojaBase = workbookBase.worksheets[0];
    }
    if (!hojaBase) {
      console.warn(`Aviso: no se encontró en ${ARCHIVO_BASE} una hoja correspondiente a "${hojaAModificar.name}". Se omite.`);
      return;
    }

    const columnaPreciosBase = encontrarColumnaPrecios(hojaBase.getRow(1));
    if (!columnaPreciosBase) {
      console.warn(`Aviso: la hoja "${hojaBase.name}" de ${ARCHIVO_BASE} no tiene columna "Precio". Se omite.`);
      return;
    }

    for (let numeroFila = 2; numeroFila <= hojaAModificar.rowCount; numeroFila++) {
      const valorBase = hojaBase.getRow(numeroFila).getCell(columnaPreciosBase).value;
      if (typeof valorBase !== "number") {
        filasOmitidas++;
        continue;
      }
      hojaAModificar.getRow(numeroFila).getCell(columnaPreciosAModificar).value = valorBase * factor;
    }

    hojasActualizadas++;
  });

  return { hojasActualizadas, filasOmitidas };
}

async function main() {
  const rutaBase = path.resolve(ARCHIVO_BASE);
  const rutaAModificar = path.resolve(ARCHIVO_A_MODIFICAR);

  if (!fs.existsSync(rutaBase)) {
    console.error(`No se encontró el archivo: ${rutaBase}`);
    process.exit(1);
  }
  if (!fs.existsSync(rutaAModificar)) {
    console.error(`No se encontró el archivo: ${rutaAModificar}`);
    process.exit(1);
  }

  let porcentaje;
  try {
    porcentaje = await pedirPorcentaje();
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }

  const workbookBase = new ExcelJS.Workbook();
  await workbookBase.xlsx.readFile(rutaBase);

  const workbookAModificar = new ExcelJS.Workbook();
  await workbookAModificar.xlsx.readFile(rutaAModificar);

  const { hojasActualizadas, filasOmitidas } = aplicarPorcentaje(workbookBase, workbookAModificar, porcentaje);

  if (hojasActualizadas === 0) {
    console.error(`No se encontró ninguna columna "Precio" en común entre ${ARCHIVO_BASE} y ${ARCHIVO_A_MODIFICAR}.`);
    process.exit(1);
  }

  await workbookAModificar.xlsx.writeFile(rutaAModificar);

  console.log(`Listo. Se aplicó ${porcentaje}% a la columna "Precio" (precio base × ${1 + porcentaje / 100}).`);
  if (filasOmitidas > 0) {
    console.log(`Se omitieron ${filasOmitidas} fila(s) sin precio base numérico.`);
  }
  console.log(`Archivo actualizado: ${rutaAModificar}`);
}

main();
