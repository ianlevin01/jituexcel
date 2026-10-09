const crypto = require("crypto");
const ExcelJS = require("exceljs");
const {
  getObjectBuffer,
  getObjectSize,
  putObjectBuffer,
  getPresignedDownloadUrl,
  deleteObject,
} = require("../../../lib/s3");
const { calcularPlanInsertarColumna } = require("../../../lib/excel-ab");
const { insertarColumnaYPegarValores } = require("../../../lib/xlsx-patch");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

async function verificarTamanio(key, tamanioEsperado) {
  if (!tamanioEsperado) return null;
  let tamanioReal;
  try {
    tamanioReal = await getObjectSize(key);
  } catch {
    return "La subida a S3 falló por completo (no llegó ningún dato). Probá de nuevo.";
  }
  if (tamanioReal !== tamanioEsperado) {
    await deleteObject(key).catch(() => {});
    return `La subida a S3 quedó incompleta (se recibieron ${tamanioReal} de ${tamanioEsperado} bytes). Probá de nuevo.`;
  }
  return null;
}

export async function POST(request) {
  let keyA, keyB;
  try {
    const body = await request.json().catch(() => ({}));
    ({ keyA, keyB } = body);
    const { tamanioEsperadoA, tamanioEsperadoB } = body;

    if (!keyA || !keyB) {
      return Response.json({ error: "Faltan los dos archivos (excel A y excel B)." }, { status: 400 });
    }

    const errorA = await verificarTamanio(keyA, tamanioEsperadoA);
    if (errorA) return Response.json({ error: `Excel A: ${errorA}` }, { status: 400 });

    const errorB = await verificarTamanio(keyB, tamanioEsperadoB);
    if (errorB) return Response.json({ error: `Excel B: ${errorB}` }, { status: 400 });

    const [bufferA, bufferB] = await Promise.all([getObjectBuffer(keyA), getObjectBuffer(keyB)]);

    const workbookA = new ExcelJS.Workbook();
    const workbookB = new ExcelJS.Workbook();
    try {
      await Promise.all([workbookA.xlsx.load(bufferA), workbookB.xlsx.load(bufferB)]);
    } catch (err) {
      await Promise.all([deleteObject(keyA).catch(() => {}), deleteObject(keyB).catch(() => {})]);
      return Response.json({ error: `Alguno de los dos excels no es válido (${err.message}).` }, { status: 400 });
    }

    let plan;
    try {
      plan = calcularPlanInsertarColumna(workbookA, workbookB);
    } catch (err) {
      await Promise.all([deleteObject(keyA).catch(() => {}), deleteObject(keyB).catch(() => {})]);
      return Response.json({ error: err.message }, { status: 400 });
    }

    let outputBuffer;
    try {
      // Igual que en "Formulas": nunca se vuelve a guardar el excel A via ExcelJS
      // completo (eso pierde imagenes/metadata que ExcelJS no entiende). Se
      // parchea quirurgicamente el XML del archivo A original.
      outputBuffer = await insertarColumnaYPegarValores(bufferA, plan);
    } catch (err) {
      console.error("Error aplicando el parche A/B:", err);
      return Response.json({ error: `No se pudo generar el excel: ${err.message}` }, { status: 500 });
    } finally {
      await Promise.all([deleteObject(keyA).catch(() => {}), deleteObject(keyB).catch(() => {})]);
    }

    const outputKey = `outputs/excel-ab-${crypto.randomUUID()}.xlsx`;
    await putObjectBuffer(outputKey, outputBuffer);
    const downloadUrl = await getPresignedDownloadUrl(outputKey, "actualizado.xlsx");

    const { valoresQ, codigosSinPrecio, ...resumen } = plan;
    return Response.json({ downloadUrl, resumen });
  } catch (err) {
    console.error("Error en /api/excel-ab:", err);
    if (keyA) deleteObject(keyA).catch(() => {});
    if (keyB) deleteObject(keyB).catch(() => {});
    return Response.json({ error: `${err.name}: ${err.message}` }, { status: 500 });
  }
}
