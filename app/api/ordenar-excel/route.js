const crypto = require("crypto");
const ExcelJS = require("exceljs");
const {
  getObjectBuffer,
  getObjectSize,
  putObjectBuffer,
  getPresignedDownloadUrl,
  deleteObject,
} = require("../../../lib/s3");
const { reordenarPorCategoria } = require("../../../lib/excel-images");
const { clasificarImagenes } = require("../../../lib/openai");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;

export async function POST(request) {
  let uploadKey;
  const inicio = Date.now();
  const log = (msg) => console.log(`[ordenar-excel] +${Date.now() - inicio}ms ${msg}`);

  try {
    const body = await request.json().catch(() => ({}));
    const { key, tamanioEsperado } = body;
    uploadKey = key;
    log(`request recibido, key=${key}, tamanioEsperado=${tamanioEsperado}`);

    if (!key) {
      return Response.json({ error: "Falta la key del archivo subido." }, { status: 400 });
    }

    if (tamanioEsperado) {
      const tamanioReal = await getObjectSize(key);
      log(`tamanio real en S3: ${tamanioReal}`);
      if (tamanioReal !== tamanioEsperado) {
        await deleteObject(key).catch(() => {});
        return Response.json(
          {
            error: `La subida a S3 quedó incompleta (se recibieron ${tamanioReal} de ${tamanioEsperado} bytes). Probá de nuevo, puede ser por una conexión lenta o inestable.`,
          },
          { status: 400 }
        );
      }
    }

    log("descargando buffer desde S3...");
    const buffer = await getObjectBuffer(key);
    log(`buffer descargado, ${buffer.length} bytes`);

    const workbook = new ExcelJS.Workbook();
    try {
      log("parseando excel con ExcelJS...");
      await workbook.xlsx.load(buffer);
      log("excel parseado ok");
    } catch (err) {
      log(`ERROR parseando excel: ${err.stack || err.message}`);
      await deleteObject(key).catch(() => {});
      return Response.json({ error: `El archivo no es un Excel válido (${err.message}).` }, { status: 400 });
    }

    let resultado;
    try {
      log("iniciando reordenamiento por categoria (incluye llamadas a OpenAI)...");
      resultado = await reordenarPorCategoria(workbook, clasificarImagenes);
      log(`reordenamiento ok, ${resultado.resumen.length} grupo(s)`);
    } catch (err) {
      log(`ERROR en reordenamiento: ${err.stack || err.message}`);
      await deleteObject(key).catch(() => {});
      return Response.json({ error: err.message }, { status: 400 });
    }

    log("generando buffer de salida...");
    const outputBuffer = await resultado.workbook.xlsx.writeBuffer();
    log(`buffer de salida generado, ${outputBuffer.length} bytes`);

    const outputKey = `outputs/ordenado-${crypto.randomUUID()}.xlsx`;
    log(`subiendo resultado a S3 (${outputKey})...`);
    await putObjectBuffer(outputKey, outputBuffer);
    log("subido, generando url de descarga...");
    const downloadUrl = await getPresignedDownloadUrl(outputKey, "ordenado.xlsx");

    deleteObject(key).catch((err) => log(`aviso: no se pudo borrar el temporal ${key}: ${err.message}`));

    log("listo, respondiendo 200");
    return Response.json({ downloadUrl, resumen: resultado.resumen });
  } catch (err) {
    log(`ERROR no manejado: ${err.stack || err.message}`);
    console.error("Error en /api/ordenar-excel:", err);
    if (uploadKey) deleteObject(uploadKey).catch(() => {});
    return Response.json({ error: `${err.name}: ${err.message}` }, { status: 500 });
  }
}
