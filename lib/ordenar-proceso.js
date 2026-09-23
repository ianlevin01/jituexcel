const crypto = require("crypto");
const ExcelJS = require("exceljs");
const {
  getObjectBuffer,
  getObjectSize,
  putObjectBuffer,
  getPresignedDownloadUrl,
  deleteObject,
} = require("./s3");
const { reordenarPorCategoria } = require("./excel-images");
const { clasificarImagenes } = require("./openai");
const { actualizarJob } = require("./jobs");

async function procesarOrdenarExcel(jobId, key, tamanioEsperado) {
  const inicio = Date.now();
  const log = (msg) => console.log(`[ordenar-excel:${jobId}] +${Date.now() - inicio}ms ${msg}`);

  try {
    await actualizarJob(jobId, { status: "procesando" });

    if (tamanioEsperado) {
      let tamanioReal;
      try {
        tamanioReal = await getObjectSize(key);
      } catch (err) {
        log(`ERROR verificando tamaño en S3: ${err.name}: ${err.message}`);
        await actualizarJob(jobId, {
          status: "error",
          error: "La subida a S3 falló por completo (no llegó ningún dato). Probá de nuevo, puede ser por una conexión lenta o inestable.",
        });
        return;
      }
      log(`tamanio real en S3: ${tamanioReal} (esperado ${tamanioEsperado})`);
      if (tamanioReal !== tamanioEsperado) {
        await deleteObject(key).catch(() => {});
        await actualizarJob(jobId, {
          status: "error",
          error: `La subida a S3 quedó incompleta (se recibieron ${tamanioReal} de ${tamanioEsperado} bytes). Probá de nuevo.`,
        });
        return;
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
      await actualizarJob(jobId, { status: "error", error: `El archivo no es un Excel válido (${err.message}).` });
      return;
    }

    let resultado;
    try {
      log("iniciando reordenamiento por categoria (incluye llamadas a OpenAI)...");
      resultado = await reordenarPorCategoria(workbook, clasificarImagenes);
      log(`reordenamiento ok, ${resultado.resumen.length} grupo(s)`);
    } catch (err) {
      log(`ERROR en reordenamiento: ${err.stack || err.message}`);
      await deleteObject(key).catch(() => {});
      await actualizarJob(jobId, { status: "error", error: err.message });
      return;
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

    log("listo");
    await actualizarJob(jobId, { status: "listo", downloadUrl, resumen: resultado.resumen });
  } catch (err) {
    log(`ERROR no manejado: ${err.stack || err.message}`);
    console.error(`Error procesando job ${jobId}:`, err);
    await actualizarJob(jobId, { status: "error", error: `${err.name}: ${err.message}` }).catch(() => {});
  }
}

module.exports = { procesarOrdenarExcel };
