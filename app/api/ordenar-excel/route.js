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
  try {
    const body = await request.json().catch(() => ({}));
    const { key, tamanioEsperado } = body;
    uploadKey = key;

    if (!key) {
      return Response.json({ error: "Falta la key del archivo subido." }, { status: 400 });
    }

    if (tamanioEsperado) {
      const tamanioReal = await getObjectSize(key);
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

    const buffer = await getObjectBuffer(key);

    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.load(buffer);
    } catch (err) {
      await deleteObject(key).catch(() => {});
      return Response.json({ error: `El archivo no es un Excel válido (${err.message}).` }, { status: 400 });
    }

    let resultado;
    try {
      resultado = await reordenarPorCategoria(workbook, clasificarImagenes);
    } catch (err) {
      await deleteObject(key).catch(() => {});
      return Response.json({ error: err.message }, { status: 400 });
    }

    const outputBuffer = await resultado.workbook.xlsx.writeBuffer();
    const outputKey = `outputs/ordenado-${crypto.randomUUID()}.xlsx`;
    await putObjectBuffer(outputKey, outputBuffer);
    const downloadUrl = await getPresignedDownloadUrl(outputKey, "ordenado.xlsx");

    deleteObject(key).catch(() => {});

    return Response.json({ downloadUrl, resumen: resultado.resumen });
  } catch (err) {
    console.error("Error en /api/ordenar-excel:", err);
    if (uploadKey) deleteObject(uploadKey).catch(() => {});
    return Response.json({ error: `${err.name}: ${err.message}` }, { status: 500 });
  }
}
