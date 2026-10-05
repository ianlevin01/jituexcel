const crypto = require("crypto");
const ExcelJS = require("exceljs");
const {
  getObjectBuffer,
  getObjectSize,
  putObjectBuffer,
  getPresignedDownloadUrl,
  deleteObject,
} = require("../../../lib/s3");
const { aplicarFormulas } = require("../../../lib/formulas-productos");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

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
      let tamanioReal;
      try {
        tamanioReal = await getObjectSize(key);
      } catch (err) {
        return Response.json(
          { error: "La subida a S3 falló por completo (no llegó ningún dato). Probá de nuevo." },
          { status: 400 }
        );
      }
      if (tamanioReal !== tamanioEsperado) {
        await deleteObject(key).catch(() => {});
        return Response.json(
          {
            error: `La subida a S3 quedó incompleta (se recibieron ${tamanioReal} de ${tamanioEsperado} bytes). Probá de nuevo.`,
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

    let resumen;
    try {
      resumen = aplicarFormulas(workbook);
    } catch (err) {
      await deleteObject(key).catch(() => {});
      return Response.json({ error: err.message }, { status: 400 });
    }

    const outputBuffer = await workbook.xlsx.writeBuffer();
    const outputKey = `outputs/formulas-${crypto.randomUUID()}.xlsx`;
    await putObjectBuffer(outputKey, outputBuffer);
    const downloadUrl = await getPresignedDownloadUrl(outputKey, "actualizado.xlsx");

    deleteObject(key).catch(() => {});

    return Response.json({ downloadUrl, resumen });
  } catch (err) {
    console.error("Error en /api/formulas:", err);
    if (uploadKey) deleteObject(uploadKey).catch(() => {});
    return Response.json({ error: `${err.name}: ${err.message}` }, { status: 500 });
  }
}
