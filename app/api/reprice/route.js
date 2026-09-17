const crypto = require("crypto");
const { getCliente } = require("../../../lib/dynamodb");
const {
  getBaseExcelBuffer,
  getObjectBuffer,
  getObjectSize,
  putObjectBuffer,
  getPresignedDownloadUrl,
  deleteObject,
} = require("../../../lib/s3");
const { repreciarDesdeBuffers } = require("../../../lib/excel");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request) {
  let uploadKey;
  try {
    const body = await request.json().catch(() => ({}));
    const { clienteId, key, tamanioEsperado } = body;
    uploadKey = key;

    if (!clienteId || !key) {
      return Response.json({ error: "Falta el cliente o el archivo." }, { status: 400 });
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

    const cliente = await getCliente(clienteId);
    if (!cliente) {
      return Response.json({ error: "Cliente no encontrado." }, { status: 404 });
    }

    let bufferBase;
    try {
      bufferBase = await getBaseExcelBuffer();
    } catch (err) {
      if (err.name === "NoSuchKey") {
        return Response.json({ error: "Todavía no se subió el excel base." }, { status: 400 });
      }
      throw err;
    }

    const bufferAModificar = await getObjectBuffer(key);

    const { buffer, hojasActualizadas, filasOmitidas } = await repreciarDesdeBuffers(
      bufferBase,
      bufferAModificar,
      cliente.porcentaje
    );

    if (hojasActualizadas === 0) {
      return Response.json(
        { error: 'No se encontraron columnas "Item No." y "Precio" en común entre los dos archivos.' },
        { status: 400 }
      );
    }

    const outputKey = `outputs/actualizados-${crypto.randomUUID()}.xlsx`;
    await putObjectBuffer(outputKey, buffer);
    const downloadUrl = await getPresignedDownloadUrl(outputKey, "actualizados.xlsx");

    deleteObject(key).catch(() => {});

    return Response.json({ downloadUrl, filasOmitidas });
  } catch (err) {
    console.error("Error en /api/reprice:", err);
    if (uploadKey) deleteObject(uploadKey).catch(() => {});
    return Response.json({ error: `${err.name}: ${err.message}` }, { status: 500 });
  }
}
