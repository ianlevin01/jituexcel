const { getCliente } = require("../../../lib/dynamodb");
const { getBaseExcelBuffer } = require("../../../lib/s3");
const { repreciarDesdeBuffers } = require("../../../lib/excel");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request) {
  try {
    let formData;
    try {
      formData = await request.formData();
    } catch (err) {
      return Response.json(
        { error: `No se pudo leer el archivo subido (${err.message}). Puede ser demasiado grande.` },
        { status: 400 }
      );
    }

    const clienteId = formData.get("clienteId");
    const file = formData.get("file");

    if (!clienteId || !file) {
      return Response.json({ error: "Falta el cliente o el archivo." }, { status: 400 });
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

    const bufferAModificar = Buffer.from(await file.arrayBuffer());

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

    return new Response(buffer, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="actualizados.xlsx"',
        "X-Filas-Omitidas": String(filasOmitidas),
      },
    });
  } catch (err) {
    console.error("Error en /api/reprice:", err);
    return Response.json({ error: `${err.name}: ${err.message}` }, { status: 500 });
  }
}
