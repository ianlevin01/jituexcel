const ExcelJS = require("exceljs");
const {
  getBaseExcelMetadata,
  getObjectBuffer,
  getObjectSize,
  promoteToBaseExcel,
  deleteObject,
} = require("../../../lib/s3");
const { workbookTieneColumnasRequeridas } = require("../../../lib/excel");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET() {
  const metadata = await getBaseExcelMetadata();
  return Response.json(metadata);
}

export async function POST(request) {
  try {
    const body = await request.json().catch(() => ({}));
    const { key, tamanioEsperado } = body;

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

    let buffer;
    try {
      buffer = await getObjectBuffer(key);
    } catch (err) {
      return Response.json(
        { error: `No se pudo leer el archivo subido a S3 (${err.name}: ${err.message}).` },
        { status: 400 }
      );
    }

    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.load(buffer);
    } catch (err) {
      await deleteObject(key).catch(() => {});
      return Response.json({ error: `El archivo no es un Excel válido (${err.message}).` }, { status: 400 });
    }

    if (!workbookTieneColumnasRequeridas(workbook)) {
      await deleteObject(key).catch(() => {});
      return Response.json(
        { error: 'El excel debe tener columnas "Item No." y "Precio".' },
        { status: 400 }
      );
    }

    await promoteToBaseExcel(key);
    const metadata = await getBaseExcelMetadata();
    return Response.json(metadata);
  } catch (err) {
    console.error("Error confirmando excel base:", err);
    return Response.json({ error: `${err.name}: ${err.message}` }, { status: 500 });
  }
}
