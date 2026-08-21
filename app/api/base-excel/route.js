const ExcelJS = require("exceljs");
const { getBaseExcelMetadata, putBaseExcelBuffer } = require("../../../lib/s3");
const { workbookTieneColumnasRequeridas } = require("../../../lib/excel");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET() {
  const metadata = await getBaseExcelMetadata();
  return Response.json(metadata);
}

export async function POST(request) {
  const formData = await request.formData();
  const file = formData.get("file");

  if (!file) {
    return Response.json({ error: "Falta el archivo." }, { status: 400 });
  }

  const buffer = Buffer.from(await file.arrayBuffer());

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(buffer);
  } catch {
    return Response.json({ error: "El archivo no es un Excel válido." }, { status: 400 });
  }

  if (!workbookTieneColumnasRequeridas(workbook)) {
    return Response.json(
      { error: 'El excel debe tener columnas "Item No." y "Precio".' },
      { status: 400 }
    );
  }

  await putBaseExcelBuffer(buffer);
  const metadata = await getBaseExcelMetadata();
  return Response.json(metadata);
}
