const { getPresignedUploadUrl } = require("../../../../lib/s3");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request) {
  try {
    const body = await request.json().catch(() => ({}));
    const prefijo = body.archivo === "b" ? "excel-ab/b-pending-" : "excel-ab/a-pending-";
    const { url, key } = await getPresignedUploadUrl(prefijo);
    return Response.json({ url, key });
  } catch (err) {
    console.error("Error generando URL prefirmada:", err);
    return Response.json({ error: `${err.name}: ${err.message}` }, { status: 500 });
  }
}
