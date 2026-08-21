const { getPresignedBaseUploadUrl } = require("../../../../lib/s3");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST() {
  try {
    const { url, key } = await getPresignedBaseUploadUrl();
    return Response.json({ url, key });
  } catch (err) {
    console.error("Error generando URL prefirmada:", err);
    return Response.json({ error: `${err.name}: ${err.message}` }, { status: 500 });
  }
}
