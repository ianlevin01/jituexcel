const { obtenerJob } = require("../../../../lib/jobs");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request) {
  try {
    const { searchParams } = new URL(request.url);
    const jobId = searchParams.get("jobId");
    if (!jobId) {
      return Response.json({ error: "Falta jobId." }, { status: 400 });
    }

    const job = await obtenerJob(jobId);
    if (!job) {
      return Response.json({ error: "Job no encontrado." }, { status: 404 });
    }

    return Response.json(job);
  } catch (err) {
    console.error("Error en /api/ordenar-excel/status:", err);
    return Response.json({ error: `${err.name}: ${err.message}` }, { status: 500 });
  }
}
