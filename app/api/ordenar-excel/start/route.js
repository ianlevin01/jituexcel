const { crearJob } = require("../../../../lib/jobs");
const { procesarOrdenarExcel } = require("../../../../lib/ordenar-proceso");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function corriendoEnLambda() {
  return Boolean(process.env.AWS_LAMBDA_FUNCTION_NAME);
}

export async function POST(request) {
  try {
    const body = await request.json().catch(() => ({}));
    const { key, tamanioEsperado } = body;

    if (!key) {
      return Response.json({ error: "Falta la key del archivo subido." }, { status: 400 });
    }

    const jobId = await crearJob();

    if (corriendoEnLambda()) {
      // Desplegado (Netlify/Lambda): invocamos la Background Function, que no tiene
      // el limite de ~10s de una funcion sincronica, y corre hasta 15 minutos.
      const siteUrl = process.env.URL || process.env.DEPLOY_URL || process.env.DEPLOY_PRIME_URL;
      if (!siteUrl) {
        console.error("[ordenar-excel/start] no se encontro la URL del sitio para invocar la background function");
        return Response.json(
          { error: "No se pudo determinar la URL del sitio para procesar en segundo plano." },
          { status: 500 }
        );
      }
      fetch(`${siteUrl}/.netlify/functions/ordenar-excel-background`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ jobId, key, tamanioEsperado }),
      }).catch((err) => console.error("[ordenar-excel/start] error invocando background function:", err));
    } else {
      // Desarrollo local (next dev): no hay background functions, procesamos
      // directo sin esperar la respuesta (el proceso de node sigue vivo igual).
      procesarOrdenarExcel(jobId, key, tamanioEsperado).catch((err) =>
        console.error("[ordenar-excel/start] error procesando en local:", err)
      );
    }

    return Response.json({ jobId });
  } catch (err) {
    console.error("Error en /api/ordenar-excel/start:", err);
    return Response.json({ error: `${err.name}: ${err.message}` }, { status: 500 });
  }
}
