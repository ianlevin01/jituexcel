const { procesarOrdenarExcel } = require("../../lib/ordenar-proceso");

exports.handler = async function (event) {
  let body;
  try {
    body = JSON.parse(event.body || "{}");
  } catch {
    console.error("[ordenar-excel-background] body invalido:", event.body);
    return { statusCode: 400 };
  }

  const { jobId, key, tamanioEsperado } = body;
  if (!jobId || !key) {
    console.error("[ordenar-excel-background] falta jobId o key:", body);
    return { statusCode: 400 };
  }

  await procesarOrdenarExcel(jobId, key, tamanioEsperado);

  return { statusCode: 200 };
};
