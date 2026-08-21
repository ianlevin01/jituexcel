const { updateCliente, deleteCliente } = require("../../../../lib/dynamodb");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function esPorcentajeValido(porcentaje) {
  return typeof porcentaje === "number" && Number.isFinite(porcentaje);
}

export async function PUT(request, { params }) {
  const { id } = await params;
  const body = await request.json();
  const nombre = typeof body.nombre === "string" ? body.nombre.trim() : "";
  const porcentaje = Number(body.porcentaje);

  if (!nombre) {
    return Response.json({ error: "El nombre es obligatorio." }, { status: 400 });
  }
  if (!esPorcentajeValido(porcentaje)) {
    return Response.json({ error: "El porcentaje debe ser un número válido." }, { status: 400 });
  }

  try {
    const cliente = await updateCliente(id, { nombre, porcentaje });
    return Response.json(cliente);
  } catch (err) {
    if (err.name === "ConditionalCheckFailedException") {
      return Response.json({ error: "Cliente no encontrado." }, { status: 404 });
    }
    throw err;
  }
}

export async function DELETE(request, { params }) {
  const { id } = await params;
  await deleteCliente(id);
  return new Response(null, { status: 204 });
}
