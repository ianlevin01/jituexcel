const { listClientes, createCliente } = require("../../../lib/dynamodb");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function esPorcentajeValido(porcentaje) {
  return typeof porcentaje === "number" && Number.isFinite(porcentaje);
}

export async function GET() {
  const clientes = await listClientes();
  return Response.json(clientes);
}

export async function POST(request) {
  const body = await request.json();
  const nombre = typeof body.nombre === "string" ? body.nombre.trim() : "";
  const porcentaje = Number(body.porcentaje);

  if (!nombre) {
    return Response.json({ error: "El nombre es obligatorio." }, { status: 400 });
  }
  if (!esPorcentajeValido(porcentaje)) {
    return Response.json({ error: "El porcentaje debe ser un número válido." }, { status: 400 });
  }

  const cliente = await createCliente({ nombre, porcentaje });
  return Response.json(cliente, { status: 201 });
}
