const MODEL = "gpt-4o-mini";
const TAMANIO_LOTE = 12;

function apiKey() {
  const key = process.env.OPENAI_API_KEY;
  if (!key) throw new Error("Falta configurar OPENAI_API_KEY.");
  return key;
}

async function clasificarLote(imagenes, categoriasConocidas) {
  const listaConocidas = categoriasConocidas.length
    ? `Categorías ya usadas (reusalas exactamente igual si el producto corresponde a alguna): ${categoriasConocidas.join(", ")}.`
    : "Todavía no hay categorías definidas, elegí vos los nombres.";

  const contenido = [
    {
      type: "text",
      text:
        `Vas a ver ${imagenes.length} fotos de productos, numeradas del 1 al ${imagenes.length} en ese orden. ` +
        `Para cada una, indicá una categoría corta en español (una o dos palabras, minúsculas, ej: "auto", "moto", "bicicleta") ` +
        `que describa el TIPO de producto. Productos del mismo tipo deben tener EXACTAMENTE el mismo texto de categoría. ` +
        `${listaConocidas}\n\n` +
        `Respondé solo JSON con este formato: {"categorias": ["categoria1", "categoria2", ...]} ` +
        `con exactamente ${imagenes.length} elementos, en el mismo orden que las fotos.`,
    },
    ...imagenes.map((img) => ({
      type: "image_url",
      image_url: {
        url: `data:image/${img.extension};base64,${img.buffer.toString("base64")}`,
        detail: "low",
      },
    })),
  ];

  const respuesta = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey()}`,
    },
    body: JSON.stringify({
      model: MODEL,
      response_format: { type: "json_object" },
      temperature: 0,
      messages: [{ role: "user", content: contenido }],
    }),
  });

  if (!respuesta.ok) {
    const texto = await respuesta.text().catch(() => "");
    throw new Error(`OpenAI respondió ${respuesta.status}: ${texto.slice(0, 300)}`);
  }

  const data = await respuesta.json();
  const contenidoRespuesta = data.choices?.[0]?.message?.content;
  if (!contenidoRespuesta) throw new Error("Respuesta de OpenAI sin contenido.");

  let parseado;
  try {
    parseado = JSON.parse(contenidoRespuesta);
  } catch {
    throw new Error("La respuesta de OpenAI no fue JSON válido.");
  }

  const categorias = parseado.categorias;
  if (!Array.isArray(categorias) || categorias.length !== imagenes.length) {
    throw new Error("OpenAI no devolvió la cantidad esperada de categorías.");
  }

  return categorias.map((c) => String(c).trim().toLowerCase());
}

async function clasificarImagenes(imagenes, tamanioLote = TAMANIO_LOTE) {
  const resultado = new Array(imagenes.length);
  const categoriasConocidas = [];

  for (let i = 0; i < imagenes.length; i += tamanioLote) {
    const lote = imagenes.slice(i, i + tamanioLote);
    const categorias = await clasificarLote(lote, categoriasConocidas);
    categorias.forEach((cat, idx) => {
      resultado[i + idx] = cat;
      if (!categoriasConocidas.includes(cat)) categoriasConocidas.push(cat);
    });
  }

  return resultado;
}

module.exports = { clasificarImagenes };
