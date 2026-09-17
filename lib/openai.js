const MODEL = "gpt-4o-mini";
const TAMANIO_LOTE = 12;
const MAX_REINTENTOS = 3;
const TIMEOUT_MS = 60000;

function apiKey() {
  const key = process.env.OPENAI_API_KEY;
  if (!key) throw new Error("Falta configurar OPENAI_API_KEY.");
  return key;
}

function esperar(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function llamarOpenAI(contenido) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), TIMEOUT_MS);

  try {
    const respuesta = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      signal: controller.signal,
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
      const error = new Error(`OpenAI respondió ${respuesta.status}: ${texto.slice(0, 300)}`);
      error.status = respuesta.status;
      throw error;
    }

    return await respuesta.json();
  } finally {
    clearTimeout(timeoutId);
  }
}

function esReintentable(err) {
  if (err.name === "AbortError") return true;
  if (typeof err.status === "number" && (err.status === 429 || err.status >= 500)) return true;
  if (err.status === undefined) return true; // error de red/fetch, no de la API
  return false;
}

async function clasificarLote(imagenes, categoriasConocidas, numeroLote) {
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

  let ultimoError;
  for (let intento = 1; intento <= MAX_REINTENTOS; intento++) {
    try {
      console.log(`[ordenar-excel] lote ${numeroLote}: intento ${intento}/${MAX_REINTENTOS}, ${imagenes.length} imagenes`);
      const data = await llamarOpenAI(contenido);

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
        throw new Error(
          `OpenAI devolvió ${Array.isArray(categorias) ? categorias.length : "0"} categorías, se esperaban ${imagenes.length}.`
        );
      }

      console.log(`[ordenar-excel] lote ${numeroLote}: ok`);
      return categorias.map((c) => String(c).trim().toLowerCase());
    } catch (err) {
      ultimoError = err;
      console.warn(`[ordenar-excel] lote ${numeroLote}: fallo intento ${intento} -> ${err.message}`);
      if (intento < MAX_REINTENTOS && esReintentable(err)) {
        await esperar(1000 * intento);
        continue;
      }
      break;
    }
  }

  throw new Error(`Falló la clasificación con OpenAI (lote ${numeroLote}): ${ultimoError.message}`);
}

async function clasificarImagenes(imagenes, tamanioLote = TAMANIO_LOTE) {
  const resultado = new Array(imagenes.length);
  const categoriasConocidas = [];
  const totalLotes = Math.ceil(imagenes.length / tamanioLote);

  console.log(`[ordenar-excel] clasificando ${imagenes.length} imagenes en ${totalLotes} lote(s)`);

  for (let i = 0; i < imagenes.length; i += tamanioLote) {
    const numeroLote = Math.floor(i / tamanioLote) + 1;
    const lote = imagenes.slice(i, i + tamanioLote);
    const categorias = await clasificarLote(lote, categoriasConocidas, numeroLote);
    categorias.forEach((cat, idx) => {
      resultado[i + idx] = cat;
      if (!categoriasConocidas.includes(cat)) categoriasConocidas.push(cat);
    });
  }

  console.log(`[ordenar-excel] clasificacion terminada, ${categoriasConocidas.length} categoria(s) distintas`);
  return resultado;
}

module.exports = { clasificarImagenes };
