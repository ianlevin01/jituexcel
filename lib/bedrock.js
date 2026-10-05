const { BedrockRuntimeClient, InvokeModelCommand } = require("@aws-sdk/client-bedrock-runtime");
const { clientConfig } = require("./aws-config");

const MODEL_ID = "amazon.titan-embed-image-v1";
const REGION_BEDROCK = process.env.APP_AWS_BEDROCK_REGION || process.env.APP_AWS_REGION || "us-east-1";

const client = new BedrockRuntimeClient({ ...clientConfig(), region: REGION_BEDROCK });

async function obtenerEmbedding(buffer, extension) {
  const body = {
    inputImage: buffer.toString("base64"),
    embeddingConfig: { outputEmbeddingLength: 256 },
  };

  const respuesta = await client.send(
    new InvokeModelCommand({
      modelId: MODEL_ID,
      contentType: "application/json",
      accept: "application/json",
      body: JSON.stringify(body),
    })
  );

  const resultado = JSON.parse(Buffer.from(respuesta.body).toString("utf8"));
  if (!Array.isArray(resultado.embedding)) {
    throw new Error("Bedrock no devolvió un embedding válido.");
  }
  return resultado.embedding;
}

async function obtenerEmbeddings(imagenes) {
  const embeddings = [];
  for (const img of imagenes) {
    const embedding = await obtenerEmbedding(img.buffer, img.extension);
    embeddings.push(embedding);
  }
  return embeddings;
}

module.exports = { obtenerEmbedding, obtenerEmbeddings };
