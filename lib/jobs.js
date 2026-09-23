const { DynamoDBClient } = require("@aws-sdk/client-dynamodb");
const { DynamoDBDocumentClient, GetCommand, PutCommand, UpdateCommand } = require("@aws-sdk/lib-dynamodb");
const crypto = require("crypto");
const { clientConfig } = require("./aws-config");

const TABLE_NAME = process.env.JOBS_TABLE_NAME || "excel-repricer-jobs";
const TTL_SEGUNDOS = 60 * 60 * 24; // 1 dia

const client = new DynamoDBClient(clientConfig());
const docClient = DynamoDBDocumentClient.from(client, {
  marshallOptions: { removeUndefinedValues: true },
});

async function crearJob() {
  const jobId = crypto.randomUUID();
  const ahora = new Date().toISOString();
  await docClient.send(
    new PutCommand({
      TableName: TABLE_NAME,
      Item: {
        jobId,
        status: "pendiente",
        createdAt: ahora,
        updatedAt: ahora,
        ttl: Math.floor(Date.now() / 1000) + TTL_SEGUNDOS,
      },
    })
  );
  return jobId;
}

async function obtenerJob(jobId) {
  const resultado = await docClient.send(new GetCommand({ TableName: TABLE_NAME, Key: { jobId } }));
  return resultado.Item || null;
}

async function actualizarJob(jobId, campos) {
  const nombres = {};
  const valores = { ":updatedAt": new Date().toISOString() };
  const sets = ["updatedAt = :updatedAt"];

  Object.entries(campos).forEach(([clave, valor], i) => {
    const nombreAttr = `#c${i}`;
    const valorAttr = `:v${i}`;
    nombres[nombreAttr] = clave;
    valores[valorAttr] = valor;
    sets.push(`${nombreAttr} = ${valorAttr}`);
  });

  await docClient.send(
    new UpdateCommand({
      TableName: TABLE_NAME,
      Key: { jobId },
      UpdateExpression: `SET ${sets.join(", ")}`,
      ExpressionAttributeNames: nombres,
      ExpressionAttributeValues: valores,
    })
  );
}

module.exports = { crearJob, obtenerJob, actualizarJob };
