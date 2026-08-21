const { DynamoDBClient } = require("@aws-sdk/client-dynamodb");
const {
  DynamoDBDocumentClient,
  ScanCommand,
  GetCommand,
  PutCommand,
  UpdateCommand,
  DeleteCommand,
} = require("@aws-sdk/lib-dynamodb");
const crypto = require("crypto");
const { clientConfig } = require("./aws-config");

const TABLE_NAME = process.env.CLIENTES_TABLE_NAME || "excel-repricer-clientes";

const client = new DynamoDBClient(clientConfig());
const docClient = DynamoDBDocumentClient.from(client, {
  marshallOptions: { removeUndefinedValues: true },
});

async function listClientes() {
  const resultado = await docClient.send(new ScanCommand({ TableName: TABLE_NAME }));
  const items = resultado.Items || [];
  items.sort((a, b) => a.nombre.localeCompare(b.nombre));
  return items;
}

async function getCliente(clienteId) {
  const resultado = await docClient.send(
    new GetCommand({ TableName: TABLE_NAME, Key: { clienteId } })
  );
  return resultado.Item || null;
}

async function createCliente({ nombre, porcentaje }) {
  const ahora = new Date().toISOString();
  const item = {
    clienteId: crypto.randomUUID(),
    nombre,
    porcentaje,
    createdAt: ahora,
    updatedAt: ahora,
  };
  await docClient.send(new PutCommand({ TableName: TABLE_NAME, Item: item }));
  return item;
}

async function updateCliente(clienteId, { nombre, porcentaje }) {
  const resultado = await docClient.send(
    new UpdateCommand({
      TableName: TABLE_NAME,
      Key: { clienteId },
      ConditionExpression: "attribute_exists(clienteId)",
      UpdateExpression: "SET nombre = :nombre, porcentaje = :porcentaje, updatedAt = :updatedAt",
      ExpressionAttributeValues: {
        ":nombre": nombre,
        ":porcentaje": porcentaje,
        ":updatedAt": new Date().toISOString(),
      },
      ReturnValues: "ALL_NEW",
    })
  );
  return resultado.Attributes;
}

async function deleteCliente(clienteId) {
  await docClient.send(new DeleteCommand({ TableName: TABLE_NAME, Key: { clienteId } }));
}

module.exports = { listClientes, getCliente, createCliente, updateCliente, deleteCliente };
