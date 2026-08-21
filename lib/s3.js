const { S3Client, GetObjectCommand, PutObjectCommand, HeadObjectCommand } = require("@aws-sdk/client-s3");
const { clientConfig } = require("./aws-config");

const BUCKET_NAME = process.env.EXCEL_BUCKET_NAME || "excel-repricer-051201154913-sa-east-1";
const BASE_EXCEL_KEY = "base/excel-base.xlsx";
const CONTENT_TYPE_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const s3 = new S3Client(clientConfig());

async function streamToBuffer(stream) {
  const chunks = [];
  for await (const chunk of stream) {
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

async function getBaseExcelBuffer() {
  const resultado = await s3.send(
    new GetObjectCommand({ Bucket: BUCKET_NAME, Key: BASE_EXCEL_KEY })
  );
  return streamToBuffer(resultado.Body);
}

async function putBaseExcelBuffer(buffer) {
  await s3.send(
    new PutObjectCommand({
      Bucket: BUCKET_NAME,
      Key: BASE_EXCEL_KEY,
      Body: buffer,
      ContentType: CONTENT_TYPE_XLSX,
    })
  );
}

async function getBaseExcelMetadata() {
  try {
    const resultado = await s3.send(
      new HeadObjectCommand({ Bucket: BUCKET_NAME, Key: BASE_EXCEL_KEY })
    );
    return {
      existe: true,
      ultimaActualizacion: resultado.LastModified,
      tamanioBytes: resultado.ContentLength,
    };
  } catch (err) {
    if (err.name === "NotFound" || err.$metadata?.httpStatusCode === 404) {
      return { existe: false };
    }
    throw err;
  }
}

module.exports = {
  getBaseExcelBuffer,
  putBaseExcelBuffer,
  getBaseExcelMetadata,
  BUCKET_NAME,
  BASE_EXCEL_KEY,
};
