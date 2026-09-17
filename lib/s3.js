const {
  S3Client,
  GetObjectCommand,
  PutObjectCommand,
  HeadObjectCommand,
  CopyObjectCommand,
  DeleteObjectCommand,
} = require("@aws-sdk/client-s3");
const { getSignedUrl } = require("@aws-sdk/s3-request-presigner");
const crypto = require("crypto");
const { clientConfig } = require("./aws-config");

const BUCKET_NAME = process.env.EXCEL_BUCKET_NAME || "excel-repricer-051201154913-sa-east-1";
const BASE_EXCEL_KEY = "base/excel-base.xlsx";
const CONTENT_TYPE_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const s3 = new S3Client({
  ...clientConfig(),
  requestChecksumCalculation: "WHEN_REQUIRED",
});

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

async function getObjectBuffer(key) {
  const resultado = await s3.send(new GetObjectCommand({ Bucket: BUCKET_NAME, Key: key }));
  return streamToBuffer(resultado.Body);
}

async function getObjectSize(key) {
  const resultado = await s3.send(new HeadObjectCommand({ Bucket: BUCKET_NAME, Key: key }));
  return resultado.ContentLength;
}

async function putObjectBuffer(key, buffer) {
  await s3.send(
    new PutObjectCommand({ Bucket: BUCKET_NAME, Key: key, Body: buffer, ContentType: CONTENT_TYPE_XLSX })
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

async function getPresignedUploadUrl(keyPrefix) {
  const key = `${keyPrefix}${crypto.randomUUID()}.xlsx`;
  const command = new PutObjectCommand({
    Bucket: BUCKET_NAME,
    Key: key,
    ContentType: CONTENT_TYPE_XLSX,
  });
  const url = await getSignedUrl(s3, command, { expiresIn: 900 });
  return { url, key };
}

async function getPresignedDownloadUrl(key, filename) {
  const command = new GetObjectCommand({
    Bucket: BUCKET_NAME,
    Key: key,
    ResponseContentDisposition: `attachment; filename="${filename}"`,
    ResponseContentType: CONTENT_TYPE_XLSX,
  });
  return getSignedUrl(s3, command, { expiresIn: 600 });
}

async function promoteToBaseExcel(tempKey) {
  await s3.send(
    new CopyObjectCommand({
      Bucket: BUCKET_NAME,
      CopySource: `/${BUCKET_NAME}/${tempKey}`,
      Key: BASE_EXCEL_KEY,
      ContentType: CONTENT_TYPE_XLSX,
    })
  );
  await deleteObject(tempKey);
}

async function deleteObject(key) {
  await s3.send(new DeleteObjectCommand({ Bucket: BUCKET_NAME, Key: key }));
}

module.exports = {
  getBaseExcelBuffer,
  getObjectBuffer,
  getObjectSize,
  putObjectBuffer,
  getBaseExcelMetadata,
  getPresignedUploadUrl,
  getPresignedDownloadUrl,
  promoteToBaseExcel,
  deleteObject,
  BUCKET_NAME,
  BASE_EXCEL_KEY,
};
