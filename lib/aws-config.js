const REGION = process.env.APP_AWS_REGION || "sa-east-1";

const CREDENTIALS =
  process.env.APP_AWS_ACCESS_KEY_ID && process.env.APP_AWS_SECRET_ACCESS_KEY
    ? {
        accessKeyId: process.env.APP_AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.APP_AWS_SECRET_ACCESS_KEY,
      }
    : undefined;

function clientConfig() {
  return CREDENTIALS ? { region: REGION, credentials: CREDENTIALS } : { region: REGION };
}

module.exports = { REGION, clientConfig };
