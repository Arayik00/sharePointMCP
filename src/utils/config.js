import dotenv from "dotenv";

// Load environment variables
dotenv.config();

function getEnvVar(name, required = true, defaultValue) {
  const value = process.env[name] || defaultValue;

  if (required && !value) {
    throw new Error(`Environment variable ${name} is required but not set`);
  }

  return value || "";
}

export const config = {
  CLIENT_ID: getEnvVar("SHP_ID_APP"),
  APP_SECRET: getEnvVar("SHP_ID_APP_SECRET", false),
  SITE_URL: getEnvVar("SHP_SITE_URL"),
  DOC_LIBRARY: getEnvVar("SHP_DOC_LIBRARY", false, ""),
  TENANT_ID: getEnvVar("SHP_TENANT_ID"),
  CERT_PATH: getEnvVar("SHP_CERT_PFX_PATH", false),
  CERT_PASSWORD: getEnvVar("SHP_CERT_PFX_PASSWORD", false),
  MAX_DEPTH: parseInt(getEnvVar("SHP_MAX_DEPTH", false, "15")),
  MAX_FOLDERS_PER_LEVEL: parseInt(
    getEnvVar("SHP_MAX_FOLDERS_PER_LEVEL", false, "100")
  ),
  LEVEL_DELAY: parseFloat(getEnvVar("SHP_LEVEL_DELAY", false, "0.5")),
};
