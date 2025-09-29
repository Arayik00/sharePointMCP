import axios from "axios";
import { ConfidentialClientApplication } from "@azure/msal-node";
import fs from "fs";
import forge from "node-forge";
import { logger } from "./utils/logger.js";
import { config } from "./utils/config.js";

export class SharePointGraphClient {
  constructor(accessToken, siteUrl) {
    this.accessToken = accessToken;
    this.siteUrl = siteUrl;
    this.siteId = null;
    this.drives = [];
    
    const urlParts = siteUrl.replace("https://", "").split("/");
    this.tenantDomain = urlParts[0];
    this.sitePath = urlParts.slice(1).join("/");
    
    this.axiosInstance = axios.create({
      baseURL: "https://graph.microsoft.com/v1.0/",
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
        "Content-Type": "application/json",
        Accept: "application/json",
      },
    });
  }

  async initialize() {
    this.siteId = await this.getSiteId();
    this.drives = await this.getDrives();
  }

  async getSiteId() {
    try {
      const response = await this.axiosInstance.get(
        `sites/${this.tenantDomain}:/${this.sitePath}`
      );
      logger.info(`Successfully retrieved site ID: ${response.data.id}`);
      return response.data.id;
    } catch (error) {
      logger.error(`Error getting site ID: ${error.message}`);
      throw error;
    }
  }

  async getDrives() {
    try {
      const response = await this.axiosInstance.get(`sites/${this.siteId}/drives`);
      logger.info(`Successfully retrieved ${response.data.value.length} drives`);
      return response.data.value;
    } catch (error) {
      logger.error(`Error getting drives: ${error.message}`);
      throw error;
    }
  }

  async getFolders(driveId, folderPath = "") {
    try {
      const endpoint = folderPath
        ? `drives/${driveId}/root:/${folderPath}:/children`
        : `drives/${driveId}/root/children`;
      
      const response = await this.axiosInstance.get(endpoint);
      const folders = response.data.value.filter(item => item.folder);
      logger.info(`Found ${folders.length} folders in ${folderPath || 'root'}`);
      return folders;
    } catch (error) {
      logger.error(`Error getting folders: ${error.message}`);
      throw error;
    }
  }

  async getFiles(driveId, folderPath = "") {
    try {
      const endpoint = folderPath
        ? `drives/${driveId}/root:/${folderPath}:/children`
        : `drives/${driveId}/root/children`;
      
      const response = await this.axiosInstance.get(endpoint);
      const files = response.data.value.filter(item => item.file);
      logger.info(`Found ${files.length} files in ${folderPath || 'root'}`);
      return files;
    } catch (error) {
      logger.error(`Error getting files: ${error.message}`);
      throw error;
    }
  }

  async getFileContent(driveId, filePath) {
    try {
      const response = await this.axiosInstance.get(
        `drives/${driveId}/root:/${filePath}:/content`,
        { responseType: 'text' }
      );
      logger.info(`Successfully retrieved content for file: ${filePath}`);
      return response.data;
    } catch (error) {
      logger.error(`Error getting file content: ${error.message}`);
      throw error;
    }
  }

  async uploadFile(driveId, filePath, content) {
    try {
      const response = await this.axiosInstance.put(
        `drives/${driveId}/root:/${filePath}:/content`,
        content,
        {
          headers: {
            "Content-Type": "application/octet-stream",
          },
        }
      );
      logger.info(`Successfully uploaded file: ${filePath}`);
      return response.data;
    } catch (error) {
      logger.error(`Error uploading file: ${error.message}`);
      throw error;
    }
  }

  async createFolder(driveId, parentPath, folderName) {
    try {
      const parentEndpoint = parentPath
        ? `drives/${driveId}/root:/${parentPath}:`
        : `drives/${driveId}/root`;
      
      const response = await this.axiosInstance.post(
        `${parentEndpoint}/children`,
        {
          name: folderName,
          folder: {},
          "@microsoft.graph.conflictBehavior": "rename",
        }
      );
      logger.info(`Successfully created folder: ${folderName}`);
      return response.data;
    } catch (error) {
      logger.error(`Error creating folder: ${error.message}`);
      throw error;
    }
  }

  async deleteItem(driveId, itemPath) {
    try {
      await this.axiosInstance.delete(`drives/${driveId}/root:/${itemPath}`);
      logger.info(`Successfully deleted item: ${itemPath}`);
      return { success: true };
    } catch (error) {
      logger.error(`Error deleting item: ${error.message}`);
      throw error;
    }
  }
}

// Certificate authentication functions
export function getCertificatePrivateKey(certPath, certPassword = "") {
  try {
    const certData = fs.readFileSync(certPath);
    const p12Asn1 = forge.asn1.fromDer(certData.toString('binary'));
    const p12 = forge.pkcs12.pkcs12FromAsn1(p12Asn1, certPassword);
    
    const keyBags = p12.getBags({ bagType: forge.pki.oids.pkcs8ShroudedKeyBag });
    const keyBag = keyBags[forge.pki.oids.pkcs8ShroudedKeyBag][0];
    
    return forge.pki.privateKeyToPem(keyBag.key);
  } catch (error) {
    logger.error(`Error reading certificate: ${error.message}`);
    throw error;
  }
}

export function getCertificateThumbprint(certPath, certPassword = "") {
  try {
    const certData = fs.readFileSync(certPath);
    const p12Asn1 = forge.asn1.fromDer(certData.toString('binary'));
    const p12 = forge.pkcs12.pkcs12FromAsn1(p12Asn1, certPassword);
    
    const certBags = p12.getBags({ bagType: forge.pki.oids.certBag });
    const certBag = certBags[forge.pki.oids.certBag][0];
    
    const cert = forge.pki.certificateFromPem(forge.pki.certificateToPem(certBag.cert));
    const der = forge.asn1.toDer(forge.pki.certificateToAsn1(cert)).getBytes();
    const hash = forge.md.sha1.create();
    hash.update(der);
    
    return hash.digest().toHex().toUpperCase();
  } catch (error) {
    logger.error(`Error getting certificate thumbprint: ${error.message}`);
    throw error;
  }
}

export async function getAccessToken() {
  try {
    const privateKey = getCertificatePrivateKey(config.CERT_PATH, config.CERT_PASSWORD);
    const thumbprint = getCertificateThumbprint(config.CERT_PATH, config.CERT_PASSWORD);

    const clientConfig = {
      auth: {
        clientId: config.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${config.TENANT_ID}`,
        clientCertificate: {
          thumbprint: thumbprint,
          privateKey: privateKey,
        },
      },
    };

    const cca = new ConfidentialClientApplication(clientConfig);

    const clientCredentialRequest = {
      scopes: ["https://graph.microsoft.com/.default"],
    };

    const response = await cca.acquireTokenByClientCredential(clientCredentialRequest);
    logger.info("Successfully acquired access token");
    return response.accessToken;
  } catch (error) {
    logger.error(`Error getting access token: ${error.message}`);
    throw error;
  }
}

export async function createSharePointClient() {
  try {
    const accessToken = await getAccessToken();
    const client = new SharePointGraphClient(accessToken, config.SITE_URL);
    await client.initialize();
    return client;
  } catch (error) {
    logger.error(`Error creating SharePoint client: ${error.message}`);
    throw error;
  }
}