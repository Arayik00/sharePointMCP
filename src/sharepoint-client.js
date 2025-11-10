import axios from "axios";
import { ConfidentialClientApplication } from "@azure/msal-node";
import fs from "fs";
import forge from "node-forge";
import { loadConfig } from "./config.js";

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
      console.log(`✅ Successfully retrieved site ID: ${response.data.id}`);
      return response.data.id;
    } catch (error) {
      console.error(`❌ Error getting site ID: ${error.message}`);
      throw error;
    }
  }

  async getDrives() {
    try {
      const response = await this.axiosInstance.get(`sites/${this.siteId}/drives`);
      console.log(`✅ Successfully retrieved ${response.data.value.length} drives`);
      return response.data.value;
    } catch (error) {
      console.error(`❌ Error getting drives: ${error.message}`);
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
      console.log(`✅ Found ${folders.length} folders in ${folderPath || 'root'}`);
      return folders;
    } catch (error) {
      console.error(`❌ Error getting folders: ${error.message}`);
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
      console.log(`✅ Found ${files.length} files in ${folderPath || 'root'}`);
      return files;
    } catch (error) {
      console.error(`❌ Error getting files: ${error.message}`);
      throw error;
    }
  }

  async getFileContent(driveId, filePath) {
    try {
      const response = await this.axiosInstance.get(
        `drives/${driveId}/root:/${filePath}:/content`,
        { responseType: 'text' }
      );
      console.log(`✅ Successfully retrieved content for file: ${filePath}`);
      return response.data;
    } catch (error) {
      console.error(`❌ Error getting file content: ${error.message}`);
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
      console.log(`✅ Successfully uploaded file: ${filePath}`);
      return response.data;
    } catch (error) {
      console.error(`❌ Error uploading file: ${error.message}`);
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
      console.log(`✅ Successfully created folder: ${folderName}`);
      return response.data;
    } catch (error) {
      console.error(`❌ Error creating folder: ${error.message}`);
      throw error;
    }
  }

  async deleteItem(driveId, itemPath) {
    try {
      await this.axiosInstance.delete(`drives/${driveId}/root:/${itemPath}`);
      console.log(`✅ Successfully deleted item: ${itemPath}`);
      return { success: true };
    } catch (error) {
      console.error(`❌ Error deleting item: ${error.message}`);
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
    console.error(`❌ Error reading certificate: ${error.message}`);
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
    console.error(`❌ Error getting certificate thumbprint: ${error.message}`);
    throw error;
  }
}

export async function getAccessToken() {
  try {
    const config = await loadConfig();
    const { sharepoint } = config;
    
    const privateKey = getCertificatePrivateKey(sharepoint.certPath, sharepoint.certPassword);
    const thumbprint = getCertificateThumbprint(sharepoint.certPath, sharepoint.certPassword);

    const clientConfig = {
      auth: {
        clientId: sharepoint.appId,
        authority: `https://login.microsoftonline.com/${sharepoint.tenantId}`,
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
    console.log("✅ Successfully acquired access token");
    return response.accessToken;
  } catch (error) {
    console.error(`❌ Error getting access token: ${error.message}`);
    throw error;
  }
}

export async function createSharePointClient() {
  try {
    const config = await loadConfig();
    const accessToken = await getAccessToken();
    const client = new SharePointGraphClient(accessToken, config.sharepoint.siteUrl);
    await client.initialize();
    return client;
  } catch (error) {
    console.error(`❌ Error creating SharePoint client: ${error.message}`);
    throw error;
  }
}