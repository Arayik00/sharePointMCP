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
    console.log(`📋 Reading certificate from: ${certPath}`);
    
    // Check if file exists
    if (!fs.existsSync(certPath)) {
      throw new Error(`Certificate file not found: ${certPath}`);
    }
    
    const certData = fs.readFileSync(certPath);
    console.log(`📋 Certificate file size: ${certData.length} bytes`);
    console.log(`📋 First 32 bytes (hex): ${certData.slice(0, 32).toString('hex')}`);
    
    // Check if this looks like a PKCS#12 file
    const magicBytes = certData.slice(0, 4);
    const hexString = certData.slice(0, 32).toString('hex');
    console.log(`📋 Magic bytes: ${magicBytes.toString('hex')}`);
    
    // Detect UTF-8 corruption (efbfbd = UTF-8 replacement character)
    if (hexString.includes('efbfbd')) {
      console.error('🚨 CERTIFICATE FILE CORRUPTION DETECTED!');
      console.error('📋 The certificate file contains UTF-8 replacement characters (�)');
      console.error('📋 This means the binary PKCS#12 file was corrupted during upload/transfer');
      console.error('');
      console.error('🔧 SOLUTIONS:');
      console.error('1. Re-upload the certificate file in binary mode');
      console.error('2. Check that the file wasn\'t opened/saved in a text editor');
      console.error('3. Verify the original certificate.pfx file is not corrupted');
      console.error('4. Use base64 encoding for safer transfer:');
      console.error('   - Convert: base64 certificate.pfx > certificate.b64');
      console.error('   - Upload the .b64 file and decode it in the environment');
      throw new Error('Certificate file is corrupted - contains UTF-8 replacement characters');
    }
    
    if (!hexString.startsWith('3082') && !hexString.startsWith('3081')) {
      console.warn('⚠️  Certificate doesn\'t start with expected PKCS#12 magic bytes');
      console.warn('📋 Expected: 3082xxxx or 3081xxxx (ASN.1 SEQUENCE)');
      console.warn(`📋 Found: ${magicBytes.toString('hex')}`);
    }
    
    // Try different encoding approaches
    let p12Asn1;
    let conversionMethod = '';
    
    try {
      // Method 1: Direct binary conversion
      conversionMethod = 'binary string';
      console.log(`🔧 Trying ${conversionMethod} conversion...`);
      p12Asn1 = forge.asn1.fromDer(certData.toString('binary'));
      console.log(`✅ Success with ${conversionMethod}`);
    } catch (binaryError) {
      console.log(`❌ ${conversionMethod} failed: ${binaryError.message}`);
      
      try {
        // Method 2: Using forge binary encoding
        conversionMethod = 'forge binary encoding';
        console.log(`🔧 Trying ${conversionMethod}...`);
        p12Asn1 = forge.asn1.fromDer(forge.util.binary.raw.encode(certData));
        console.log(`✅ Success with ${conversionMethod}`);
      } catch (forgeError) {
        console.log(`❌ ${conversionMethod} failed: ${forgeError.message}`);
        
        try {
          // Method 3: Manual byte-by-byte conversion
          conversionMethod = 'manual byte conversion';
          console.log(`🔧 Trying ${conversionMethod}...`);
          let binaryString = '';
          for (let i = 0; i < certData.length; i++) {
            binaryString += String.fromCharCode(certData[i]);
          }
          p12Asn1 = forge.asn1.fromDer(binaryString);
          console.log(`✅ Success with ${conversionMethod}`);
        } catch (manualError) {
          console.log(`❌ ${conversionMethod} failed: ${manualError.message}`);
          
          // Final attempt: Check if file might be base64 encoded
          try {
            conversionMethod = 'base64 decoding';
            console.log(`🔧 Trying ${conversionMethod} (in case file is base64)...`);
            const base64Data = certData.toString('utf8').replace(/\s/g, '');
            const decodedData = Buffer.from(base64Data, 'base64');
            p12Asn1 = forge.asn1.fromDer(decodedData.toString('binary'));
            console.log(`✅ Success with ${conversionMethod}`);
          } catch (base64Error) {
            console.log(`❌ ${conversionMethod} failed: ${base64Error.message}`);
            
            throw new Error(`All certificate parsing methods failed. Original errors:
  Binary: ${binaryError.message}
  Forge: ${forgeError.message}  
  Manual: ${manualError.message}
  Base64: ${base64Error.message}`);
          }
        }
      }
    }
    
    const p12 = forge.pkcs12.pkcs12FromAsn1(p12Asn1, certPassword);
    
    const keyBags = p12.getBags({ bagType: forge.pki.oids.pkcs8ShroudedKeyBag });
    const keyBag = keyBags[forge.pki.oids.pkcs8ShroudedKeyBag][0];
    
    if (!keyBag || !keyBag.key) {
      throw new Error('No private key found in certificate');
    }
    
    console.log(`✅ Successfully extracted private key from certificate`);
    return forge.pki.privateKeyToPem(keyBag.key);
  } catch (error) {
    console.error(`❌ Error reading certificate: ${error.message}`);
    throw error;
  }
}

export function getCertificateThumbprint(certPath, certPassword = "") {
  try {
    console.log(`📋 Reading certificate thumbprint from: ${certPath}`);
    
    if (!fs.existsSync(certPath)) {
      throw new Error(`Certificate file not found: ${certPath}`);
    }
    
    const certData = fs.readFileSync(certPath);
    
    // Use the same robust parsing logic as getCertificatePrivateKey
    let p12Asn1;
    let conversionMethod = '';
    
    try {
      // Method 1: Direct binary conversion
      conversionMethod = 'binary string';
      p12Asn1 = forge.asn1.fromDer(certData.toString('binary'));
    } catch (binaryError) {
      try {
        // Method 2: Using forge binary encoding
        conversionMethod = 'forge binary encoding';
        p12Asn1 = forge.asn1.fromDer(forge.util.binary.raw.encode(certData));
      } catch (forgeError) {
        try {
          // Method 3: Manual byte-by-byte conversion
          conversionMethod = 'manual byte conversion';
          let binaryString = '';
          for (let i = 0; i < certData.length; i++) {
            binaryString += String.fromCharCode(certData[i]);
          }
          p12Asn1 = forge.asn1.fromDer(binaryString);
        } catch (manualError) {
          // Final attempt: Check if file might be base64 encoded
          try {
            conversionMethod = 'base64 decoding';
            const base64Data = certData.toString('utf8').replace(/\s/g, '');
            const decodedData = Buffer.from(base64Data, 'base64');
            p12Asn1 = forge.asn1.fromDer(decodedData.toString('binary'));
          } catch (base64Error) {
            throw new Error(`All thumbprint parsing methods failed: ${binaryError.message}`);
          }
        }
      }
    }
    
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