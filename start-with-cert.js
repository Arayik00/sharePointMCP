#!/usr/bin/env node

/**
 * Base64 Certificate Handler for Render Deployment
 * This script handles base64-encoded certificates to avoid binary corruption
 */

import fs from 'fs';
import path from 'path';

console.log('🔄 Base64 Certificate Handler');
console.log('=============================');

// Check if we should use base64 encoded certificate
const useBase64 = process.env.USE_BASE64_CERT === 'true';
const base64CertPath = process.env.SHP_CERT_BASE64_PATH || './hearst-sharepoint.b64';
const binaryCertPath = process.env.SHP_CERT_PFX_PATH || './certificate.pfx';

if (useBase64) {
  console.log('🔧 Base64 mode enabled');
  
  if (!fs.existsSync(base64CertPath)) {
    console.error(`❌ Base64 certificate file not found: ${base64CertPath}`);
    process.exit(1);
  }
  
  try {
    console.log('📂 Reading base64 certificate...');
    const base64Data = fs.readFileSync(base64CertPath, 'utf8').trim();
    
    console.log('🔄 Converting base64 to binary...');
    const binaryData = Buffer.from(base64Data, 'base64');
    
    // Create the binary certificate file
    const outputPath = binaryCertPath;
    fs.writeFileSync(outputPath, binaryData);
    
    console.log(`✅ Binary certificate created: ${outputPath} (${binaryData.length} bytes)`);
    console.log(`📋 First 8 bytes: ${binaryData.slice(0, 8).toString('hex')}`);
    
    // Verify it looks like a PKCS#12 file
    if (binaryData.slice(0, 2).toString('hex').startsWith('30')) {
      console.log('✅ Certificate looks valid (starts with ASN.1 SEQUENCE)');
    } else {
      console.warn('⚠️  Certificate may not be valid PKCS#12 format');
    }
    
  } catch (error) {
    console.error('❌ Error processing base64 certificate:', error.message);
    process.exit(1);
  }
} else {
  console.log('📋 Using binary certificate directly');
}

console.log('🎉 Certificate processing completed');

// Now start the actual server
console.log('🚀 Starting SharePoint MCP Server...');
try {
  await import('./src/server.js');
} catch (error) {
  console.error('❌ Failed to start server:', error);
  process.exit(1);
}