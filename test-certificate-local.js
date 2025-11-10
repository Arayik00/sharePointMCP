#!/usr/bin/env node

/**
 * Test certificate parsing locally to debug issues
 */

import fs from 'fs';
import forge from 'node-forge';

console.log('🧪 Certificate Test Script');
console.log('==========================');

const certPath = './certificate.pfx';
const certPassword = 'y8DauLqAYm8a6M1riHPFN'; // Your certificate password

if (!fs.existsSync(certPath)) {
  console.log('❌ Certificate file not found at:', certPath);
  console.log('Please copy your certificate.pfx to the project root for testing');
  process.exit(1);
}

try {
  const certData = fs.readFileSync(certPath);
  console.log(`📋 Certificate file size: ${certData.length} bytes`);
  console.log(`📋 First 32 bytes (hex): ${certData.slice(0, 32).toString('hex')}`);
  
  // Check magic bytes for PKCS#12
  const magicBytes = certData.slice(0, 4).toString('hex');
  console.log(`📋 Magic bytes: ${magicBytes}`);
  
  if (magicBytes === '308230' || magicBytes.startsWith('3082')) {
    console.log('✅ Looks like a valid PKCS#12 file (starts with ASN.1 sequence)');
  } else {
    console.log('⚠️  Unusual magic bytes - might be corrupted or different format');
  }

  // Test parsing methods
  const methods = [
    {
      name: 'Binary String',
      fn: () => forge.asn1.fromDer(certData.toString('binary'))
    },
    {
      name: 'Forge Binary Encoding', 
      fn: () => forge.asn1.fromDer(forge.util.binary.raw.encode(certData))
    },
    {
      name: 'Manual Byte Conversion',
      fn: () => {
        let binaryString = '';
        for (let i = 0; i < certData.length; i++) {
          binaryString += String.fromCharCode(certData[i]);
        }
        return forge.asn1.fromDer(binaryString);
      }
    },
    {
      name: 'Base64 Decode',
      fn: () => {
        const base64Data = certData.toString('utf8').replace(/\s/g, '');
        const decodedData = Buffer.from(base64Data, 'base64');
        return forge.asn1.fromDer(decodedData.toString('binary'));
      }
    }
  ];

  let successfulMethod = null;
  let p12Asn1 = null;

  for (const method of methods) {
    try {
      console.log(`🔧 Testing ${method.name}...`);
      p12Asn1 = method.fn();
      console.log(`✅ ${method.name} - SUCCESS`);
      successfulMethod = method.name;
      break;
    } catch (error) {
      console.log(`❌ ${method.name} - FAILED: ${error.message}`);
    }
  }

  if (!p12Asn1) {
    console.log('❌ All parsing methods failed');
    process.exit(1);
  }

  console.log(`\n🎉 Successfully parsed certificate using: ${successfulMethod}`);

  // Try to extract the PKCS#12
  const p12 = forge.pkcs12.pkcs12FromAsn1(p12Asn1, certPassword);
  console.log('✅ PKCS#12 structure extracted successfully');

  // Try to get private key
  const keyBags = p12.getBags({ bagType: forge.pki.oids.pkcs8ShroudedKeyBag });
  if (keyBags[forge.pki.oids.pkcs8ShroudedKeyBag] && keyBags[forge.pki.oids.pkcs8ShroudedKeyBag].length > 0) {
    console.log('✅ Private key found in certificate');
  } else {
    console.log('❌ No private key found in certificate');
  }

  // Try to get certificate
  const certBags = p12.getBags({ bagType: forge.pki.oids.certBag });
  if (certBags[forge.pki.oids.certBag] && certBags[forge.pki.oids.certBag].length > 0) {
    console.log('✅ Certificate found in PKCS#12');
    
    const cert = certBags[forge.pki.oids.certBag][0].cert;
    console.log(`📋 Subject: ${cert.subject.getField('CN').value}`);
    console.log(`📋 Issuer: ${cert.issuer.getField('CN').value}`);
    console.log(`📋 Valid from: ${cert.validity.notBefore}`);
    console.log(`📋 Valid to: ${cert.validity.notAfter}`);
    
    // Calculate thumbprint
    const der = forge.asn1.toDer(forge.pki.certificateToAsn1(cert)).getBytes();
    const md = forge.md.sha1.create();
    md.update(der);
    const thumbprint = md.digest().toHex().toUpperCase();
    console.log(`📋 Thumbprint: ${thumbprint}`);
  } else {
    console.log('❌ No certificate found in PKCS#12');
  }

  console.log('\n🎉 Certificate validation completed successfully!');

} catch (error) {
  console.error('❌ Certificate test failed:', error);
  process.exit(1);
}