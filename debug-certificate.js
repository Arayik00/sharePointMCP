#!/usr/bin/env node

/**
 * Certificate debugging script for Render deployment
 */

console.log('🔍 Certificate Debugging Information');
console.log('====================================');

console.log('\n📋 Environment Variables:');
console.log('SHP_CERT_PFX_PATH:', process.env.SHP_CERT_PFX_PATH);
console.log('SHP_CERT_PFX_PASSWORD:', process.env.SHP_CERT_PFX_PASSWORD ? '[SET]' : '[NOT SET]');

console.log('\n📂 Working Directory:', process.cwd());

import fs from 'fs';
import path from 'path';

const certPath = process.env.SHP_CERT_PFX_PATH;

if (!certPath) {
  console.log('❌ SHP_CERT_PFX_PATH not set');
  process.exit(1);
}

console.log('\n🔍 Certificate Path Check:');
console.log('Path:', certPath);
console.log('Is Absolute:', path.isAbsolute(certPath));

try {
  console.log('📂 Directory listing of:', path.dirname(certPath));
  const dir = path.dirname(certPath);
  if (fs.existsSync(dir)) {
    const files = fs.readdirSync(dir);
    console.log('Files in directory:', files);
  } else {
    console.log('❌ Directory does not exist');
  }
} catch (error) {
  console.log('❌ Error listing directory:', error.message);
}

console.log('\n📄 Certificate File Check:');
try {
  if (fs.existsSync(certPath)) {
    const stats = fs.statSync(certPath);
    console.log('✅ File exists');
    console.log('File size:', stats.size, 'bytes');
    console.log('File permissions:', (stats.mode & parseInt('777', 8)).toString(8));
    
    // Try to read first few bytes
    const buffer = fs.readFileSync(certPath, { encoding: null });
    console.log('First 16 bytes (hex):', buffer.slice(0, 16).toString('hex'));
    
  } else {
    console.log('❌ File does not exist');
    
    // Check common alternative paths
    const alternatives = [
      './certificate.pfx',
      '/opt/render/project/certificate.pfx',
      process.cwd() + '/certificate.pfx'
    ];
    
    console.log('\n🔍 Checking alternative paths:');
    alternatives.forEach(alt => {
      if (fs.existsSync(alt)) {
        console.log('✅ Found at:', alt);
      } else {
        console.log('❌ Not found:', alt);
      }
    });
  }
} catch (error) {
  console.log('❌ Error checking file:', error.message);
}