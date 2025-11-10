#!/usr/bin/env node

/**
 * SharePoint MCP Server - Credential Validation Script
 * This script validates all required credentials and configurations before starting the server
 */

import { loadConfig, getConfigLoader } from '../src/config.js';
import { createSharePointClient } from '../src/sharepoint-client.js';
import fs from 'fs';
import path from 'path';

const VALIDATION_RESULTS = {
  config: false,
  certificate: false,
  sharepoint: false,
  permissions: false
};

console.log('🔍 SharePoint MCP Server - Credential Validation');
console.log('===============================================');

async function validateConfiguration() {
  try {
    console.log('\n1️⃣  Validating configuration...');
    const config = await loadConfig();
    console.log('   ✅ Configuration loaded successfully');
    console.log(`   📋 Server mode: ${config.server.mode}`);
    console.log(`   📋 SharePoint site: ${config.sharepoint.siteUrl}`);
    console.log(`   📋 Certificate path: ${config.sharepoint.certPath}`);
    VALIDATION_RESULTS.config = true;
    return config;
  } catch (error) {
    console.error('   ❌ Configuration validation failed:', error.message);
    return null;
  }
}

async function validateCertificate(config) {
  try {
    console.log('\n2️⃣  Validating certificate...');
    const certPath = config.sharepoint.certPath;
    
    // Check file exists
    if (!fs.existsSync(certPath)) {
      throw new Error(`Certificate file not found: ${certPath}`);
    }
    
    // Check file permissions
    const stats = fs.statSync(certPath);
    const permissions = (stats.mode & parseInt('777', 8)).toString(8);
    
    if (permissions !== '600' && permissions !== '400') {
      console.warn(`   ⚠️  Certificate file has permissions ${permissions}, recommend 600`);
    }
    
    // Check file size (basic validation)
    if (stats.size < 100) {
      throw new Error('Certificate file appears to be empty or corrupted');
    }
    
    console.log('   ✅ Certificate file found and accessible');
    console.log(`   📋 File size: ${stats.size} bytes`);
    console.log(`   📋 Permissions: ${permissions}`);
    
    VALIDATION_RESULTS.certificate = true;
    return true;
  } catch (error) {
    console.error('   ❌ Certificate validation failed:', error.message);
    return false;
  }
}

async function validateSharePointConnection() {
  try {
    console.log('\n3️⃣  Validating SharePoint connection...');
    console.log('   🔄 Attempting to authenticate with SharePoint...');
    
    const client = await createSharePointClient();
    
    console.log('   ✅ Successfully authenticated with SharePoint');
    console.log(`   📋 Site ID: ${client.siteId}`);
    console.log(`   📋 Available drives: ${client.drives.length}`);
    
    // Test basic functionality
    console.log('   🔄 Testing folder listing...');
    const folders = await client.getFolders('');
    console.log(`   ✅ Successfully retrieved ${folders.length} root folders`);
    
    VALIDATION_RESULTS.sharepoint = true;
    return true;
  } catch (error) {
    console.error('   ❌ SharePoint connection failed:', error.message);
    console.error('   💡 Check your credentials, certificate, and network connectivity');
    return false;
  }
}

async function validatePermissions() {
  try {
    console.log('\n4️⃣  Validating file system permissions...');
    
    // Check current user
    const userId = process.getuid();
    const groupId = process.getgid();
    console.log(`   📋 Running as user ID: ${userId}, group ID: ${groupId}`);
    
    // Check write permissions for logs (if in production layout)
    const logPaths = [
      '/var/log/sharepoint-mcp',
      './logs',
      '.'
    ];
    
    let canWriteLogs = false;
    for (const logPath of logPaths) {
      try {
        if (fs.existsSync(logPath)) {
          fs.accessSync(logPath, fs.constants.W_OK);
          console.log(`   ✅ Can write to log directory: ${logPath}`);
          canWriteLogs = true;
          break;
        }
      } catch (error) {
        // Continue to next path
      }
    }
    
    if (!canWriteLogs) {
      console.warn('   ⚠️  No writable log directory found, using console logging');
    }
    
    VALIDATION_RESULTS.permissions = true;
    return true;
  } catch (error) {
    console.error('   ❌ Permission validation failed:', error.message);
    return false;
  }
}

function printValidationSummary() {
  console.log('\n📊 Validation Summary');
  console.log('===================');
  
  const results = [
    { name: 'Configuration', status: VALIDATION_RESULTS.config },
    { name: 'Certificate', status: VALIDATION_RESULTS.certificate },
    { name: 'SharePoint Connection', status: VALIDATION_RESULTS.sharepoint },
    { name: 'File Permissions', status: VALIDATION_RESULTS.permissions }
  ];
  
  results.forEach(result => {
    const icon = result.status ? '✅' : '❌';
    const status = result.status ? 'PASS' : 'FAIL';
    console.log(`   ${icon} ${result.name}: ${status}`);
  });
  
  const allPassed = Object.values(VALIDATION_RESULTS).every(result => result);
  
  console.log('\n' + '='.repeat(50));
  
  if (allPassed) {
    console.log('🎉 ALL VALIDATIONS PASSED!');
    console.log('🚀 Your SharePoint MCP Server is ready to start');
    console.log('\n💡 To start the server:');
    console.log('   MCP mode: node src/server.js mcp');
    console.log('   API mode: SERVER_MODE=api node src/server.js');
    process.exit(0);
  } else {
    console.log('❌ VALIDATION FAILED!');
    console.log('🔧 Please fix the issues above before starting the server');
    
    console.log('\n💡 Common solutions:');
    console.log('   • Check your .env file configuration');
    console.log('   • Verify certificate file path and permissions');
    console.log('   • Ensure SharePoint site URL and credentials are correct');
    console.log('   • Check network connectivity to SharePoint');
    
    process.exit(1);
  }
}

async function main() {
  const config = await validateConfiguration();
  
  if (!config) {
    console.log('\n❌ Cannot proceed without valid configuration');
    process.exit(1);
  }
  
  await validateCertificate(config);
  await validateSharePointConnection();
  await validatePermissions();
  
  printValidationSummary();
}

// Handle process termination gracefully
process.on('SIGINT', () => {
  console.log('\n\n🛑 Validation interrupted by user');
  process.exit(1);
});

process.on('unhandledRejection', (error) => {
  console.error('\n❌ Unhandled error during validation:', error.message);
  process.exit(1);
});

// Run validation
main().catch((error) => {
  console.error('\n❌ Validation script failed:', error.message);
  process.exit(1);
});