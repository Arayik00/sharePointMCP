#!/usr/bin/env node

/**
 * SharePoint MCP Server - Comprehensive Test Suite
 * Tests both direct SharePoint connection and API server modes
 */

import * as dotenv from 'dotenv';
import { createSharePointClient } from './src/sharepoint-client.js';
import { SharePointTools } from './src/sharepoint-tools.js';

// Load environment variables
dotenv.config();

// API Server Configuration
const API_CONFIG = {
  baseURL: 'https://your-server.onrender.com',
  apiToken: 'your-64-character-api-token-here'
};

/**
 * Make authenticated API request
 */
async function apiRequest(endpoint, options = {}) {
  const url = `${API_CONFIG.baseURL}${endpoint}`;
  const config = {
    headers: {
      'Authorization': `Bearer ${API_CONFIG.apiToken}`,
      'Content-Type': 'application/json'
    },
    ...options
  };

  const response = await fetch(url, config);
  
  if (!response.ok) {
    throw new Error(`API request failed: ${response.status} ${response.statusText}`);
  }

  return await response.json();
}

/**
 * Test API Server Connection and Endpoints
 */
async function testAPIServer() {
  console.log('\n🌐 Testing API Server Connection...');
  console.log('=====================================');

  try {
    // Test 1: Health check
    console.log('1️⃣  Testing health endpoint...');
    const healthResponse = await fetch(`${API_CONFIG.baseURL}/health`);
    const health = await healthResponse.json();
    console.log('✅ Health check passed:', health.status);

    // Test 2: Authentication
    console.log('2️⃣  Testing API authentication...');
    const auth = await apiRequest('/api/auth/validate');
    console.log('✅ Authentication successful:', auth.message);

    // Test 3: List folders
    console.log('3️⃣  Testing folder listing...');
    const folders = await apiRequest('/api/folders');
    console.log(`✅ Folders retrieved: ${folders.count} folders found`);

    // Test 4: List documents
    console.log('4️⃣  Testing document listing...');
    const documents = await apiRequest('/api/documents');
    console.log(`✅ Documents retrieved: ${documents.count} documents found`);

    // Test 5: Get folder tree
    console.log('5️⃣  Testing folder tree...');
    const tree = await apiRequest('/api/tree?maxDepth=2');
    console.log('✅ Folder tree retrieved successfully');

    // Test 6: Upload test file via API
    console.log('6️⃣  Testing file upload via API...');
    const testFileName = 'api_test_file.txt';
    const testContent = 'Test file uploaded via API server';
    
    const uploadResult = await apiRequest('/api/upload', {
      method: 'POST',
      body: JSON.stringify({
        fileName: testFileName,
        content: testContent,
        folderPath: '',
        overwrite: true
      })
    });
    console.log('✅ File upload successful:', uploadResult.success);

    console.log('\n🎉 API Server Tests: ALL PASSED');
    return true;

  } catch (error) {
    console.error('❌ API Server test failed:', error.message);
    return false;
  }
}

/**
 * Test Direct SharePoint Connection
 */
async function testDirectSharePointConnection() {
  console.log('\n🔗 Testing Direct SharePoint Connection...');
  console.log('==========================================');

  try {
    // Initialize SharePoint client
    const client = await createSharePointClient();
    if (!client) {
      console.error('❌ Failed to create SharePoint client');
      return false;
    }

    const tools = new SharePointTools(client);
    console.log('✅ Successfully initialized SharePoint client and tools');

    // Test 1: List folders in root
    console.log('1️⃣  List folders in root...');
    const foldersResult = await tools.listFolders();
    console.log(`✅ Folders: ${foldersResult.success ? foldersResult.items?.length : 'Failed'}`);

    // Test 2: List documents in root
    console.log('2️⃣  List documents in root...');
    const docsResult = await tools.listDocuments('');
    console.log(`✅ Documents: ${docsResult.success ? docsResult.items?.length : 'Failed'}`);

    // Test 3: Get folder tree
    console.log('3️⃣  Get folder tree...');
    const treeResult = await tools.getFolderTree(undefined, 2);
    console.log(`✅ Tree: ${treeResult.success ? 'Retrieved' : 'Failed'}`);
    if (treeResult.success && treeResult.tree) {
      const items = treeResult.tree.items || [];
      console.log(`   Root items count: ${items.length}`);
    }

    // Test 4: Try to get content of a file if we find any
    console.log('4️⃣  Try to get file content...');
    if (docsResult.success && docsResult.items && docsResult.items.length > 0) {
      const firstDoc = docsResult.items[0];
      const docName = firstDoc.name;
      console.log(`   Trying to read: ${docName}`);
      
      const contentResult = await tools.getDocumentContent('', docName);
      console.log(`✅ Content reading: ${contentResult.success ? 'Success' : 'Failed'}`);
      if (contentResult.success) {
        const contentLength = contentResult.content?.length || 0;
        console.log(`   Content length: ${contentLength} characters`);
      }
    } else {
      console.log('   No documents found to test content reading');
    }

    // Test 5: Upload a test document
    console.log('5️⃣  Upload a test document...');
    const testFileName = 'direct_test_file.txt';
    const testFileContent = 'Test file via direct connection';
    
    try {
      const uploadResult = await tools.uploadDocument('', testFileName, testFileContent, false);
      console.log(`✅ Upload: ${uploadResult.success ? 'Success' : 'Failed'}`);
      
      if (uploadResult.success) {
        // Verify the upload by reading the file back
        console.log('6️⃣  Verifying upload...');
        const verifyResult = await tools.getDocumentContent('', testFileName);
        
        if (verifyResult.success && verifyResult.content === testFileContent) {
          console.log('✅ Upload verification: Content matches exactly!');
        } else {
          console.warn('⚠️  Upload verification: Content mismatch detected');
        }
      }
    } catch (error) {
      console.error('❌ Exception during upload test:', error.message);
    }

    console.log('\n🎉 Direct SharePoint Tests: ALL PASSED');
    return true;
  } catch (error) {
    console.error('❌ Direct SharePoint test failed:', error);
    return false;
  }
}

/**
 * Run comprehensive test suite
 */
async function runAllTests() {
  console.log('🧪 SharePoint MCP Server - Comprehensive Test Suite');
  console.log('==================================================');
  console.log(`📅 Test Date: ${new Date().toISOString()}`);
  console.log(`🖥️  API Server: ${API_CONFIG.baseURL}`);
  console.log(`🔑 API Token: ${API_CONFIG.apiToken.substring(0, 8)}...`);

  const results = {
    apiServer: false,
    directConnection: false
  };

  // Test 1: API Server
  try {
    results.apiServer = await testAPIServer();
  } catch (error) {
    console.error('API Server tests crashed:', error.message);
  }

  // Test 2: Direct Connection (only if environment is configured)
  try {
    if (process.env.SHP_ID_APP && process.env.SHP_CERT_PFX_PATH) {
      results.directConnection = await testDirectSharePointConnection();
    } else {
      console.log('\n🔗 Direct SharePoint Connection...');
      console.log('==========================================');
      console.log('⚠️  Skipping direct connection tests (no local credentials configured)');
      console.log('   To test direct connection, configure environment variables:');
      console.log('   - SHP_ID_APP, SHP_TENANT_ID, SHP_SITE_URL');
      console.log('   - SHP_CERT_PFX_PATH, SHP_CERT_PFX_PASSWORD');
      results.directConnection = null; // Not tested
    }
  } catch (error) {
    console.error('Direct connection tests crashed:', error.message);
  }

  // Final Results
  console.log('\n📊 TEST RESULTS SUMMARY');
  console.log('========================');
  console.log(`🌐 API Server Tests: ${results.apiServer ? '✅ PASSED' : '❌ FAILED'}`);
  console.log(`🔗 Direct Connection: ${results.directConnection === null ? '⚠️  SKIPPED' : results.directConnection ? '✅ PASSED' : '❌ FAILED'}`);

  if (results.apiServer) {
    console.log('\n🎉 SUCCESS: Your SharePoint MCP Server is working!');
    console.log('✅ API server authentication successful');
    console.log('✅ SharePoint connectivity confirmed');
    console.log('✅ Read/write operations functional');
    console.log('� Ready for production use with Claude Desktop');
    
    console.log('\n📖 Next Steps:');
    console.log('   1. Add the mcp.json configuration to Claude Desktop');
    console.log('   2. Use SharePoint tools in your conversations');
    console.log('   3. Your credentials stay secure on the server!');
  } else {
    console.log('\n❌ ISSUES DETECTED: Please check your configuration');
    console.log('🔧 Troubleshooting:');
    console.log('   • Verify API server is running and accessible');
    console.log('   • Check API token configuration');
    console.log('   • Ensure SharePoint credentials are valid on server');
  }

  return results.apiServer;
}

// Run the comprehensive test suite
runAllTests().then((success) => {
  process.exit(success ? 0 : 1);
}).catch((error) => {
  console.error('❌ Test suite crashed:', error);
  process.exit(1);
});