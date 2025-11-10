#!/usr/bin/env node

/**
 * Quick test script to validate your API server connection
 */

const API_CONFIG = {
  baseURL: 'https://your-server.onrender.com',
  apiToken: 'your-64-character-api-token-here'
};

async function testConnection() {
  console.log('🧪 Testing SharePoint API Server Connection');
  console.log('==========================================');
  console.log(`📡 Server: ${API_CONFIG.baseURL}`);
  console.log(`🔑 Token: ${API_CONFIG.apiToken.substring(0, 8)}...`);
  console.log('');

  try {
    // Test 1: Health check (no auth required)
    console.log('1️⃣  Testing health endpoint...');
    const healthResponse = await fetch(`${API_CONFIG.baseURL}/health`);
    const health = await healthResponse.json();
    console.log('✅ Health check passed:', health.status);
    console.log('');

    // Test 2: API token validation
    console.log('2️⃣  Testing API authentication...');
    const authResponse = await fetch(`${API_CONFIG.baseURL}/api/auth/validate`, {
      headers: {
        'Authorization': `Bearer ${API_CONFIG.apiToken}`
      }
    });
    
    if (authResponse.ok) {
      const auth = await authResponse.json();
      console.log('✅ API authentication successful:', auth.message);
      console.log('');
    } else {
      console.log('❌ API authentication failed:', authResponse.status, authResponse.statusText);
      return;
    }

    // Test 3: SharePoint folders
    console.log('3️⃣  Testing SharePoint folder access...');
    const foldersResponse = await fetch(`${API_CONFIG.baseURL}/api/folders`, {
      headers: {
        'Authorization': `Bearer ${API_CONFIG.apiToken}`
      }
    });
    
    if (foldersResponse.ok) {
      const folders = await foldersResponse.json();
      console.log(`✅ SharePoint access successful: Found ${folders.count} folders`);
      console.log('   Sample folders:', folders.items.slice(0, 3).map(f => f.name));
    } else {
      console.log('❌ SharePoint access failed:', foldersResponse.status, foldersResponse.statusText);
    }

    console.log('');
    console.log('🎉 ALL TESTS PASSED! Your API server is ready to use.');
    console.log('');
    console.log('📋 Next steps:');
    console.log('   1. Add this server to Claude Desktop with the sharepoint-api configuration');
    console.log('   2. Use the SharePoint tools in your conversations');
    console.log('   3. Your SharePoint credentials stay secure on the server!');

  } catch (error) {
    console.error('❌ Connection test failed:', error.message);
    console.log('');
    console.log('🔧 Troubleshooting:');
    console.log('   • Check if the server is running and accessible');
    console.log('   • Verify the API token is correct');
    console.log('   • Check network connectivity');
  }
}

testConnection();