#!/usr/bin/env node

// Example client for SharePoint API Server
import axios from 'axios';

class SharePointAPIClient {
  constructor(baseURL, authToken) {
    this.baseURL = baseURL;
    this.authToken = authToken;
    
    this.client = axios.create({
      baseURL: this.baseURL,
      headers: {
        'Authorization': `Bearer ${this.authToken}`,
        'Content-Type': 'application/json'
      }
    });
  }

  async listFolders(parentFolder = '') {
    const response = await this.client.get('/api/folders', {
      params: { parentFolder }
    });
    return response.data;
  }

  async listDocuments(folderName = '') {
    const response = await this.client.get('/api/documents', {
      params: { folderName }
    });
    return response.data;
  }

  async getFolderTree(folderPath = '', maxDepth = 5) {
    const response = await this.client.get('/api/tree', {
      params: { folderPath, maxDepth }
    });
    return response.data;
  }

  async getDocumentContent(documentPath) {
    const response = await this.client.get(`/api/document/${encodeURIComponent(documentPath)}/content`);
    return response.data;
  }

  async checkHealth() {
    const response = await this.client.get('/health');
    return response.data;
  }
}

// Example usage
async function example() {
  const client = new SharePointAPIClient(
    'http://localhost:3000',
    'your-api-token-here'
  );

  try {
    console.log('🔍 Checking server health...');
    const health = await client.checkHealth();
    console.log('Health:', health);

    console.log('\n📁 Listing root folders...');
    const folders = await client.listFolders();
    console.log('Folders:', JSON.stringify(folders, null, 2));

    console.log('\n📄 Listing root documents...');
    const documents = await client.listDocuments();
    console.log('Documents:', JSON.stringify(documents, null, 2));

    console.log('\n🌳 Getting folder tree...');
    const tree = await client.getFolderTree('', 2);
    console.log('Tree:', JSON.stringify(tree, null, 2));

  } catch (error) {
    console.error('❌ Error:', error.response?.data || error.message);
  }
}

// Run example if called directly
if (import.meta.url === `file://${process.argv[1]}`) {
  example();
}

export { SharePointAPIClient };