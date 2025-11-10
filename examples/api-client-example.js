/**
 * SharePoint API Client Example
 * Demonstrates how to authenticate and use the SharePoint MCP Server API
 */

// Your API configuration - PRODUCTION SERVER
const SHAREPOINT_API_CONFIG = {
  // Your live server URL
  baseURL: process.env.SHAREPOINT_API_URL || 'https://your-server.onrender.com',
  
  // Your secure API token (64 characters)
  apiToken: process.env.SHAREPOINT_API_TOKEN || 'your-64-character-api-token-here',
  
  // Request headers with authentication
  headers: {
    'Authorization': `Bearer ${process.env.SHAREPOINT_API_TOKEN || 'your-64-character-api-token-here'}`,
    'Content-Type': 'application/json'
  }
};

/**
 * SharePoint API Client Class
 */
class SharePointAPIClient {
  constructor(config = SHAREPOINT_API_CONFIG) {
    this.config = config;
  }

  /**
   * Make authenticated API request
   */
  async request(endpoint, options = {}) {
    const url = `${this.config.baseURL}${endpoint}`;
    const config = {
      headers: this.config.headers,
      ...options
    };

    try {
      const response = await fetch(url, config);
      
      if (!response.ok) {
        if (response.status === 401) {
          throw new Error('Authentication failed - check your API token');
        }
        throw new Error(`API request failed: ${response.status} ${response.statusText}`);
      }

      return await response.json();
    } catch (error) {
      console.error(`SharePoint API Error:`, error.message);
      throw error;
    }
  }

  /**
   * Validate API token
   */
  async validateToken() {
    return await this.request('/api/auth/validate');
  }

  /**
   * Get server health status
   */
  async getHealth() {
    // Health endpoint doesn't require authentication
    const response = await fetch(`${this.config.baseURL}/health`);
    return await response.json();
  }

  /**
   * List SharePoint folders
   */
  async listFolders(parentFolder = '') {
    const params = parentFolder ? `?parentFolder=${encodeURIComponent(parentFolder)}` : '';
    return await this.request(`/api/folders${params}`);
  }

  /**
   * List SharePoint documents
   */
  async listDocuments(folderName = '') {
    const params = folderName ? `?folderName=${encodeURIComponent(folderName)}` : '';
    return await this.request(`/api/documents${params}`);
  }

  /**
   * Get folder tree structure
   */
  async getFolderTree(folderPath = '', maxDepth = 3) {
    const params = new URLSearchParams();
    if (folderPath) params.append('folderPath', folderPath);
    if (maxDepth) params.append('maxDepth', maxDepth);
    
    const queryString = params.toString() ? `?${params.toString()}` : '';
    return await this.request(`/api/tree${queryString}`);
  }

  /**
   * Get document content
   */
  async getDocumentContent(documentPath) {
    return await this.request(`/api/document/${encodeURIComponent(documentPath)}/content`);
  }

  /**
   * Upload document
   */
  async uploadDocument(fileName, content, folderPath = '', overwrite = false) {
    return await this.request('/api/upload', {
      method: 'POST',
      body: JSON.stringify({
        fileName,
        content,
        folderPath,
        overwrite
      })
    });
  }
}

/**
 * Example usage
 */
async function demonstrateAPI() {
  const client = new SharePointAPIClient();

  try {
    console.log('🔍 Testing SharePoint API Client...\n');

    // 1. Check server health
    console.log('1. Server Health Check:');
    const health = await client.getHealth();
    console.log(health);

    // 2. Validate API token
    console.log('\n2. Validate API Token:');
    const tokenValidation = await client.validateToken();
    console.log(tokenValidation);

    // 3. List root folders
    console.log('\n3. List Root Folders:');
    const folders = await client.listFolders();
    console.log(`Found ${folders.count} folders:`, folders.items);

    // 4. List documents in root
    console.log('\n4. List Root Documents:');
    const documents = await client.listDocuments();
    console.log(`Found ${documents.count} documents:`, documents.items);

    // 5. Get folder tree
    console.log('\n5. Get Folder Tree:');
    const tree = await client.getFolderTree('', 2);
    console.log('Folder structure:', tree);

    console.log('\n✅ API demonstration completed successfully!');

  } catch (error) {
    console.error('\n❌ API demonstration failed:', error.message);
    
    if (error.message.includes('Authentication failed')) {
      console.log('\n💡 Fix: Update the apiToken in SHAREPOINT_API_CONFIG with your server\'s API token');
    }
  }
}

/**
 * WebSocket Example
 */
function demonstrateWebSocket() {
  // Convert HTTPS URL to WSS (WebSocket Secure) for production
  const wsUrl = SHAREPOINT_API_CONFIG.baseURL.replace('https://', 'wss://') + `?token=${SHAREPOINT_API_CONFIG.apiToken}`;
  
  console.log('🔌 Connecting to WebSocket:', wsUrl.replace(SHAREPOINT_API_CONFIG.apiToken, 'TOKEN_HIDDEN'));
  const ws = new WebSocket(wsUrl);

  ws.onopen = () => {
    console.log('✅ WebSocket connected');
    
    // Send a request
    ws.send(JSON.stringify({
      action: 'listFolders',
      params: { parentFolder: '' }
    }));
  };

  ws.onmessage = (event) => {
    const response = JSON.parse(event.data);
    console.log('📨 WebSocket response:', response);
  };

  ws.onerror = (error) => {
    console.error('❌ WebSocket error:', error);
  };

  ws.onclose = (event) => {
    console.log('🔌 WebSocket closed:', event.code, event.reason);
  };
}

// Export for use in other modules
export { SharePointAPIClient, SHAREPOINT_API_CONFIG };

// Run demonstration if this file is executed directly
if (typeof window === 'undefined' && import.meta.url === `file://${process.argv[1]}`) {
  demonstrateAPI();
}