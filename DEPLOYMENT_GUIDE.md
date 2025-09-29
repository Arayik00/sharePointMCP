# 🔐 Secure SharePoint MCP Server Deployment

## 🏗️ **Architecture Overview**

Your SharePoint MCP server now supports **two deployment modes**:

### **Mode 1: API Server (Recommended for Production)**
- ✅ **Server-side authentication**: All SharePoint credentials stored securely on server
- ✅ **Custom API tokens**: Clients authenticate with your own tokens
- ✅ **HTTP/WebSocket API**: RESTful endpoints + real-time WebSocket
- ✅ **Multi-client support**: Multiple clients can connect simultaneously

### **Mode 2: Traditional MCP Server**
- ✅ **Direct MCP protocol**: For Claude Desktop integration
- ✅ **Single client**: Designed for one-to-one communication

---

## 🔧 **Server Environment Setup**

### **1. Environment Variables (`.env` file on server):**
```bash
# SharePoint Configuration (SECURE - Server Only)
SHP_ID_APP=your-azure-app-id
SHP_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site
SHP_DOC_LIBRARY=
SHP_TENANT_ID=your-tenant-id
SHP_MAX_DEPTH=15
SHP_MAX_FOLDERS_PER_LEVEL=100
SHP_LEVEL_DELAY=0.5
SHP_CERT_PFX_PATH=/opt/mcp-sharepoint/certificate.pfx
SHP_CERT_PFX_PASSWORD=your-cert-password

# API Server Configuration  
SERVER_MODE=api
PORT=3000

# Custom Authentication Tokens (Generate strong tokens!)
API_AUTH_TOKENS=sp_live_sk_abc123xyz789,sp_live_sk_def456uvw012,sp_live_sk_ghi789rst345
```

### **2. Generate Secure API Tokens:**
```bash
# Generate strong random tokens
node -e "console.log('sp_live_sk_' + require('crypto').randomBytes(32).toString('hex'))"
node -e "console.log('sp_live_sk_' + require('crypto').randomBytes(32).toString('hex'))"
node -e "console.log('sp_live_sk_' + require('crypto').randomBytes(32).toString('hex'))"
```

---

## 🚀 **Deployment Commands**

### **Deploy API Server (Recommended)**
```bash
# 1. Deploy code
git clone <your-repo> /opt/mcp-sharepoint
cd /opt/mcp-sharepoint
npm install --production

# 2. Setup environment
cp .env.example .env
# Edit .env with your actual values including API_AUTH_TOKENS

# 3. Upload certificate securely
sudo chmod 600 /opt/mcp-sharepoint/certificate.pfx
sudo chown mcp:mcp /opt/mcp-sharepoint/certificate.pfx

# 4. Start API server
npm run start:api

# OR with PM2 (recommended)
pm2 start ecosystem.config.json --only mcp-sharepoint-api
pm2 save
pm2 startup
```

### **Deploy Traditional MCP Server**
```bash
# Start MCP server
npm run start:mcp

# OR with PM2
pm2 start ecosystem.config.json --only mcp-sharepoint-mcp
```

---

## 🌐 **Client Usage Examples**

### **Option 1: HTTP API Client**
```javascript
import axios from 'axios';

const client = axios.create({
  baseURL: 'https://your-server.com:3000',
  headers: {
    'Authorization': 'Bearer sp_live_sk_abc123xyz789'
  }
});

// List folders
const folders = await client.get('/api/folders?parentFolder=General');

// List documents  
const docs = await client.get('/api/documents?folderName=Reports');

// Get folder tree
const tree = await client.get('/api/tree?folderPath=&maxDepth=3');

// Get document content
const content = await client.get('/api/document/MyFile.txt/content');
```

### **Option 2: WebSocket Client**
```javascript
const ws = new WebSocket('wss://your-server.com:3000?token=sp_live_sk_abc123xyz789');

ws.send(JSON.stringify({
  action: 'listFolders',
  params: { parentFolder: 'General' }
}));

ws.onmessage = (event) => {
  const response = JSON.parse(event.data);
  console.log(response);
};
```

### **Option 3: Traditional MCP (Claude Desktop)**
```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": ["/opt/mcp-sharepoint/src/server.js", "mcp"]
    }
  }
}
```

---

## 🔒 **Security Benefits**

### **✅ Server-Side Security:**
- SharePoint credentials **never leave the server**
- Certificate file stored securely with restricted permissions
- All authentication happens server-side

### **✅ Client-Side Simplicity:**
- Clients only need your custom API token
- No Azure credentials on client machines  
- Easy token rotation and revocation

### **✅ Access Control:**
- Multiple independent API tokens
- Easy to track which client is making requests
- Granular access control (future: token-specific permissions)

---

## 📊 **API Endpoints Reference**

### **Public Endpoints:**
- `GET /health` - Server health check (no auth required)

### **Authenticated Endpoints:**
- `GET /api` - List available endpoints
- `GET /api/folders?parentFolder=path` - List folders
- `GET /api/documents?folderName=path` - List documents  
- `GET /api/tree?folderPath=path&maxDepth=N` - Get folder tree
- `GET /api/document/:path/content` - Get document content
- `POST /api/upload` - Upload file (future)

### **Authentication Methods:**
```bash
# Method 1: Authorization header
curl -H "Authorization: Bearer sp_live_sk_abc123" https://server:3000/api/folders

# Method 2: X-API-Token header  
curl -H "X-API-Token: sp_live_sk_abc123" https://server:3000/api/folders

# Method 3: Query parameter
curl "https://server:3000/api/folders?token=sp_live_sk_abc123"
```

---

## 🎯 **What to Push to Repository**

### **✅ Push These Files:**
```
src/
├── server.js              # New dual-mode entry point
├── api-server.js          # New API server implementation
├── sharepoint-client.js   # Existing SharePoint client
├── sharepoint-tools.js    # Existing tools wrapper
├── sharepoint-server.js   # Existing MCP server
├── index.js               # Legacy entry point (keep for compatibility)
└── utils/
    ├── config.js          # Environment configuration
    └── logger.js          # Logging utility

examples/
└── client-example.js      # API client example

package.json               # Updated with new dependencies & scripts
ecosystem.config.json      # Updated PM2 configuration
docker-compose.yml         # Docker setup
Dockerfile                 # Container definition
.env.example              # Environment template
.gitignore                # Git ignore rules
README.md                 # Documentation
```

### **❌ Never Push These Files:**
```
.env                      # Contains real API tokens & credentials
certificate.pfx          # Azure certificate
node_modules/            # Dependencies
*.log                    # Log files
```

---

## 🎉 **Result: Ultimate Security**

### **Before (Traditional):**
```json
{
  "mcpServers": {
    "sharepoint": {
      "env": {
        "SHP_ID_APP": "your-azure-app-id",
        "SHP_TENANT_ID": "your-tenant-id",
        "SHP_CERT_PFX_PASSWORD": "y8DauLqAYm8a6M1riHPFN"
      }
    }
  }
}
```
❌ **Problem**: SharePoint credentials exposed to every client

### **After (Secure API):**
```javascript
const client = axios.create({
  baseURL: 'https://your-server.com:3000',
  headers: { 'Authorization': 'Bearer sp_live_sk_abc123' }
});
```
✅ **Solution**: Only your custom token on client, SharePoint credentials secure on server

---

**Your SharePoint MCP server is now enterprise-ready with bank-level security! 🔐🚀**