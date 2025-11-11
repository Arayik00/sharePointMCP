# 🚀 SharePoint MCP Server - Next Generation

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Node.js](https://img.shields.io/badge/Node.js-18+-green.svg)](https://nodejs.org/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-blue.svg)](https://modelcontextprotocol.io/)

A **next-generation MCP Server** for seamless integration with Microsoft SharePoint, providing both traditional MCP protocol support and modern API capabilities. Built with JavaScript/Node.js for efficiency and enterprise-grade security.

> **� Based on Original Work**: This project is inspired by and builds upon the excellent [mcp-sharepoint](https://github.com/Sofias-ai/mcp-sharepoint) Python implementation by Sofias-ai.
>
> **🔄 Enhanced JavaScript Port**: We've completely rewritten the codebase in JavaScript and added significant enhancements:
> - **🔐 Enhanced Security**: Upgraded from client secret to certificate-based authentication
> - **🏗️ Dual-Mode Architecture**: Added HTTP/WebSocket API server alongside traditional MCP
> - **🛡️ Enterprise Security**: Server-side credential storage with custom API tokens
> - **🚀 Production Ready**: Docker, PM2, and systemd deployment configurations
> - **🌐 Multi-Client Support**: Simultaneous client connections via HTTP/WebSocket

---

## 🎯 **What We've Built**

### **🏗️ Dual-Mode Architecture**

#### **Mode 1: Traditional MCP Server**
- ✅ **100% MCP Protocol Compliant** - Works with Claude Desktop
- ✅ **Stdio Communication** - Standard MCP transport
- ✅ **Tool-based Interface** - Registered MCP tools for SharePoint operations

#### **Mode 2: Secure API Server**
- ✅ **HTTP REST API** - RESTful endpoints for all SharePoint operations
- ✅ **WebSocket Support** - Real-time communication capabilities
- ✅ **Custom Authentication** - Your own API tokens instead of SharePoint credentials
- ✅ **Multi-Client Support** - Multiple clients can connect simultaneously
- ✅ **Enterprise Security** - Server-side credential storage

---

## ✨ **Key Features**

### **📁 Complete CRUD Operations**
- **� READ** - List folders, browse documents, get content, navigate trees
- **� CREATE** - Upload documents, create folders, organize content
- **✏️ UPDATE** - Modify document content, update file properties
- **🗑️ DELETE** - Remove files and folders with comprehensive cleanup
- **🌳 Tree Navigation** - Recursive folder structures with configurable depth
- **🔍 Content Access** - Document content with automatic text/binary detection
- **🏷️ Rich Metadata** - File properties, sizes, modification dates, SharePoint URLs

### **🔐 Security & Authentication**
- **🔒 Certificate Authentication** - Secure X.509 certificate-based auth
- **🛡️ Server-side Credentials** - SharePoint credentials never leave the server
- **🎫 Custom API Tokens** - Generate your own authentication tokens
- **🔑 Token Management** - Multiple tokens for different clients/environments
- **🚫 Zero Credential Exposure** - Clients only need your custom tokens

### **🚀 Deployment Options**
- **🐳 Docker Support** - Containerized deployment
- **⚙️ PM2 Configuration** - Production process management
- **🔄 Systemd Service** - Linux service integration
- **📊 Health Monitoring** - Built-in health checks and logging

---

## 🛠️ **Available Operations**

### **Traditional MCP Tools** (Claude Desktop)
1. **`List_SharePoint_Folders`** - List folders in specified directory
2. **`List_SharePoint_Documents`** - List documents in a folder
3. **`Get_SharePoint_Tree`** - Get recursive folder tree structure
4. **`Get_SharePoint_Document_Content`** - Read document content
5. **`Create_SharePoint_Folder`** - Create new folders
6. **`Upload_SharePoint_Document`** - Upload documents
7. **`Upload_SharePoint_Document_From_Path`** - Upload from local path
8. **`Update_SharePoint_Document`** - Update existing documents
9. **`Delete_SharePoint_Item`** - Delete files or folders

### **API Server Endpoints** (HTTP/WebSocket)
**📖 READ Operations:**
- **`GET /api/folders`** - List folders
- **`GET /api/documents`** - List documents
- **`GET /api/tree`** - Get folder tree
- **`GET /api/document/:path/content`** - Get document content

**📝 CREATE Operations:**
- **`POST /api/upload`** - Upload documents
- **`POST /api/folder`** - Create new folders

**✏️ UPDATE Operations:**
- **`PUT /api/document/:path`** - Update document content

**🗑️ DELETE Operations:**
- **`DELETE /api/item/:path`** - Delete any file or folder
- **`DELETE /api/document/:path`** - Delete document (alias)
- **`DELETE /api/folder/:path`** - Delete folder (alias)
- **`POST /api/upload`** - Upload files
- **`GET /health`** - Server health check

---

## 🚀 **Quick Start**

### **1. Installation**
```bash
git clone <your-repo>
cd sharePointMCP
npm install
```

### **2. Environment Setup**
```bash
# Copy environment template
cp .env.example .env

# Edit .env with your values:
SHP_ID_APP=your-azure-app-id
SHP_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site
SHP_TENANT_ID=your-tenant-id
SHP_CERT_PFX_PATH=./certificate.pfx
SHP_CERT_PFX_PASSWORD=your-cert-password

# For API mode, add custom tokens:
API_AUTH_TOKENS=your-secret-token-1,your-secret-token-2
```

### **3. Run the Server**

#### **MCP Mode (Claude Desktop)**
```bash
npm run start:mcp
```

#### **API Mode (HTTP/WebSocket)**
```bash
npm run start:api
# Server will be available at http://localhost:3000
```

---

## 📋 **Usage Examples**

### **Claude Desktop Integration**
Add to your `mcp.json`:
```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": ["/path/to/sharePointMCP/src/server.js", "mcp"]
    }
  }
}
```

### **HTTP API Usage**
```bash
# List folders (with authentication)
curl -H "Authorization: Bearer your-api-token" \
     http://localhost:3000/api/folders

# Get document content
curl -H "Authorization: Bearer your-api-token" \
     http://localhost:3000/api/document/MyFile.txt/content

# Get folder tree
curl -H "Authorization: Bearer your-api-token" \
     "http://localhost:3000/api/tree?maxDepth=3"
```

### **JavaScript Client Example**
```javascript
import axios from 'axios';

const client = axios.create({
  baseURL: 'http://localhost:3000',
  headers: { 'Authorization': 'Bearer your-api-token' }
});

// List folders
const folders = await client.get('/api/folders');

// Get document content
const content = await client.get('/api/document/report.pdf/content');
```

### **WebSocket Example**
```javascript
const ws = new WebSocket('ws://localhost:3000?token=your-api-token');

ws.send(JSON.stringify({
  action: 'listFolders',
  params: { parentFolder: 'Documents' }
}));

ws.onmessage = (event) => {
  const response = JSON.parse(event.data);
  console.log('Folders:', response);
};
```

---

## 🔧 **Environment Variables**

### **SharePoint Configuration**
```bash
SHP_ID_APP=your-azure-app-id               # Azure AD Application ID
SHP_SITE_URL=https://tenant.sharepoint.com # SharePoint site URL
SHP_DOC_LIBRARY=Shared Documents           # Document library path
SHP_TENANT_ID=your-tenant-id               # Microsoft tenant ID
SHP_CERT_PFX_PATH=./certificate.pfx        # Certificate file path
SHP_CERT_PFX_PASSWORD=cert-password        # Certificate password
```

### **Server Configuration**
```bash
SERVER_MODE=api                             # Mode: 'mcp' or 'api'
PORT=3000                                   # API server port
API_AUTH_TOKENS=token1,token2,token3        # Custom authentication tokens
NODE_ENV=production                         # Environment mode
```

### **Operational Settings**
```bash
SHP_MAX_DEPTH=15                           # Maximum folder tree depth
SHP_MAX_FOLDERS_PER_LEVEL=100              # Max folders per level
SHP_LEVEL_DELAY=0.5                        # Delay between levels (seconds)
```

---

## � **Server-Side Credential Management**

### **🛡️ Why Server-Side Credentials?**

This MCP server implements **enterprise-grade security** by storing SharePoint credentials securely on the server and providing your own API authentication tokens. This approach offers several advantages:

- **🔒 Enhanced Security** - SharePoint credentials never leave the server
- **🎫 Custom Authentication** - Generate your own API tokens
- **🔄 Easy Rotation** - Update credentials without touching client configurations
- **📊 Audit Trail** - Track API usage with your own tokens
- **🚀 Scalability** - Multiple clients can use the same SharePoint connection

### **📁 Secure Configuration Structure**

```bash
/opt/sharepoint-mcp/                    # Production directory structure
├── config/
│   └── .env.production                 # Environment variables (600 permissions)
├── certs/
│   └── certificate.pfx                 # X.509 certificate (600 permissions)
├── src/                               # Application source code
├── logs/                              # Application logs
└── scripts/                           # Deployment scripts
```

### **⚙️ Configuration Management**

The server uses a **secure configuration loader** that:

1. **Validates all credentials** before starting
2. **Checks certificate file permissions** and accessibility  
3. **Validates Azure AD configuration** (GUIDs, URLs)
4. **Tests SharePoint connectivity** during startup
5. **Provides detailed error messages** for troubleshooting

### **🚀 Quick Setup Commands**

```bash
# 1. Validate your configuration and credentials
npm run validate

# 2. Set up secure deployment structure
npm run setup:secure

# 3. Configure systemd service (Linux)
npm run setup:systemd

# 4. Start with validated configuration
npm start
```

### **🔧 Environment File Structure**

Create `.env.production` with your server credentials:

```bash
# SharePoint Configuration
SHP_ID_APP=12345678-1234-1234-1234-123456789012
SHP_TENANT_ID=87654321-4321-4321-4321-210987654321  
SHP_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site
SHP_CERT_PFX_PATH=/opt/sharepoint-mcp/certs/certificate.pfx
SHP_CERT_PFX_PASSWORD=your-certificate-password

# Server Configuration  
SERVER_MODE=api                                      # or 'mcp' for MCP-only
PORT=3000
LOG_LEVEL=info

# API Security (Generate with: openssl rand -hex 32)
API_AUTH_TOKENS=64-char-token-1,64-char-token-2
CORS_ORIGINS=https://your-domain.com
RATE_LIMIT_WINDOW=900000                            # 15 minutes
RATE_LIMIT_MAX=100                                  # requests per window
```

### **🔍 Credential Validation**

Before starting the server, validate your setup:

```bash
npm run validate
```

This comprehensive validation checks:
- ✅ Configuration file syntax and required values
- ✅ Certificate file existence and permissions  
- ✅ SharePoint authentication and connectivity
- ✅ File system permissions and security
- ✅ Network connectivity to Microsoft Graph API

### **⚡ Production Deployment**

```bash
# 1. Set up secure directory structure
sudo ./scripts/setup-secure-deployment.sh

# 2. Copy your files to secure locations
sudo cp certificate.pfx /opt/sharepoint-mcp/certs/
sudo cp .env.production /opt/sharepoint-mcp/config/

# 3. Set secure file permissions  
sudo chmod 600 /opt/sharepoint-mcp/certs/certificate.pfx
sudo chmod 600 /opt/sharepoint-mcp/config/.env.production

# 4. Configure systemd service
sudo ./scripts/deploy-systemd.sh

# 5. Start the service
sudo systemctl start sharepoint-mcp
```

### **🎫 Client Authentication**

Clients authenticate using **your custom API tokens** instead of SharePoint credentials:

```javascript
// HTTP API Authentication
const response = await fetch('http://your-server:3000/api/folders', {
  headers: { 'Authorization': 'Bearer your-custom-token' }
});

// WebSocket Authentication  
const ws = new WebSocket('ws://your-server:3000?token=your-custom-token');
```

**Benefits:**
- 🔒 SharePoint credentials stay on server
- 🎛️ Control access with your own tokens
- 🔄 Rotate tokens without updating SharePoint
- 📊 Track usage by token/client

---

## �🐳 **Deployment Options**

### **Docker Deployment**
```bash
# Build and run with Docker Compose
docker-compose up -d

# Or build manually
docker build -t mcp-sharepoint .
docker run -d --env-file .env mcp-sharepoint
```

### **PM2 Process Manager**
```bash
# Install PM2 globally
npm install -g pm2

# Start both MCP and API servers
pm2 start ecosystem.config.json

# Or start individual services
pm2 start ecosystem.config.json --only mcp-sharepoint-api
pm2 start ecosystem.config.json --only mcp-sharepoint-mcp
```

### **Systemd Service**
```bash
# Copy service file
sudo cp mcp-sharepoint.service /etc/systemd/system/

# Enable and start service
sudo systemctl enable mcp-sharepoint
sudo systemctl start mcp-sharepoint
```

---

## 🔐 **Security Features**

### **✅ Enterprise-Grade Security**
- **Server-side Credentials** - SharePoint credentials never exposed to clients
- **Custom Authentication** - Generate your own API tokens
- **Certificate-based Auth** - X.509 certificate authentication with Azure
- **Token Rotation** - Easy token management and rotation
- **Access Control** - Multiple tokens for different access levels

### **✅ Deployment Security**
- **Environment Variables** - Sensitive data in environment, not code
- **Docker Secrets** - Container secret management
- **File Permissions** - Restricted certificate file access
- **Process Isolation** - Non-root user execution

---

## 📊 **Architecture Benefits**

### **🔄 Traditional MCP vs Enhanced API**

| Feature | Traditional MCP | Our Enhanced Server |
|---------|----------------|-------------------|
| **Protocol Compliance** | ✅ MCP Standard | ✅ MCP + HTTP + WebSocket |
| **Client Support** | 🔒 Single (Claude) | 🌐 Multiple simultaneous |
| **Authentication** | 😰 Credentials exposed | 🔐 Custom tokens only |
| **Deployment** | 🏠 Local only | ☁️ Server/Cloud ready |
| **Transport** | 📡 Stdio only | 🚀 Stdio + HTTP + WebSocket |
| **Security Model** | ⚠️ Client-side creds | 🛡️ Server-side secure |

### **🏗️ Technical Architecture**
```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   MCP Client    │◄──►│  Dual-Mode       │◄──►│   SharePoint    │
│ (Claude Desktop)│    │  Server          │    │   Graph API     │
└─────────────────┘    │                  │    └─────────────────┘
                       │  ┌─────────────┐ │
┌─────────────────┐    │  │   MCP Core  │ │
│  HTTP Client    │◄──►│  └─────────────┘ │
└─────────────────┘    │  ┌─────────────┐ │
                       │  │  API Server │ │
┌─────────────────┐    │  └─────────────┘ │
│ WebSocket Client│◄──►│  ┌─────────────┐ │
└─────────────────┘    │  │ Auth Layer  │ │
                       │  └─────────────┘ │
                       └──────────────────┘
```

---

## 🧪 **Testing & Validation**

### **Test the Installation**
```bash
# Test MCP mode
npm test

# Test API mode
npm run dev:api
curl http://localhost:3000/health
```

### **What's Been Tested**
- ✅ **Authentication** - Certificate-based Azure authentication working
- ✅ **SharePoint Access** - Successfully connects to SharePoint site
- ✅ **Folder Operations** - Lists folders and documents correctly
- ✅ **API Endpoints** - HTTP endpoints responding with proper data
- ✅ **Security** - Custom token authentication functioning
- ✅ **MCP Compliance** - Compatible with Claude Desktop integration

---

## 📈 **What We've Achieved**

### **🎯 From Python to JavaScript - Key Improvements**

| Aspect | Original Python Version | Our Enhanced JavaScript Version |
|--------|------------------------|--------------------------------|
| **Language** | 🐍 Python | ⚡ JavaScript/Node.js |
| **Authentication** | 🔑 Client Secret | 🔐 X.509 Certificate |
| **Architecture** | 📡 MCP Only | 🏗️ Dual-Mode (MCP + API) |
| **Client Support** | 🔒 Single Client | 🌐 Multi-Client |
| **Transport** | 📨 Stdio Only | 🚀 Stdio + HTTP + WebSocket |
| **Security Model** | ⚠️ Credentials in Config | 🛡️ Server-side + Custom Tokens |
| **Deployment** | 🏠 Local Development | ☁️ Production Ready |
| **Authentication Exposure** | 😰 Client-side Secrets | 🔐 Zero Credential Exposure |

### **🚀 Beyond the Original**
- ✅ **Complete Language Migration** - Full Python → JavaScript conversion
- ✅ **Enhanced Security** - Certificate-based auth + server-side credential protection  
- ✅ **Modern Architecture** - HTTP/WebSocket API server capabilities
- ✅ **Production Deployment** - Docker, PM2, systemd configurations
- ✅ **Enterprise Features** - Multi-client support, custom authentication
- ✅ **Developer Experience** - Multiple integration patterns and examples

### **🙏 Credits to Original Authors**
Special thanks to the [Sofias-ai team](https://github.com/Sofias-ai) for creating the foundational Python SharePoint MCP server that inspired this enhanced JavaScript implementation. Their work provided the blueprint for SharePoint integration patterns and MCP tool definitions that we've built upon and extended.

---

## 🔮 **Future Possibilities**

### **Planned Enhancements**
- 🔄 **Real-time Sync** - WebSocket-based file change notifications
- 🎯 **Advanced Permissions** - Token-based access control per folder/operation
- 📊 **Analytics Dashboard** - Usage monitoring and analytics
- 🔌 **Plugin System** - Extensible operation plugins
- 🌐 **Multi-Tenant** - Support multiple SharePoint sites per instance

### **Integration Opportunities**
- 🤖 **AI Assistants** - Beyond Claude (OpenAI, Anthropic, etc.)
- 🔗 **Webhooks** - SharePoint change notifications
- 📱 **Mobile SDKs** - Native mobile app integration
- 🌍 **GraphQL API** - Modern GraphQL endpoint
- 🔄 **Sync Services** - Bidirectional file synchronization

---

## 🤝 **Contributing**

We welcome contributions! This project demonstrates:
- Modern JavaScript/Node.js patterns
- MCP protocol implementation
- Enterprise security practices
- Multi-modal server architecture

---

## 📄 **License**

MIT License - see LICENSE file for details.

---

## 🎉 **Success Story**

**From Python to Production-Ready JavaScript MCP Server with Enterprise Security! 🚀**

This project successfully demonstrates how to build a next-generation MCP server that maintains full protocol compliance while adding modern API capabilities and enterprise-grade security features.

### Optional Configuration Variables
- `SHP_MAX_DEPTH`: Maximum folder depth for tree operations (default: 15)
- `SHP_MAX_FOLDERS_PER_LEVEL`: Maximum folders to process per level (default: 100)
- `SHP_LEVEL_DELAY`: Delay in seconds between processing levels (default: 0.5)

## Installation

### Prerequisites

- Node.js 18+ 
- npm or yarn
- A registered Azure AD application with SharePoint permissions
- A certificate (PFX format) for authentication

### Install from Source

```bash
# Clone the repository
git clone <repository-url>
cd sharePointMCP

# Install dependencies
npm install

# Build the TypeScript code
npm run build

# Set up environment variables
cp .env.example .env
# Edit .env with your configuration
```

## Usage

### Running the Server

```bash
# Development mode (with TypeScript compilation)
npm run dev

# Production mode (compiled JavaScript)
npm start
```

### Claude Desktop Integration

Add to your Claude Desktop configuration:

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": ["/path/to/sharePointMCP/dist/index.js"],
      "env": {
        "SHP_ID_APP": "your-app-id",
        "SHP_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site",
        "SHP_DOC_LIBRARY": "Shared Documents/your-folder",
        "SHP_TENANT_ID": "your-tenant-id",
        "SHP_CERT_PFX_PATH": "./certificate.pfx",
        "SHP_CERT_PFX_PASSWORD": "your-cert-password",
        "SHP_MAX_DEPTH": "15",
        "SHP_MAX_FOLDERS_PER_LEVEL": "100",
        "SHP_LEVEL_DELAY": "0.5"
      }
    }
  }
}
```

## Available Tools

### 📁 Folder Operations
- **List_SharePoint_Folders**: List folders in a directory
- **Get_SharePoint_Tree**: Get recursive folder structure
- **Create_Folder**: Create new folders
- **Delete_Folder**: Delete empty folders

### 📄 Document Operations
- **List_SharePoint_Documents**: List documents in a folder
- **Get_Document_Content**: Read document content (text/binary)
- **Upload_Document**: Upload files from content
- **Upload_Document_From_Path**: Upload files from local path
- **Update_Document**: Update existing documents
- **Delete_Document**: Delete documents
- **Download_Document**: Download files to local system

## Development

### Requirements

- Node.js 18+
- TypeScript 5+
- Dependencies listed in `package.json`

### Local Development

1. Clone the repository
2. Install dependencies:
   ```bash
   npm install
   ```
3. Create a `.env` file with your SharePoint credentials (use `.env.example` as template)
4. Build and run:
   ```bash
   npm run build
   npm start
   ```

## 🧪 **API Usage Examples**

### **Authentication**
All API endpoints require authentication. Use your API token:

```bash
# Set your API token
export API_TOKEN="your-64-character-api-token-here"
export SERVER_URL="https://your-server.onrender.com"
```

### **📖 READ Operations**

```bash
# List all folders
curl -H "Authorization: Bearer $API_TOKEN" "$SERVER_URL/api/folders"

# List documents in a specific folder
curl -H "Authorization: Bearer $API_TOKEN" "$SERVER_URL/api/documents?folderName=General"

# Get folder tree structure
curl -H "Authorization: Bearer $API_TOKEN" "$SERVER_URL/api/tree"

# Get document content
curl -H "Authorization: Bearer $API_TOKEN" "$SERVER_URL/api/document/test.txt/content"
```

### **📝 CREATE Operations**

```bash
# Upload a document (JSON-based)
curl -X POST \
  -H "Authorization: Bearer $API_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "fileName": "hello-world.txt",
    "content": "Hello from SharePoint MCP!",
    "folderPath": ""
  }' \
  "$SERVER_URL/api/upload"

# Create a new folder
curl -X POST \
  -H "Authorization: Bearer $API_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "folderName": "New Project",
    "parentPath": ""
  }' \
  "$SERVER_URL/api/folder"
```

### **✏️ UPDATE Operations**

```bash
# Update document content
curl -X PUT \
  -H "Authorization: Bearer $API_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "content": "Updated content for the document"
  }' \
  "$SERVER_URL/api/document/hello-world.txt"
```

### **🗑️ DELETE Operations**

```bash
# Delete a document
curl -X DELETE \
  -H "Authorization: Bearer $API_TOKEN" \
  "$SERVER_URL/api/document/hello-world.txt"

# Delete a folder
curl -X DELETE \
  -H "Authorization: Bearer $API_TOKEN" \
  "$SERVER_URL/api/folder/Old%20Project"

# Delete any item (universal endpoint)
curl -X DELETE \
  -H "Authorization: Bearer $API_TOKEN" \
  "$SERVER_URL/api/item/path/to/item"
```

### **🏥 Health Check**

```bash
# Check server status
curl "$SERVER_URL/health"

# Validate API token
curl -H "Authorization: Bearer $API_TOKEN" "$SERVER_URL/api/auth/validate"

# List all available endpoints
curl -H "Authorization: Bearer $API_TOKEN" "$SERVER_URL/api"
```

---

### Testing

```bash
# Run tests
npm test
```

### Debugging

For debugging the MCP server, you can use the [MCP Inspector](https://github.com/modelcontextprotocol/inspector):

```bash
npx @modelcontextprotocol/inspector -- node dist/index.js
```

## Security Considerations

- Store certificates securely and never commit them to version control
- Use environment variables for all sensitive configuration
- Regularly rotate certificates and access tokens
- Follow your organization's security policies for SharePoint access

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

MIT License. See LICENSE file for details.

## Support

For issues and questions:
1. Check the existing issues
2. Create a new issue with detailed information
3. Include logs and configuration (redact sensitive information)

## Differences from Python Version

This JavaScript/TypeScript version maintains feature parity with the original Python version while offering:

- **Better Node.js integration** for JavaScript-first environments
- **Type safety** with TypeScript
- **Modern async/await** patterns throughout
- **Improved error handling** with structured logging
- **NPM ecosystem** access for additional functionality
- **Faster startup times** in some environments