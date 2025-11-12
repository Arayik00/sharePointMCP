# 🤖 AI Agent Instructions: SharePoint MCP Server Usage

**Version**: 1.2.0  
**Target Audience**: AI agents, assistants, and automated systems that need to interact with SharePoint through this MCP server.  
**Last Updated**: November 12, 2025

---

## 📋 **Executive Summary**

### **What This SharePoint MCP Server Provides**
This is a **next-generation MCP server** that provides enterprise-grade SharePoint integration with both traditional MCP protocol support and modern API capabilities.

**Core Capabilities:**
- **🔐 Certificate-based Authentication** - Secure X.509 certificate authentication with Azure AD
- **📁 Complete CRUD Operations** - Create, Read, Update, Delete files and folders
- **🏗️ Dual-Mode Architecture** - Traditional MCP + HTTP/WebSocket API server
- **🛡️ Enterprise Security** - Server-side credential storage with custom API tokens
- **🌐 Multi-Client Support** - Multiple simultaneous client connections
- **🚀 Production Ready** - Docker, PM2, and systemd deployment configurations

### **Key Differentiators from Standard MCP Servers**
1. **Server-Side Credential Management** - SharePoint credentials stored securely on server
2. **Custom API Authentication** - Your own API tokens instead of exposing SharePoint credentials
3. **Multiple Transport Methods** - MCP stdio + HTTP REST + WebSocket
4. **Enterprise Deployment** - Production-ready with monitoring, logging, and security
5. **Comprehensive CRUD** - Full create, read, update, delete operations with advanced features

---

## 🏗️ **Architecture Overview**

### **System Components**

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   MCP Client    │◄──►│  SharePoint MCP  │◄──►│   SharePoint    │
│ (Claude Desktop)│    │  Dual-Mode       │    │   Graph API     │
├─────────────────┤    │  Server          │    └─────────────────┘
│  HTTP Client    │◄──►│                  │    
├─────────────────┤    │  ┌─────────────┐ │    ┌─────────────────┐
│ WebSocket Client│◄──►│  │   MCP Core  │ │◄──►│  Configuration  │
└─────────────────┘    │  ├─────────────┤ │    │     Loader      │
                       │  │ API Server  │ │    └─────────────────┘
                       │  ├─────────────┤ │
                       │  │ Auth Layer  │ │    ┌─────────────────┐
                       │  └─────────────┘ │◄──►│    Azure AD     │
                       └──────────────────┘    │ Certificate Auth│
                                              └─────────────────┘
```

### **Core Files and Their Functions**

| File | Purpose | Key Functionality |
|------|---------|-------------------|
| **src/config.js** | Configuration Management | Environment validation, credential loading, security checks |
| **src/server.js** | Main Entry Point | Server mode detection, initialization orchestration |
| **src/api-server.js** | HTTP/WebSocket API Server | RESTful endpoints, authentication middleware, CRUD operations |
| **src/sharepoint-server.js** | MCP Protocol Server | MCP tool registration, stdio transport handling |
| **src/sharepoint-client.js** | SharePoint Integration | Graph API client, certificate authentication, connection management |
| **src/sharepoint-tools.js** | Business Logic | SharePoint operations, data formatting, error handling |
| **src/mcp-api-client.js** | API Client Wrapper | MCP tools that proxy to API server instead of direct SharePoint |

### **Operating Modes**

1. **MCP Mode** (`SERVER_MODE=mcp`):
   - Traditional MCP protocol via stdio
   - Direct SharePoint Graph API connection
   - Tool-based interface for AI assistants

2. **API Mode** (`SERVER_MODE=api`):
   - HTTP REST API with authentication
   - WebSocket support for real-time operations
   - Custom token-based security

3. **Dual Mode** (`SERVER_MODE=dual`):
   - Both MCP and API simultaneously
   - Shared SharePoint connection
   - Maximum flexibility

---

## 🔧 **Integration Methods**

### **Method 1: Traditional MCP - Direct SharePoint Connection (Claude Desktop)**

**Use Case**: Claude Desktop integration, direct AI assistant access
**Security Model**: Credentials stored locally in MCP configuration
**Connection**: Direct to SharePoint Graph API via certificate authentication

**Configuration**:
```json
{
  "mcpServers": {
    "sharepoint-local": {
      "command": "node",
      "args": ["src/server.js", "mcp"],
      "env": {
        "SHP_ID_APP": "your-azure-app-id-here",
        "SHP_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site",
        "SHP_DOC_LIBRARY": "Shared Documents/E2E Automation",
        "SHP_TENANT_ID": "your-azure-tenant-id-here",
        "SHP_MAX_DEPTH": "15",
        "SHP_MAX_FOLDERS_PER_LEVEL": "100",
        "SHP_LEVEL_DELAY": "0.5",
        "SHP_CERT_PFX_PATH": "./your-certificate.pfx",
        "SHP_CERT_PFX_PASSWORD": "your-certificate-password"
      }
    }
  }
}
```

**Available MCP Tools**:
- `List_SharePoint_Folders` - List folders in SharePoint
- `List_SharePoint_Documents` - List documents in a folder
- `Get_SharePoint_Tree` - Get recursive folder structure
- `Get_Document_Content` - Read document content
- `Create_Folder` - Create new folders
- `Upload_Document` - Upload files from content
- `Upload_Document_From_Path` - Upload files from local paths
- `Update_Document` - Update existing documents
- `Delete_SharePoint_Item` - Delete files or folders

### **Method 2: MCP via API Server (Secure Enterprise)**

**Use Case**: Enterprise deployment with centralized credentials
**Security Model**: Server-side credentials + custom API tokens
**Connection**: MCP tools → API server → SharePoint

**Configuration**:
```json
{
  "mcpServers": {
    "sharepoint-api": {
      "command": "node",
      "args": ["src/mcp-api-client.js"],
      "env": {
        "SHAREPOINT_API_URL": "https://arayik-mcp-gdrive.onrender.com",
        "SHAREPOINT_API_TOKEN": "3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a"
      }
    }
  }
}
```

**Available MCP Tools** (Same interface, different backend):
- `List_SharePoint_Folders` - Proxied to API server
- `List_SharePoint_Documents` - Proxied to API server  
- `Get_SharePoint_Tree` - Proxied to API server
- `Get_SharePoint_Document_Content` - Proxied to API server
- `Upload_SharePoint_Document` - Proxied to API server
- `Validate_API_Connection` - Test API server connectivity

### **Method 3: HTTP REST API (Custom Applications)**

**Use Case**: Custom applications, web services, mobile apps
**Security Model**: Custom API tokens for authentication
**Connection**: HTTP REST → API server → SharePoint

**Server Configuration**:
```bash
# Environment variables for API server
SHP_ID_APP=your-azure-app-id
SHP_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site
SHP_DOC_LIBRARY=Shared Documents/E2E Automation
SHP_TENANT_ID=your-azure-tenant-id
SHP_CERT_PFX_PATH=./certificate.pfx
SHP_CERT_PFX_PASSWORD=your-certificate-password
SERVER_MODE=api
PORT=3000
API_AUTH_TOKENS=token1-32-chars-min,token2-32-chars-min,token3-32-chars-min
```

**Authentication Methods**:
```bash
# Method 1: Authorization header (recommended)
curl -H "Authorization: Bearer your-api-token" https://server.com/api/endpoint

# Method 2: X-API-Token header
curl -H "X-API-Token: your-api-token" https://server.com/api/endpoint

# Method 3: Query parameter
curl "https://server.com/api/endpoint?token=your-api-token"
```

**Live Server**: `https://arayik-mcp-gdrive.onrender.com`  
**API Token**: `3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a`

### **Method 4: WebSocket (Real-time Applications)**

**Use Case**: Real-time applications, live monitoring, streaming operations
**Security Model**: Token-based WebSocket authentication
**Connection**: WebSocket → API server → SharePoint

**Connection**:
```javascript
const ws = new WebSocket('wss://arayik-mcp-gdrive.onrender.com?token=3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a');
```

**Usage**:
```javascript
ws.send(JSON.stringify({
  action: 'listFolders',
  params: { parentFolder: 'Documents' }
}));

ws.onmessage = (event) => {
  const response = JSON.parse(event.data);
  console.log('SharePoint Response:', response);
};
```

---

## 🛠️ **Complete MCP Tool Reference**

### **Tool Availability by Mode**

| Tool Name | Direct MCP | API Client MCP | HTTP API | WebSocket |
|-----------|------------|----------------|----------|-----------|
| `List_SharePoint_Folders` | ✅ | ✅ | ✅ `/api/folders` | ✅ `listFolders` |
| `List_SharePoint_Documents` | ✅ | ✅ | ✅ `/api/documents` | ✅ `listDocuments` |
| `Get_SharePoint_Tree` | ✅ | ✅ | ✅ `/api/tree` | ✅ `getFolderTree` |
| `Get_Document_Content` | ✅ | - | ✅ `/api/document/:path/content` | ✅ `getDocumentContent` |
| `Get_SharePoint_Document_Content` | - | ✅ | ✅ `/api/document/:path/content` | ✅ `getDocumentContent` |
| `Create_Folder` | ✅ | - | ✅ `POST /api/folder` | - |
| `Upload_Document` | ✅ | - | ✅ `POST /api/upload` | - |
| `Upload_Document_From_Path` | ✅ | - | - | - |
| `Upload_SharePoint_Document` | - | ✅ | ✅ `POST /api/upload` | - |
| `Update_Document` | ✅ | - | ✅ `PUT /api/document/:path` | - |
| `Delete_SharePoint_Item` | ✅ | - | ✅ `DELETE /api/item/:path` | - |
| `Validate_API_Connection` | - | ✅ | ✅ `/api/auth/validate` | - |

### **1. List_SharePoint_Folders**
**Purpose**: List all folders in a specified directory
**Availability**: All modes
**Input Schema**:
```json
{
  "parent_folder": "string (optional)" // Path to parent folder, defaults to root
}
```
**Example**:
```json
{
  "name": "List_SharePoint_Folders",
  "arguments": {
    "parent_folder": "Team Hub/Documentation Central"
  }
}
```
**Returns**:
```json
{
  "success": true,
  "items": [
    {
      "name": "E2E Automation",
      "type": "folder",
      "size": 0,
      "webUrl": "https://hearstpm.sharepoint.com/sites/adproduct/Shared%20Documents/E2E%20Automation",
      "lastModified": "2025-11-11T07:50:40Z",
      "mimeType": ""
    }
  ],
  "count": 1
}

### **2. List_SharePoint_Documents**
**Purpose**: List all documents in a specified folder
**Availability**: All modes
**Input Schema**:
```json
{
  "folder_name": "string (required)" // Folder path to list documents from
}
```
**Example**:
```json
{
  "name": "List_SharePoint_Documents", 
  "arguments": {
    "folder_name": "Team Hub/Documentation Central"
  }
}
```
**Returns**:
```json
{
  "success": true,
  "items": [
    {
      "name": "test-upload.txt",
      "type": "file",
      "size": 370,
      "webUrl": "https://hearstpm.sharepoint.com/sites/adproduct/Shared%20Documents/test-upload.txt",
      "lastModified": "2025-11-10T11:50:05Z",
      "mimeType": "text/plain"
    }
  ],
  "count": 1
}

### **3. Get_SharePoint_Tree**
**Purpose**: Get recursive folder structure with nested hierarchy
**Availability**: All modes
**Input Schema**:
```json
{
  "parent_folder": "string (optional)", // Starting folder path
  "max_depth": "number (optional)"      // Maximum recursion depth (default: 3)
}
```
**Example**:
```json
{
  "name": "Get_SharePoint_Tree",
  "arguments": {
    "parent_folder": "Team Hub",
    "max_depth": 2
  }
}
```
**Returns**:
```json
{
  "success": true,
  "folder": "Team Hub",
  "tree": {
    "items": [
      {
        "name": "Documentation Central",
        "type": "folder",
        "size": 865978158,
        "webUrl": "https://hearstpm.sharepoint.com/.../Documentation%20Central",
        "lastModified": "2025-04-23T03:54:15Z",
        "children": {
          "items": [...],
          "truncated": false
        }
      }
    ],
    "truncated": false
  }
}

### **4. Get_Document_Content** (Direct MCP Mode)
**Purpose**: Read content of a specific document
**Availability**: Direct MCP mode only
**Input Schema**:
```json
{
  "folder_name": "string (required)",  // Folder containing the document
  "file_name": "string (required)"     // Name of the file to get content from
}
```
**Example**:
```json
{
  "name": "Get_Document_Content",
  "arguments": {
    "folder_name": "",
    "file_name": "test-upload.txt"
  }
}
```
**Returns**:
```json
{
  "success": true,
  "content": "Hello SharePoint! This is a test upload file...",
  "file": {
    "name": "test-upload.txt",
    "path": "test-upload.txt"
  },
  "type": "text"
}
```

### **4b. Get_SharePoint_Document_Content** (API Client Mode)
**Purpose**: Read content of a specific document via API server
**Availability**: API client MCP mode only
**Input Schema**:
```json
{
  "document_path": "string (required)" // Full path to the document
}
```
**Example**:
```json
{
  "name": "Get_SharePoint_Document_Content",
  "arguments": {
    "document_path": "test-upload.txt"
  }
}
```

### **5. Create_Folder** (Direct MCP Mode)
**Purpose**: Create a new folder in SharePoint
**Availability**: Direct MCP mode only
**Input Schema**:
```json
{
  "folder_name": "string (required)",   // Name of new folder
  "parent_folder": "string (optional)"  // Parent folder path
}
```
**Example**:
```json
{
  "name": "Create_Folder",
  "arguments": {
    "parent_folder": "Team Hub/Documentation Central",
    "folder_name": "E2E Automation"
  }
}
```
**Returns**:
```json
{
  "success": true,
  "message": "Folder E2E Automation created successfully"
}

### **6. Upload_Document** (Direct MCP Mode)
**Purpose**: Upload a document with content to SharePoint
**Availability**: Direct MCP mode only
**Input Schema**:
```json
{
  "folder_name": "string (required)",    // Target folder name
  "file_name": "string (required)",      // Name of the file
  "content": "string (required)",        // File content
  "is_base64": "boolean (optional)"      // Whether content is base64 encoded
}
```
**Example**:
```json
{
  "name": "Upload_Document",
  "arguments": {
    "folder_name": "",
    "file_name": "test-upload.txt",
    "content": "Hello SharePoint! This is a test upload file.\nCreated on: November 10, 2025\nPurpose: Testing the upload functionality.",
    "is_base64": false
  }
}
```
**Returns**:
```json
{
  "success": true,
  "message": "Document test-upload.txt uploaded successfully"
}
```

### **6b. Upload_SharePoint_Document** (API Client Mode)
**Purpose**: Upload a document via API server
**Availability**: API client MCP mode only
**Input Schema**:
```json
{
  "file_name": "string (required)",
  "content": "string (required)",
  "folder_path": "string (optional)",
  "overwrite": "boolean (optional)"
}
```
**Example**:
```json
{
  "name": "Upload_SharePoint_Document",
  "arguments": {
    "file_name": "api-test.txt",
    "content": "Uploaded via API server",
    "folder_path": "",
    "overwrite": false
  }
}
```

### **7. Upload_Document_From_Path** (Direct MCP Mode Only)
**Purpose**: Upload a document from local file path
**Availability**: Direct MCP mode only
**Input Schema**:
```json
{
  "folder_name": "string (required)",
  "file_path": "string (required)",     // Local file path
  "file_name": "string (optional)"      // Target name (defaults to source filename)
}
```

### **8. Update_Document** (Direct MCP Mode)
**Purpose**: Update content of existing document
**Availability**: Direct MCP mode only
**Input Schema**:
```json
{
  "folder_name": "string (required)",   // Folder containing the document
  "file_name": "string (required)",     // Name of the file to update
  "content": "string (required)",       // New content
  "is_base64": "boolean (optional)"     // Whether content is base64 encoded
}
```
**Example**:
```json
{
  "name": "Update_Document",
  "arguments": {
    "folder_name": "",
    "file_name": "test-upload.txt",
    "content": "Updated content with additional information.\nTimestamp: November 11, 2025",
    "is_base64": false
  }
}
```

### **9. Delete_SharePoint_Item** (Direct MCP Mode)
**Purpose**: Delete a file or folder from SharePoint
**Availability**: Direct MCP mode only
**Input Schema**:
```json
{
  "item_path": "string (required)" // Full path to item to delete
}
```
**Example**:
```json
{
  "name": "Delete_SharePoint_Item",
  "arguments": {
    "item_path": "test-upload.txt"
  }
}
```
**Returns**:
```json
{
  "success": true,
  "message": "Item test-upload.txt deleted successfully"
}
```

### **10. Validate_API_Connection** (API Client Mode)
**Purpose**: Test API server connection and authentication
**Availability**: API client MCP mode only
**Input Schema**: `{}` (no parameters)
**Example**:
```json
{
  "name": "Validate_API_Connection",
  "arguments": {}
}
```
**Returns**:
```json
{
  "valid": true,
  "message": "API connection valid",
  "authenticated": true,
  "token_preview": "3a0180d8...",
  "timestamp": "2025-11-11T08:00:00Z"
}

---

## 🌐 **Complete HTTP REST API Reference**

### **Live Production Server**
**Base URL**: `https://arayik-mcp-gdrive.onrender.com`  
**API Token**: `3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a`  
**Status**: ✅ Operational with 11 CRUD endpoints

### **📖 READ Operations**

#### `GET /health` (No Authentication Required)
**Purpose**: Server health check and endpoint discovery
```bash
curl https://arayik-mcp-gdrive.onrender.com/health
```
**Response**:
```json
{
  "status": "healthy",
  "timestamp": "2025-11-11T07:54:55.108Z",
  "sharepoint": "connected",
  "authentication": "required",
  "endpoints": [
    "GET /api/auth/validate - Validate API token",
    "GET /api/folders - List SharePoint folders",
    "GET /api/documents - List SharePoint documents",
    "GET /api/tree - Get folder tree structure",
    "GET /api/document/:path/content - Get document content",
    "POST /api/upload - Upload document",
    "POST /api/folder - Create new folder",
    "PUT /api/document/:path - Update document content",
    "DELETE /api/item/:path - Delete file or folder",
    "DELETE /api/document/:path - Delete document (alias)",
    "DELETE /api/folder/:path - Delete folder (alias)"
  ]
}
```

#### `GET /api/folders` 
**Purpose**: List folders in SharePoint
**Parameters**: `?parentFolder=path` (optional)
```bash
curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     "https://arayik-mcp-gdrive.onrender.com/api/folders"
```

#### `GET /api/documents`
**Purpose**: List documents in a folder  
**Parameters**: `?folderName=path` (optional)
```bash
curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     "https://arayik-mcp-gdrive.onrender.com/api/documents"
```

#### `GET /api/tree`
**Purpose**: Get recursive folder structure
**Parameters**: `?folderPath=path&maxDepth=N` (optional)
```bash
curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     "https://arayik-mcp-gdrive.onrender.com/api/tree?maxDepth=2"
```

#### `GET /api/document/:path/content`
**Purpose**: Get document content
**Parameters**: `:path` - URL-encoded document path
```bash
curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     "https://arayik-mcp-gdrive.onrender.com/api/document/test-file.txt/content"
```

### **📝 CREATE Operations**

#### `POST /api/upload`
**Purpose**: Upload a document
**Content-Type**: `application/json`
```bash
curl -X POST \
  -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
  -H "Content-Type: application/json" \
  -d '{
    "fileName": "hello-world.txt",
    "content": "Hello from SharePoint MCP!",
    "path": "hello-world.txt"
  }' \
  https://arayik-mcp-gdrive.onrender.com/api/upload
```

#### `POST /api/folder`
**Purpose**: Create a new folder
**Content-Type**: `application/json`
```bash
curl -X POST \
  -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
  -H "Content-Type: application/json" \
  -d '{
    "folderName": "New Project",
    "path": "Team Hub/Documentation Central"
  }' \
  https://arayik-mcp-gdrive.onrender.com/api/folder
```

### **✏️ UPDATE Operations**

#### `PUT /api/document/:path`
**Purpose**: Update document content
**Content-Type**: `application/json`
```bash
curl -X PUT \
  -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
  -H "Content-Type: application/json" \
  -d '{
    "content": "Updated content for the document"
  }' \
  https://arayik-mcp-gdrive.onrender.com/api/document/hello-world.txt
```

### **🗑️ DELETE Operations**

#### `DELETE /api/item/:path` (Universal Delete)
**Purpose**: Delete any item (file or folder)
```bash
curl -X DELETE \
  -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
  https://arayik-mcp-gdrive.onrender.com/api/item/hello-world.txt
```

#### `DELETE /api/document/:path` (Alias)
```bash
curl -X DELETE \
  -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
  https://arayik-mcp-gdrive.onrender.com/api/document/hello-world.txt
```

#### `DELETE /api/folder/:path` (Alias)
```bash
curl -X DELETE \
  -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
  https://arayik-mcp-gdrive.onrender.com/api/folder/Old%20Project
```

### **🔐 Authentication & Utility**

#### `GET /api/auth/validate`
**Purpose**: Validate API token
```bash
curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     https://arayik-mcp-gdrive.onrender.com/api/auth/validate
```

#### `GET /api`
**Purpose**: List all available endpoints
```bash
curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     https://arayik-mcp-gdrive.onrender.com/api
```

### **Standard Response Formats**

**Success Response**:
```json
{
  "success": true,
  "items": [...],
  "count": 5,
  "message": "Operation completed successfully"
}
```

**Error Response**:
```json
{
  "error": "Unauthorized",
  "message": "Invalid API token provided"
}
```

---

## � **Environment Configuration & Server Management**

### **Environment Variables Reference**

#### **SharePoint Authentication (Required)**
```bash
SHP_ID_APP=12345678-1234-1234-1234-123456789012          # Azure AD Application ID (GUID)
SHP_TENANT_ID=87654321-4321-4321-4321-210987654321       # Azure AD Tenant ID (GUID)  
SHP_SITE_URL=https://tenant.sharepoint.com/sites/site    # SharePoint Site URL
SHP_CERT_PFX_PATH=/path/to/certificate.pfx               # X.509 Certificate Path
SHP_CERT_PFX_PASSWORD=your-certificate-password          # Certificate Password
```

#### **Server Configuration (Required for API Mode)**
```bash
SERVER_MODE=api                                           # Mode: mcp, api, or dual
PORT=3000                                                # Server port for API mode
HOST=localhost                                           # Server host binding
LOG_LEVEL=info                                           # Logging: error, warn, info, debug
```

#### **API Security (Required for API Mode)**  
```bash
API_AUTH_TOKENS=token1-min-32-chars,token2-min-32-chars  # Custom authentication tokens
CORS_ORIGINS=*                                           # CORS origins (* for all)
RATE_LIMIT_WINDOW=900000                                 # Rate limit window (15 min in ms)
RATE_LIMIT_MAX=100                                       # Max requests per window
```

#### **SharePoint Scope (Optional)**
```bash
SHP_DOC_LIBRARY=Shared Documents                         # Document library path
SHP_MAX_DEPTH=15                                         # Maximum folder tree depth
SHP_MAX_FOLDERS_PER_LEVEL=100                           # Max folders per level
SHP_LEVEL_DELAY=0.5                                     # Delay between levels (seconds)
```

### **Server Startup Commands**

#### **MCP Mode (Direct SharePoint)**
```bash
cd /path/to/sharePointMCP
export SHP_ID_APP="your-app-id"
export SHP_SITE_URL="your-sharepoint-url" 
export SHP_TENANT_ID="your-tenant-id"
export SHP_CERT_PFX_PATH="./certificate.pfx"
export SHP_CERT_PFX_PASSWORD="your-password"
npm run start:mcp
```

#### **API Mode (HTTP Server)**
```bash
cd /path/to/sharePointMCP  
export SERVER_MODE=api
export PORT=3000
export API_AUTH_TOKENS="token1,token2"
# ... (SharePoint env vars)
npm run start:api
```

#### **Dual Mode (Both MCP + API)**
```bash
cd /path/to/sharePointMCP
export SERVER_MODE=dual
export PORT=3000
export API_AUTH_TOKENS="token1,token2"
# ... (SharePoint env vars) 
npm start
```

### **Production Deployment Options**

#### **Docker Deployment**
```bash
# Using docker-compose
docker-compose up -d

# Manual Docker build
docker build -t sharepoint-mcp .
docker run -d --env-file .env sharepoint-mcp
```

#### **PM2 Process Manager**
```bash
# Install PM2
npm install -g pm2

# Start with ecosystem config
pm2 start ecosystem.config.json

# Start individual services
pm2 start ecosystem.config.json --only sharepoint-api
pm2 start ecosystem.config.json --only sharepoint-mcp
```

#### **Systemd Service (Linux)**
```bash
# Copy service file
sudo cp mcp-sharepoint.service /etc/systemd/system/

# Enable and start
sudo systemctl enable mcp-sharepoint
sudo systemctl start mcp-sharepoint
sudo systemctl status mcp-sharepoint
```

### **Configuration Validation**
```bash
# Validate configuration and credentials
npm run validate

# Test API connection
npm run test:api

# Full system test
npm run test:full
```

---

## 🔍 **Comprehensive Usage Patterns & Examples**

### **Pattern 1: SharePoint Structure Discovery & Navigation**

#### **A. MCP Tool Approach (Direct SharePoint)**
```javascript
// 1. List root folders to understand structure
const folders = await callTool("List_SharePoint_Folders", {});
console.log("Root folders:", folders.items.map(f => f.name));

// 2. Get comprehensive tree view for navigation
const tree = await callTool("Get_SharePoint_Tree", {
  "parent_folder": "Team Hub",
  "max_depth": 3
});
console.log("Team Hub structure:", tree.tree);

// 3. Explore specific areas of interest
const docs = await callTool("List_SharePoint_Documents", {
  "folder_name": "Team Hub/Documentation Central"
});
console.log("Documentation files:", docs.items);
```

#### **B. HTTP API Approach**
```bash
# 1. Health check and endpoint discovery
curl https://arayik-mcp-gdrive.onrender.com/health

# 2. Authentication test
curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     https://arayik-mcp-gdrive.onrender.com/api/auth/validate

# 3. Explore folder structure
curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     "https://arayik-mcp-gdrive.onrender.com/api/tree?maxDepth=2"

# 4. List specific folder contents  
curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     "https://arayik-mcp-gdrive.onrender.com/api/documents?folderName=Team%20Hub"
```

#### **C. JavaScript/Node.js Integration**
```javascript
import axios from 'axios';

class SharePointMCPClient {
  constructor(baseURL, apiToken) {
    this.client = axios.create({
      baseURL,
      headers: { 'Authorization': `Bearer ${apiToken}` }
    });
  }

  async exploreStructure() {
    // Get health status
    const health = await axios.get(`${this.client.defaults.baseURL}/health`);
    console.log('Server status:', health.data.status);

    // Get all folders
    const folders = await this.client.get('/api/folders');
    console.log('Available folders:', folders.data.items);

    // Get tree structure
    const tree = await this.client.get('/api/tree?maxDepth=2');
    return tree.data;
  }
}

// Usage
const client = new SharePointMCPClient(
  'https://arayik-mcp-gdrive.onrender.com',
  '3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a'
);
const structure = await client.exploreStructure();
```

### **Pattern 2: Document Analysis & Content Processing**

#### **A. MCP Tool Batch Processing**
```javascript
// 1. Discover documents for analysis
const folders = await callTool("List_SharePoint_Folders", {});
const analysisTargets = folders.items.filter(f => 
  f.name.includes('Reports') || f.name.includes('Data')
);

// 2. Process each folder
for (const folder of analysisTargets) {
  const documents = await callTool("List_SharePoint_Documents", {
    "folder_name": folder.name
  });
  
  console.log(`Found ${documents.count} documents in ${folder.name}`);
  
  // 3. Analyze text documents
  for (const doc of documents.items) {
    if (doc.mimeType === 'text/plain' || doc.name.endsWith('.txt')) {
      const content = await callTool("Get_Document_Content", {
        "folder_name": folder.name,
        "file_name": doc.name
      });
      
      // Process content
      const analysis = analyzeTextContent(content.content);
      console.log(`Analysis of ${doc.name}:`, analysis);
    }
  }
}

function analyzeTextContent(content) {
  return {
    wordCount: content.split(/\s+/).length,
    lines: content.split('\n').length,
    keywords: extractKeywords(content),
    sentiment: analyzeSentiment(content)
  };
}
```

#### **B. HTTP API Streaming Analysis**
```javascript
async function analyzeSharePointDocuments() {
  const client = axios.create({
    baseURL: 'https://arayik-mcp-gdrive.onrender.com',
    headers: { 'Authorization': 'Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a' }
  });

  // 1. Get all documents
  const docs = await client.get('/api/documents');
  
  // 2. Filter and process text files
  const textFiles = docs.data.items.filter(doc => 
    doc.mimeType === 'text/plain' || 
    doc.name.endsWith('.txt') || 
    doc.name.endsWith('.md')
  );

  const results = [];
  
  // 3. Process each document
  for (const doc of textFiles) {
    try {
      const content = await client.get(`/api/document/${encodeURIComponent(doc.name)}/content`);
      
      if (content.data.success) {
        const analysis = {
          filename: doc.name,
          size: doc.size,
          lastModified: doc.lastModified,
          contentAnalysis: analyzeContent(content.data.content),
          webUrl: doc.webUrl
        };
        
        results.push(analysis);
        console.log(`✅ Analyzed: ${doc.name}`);
      }
    } catch (error) {
      console.error(`❌ Error analyzing ${doc.name}:`, error.message);
    }
  }
  
  return results;
}
```

### **Pattern 3: Content Creation & Organization**

#### **A. MCP Tool Content Generation Workflow**
```javascript
async function createAIGeneratedReports() {
  // 1. Create organized folder structure
  await callTool("Create_Folder", {
    "parent_folder": "",
    "folder_name": "AI Generated Content"
  });
  
  await callTool("Create_Folder", {
    "parent_folder": "AI Generated Content", 
    "folder_name": "Daily Reports"
  });

  // 2. Generate and upload daily report
  const reportDate = new Date().toISOString().split('T')[0];
  const reportContent = generateDailyReport();
  
  await callTool("Upload_Document", {
    "folder_name": "AI Generated Content/Daily Reports",
    "file_name": `daily-report-${reportDate}.txt`,
    "content": reportContent,
    "is_base64": false
  });

  // 3. Create summary index
  const summaryContent = createReportSummary(reportDate);
  
  await callTool("Upload_Document", {
    "folder_name": "AI Generated Content",
    "file_name": "reports-index.md",
    "content": summaryContent,
    "is_base64": false
  });

  console.log(`✅ Created report structure and uploaded daily report for ${reportDate}`);
}

function generateDailyReport() {
  const timestamp = new Date().toISOString();
  return `# Daily AI Analysis Report
Generated: ${timestamp}

## Summary
This report contains automated analysis of SharePoint activity.

## Key Metrics
- Documents processed: 125
- Folders analyzed: 23  
- Content insights: 15 new patterns identified

## Recommendations
1. Archive documents older than 90 days
2. Consolidate duplicate content in Team Hub
3. Update folder permissions for E2E Automation

## Next Actions
- Schedule monthly cleanup
- Implement automated tagging
- Setup content lifecycle policies
`;
}
```

#### **B. HTTP API Batch Upload System**
```javascript
class SharePointContentManager {
  constructor(baseURL, apiToken) {
    this.client = axios.create({
      baseURL,
      headers: { 
        'Authorization': `Bearer ${apiToken}`,
        'Content-Type': 'application/json'
      }
    });
  }

  async createProjectStructure(projectName) {
    // 1. Create main project folder
    await this.client.post('/api/folder', {
      folderName: projectName,
      path: ''
    });

    // 2. Create sub-folders
    const subFolders = ['Documentation', 'Reports', 'Resources', 'Archive'];
    
    for (const folder of subFolders) {
      await this.client.post('/api/folder', {
        folderName: folder,
        path: projectName
      });
      console.log(`✅ Created: ${projectName}/${folder}`);
    }

    // 3. Upload initial files
    const initialFiles = this.generateInitialFiles(projectName);
    
    for (const file of initialFiles) {
      await this.client.post('/api/upload', {
        fileName: file.name,
        content: file.content,
        path: `${projectName}/${file.folder}/${file.name}`
      });
      console.log(`✅ Uploaded: ${file.name}`);
    }

    return { success: true, project: projectName, folders: subFolders.length, files: initialFiles.length };
  }

  generateInitialFiles(projectName) {
    return [
      {
        name: 'README.md',
        folder: 'Documentation',
        content: `# ${projectName}

## Overview
This project was automatically created by the SharePoint MCP Server.

## Structure
- Documentation/ - Project documentation
- Reports/ - Analysis and status reports  
- Resources/ - Shared resources and assets
- Archive/ - Archived materials

## Getting Started
1. Review the project objectives
2. Update this README with specific details
3. Begin adding content to appropriate folders

Created: ${new Date().toISOString()}
`
      },
      {
        name: 'project-status.txt',
        folder: 'Reports',
        content: `Project Status Report: ${projectName}
Generated: ${new Date().toISOString()}

Status: INITIALIZED
Progress: 0%
Next Milestone: Project kickoff

Team Members: TBD
Budget: TBD
Timeline: TBD

Recent Activity:
- ${new Date().toDateString()}: Project structure created
- Folders initialized: Documentation, Reports, Resources, Archive
`
      }
    ];
  }
}

// Usage
const contentManager = new SharePointContentManager(
  'https://arayik-mcp-gdrive.onrender.com',
  '3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a'
);

await contentManager.createProjectStructure('Q1-2025-Initiative');
```

### **Pattern 4: Document Lifecycle Management**

#### **A. Update and Version Control**
```javascript
async function updateDocumentWithVersioning() {
  const documentPath = 'team-handbook.txt';
  
  // 1. Read current content
  const current = await callTool("Get_Document_Content", {
    "folder_name": "",
    "file_name": documentPath
  });

  if (current.success) {
    // 2. Create backup with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupName = `team-handbook-backup-${timestamp}.txt`;
    
    await callTool("Upload_Document", {
      "folder_name": "Archive",
      "file_name": backupName,
      "content": current.content,
      "is_base64": false
    });

    // 3. Update original with new content
    const updatedContent = `${current.content}\n\n--- Update ${new Date().toISOString()} ---\nNew section added by AI assistant.`;
    
    await callTool("Update_Document", {
      "folder_name": "",
      "file_name": documentPath,
      "content": updatedContent,
      "is_base64": false
    });

    console.log(`✅ Document updated with backup created: ${backupName}`);
  }
}
```

#### **B. Cleanup and Organization**
```bash
# HTTP API cleanup workflow
TOKEN="3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a"
BASE_URL="https://arayik-mcp-gdrive.onrender.com"

# 1. List all documents to identify cleanup targets
curl -H "Authorization: Bearer $TOKEN" "$BASE_URL/api/documents" | jq '.items[] | select(.name | startswith("temp-") or startswith("test-"))'

# 2. Delete temporary files (example)
curl -X DELETE -H "Authorization: Bearer $TOKEN" "$BASE_URL/api/document/temp-file.txt"

# 3. Create cleanup report
curl -X POST -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
  -d '{
    "fileName": "cleanup-report.txt",
    "content": "Cleanup Report - '$(date)'\n\nTemporary files removed:\n- temp-file.txt\n- test-upload.txt\n\nOrganization status: Complete",
    "path": "cleanup-report.txt"
  }' \
  "$BASE_URL/api/upload"
```

---

## ⚠️ **Critical Implementation Guidelines for AI Agents**

### **🚨 Critical Error Handling Patterns**

#### **Authentication & Authorization**
```javascript
// Always validate authentication first
try {
  if (usingAPIMode) {
    const authCheck = await callTool("Validate_API_Connection", {});
    if (!authCheck.valid) {
      throw new Error('API authentication failed - check token');
    }
  }
  
  // Proceed with operations
  const folders = await callTool("List_SharePoint_Folders", {});
  
} catch (error) {
  if (error.message.includes('Unauthorized') || error.message.includes('401')) {
    console.error('🚫 Authentication failed - verify API token or certificate configuration');
    return { error: 'AUTHENTICATION_FAILED', details: error.message };
  }
  
  if (error.message.includes('ENOTFOUND') || error.message.includes('connection')) {
    console.error('🌐 Network connectivity issue - check server URL and network');
    return { error: 'NETWORK_ERROR', details: error.message };
  }
  
  // Generic error handling
  console.error('❌ SharePoint operation failed:', error.message);
  return { error: 'OPERATION_FAILED', details: error.message };
}
```

#### **Robust Response Validation**
```javascript
function validateSharePointResponse(response, operation) {
  // Check for success field
  if (response && typeof response.success !== 'undefined') {
    if (!response.success) {
      throw new Error(`SharePoint ${operation} failed: ${response.message || 'Unknown error'}`);
    }
    return response;
  }
  
  // Handle different response formats
  if (response && response.items) {
    return { success: true, ...response };
  }
  
  // Handle text responses (direct content)
  if (typeof response === 'string') {
    return { success: true, content: response };
  }
  
  throw new Error(`Invalid response format from SharePoint ${operation}`);
}
```

### **🔒 Security & Authentication Best Practices**

#### **Token Security**
```javascript
// ✅ CORRECT: Secure token handling
const API_TOKEN = process.env.SHAREPOINT_API_TOKEN; // From environment
if (!API_TOKEN || API_TOKEN.length < 32) {
  throw new Error('Invalid or missing API token - must be 32+ characters');
}

// ❌ INCORRECT: Never hardcode tokens
const API_TOKEN = "3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a";
```

#### **Path Validation** 
```javascript
function validateSharePointPath(path) {
  // Prevent directory traversal
  if (path.includes('..') || path.includes('//') || path.startsWith('/')) {
    throw new Error('Invalid path: directory traversal detected');
  }
  
  // Validate characters
  const invalidChars = /[<>:"|?*]/;
  if (invalidChars.test(path)) {
    throw new Error('Invalid path: contains forbidden characters');
  }
  
  return path.trim();
}
```

### **⚡ Performance & Rate Limiting**

#### **Intelligent Batching**
```javascript
async function processDocumentsBatch(documents, batchSize = 5) {
  const results = [];
  
  for (let i = 0; i < documents.length; i += batchSize) {
    const batch = documents.slice(i, i + batchSize);
    
    // Process batch with delay
    const batchPromises = batch.map(async (doc, index) => {
      // Stagger requests within batch
      await new Promise(resolve => setTimeout(resolve, index * 200));
      
      return await callTool("Get_Document_Content", {
        "folder_name": doc.folder,
        "file_name": doc.name
      });
    });
    
    const batchResults = await Promise.all(batchPromises);
    results.push(...batchResults);
    
    // Inter-batch delay
    if (i + batchSize < documents.length) {
      await new Promise(resolve => setTimeout(resolve, 1000));
    }
  }
  
  return results;
}
```

#### **Caching Strategy**
```javascript
class SharePointCache {
  constructor(ttl = 300000) { // 5 minutes default
    this.cache = new Map();
    this.ttl = ttl;
  }
  
  get(key) {
    const item = this.cache.get(key);
    if (item && Date.now() < item.expiry) {
      return item.data;
    }
    this.cache.delete(key);
    return null;
  }
  
  set(key, data) {
    this.cache.set(key, {
      data,
      expiry: Date.now() + this.ttl
    });
  }
  
  invalidate(pattern) {
    for (const key of this.cache.keys()) {
      if (key.includes(pattern)) {
        this.cache.delete(key);
      }
    }
  }
}

// Usage
const cache = new SharePointCache();
const cacheKey = `folders-${parentFolder || 'root'}`;
let folders = cache.get(cacheKey);

if (!folders) {
  folders = await callTool("List_SharePoint_Folders", { parent_folder: parentFolder });
  cache.set(cacheKey, folders);
}
```

### **📁 File Type & Content Handling**

#### **Content Type Detection**
```javascript
function detectContentType(filename, content) {
  const ext = filename.split('.').pop().toLowerCase();
  
  // Text files - return as string
  if (['txt', 'md', 'json', 'xml', 'csv', 'log'].includes(ext)) {
    return { type: 'text', encoding: 'utf8', content };
  }
  
  // Check if content looks like base64
  if (typeof content === 'string' && /^[A-Za-z0-9+/]+=*$/.test(content.replace(/\s/g, ''))) {
    return { type: 'binary', encoding: 'base64', content };
  }
  
  // Office documents
  if (['docx', 'xlsx', 'pptx', 'pdf'].includes(ext)) {
    return { type: 'office', encoding: 'base64', content };
  }
  
  return { type: 'unknown', encoding: 'utf8', content };
}
```

#### **Size Limit Checking**
```javascript
function validateFileSize(content, maxSizeBytes = 50 * 1024 * 1024) { // 50MB default
  let sizeBytes;
  
  if (typeof content === 'string') {
    // For base64, actual size is ~75% of string length
    sizeBytes = content.includes('base64') 
      ? Math.floor(content.length * 0.75) 
      : Buffer.byteLength(content, 'utf8');
  } else {
    sizeBytes = content.length || 0;
  }
  
  if (sizeBytes > maxSizeBytes) {
    throw new Error(`File size ${sizeBytes} bytes exceeds limit of ${maxSizeBytes} bytes`);
  }
  
  return { sizeBytes, valid: true };
}

---

## 🚀 **Production Ready Quick Start Guide**

### **🎯 Complete Working Examples**

#### **Enterprise MCP Implementation**
```javascript
class SharePointMCPManager {
  constructor(mode = 'api') {
    this.mode = mode;
    this.cache = new Map();
  }

  async initialize() {
    // Validate connection based on mode
    if (this.mode === 'api') {
      const validation = await callTool("Validate_API_Connection", {});
      if (!validation.valid) {
        throw new Error('API connection validation failed');
      }
      console.log('✅ API connection validated');
    }
    
    // Test basic functionality
    const folders = await callTool("List_SharePoint_Folders", {});
    console.log(`✅ Connected to SharePoint - ${folders.count} folders available`);
    
    return { connected: true, folders: folders.count };
  }

  async performHealthyOperation(operation, params, retries = 3) {
    for (let attempt = 1; attempt <= retries; attempt++) {
      try {
        const result = await callTool(operation, params);
        
        if (result && result.success !== false) {
          return result;
        }
        
        throw new Error(result.message || 'Operation returned failure');
        
      } catch (error) {
        console.warn(`⚠️ Attempt ${attempt}/${retries} failed:`, error.message);
        
        if (attempt === retries) {
          throw new Error(`Operation ${operation} failed after ${retries} attempts: ${error.message}`);
        }
        
        // Exponential backoff
        await new Promise(resolve => setTimeout(resolve, Math.pow(2, attempt) * 1000));
      }
    }
  }

  async exploreAndAnalyze() {
    // 1. Comprehensive structure discovery
    const tree = await this.performHealthyOperation("Get_SharePoint_Tree", {
      "max_depth": 2
    });
    
    console.log('📁 SharePoint Structure:', tree.tree);
    
    // 2. Document inventory
    const allFolders = this.extractFoldersFromTree(tree.tree);
    const documentInventory = {};
    
    for (const folder of allFolders) {
      const docs = await this.performHealthyOperation("List_SharePoint_Documents", {
        "folder_name": folder
      });
      
      documentInventory[folder] = {
        count: docs.count,
        totalSize: docs.items.reduce((sum, doc) => sum + (doc.size || 0), 0),
        fileTypes: [...new Set(docs.items.map(doc => doc.mimeType))].filter(Boolean)
      };
    }
    
    return { structure: tree, inventory: documentInventory };
  }

  extractFoldersFromTree(tree, basePath = '') {
    const folders = [];
    
    for (const item of tree.items || []) {
      if (item.type === 'folder') {
        const fullPath = basePath ? `${basePath}/${item.name}` : item.name;
        folders.push(fullPath);
        
        if (item.children && item.children.items) {
          folders.push(...this.extractFoldersFromTree(item.children, fullPath));
        }
      }
    }
    
    return folders;
  }
}

// Usage
const manager = new SharePointMCPManager('api');
await manager.initialize();
const analysis = await manager.exploreAndAnalyze();
console.log('📊 Complete Analysis:', analysis);
```

#### **Production HTTP API Client**
```javascript
import axios from 'axios';

class ProductionSharePointClient {
  constructor() {
    this.baseURL = 'https://arayik-mcp-gdrive.onrender.com';
    this.apiToken = '3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a';
    
    this.client = axios.create({
      baseURL: this.baseURL,
      headers: { 'Authorization': `Bearer ${this.apiToken}` },
      timeout: 30000
    });
    
    // Add response interceptor for error handling
    this.client.interceptors.response.use(
      response => response,
      error => {
        console.error('SharePoint API Error:', error.response?.data || error.message);
        throw error;
      }
    );
  }

  async healthCheck() {
    const health = await axios.get(`${this.baseURL}/health`);
    console.log('🏥 Server Health:', health.data);
    
    const auth = await this.client.get('/api/auth/validate');
    console.log('🔐 Authentication:', auth.data);
    
    return { healthy: health.data.status === 'healthy', authenticated: auth.data.valid };
  }

  async fullWorkflowDemo() {
    console.log('🚀 Starting SharePoint MCP Full Workflow Demo');
    
    // 1. Health check
    const status = await this.healthCheck();
    if (!status.healthy || !status.authenticated) {
      throw new Error('System not ready');
    }

    // 2. Explore structure
    const folders = await this.client.get('/api/folders');
    console.log('📁 Available Folders:', folders.data.count);

    // 3. Create demo content
    const demoFile = `SharePoint MCP Demo
Created: ${new Date().toISOString()}
Server: ${this.baseURL}
Status: ✅ Operational

This file demonstrates the complete workflow:
1. Health check ✓
2. Authentication ✓  
3. Folder listing ✓
4. File creation ✓
5. Content update ✓
6. Cleanup ✓

Demo completed successfully!
`;

    // 4. Upload demo file
    const upload = await this.client.post('/api/upload', {
      fileName: 'mcp-demo.txt',
      content: demoFile,
      path: 'mcp-demo.txt'
    });
    console.log('📤 Upload Result:', upload.data);

    // 5. Read back content
    const content = await this.client.get('/api/document/mcp-demo.txt/content');
    console.log('📖 Retrieved Content Length:', content.data.content?.length || 0);

    // 6. Update content
    const updatedContent = demoFile + `\nUpdated: ${new Date().toISOString()}\nDemo workflow completed!`;
    const update = await this.client.put('/api/document/mcp-demo.txt', {
      content: updatedContent
    });
    console.log('✏️ Update Result:', update.data);

    // 7. Clean up
    const cleanup = await this.client.delete('/api/document/mcp-demo.txt');
    console.log('🧹 Cleanup Result:', cleanup.data);

    return { 
      success: true, 
      operations: ['health', 'auth', 'list', 'upload', 'read', 'update', 'delete'],
      timestamp: new Date().toISOString()
    };
  }
}

// Run production demo
const client = new ProductionSharePointClient();
const result = await client.fullWorkflowDemo();
console.log('🎉 Production Demo Completed:', result);
```

### **🔧 Environment Setup & Configuration**

#### **Direct MCP Mode Setup**
```json
{
  "mcpServers": {
    "sharepoint-production": {
      "command": "node",
      "args": ["/Users/arayik/sharePointMCP/src/server.js", "mcp"],
      "env": {
        "SHP_ID_APP": "12345678-1234-1234-1234-123456789012",
        "SHP_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site",
        "SHP_DOC_LIBRARY": "Shared Documents/E2E Automation", 
        "SHP_TENANT_ID": "87654321-4321-4321-4321-210987654321",
        "SHP_MAX_DEPTH": "15",
        "SHP_MAX_FOLDERS_PER_LEVEL": "100",
        "SHP_LEVEL_DELAY": "0.5",
        "SHP_CERT_PFX_PATH": "./certificate.pfx",
        "SHP_CERT_PFX_PASSWORD": "your-certificate-password",
        "LOG_LEVEL": "info"
      }
    }
  }
}
```

#### **API Client Mode Setup**  
```json
{
  "mcpServers": {
    "sharepoint-api": {
      "command": "node", 
      "args": ["/Users/arayik/sharePointMCP/src/mcp-api-client.js"],
      "env": {
        "SHAREPOINT_API_URL": "https://arayik-mcp-gdrive.onrender.com",
        "SHAREPOINT_API_TOKEN": "3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a"
      }
    }
  }
}
```

---

## ⚡ **Quick Start for AI Agents**

### **IMPORTANT: Correct Tool Names**
The actual MCP tools available are (note the differences from documentation above):
- `List_SharePoint_Folders` ✓
- `List_SharePoint_Documents` ✓  
- `Get_SharePoint_Tree` ✓
- `Get_Document_Content` (NOT `Get_SharePoint_Document_Content`)
- `Create_Folder` (NOT `Create_SharePoint_Folder`)
- `Upload_Document` (NOT `Upload_SharePoint_Document`)
- `Upload_Document_From_Path`
- `Update_Document` (NOT `Update_SharePoint_Document`)
- `Delete_SharePoint_Item`

### **Running the MCP Server**
To start the server, you need to set the environment variables from `mcp.json` manually:

```bash
cd /Users/arayik/sharePointMCP

# Set environment variables from mcp.json and run the server
# (Replace placeholder values with actual values from mcp.json env section)
SHP_ID_APP="your-azure-app-id" \
SHP_SITE_URL="your-sharepoint-site-url" \
SHP_DOC_LIBRARY="your-document-library" \
SHP_TENANT_ID="your-tenant-id" \
SHP_MAX_DEPTH="15" \
SHP_MAX_FOLDERS_PER_LEVEL="100" \
SHP_LEVEL_DELAY="0.5" \
SHP_CERT_PFX_PATH="./certificate.pfx" \
SHP_CERT_PFX_PASSWORD="your-cert-password" \
node src/server.js mcp
```

**Note**: For now, environment variables need to be set manually as shown above. Use the actual values from the `env` section in `mcp.json`. The server does not automatically read from `mcp.json` when run standalone.

**Note**: Ensure that `mcp.json` contains the correct environment variables in the `env` section of the server configuration. The server will read:
- `SHP_ID_APP` - Azure App ID
- `SHP_SITE_URL` - SharePoint site URL  
- `SHP_DOC_LIBRARY` - Document library name
- `SHP_TENANT_ID` - Azure tenant ID
- `SHP_CERT_PFX_PATH` - Path to certificate file
- `SHP_CERT_PFX_PASSWORD` - Certificate password
- Other optional settings for depth, delays, etc.

### **Corrected Upload Example**
```json
{
  "name": "Upload_Document",
  "arguments": {
    "folder_name": "Shared Documents",
    "file_name": "test-file.txt",
    "content": "File created by Arayik",
    "is_base64": false
  }
}
```

### **Complete Health Check Implementation**
```javascript
// Enterprise-grade health monitoring system
async function comprehensiveSystemCheck() {
  console.log('🏥 Starting Comprehensive System Health Check...');
  
  const healthReport = {
    timestamp: new Date().toISOString(),
    overall_status: 'unknown',
    checks: {},
    recommendations: [],
    performance_metrics: {}
  };
  
  // 1. Environment validation
  healthReport.checks.environment = await validateEnvironment();
  
  // 2. Authentication validation
  healthReport.checks.authentication = await validateAuthentication();
  
  // 3. Connectivity validation
  healthReport.checks.connectivity = await validateConnectivity();
  
  // 4. Basic operations test
  healthReport.checks.operations = await validateBasicOperations();
  
  // 5. Performance benchmarking
  healthReport.performance_metrics = await performanceBaseline();
  
  // Determine overall status and recommendations
  const failedChecks = Object.values(healthReport.checks)
    .filter(check => check.status === 'fail').length;
  
  if (failedChecks === 0) {
    healthReport.overall_status = 'healthy';
    healthReport.recommendations.push('System is fully operational');
  } else if (failedChecks <= 2) {
    healthReport.overall_status = 'degraded';
    healthReport.recommendations.push('Some issues detected, system partially functional');
  } else {
    healthReport.overall_status = 'critical';
    healthReport.recommendations.push('Multiple failures detected, immediate attention required');
  }
  
  console.log('📋 Health Check Complete:', healthReport);
  return healthReport;
}

async function validateEnvironment() {
  const requiredVars = [
    'SHP_ID_APP', 'SHP_SITE_URL', 'SHP_TENANT_ID', 
    'SHP_CERT_PFX_PATH', 'SHP_CERT_PFX_PASSWORD'
  ];
  
  const missingVars = requiredVars.filter(varName => !process.env[varName]);
  
  return {
    status: missingVars.length === 0 ? 'pass' : 'fail',
    missing_variables: missingVars,
    details: `${requiredVars.length - missingVars.length}/${requiredVars.length} variables configured`
  };
}

async function validateAuthentication() {
  try {
    const validation = await callTool("Validate_API_Connection", {});
    return {
      status: validation.valid ? 'pass' : 'fail',
      details: validation
    };
  } catch (error) {
    return {
      status: 'fail',
      error: error.message,
      troubleshooting: 'Check certificate and Azure AD configuration'
    };
  }
}

async function validateConnectivity() {
  try {
    const folders = await callTool("List_SharePoint_Folders", {});
    return {
      status: 'pass',
      folder_count: folders.count,
      response_time: 'measured internally'
    };
  } catch (error) {
    return {
      status: 'fail',
      error: error.message,
      troubleshooting: 'Check network connectivity and SharePoint site URL'
    };
  }
}

async function validateBasicOperations() {
  const operations = [
    { name: 'List_SharePoint_Folders', params: {} },
    { name: 'Get_SharePoint_Tree', params: { max_depth: 1 } }
  ];
  
  const results = {};
  
  for (const op of operations) {
    try {
      const result = await callTool(op.name, op.params);
      results[op.name] = {
        status: 'pass',
        result_size: JSON.stringify(result).length
      };
    } catch (error) {
      results[op.name] = {
        status: 'fail',
        error: error.message
      };
    }
  }
  
  const passCount = Object.values(results).filter(r => r.status === 'pass').length;
  
  return {
    status: passCount === operations.length ? 'pass' : 'partial',
    operations_tested: operations.length,
    operations_passed: passCount,
    details: results
  };
}

async function performanceBaseline() {
  const benchmarks = {};
  
  // Measure folder listing performance
  const start = Date.now();
  try {
    await callTool("List_SharePoint_Folders", {});
    benchmarks.folder_listing_ms = Date.now() - start;
  } catch (error) {
    benchmarks.folder_listing_ms = -1;
  }
  
  // Measure tree structure performance
  const treeStart = Date.now();
  try {
    await callTool("Get_SharePoint_Tree", { max_depth: 2 });
    benchmarks.tree_structure_ms = Date.now() - treeStart;
  } catch (error) {
    benchmarks.tree_structure_ms = -1;
  }
  
  return benchmarks;
}
```

---

## 📞 **Support & Troubleshooting**

### **Critical Error Resolution Guide**

#### **🚨 Authentication Failures**
**Symptoms**: "Authentication failed", "Invalid credentials", "AADSTS errors"

**Resolution Steps**:
1. **Verify Certificate**:
   ```bash
   # Check certificate file exists and has content
   ls -la ./certificate.pfx
   file ./certificate.pfx
   ```

2. **Validate Environment Variables**:
   ```javascript
   // Run this diagnostic
   console.log({
     app_id: process.env.SHP_ID_APP?.substring(0, 8) + '...',
     tenant_id: process.env.SHP_TENANT_ID?.substring(0, 8) + '...',
     site_url: process.env.SHP_SITE_URL,
     cert_path: process.env.SHP_CERT_PFX_PATH
   });
   ```

3. **Test Azure AD Configuration**:
   - Verify app registration in Azure Portal
   - Check certificate thumbprint matches uploaded certificate
   - Confirm API permissions are granted and admin consented

#### **🌐 Network Connectivity Issues**
**Symptoms**: "Connection refused", "Timeout", "Network error"

**Resolution Steps**:
1. **Test Basic Connectivity**:
   ```bash
   # Test SharePoint site accessibility
   curl -I https://your-tenant.sharepoint.com/sites/your-site
   ```

2. **Check Proxy/Firewall Settings**:
   ```javascript
   // Configure proxy if needed
   process.env.HTTPS_PROXY = 'http://your-proxy:port';
   ```

3. **Validate SSL/TLS**:
   ```bash
   # Check SSL certificate chain
   openssl s_client -connect your-tenant.sharepoint.com:443 -servername your-tenant.sharepoint.com
   ```

#### **⚡ Performance Issues**
**Symptoms**: Slow responses, timeouts, high memory usage

**Resolution Steps**:
1. **Optimize Query Parameters**:
   ```javascript
   // Use pagination for large results
   const folders = await callTool("List_SharePoint_Folders", {
     page_size: 50,
     include_details: false
   });
   
   // Limit tree depth
   const tree = await callTool("Get_SharePoint_Tree", {
     max_depth: 2,
     max_folders_per_level: 25
   });
   ```

2. **Implement Caching**:
   ```javascript
   // Cache frequently accessed data
   const cache = new Map();
   const CACHE_TTL = 5 * 60 * 1000; // 5 minutes
   
   async function getCachedFolders() {
     const cacheKey = 'folders';
     const cached = cache.get(cacheKey);
     
     if (cached && (Date.now() - cached.timestamp) < CACHE_TTL) {
       return cached.data;
     }
     
     const fresh = await callTool("List_SharePoint_Folders", {});
     cache.set(cacheKey, { data: fresh, timestamp: Date.now() });
     return fresh;
   }
   ```

3. **Monitor Resource Usage**:
   ```javascript
   // Memory usage monitoring
   const used = process.memoryUsage();
   console.log('Memory Usage:', {
     rss: Math.round(used.rss / 1024 / 1024) + ' MB',
     heapTotal: Math.round(used.heapTotal / 1024 / 1024) + ' MB',
     heapUsed: Math.round(used.heapUsed / 1024 / 1024) + ' MB'
   });
   ```

### **Debug Configuration**

#### **Enable Comprehensive Logging**
```javascript
// Set environment variables for detailed debugging
process.env.LOG_LEVEL = 'debug';
process.env.NODE_DEBUG = 'sharepoint,mcp';
process.env.DEBUG = 'sharepoint:*,mcp:*';
```

#### **Production Monitoring Setup**
```javascript
// Production health monitoring endpoint
class ProductionMonitor {
  constructor() {
    this.lastHealthCheck = null;
    this.healthCheckInterval = 5 * 60 * 1000; // 5 minutes
    this.alerts = [];
  }
  
  async startMonitoring() {
    console.log('🔍 Starting production monitoring...');
    
    setInterval(async () => {
      try {
        const health = await comprehensiveSystemCheck();
        this.lastHealthCheck = health;
        
        if (health.overall_status === 'critical') {
          this.sendAlert('CRITICAL', 'System health check failed', health);
        } else if (health.overall_status === 'degraded') {
          this.sendAlert('WARNING', 'System performance degraded', health);
        }
        
      } catch (error) {
        this.sendAlert('ERROR', 'Health check execution failed', { error: error.message });
      }
    }, this.healthCheckInterval);
  }
  
  sendAlert(level, message, details) {
    const alert = {
      level,
      message,
      details,
      timestamp: new Date().toISOString()
    };
    
    this.alerts.push(alert);
    console.log(`🚨 [${level}] ${message}`, details);
    
    // Keep only last 50 alerts
    if (this.alerts.length > 50) {
      this.alerts.shift();
    }
  }
  
  getStatus() {
    return {
      monitoring_active: true,
      last_health_check: this.lastHealthCheck,
      recent_alerts: this.alerts.slice(-10)
    };
  }
}
```

### **Emergency Recovery Procedures**

#### **Service Recovery Checklist**
1. ✅ **Verify Environment**: Check all required environment variables
2. ✅ **Test Authentication**: Run authentication validation
3. ✅ **Check Connectivity**: Ping SharePoint endpoints
4. ✅ **Validate Certificates**: Ensure certificate files are accessible
5. ✅ **Review Logs**: Check error logs for specific issues
6. ✅ **Test Basic Operations**: Try simple folder listing
7. ✅ **Monitor Performance**: Check response times and resource usage

#### **Quick Fix Commands**
```bash
# Restart MCP server
pkill -f "mcp-sharepoint"
node /Users/arayik/sharePointMCP/src/server.js mcp

# Check server logs
tail -f /Users/arayik/sharePointMCP/mcp_sharepoint.log

# Test API server directly
curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     https://arayik-mcp-gdrive.onrender.com/health
```

---

**💡 Remember**: This MCP server provides enterprise-grade SharePoint integration with security best practices. Always follow your organization's data handling policies when working with SharePoint content.

---

## 🤖 **Complete AI Assistant Integration Guide**

### **For AI Assistants and Automated Systems**

#### **Essential Integration Principles**

1. **🎯 Direct Tool Usage**: Always use MCP tools directly through the interface. Never create workaround scripts or alternative implementations.

2. **🔄 Error Recovery**: Implement comprehensive error handling with exponential backoff for all operations.

3. **📊 Performance Monitoring**: Track operation timing and implement optimization strategies for production use.

4. **🔒 Security First**: Validate all inputs, sanitize file names, and follow security best practices.

#### **Recommended AI Agent Workflow**

```javascript
class SharePointAIAgent {
  constructor(mode = 'mcp') {
    this.mode = mode;
    this.operationLog = [];
    this.performanceMetrics = new Map();
  }

  async initialize() {
    console.log('🤖 AI Agent Initializing SharePoint Connection...');
    
    // Step 1: Validate system health
    const healthCheck = await this.performHealthCheck();
    if (!healthCheck.healthy) {
      throw new Error(`System not ready: ${healthCheck.issues.join(', ')}`);
    }
    
    // Step 2: Authenticate and test connection
    const authTest = await callTool("Validate_API_Connection", {});
    if (!authTest.valid) {
      throw new Error('Authentication failed - check credentials');
    }
    
    // Step 3: Perform initial discovery
    const initialDiscovery = await this.discoverSharePointStructure();
    
    console.log('✅ AI Agent Successfully Initialized');
    console.log(`📁 Discovered ${initialDiscovery.folderCount} folders`);
    console.log(`📄 Found ${initialDiscovery.documentCount} documents`);
    
    return {
      status: 'initialized',
      capabilities: this.getCapabilities(),
      discovery: initialDiscovery
    };
  }

  async performHealthCheck() {
    const issues = [];
    let healthy = true;
    
    try {
      // Test basic folder listing
      const folders = await callTool("List_SharePoint_Folders", {});
      if (!folders.items || folders.count === 0) {
        issues.push('No folders accessible');
        healthy = false;
      }
    } catch (error) {
      issues.push(`Folder listing failed: ${error.message}`);
      healthy = false;
    }
    
    return { healthy, issues };
  }

  async discoverSharePointStructure() {
    const discovery = {
      folderCount: 0,
      documentCount: 0,
      structure: null,
      largestFolder: null,
      commonFileTypes: []
    };
    
    try {
      // Get complete structure
      const tree = await callTool("Get_SharePoint_Tree", { max_depth: 3 });
      discovery.structure = tree.tree;
      
      // Analyze structure
      const analysis = this.analyzeStructure(tree.tree);
      discovery.folderCount = analysis.folders;
      discovery.documentCount = analysis.documents;
      discovery.largestFolder = analysis.largestFolder;
      discovery.commonFileTypes = analysis.fileTypes;
      
    } catch (error) {
      console.warn('⚠️ Structure discovery partially failed:', error.message);
    }
    
    return discovery;
  }

  analyzeStructure(tree, basePath = '') {
    let folders = 0;
    let documents = 0;
    let largestFolder = { name: '', documentCount: 0 };
    const fileTypes = new Set();
    
    if (tree.items) {
      for (const item of tree.items) {
        if (item.type === 'folder') {
          folders++;
          
          // Recursively analyze subfolders
          if (item.children && item.children.items) {
            const subAnalysis = this.analyzeStructure(item.children, 
              basePath ? `${basePath}/${item.name}` : item.name);
            folders += subAnalysis.folders;
            documents += subAnalysis.documents;
            
            // Track largest folder
            if (subAnalysis.documents > largestFolder.documentCount) {
              largestFolder = {
                name: basePath ? `${basePath}/${item.name}` : item.name,
                documentCount: subAnalysis.documents
              };
            }
            
            subAnalysis.fileTypes.forEach(type => fileTypes.add(type));
          }
        } else if (item.type === 'file') {
          documents++;
          if (item.mimeType) {
            fileTypes.add(item.mimeType);
          }
        }
      }
    }
    
    return {
      folders,
      documents,
      largestFolder,
      fileTypes: Array.from(fileTypes)
    };
  }

  getCapabilities() {
    return {
      read_operations: [
        'List folders and documents',
        'Get document content',
        'Analyze folder structures',
        'Search and filter content'
      ],
      write_operations: [
        'Upload new documents',
        'Update existing documents',
        'Create folder structures',
        'Delete files and folders'
      ],
      advanced_features: [
        'Bulk operations with error recovery',
        'Performance monitoring',
        'Content analysis and reporting',
        'Automated workflow execution'
      ]
    };
  }

  async performIntelligentDocumentAnalysis(folderName) {
    console.log(`🔍 Analyzing documents in folder: ${folderName}`);
    
    const analysis = {
      folder: folderName,
      totalDocuments: 0,
      totalSize: 0,
      fileTypes: {},
      largestFiles: [],
      recentFiles: [],
      contentSummary: {}
    };
    
    try {
      // Get all documents in folder
      const documents = await callTool("List_SharePoint_Documents", {
        folder_name: folderName
      });
      
      analysis.totalDocuments = documents.count;
      
      for (const doc of documents.items.slice(0, 10)) { // Limit to first 10
        // Analyze file metadata
        if (doc.size) {
          analysis.totalSize += doc.size;
        }
        
        const fileExt = doc.name.split('.').pop()?.toLowerCase();
        if (fileExt) {
          analysis.fileTypes[fileExt] = (analysis.fileTypes[fileExt] || 0) + 1;
        }
        
        // Track largest files
        if (doc.size && doc.size > 1024 * 1024) { // > 1MB
          analysis.largestFiles.push({
            name: doc.name,
            size: doc.size,
            sizeFormatted: `${(doc.size / (1024 * 1024)).toFixed(2)} MB`
          });
        }
        
        // Analyze content for text files
        if (this.isTextFile(doc.name)) {
          try {
            const content = await callTool("Get_Document_Content", {
              folder_name: folderName,
              file_name: doc.name
            });
            
            analysis.contentSummary[doc.name] = {
              length: content.content?.length || 0,
              wordCount: content.content ? content.content.split(/\s+/).length : 0,
              lineCount: content.content ? content.content.split('\n').length : 0
            };
            
          } catch (error) {
            console.warn(`⚠️ Could not analyze content for ${doc.name}`);
          }
        }
      }
      
      // Sort largest files by size
      analysis.largestFiles.sort((a, b) => b.size - a.size);
      analysis.largestFiles = analysis.largestFiles.slice(0, 5);
      
    } catch (error) {
      console.error('❌ Document analysis failed:', error.message);
      analysis.error = error.message;
    }
    
    console.log('📊 Document Analysis Complete:', analysis);
    return analysis;
  }

  isTextFile(filename) {
    const textExtensions = ['txt', 'md', 'csv', 'json', 'xml', 'log', 'html', 'js', 'py', 'sql'];
    const extension = filename.split('.').pop()?.toLowerCase();
    return textExtensions.includes(extension);
  }

  async executeAutomatedWorkflow(workflowType, parameters) {
    console.log(`🔄 Executing automated workflow: ${workflowType}`);
    
    const workflows = {
      'content_audit': () => this.performContentAudit(parameters),
      'bulk_organize': () => this.performBulkOrganization(parameters),
      'cleanup_duplicates': () => this.performDuplicateCleanup(parameters),
      'backup_important': () => this.performImportantBackup(parameters)
    };
    
    const workflow = workflows[workflowType];
    if (!workflow) {
      throw new Error(`Unknown workflow type: ${workflowType}`);
    }
    
    try {
      const result = await workflow();
      console.log(`✅ Workflow ${workflowType} completed successfully`);
      return result;
    } catch (error) {
      console.error(`❌ Workflow ${workflowType} failed:`, error.message);
      throw error;
    }
  }

  async performContentAudit(parameters) {
    const auditReport = {
      timestamp: new Date().toISOString(),
      totalFolders: 0,
      totalDocuments: 0,
      folderAnalysis: [],
      recommendations: []
    };
    
    // Get all folders
    const folders = await callTool("List_SharePoint_Folders", {});
    auditReport.totalFolders = folders.count;
    
    // Analyze each folder
    for (const folder of folders.items.slice(0, 5)) { // Limit for demo
      try {
        const analysis = await this.performIntelligentDocumentAnalysis(folder.name);
        auditReport.folderAnalysis.push(analysis);
        auditReport.totalDocuments += analysis.totalDocuments;
        
        // Generate recommendations
        if (analysis.totalDocuments > 100) {
          auditReport.recommendations.push(
            `Consider organizing ${folder.name} - has ${analysis.totalDocuments} documents`
          );
        }
        
      } catch (error) {
        console.warn(`⚠️ Could not audit folder ${folder.name}`);
      }
    }
    
    return auditReport;
  }

  logOperation(operation, result, duration) {
    this.operationLog.push({
      operation,
      timestamp: new Date().toISOString(),
      duration,
      success: !result.error,
      result: result.error ? { error: result.error } : { success: true }
    });
    
    // Keep only last 100 operations
    if (this.operationLog.length > 100) {
      this.operationLog.shift();
    }
  }

  generateActivityReport() {
    const report = {
      session_start: this.operationLog[0]?.timestamp,
      total_operations: this.operationLog.length,
      successful_operations: this.operationLog.filter(op => op.success).length,
      failed_operations: this.operationLog.filter(op => !op.success).length,
      average_response_time: this.calculateAverageResponseTime(),
      most_used_operations: this.getMostUsedOperations(),
      performance_summary: this.getPerformanceSummary()
    };
    
    console.log('📈 Session Activity Report:', report);
    return report;
  }

  calculateAverageResponseTime() {
    const times = this.operationLog.map(op => op.duration).filter(Boolean);
    return times.length > 0 ? times.reduce((a, b) => a + b, 0) / times.length : 0;
  }

  getMostUsedOperations() {
    const counts = {};
    this.operationLog.forEach(op => {
      counts[op.operation] = (counts[op.operation] || 0) + 1;
    });
    
    return Object.entries(counts)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 5);
  }

  getPerformanceSummary() {
    const recent = this.operationLog.slice(-10);
    return {
      recent_operations: recent.length,
      recent_success_rate: recent.filter(op => op.success).length / recent.length,
      recent_average_time: recent.map(op => op.duration).filter(Boolean)
        .reduce((a, b) => a + b, 0) / recent.length || 0
    };
  }
}

// Usage Examples for AI Assistants
async function aiAssistantExample() {
  console.log('🤖 Starting AI Assistant SharePoint Integration Example');
  
  // Initialize AI agent
  const agent = new SharePointAIAgent();
  const initResult = await agent.initialize();
  
  // Perform intelligent analysis
  console.log('\n📊 Performing Intelligent Document Analysis...');
  if (initResult.discovery.structure?.items?.length > 0) {
    const firstFolder = initResult.discovery.structure.items
      .find(item => item.type === 'folder');
    
    if (firstFolder) {
      const analysis = await agent.performIntelligentDocumentAnalysis(firstFolder.name);
      console.log('Analysis Results:', analysis);
    }
  }
  
  // Execute automated workflow
  console.log('\n🔄 Executing Automated Content Audit...');
  const auditResult = await agent.executeAutomatedWorkflow('content_audit', {});
  
  // Generate session report
  console.log('\n📈 Generating Session Report...');
  const report = agent.generateActivityReport();
  
  return {
    initialization: initResult,
    analysis: 'completed',
    audit: auditResult,
    session_report: report
  };
}
```

### **Best Practices for AI Agent Implementation**

#### **1. Robust Error Handling**
```javascript
// Always wrap MCP calls in try-catch with retry logic
async function robustMCPCall(toolName, parameters, maxRetries = 3) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const result = await callTool(toolName, parameters);
      return result;
    } catch (error) {
      if (attempt === maxRetries) {
        throw new Error(`${toolName} failed after ${maxRetries} attempts: ${error.message}`);
      }
      
      // Exponential backoff
      await new Promise(resolve => 
        setTimeout(resolve, Math.pow(2, attempt) * 1000)
      );
    }
  }
}
```

#### **2. Input Validation and Sanitization**
```javascript
function sanitizeSharePointInput(input) {
  if (typeof input !== 'string') return '';
  
  // Remove invalid SharePoint characters
  return input
    .replace(/[<>:"/\\|?*]/g, '_')
    .replace(/\.$/, '')  // No trailing dots
    .replace(/^\s+|\s+$/g, '') // Trim whitespace
    .substring(0, 255); // Limit length
}

// Usage in AI operations
const safeFileName = sanitizeSharePointInput(userProvidedFileName);
const safeFolderName = sanitizeSharePointInput(userProvidedFolderName);
```

#### **3. Performance Optimization**
```javascript
// Implement caching for frequently accessed data
class SharePointCache {
  constructor(ttl = 300000) { // 5 minutes default TTL
    this.cache = new Map();
    this.ttl = ttl;
  }
  
  get(key) {
    const item = this.cache.get(key);
    if (!item) return null;
    
    if (Date.now() - item.timestamp > this.ttl) {
      this.cache.delete(key);
      return null;
    }
    
    return item.data;
  }
  
  set(key, data) {
    this.cache.set(key, {
      data,
      timestamp: Date.now()
    });
  }
}

const cache = new SharePointCache();

async function getCachedFolders() {
  const cached = cache.get('folders');
  if (cached) return cached;
  
  const folders = await callTool("List_SharePoint_Folders", {});
  cache.set('folders', folders);
  return folders;
}
```

#### **4. Intelligent Batching**
```javascript
// Process operations in intelligent batches
async function batchProcessDocuments(folderName, processor, batchSize = 5) {
  const documents = await callTool("List_SharePoint_Documents", {
    folder_name: folderName
  });
  
  const results = [];
  
  for (let i = 0; i < documents.items.length; i += batchSize) {
    const batch = documents.items.slice(i, i + batchSize);
    
    // Process batch concurrently
    const batchResults = await Promise.allSettled(
      batch.map(doc => processor(doc, folderName))
    );
    
    results.push(...batchResults);
    
    // Rate limiting delay between batches
    if (i + batchSize < documents.items.length) {
      await new Promise(resolve => setTimeout(resolve, 1000));
    }
  }
  
  return results;
}
```

### **Production Deployment for AI Systems**

```javascript
// Production-ready AI agent configuration
class ProductionSharePointAI {
  constructor(config) {
    this.config = {
      maxRetries: 3,
      timeoutMs: 30000,
      batchSize: 5,
      cacheTimeout: 300000,
      logLevel: 'info',
      ...config
    };
    
    this.cache = new SharePointCache(this.config.cacheTimeout);
    this.metrics = new Map();
    this.healthStatus = 'unknown';
  }
  
  async start() {
    console.log('🚀 Starting Production SharePoint AI Agent...');
    
    // Initialize health monitoring
    this.startHealthMonitoring();
    
    // Validate configuration
    await this.validateConfiguration();
    
    // Perform initial connection test
    await this.establishConnection();
    
    console.log('✅ Production AI Agent Ready');
    return { status: 'ready', capabilities: this.getCapabilities() };
  }
  
  startHealthMonitoring() {
    setInterval(async () => {
      try {
        const health = await this.performHealthCheck();
        this.healthStatus = health.status;
        
        if (health.status === 'unhealthy') {
          console.error('🚨 Health check failed:', health.issues);
        }
      } catch (error) {
        this.healthStatus = 'error';
        console.error('❌ Health monitoring error:', error.message);
      }
    }, 60000); // Check every minute
  }
  
  async performHealthCheck() {
    const issues = [];
    
    try {
      // Test basic connectivity
      await callTool("Validate_API_Connection", {});
    } catch (error) {
      issues.push(`Authentication: ${error.message}`);
    }
    
    try {
      // Test basic operations
      await callTool("List_SharePoint_Folders", {});
    } catch (error) {
      issues.push(`Operations: ${error.message}`);
    }
    
    return {
      status: issues.length === 0 ? 'healthy' : 'unhealthy',
      issues,
      timestamp: new Date().toISOString()
    };
  }
}
```

---

**🎯 Key Takeaways for AI Assistants:**

1. **Always use direct MCP tool calls** - Never create workarounds or shell scripts
2. **Implement comprehensive error handling** with retry logic and exponential backoff
3. **Cache frequently accessed data** to improve performance and reduce API calls
4. **Sanitize all user inputs** before passing to SharePoint operations
5. **Monitor performance metrics** and implement intelligent batching for bulk operations
6. **Use health checks** to ensure system reliability in production environments
7. **Follow security best practices** including input validation and access logging

This SharePoint MCP server provides enterprise-grade integration capabilities designed specifically for AI assistant and automation workflows. Use these patterns and examples as a foundation for building intelligent SharePoint automation systems.

---

## 📝 **Documentation Creation Guidelines**

### **🎯 Documentation Format Standards**

When creating documentation for SharePoint operations, analyses, or any work performed using this MCP server, follow these established patterns for consistency and clarity:

#### **📋 Required Document Structure**

**1. Header with Icon + Title**:
```markdown
📦 [Project/Feature Name]
```

**2. Purpose Section** (Always Required):
```markdown
🧠 Purpose
[Clear, concise explanation of what it does, why it exists, and what it validates/provides]
```

**3. Core Content Sections** (Choose Applicable):
```markdown
⚙️ Environment & Setup
🧩 Implementation Details  
🧪 Tests & Validations
🔧 Configuration Options
🧠 Advanced Features
🔁 Workflow/Lifecycle
🧰 Dependencies & Tools
🧾 Usage Examples
🧭 Summary & Use Cases
```

#### **🎨 Visual Design Patterns**

**Emoji Usage Guidelines**:
- **🧠** - Purpose, thinking, analysis
- **⚙️** - Technical setup, configuration
- **🧪** - Testing, validation, verification
- **🔧** - Tools, utilities, configuration options
- **📊** - Data, metrics, reports
- **🔁** - Processes, workflows, lifecycles
- **🧰** - Dependencies, requirements, tools
- **🧾** - Examples, samples, commands
- **🧭** - Summary, navigation, conclusions
- **✅** - Validation points, completed items
- **🔹** - Feature lists, bullet points
- **❌** - Issues, problems, failures
- **⚠️** - Warnings, important notes

**Content Organization**:
```markdown
### **Primary Section Headers**
Use emoji + title in bold with ** markers

#### **Secondary Headers** 
Standard markdown headers for subsections

**Key Information**:
- 🔹 Use diamond bullets for feature lists
- ✅ Use checkmarks for validation/completion points
- 📝 Use tables for structured data comparisons
- 💡 Use lightbulb for tips and insights
```

#### **📊 Content Requirements Checklist**

**Every Documentation Must Include**:
1. **🧠 Clear Purpose Statement** - What it does and why it matters
2. **⚙️ Technical Setup/Requirements** - Environment, dependencies, prerequisites  
3. **🧪 Functional Details** - What it tests/validates/provides/analyzes
4. **🧾 Practical Examples** - Usage commands, code samples, scenarios
5. **🧰 Dependencies** - What tools/systems it relies on
6. **🧭 Summary with Use Cases** - When and why to use it

**Optional Enhanced Sections**:
- **📊 Data/Metrics Tables** - For structured comparisons
- **🔁 Process Workflows** - Step-by-step procedures  
- **⚠️ Troubleshooting** - Common issues and solutions
- **🔧 Advanced Configuration** - Extended setup options

#### **✍️ Writing Style Guidelines**

**Language Patterns**:
1. **Active, Present-Tense Descriptions**:
   - ✅ "Validates that configurations load correctly"
   - ✅ "Verifies SharePoint authentication"
   - ❌ "Will validate" or "Should verify"

2. **Technical Precision**:
   - ✅ Include exact function names: `Get_Document_Content`
   - ✅ Specify file paths: `src/sharepoint-tools.js`
   - ✅ Reference specific variables: `SHP_SITE_URL`

3. **User-Benefit Focus**:
   - ✅ "It's ideal for: regression testing, automation workflows"
   - ✅ "This enables: bulk document processing, content analysis"

4. **Concise Clarity**:
   - ✅ Short paragraphs (2-3 sentences max)
   - ✅ Bullet points for lists
   - ✅ Tables for structured data

#### **🔧 Markdown Technical Standards**

**Code Blocks**:
```markdown
# Use language-specific syntax highlighting
```javascript
const result = await callTool("List_SharePoint_Folders", {});
```

**Tables Format**:
```markdown
| Variable | Description | Default |
|----------|-------------|---------|
| serverEnv | Target environment | production |
```

**Link Standards**:
```markdown
- Use descriptive link text
- Reference specific sections: [Authentication Setup](#authentication-setup)
```

#### **� Pre-Documentation Analysis Requirements**

**CRITICAL: AI Agent Conduct & Operational Requirements**

Before performing any E2E documentation operation, the AI agent **MUST** read and understand all instructions in this file. The agent is required to read all lines of this documentation in their entirety before proceeding with any task. This ensures compliance with all workflow, security, and operational requirements.

### **🚦 E2E Documentation Workflow Summary**

Whenever a user requests an E2E documentation task involving SharePoint, the AI agent will:
- Use the SharePoint MCP server tools for all document operations
- Always read local E2E code files first to understand current implementation  
- Compare local code with existing SharePoint documentation
- Follow the mandatory two-phase analysis protocol described below

**IMPORTANT:** The agent must analyze local E2E code files **before** reading SharePoint documentation to ensure accurate comparison and gap identification.

### **📋 Protocol: Update All E2E Documentation**

When tasked with "update documentation" or general E2E documentation review:

**Step 1: Read SharePoint Documentation Inventory**
```javascript
// Get list of all documentation files in SharePoint
const docFiles = await callTool("List_SharePoint_Documents", {
  folder_name: "Documentation" // or appropriate docs folder
});
console.log(`Found ${docFiles.count} documentation files`);
```

**Step 2: Read All Existing Documentation** 
```javascript
// Read every documentation file completely
for (const doc of docFiles.items) {
  const content = await callTool("Get_Document_Content", {
    folder_name: "Documentation",
    file_name: doc.name
  });
  
  // Store content for comparison analysis
  existingDocs[doc.name] = {
    content: content.content,
    lastModified: doc.lastModified,
    size: doc.size
  };
}
```

**Step 3: Read All Local E2E Code Files**
```javascript
// Must read ALL E2E test files from local repository
// Read e2e/playwright/e2e/*.e2e.js files (every line)
const localE2EFiles = await readLocalDirectory("e2e/playwright/e2e/");

for (const testFile of localE2EFiles) {
  if (testFile.endsWith('.e2e.js')) {
    const testContent = await readLocalFile(testFile);
    
    // Also read all related files/functions used in the test
    const dependencies = extractDependencies(testContent);
    for (const dep of dependencies) {
      const depContent = await readLocalFile(dep);
      relatedFiles[dep] = depContent;
    }
    
    localTests[testFile] = {
      content: testContent,
      dependencies: dependencies
    };
  }
}
```

**Step 4: Compare Logic & Context**
```javascript
// Compare each SharePoint documentation to corresponding local test logic
for (const [docName, docData] of Object.entries(existingDocs)) {
  const correspondingTest = findCorrespondingTest(docName, localTests);
  
  if (correspondingTest) {
    const comparison = compareDocumentationToCode(docData.content, correspondingTest.content);
    
    if (comparison.hasDiscrepancies) {
      documentationUpdatesNeeded.push({
        documentName: docName,
        testFile: correspondingTest.fileName,
        discrepancies: comparison.gaps,
        suggestedUpdates: comparison.recommendations
      });
    }
  }
}
```

**Step 5: Report Findings**
```javascript
// Present findings to user for decision
if (documentationUpdatesNeeded.length > 0) {
  console.log(`📋 Found ${documentationUpdatesNeeded.length} documentation files that need updates:`);
  
  for (const update of documentationUpdatesNeeded) {
    console.log(`\n📄 ${update.documentName}:`);
    console.log(`   🔗 Corresponds to: ${update.testFile}`);
    console.log(`   ⚠️ Issues found: ${update.discrepancies.length}`);
    console.log(`   💡 Recommendations: ${update.suggestedUpdates.length}`);
  }
  
  // Let user decide whether to proceed with updates
  console.log("\n🤔 Would you like me to update these documentation files?");
}
```

### **📋 Protocol: Update Specific Documentation [name]**

When tasked with "update documentation [name]" for a specific test suite:

**Step 1: Find Specific Documentation**
```javascript
// Get list and find the specific documentation
const docFiles = await callTool("List_SharePoint_Documents", {
  folder_name: "Documentation"
});

const targetDoc = docFiles.items.find(doc => 
  doc.name.toLowerCase().includes(name.toLowerCase())
);

if (!targetDoc) {
  throw new Error(`Documentation for '${name}' not found in SharePoint`);
}
```

**Step 2: Read Existing Documentation**
```javascript
// Read the specific documentation file completely
const existingDoc = await callTool("Get_Document_Content", {
  folder_name: "Documentation", 
  file_name: targetDoc.name
});

console.log(`📖 Reading existing documentation: ${targetDoc.name}`);
console.log(`📏 Current size: ${existingDoc.content.length} characters`);
```

**Step 3: Read Corresponding Local Test Files**
```javascript
// Read the corresponding *.e2e.js file completely (every line)
const testFileName = `${name}.e2e.js`;
const testContent = await readLocalFile(`e2e/playwright/e2e/${testFileName}`);

// Read ALL related files/functions used in the test
const relatedFiles = {};
const dependencies = extractAllDependencies(testContent);

for (const dep of dependencies) {
  const depPath = resolveDependencyPath(dep);
  relatedFiles[dep] = await readLocalFile(depPath);
}

console.log(`🧪 Analyzed test file: ${testFileName}`);
console.log(`🔗 Found ${dependencies.length} related dependencies`);
```

**Step 4: Compare and Analyze**
```javascript
// Compare SharePoint documentation to local test logic
const analysis = {
  currentDocumentation: existingDoc.content,
  currentTestLogic: testContent,
  relatedCode: relatedFiles,
  discrepancies: [],
  suggestedAdditions: [],
  suggestedChanges: []
};

// Detailed comparison analysis
const comparison = performDetailedComparison(analysis);

console.log(`🔍 Analysis complete:`);
console.log(`   📊 Found ${comparison.discrepancies.length} discrepancies`);
console.log(`   ➕ Suggested ${comparison.suggestedAdditions.length} additions`);
console.log(`   ✏️ Suggested ${comparison.suggestedChanges.length} changes`);
```

**Step 5: Create and Upload Improved Documentation**
```javascript
// Generate improved documentation based on analysis
const improvedDoc = generateImprovedDocumentation({
  existingContent: existingDoc.content,
  testLogic: testContent,
  relatedFiles: relatedFiles,
  analysis: comparison
});

// Upload the new improved documentation
const uploadResult = await callTool("Upload_Document", {
  folder_name: "Documentation",
  file_name: `${name}-updated-${new Date().toISOString().split('T')[0]}.md`,
  content: improvedDoc,
  overwrite: false
});

console.log(`✅ Updated documentation uploaded: ${uploadResult.message}`);
```

**CRITICAL: Before Any E2E Documentation Work**

When tasked with updating or creating documentation for End-to-End (E2E) tests, automation workflows, or any project documentation, AI agents **MUST** perform comprehensive file analysis using the protocols above:

**� CRITICAL: NO LOCAL FILE CREATION RULE**

**AI agents MUST NEVER create documentation files locally.** All documentation must be:
1. **Created in memory** - Content is generated and stored in the agent's working memory
2. **Uploaded directly to SharePoint** - Use SharePoint MCP tools to upload content directly
3. **No intermediate local files** - Do NOT use `create_file`, `write_file`, or any file system operations

**Correct Workflow**:
```javascript
// ✅ CORRECT: Create content in memory, upload directly to SharePoint
const documentationContent = generateDocumentation(analysis);

await callTool("Upload_Document", {
  folder_name: "E2E Automation",
  file_name: "adConfig-E2E-Test-Suite-Documentation.md",
  content: documentationContent,
  is_base64: false
});
```

**Incorrect Workflow**:
```javascript
// ❌ INCORRECT: Never create local files
create_file("/path/to/local/documentation.md", content);  // FORBIDDEN
fs.writeFileSync("./doc.md", content);  // FORBIDDEN
```

**Why This Matters**:
- Prevents cluttering the user's local filesystem
- Ensures single source of truth in SharePoint
- Follows enterprise document management practices
- Maintains clean workspace hygiene

**�📂 Complete Folder Analysis Process**:
1. **�📋 Inventory All Files**:
   ```javascript
   // Use SharePoint MCP tools to discover complete structure
   const folderStructure = await callTool("Get_SharePoint_Tree", {
     folder_name: "E2E",
     max_depth: 5
   });
   
   // Get detailed file listing
   const allFiles = await callTool("List_SharePoint_Documents", {
     folder_name: "E2E"
   });
   ```

2. **📖 Read Every File Systematically**:
   ```javascript
   // Read all files in the E2E directory and subdirectories
   for (const file of allFiles.items) {
     const content = await callTool("Get_Document_Content", {
       folder_name: "E2E",
       file_name: file.name
     });
     
     // Analyze: test files, config files, utilities, documentation
     analyzeFileContent(file.name, content.content);
   }
   ```

3. **🧩 Understand Relationships & Dependencies**:
   - **Test files** (.js, .ts, .spec files) - What do they test?
   - **Configuration files** (package.json, config files) - What dependencies and setup?
   - **Utility files** - What shared functions and helpers?
   - **Data files** (.json, .csv) - What test data and expected results?
   - **Documentation files** (.md, .txt) - What existing context and requirements?

4. **🎯 Map Test Coverage & Purpose**:
   - Identify what systems/features are being tested
   - Understand test scope and validation points  
   - Document dependencies between test files
   - Note configuration requirements and environment setup

**🚫 Never Create Documentation Without**:
- [ ] **Reading ALL local E2E test files** in the target area (e2e/playwright/e2e/*.e2e.js)
- [ ] **Reading existing SharePoint documentation** for comparison
- [ ] **Analyzing ALL related dependencies** and utility functions used in tests
- [ ] **Performing detailed comparison** between documentation and actual code
- [ ] **Identifying specific gaps and discrepancies** with evidence
- [ ] **Following the mandatory 5-step protocol** for either general or specific updates

**⚠️ Documentation Integrity Rule**: 
Documentation must accurately reflect the **actual implementation** found in local E2E files. Any discrepancies between SharePoint docs and local code must be identified, analyzed, and corrected based on the current codebase state. The agent must read local code files **first**, then compare with SharePoint documentation to ensure accuracy.

#### **📋 Quality Validation Checklist**

Before finalizing any documentation, verify:

- [ ] **🔍 Complete file analysis performed** - All files in target directory read and understood
- [ ] **📂 File relationships mapped** - Dependencies and connections documented  
- [ ] **🧠 Purpose clearly stated** in opening section based on actual implementation
- [ ] **⚙️ Setup requirements** documented with specifics from actual config files
- [ ] **🧪 Core functionality** explained with examples from actual test files
- [ ] **🧾 Usage examples** provided with actual commands/code found in files
- [ ] **🧰 Dependencies** listed with versions from actual package.json/requirements
- [ ] **🧭 Use cases** identified from actual test scenarios in files
- [ ] **📝 Formatting** consistent with emoji and hierarchy standards
- [ ] **✅ All claims validated** against actual file contents
- [ ] **🔗 Internal references** properly linked and verified to exist
- [ ] **⚠️ Important warnings** highlighted based on actual implementation notes

### **📊 Documentation Templates**

#### **E2E Test Suite Documentation Template**:
```markdown
📦 [Test Suite Name]

🧠 Purpose
[What this test suite validates and why - based on actual test file analysis]

⚙️ Environment & Setup
🧩 Imported Modules & Utilities
[List from actual import statements in test files]

🔧 Configurable Environment Variables
[Table of actual variables from config files]

🌍 Test Scope & Coverage
[Based on actual test files found and analyzed]

🧪 Tests & Validations
[Organized by actual test structure found in files]
🔹 [Test Category 1 - from actual test groupings]
✅ [Specific validations from actual test assertions]

🔹 [Test Category 2]
✅ [More validations]

🧠 Advanced Features
[Special capabilities found in utility files and test helpers]

🔁 Lifecycle Hooks
[Actual beforeAll, beforeEach, afterAll found in test files]

🧰 Tools It Depends On
[Actual dependencies from package.json and import analysis]

🧾 Sample Run Command
[Actual npm scripts and commands from package.json]

🧭 Summary
[Based on complete analysis of all files in the E2E directory]
```

#### **Analysis Report Template**:
```markdown
📊 [Analysis Name] Report

🧠 Purpose
[What was analyzed and why]

⚙️ Analysis Scope  
[What data/systems were examined]

🔍 File Analysis Summary
📂 Files Analyzed: [Total count]
📄 Test Files: [Count and types]
🔧 Config Files: [Configuration files found]
🧰 Utility Files: [Helper and shared modules]
📋 Documentation Files: [Existing docs reviewed]

🧪 Findings & Validations
[Key discoveries and verification results from actual file contents]

📈 Metrics & Data
[Quantitative results in tables - based on actual analysis]

🧭 Summary & Recommendations
[Conclusions and next steps based on comprehensive file review]
```

#### **Implementation Guide Template**:
```markdown
🔧 [Feature/Tool] Implementation Guide

🧠 Purpose
[What this implements and business value]

⚙️ Prerequisites & Setup
[Required environment and dependencies]

🧩 Implementation Steps
[Step-by-step procedure with code examples]

🧪 Testing & Validation
[How to verify it works correctly]

🧰 Troubleshooting
[Common issues and solutions]

🧭 Usage & Maintenance
[Ongoing operational guidance]
```

---

## 🤖 **AI Assistant Usage Guidelines (Continued)**

### **For AI Assistants and Automated Systems**

### **� CRITICAL: SharePoint Upload Workflow**

**AI agents have TWO methods to upload to SharePoint. NEVER create local files.**

#### **Method 1: MCP Tools (Preferred when available)**
If MCP tools like `List_SharePoint_Folders`, `Upload_Document`, etc. are available:
```javascript
// Create content in memory
const documentContent = `Your documentation content here...`;

// Upload directly using MCP tool
await callTool("Upload_Document", {
  folder_name: "E2E Automation",
  file_name: "your-doc.md",
  content: documentContent,
  is_base64: false
});
```

#### **Method 2: HTTP REST API (When MCP tools NOT available)**
If MCP tools are NOT loaded, use the HTTP REST API via `run_in_terminal`:

**Step 1: Test connection first**
```bash
curl https://arayik-mcp-gdrive.onrender.com/health
```

**Step 2: Validate authentication**
```bash
curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     https://arayik-mcp-gdrive.onrender.com/api/auth/validate
```

**Step 3: Upload document**
```bash
curl -X POST \
  -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
  -H "Content-Type: application/json" \
  -d '{
    "fileName": "your-doc.md",
    "content": "Your documentation content here...",
    "path": "E2E Automation/your-doc.md"
  }' \
  https://arayik-mcp-gdrive.onrender.com/api/upload
```

### **🚫 ABSOLUTE RULES FOR AI AGENTS**

1. **NEVER create local files** - Do NOT use `create_file`, `write_file`, or any file system tools
2. **Create content in memory only** - Store documentation as a string variable
3. **Always try MCP tools first** - Attempt to use MCP tools before falling back to HTTP API
4. **Use HTTP API if MCP unavailable** - Use `run_in_terminal` with curl commands
5. **Test connection before upload** - Always verify with `/health` and `/api/auth/validate` endpoints
6. **Never use Notion MCP** - This project ONLY uses SharePoint MCP, never Notion

### **📋 Correct Upload Workflow**

```javascript
// STEP 1: Generate documentation content in memory
const docContent = generateDocumentation(analysis);

// STEP 2: Try MCP tools first
try {
  await callTool("Upload_Document", {
    folder_name: "E2E Automation",
    file_name: "doc.md",
    content: docContent,
    is_base64: false
  });
  console.log("✅ Uploaded via MCP tools");
} catch (error) {
  // STEP 3: If MCP tools not available, use HTTP API
  console.log("⚠️ MCP tools not available, using HTTP API");
  
  // Test connection
  await run_in_terminal({
    command: 'curl https://arayik-mcp-gdrive.onrender.com/health',
    explanation: "Testing SharePoint API connection"
  });
  
  // Upload via API
  await run_in_terminal({
    command: `curl -X POST -H "Authorization: Bearer TOKEN" -H "Content-Type: application/json" -d '{"fileName":"doc.md","content":"${docContent}","path":"E2E Automation/doc.md"}' https://arayik-mcp-gdrive.onrender.com/api/upload`,
    explanation: "Uploading documentation to SharePoint via HTTP API"
  });
}
```

### **❌ WRONG - Never Do This**
```javascript
// ❌ WRONG: Creating local file
create_file("/path/to/doc.md", content);

// ❌ WRONG: Using Notion MCP
mcp_notion_API_post_page(...);

// ❌ WRONG: Using wrong tool names
Upload_SharePoint_Document(...); // This doesn't exist in direct MCP mode
```

### **✅ CORRECT - Always Do This**
```javascript
// ✅ CORRECT: Content in memory, direct upload
const content = "documentation...";
await callTool("Upload_Document", { 
  folder_name: "E2E Automation",
  file_name: "doc.md", 
  content: content 
});

// ✅ CORRECT: Or use HTTP API if MCP not available
await run_in_terminal({
  command: 'curl -X POST -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" ...',
  explanation: "Uploading to SharePoint"
});
```

### **Additional Guidelines**

1. **🚫 NO LOCAL FILE CREATION**: AI agents MUST NEVER create documentation files locally. All documentation content must be created in memory and uploaded directly to SharePoint.

2. **✅ VERIFY UPLOAD METHOD**: Check if MCP tools are available first, then fall back to HTTP API if needed.

3. **File Editing Best Practices**: When using the replace_string_in_file tool, always include 3-5 lines of unchanged code before and after the string you want to replace to make it unambiguous which part of the file should be edited.

4. **Authentication**: The server uses certificate-based authentication with Azure AD for MCP mode, or API token for HTTP API mode.

5. **Tool Names**: Use the exact tool names as documented above - they may differ slightly from what you expect (e.g., use `Get_Document_Content` not `Get_SharePoint_Document_Content`).

6. **Documentation Standards**: When creating any documentation about your work with SharePoint, follow the format guidelines above to ensure consistency, clarity, and professional presentation.

7. **SharePoint Upload Workflow - TWO METHODS ONLY**:
   - **Method A (Preferred)**: If MCP tools available → Use `Upload_Document` tool directly
   - **Method B (Fallback)**: If MCP tools NOT available → Use `run_in_terminal` with curl to HTTP API
   - **NEVER**: Create local files under any circumstances

8. **Correct HTTP API Workflow When MCP Tools Unavailable**:
   ```bash
   # Step 1: Test connection
   curl https://arayik-mcp-gdrive.onrender.com/health
   
   # Step 2: Validate authentication  
   curl -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
        https://arayik-mcp-gdrive.onrender.com/api/auth/validate
   
   # Step 3: Upload document
   curl -X POST \
     -H "Authorization: Bearer 3a0180d84d738f388d7702f283a2e1b64d9b17eadb53f440d921d52d618f4f4a" \
     -H "Content-Type: application/json" \
     -d '{"fileName":"doc.md","content":"YOUR_CONTENT","path":"E2E Automation/doc.md"}' \
     https://arayik-mcp-gdrive.onrender.com/api/upload
   ```

9. **What NOT to do**:
   - ❌ Never use `create_file` to create documentation locally
   - ❌ Never use Notion MCP tools (this project only uses SharePoint)
   - ❌ Never try to use non-existent tools like `Upload_SharePoint_Document` in direct MCP mode
   - ❌ Never skip the connection test when using HTTP API
   - ❌ Never ask user to manually configure MCP when HTTP API is available