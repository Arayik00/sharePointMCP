# 🤖 AI Agent Instructions: SharePoint MCP Server Usage

**Target Audience**: AI agents, assistants, and automated systems that need to interact with SharePoint through this MCP server.

---

## 📋 **Quick Reference**

### **What This MCP Server Does**
- **Connects to Microsoft SharePoint** via Graph API with certificate authentication
- **Provides file and folder operations** (list, read, create, upload, delete)
- **Offers dual communication modes**: Traditional MCP + HTTP/WebSocket API
- **Maintains enterprise security** with server-side credential storage

### **Available Operations**
1. **List folders** in SharePoint directories
2. **List documents** in specific folders
3. **Get folder tree** structure recursively
4. **Read document content** (text and binary files)
5. **Create new folders**
6. **Upload documents** from content or file paths
7. **Update existing documents**
8. **Delete files and folders**

---

## 🔧 **Integration Methods**

### **Method 1: Traditional MCP (Recommended for Claude)**

**Setup**: Add to MCP client configuration
```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": ["/path/to/sharePointMCP/src/server.js", "mcp"],
      "env": {
        "SHP_ID_APP": "your-azure-app-id",
        "SHP_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site",
        "SHP_TENANT_ID": "your-tenant-id",
        "SHP_CERT_PFX_PATH": "/path/to/certificate.pfx",
        "SHP_CERT_PFX_PASSWORD": "your-cert-password"
      }
    }
  }
}
```

**Usage**: Use standard MCP tool calls:
- `List_SharePoint_Folders`
- `List_SharePoint_Documents` 
- `Get_SharePoint_Tree`
- `Get_Document_Content` (note: NOT Get_SharePoint_Document_Content)
- `Create_Folder` (note: NOT Create_SharePoint_Folder)
- `Upload_Document` (note: NOT Upload_SharePoint_Document)
- `Upload_Document_From_Path`
- `Update_Document` (note: NOT Update_SharePoint_Document)
- `Delete_SharePoint_Item`

### **Method 2: HTTP API (Recommended for Custom Agents)**

**Setup**: Start API server
```bash
cd /path/to/sharePointMCP
SERVER_MODE=api PORT=3000 API_AUTH_TOKENS=your-secret-token npm start
```

**Usage**: Make HTTP requests with authentication
```bash
# List folders
curl -H "Authorization: Bearer your-secret-token" \
     "http://localhost:3000/api/folders?parentFolder=Documents"

# Get document content
curl -H "Authorization: Bearer your-secret-token" \
     "http://localhost:3000/api/document/file.txt/content"
```

### **Method 3: WebSocket (Real-time Applications)**

**Setup**: Connect to WebSocket endpoint
```javascript
const ws = new WebSocket('ws://localhost:3000?token=your-secret-token');
```

**Usage**: Send JSON commands
```javascript
ws.send(JSON.stringify({
  action: 'listFolders',
  params: { parentFolder: 'Documents' }
}));
```

---

## 🛠️ **MCP Tool Reference**

### **1. List_SharePoint_Folders**
**Purpose**: List all folders in a specified directory
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
    "parent_folder": "Documents/Projects"
  }
}
```
**Returns**: Array of folder objects with name, size, URL, lastModified

### **2. List_SharePoint_Documents**
**Purpose**: List all documents in a specified folder
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
    "folder_name": "Documents/Reports"
  }
}
```
**Returns**: Array of document objects with name, size, mimeType, URL, lastModified

### **3. Get_SharePoint_Tree**
**Purpose**: Get recursive folder structure
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
    "parent_folder": "Documents",
    "max_depth": 2
  }
}
```
**Returns**: Nested tree structure with folders and files

### **4. Get_Document_Content**
**Purpose**: Read content of a specific document
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
    "folder_name": "Documents/Reports",
    "file_name": "quarterly-report.txt"
  }
}
```
**Returns**: Document content as text (for text files) or base64 (for binary files)

### **5. Create_Folder**
**Purpose**: Create a new folder
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
    "parent_folder": "Documents/Projects",
    "folder_name": "New Project"
  }
}
```
**Returns**: Success status and created folder information

### **6. Upload_Document**
**Purpose**: Upload a document with content
**Input Schema**:**
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
    "folder_name": "Documents/Reports",
    "file_name": "new-report.txt",
    "content": "This is the content of the report...",
    "is_base64": false
  }
}
```

### **7. Update_Document**
**Purpose**: Update content of existing document
**Input Schema**:
```json
{
  "folder_name": "string (required)",   // Folder containing the document
  "file_name": "string (required)",     // Name of the file to update
  "content": "string (required)",       // New content
  "is_base64": "boolean (optional)"     // Whether content is base64 encoded
}
```

### **8. Delete_SharePoint_Item**
**Purpose**: Delete a file or folder
**Input Schema**:
```json
{
  "item_path": "string (required)" // Full path to item to delete
}
```

---

## 🌐 **HTTP API Reference**

### **Base URL**: `http://localhost:3000` (or your server URL)
### **Authentication**: Include `Authorization: Bearer your-token` header

### **Endpoints**:
- `GET /health` - Server health check (no auth required)
- `GET /api/folders?parentFolder=path` - List folders
- `GET /api/documents?folderName=path` - List documents
- `GET /api/tree?folderPath=path&maxDepth=N` - Get folder tree
- `GET /api/document/:path/content` - Get document content
- `POST /api/upload` - Upload file

### **Response Format**:
```json
{
  "success": true,
  "items": [...],
  "count": 5,
  "message": "Operation completed successfully"
}
```

---

## 🔍 **Common Usage Patterns**

### **Pattern 1: Explore SharePoint Structure**
```javascript
// 1. List root folders
const folders = await callTool("List_SharePoint_Folders", {});

// 2. Get detailed tree view
const tree = await callTool("Get_SharePoint_Tree", {
  "parent_folder": "",
  "max_depth": 2
});

// 3. List documents in specific folder
const docs = await callTool("List_SharePoint_Documents", {
  "folder_name": "Documents/Reports"
});
```

### **Pattern 2: Read and Analyze Documents**
```javascript
// 1. Find documents
const documents = await callTool("List_SharePoint_Documents", {
  "folder_name": "Documents/Data"
});

// 2. Read document content
for (const doc of documents.items) {
  const content = await callTool("Get_Document_Content", {
    "folder_name": "Documents/Data",
    "file_name": doc.name
  });
  
  // Process content...
  analyzeDocument(content);
}
```

### **Pattern 3: Create and Upload Content**
```javascript
// 1. Create folder structure
await callTool("Create_Folder", {
  "parent_folder": "Documents",
  "folder_name": "AI Generated Reports"
});

// 2. Upload generated content
await callTool("Upload_Document", {
  "folder_name": "Documents/AI Generated Reports",
  "file_name": "analysis-report.txt",
  "content": generatedReportContent,
  "is_base64": false
});
```

---

## ⚠️ **Important Notes for Agents**

### **Error Handling**
- Always check the `success` field in responses
- Handle authentication errors gracefully
- Implement retry logic for network issues
- Respect rate limits (avoid rapid consecutive calls)

### **Security Considerations**
- Never log or expose authentication tokens
- Use HTTPS in production environments
- Validate file paths to prevent directory traversal
- Be cautious with file uploads and content

### **Performance Tips**
- Use `Get_SharePoint_Tree` for bulk folder exploration
- Cache folder listings when possible
- Limit `max_depth` for large directory structures
- Process large files in chunks when possible

### **File Types**
- **Text files**: Content returned as plain text
- **Binary files**: Content returned as base64-encoded string
- **Supported formats**: All SharePoint-supported file types
- **Size limits**: Respect SharePoint's file size limitations

---

## 🚀 **Quick Start for Agents**

### **Minimal Working Example (MCP)**:
```javascript
// Check available tools
const tools = await listTools();

// List root folders
const folders = await callTool("List_SharePoint_Folders", {});
console.log("Available folders:", folders.items);

// Read a document
const content = await callTool("Get_Document_Content", {
  "folder_name": "Documents",
  "file_name": "readme.txt"
});
console.log("Document content:", content);
```

### **Minimal Working Example (HTTP)**:
```javascript
const baseURL = 'http://localhost:3000';
const token = 'your-secret-token';

// List folders
const response = await fetch(`${baseURL}/api/folders`, {
  headers: { 'Authorization': `Bearer ${token}` }
});
const folders = await response.json();
console.log("Folders:", folders.items);
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
node src/index.js
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

---

## 📞 **Support & Troubleshooting**

### **Common Issues**:
1. **Authentication failed**: Check certificate and credentials in mcp.json
2. **Connection refused**: Ensure server is running and accessible
3. **File not found**: Verify file paths and permissions
4. **Rate limited**: Implement delays between requests
5. **Environment variables missing**: Check that mcp.json has all required env variables in the server configuration

### **Debug Mode**:
Set environment variable `LOG_LEVEL=debug` for detailed logging

### **Health Check**:
Always start with `/health` endpoint to verify server status

---

**💡 Remember**: This MCP server provides enterprise-grade SharePoint integration with security best practices. Always follow your organization's data handling policies when working with SharePoint content.

---

## 🤖 **AI Assistant Usage Guidelines**

### **For AI Assistants and Automated Systems**:

1. **Direct Tool Usage**: Always use the MCP tools directly instead of creating separate scripts or workarounds. The tools are designed to be used directly through the MCP interface.

2. **File Editing Best Practices**: When using the replace_string_in_file tool, always include 3-5 lines of unchanged code before and after the string you want to replace to make it unambiguous which part of the file should be edited.

3. **Authentication**: The server uses certificate-based authentication with Azure AD. Ensure the certificate is properly configured before attempting operations.

4. **Tool Names**: Use the exact tool names as documented above - they may differ slightly from what you expect (e.g., use `Get_Document_Content` not `Get_SharePoint_Document_Content`).