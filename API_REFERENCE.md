# 📚 SharePoint MCP Server - API Reference

Complete REST API documentation for the SharePoint MCP Server.

## 🔐 Authentication

All API endpoints (except `/health`) require authentication via API token:

```bash
# Method 1: Authorization header (recommended)
curl -H "Authorization: Bearer YOUR_TOKEN" https://your-server.com/api/endpoint

# Method 2: X-API-Token header
curl -H "X-API-Token: YOUR_TOKEN" https://your-server.com/api/endpoint

# Method 3: Query parameter
curl "https://your-server.com/api/endpoint?token=YOUR_TOKEN"
```

---

## 📖 READ Operations

### `GET /api/folders`
List folders in SharePoint root or specified parent folder.

**Parameters:**
- `parentFolder` (query, optional): Parent folder path

**Response:**
```json
{
  "success": true,
  "items": [
    {
      "name": "General",
      "type": "folder", 
      "size": 377402,
      "webUrl": "https://...",
      "lastModified": "2025-04-22T20:52:21Z",
      "mimeType": ""
    }
  ],
  "count": 1
}
```

### `GET /api/documents`
List documents in SharePoint root or specified folder.

**Parameters:**
- `folderName` (query, optional): Folder name to list documents from

**Response:**
```json
{
  "success": true,
  "items": [
    {
      "name": "document.txt",
      "type": "file",
      "size": 1024,
      "webUrl": "https://...",
      "lastModified": "2025-11-10T12:00:00Z",
      "mimeType": "text/plain"
    }
  ],
  "count": 1
}
```

### `GET /api/tree`
Get hierarchical folder tree structure.

**Parameters:**
- `folderPath` (query, optional): Root folder path for tree
- `maxDepth` (query, optional): Maximum depth to traverse (default: 3)

**Response:**
```json
{
  "success": true,
  "folder": "root",
  "tree": {
    "items": [
      {
        "name": "General",
        "type": "folder",
        "children": {
          "items": [...],
          "truncated": false
        }
      }
    ],
    "truncated": false
  }
}
```

### `GET /api/document/:path/content`
Get document content.

**Parameters:**
- `path` (URL parameter): Document path (URL encoded)

**Response:**
```json
{
  "success": true,
  "content": "Document content here...",
  "mimeType": "text/plain",
  "size": 1024
}
```

---

## 📝 CREATE Operations

### `POST /api/upload`
Upload a document to SharePoint.

**Request Body:**
```json
{
  "fileName": "document.txt",
  "content": "Document content or base64 for binary files",
  "folderPath": "optional/folder/path"
}
```

**Response:**
```json
{
  "success": true,
  "message": "Document document.txt uploaded successfully",
  "item": {
    "name": "document.txt",
    "path": "optional/folder/path/document.txt",
    "webUrl": "https://..."
  }
}
```

### `POST /api/folder`
Create a new folder.

**Request Body:**
```json
{
  "folderName": "New Folder",
  "parentPath": "optional/parent/path"
}
```

**Response:**
```json
{
  "success": true,
  "message": "Folder 'New Folder' created successfully",
  "folder": {
    "name": "New Folder",
    "path": "optional/parent/path/New Folder",
    "webUrl": "https://..."
  }
}
```

---

## ✏️ UPDATE Operations

### `PUT /api/document/:path`
Update document content.

**Parameters:**
- `path` (URL parameter): Document path (URL encoded)

**Request Body:**
```json
{
  "content": "Updated document content"
}
```

**Response:**
```json
{
  "success": true,
  "message": "Document updated successfully",
  "item": {
    "name": "document.txt",
    "path": "folder/document.txt",
    "webUrl": "https://..."
  }
}
```

---

## 🗑️ DELETE Operations

### `DELETE /api/item/:path`
Delete any item (file or folder).

**Parameters:**
- `path` (URL parameter): Item path (URL encoded)

**Response:**
```json
{
  "success": true,
  "message": "Item deleted successfully",
  "deletedItem": {
    "path": "folder/item",
    "type": "file"
  }
}
```

### `DELETE /api/document/:path`
Delete a document (alias for `/api/item/:path`).

### `DELETE /api/folder/:path` 
Delete a folder (alias for `/api/item/:path`).

---

## 🏥 Utility Endpoints

### `GET /health`
Server health check (no authentication required).

**Response:**
```json
{
  "status": "healthy",
  "timestamp": "2025-11-10T12:00:00Z",
  "sharepoint": "connected",
  "authentication": "required",
  "endpoints": [...]
}
```

### `GET /api/auth/validate`
Validate API token.

**Response:**
```json
{
  "valid": true,
  "message": "API token is valid",
  "authenticated": true,
  "token_preview": "1098b157...",
  "timestamp": "2025-11-10T12:00:00Z"
}
```

### `GET /api`
List all available endpoints.

**Response:**
```json
{
  "message": "SharePoint MCP API Server",
  "endpoints": {
    "GET /api/folders": "List folders (query: parentFolder)",
    "POST /api/upload": "Upload file (body: folderPath, fileName, content)",
    // ... all endpoints
  },
  "authentication": "Bearer token or X-API-Token header required"
}
```

---

## 🚨 Error Responses

All endpoints return consistent error responses:

```json
{
  "error": "Error type",
  "message": "Human readable error message",
  "details": "Additional context (optional)"
}
```

**Common HTTP Status Codes:**
- `200` - Success
- `400` - Bad Request (missing/invalid parameters)
- `401` - Unauthorized (invalid/missing API token)
- `404` - Not Found (item doesn't exist)
- `500` - Internal Server Error
- `503` - Service Unavailable (SharePoint not connected)

---

## 🎯 Rate Limiting

**Current Limits:**
- 100 requests per 15-minute window per IP
- No limits on authenticated requests from trusted tokens

**Headers:**
- `X-RateLimit-Limit` - Request limit per window
- `X-RateLimit-Remaining` - Remaining requests in window
- `X-RateLimit-Reset` - Window reset time (Unix timestamp)