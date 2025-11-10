#!/usr/bin/env node

/**
 * SharePoint MCP Server - API Client Wrapper
 * This creates MCP tools that connect to your API server instead of directly to SharePoint
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

// API Configuration from environment
const API_CONFIG = {
  baseURL: process.env.SHAREPOINT_API_URL || 'https://your-server.onrender.com',
  apiToken: process.env.SHAREPOINT_API_TOKEN || 'your-64-character-api-token-here'
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

  try {
    const response = await fetch(url, config);
    
    if (!response.ok) {
      if (response.status === 401) {
        throw new Error('Authentication failed - check API token');
      }
      throw new Error(`API request failed: ${response.status} ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    console.error(`SharePoint API Error:`, error.message);
    throw error;
  }
}

// Create MCP Server
const server = new Server(
  {
    name: "sharepoint-api-client",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// MCP Tool: List SharePoint Folders
server.setRequestHandler("tools/call", async (request) => {
  const { name, arguments: args } = request.params;

  switch (name) {
    case "List_SharePoint_Folders":
      try {
        const params = args.parent_folder ? `?parentFolder=${encodeURIComponent(args.parent_folder)}` : '';
        const result = await apiRequest(`/api/folders${params}`);
        
        return {
          content: [{
            type: "text",
            text: `Found ${result.count} folders:\n${JSON.stringify(result.items, null, 2)}`
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: "text", 
            text: `Error listing folders: ${error.message}`
          }],
          isError: true
        };
      }

    case "List_SharePoint_Documents":
      try {
        const params = args.folder_name ? `?folderName=${encodeURIComponent(args.folder_name)}` : '';
        const result = await apiRequest(`/api/documents${params}`);
        
        return {
          content: [{
            type: "text",
            text: `Found ${result.count} documents:\n${JSON.stringify(result.items, null, 2)}`
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `Error listing documents: ${error.message}`
          }],
          isError: true
        };
      }

    case "Get_SharePoint_Tree":
      try {
        const params = new URLSearchParams();
        if (args.parent_folder) params.append('folderPath', args.parent_folder);
        if (args.max_depth) params.append('maxDepth', args.max_depth);
        
        const queryString = params.toString() ? `?${params.toString()}` : '';
        const result = await apiRequest(`/api/tree${queryString}`);
        
        return {
          content: [{
            type: "text",
            text: `Folder tree structure:\n${JSON.stringify(result, null, 2)}`
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `Error getting folder tree: ${error.message}`
          }],
          isError: true
        };
      }

    case "Get_SharePoint_Document_Content":
      try {
        const result = await apiRequest(`/api/document/${encodeURIComponent(args.document_path)}/content`);
        
        return {
          content: [{
            type: "text",
            text: `Document content:\n${result.content}`
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `Error getting document content: ${error.message}`
          }],
          isError: true
        };
      }

    case "Upload_SharePoint_Document":
      try {
        const result = await apiRequest('/api/upload', {
          method: 'POST',
          body: JSON.stringify({
            fileName: args.file_name,
            content: args.content,
            folderPath: args.folder_path || '',
            overwrite: args.overwrite || false
          })
        });
        
        return {
          content: [{
            type: "text",
            text: `Document uploaded successfully:\n${JSON.stringify(result, null, 2)}`
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `Error uploading document: ${error.message}`
          }],
          isError: true
        };
      }

    case "Validate_API_Connection":
      try {
        const result = await apiRequest('/api/auth/validate');
        
        return {
          content: [{
            type: "text",
            text: `API connection valid:\n${JSON.stringify(result, null, 2)}`
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `API connection failed: ${error.message}`
          }],
          isError: true
        };
      }

    default:
      throw new Error(`Unknown tool: ${name}`);
  }
});

// List available tools
server.setRequestHandler("tools/list", async () => {
  return {
    tools: [
      {
        name: "List_SharePoint_Folders",
        description: "List folders in SharePoint via API server",
        inputSchema: {
          type: "object",
          properties: {
            parent_folder: {
              type: "string",
              description: "Parent folder path (optional)"
            }
          }
        }
      },
      {
        name: "List_SharePoint_Documents", 
        description: "List documents in a SharePoint folder via API server",
        inputSchema: {
          type: "object",
          properties: {
            folder_name: {
              type: "string",
              description: "Folder name to list documents from"
            }
          },
          required: ["folder_name"]
        }
      },
      {
        name: "Get_SharePoint_Tree",
        description: "Get SharePoint folder tree structure via API server",
        inputSchema: {
          type: "object",
          properties: {
            parent_folder: {
              type: "string", 
              description: "Parent folder path (optional)"
            },
            max_depth: {
              type: "number",
              description: "Maximum tree depth (default: 3)"
            }
          }
        }
      },
      {
        name: "Get_SharePoint_Document_Content",
        description: "Get content of a SharePoint document via API server",
        inputSchema: {
          type: "object",
          properties: {
            document_path: {
              type: "string",
              description: "Full path to the document"
            }
          },
          required: ["document_path"]
        }
      },
      {
        name: "Upload_SharePoint_Document",
        description: "Upload a document to SharePoint via API server", 
        inputSchema: {
          type: "object",
          properties: {
            file_name: {
              type: "string",
              description: "Name of the file to upload"
            },
            content: {
              type: "string", 
              description: "Content of the file"
            },
            folder_path: {
              type: "string",
              description: "Folder path to upload to (optional)"
            },
            overwrite: {
              type: "boolean",
              description: "Whether to overwrite existing file (default: false)"
            }
          },
          required: ["file_name", "content"]
        }
      },
      {
        name: "Validate_API_Connection",
        description: "Test API server connection and authentication",
        inputSchema: {
          type: "object",
          properties: {}
        }
      }
    ]
  };
});

async function main() {
  console.log('🚀 Starting SharePoint API MCP Client...');
  console.log(`📡 API Server: ${API_CONFIG.baseURL}`);
  console.log(`🔑 Using API Token: ${API_CONFIG.apiToken.substring(0, 8)}...`);

  const transport = new StdioServerTransport();
  await server.connect(transport);
  
  console.log('✅ SharePoint API MCP Client connected successfully');
}

main().catch((error) => {
  console.error('❌ Failed to start SharePoint API MCP Client:', error);
  process.exit(1);
});