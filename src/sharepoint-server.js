import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { createSharePointClient } from "./sharepoint-client.js";
import { SharePointTools } from "./sharepoint-tools.js";
import { logger } from "./utils/logger.js";

export class SharePointServer {
  constructor(server) {
    this.server = server;
    this.tools = null;
  }

  async initialize() {
    try {
      // Initialize SharePoint client
      const client = await createSharePointClient();
      if (!client) {
        throw new Error("Failed to create SharePoint client");
      }

      this.tools = new SharePointTools(client);

      // Set up MCP handlers
      this.setupHandlers();

      logger.info("SharePoint server initialized successfully");
    } catch (error) {
      logger.error("Failed to initialize SharePoint server:", error);
      throw error;
    }
  }

  setupHandlers() {
    // List tools handler
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: [
          {
            name: "List_SharePoint_Folders",
            description:
              "List folders in the specified SharePoint directory or root if not specified",
            inputSchema: {
              type: "object",
              properties: {
                parent_folder: {
                  type: "string",
                  description: "Parent folder path (optional, defaults to root)",
                },
              },
            },
          },
          {
            name: "List_SharePoint_Documents",
            description: "List all documents in a specified SharePoint folder",
            inputSchema: {
              type: "object",
              properties: {
                folder_name: {
                  type: "string",
                  description: "Folder name to list documents from",
                },
              },
              required: ["folder_name"],
            },
          },
          {
            name: "Get_SharePoint_Tree",
            description: "Get a recursive tree view of a SharePoint folder",
            inputSchema: {
              type: "object",
              properties: {
                parent_folder: {
                  type: "string",
                  description: "Parent folder path (optional, defaults to root)",
                },
                max_depth: {
                  type: "number",
                  description: "Maximum depth for recursion (default: 3)",
                  default: 3,
                },
              },
            },
          },
          {
            name: "Get_Document_Content",
            description: "Get content of a document in SharePoint",
            inputSchema: {
              type: "object",
              properties: {
                folder_name: {
                  type: "string",
                  description: "Folder containing the document",
                },
                file_name: {
                  type: "string",
                  description: "Name of the file to get content from",
                },
              },
              required: ["folder_name", "file_name"],
            },
          },
          {
            name: "Create_Folder",
            description:
              "Create a new folder in the specified directory or root if not specified",
            inputSchema: {
              type: "object",
              properties: {
                folder_name: {
                  type: "string",
                  description: "Name of the folder to create",
                },
                parent_folder: {
                  type: "string",
                  description: "Parent folder path (optional, defaults to root)",
                },
              },
              required: ["folder_name"],
            },
          },
          {
            name: "Upload_Document",
            description: "Upload a new file to SharePoint",
            inputSchema: {
              type: "object",
              properties: {
                folder_name: {
                  type: "string",
                  description: "Target folder name",
                },
                file_name: {
                  type: "string",
                  description: "Name of the file to upload",
                },
                content: {
                  type: "string",
                  description: "File content (text or base64)",
                },
                is_base64: {
                  type: "boolean",
                  description: "Whether the content is base64 encoded",
                  default: false,
                },
              },
              required: ["folder_name", "file_name", "content"],
            },
          },
          {
            name: "Upload_Document_From_Path",
            description: "Upload a file directly from a file path to SharePoint",
            inputSchema: {
              type: "object",
              properties: {
                folder_name: {
                  type: "string",
                  description: "Target folder name",
                },
                file_path: {
                  type: "string",
                  description: "Local file path to upload",
                },
                new_file_name: {
                  type: "string",
                  description:
                    "New file name (optional, uses original name if not provided)",
                },
              },
              required: ["folder_name", "file_path"],
            },
          },
          {
            name: "Update_Document",
            description: "Update an existing document in SharePoint",
            inputSchema: {
              type: "object",
              properties: {
                folder_name: {
                  type: "string",
                  description: "Folder containing the document",
                },
                file_name: {
                  type: "string",
                  description: "Name of the file to update",
                },
                content: {
                  type: "string",
                  description: "New file content (text or base64)",
                },
                is_base64: {
                  type: "boolean",
                  description: "Whether the content is base64 encoded",
                  default: false,
                },
              },
              required: ["folder_name", "file_name", "content"],
            },
          },
          {
            name: "Delete_Document",
            description: "Delete a document from SharePoint",
            inputSchema: {
              type: "object",
              properties: {
                folder_name: {
                  type: "string",
                  description: "Folder containing the document",
                },
                file_name: {
                  type: "string",
                  description: "Name of the file to delete",
                },
              },
              required: ["folder_name", "file_name"],
            },
          },
          {
            name: "Delete_Folder",
            description: "Delete an empty folder from SharePoint",
            inputSchema: {
              type: "object",
              properties: {
                folder_path: {
                  type: "string",
                  description: "Path of the folder to delete",
                },
              },
              required: ["folder_path"],
            },
          },
          {
            name: "Download_Document",
            description: "Download a document to local filesystem",
            inputSchema: {
              type: "object",
              properties: {
                folder_name: {
                  type: "string",
                  description: "Folder containing the document",
                },
                file_name: {
                  type: "string",
                  description: "Name of the file to download",
                },
                local_path: {
                  type: "string",
                  description: "Local path where to save the file",
                },
              },
              required: ["folder_name", "file_name", "local_path"],
            },
          },
        ],
      };
    });

    // Call tool handler
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      if (!this.tools) {
        throw new Error("SharePoint tools not initialized");
      }

      const { name, arguments: args } = request.params;

      try {
        let result;

        switch (name) {
          case "List_SharePoint_Folders":
            result = await this.tools.listFolders(args?.parent_folder);
            break;

          case "List_SharePoint_Documents":
            result = await this.tools.listDocuments(args?.folder_name);
            break;

          case "Get_SharePoint_Tree":
            result = await this.tools.getFolderTree(
              args?.parent_folder,
              args?.max_depth || 3
            );
            break;

          case "Get_Document_Content":
            result = await this.tools.getDocumentContent(
              args?.folder_name,
              args?.file_name
            );
            break;

          case "Create_Folder":
            result = await this.tools.createFolder(
              args?.folder_name,
              args?.parent_folder
            );
            break;

          case "Upload_Document":
            result = await this.tools.uploadDocument(
              args?.folder_name,
              args?.file_name,
              args?.content,
              args?.is_base64 || false
            );
            break;

          case "Upload_Document_From_Path":
            result = await this.tools.uploadDocumentFromPath(
              args?.folder_name,
              args?.file_path,
              args?.new_file_name
            );
            break;

          case "Update_Document":
            result = await this.tools.updateDocument(
              args?.folder_name,
              args?.file_name,
              args?.content,
              args?.is_base64 || false
            );
            break;

          case "Delete_Document":
            result = await this.tools.deleteDocument(
              args?.folder_name,
              args?.file_name
            );
            break;

          case "Delete_Folder":
            result = await this.tools.deleteFolder(args?.folder_path);
            break;

          case "Download_Document":
            result = await this.tools.downloadDocument(
              args?.folder_name,
              args?.file_name,
              args?.local_path
            );
            break;

          default:
            throw new Error(`Unknown tool: ${name}`);
        }

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        logger.error(`Error executing tool ${name}:`, error);
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  success: false,
                  message: `Tool execution failed: ${error}`,
                },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }
    });
  }
}