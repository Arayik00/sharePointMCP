import { logger } from "./utils/logger.js";
import { config } from "./utils/config.js";

export class SharePointTools {
  constructor(client) {
    this.client = client;
  }

  getDrivePath(folder = "") {
    // When config.docLibrary is empty, we're working directly in the Documents drive root
    if (!config.docLibrary || config.docLibrary.trim() === "") {
      return folder.replace(/^\/+|\/+$/g, ""); // Remove leading/trailing slashes
    }

    // For non-empty config.docLibrary, combine with folder
    const docLibParts = config.docLibrary.replace(/^\/+|\/+$/g, "").split("/");
    const basePath = docLibParts.join("/");

    if (folder) {
      const cleanFolder = folder.replace(/^\/+|\/+$/g, "");
      return basePath ? `${basePath}/${cleanFolder}` : cleanFolder;
    }
    return basePath;
  }

  formatFileInfo(item) {
    const isFolder = "folder" in item;
    return {
      name: item.name || "Unknown",
      type: isFolder ? "folder" : "file",
      size: item.size || 0,
      webUrl: item.webUrl || "",
      lastModified: item.lastModifiedDateTime || "",
      mimeType: !isFolder ? item.file?.mimeType || "" : "",
    };
  }

  async listFolders(parentFolder) {
    try {
      const folderPath = this.getDrivePath(parentFolder || "");
      logger.info(
        `Listing folders in path: '${folderPath}' (parent_folder: '${parentFolder}')`
      );

      const driveId = this.client.drives[0]?.id;
      if (!driveId) {
        throw new Error("No drives available");
      }

      const folders = await this.client.getFolders(driveId, folderPath);
      const formattedFolders = folders.map((item) => this.formatFileInfo(item));

      logger.info(`Found ${formattedFolders.length} folders`);
      return {
        success: true,
        items: formattedFolders,
        count: formattedFolders.length,
      };
    } catch (error) {
      logger.error("Error listing folders:", error);
      return {
        success: false,
        message: `Failed to list folders: ${error}`,
      };
    }
  }

  async listDocuments(folderName) {
    try {
      const folderPath = this.getDrivePath(folderName);
      logger.info(
        `Listing documents in path: '${folderPath}' (folder_name: '${folderName}')`
      );

      const driveId = this.client.drives[0]?.id;
      if (!driveId) {
        throw new Error("No drives available");
      }

      const files = await this.client.getFiles(driveId, folderPath);
      const documents = files.map((item) => this.formatFileInfo(item));

      logger.info(`Found ${documents.length} documents`);
      return {
        success: true,
        items: documents,
        count: documents.length,
      };
    } catch (error) {
      logger.error("Error listing documents:", error);
      return {
        success: false,
        message: `Failed to list documents: ${error}`,
      };
    }
  }

  async getDocumentContent(folderName, fileName) {
    try {
      const filePath = `${this.getDrivePath(folderName)}/${fileName}`.replace(
        /\/+/g,
        "/"
      );
      logger.info(`Getting content for: ${filePath}`);

      const content = await this.client.getFileContent("Documents", filePath);

      if (!content) {
        return {
          success: false,
          message: `File '${fileName}' not found`,
        };
      }

      // Try to decode as text for common text files
      try {
        const textContent = content.toString("utf-8");
        return {
          success: true,
          content: textContent,
          file: { name: fileName, path: filePath },
          type: "text",
        };
      } catch (error) {
        // Return as base64 for binary files
        return {
          success: true,
          content: content.toString("base64"),
          file: { name: fileName, path: filePath },
          type: "binary",
        };
      }
    } catch (error) {
      logger.error("Error getting document content:", error);
      return {
        success: false,
        message: `Failed to get content: ${error}`,
      };
    }
  }

  async getFolderTree(parentFolder, maxDepth = 3) {
    try {
      const folderPath = this.getDrivePath(parentFolder || "");
      logger.info(`Getting folder tree for: ${folderPath}`);

      const buildTree = async (path, depth) => {
        if (depth <= 0) {
          return { items: [], truncated: true };
        }

        try {
          const driveId = this.client.drives[0]?.id;
          if (!driveId) {
            throw new Error("No drives available");
          }

          // Get both files and folders
          const [folders, files] = await Promise.all([
            this.client.getFolders(driveId, path),
            this.client.getFiles(driveId, path)
          ]);
          
          const items = [...folders, ...files];
          const treeItems = [];

          for (const item of items) {
            const itemInfo = this.formatFileInfo(item);

            if (itemInfo.type === "folder") {
              const childPath = `${path}/${itemInfo.name}`.replace(/\/+/g, "/");
              itemInfo.children = await buildTree(childPath, depth - 1);
            }

            treeItems.push(itemInfo);
          }

          return { items: treeItems, truncated: false };
        } catch (error) {
          logger.error(`Error building tree for ${path}:`, error);
          return { items: [], error: String(error) };
        }
      };

      const tree = await buildTree(folderPath, maxDepth);
      return {
        success: true,
        folder: parentFolder || "root",
        tree,
      };
    } catch (error) {
      logger.error("Error getting folder tree:", error);
      return {
        success: false,
        message: `Failed to get tree: ${error}`,
      };
    }
  }
}