#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { SharePointServer } from "./sharepoint-server.js";
import { logger } from "./utils/logger.js";

async function main() {
  logger.info("Starting SharePoint MCP server...");

  const server = new Server(
    {
      name: "mcp-sharepoint-js",
      version: "0.1.0",
    },
    {
      capabilities: {
        tools: {},
      },
    }
  );

  // Initialize SharePoint server
  const sharepointServer = new SharePointServer(server);
  await sharepointServer.initialize();

  // Connect to stdio transport
  const transport = new StdioServerTransport();
  await server.connect(transport);

  logger.info("SharePoint MCP server is running");
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    logger.error("Failed to start server:", error);
    process.exit(1);
  });
}
