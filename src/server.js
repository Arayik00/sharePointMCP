#!/usr/bin/env node

import { logger } from "./utils/logger.js";

async function main() {
  const mode = process.env.SERVER_MODE || process.argv[2] || 'mcp';

  logger.info(`Starting SharePoint server in ${mode} mode...`);

  try {
    if (mode === 'api') {
      // Run as HTTP/WebSocket API server
      const { SharePointAPIServer } = await import('./api-server.js');
      const apiServer = new SharePointAPIServer();
      
      const initialized = await apiServer.initialize();
      if (!initialized) {
        logger.error('Failed to initialize SharePoint client');
        process.exit(1);
      }
      
      const port = process.env.PORT || 3000;
      apiServer.start(port);
      
      // Graceful shutdown
      process.on('SIGINT', () => {
        logger.info('Shutting down API server...');
        apiServer.stop();
        process.exit(0);
      });
      
    } else {
      // Run as MCP server (default)
      const { Server } = await import("@modelcontextprotocol/sdk/server/index.js");
      const { StdioServerTransport } = await import("@modelcontextprotocol/sdk/server/stdio.js");
      const { SharePointServer } = await import("./sharepoint-server.js");

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

      const sharePointServer = new SharePointServer();
      const initialized = await sharePointServer.initialize();

      if (!initialized) {
        logger.error("Failed to initialize SharePoint server");
        process.exit(1);
      }

      // Register tools
      sharePointServer.registerTools(server);

      // Start MCP server
      const transport = new StdioServerTransport();
      await server.connect(transport);
      logger.info("SharePoint MCP server started successfully");
    }
  } catch (error) {
    logger.error("Error starting server:", error);
    process.exit(1);
  }
}

main().catch((error) => {
  logger.error("Unhandled error:", error);
  process.exit(1);
});