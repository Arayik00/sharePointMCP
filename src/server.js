#!/usr/bin/env node

import { loadConfig, getConfigLoader } from "./config.js";

async function main() {
  try {
    // Load and validate configuration first
    console.log('🚀 Starting SharePoint MCP Server...');
    const config = await loadConfig();
    const mode = config.server.mode;
    
    console.log(`🔧 Starting SharePoint server in ${mode} mode...`);

    if (mode === 'api') {
      // Run as HTTP/WebSocket API server
      const { SharePointAPIServer } = await import('./api-server.js');
      const apiServer = new SharePointAPIServer(config);
      
      const initialized = await apiServer.initialize();
      if (!initialized) {
        console.error('❌ Failed to initialize SharePoint client - starting server anyway for debugging');
        console.error('🔧 Server will be available but SharePoint operations will fail');
        console.error('📋 Check /health endpoint for more information');
      }
      
      const port = config.server.port;
      apiServer.start(port);
      
      // Graceful shutdown
      process.on('SIGINT', () => {
        console.log('🛑 Shutting down API server...');
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

      const sharePointServer = new SharePointServer(config);
      const initialized = await sharePointServer.initialize();

      if (!initialized) {
        console.error("❌ Failed to initialize SharePoint server");
        process.exit(1);
      }

      // Register tools
      sharePointServer.registerTools(server);

      // Start MCP server
      const transport = new StdioServerTransport();
      await server.connect(transport);
      console.log("✅ SharePoint MCP server started successfully");
    }
  } catch (error) {
    console.error("❌ Error starting server:", error);
    process.exit(1);
  }
}

main().catch((error) => {
  console.error("❌ Unhandled error:", error);
  process.exit(1);
});