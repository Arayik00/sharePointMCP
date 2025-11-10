#!/usr/bin/env node

/**
 * Root-level server.js for Render deployment compatibility
 * This file redirects to the actual server in src/server.js
 */

console.log('🔄 Root server.js - Redirecting to src/server.js...');

// Import and run the actual server
import('./src/server.js').catch(error => {
  console.error('❌ Failed to start SharePoint MCP Server:', error);
  
  // Try alternative entry point
  console.log('🔄 Trying alternative entry point...');
  import('./src/start.js').catch(fallbackError => {
    console.error('❌ All entry points failed:', fallbackError);
    process.exit(1);
  });
});