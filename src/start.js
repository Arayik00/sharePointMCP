#!/usr/bin/env node

/**
 * Robust entry point for SharePoint MCP Server
 * Handles deployment issues and provides better error reporting
 */

import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { existsSync } from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

console.log('🚀 SharePoint MCP Server - Starting...');
console.log(`📁 Working directory: ${process.cwd()}`);
console.log(`📄 Entry point: ${__filename}`);
console.log(`📂 Source directory: ${__dirname}`);

// Check if required files exist
const requiredFiles = ['config.js', 'api-server.js', 'sharepoint-server.js'];
const missingFiles = requiredFiles.filter(file => !existsSync(join(__dirname, file)));

if (missingFiles.length > 0) {
  console.error('❌ Missing required files:', missingFiles);
  console.error('📂 Available files in src:');
  try {
    const { readdirSync } = await import('fs');
    console.log(readdirSync(__dirname));
  } catch (e) {
    console.error('Could not read directory');
  }
  process.exit(1);
}

// Check environment variables
console.log('🔧 Environment check:');
console.log(`   NODE_ENV: ${process.env.NODE_ENV || 'not set'}`);
console.log(`   SERVER_MODE: ${process.env.SERVER_MODE || 'not set'}`);
console.log(`   PORT: ${process.env.PORT || 'not set'}`);

try {
  // Import and run the main server
  const { default: mainServer } = await import('./server.js');
} catch (error) {
  console.error('❌ Failed to import server.js:', error);
  
  // Try to provide helpful debugging info
  console.log('🔍 Debugging info:');
  console.log('Available modules:');
  try {
    const { readdirSync } = await import('fs');
    const files = readdirSync(__dirname).filter(f => f.endsWith('.js'));
    console.log(files);
  } catch (e) {
    console.error('Could not list files');
  }
  
  process.exit(1);
}