#!/bin/bash

# Deployment troubleshooting script for Render
echo "🔍 SharePoint MCP Server - Deployment Diagnostics"
echo "================================================="

echo "📁 Current directory:"
pwd

echo -e "\n📋 Environment variables:"
echo "NODE_ENV: $NODE_ENV"
echo "SERVER_MODE: $SERVER_MODE" 
echo "PORT: $PORT"

echo -e "\n📂 File structure check:"
ls -la

echo -e "\n📂 Source directory check:"
ls -la src/

echo -e "\n📄 Package.json main field:"
grep '"main"' package.json

echo -e "\n🚀 Starting server with debug info..."
echo "Command: node src/server.js"
echo "Starting in 3 seconds..."
sleep 3

# Start with extra debugging
NODE_ENV=production node src/server.js