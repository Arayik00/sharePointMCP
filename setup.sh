#!/bin/bash

echo "🚀 Setting up SharePoint MCP Server (JavaScript)"

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install Node.js 18+ first."
    exit 1
fi

# Check Node.js version
NODE_VERSION=$(node -v | cut -d'v' -f2 | cut -d'.' -f1)
if [ "$NODE_VERSION" -lt 18 ]; then
    echo "❌ Node.js version 18+ is required. Current version: $(node -v)"
    exit 1
fi

echo "✅ Node.js $(node -v) detected"

# Install dependencies
echo "📦 Installing dependencies..."
npm install

if [ $? -ne 0 ]; then
    echo "❌ Failed to install dependencies"
    exit 1
fi

echo "✅ Dependencies installed"

# Check if .env exists
if [ ! -f ".env" ]; then
    echo "⚠️ .env file not found. Copying from .env.example..."
    cp .env.example .env
    echo "📝 Please edit .env with your SharePoint configuration"
fi

# Check if certificate exists
if [ ! -f "certificate.pfx" ]; then
    echo "⚠️ certificate.pfx not found. Please place your certificate file in the project root."
fi

echo ""
echo "🎉 Setup complete!"
echo ""
echo "Next steps:"
echo "1. Edit .env with your SharePoint configuration"
echo "2. Place your certificate.pfx file in the project root"
echo "3. Run 'npm start' to start the server"
echo "4. Or run 'npm run dev' for development mode"
echo ""
echo "To test the setup, run: node test.js"