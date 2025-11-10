#!/bin/bash

# SharePoint MCP Server - Secure Deployment Script
# This script sets up the server environment with proper security measures

set -e  # Exit on any error

echo "🔐 SharePoint MCP Server - Secure Deployment Setup"
echo "=================================================="

# Check if running as root
if [ "$EUID" -eq 0 ]; then
    echo "❌ Do not run this script as root for security reasons"
    exit 1
fi

# Configuration
APP_USER="mcp-sharepoint"
APP_DIR="/opt/sharepoint-mcp"
CERT_DIR="/opt/sharepoint-mcp/certs"
CONFIG_DIR="/opt/sharepoint-mcp/config"
LOG_DIR="/var/log/sharepoint-mcp"

echo "📁 Setting up directory structure..."

# Create application user (if it doesn't exist)
if ! id "$APP_USER" &>/dev/null; then
    sudo useradd -r -s /bin/false -d "$APP_DIR" "$APP_USER"
    echo "✅ Created application user: $APP_USER"
fi

# Create directories
sudo mkdir -p "$APP_DIR" "$CERT_DIR" "$CONFIG_DIR" "$LOG_DIR"

# Set directory ownership and permissions
sudo chown -R "$APP_USER:$APP_USER" "$APP_DIR"
sudo chown -R "$APP_USER:$APP_USER" "$LOG_DIR"

# Secure directory permissions
sudo chmod 750 "$APP_DIR"
sudo chmod 700 "$CERT_DIR"      # Only app user can access certificates
sudo chmod 750 "$CONFIG_DIR"
sudo chmod 755 "$LOG_DIR"

echo "✅ Directory structure created with secure permissions"

echo "📋 Next steps:"
echo "1. Copy your application files to: $APP_DIR"
echo "2. Copy your certificate.pfx to: $CERT_DIR/"
echo "3. Copy your .env.production to: $CONFIG_DIR/"
echo "4. Set secure file permissions:"
echo "   sudo chmod 600 $CERT_DIR/certificate.pfx"
echo "   sudo chmod 600 $CONFIG_DIR/.env.production"
echo "5. Configure systemd service (see deploy-systemd.sh)"

echo ""
echo "🔒 Security checklist:"
echo "- Certificate file permissions: 600 (owner read/write only)"
echo "- Environment file permissions: 600 (owner read/write only)" 
echo "- Application runs as non-root user: $APP_USER"
echo "- Log directory accessible for monitoring"
echo "- Firewall configured for required ports only"

echo ""
echo "📁 Directory layout:"
echo "$APP_DIR/                   # Application files"
echo "├── src/                   # Source code"
echo "├── package.json           # Dependencies"
echo "├── node_modules/          # Installed packages"
echo "$CERT_DIR/                 # Certificates (secure)"
echo "└── certificate.pfx        # Your SharePoint certificate"
echo "$CONFIG_DIR/               # Configuration (secure)"
echo "└── .env.production        # Environment variables"
echo "$LOG_DIR/                  # Application logs"
echo "└── sharepoint-mcp.log     # Main log file"

echo ""
echo "✅ Secure deployment structure ready!"
echo "⚠️  Remember to configure your firewall and SSL certificates for production use"