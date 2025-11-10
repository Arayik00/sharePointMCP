#!/bin/bash

# SharePoint MCP Server - systemd Service Configuration
# This script creates a systemd service for production deployment

set -e

echo "🔧 Setting up systemd service for SharePoint MCP Server"
echo "======================================================"

# Configuration
SERVICE_NAME="sharepoint-mcp"
APP_USER="mcp-sharepoint"
APP_DIR="/opt/sharepoint-mcp"
CONFIG_DIR="/opt/sharepoint-mcp/config"
LOG_DIR="/var/log/sharepoint-mcp"
NODE_PATH=$(which node)

# Check if Node.js is installed
if [ ! -f "$NODE_PATH" ]; then
    echo "❌ Node.js not found. Please install Node.js first."
    exit 1
fi

echo "📝 Creating systemd service file..."

# Create systemd service file
sudo tee /etc/systemd/system/${SERVICE_NAME}.service > /dev/null <<EOF
[Unit]
Description=SharePoint MCP Server
Documentation=https://github.com/Arayik00/sharePointMCP
After=network.target

[Service]
Type=simple
User=$APP_USER
Group=$APP_USER
WorkingDirectory=$APP_DIR
Environment=NODE_ENV=production
EnvironmentFile=$CONFIG_DIR/.env.production
ExecStart=$NODE_PATH src/server.js
Restart=always
RestartSec=10
StandardOutput=append:$LOG_DIR/sharepoint-mcp.log
StandardError=append:$LOG_DIR/sharepoint-mcp-error.log

# Security settings
NoNewPrivileges=true
ProtectSystem=strict
ProtectHome=true
ReadWritePaths=$LOG_DIR
PrivateTmp=true
ProtectKernelTunables=true
ProtectKernelModules=true
ProtectControlGroups=true

[Install]
WantedBy=multi-user.target
EOF

echo "✅ Created systemd service: /etc/systemd/system/${SERVICE_NAME}.service"

# Create log rotation configuration
echo "📝 Setting up log rotation..."

sudo tee /etc/logrotate.d/${SERVICE_NAME} > /dev/null <<EOF
$LOG_DIR/*.log {
    daily
    missingok
    rotate 14
    compress
    delaycompress
    notifempty
    create 644 $APP_USER $APP_USER
    postrotate
        systemctl reload $SERVICE_NAME > /dev/null 2>&1 || true
    endscript
}
EOF

echo "✅ Created log rotation: /etc/logrotate.d/${SERVICE_NAME}"

# Reload systemd and enable service
echo "🔄 Reloading systemd daemon..."
sudo systemctl daemon-reload

echo "⚡ Enabling service to start on boot..."
sudo systemctl enable ${SERVICE_NAME}

echo ""
echo "🎉 SystemD service setup complete!"
echo ""
echo "📋 Service management commands:"
echo "Start service:    sudo systemctl start ${SERVICE_NAME}"
echo "Stop service:     sudo systemctl stop ${SERVICE_NAME}"
echo "Restart service:  sudo systemctl restart ${SERVICE_NAME}"
echo "Check status:     sudo systemctl status ${SERVICE_NAME}"
echo "View logs:        sudo journalctl -u ${SERVICE_NAME} -f"
echo "View app logs:    tail -f $LOG_DIR/sharepoint-mcp.log"
echo ""
echo "⚠️  Before starting the service, ensure:"
echo "1. Application files are in: $APP_DIR"
echo "2. Environment file exists: $CONFIG_DIR/.env.production"
echo "3. Certificate file exists with correct permissions"
echo "4. Dependencies installed: cd $APP_DIR && npm install --production"
echo ""
echo "🚀 To start now: sudo systemctl start ${SERVICE_NAME}"