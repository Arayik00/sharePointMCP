#!/bin/bash

# Certificate Preparation Script for Render Deployment
# This script helps prepare your certificate for safer upload

echo "🔧 Certificate Preparation for Render"
echo "====================================="

CERT_FILE="certificate.pfx"
BASE64_FILE="hearst-sharepoint.b64"

if [ ! -f "$CERT_FILE" ]; then
  echo "❌ Certificate file not found: $CERT_FILE"
  echo "Please place your certificate.pfx file in the project root"
  exit 1
fi

echo "📋 Original certificate:"
ls -la "$CERT_FILE"
echo "📋 Size: $(wc -c < "$CERT_FILE") bytes"
echo "📋 First 16 bytes (hex): $(hexdump -C "$CERT_FILE" | head -1 | cut -d' ' -f2-9)"

# Create base64 version
echo ""
echo "🔄 Creating base64 version..."
base64 "$CERT_FILE" > "$BASE64_FILE"

echo "✅ Base64 certificate created: $BASE64_FILE"
ls -la "$BASE64_FILE"

# Verify the conversion
echo ""
echo "🧪 Verifying conversion..."
TEMP_CERT="temp_verify.pfx"
base64 -d "$BASE64_FILE" > "$TEMP_CERT"

if cmp -s "$CERT_FILE" "$TEMP_CERT"; then
  echo "✅ Base64 conversion verified successfully"
  rm "$TEMP_CERT"
else
  echo "❌ Base64 conversion failed verification"
  rm "$TEMP_CERT"
  exit 1
fi

echo ""
echo "📋 Upload Instructions for Render:"
echo "=================================="
echo "1. Upload the '$BASE64_FILE' file to your Render service"
echo "2. Set these environment variables in Render:"
echo "   USE_BASE64_CERT=true"
echo "   SHP_CERT_BASE64_PATH=/opt/render/project/$BASE64_FILE"
echo "   SHP_CERT_PFX_PATH=/opt/render/project/$CERT_FILE"
echo "3. Update your start command to: node start-with-cert.js"
echo ""
echo "This will automatically convert the base64 file back to binary at startup."