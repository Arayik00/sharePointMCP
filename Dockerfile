# Use Node.js LTS
FROM node:18-alpine

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install --production

# Copy source code
COPY src/ ./src/

# Create directory for certificate
RUN mkdir -p /app/certs

# Set permissions
RUN chmod 755 /app/certs

# Expose port (if needed for health checks)
EXPOSE 3000

# Set user for security
USER node

# Start the MCP server
CMD ["npm", "start"]