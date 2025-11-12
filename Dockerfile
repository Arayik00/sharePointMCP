# Use Node.js LTS (Alpine for smaller image size)
FROM node:18-alpine

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install production dependencies only
RUN npm ci --only=production && npm cache clean --force

# Copy source code and necessary files
COPY src/ ./src/
COPY server.js ./

# Create directory for certificate
RUN mkdir -p /app/certs

# Set permissions
RUN chmod 755 /app/certs

# Expose port for API server mode
EXPOSE 3000

# Add health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=40s --retries=3 \
    CMD node -e "require('http').get('http://localhost:3000/health', (r) => {process.exit(r.statusCode === 200 ? 0 : 1)})"

# Set user for security (run as non-root)
USER node

# Set default environment
ENV NODE_ENV=production

# Start the server (defaults to start.js for robustness)
CMD ["npm", "start"]