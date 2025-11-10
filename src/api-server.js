import express from 'express';
import cors from 'cors';
import { createServer } from 'http';
import { WebSocketServer } from 'ws';
import crypto from 'crypto';
import { SharePointTools } from './sharepoint-tools.js';
import { createSharePointClient } from './sharepoint-client.js';

export class SharePointAPIServer {
  constructor(config) {
    this.config = config;
    this.app = express();
    this.server = createServer(this.app);
    this.wss = new WebSocketServer({ server: this.server });
    this.sharePointTools = null;
    this.authTokens = new Set();
    
    // Initialize auth tokens from config
    this.initializeAuthTokens();
    this.setupMiddleware();
    this.setupRoutes();
    this.setupWebSocket();
  }

  initializeAuthTokens() {
    // Get auth tokens from config
    const tokens = this.config.api.authTokens || [];
    
    if (tokens.length > 0) {
      tokens.forEach(token => {
        if (token && token.trim()) {
          // Validate token strength (minimum 32 characters)
          if (token.trim().length < 32) {
            console.warn(`⚠️  API token is shorter than recommended 32 characters`);
          }
          this.authTokens.add(token.trim());
        }
      });
      console.log(`🔐 Loaded ${this.authTokens.size} API authentication tokens from environment`);
      console.log(`🛡️  All API endpoints now require valid token authentication`);
    } else {
      // This should not happen in production due to config validation
      const defaultToken = crypto.randomBytes(32).toString('hex');
      this.authTokens.add(defaultToken);
      console.warn(`⚠️  No API_AUTH_TOKENS provided in environment!`);
      console.warn(`🔑 Generated temporary token: ${defaultToken}`);
      console.warn(`🚨 This is insecure! Set API_AUTH_TOKENS in your .env.production file`);
    }
  }

  setupMiddleware() {
    this.app.use(express.json());
    this.app.use(cors());
    
    // Custom authentication middleware - ALL /api/* endpoints require valid API key
    this.app.use('/api', (req, res, next) => {
      const authHeader = req.headers.authorization;
      const token = authHeader && authHeader.startsWith('Bearer ') 
        ? authHeader.slice(7) 
        : req.headers['x-api-token'] || req.query.token;

      // Log authentication attempts for security monitoring
      const clientIP = req.ip || req.connection.remoteAddress;
      const endpoint = req.originalUrl;

      if (!token) {
        console.warn(`🚫 API access denied - No token provided | IP: ${clientIP} | Endpoint: ${endpoint}`);
        return res.status(401).json({ 
          error: 'Unauthorized', 
          message: 'API token required. Provide via Authorization header, X-API-Token header, or ?token= query parameter' 
        });
      }

      if (!this.authTokens.has(token)) {
        console.warn(`🚫 API access denied - Invalid token | IP: ${clientIP} | Endpoint: ${endpoint} | Token: ${token.substring(0, 8)}...`);
        return res.status(401).json({ 
          error: 'Unauthorized', 
          message: 'Invalid API token provided' 
        });
      }

      // Successful authentication
      console.log(`✅ API access granted | IP: ${clientIP} | Endpoint: ${endpoint}`);
      req.authenticated = true;
      req.apiToken = token;
      next();
    });
  }

  setupRoutes() {
    // Health check (no auth required)
    this.app.get('/health', (req, res) => {
      res.json({ 
        status: 'healthy', 
        timestamp: new Date().toISOString(),
        sharepoint: this.sharePointTools ? 'connected' : 'disconnected',
        authentication: 'required',
        endpoints: [
          'GET /api/auth/validate - Validate API token',
          'GET /api/folders - List SharePoint folders', 
          'GET /api/documents - List SharePoint documents',
          'GET /api/tree - Get folder tree structure',
          'GET /api/document/:path/content - Get document content',
          'POST /api/upload - Upload document'
        ]
      });
    });

    // Token validation endpoint (requires auth)
    this.app.get('/api/auth/validate', (req, res) => {
      // If we get here, auth middleware already validated the token
      res.json({
        valid: true,
        message: 'API token is valid',
        authenticated: true,
        token_preview: req.apiToken.substring(0, 8) + '...',
        timestamp: new Date().toISOString()
      });
    });

    // SharePoint API endpoints (all require auth)
    this.app.get('/api/folders', async (req, res) => {
      try {
        const { parentFolder } = req.query;
        const result = await this.sharePointTools.listFolders(parentFolder);
        res.json(result);
      } catch (error) {
        console.error('❌ API Error - listFolders:', error);
        res.status(500).json({ error: 'Internal server error', message: error.message });
      }
    });

    this.app.get('/api/documents', async (req, res) => {
      try {
        const { folderName } = req.query;
        const result = await this.sharePointTools.listDocuments(folderName);
        res.json(result);
      } catch (error) {
        console.error('❌ API Error - listDocuments:', error);
        res.status(500).json({ error: 'Internal server error', message: error.message });
      }
    });

    this.app.get('/api/tree', async (req, res) => {
      try {
        const { folderPath, maxDepth } = req.query;
        const result = await this.sharePointTools.getFolderTree(folderPath, maxDepth ? parseInt(maxDepth) : undefined);
        res.json(result);
      } catch (error) {
        console.error('❌ API Error - getFolderTree:', error);
        res.status(500).json({ error: 'Internal server error', message: error.message });
      }
    });

    this.app.get('/api/document/:path/content', async (req, res) => {
      try {
        const documentPath = decodeURIComponent(req.params.path);
        const result = await this.sharePointTools.getDocumentContent(documentPath);
        
        if (result.success) {
          res.json(result);
        } else {
          res.status(404).json(result);
        }
      } catch (error) {
        console.error('❌ API Error - getDocumentContent:', error);
        res.status(500).json({ error: 'Internal server error', message: error.message });
      }
    });

    // Upload endpoint (if needed later)
    this.app.post('/api/upload', async (req, res) => {
      try {
        const { folderPath, fileName, content } = req.body;
        // Implementation would go here
        res.json({ message: 'Upload endpoint - not implemented yet' });
      } catch (error) {
        console.error('❌ API Error - upload:', error);
        res.status(500).json({ error: 'Internal server error', message: error.message });
      }
    });

    // List available API endpoints
    this.app.get('/api', (req, res) => {
      res.json({
        message: 'SharePoint MCP API Server',
        endpoints: {
          'GET /api/folders': 'List folders (query: parentFolder)',
          'GET /api/documents': 'List documents (query: folderName)',
          'GET /api/tree': 'Get folder tree (query: folderPath, maxDepth)',
          'GET /api/document/:path/content': 'Get document content',
          'POST /api/upload': 'Upload file (body: folderPath, fileName, content)'
        },
        authentication: 'Bearer token or X-API-Token header required'
      });
    });
  }

  setupWebSocket() {
    this.wss.on('connection', (ws, req) => {
      // Check auth for WebSocket connections
      const url = new URL(req.url, `http://${req.headers.host}`);
      const token = url.searchParams.get('token');
      const clientIP = req.socket.remoteAddress;
      
      if (!token) {
        console.warn(`🚫 WebSocket connection denied - No token provided | IP: ${clientIP}`);
        ws.close(1008, 'Unauthorized - Token required');
        return;
      }

      if (!this.authTokens.has(token)) {
        console.warn(`🚫 WebSocket connection denied - Invalid token | IP: ${clientIP} | Token: ${token.substring(0, 8)}...`);
        ws.close(1008, 'Unauthorized - Invalid token');
        return;
      }

      console.log(`🔌 WebSocket client connected with valid authentication | IP: ${clientIP}`);
      
      ws.on('message', async (message) => {
        try {
          const request = JSON.parse(message.toString());
          const response = await this.handleWebSocketRequest(request);
          ws.send(JSON.stringify(response));
        } catch (error) {
          console.error('❌ WebSocket error:', error);
          ws.send(JSON.stringify({ 
            error: 'Internal server error', 
            message: error.message 
          }));
        }
      });

      ws.on('close', () => {
        console.log('🔌 WebSocket client disconnected');
      });
    });
  }

  async handleWebSocketRequest(request) {
    const { action, params = {} } = request;

    switch (action) {
      case 'listFolders':
        return await this.sharePointTools.listFolders(params.parentFolder);
      
      case 'listDocuments':
        return await this.sharePointTools.listDocuments(params.folderName);
      
      case 'getFolderTree':
        return await this.sharePointTools.getFolderTree(params.folderPath, params.maxDepth);
      
      case 'getDocumentContent':
        return await this.sharePointTools.getDocumentContent(params.documentPath);
      
      default:
        return { error: 'Unknown action', action };
    }
  }

  async initialize() {
    try {
      // Initialize SharePoint client with server-side credentials
      console.log('🔧 Initializing SharePoint client...');
      const client = await createSharePointClient(this.config);
      this.sharePointTools = new SharePointTools(client);
      console.log('✅ SharePoint client initialized successfully');
      return true;
    } catch (error) {
      console.error('❌ Failed to initialize SharePoint client:', error);
      return false;
    }
  }

  start(port = process.env.PORT || 3000) {
    this.server.listen(port, () => {
      console.log(`🚀 SharePoint API Server listening on port ${port}`);
      console.log(`📡 API endpoints available at http://localhost:${port}/api`);
      console.log(`🔌 WebSocket available at ws://localhost:${port}?token=YOUR_TOKEN`);
      
      if (this.authTokens.size === 1) {
        console.warn(`🔑 Use this token for API access: ${Array.from(this.authTokens)[0]}`);
      }
    });
  }

  stop() {
    this.server.close();
    this.wss.close();
  }
}