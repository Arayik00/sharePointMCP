import express from 'express';
import cors from 'cors';
import { createServer } from 'http';
import { WebSocketServer } from 'ws';
import crypto from 'crypto';
import { SharePointTools } from './sharepoint-tools.js';
import { createSharePointClient } from './sharepoint-client.js';
import { logger } from './utils/logger.js';
import { config } from './utils/config.js';

export class SharePointAPIServer {
  constructor() {
    this.app = express();
    this.server = createServer(this.app);
    this.wss = new WebSocketServer({ server: this.server });
    this.sharePointTools = null;
    this.authTokens = new Set();
    
    // Generate server auth tokens from environment
    this.initializeAuthTokens();
    this.setupMiddleware();
    this.setupRoutes();
    this.setupWebSocket();
  }

  initializeAuthTokens() {
    // Get auth tokens from environment (comma-separated)
    const tokens = process.env.API_AUTH_TOKENS || '';
    if (tokens) {
      tokens.split(',').forEach(token => {
        this.authTokens.add(token.trim());
      });
      logger.info(`Loaded ${this.authTokens.size} auth tokens`);
    } else {
      // Generate a default token if none provided
      const defaultToken = crypto.randomBytes(32).toString('hex');
      this.authTokens.add(defaultToken);
      logger.warn(`No auth tokens provided. Generated default token: ${defaultToken}`);
    }
  }

  setupMiddleware() {
    this.app.use(express.json());
    this.app.use(cors());
    
    // Custom authentication middleware
    this.app.use('/api', (req, res, next) => {
      const authHeader = req.headers.authorization;
      const token = authHeader && authHeader.startsWith('Bearer ') 
        ? authHeader.slice(7) 
        : req.headers['x-api-token'] || req.query.token;

      if (!token || !this.authTokens.has(token)) {
        return res.status(401).json({ 
          error: 'Unauthorized', 
          message: 'Valid API token required' 
        });
      }

      req.authenticated = true;
      next();
    });
  }

  setupRoutes() {
    // Health check (no auth required)
    this.app.get('/health', (req, res) => {
      res.json({ 
        status: 'healthy', 
        timestamp: new Date().toISOString(),
        sharepoint: this.sharePointTools ? 'connected' : 'disconnected'
      });
    });

    // SharePoint API endpoints (all require auth)
    this.app.get('/api/folders', async (req, res) => {
      try {
        const { parentFolder } = req.query;
        const result = await this.sharePointTools.listFolders(parentFolder);
        res.json(result);
      } catch (error) {
        logger.error('API Error - listFolders:', error);
        res.status(500).json({ error: 'Internal server error', message: error.message });
      }
    });

    this.app.get('/api/documents', async (req, res) => {
      try {
        const { folderName } = req.query;
        const result = await this.sharePointTools.listDocuments(folderName);
        res.json(result);
      } catch (error) {
        logger.error('API Error - listDocuments:', error);
        res.status(500).json({ error: 'Internal server error', message: error.message });
      }
    });

    this.app.get('/api/tree', async (req, res) => {
      try {
        const { folderPath, maxDepth } = req.query;
        const result = await this.sharePointTools.getFolderTree(folderPath, maxDepth ? parseInt(maxDepth) : undefined);
        res.json(result);
      } catch (error) {
        logger.error('API Error - getFolderTree:', error);
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
        logger.error('API Error - getDocumentContent:', error);
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
        logger.error('API Error - upload:', error);
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
      
      if (!token || !this.authTokens.has(token)) {
        ws.close(1008, 'Unauthorized');
        return;
      }

      logger.info('WebSocket client connected');
      
      ws.on('message', async (message) => {
        try {
          const request = JSON.parse(message.toString());
          const response = await this.handleWebSocketRequest(request);
          ws.send(JSON.stringify(response));
        } catch (error) {
          logger.error('WebSocket error:', error);
          ws.send(JSON.stringify({ 
            error: 'Internal server error', 
            message: error.message 
          }));
        }
      });

      ws.on('close', () => {
        logger.info('WebSocket client disconnected');
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
      logger.info('Initializing SharePoint client...');
      const client = await createSharePointClient();
      this.sharePointTools = new SharePointTools(client);
      logger.info('SharePoint client initialized successfully');
      return true;
    } catch (error) {
      logger.error('Failed to initialize SharePoint client:', error);
      return false;
    }
  }

  start(port = process.env.PORT || 3000) {
    this.server.listen(port, () => {
      logger.info(`SharePoint API Server listening on port ${port}`);
      logger.info(`API endpoints available at http://localhost:${port}/api`);
      logger.info(`WebSocket available at ws://localhost:${port}?token=YOUR_TOKEN`);
      
      if (this.authTokens.size === 1 && !process.env.API_AUTH_TOKENS) {
        logger.warn(`Use this token for API access: ${Array.from(this.authTokens)[0]}`);
      }
    });
  }

  stop() {
    this.server.close();
    this.wss.close();
  }
}