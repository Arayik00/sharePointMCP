import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/**
 * Secure Configuration Loader for SharePoint MCP Server
 * Loads credentials from server environment variables and validates them
 */
class ConfigLoader {
  constructor() {
    this.config = null;
    this.isLoaded = false;
  }

  /**
   * Auto-detect server mode based on environment
   */
  detectServerMode() {
    // Check for common cloud deployment environments
    if (process.env.RENDER || process.env.RENDER_SERVICE_NAME) {
      console.log('🌐 Detected Render environment - using API mode');
      return 'api';
    }
    if (process.env.HEROKU_APP_NAME || process.env.DYNO) {
      console.log('🌐 Detected Heroku environment - using API mode');
      return 'api';
    }
    if (process.env.VERCEL || process.env.NOW_REGION) {
      console.log('🌐 Detected Vercel environment - using API mode');
      return 'api';
    }
    if (process.env.NODE_ENV === 'production' && process.env.PORT) {
      console.log('🌐 Detected production environment with PORT - using API mode');
      return 'api';
    }
    
    console.log('🖥️  Default to MCP mode for local development');
    return 'mcp';
  }

  /**
   * Load and validate configuration from environment
   * @returns {Object} Configuration object
   */
  async load() {
    if (this.isLoaded && this.config) {
      return this.config;
    }

    console.log('🔧 Loading server configuration...');

    // Load configuration from environment variables
    this.config = {
      sharepoint: {
        appId: this.getRequiredEnv('SHP_ID_APP'),
        tenantId: this.getRequiredEnv('SHP_TENANT_ID'),
        siteUrl: this.getRequiredEnv('SHP_SITE_URL'),
        certPath: this.getRequiredEnv('SHP_CERT_PFX_PATH'),
        certPassword: this.getRequiredEnv('SHP_CERT_PFX_PASSWORD'),
        docLibrary: this.getOptionalEnv('SHP_DOC_LIBRARY', '')
      },
      server: {
        mode: process.env.SERVER_MODE || this.detectServerMode(),
        port: parseInt(process.env.PORT || '3000'),
        host: process.env.HOST || 'localhost',
        logLevel: process.env.LOG_LEVEL || 'info'
      },
      api: {
        authTokens: this.getOptionalEnv('API_AUTH_TOKENS', '').split(',').filter(t => t.trim()),
        corsOrigins: this.getOptionalEnv('CORS_ORIGINS', '*').split(',').map(o => o.trim()),
        rateLimitWindow: parseInt(process.env.RATE_LIMIT_WINDOW || '900000'), // 15 minutes
        rateLimitMax: parseInt(process.env.RATE_LIMIT_MAX || '100')
      }
    };

    // Validate configuration
    await this.validate();
    
    this.isLoaded = true;
    console.log('✅ Configuration loaded successfully');
    return this.config;
  }

  /**
   * Get required environment variable
   * @param {string} name - Environment variable name
   * @returns {string} Variable value
   * @throws {Error} If variable is not set
   */
  getRequiredEnv(name) {
    const value = process.env[name];
    if (!value) {
      throw new Error(`❌ Required environment variable ${name} is not set`);
    }
    return value;
  }

  /**
   * Get optional environment variable with default
   * @param {string} name - Environment variable name
   * @param {string} defaultValue - Default value if not set
   * @returns {string} Variable value or default
   */
  getOptionalEnv(name, defaultValue = '') {
    return process.env[name] || defaultValue;
  }

  /**
   * Validate configuration
   * @throws {Error} If configuration is invalid
   */
  async validate() {
    console.log('🔍 Validating configuration...');

    // Validate SharePoint configuration
    await this.validateSharePointConfig();
    
    // Validate server configuration
    this.validateServerConfig();
    
    // Validate API configuration
    this.validateApiConfig();

    console.log('✅ Configuration validation passed');
  }

  /**
   * Validate SharePoint-specific configuration
   */
  async validateSharePointConfig() {
    const { sharepoint } = this.config;

    // Validate App ID format (GUID)
    const guidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
    if (!guidRegex.test(sharepoint.appId)) {
      throw new Error('❌ SHP_ID_APP must be a valid GUID format');
    }

    // Validate Tenant ID format (GUID)
    if (!guidRegex.test(sharepoint.tenantId)) {
      throw new Error('❌ SHP_TENANT_ID must be a valid GUID format');
    }

    // Validate Site URL format
    const urlRegex = /^https:\/\/[a-zA-Z0-9-]+\.sharepoint\.com\/sites\/[a-zA-Z0-9-]+$/;
    if (!urlRegex.test(sharepoint.siteUrl)) {
      throw new Error('❌ SHP_SITE_URL must be a valid SharePoint site URL');
    }

    // Validate certificate file exists and is accessible
    await this.validateCertificate();
  }

  /**
   * Validate certificate file
   */
  async validateCertificate() {
    const { certPath, certPassword } = this.config.sharepoint;

    // For cloud deployments, provide more flexible validation
    const isCloudDeployment = process.env.RENDER || process.env.HEROKU_APP_NAME || 
                             process.env.VERCEL || process.env.NODE_ENV === 'production';

    // Check if certificate path is absolute
    if (!path.isAbsolute(certPath)) {
      if (isCloudDeployment) {
        console.warn('⚠️  Certificate path is not absolute, attempting to resolve relative to project root');
        // Try to make it absolute relative to the project root
        const projectRoot = path.resolve(process.cwd());
        this.config.sharepoint.certPath = path.resolve(projectRoot, certPath);
        console.log(`🔧 Resolved certificate path to: ${this.config.sharepoint.certPath}`);
      } else {
        throw new Error('❌ SHP_CERT_PFX_PATH must be an absolute path');
      }
    }

    // Check if certificate file exists
    try {
      const resolvedPath = this.config.sharepoint.certPath;
      console.log(`🔍 Checking certificate file: ${resolvedPath}`);
      
      const stats = await fs.promises.stat(resolvedPath);
      if (!stats.isFile()) {
        throw new Error('❌ Certificate path does not point to a file');
      }

      // Check file permissions (should be readable only by owner)
      const mode = stats.mode & parseInt('777', 8);
      if (mode > parseInt('600', 8) && !isCloudDeployment) {
        console.warn('⚠️  Certificate file has overly permissive permissions. Consider: chmod 600 ' + resolvedPath);
      }

      console.log(`✅ Certificate file found and accessible (${stats.size} bytes)`);
    } catch (error) {
      if (error.code === 'ENOENT') {
        const suggestion = isCloudDeployment 
          ? 'Make sure the certificate file is properly uploaded to your deployment environment.'
          : 'Check the file path and ensure the certificate file exists.';
        throw new Error(`❌ Certificate file not found: ${this.config.sharepoint.certPath}\n   ${suggestion}`);
      }
      if (error.code === 'EACCES') {
        throw new Error(`❌ Cannot access certificate file: ${this.config.sharepoint.certPath}`);
      }
      throw error;
    }

    // Validate certificate password is not empty
    if (!certPassword || certPassword.trim().length === 0) {
      throw new Error('❌ SHP_CERT_PFX_PASSWORD cannot be empty');
    }
  }

  /**
   * Validate server configuration
   */
  validateServerConfig() {
    const { server } = this.config;

    // Validate server mode
    if (!['mcp', 'api', 'dual'].includes(server.mode)) {
      throw new Error('❌ SERVER_MODE must be one of: mcp, api, dual');
    }

    // Validate port range
    if (server.port < 1 || server.port > 65535) {
      throw new Error('❌ PORT must be between 1 and 65535');
    }

    // Validate log level
    const validLogLevels = ['error', 'warn', 'info', 'debug'];
    if (!validLogLevels.includes(server.logLevel)) {
      throw new Error('❌ LOG_LEVEL must be one of: ' + validLogLevels.join(', '));
    }
  }

  /**
   * Validate API configuration
   */
  validateApiConfig() {
    const { api, server } = this.config;

    // If running API server, validate auth tokens
    if (['api', 'dual'].includes(server.mode)) {
      if (api.authTokens.length === 0) {
        throw new Error('❌ API_AUTH_TOKENS is required when running in API mode');
      }

      // Validate token strength
      for (const token of api.authTokens) {
        if (token.length < 32) {
          throw new Error('❌ API auth tokens must be at least 32 characters long');
        }
      }
    }

    // Validate rate limiting values
    if (api.rateLimitWindow < 60000) { // 1 minute minimum
      throw new Error('❌ RATE_LIMIT_WINDOW must be at least 60000ms (1 minute)');
    }

    if (api.rateLimitMax < 1) {
      throw new Error('❌ RATE_LIMIT_MAX must be at least 1');
    }
  }

  /**
   * Get current configuration
   * @returns {Object|null} Current configuration or null if not loaded
   */
  getConfig() {
    return this.config;
  }

  /**
   * Check if configuration is loaded
   * @returns {boolean} True if configuration is loaded
   */
  isConfigLoaded() {
    return this.isLoaded;
  }

  /**
   * Reload configuration (useful for credential rotation)
   */
  async reload() {
    console.log('🔄 Reloading configuration...');
    this.config = null;
    this.isLoaded = false;
    return await this.load();
  }

  /**
   * Get sanitized configuration for logging (removes sensitive data)
   * @returns {Object} Sanitized configuration
   */
  getSanitizedConfig() {
    if (!this.config) {
      return null;
    }

    return {
      sharepoint: {
        appId: this.config.sharepoint.appId,
        tenantId: this.config.sharepoint.tenantId,
        siteUrl: this.config.sharepoint.siteUrl,
        certPath: this.config.sharepoint.certPath,
        certPassword: '***REDACTED***'
      },
      server: { ...this.config.server },
      api: {
        ...this.config.api,
        authTokens: this.config.api.authTokens.map(() => '***REDACTED***')
      }
    };
  }
}

// Create singleton instance
const configLoader = new ConfigLoader();

/**
 * Get the configuration loader instance
 * @returns {ConfigLoader} Configuration loader instance
 */
export function getConfigLoader() {
  return configLoader;
}

/**
 * Load configuration (convenience function)
 * @returns {Promise<Object>} Configuration object
 */
export async function loadConfig() {
  return await configLoader.load();
}

/**
 * Get current configuration (convenience function)
 * @returns {Object|null} Current configuration or null if not loaded
 */
export function getConfig() {
  return configLoader.getConfig();
}

export default configLoader;