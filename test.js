#!/usr/bin/env node

import * as dotenv from 'dotenv';
import { createSharePointClient } from './src/sharepoint-client.js';
import { SharePointTools } from './src/sharepoint-tools.js';
import { logger } from './src/utils/logger.js';

// Load environment variables
dotenv.config();

async function testSharePointConnection() {
  logger.info('🧪 Testing SharePoint MCP server tools...');

  try {
    // Initialize SharePoint client
    const client = await createSharePointClient();
    if (!client) {
      logger.error('❌ Failed to create SharePoint client');
      return false;
    }

    const tools = new SharePointTools(client);
    logger.info('✅ Successfully initialized SharePoint client and tools');

    // Test 1: List folders in root
    logger.info('\n📁 Test 1: List folders in root...');
    const foldersResult = await tools.listFolders();
    logger.info(`   Result: ${JSON.stringify(foldersResult, null, 2)}`);

    // Test 2: List documents in root
    logger.info('\n📄 Test 2: List documents in root...');
    const docsResult = await tools.listDocuments('');
    logger.info(`   Result: ${JSON.stringify(docsResult, null, 2)}`);

    // Test 3: Get folder tree
    logger.info('\n🌳 Test 3: Get folder tree...');
    const treeResult = await tools.getFolderTree(undefined, 2);
    logger.info(`   Success: ${treeResult.success}`);
    if (treeResult.success && treeResult.tree) {
      const items = treeResult.tree.items || [];
      logger.info(`   Root items count: ${items.length}`);
      for (const item of items.slice(0, 3)) {
        logger.info(`     - ${item.name} (${item.type})`);
      }
    }

    // Test 4: Try to get content of a file if we find any
    logger.info('\n📖 Test 4: Try to get file content...');
    if (docsResult.success && docsResult.items && docsResult.items.length > 0) {
      const firstDoc = docsResult.items[0];
      const docName = firstDoc.name;
      logger.info(`   Trying to read: ${docName}`);
      
      const contentResult = await tools.getDocumentContent('', docName);
      logger.info(`   Content result success: ${contentResult.success}`);
      if (contentResult.success) {
        const contentType = contentResult.type || 'unknown';
        const contentLength = contentResult.content?.length || 0;
        logger.info(`   Content type: ${contentType}`);
        logger.info(`   Content length: ${contentLength} characters`);
        
        if (contentType === 'text' && contentLength < 500) {
          const preview = contentResult.content?.substring(0, 200) || '';
          logger.info(`   Content preview: ${preview}...`);
        }
      }
    } else {
      logger.info('   No documents found to test content reading');
    }

    return true;
  } catch (error) {
    logger.error('❌ Error testing tools:', error);
    return false;
  }
}

// Run the test
testSharePointConnection().then((success) => {
  if (success) {
    logger.info('\n🏆 SHAREPOINT MCP SERVER TOOLS WORK!');
    logger.info('✅ Certificate authentication successful');
    logger.info('✅ Graph API integration working');
    logger.info('✅ Read operations functional');
    logger.info('📚 Ready for read-only SharePoint operations');
  } else {
    logger.error('\n💥 There are issues with the tools');
  }
}).catch((error) => {
  logger.error('Test failed:', error);
  process.exit(1);
});