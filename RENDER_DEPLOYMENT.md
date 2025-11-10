# Render Deployment Guide

## Quick Fix for "Cannot find module" Error

If you're seeing this error:
```
Error: Cannot find module '/opt/render/project/src/server.js'
```

### Solution Options:

#### Option 1: Update Render Configuration
1. In your Render dashboard, go to your service settings
2. Change the **Start Command** from `npm start` to:
   ```
   node src/start.js
   ```

#### Option 2: Manual Deploy with render.yaml
1. Push the updated code with `render.yaml` included
2. Render should automatically detect and use the configuration

#### Option 3: Environment Variables Check
Ensure these environment variables are set in Render:

**Required:**
- `NODE_ENV=production`
- `SERVER_MODE=api`
- `PORT` (set by Render automatically)

**SharePoint Configuration:**
- `SHAREPOINT_TENANT_ID=your-tenant-id`
- `SHAREPOINT_CLIENT_ID=your-client-id`
- `SHAREPOINT_SITE_URL=https://your-company.sharepoint.com/sites/your-site`
- `SHAREPOINT_CERT_PATH=./certificate.pfx`
- `SHAREPOINT_CERT_PASSWORD=your-cert-password`
- `API_AUTH_TOKENS=your-api-token-here`

## Deployment Steps:

1. **Connect Repository:** Connect your GitHub repository to Render
2. **Auto-Deploy:** Render will automatically deploy when you push to main branch
3. **Environment Variables:** Set all required environment variables in Render dashboard
4. **Certificate Upload:** Upload your `certificate.pfx` file to Render (if using cert auth)
5. **Health Check:** Render will use `/health` endpoint to verify service is running

## Debugging:

If deployment fails:

1. Check Render logs for specific errors
2. Verify all environment variables are set
3. Ensure `certificate.pfx` file is accessible (if using certificate auth)
4. Test locally with: `npm run start:api`

## File Structure Required:

```
src/
├── server.js       (main server entry)
├── start.js        (robust entry point)
├── config.js       (configuration loader)
├── api-server.js   (API server)
└── other files...
package.json
render.yaml         (Render configuration)
```

## Health Check:

Once deployed, test your service:
```bash
curl https://your-app.onrender.com/health
```

Should return:
```json
{
  "status": "healthy",
  "sharepoint": "connected",
  "authentication": "required"
}
```