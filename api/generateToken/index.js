/**
 * Azure Function to generate a Power BI Embed Token
 * 
 * To run this locally:
 * 1. Ensure you have the Azure Functions Core Tools installed.
 * 2. Run `npm install node-fetch@2` in your function directory.
 * 3. Add your configurations to local.settings.json
 */

const fetch = require('node-fetch');

module.exports = async function (context, req) {
    context.log('HTTP trigger function processing request for Embed Token.');

    // 1. Get configuration from environment variables (local.settings.json or Azure portal)
    const tenantId = process.env["TENANT_ID"];
    const clientId = process.env["CLIENT_ID"];
    const clientSecret = process.env["CLIENT_SECRET"];
    
    // 2. Extract Group ID and Report ID from the incoming request body
    const groupId = req.body && req.body.groupid;
    const reportId = req.body && req.body.reportid;

    if (!groupId || !reportId) {
        context.res = {
            status: 400,
            body: "Please pass a groupid and reportid in the request body"
        };
        return;
    }

    try {
        // 3. Authenticate with Azure AD to get an Access Token using Service Principal
        const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
        
        const tokenBody = new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            grant_type: 'client_credentials',
            scope: 'https://analysis.windows.net/powerbi/api/.default'
        });

        const tokenResponse = await fetch(tokenUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: tokenBody
        });

        if (!tokenResponse.ok) {
            const errDetails = await tokenResponse.text();
            throw new Error(`Failed to get AAD Token: ${tokenResponse.status} ${errDetails}`);
        }

        const tokenData = await tokenResponse.json();
        const aadToken = tokenData.access_token;

        // 4. Use the AAD Token to request a Power BI Embed Token
        const pbiUrl = `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}/GenerateToken`;
        
        const pbiBody = {
            accessLevel: "View",
            allowSaveAs: "false"
        };

        const pbiResponse = await fetch(pbiUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${aadToken}`
            },
            body: JSON.stringify(pbiBody)
        });

        if (!pbiResponse.ok) {
            const errDetails = await pbiResponse.text();
            throw new Error(`Failed to get Embed Token: ${pbiResponse.status} ${errDetails}`);
        }

        const pbiData = await pbiResponse.json();

        // 5. Return the Embed Token to the frontend
        context.res = {
            status: 200,
            body: { 
                token: pbiData.token,
                expiration: pbiData.expiration
            }
        };

    } catch (error) {
        context.log.error(error);
        context.res = {
            status: 500,
            body: `Error generating token: ${error.message}`
        };
    }
};
