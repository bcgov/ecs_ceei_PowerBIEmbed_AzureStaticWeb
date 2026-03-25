const express = require('express');
const fetch = require('node-fetch');
const path = require('path');
const app = express();
const port = process.env.PORT || 8080;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Helper to pause execution for polling
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// --- Generate Token Route ---
app.post('/api/generateToken', async (req, res) => {
    console.log('API /api/generateToken hit.');

    const tenantId = process.env["TENANT_ID"];
    const clientId = process.env["CLIENT_ID"];
    const clientSecret = process.env["CLIENT_SECRET"];

    const groupId = req.body && req.body.groupid;
    const reportId = req.body && req.body.reportid;

    if (!groupId || !reportId) {
        return res.status(400).send("Please pass a groupid and reportid in the request body");
    }

    try {
        const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
        const tokenBody = new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            grant_type: 'client_credentials',
            scope: 'https://analysis.windows.net/powerbi/api/.default'
        });

        const tokenResponse = await fetch(tokenUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: tokenBody
        });

        if (!tokenResponse.ok) {
            const errDetails = await tokenResponse.text();
            throw new Error(`Failed to get AAD Token: ${tokenResponse.status} ${errDetails}`);
        }

        const tokenData = await tokenResponse.json();
        const aadToken = tokenData.access_token;

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

        res.status(200).json({ 
            token: pbiData.token,
            expiration: pbiData.expiration
        });

    } catch (error) {
        console.error(error);
        res.status(500).send(`Error generating token: ${error.message}`);
    }
});

// --- Export Report Route ---
app.post('/api/exportReport', async (req, res) => {
    console.log('API /api/exportReport hit.');

    const tenantId = process.env["TENANT_ID"];
    const clientId = process.env["CLIENT_ID"];
    const clientSecret = process.env["CLIENT_SECRET"];
    
    const groupId = req.body && req.body.groupid;
    const reportId = req.body && req.body.reportid;
    const bookmarkState = req.body && req.body.state;

    if (!groupId || !reportId) {
        return res.status(400).send("Please pass a groupid and reportid in the request body.");
    }

    try {
        const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
        const tokenBody = new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            grant_type: 'client_credentials',
            scope: 'https://analysis.windows.net/powerbi/api/.default'
        });

        const tokenResponse = await fetch(tokenUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: tokenBody
        });

        if (!tokenResponse.ok) {
            const errDetails = await tokenResponse.text();
            throw new Error(`Failed to get AAD Token: ${tokenResponse.status} ${errDetails}`);
        }

        const tokenData = await tokenResponse.json();
        const aadToken = tokenData.access_token;

        const exportUrl = `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}/ExportTo`;
        const exportPayload = {
            format: "PDF",
            powerBIReportConfiguration: {}
        };

        if (bookmarkState) {
            exportPayload.powerBIReportConfiguration = {
                defaultBookmark: { state: bookmarkState }
            };
        }

        const exportResponse = await fetch(exportUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${aadToken}`
            },
            body: JSON.stringify(exportPayload)
        });

        if (!exportResponse.ok) {
            const errDetails = await exportResponse.text();
            throw new Error(`Failed to start Export job: ${exportResponse.status} ${errDetails}`);
        }

        const exportData = await exportResponse.json();
        const exportId = exportData.id;
        console.log(`Export job started successfully. Export ID: ${exportId}`);

        const statusUrl = `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}/exports/${exportId}`;
        let exportStatus = "NotStarted";
        let downloadUrl = null;

        for (let i = 0; i < 30; i++) {
            await sleep(5000);

            const statusResponse = await fetch(statusUrl, {
                method: 'GET',
                headers: { 'Authorization': `Bearer ${aadToken}` }
            });

            if (!statusResponse.ok) {
                const errDetails = await statusResponse.text();
                throw new Error(`Failed to get Export Status: ${statusResponse.status} ${errDetails}`);
            }

            const statusData = await statusResponse.json();
            exportStatus = statusData.status;
            console.log(`Iteration ${i+1}: Export status is ${exportStatus}`);

            if (exportStatus === "Succeeded") {
                downloadUrl = statusData.resourceLocation;
                break;
            } else if (exportStatus === "Failed") {
                throw new Error("Export job failed on Power BI servers.");
            }
        }

        if (exportStatus !== "Succeeded" || !downloadUrl) {
            throw new Error(`Export job timed out after 2.5 minutes. Final status: ${exportStatus}`);
        }

        console.log("Export succeeded. Downloading file from Power BI...");
        const fileResponse = await fetch(downloadUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${aadToken}` }
        });

        if (!fileResponse.ok) {
            const errDetails = await fileResponse.text();
            throw new Error(`Failed to download exported file: ${fileResponse.status} ${errDetails}`);
        }

        const arrayBuffer = await fileResponse.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);

        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", 'attachment; filename="ReportExport.pdf"');
        res.status(200).send(buffer);
        console.log("File downloaded and sent to client successfully.");

    } catch (error) {
        console.error("Export Error: ", error);
        res.status(500).send(`Error processing export: ${error.message}`);
    }
});

// For any other basic routes, serve index.html (SPA-like fallback, though not strictly an SPA)
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// --- Start the server ---
app.listen(port, () => {
    console.log(`Power BI Embed App Service listening at http://localhost:${port}`);
});
