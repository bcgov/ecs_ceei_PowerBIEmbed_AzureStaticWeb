/**
 * Azure Function to handle Export-To-File for Power BI Embedded
 * 
 * Flow:
 * 1. Receive groupId, reportId, and optionally bookMark from frontend.
 * 2. Authenticate with AAD to get an Access Token.
 * 3. Start the Export-To-File async job on Power BI API.
 * 4. Poll the status URL until the job succeeds.
 * 5. Download the final PDF and stream it back to the client.
 */

const fetch = require('node-fetch');

// Helper to pause execution for polling
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

module.exports = async function (context, req) {
    context.log('HTTP trigger processing request for Power BI Export.');

    const tenantId = process.env["TENANT_ID"];
    const clientId = process.env["CLIENT_ID"];
    const clientSecret = process.env["CLIENT_SECRET"];
    
    // Parse input
    const groupId = req.body && req.body.groupid;
    const reportId = req.body && req.body.reportid;
    const bookmarkState = req.body && req.body.state; // The active filters

    if (!groupId || !reportId) {
        context.res = {
            status: 400,
            body: "Please pass a groupid and reportid in the request body."
        };
        return;
    }

    try {
        // --- 1. Authenticate with Azure AD ---
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


        // --- 2. Start the Export-To-File Job ---
        const exportUrl = `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}/ExportTo`;
        
        const exportPayload = {
            format: "PDF",
            powerBIReportConfiguration: {}
        };

        // If a bookmark was passed, inject it into the payload
        if (bookmarkState) {
            exportPayload.powerBIReportConfiguration = {
                defaultBookmark: {
                    state: bookmarkState
                }
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
        context.log(`Export job started successfully. Export ID: ${exportId}`);


        // --- 3. Poll for Export Status ---
        const statusUrl = `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}/exports/${exportId}`;
        let exportStatus = "NotStarted";
        let downloadUrl = null;

        // Poll up to 30 times (e.g., waiting 5 seconds each time = max ~2.5 mins)
        for (let i = 0; i < 30; i++) {
            await sleep(5000); // Wait 5 seconds between polls

            const statusResponse = await fetch(statusUrl, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${aadToken}`
                }
            });

            if (!statusResponse.ok) {
                const errDetails = await statusResponse.text();
                throw new Error(`Failed to get Export Status: ${statusResponse.status} ${errDetails}`);
            }

            const statusData = await statusResponse.json();
            exportStatus = statusData.status;

            context.log(`Iteration ${i+1}: Export status is ${exportStatus}`);

            if (exportStatus === "Succeeded") {
                // The export is complete, get the download URL
                downloadUrl = statusData.resourceLocation;
                break;
            } else if (exportStatus === "Failed") {
                throw new Error("Export job failed on Power BI servers.");
            }
            
            // If Running or NotStarted, loop continues
        }

        if (exportStatus !== "Succeeded" || !downloadUrl) {
            throw new Error(`Export job timed out after 2.5 minutes. Final status: ${exportStatus}`);
        }

        // --- 4. Download and Return the PDF File ---
        context.log("Export succeeded. Downloading file from Power BI...");
        const fileResponse = await fetch(downloadUrl, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${aadToken}`
            }
        });

        if (!fileResponse.ok) {
            const errDetails = await fileResponse.text();
            throw new Error(`Failed to download exported file: ${fileResponse.status} ${errDetails}`);
        }

        // Get the binary buffer from the response
        const arrayBuffer = await fileResponse.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);

        context.res = {
            status: 200,
            headers: {
                "Content-Type": "application/pdf",
                "Content-Disposition": `attachment; filename="ReportExport.pdf"`
            },
            isRaw: true,
            body: buffer
        };
        context.log("File downloaded and sent to client successfully.");

    } catch (error) {
        context.log.error("Export Error: ", error);
        context.res = {
            status: 500,
            body: `Error processing export: ${error.message}`
        };
    }
};
