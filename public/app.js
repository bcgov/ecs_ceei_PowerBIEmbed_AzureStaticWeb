/**
 * ⚠️ CONFIGURATION REQUIRED ⚠️
 * Replace these values before running the application.
 */
const MS_FLOW_URL = "/api/generateToken";
// Dictionary of available reports mapping a friendly name to their IDs
const REPORTS = {
    lgcap: {
        groupId: "09a7f57f-a3f2-4571-9a88-f6637406685a", // Replace with LGCAP Group ID
        reportId: "3e2def4d-bbbb-48b8-b4be-b7d1cea33c24"  // Replace with LGCAP Report ID
    },
    ceei: {
        groupId: "09a7f57f-a3f2-4571-9a88-f6637406685a", // Replace with CEEI Group ID
        reportId: "fda20a6d-87c0-48e3-9989-d8f17404499a"  // Replace with CEEI Report ID
    }
};

// Parse the current URL query parameters (e.g. ?report=lgcap)
const urlParams = new URLSearchParams(window.location.search);
const requestedReport = urlParams.get('report') ? urlParams.get('report').toLowerCase() : 'lgcap'; // Default to lgcap

// Get the specific IDs for the requested report
const currentReportConfig = REPORTS[requestedReport] || REPORTS['lgcap'];

const EMBED_GROUP_ID = currentReportConfig.groupId;
const EMBED_REPORT_ID = currentReportConfig.reportId;
const EMBED_TYPE = "report"; // "report", "dashboard", or "tile"
const TOKEN_TYPE = 1; // 1 = powerbi.models.TokenType.Embed, 0 = Aad

// UI Elements
const reportContainer = document.getElementById("reportContainer");
const loadingMessage = document.getElementById("loadingMessage");
const btnFullscreen = document.getElementById("btnFullscreen");
const btnPrint = document.getElementById("btnPrint");

let activeReport = null;

/**
 * Initializes the application by fetching the token and rendering the report.
 */
async function initializeApp() {
    try {
        console.log("Requesting token from MS Flow...");

        // 1. Fetch the token from your backend proxy (MS Flow)
        const response = await fetch(MS_FLOW_URL, {
            method: 'POST',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                groupid: EMBED_GROUP_ID,
                reportid: EMBED_REPORT_ID
            })
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        const accessToken = data.token; // Ensure this matches the JSON property output by your Flow

        console.log("Token retrieved. Embedding report...");

        // 2. Embed the report
        embedPowerBiReport(accessToken);

    } catch (error) {
        console.error("Error initializing app:", error);
        loadingMessage.innerText = "Failed to load report. Check console for details. Ensure configuration variables in app.js are populated.";
        loadingMessage.style.color = "red";
    }
}

/**
 * Embeds the Power BI report into the container 
 * using the provided access token.
 */
function embedPowerBiReport(accessToken) {
    // Hide loading message and show container
    loadingMessage.style.display = "none";
    reportContainer.style.display = "block";

    const embedURL = `https://app.powerbi.com/reportEmbed?groupId=${EMBED_GROUP_ID}`;

    // Configuration object for embedding
    const config = {
        type: EMBED_TYPE,
        tokenType: TOKEN_TYPE,
        accessToken: accessToken,
        embedUrl: embedURL,
        id: EMBED_REPORT_ID,
        permissions: window['powerbi-client'].models.Permissions.All, // or Permissions.Read
        settings: {
            panes: {
                filters: {
                    expanded: false,
                    visible: false
                },
                pageNavigation: {
                    visible: true
                }
            },
            bars: {
                actionBar: {
                    visible: true
                }
            },
            layoutType: window['powerbi-client'].models.LayoutType.Custom,
            customLayout: {
                displayOption: window['powerbi-client'].models.DisplayOption.FitToPage
            }
        }
    };

    // Embed the report using the powerbi service available on the global window object (loaded via CDN)
    activeReport = powerbi.embed(reportContainer, config);

    // Wire up events
    activeReport.on("loaded", function () {
        console.log("Report loaded successfully.");
        // Enable buttons once loaded
        btnFullscreen.disabled = false;
        btnPrint.disabled = false;
    });

    activeReport.on("rendered", function () {
        console.log("Report rendered.");
    });

    activeReport.on("error", function (event) {
        console.error("Report encountered an error:", event.detail);
    });

    activeReport.on("pageChanged", function (event) {
        console.log("Page changed:", event.detail);
    });
}

// Wire up UI interactions
btnFullscreen.addEventListener("click", () => {
    if (activeReport) activeReport.fullscreen();
});

btnPrint.addEventListener("click", () => {
    if (activeReport) activeReport.print();
});

// Start the app!
initializeApp();
