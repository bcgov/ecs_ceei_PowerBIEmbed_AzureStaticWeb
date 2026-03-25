# Power BI Embed with Azure Static Web Apps

This repository contains a full-stack solution for embedding Power BI Reports securely. 

It is designed to be deployed directly to **Azure Static Web Apps (SWA)**, which provides free global hosting for the frontend and an integrated Azure Serverless Function for the backend.

## Why this Architecture?

Embedding Power BI requires an **Embed Token**. Generating this securely means authenticating with Azure AD using a Service Principal (Client ID and Secret). If you place these credentials directly in the browser's JavaScript, anyone could steal them. 

Azure Static Web Apps solves this by:
1. Hosting your `index.html` and `app.js` publicly.
2. Hosting the files in the `/api` folder privately as an Azure Serverless Function. The browser asks the API for a token, the API securely authenticates with Azure AD, and passes the token back. The browser never sees the secrets.

---

## Local Development Setup

To run this locally, you need the Azure Static Web Apps CLI and Azure Functions Core Tools installed.

### 1. Configure the Backend (API)
Open `api/local.settings.json` and provide your Azure AD **Service Principal** credentials:
- `TENANT_ID`: Your Azure AD Tenant ID
- `CLIENT_ID`: Your registered App ID
- `CLIENT_SECRET`: Your App Secret

### 2. Configure the Frontend
Open `app.js` and replace the configuration variables:
- `EMBED_GROUP_ID`: Your Power BI Workspace ID.
- `EMBED_REPORT_ID`: Your Power BI Report ID.
*(Leave `MS_FLOW_URL` as `/api/generateToken`)*

### 3. Run the App
Install dependencies and use the SWA CLI to start both the frontend and backend servers together:

```bash
cd api
npm install
cd ..
swa start . --api-location ./api
```

Open your browser to `http://localhost:4280`.

---

## Deployment to Azure Static Web Apps

Deploying to production is incredibly simple:

1. Push this code to a GitHub repository.
2. Go to the Azure Portal and create a new **Static Web App**.
3. Connect your GitHub repository. Azure will auto-generate a GitHub Actions workflow file.
4. **Important**: Go to the "Configuration" menu on your new Static Web App in the portal, and add your three Application Settings (`TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`).


