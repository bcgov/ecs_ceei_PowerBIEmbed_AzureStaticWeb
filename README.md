# Power BI Embed with Azure Web Apps

This repository contains a full-stack solution for embedding Power BI Reports securely. 

It is designed to be deployed directly to an **Azure App Service (Web App)** running Node.js, which hosts both the frontend UI and the backend API required for secure token generation.

## Why this Architecture?

Embedding Power BI requires an **Embed Token**. Generating this securely means authenticating with Azure AD using a Service Principal (Client ID and Secret). If you place these credentials directly in the browser's JavaScript, anyone could steal them. 

Azure App Service solves this by:
1. Hosting your frontend HTML, CSS, and JS files securely.
2. Running a Node.js Express server backend. The browser asks the backend API (`/api/generateToken`) for a token, the API securely authenticates with Azure AD, and passes the token back to the frontend. The browser never sees the secrets.

---

## Local Development Setup

To run this locally, you need Node.js installed on your machine.

### 1. Configure the Backend (API)
Create a `.env` file in the `AppServiceSolution` folder and provide your Azure AD **Service Principal** credentials:
```env
TENANT_ID=Your Azure AD Tenant ID
CLIENT_ID=Your registered App ID
CLIENT_SECRET=Your App Secret
```

### 2. Configure the Frontend
Open `app.js` and replace the configuration variables:
- `EMBED_GROUP_ID`: Your Power BI Workspace ID.
- `EMBED_REPORT_ID`: Your Power BI Report ID.
Open `public/app.js` and ensure it points to the correct endpoint:
- `MS_FLOW_URL`: Set to `/api/generateToken`

### 3. Run the App
Navigate into the solution directory, install dependencies, and start the local development server:

```bash
cd AppServiceSolution
npm install
node index.js
```

Open your browser to `http://localhost:8080`.

---

## Deployment to Azure App Service

Deploying to production is straightforward using Kudu ZipDeploy:

### 1. Create a Zip Archive
Since Azure App Service instances running Linux require forward slashes in file paths, generate a compatible zip using `bestzip`:

```bash
cd AppServiceSolution
npx -y bestzip ../AppServiceSolution.zip *
```

### 2. Deploy via Kudu ZipDeploy
1. Go to the Azure Portal and ensure you have created an **Azure Web App** running Node.js (e.g., Node 20 LTS or Node 22 LTS).
2. Add your three Environment Variables (`TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`) in the **Settings > Environment variables** menu of your Web App.
3. Keep the Startup Command empty so Azure automatically detects and runs `index.js`.
4. Open the Kudu ZipDeploy UI by navigating to:
   `https://<YOUR_APP_SERVICE_NAME>.scm.azurewebsites.net/ZipDeployUI`
5. Drag and drop the `AppServiceSolution.zip` file into the browser window.
6. Azure will automatically upload, extract, and install the dependencies (`npm install`) on the server. Your app is now live!
