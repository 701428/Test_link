# SharePoint Test Report Linker

Automatically inserts SharePoint folder hyperlinks into the **Test Report path** cells of your Excel test plan — directly on SharePoint, no file download needed.

---

## How it works

1. User signs in with their Microsoft account (OAuth2 popup)
2. The app reads all worksheet names from your Excel file on SharePoint via Graph API
3. For each sheet, it finds the cell labelled **"Test Report path"** and inserts a hyperlink in the adjacent cell pointing to `<Base Folder URL>/<Sheet Name>`
4. The updated file is saved back to SharePoint in-place

---

## Prerequisites

- Python 3.10+
- Node.js 18+
- An **Azure App Registration** (see setup below)

---

## Azure App Registration Setup (one-time, ~5 min)

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name it anything (e.g. `TestReportLinker`)
3. **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts*
4. **Redirect URI:** Platform = `Single-page application (SPA)`, URI = `http://localhost:5173`
   - For production, also add your server URL
5. Click **Register**
6. Copy the **Application (client) ID** — you'll need it below
7. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
   - Add: `Files.ReadWrite`
8. Click **Grant admin consent** (or ask your admin)

---

## Local Development

### 1. Backend

```bash
cd backend
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

Backend runs at `http://localhost:8000`

### 2. Frontend

```bash
cd frontend
cp .env.example .env
# Edit .env — set VITE_CLIENT_ID to your Azure App client ID
npm install
npm run dev
```

Frontend runs at `http://localhost:5173`

---

## Usage

1. Open `http://localhost:5173`
2. Click **Sign in with Microsoft** and log in with the account that has access to the SharePoint file
3. Verify the **Excel file URL** (pre-filled with your file)
4. Enter the **Base SharePoint Folder URL** (the folder containing `IMDS_001`, `IMDS_002`, etc.)
5. Click **Insert Hyperlinks & Save to SharePoint**
6. Done — open the Excel file on SharePoint to see the hyperlinks

---

## Deploying to a Server

### Backend (e.g. Railway, Render, Azure App Service)
- Deploy the `backend/` folder as a Python app
- Set start command: `uvicorn main:app --host 0.0.0.0 --port 8000`

### Frontend (e.g. Vercel, Netlify, Azure Static Web Apps)
- Deploy the `frontend/` folder
- Set environment variables:
  - `VITE_CLIENT_ID` = your Azure client ID
  - `VITE_BACKEND_URL` = your deployed backend URL
- Add your production domain as a **Redirect URI** in the Azure App Registration (SPA platform)

---

## Project Structure

```
Test_link/
├── backend/
│   ├── main.py          # FastAPI — Graph API calls + openpyxl processing
│   └── requirements.txt
├── frontend/
│   ├── src/
│   │   ├── App.jsx      # Main React UI
│   │   ├── App.css
│   │   └── main.jsx
│   ├── index.html
│   ├── package.json
│   ├── vite.config.js
│   └── .env.example
└── README.md
```
