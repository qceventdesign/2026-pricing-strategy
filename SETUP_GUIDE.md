# Quill Creative Pricing Engine — Setup Guide

Step-by-step instructions for running the app locally on your machine.

---

## Prerequisites

Make sure you have the following installed before starting:

| Tool    | Minimum Version | Check with          |
|---------|----------------|----------------------|
| Python  | 3.13+          | `python3 --version`  |
| Node.js | 18+            | `node --version`     |
| npm     | 9+             | `npm --version`      |
| Git     | any            | `git --version`      |

---

## Step 1 — Clone the Repository

```bash
git clone <your-repo-url>
cd 2026-pricing-strategy
```

If you already have the repo, just navigate to the project root:

```bash
cd /path/to/2026-pricing-strategy
```

---

## Step 2 — Set Up the Python Backend

### 2a. Create a virtual environment

```bash
python3 -m venv venv
```

### 2b. Activate the virtual environment

**macOS / Linux:**
```bash
source venv/bin/activate
```

**Windows (PowerShell):**
```powershell
.\venv\Scripts\Activate
```

You should see `(venv)` at the beginning of your terminal prompt.

### 2c. Install Python dependencies

```bash
pip install -r app/backend/requirements.txt
```

This installs:
- **FastAPI** — web framework
- **Uvicorn** — ASGI server
- **Pydantic** — request/response validation
- **openpyxl** — Excel export
- **python-multipart** — form data handling

---

## Step 3 — Start the Backend Server

From the **project root** directory (`2026-pricing-strategy/`), run:

```bash
uvicorn app.backend.main:app --host 0.0.0.0 --port 8000
```

You should see output like:

```
INFO:     Uvicorn running on http://0.0.0.0:8000
INFO:     Started server process
```

### Verify it's working

Open a new terminal tab and run:

```bash
curl http://localhost:8000/api/health
```

Expected response:

```json
{"status": "ok", "version": "1.0.0"}
```

> Keep this terminal running. The backend needs to stay active.

---

## Step 4 — Set Up the Frontend

Open a **new terminal tab/window** and navigate to the frontend directory:

```bash
cd /path/to/2026-pricing-strategy/app/frontend
```

### 4a. Install Node dependencies

```bash
npm install
```

This installs React, Vite, Tailwind CSS, and all other frontend packages.

### 4b. Start the frontend dev server

```bash
npm run dev
```

You should see output like:

```
  VITE v6.x.x  ready in XXX ms

  ➜  Local:   http://localhost:5173/
  ➜  Network: http://192.168.x.x:5173/
```

---

## Step 5 — Open the App

Open your browser and go to:

### **http://localhost:5173**

That's it — you should see the Quill Creative Pricing Engine dashboard.

---

## Quick Reference

| Service  | URL                          | Purpose                  |
|----------|------------------------------|--------------------------|
| Frontend | http://localhost:5173        | The app (open this one)  |
| Backend  | http://localhost:8000        | API server               |
| API Docs | http://localhost:8000/docs   | Interactive Swagger docs |
| Health   | http://localhost:8000/api/health | Backend status check |

---

## Project Structure

```
2026-pricing-strategy/
├── app/
│   ├── backend/                  # FastAPI server
│   │   ├── main.py               # API routes
│   │   ├── pricing_engine.py     # Pricing calculation logic
│   │   ├── config_loader.py      # Loads JSON config files
│   │   ├── schemas.py            # Request/response models
│   │   ├── storage.py            # JSON file storage for estimates
│   │   └── requirements.txt      # Python dependencies
│   ├── frontend/                 # React + Vite app
│   │   ├── src/
│   │   │   ├── pages/            # Dashboard, NewEstimate, EstimateDetail, ConfigView
│   │   │   ├── components/       # LineItemRow, PnLPanel, HealthBadge
│   │   │   ├── api.ts            # API client
│   │   │   └── types.ts          # TypeScript types
│   │   ├── package.json
│   │   └── vite.config.ts        # Proxies /api → backend
│   └── data/
│       └── estimates.json        # Saved estimates (auto-created)
├── config/                       # Pricing config (JSON files)
│   ├── markups.json              # Category markup percentages
│   ├── rates.json                # Hourly/standard rates
│   ├── fees.json                 # Fee defaults
│   ├── commissions.json          # Commission scenarios
│   └── thresholds.json           # Margin health targets
└── venv/                         # Python virtual environment
```

---

## Troubleshooting

### "Module not found" when starting the backend

Make sure you are running `uvicorn` from the **project root** directory (not from inside `app/backend/`). The command `app.backend.main:app` relies on the Python module path from the root.

### Frontend shows a blank page or API errors

- Confirm the backend is running on port 8000 (check the terminal where you started it)
- The frontend proxies `/api` requests to `http://localhost:8000` — both servers must be running simultaneously

### Port already in use

If port 8000 or 5173 is taken, stop whatever is using it:

```bash
# Find and kill process on port 8000
lsof -ti:8000 | xargs kill

# Find and kill process on port 5173
lsof -ti:5173 | xargs kill
```

Then restart the servers.

### Python virtual environment issues

If `pip install` fails, try recreating the virtual environment:

```bash
rm -rf venv
python3 -m venv venv
source venv/bin/activate
pip install -r app/backend/requirements.txt
```

---

## Stopping the App

Press `Ctrl + C` in each terminal window (backend and frontend) to stop the servers.
