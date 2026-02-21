# Render Deployment Guide

This guide covers deploying your new stateless, job-based Excel-to-CSV filtering server on Render.

## Prerequisites
Your `server_v2` folder is now self-contained. It includes:
1. `main.py` - The FastAPI backend.
2. `requirements.txt` - Minimal dependencies for fast builds.

## Deployment Steps

1. **Commit to GitHub:**
   Commit the `server_v2` folder to your GitHub repository. (If you prefer, you can make `server_v2` the root of a new repository to keep it entirely separate from legacy code).

2. **Create New Web Service on Render:**
   - Go to [Render Dashboard](https://dashboard.render.com/).
   - Click **New** -> **Web Service**.
   - Connect your GitHub repository.

3. **Configure the Service:**
   - **Name:** e.g., `xlsx-filter-service`
   - **Environment:** `Python 3`
   - **Branch:** `main` (or whatever branch you committed to)
   - **Root Directory:** `server_v2` (Very important so Render finds the correct files).
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `uvicorn main:app --host 0.0.0.0 --port $PORT`

4. **Instance Type:**
   - Select the Free tier if testing, but ideally a **Starter ($7/mo) 512MB RAM** tier. The new code is highly optimized strictly for this explicit limit.

5. **Deploy:**
   - Click **Create Web Service**. 
   - Wait 2-3 minutes for the build to finish.

## Testing Production
Once deployed, Render provides a URL like `https://xlsx-filter-service.onrender.com`.

You can immediately test the UI tool directly from the cloud:
1. Visit `https://xlsx-filter-service.onrender.com/test` in your browser.
2. Select your test `.xlsx` file and provide a job ID.
3. Click "Process File". 

The endpoint will respond efficiently with the isolated CSV based on DC vs Hub headers.
