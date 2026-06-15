# Textile Prospector — Smartex

AI-powered B2B prospecting agent for textile manufacturers. Finds weaving mills, dyehouses, and finishing houses with active buying signals and generates personalised LinkedIn + email outreach.

---

## 🚀 Deploy to Vercel (5 minutes)

### 1. Create a GitHub repo

1. Go to [github.com/new](https://github.com/new)
2. Name it `smartex-prospector`, set to **Private**
3. Upload all files from this ZIP (drag & drop into the repo)

### 2. Deploy on Vercel

1. Go to [vercel.com](https://vercel.com) → **Add New Project**
2. Import your GitHub repo
3. Framework: **Vite** (auto-detected)
4. Click **Deploy** — wait ~1 min

### 3. Add your Anthropic API key

1. In Vercel → your project → **Settings → Environment Variables**
2. Add:
   - **Name:** `ANTHROPIC_API_KEY`
   - **Value:** `sk-ant-...` (your key from [console.anthropic.com](https://console.anthropic.com))
3. Click **Save**
4. Go to **Deployments** → click the three dots on the latest → **Redeploy**

Your app is now live at `https://smartex-prospector.vercel.app` (or similar).

---

## 💻 Run locally

```bash
# 1. Install dependencies
npm install

# 2. Add your API key
cp .env.example .env.local
# Edit .env.local and paste your key

# 3. Start dev server
npm run dev
# Open http://localhost:5173
```

---

## 🔑 Getting an Anthropic API key

1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Sign up / log in
3. **API Keys** → **Create Key**
4. Copy the `sk-ant-...` key

Typical cost: ~$0.01–0.05 per prospecting run (5–30 leads).

---

## 📁 Project structure

```
smartex-prospector/
├── api/
│   └── claude.js        ← Vercel serverless proxy (keeps API key secure)
├── src/
│   ├── main.jsx         ← React entry point
│   └── App.jsx          ← Main app component
├── index.html
├── vite.config.js
├── vercel.json
└── package.json
```

---

## ✉️ Sharing with your team

Once deployed, share the Vercel URL. Anyone with the link can use the app — no login required, no installation.

To restrict access, enable **Vercel Authentication** in Project Settings → Deployment Protection.
