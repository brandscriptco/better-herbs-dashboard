# Better Herbs — Performance Dashboard

Interactive ad performance dashboard with product-wise filtering, ad-level breakdown, and creator search.

## Quick Deploy to Vercel

### Step 1: Install Node.js
Download and install from https://nodejs.org (LTS version)

### Step 2: Push to GitHub
1. Go to https://github.com/new and create a new repository called `better-herbs-dashboard`
2. Open your terminal and run:

```bash
cd better-herbs
git init
git add .
git commit -m "Initial dashboard"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/better-herbs-dashboard.git
git push -u origin main
```

### Step 3: Deploy on Vercel
1. Go to https://vercel.com and sign up with your GitHub account
2. Click "Add New" → "Project"
3. Import your `better-herbs-dashboard` repo
4. Vercel auto-detects Vite — just click "Deploy"
5. Done! Your dashboard is live at `better-herbs-dashboard.vercel.app`

## Local Development

```bash
npm install
npm run dev
```

Opens at http://localhost:5173

## Updating Data

Edit the `DATA` object in `src/App.jsx` with your latest Meta & Google Ads numbers.
