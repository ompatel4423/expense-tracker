# 💳 Expense Tracker (PWA)

A high-performance, dark-mode Progressive Web Application (PWA) for automated personal finance tracking, integrated with Axis Bank transaction alerts via Google Sheets and Google Apps Script.

---

## 🚀 Live Links & Resources
- **Google Sheet**: [Axis Bank Expense Tracker Sheet](https://docs.google.com/spreadsheets/d/11d93LcZgoc1fOm1TLDsjT19PK4uw7zSf4juzFaLf_ck/edit?usp=drive_link)
- **Google Apps Script API**: `https://script.google.com/macros/s/AKfycbz4LphVWeHvjVxweBjqJ7zP4ISgzxeGiuT4zNH_vkeAYizlQG8MpRCChwfrpLEaH4yRoQ/exec`
- **Netlify Deployment**: [zesty-nougat-65e4db.netlify.app](https://zesty-nougat-65e4db.netlify.app)

---

## ✨ Features
- **Real-Time Automated Tracking**: Pulls debited transaction alerts directly from Axis Bank via Gmail into Google Sheets.
- **REST API Backend**: Google Apps Script acts as a fast JSON API for the frontend.
- **Progressive Web App (PWA)**: Installable on iOS (via Home Screen or `.mobileconfig` profile) and Android with offline caching via Service Worker.
- **Fluid WebGL Background**: Ultra-modern reactive liquid metal aesthetic.
- **Analytics & Visualizations**: Daily & monthly spending meters, merchant breakdowns, budget alerts, and transaction history.
- **Edge Deployment Ready**: Optimized for zero-config Vercel hosting with tailored cache-busting headers for instant updates.

---

## 🛠️ Deploy to Vercel

### Option 1: Import via GitHub (Recommended)
1. Push this repository to GitHub:
   ```bash
   git push -u origin main
   ```
2. Go to [vercel.com/new](https://vercel.com/new).
3. Select your `expense-tracker` repository from the list.
4. Leave all settings at default (**Framework Preset: Other**, **Root Directory: `./`**).
5. Click **Deploy**.

> Any future `git push` will automatically trigger a new deployment on Vercel with zero downtime.

### Option 2: Deploy using Vercel CLI
```bash
npx vercel
```
Follow the interactive prompts to link and deploy your site in seconds.

---

## 📱 Mobile Installation

### iPhone / iOS
- **Method A (Safari PWA)**: Open your Vercel URL in Safari -> tap **Share** -> tap **Add to Home Screen**.
- **Method B (Native Standalone WebClip Profile)**: Tap your profile avatar in the app header -> click **"Download iOS Profile"** -> Open iOS **Settings** -> **Profile Downloaded** -> tap **Install**. The app will launch in dedicated fullscreen mode.

### Android
- Open your Vercel URL in Chrome -> tap the install banner or 3-dots menu -> tap **Install App**.

---

## 📁 Project Structure
```text
expense-tracker/
├── index.html                 # Main PWA application (UI, WebGL, animations, API client)
├── manifest.json              # PWA manifest
├── sw.js                      # Service Worker (offline cache & dynamic sync)
├── app_logo.jpg               # High-res app icon & maskable assets
├── vercel.json                # Vercel caching & routing headers
├── Treasury.mobileconfig      # Standalone iOS WebClip profile
└── google-apps-script/
    ├── Code.gs                # Google Apps Script backend source code
    └── README.md              # Script API documentation
```
