# 🚀 D2C Scaler Performance Command Center
### Shareable USB & Offline Dashboard Bundle

Welcome to the **D2C Scaler Performance Command Center**! This bundle is completely self-contained and pre-packaged to run directly from a USB drive, local hard disk, or network share without requiring an active internet connection or any complex server installations.

---

## 📂 Bundle File Structure

When you share this folder (e.g. copy it to a USB drive), ensure the following files remain together in the same directory:

```text
📁 D2C_Scaler_Shareable/
├── 🌐 index.html                          <-- Double-click to launch the Performance Dashboard
├── 📊 D2C_Marketing_Dashboard_Template.xlsx <-- Main Excel Template (Meta + Google + Sales)
├── 🏷️ D2C_Scaler_Ads_Nomenclature.xlsx     <-- Google/Meta naming convention tracker
├── 📈 Ads_Tracker_Template.xlsx            <-- Daily ads performance ledger template
├── 📦 d2c_app_icon.png                     <-- PWA Icon
├── 📄 manifest.json                        <-- PWA Configuration
├── ⚙️ service-worker.js                    <-- PWA Caching Service (enabled on HTTP/HTTPS)
├── 📚 chart.umd.min.js                     <-- Offline Chart.js library
├── 📚 papaparse.min.js                     <-- Offline CSV Parser library
└── 📚 xlsx.full.min.js                     <-- Offline Excel Parser library
```

---

## ⚡ Quick-Start Guide (Offline / USB Mode)

### 1. Launch the Dashboard
Double-click the **`index.html`** file on your USB drive. It will open instantly in your default web browser (Chrome, Safari, Edge, or Firefox).

### 2. Instant Simulated Experience (Plug-and-Play)
If you are opening the dashboard for the first time on a new computer:
- The dashboard automatically detects a blank state.
- It will instantly load **30 days of simulated D2C mock data** (funnels, waterfall diagnostic charts, cohort heatmaps, daily ledgers) so the user gets a working demo without any configuration.
- **No data errors, blank sheets, or empty screen states will be displayed!**

---

## 📥 Loading Your Own D2C Data

You can load your own live store performance metrics in two ways:

### Method A: Local Excel Drag-and-Drop (100% Offline)
If you don't have internet access or want to keep your data private:
1. Open the included **`D2C_Marketing_Dashboard_Template.xlsx`** and paste your Meta, Google, Shopify/Sales, and Product Margin data into their respective tabs.
2. Open the dashboard (`index.html`) in your browser.
3. Go to the **Data Ingestion & Sync** tab (the last navigation link).
4. Drag and drop your saved `.xlsx` template directly onto the drop area, or click to select the file.
5. The dashboard will instantly extract, calculate, and render your charts and cohort matrices locally in less than 500ms!

### Method B: Google Sheets Live Sync (Online)
If you want to sync live with a shared Google Sheet:
1. Upload the template **`D2C_Marketing_Dashboard_Template.xlsx`** to Google Drive and open it as a Google Sheet.
2. Click **File** -> **Share** -> **Share with others** and set the permissions to **"Anyone with the link can view"** (Viewer permission is required for the GViz feed).
3. Copy the Google Sheet URL.
4. Paste the URL into the **Google Sheet URL or ID** field in the **Data Ingestion & Sync** tab and click **Sync Google Sheet**.
5. The dashboard will connect via secure JSONP (bypassing browser CORS blocks) and pull your live sheet updates instantly!

---

## 💡 Troubleshooting & Security Details

- **Why are there warnings in the Console?**
  When running locally from a USB drive via the `file://` protocol, modern browsers restrict certain advanced features (like installing PWAs via Service Workers or manifest validation) for local security reasons. We have bypassed these restrictions gracefully—the app automatically turns off service workers on local drives to avoid errors, while retaining 100% of the dashboard logic, charts, and Excel parsers.
- **Can I run this offline?**
  Yes! Because the core libraries (`Chart.js`, `PapaParse`, `XLSX`) are packaged locally, all calculation engines, cohort calculations, and interactive visual charts remain fully operational even if you have zero internet connectivity.

---

*Built for D2C Founders, Marketers, and growth agencies.*
