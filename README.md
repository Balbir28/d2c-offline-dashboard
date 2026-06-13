# 📊 D2C Marketing Command Center (User Manual)

Welcome to your D2C Marketing Dashboard! This repository builds two beautiful, high-tech marketing command centers that synchronize to give you deep media-buying intelligence:

1. **Spreadsheet-Integrated Native Dashboard** (direct tabs inside Google Sheets, including the daily stats ledger).
2. **Standalone HTML Dashboard File** (`index.html`) — a premium, standalone visual interface featuring glassmorphic cards, active charts, cohorted LTV analysis, and creator optimization panels that you can open offline on your computer.

---

## 🚀 How to Set Up the Google Sheet (One-Time)

### Step 1: Initialize Sheets Headers
1. Open your Google Sheet spreadsheet.
2. In the top menu, select **`🚀 D2C Intelligence`** > **`⚙️ Setup Sheets`**.
   * *This automatically creates the standard sheet tabs (`META_DATA`, `GOOGLE_DATA`, `DAILY_STATS`, `PRODUCT_MASTER`, `SALES_DATA`, `CONFIG`) and configures their column headers in the correct order.*

### Step 2: Load Test Data (Optional)
If you want to play with simulated 90-day data:
1. In the top menu, select **`🚀 D2C Intelligence`** > **`🌱 Insert Sample Data`**.
2. This loads mock spend, conversion, video creative hook/hold rates, and Shopify order data.

### Step 3: Refresh and View the Spreadsheet Dashboard
1. Select **`🚀 D2C Intelligence`** > **`🔄 Update Dashboard (Native)`**.
   * *This aggregates your raw inputs and updates the native spreadsheet scorecard block, Top/Bottom Ad Sets, and product daily trends in the `DASHBOARD` and `DAILY_STATS` tabs.*

---

## 💻 How to Use the Standalone HTML Dashboard (`index.html`)

The standalone dashboard is a single HTML file (`index.html`) that runs entirely inside your web browser. It does not require external servers and stores data securely on your own computer.

### Step 1: Open the file
Locate [index.html](file:///Users/balbaasaur/Documents/D2C%20Google%20Sheet%20Dashboard%20Builder/index.html) in your computer's file explorer and **double-click it** to open it in your browser (e.g. Chrome or Safari).

### Step 2: Copy-Paste Data from Google Sheets
To import your numbers from Google Sheets into the browser file:
1. Open your Google Sheet.
2. Go to the **`META_DATA`** tab.
3. Click anywhere in the sheet, press **`Cmd+A`** (Mac) or **`Ctrl+A`** (Windows) to select all cells, and copy them (**`Cmd+C`** / **`Ctrl+C`**).
4. Go back to the browser window running `index.html`.
5. Click **`📥 Data Import`** in the sidebar.
6. Paste (**`Cmd+V`** / **`Ctrl+V`**) the copied data into the **META_DATA Textarea**.
7. **Repeat this copy-paste process** for the remaining tabs:
   * `GOOGLE_DATA`
   * `SALES_DATA`
   * `PRODUCT_MASTER`
   * `CONFIG`
8. Click the green **`💾 Save & Update Dashboard`** button at the bottom of the page.

> 💡 **Tip:** Your browser caches this data locally (using `localStorage`). The next time you open `index.html`, your dashboard will load instantly with your last saved data!

---

## 🧠 Navigation & Features (Layman's Guide)

### 👑 CEO Command
* View your 4 key scorecards: **Blended MER** (ROAS), **Total Revenue**, **Total Media Spend**, and **Blended CAC** (Customer Acquisition Cost).
* Displays a quick visual strip of your marketing funnel conversion rates (Impressions ➔ Clicks ➔ Landing Page Views ➔ Add to Cart ➔ Conversions).

### 🤖 AI Smart Optimizer
* **Top/Bottom Ad Sets**: Highlights the top 5 and bottom 5 performing ad sets based on ROAS (only counts campaigns with spend > ₹500 to filter out noise).
* **Creative Messaging Narrative & Language Optimizer**: Tells you exactly which narrative angles (e.g. *Offer-led*, *Founder Story*) and ad languages (e.g. *Hindi*, *Hinglish*, *English*) are working (ROAS ≥ target) or underperforming (ROAS < target).

### 📈 Trends
* Visualizes performance over time.
* Shows daily funnel flows and Meta vs Google revenue splits.

### 📅 Daily Stats Ledger (New!)
* Shows a clean chronological daily timeline table detailing daily media Spend, total Conversions, CAC, Average Order Value (AOV), and ROAS.
* Use this to track day-to-day scaling pacing and efficiency drops.

### 👥 LTV & Cohorts
* Displays customer lifetime repeat purchase behaviors.
* Heatmap format shows cohort retention decay and customer LTV growth month-over-month.

---

## ⚙️ How to Tune Thresholds

Open the **`CONFIG`** sheet tab in Google Sheets (or edit the Config textarea in `index.html` data import):
* **Target ROAS**: Tells the system when to recommend scaling (green) or stopping (red) campaigns.
* **Min Spend**: Filters out tiny test ads from bloating recommendations.
* **Hook/Hold Rate benchmarks**: Adjusts the flags highlighting whether video hooks or visual assets are fatiguing.
