# 🚀 Push to Google Sheets via CLI (clasp)
### No copy-pasting code manually — do it in 3 terminal commands

---

## 📦 What is clasp?
**clasp** = Command Line Apps Script  
It's Google's official tool that lets you push code files from your laptop directly into Google Apps Script — like uploading files, but for code.

---

## ✅ One-Time Setup (do this only once)

### 1. Open Terminal
On your Mac: Press `Cmd + Space` → type `Terminal` → press Enter

### 2. Install clasp
Type this and press Enter:
```
npm install -g @google/clasp
```
(This installs clasp on your computer. Takes ~30 seconds.)

### 3. Login to your Google account
```
clasp login
```
Your browser will open → sign in with the same Google account that has your Sheet → click Allow.

---

## 🗂️ Link Your Script ID (One Time Per Project)

### 4. Get Your Script ID
- Open your Google Sheet → `Extensions → Apps Script`
- In the editor, click ⚙️ **Project Settings** (left sidebar gear icon)
- Copy the **Script ID** (looks like: `1BxSgXKnKpJ3H2...`)

### 5. Paste it into `.clasp.json`
Open the file at:
```
Documents/D2C Google Sheet Dashboard Builder/.clasp.json
```
Replace `PASTE_YOUR_SCRIPT_ID_HERE` with your actual Script ID. Save the file.

---

## 🚢 Push Your Code (Every Time You Make Changes)

### 6. Navigate to Your Project Folder
In Terminal, type:
```
cd ~/Documents/D2C\ Google\ Sheet\ Dashboard\ Builder
```

### 7. Push All Files
```
clasp push
```
That's it! All 3 files (Code.gs, SampleData.gs, Dashboard.html) get uploaded instantly.

You'll see:
```
└─ appsscript.json
└─ Code.gs
└─ SampleData.gs
└─ Dashboard.html
Pushed 4 files.
```

---

## 🔄 If You Make Changes Later
Just run these 2 commands:
```
cd ~/Documents/D2C\ Google\ Sheet\ Dashboard\ Builder
clasp push
```
Done. No manual copy-paste ever again.

---

## 📤 Sharing with Someone Else
If you want to share this dashboard with a colleague or client:
1. Share the Google Sheet with them (normal Share button)
2. OR share a copy of your Sheet: `File → Make a Copy` → share that copy

---

## 🆘 If Something Goes Wrong
Run this to check everything is connected:
```
clasp status
```
It will show you which files are linked and ready to push.

---

> **Summary in one line:**
> Install clasp → Login → Paste Script ID → `clasp push` → you're live 🎉
