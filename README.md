# ✂️ AI Cut-Plan Optimizer — Diamond Fabrics
## Complete Setup & Deployment Guide

---

## 📦 Project Structure

```
diamond-cutplan/
├── app.py                    ← Main Streamlit application
├── requirements.txt          ← Python dependencies
├── .streamlit/
│   └── config.toml           ← Dark theme + server config
└── README.md                 ← This file
```

---

## 🖥️ LOCAL SETUP (Run on Your PC)

### Step 1: Install Python
Download Python 3.10+ from https://python.org/downloads/

### Step 2: Create project folder
```bash
mkdir diamond-cutplan
cd diamond-cutplan
```

### Step 3: Copy files
Place `app.py`, `requirements.txt`, and `.streamlit/config.toml` in the folder.

### Step 4: Install dependencies
```bash
pip install -r requirements.txt
```

### Step 5: Run the app
```bash
streamlit run app.py
```

App opens automatically at: **http://localhost:8501**

---

## ☁️ FREE HOSTING ON STREAMLIT CLOUD

### Step 1: Create accounts (both free)
- GitHub: https://github.com/signup
- Streamlit Cloud: https://share.streamlit.io

### Step 2: Push code to GitHub
```bash
git init
git add .
git commit -m "Diamond Fabrics AI Cut Plan Optimizer"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/diamond-cutplan.git
git push -u origin main
```

### Step 3: Deploy on Streamlit Cloud
1. Go to https://share.streamlit.io
2. Click **"New app"**
3. Select your GitHub repo: `diamond-cutplan`
4. Main file: `app.py`
5. Click **"Deploy!"**

✅ Your app will be live at:
`https://YOUR_USERNAME-diamond-cutplan-app-XXXX.streamlit.app`

**Completely FREE. No credit card. No expiry.**

---

## 📄 HOW TO USE THE APP

### Option A: Upload a Real PDF
1. Open the app
2. Adjust **Fabric Width**, **Shrinkage %**, and **Max Plies** in the sidebar
3. Click **"Browse files"** → upload Diamond Fabrics Order PDF
4. App auto-extracts: Order No, Style, Size-wise quantities
5. Verify/edit quantities in the **"Edit Order Data"** expander
6. Download your files instantly

### Option B: Demo Mode (No PDF needed)
1. Check **"Use Demo Order"** in the sidebar
2. A sample 2,400-piece order loads automatically
3. Explore all features and exports

---

## 📥 EXPORT FILES EXPLAINED

### 1. AccuMark CSV (Gerber Ready)
**Headers:** `ORDER_NUMBER, STYLE_CODE, MARKER_NAME, SIZE, QUANTITY, PLIES, FABRIC_WIDTH, MARKER_LENGTH, SHRINK_L, SHRINK_W, DATE`

**Usage in Gerber AccuMark v14:**
1. Open AccuMark Storage Manager
2. Go to **Easy Order** → **Import**
3. Select the downloaded `.csv` file
4. Map columns if prompted (they're pre-mapped)
5. Click **Process** → Ready for cutter!

### 2. Professional Excel Cut Plan
- **Sheet 1 "Cut Plan":** Full visual cut plan in Sapphire/Diamond format
  - Order header block
  - Size-wise breakdown table
  - Marker plan with efficiency bars
  - Color-coded by efficiency (Green ≥90%, Amber ≥85%, Red <85%)
- **Sheet 2 "AccuMark Data":** Raw tabular data for ERP/system import

### 3. JSON Plan Data
Machine-readable full plan for integration with ERP, SAP, or custom systems.

---

## 🧠 AI ENGINE LOGIC

### Marker Ratio Optimisation
1. Calculates **GCD** of all size quantities
2. Divides each size quantity by GCD to get base ratios
3. If total ratio sum ≤ 8 → **Single marker** (all sizes in one pass)
4. If total ratio sum > 8 → **Split into 2 markers** (by size grouping)
5. Assigns **plies = GCD** (capped at Max Plies setting)

### Shrinkage Adjustment
```
Effective Width  = Fabric Width ÷ (1 + Shrink_Width% / 100)
Garment Length   = Base Length × (1 + Shrink_Length% / 100)
```

### Marker Efficiency Estimation
```
Efficiency = min(88 + ratio_sum × 0.4, 96)%
```
*(Based on industry averages; actual efficiency depends on pattern shapes in AccuMark)*

### Fabric Consumption
```
Marker Length = Garment Length × Max Ratio in Marker + 5cm (seam allowance)
Total Fabric  = Marker Length / 100 × Plies  (in meters)
```

---

## 🔧 CUSTOMISATION

### Add more sizes (e.g., 28, 30, 48)
In `app.py`, line 14:
```python
SIZE_COLS = ["28", "30", "32", "34", "36", "38", "40", "42", "44", "46", "48"]
```
And update `SIZE_CONSUMPTION` dict below it.

### Change company name / logo
Edit the `top-header` HTML block and footer strings in `app.py`.

### Add more PDF patterns
In `parse_pdf()` function, add regex patterns to the `for pat in [...]` lists.

---

## 🛠️ TROUBLESHOOTING

| Issue | Fix |
|-------|-----|
| PDF quantities show 0 | Use the Edit panel to type quantities manually |
| PDF parse errors | Check PDF is text-based (not scanned image) |
| Streamlit Cloud deploy fails | Ensure `requirements.txt` is in root folder |
| Excel file doesn't open | Update Microsoft Office or use LibreOffice |
| AccuMark import fails | Verify headers match your AccuMark version's Easy Order template |

---

## 📞 SUPPORT

Built for **Diamond Fabrics (Pvt) Ltd — Ferozewala, Pakistan**
Compatible with **Gerber AccuMark v14** Easy Order Import

*For AccuMark-specific column mapping, consult your Gerber support rep or
the AccuMark v14 User Guide → Chapter 7: Easy Order.*
