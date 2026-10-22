# VTU Result Automation 🎓

A simple tool that automatically fetches student results from the VTU website and saves them into **Excel**, **Word (DOCX)**, **PDF**, or **CSV** files.

No more manually typing captchas or checking results one by one!

---

## ⚡ Quick 3-Step Setup

### Step 1: Open Terminal / Command Prompt
Navigate to the project folder:
```bash
cd vtu-result-automation
```

### Step 2: Install Dependencies
Run this single command to install everything (works on Windows, Mac, and Linux):
```bash
pip install -r requirements.txt
```

### Step 3: Start the App
```bash
python src/main.py
```

---

## 🖥️ How to Use the App

When you launch the app, a simple window will appear with a few boxes to fill:

| Field | What to put in |
|---|---|
| **CHROMEDRIVER PATH** | *(Optional)* Leave blank if you have Chrome installed! Selenium will handle it automatically. |
| **USN FILE PATH** | Click **CHOOSE** and pick your file containing student USNs. Can be an **Excel (.xlsx/.xls)**, **Word (.docx)**, **PDF (.pdf)**, **CSV (.csv)**, or **Text (.txt)** file. |
| **WEBSITE ADDRESS** | Paste the VTU results link (e.g., `https://results.vtu.ac.in/MJ26cbcs/index.php`). |
| **SAVE PATH** | Click **CHOOSE** and pick where to save your output file. Choose your format by typing the extension: `.xlsx`, `.docx`, `.pdf`, or `.csv`. |
| **USN START / END (Row)** | *(Optional)* Choose which students to fetch (e.g., start at `1` and end at `60`). Leave blank to process **everyone** in the file! |
| **Auto-solve CAPTCHA** | Keep this **checked**! The tool will read and solve the captcha automatically without bothering you. |

Click **SUBMIT** and sit back! 🚀

---

## 📁 Supported Input Files (Where you keep your USNs)

You don't need to format your USN list in a special way. The tool automatically detects VTU USNs (like `1HC24CS058`, `1DS21IS020`, etc.) from:
- 📊 **Excel Files** (`.xlsx`, `.xls`)
- 📝 **Word Documents** (`.docx`) — tables or lists
- 📄 **PDF Files** (`.pdf`)
- 📑 **CSV Files** (`.csv`)
- 📄 **Plain Text Files** (`.txt`)

---

## 📊 Supported Output Files (How your results are saved)

Just choose your preferred extension when saving:
- **`.xlsx` (Excel)**: Full spreadsheet with subject names, marks (IA, SEE, Total), and colored Pass/Fail tags.
- **`.docx` (Word)**: A formatted departmental table ready to print or share.
- **`.pdf` (PDF)**: A landscape result sheet with Pass (green) and Fail (red) highlights.
- **`.csv` (CSV)**: Raw data for data analysis or importing into other tools.

---

## ❓ Frequently Asked Questions

#### Do I have to enter the captcha manually?
**No.** The app takes a picture of the captcha, cleans up the background noise, reads the letters, and types it in for you.

#### What happens if the captcha fails?
If VTU rejects a captcha, the script automatically detects the alert, refreshes the captcha, and tries again up to 5 times.

#### Does this work on Windows, Mac, and Linux?
**Yes.** All required tools (including the OCR captcha reader) are bundled inside `requirements.txt`.

#### Can I run it without the GUI?
Yes! For quick batch runs from the terminal, run:
```bash
python src/temp_test_runner.py
```
This automatically processes target USNs and exports `.xlsx`, `.docx`, `.pdf`, and `.csv` files all at once.

---

## 👤 Credits

Originally created by **Samarth Kashyap** (Department of CSE).  
Updated with automatic CAPTCHA solving and universal document format support (DOCX, PDF, CSV, Excel).
