# AppRiver Invoice → Excel Converter

This project converts AppRiver invoice PDFs into structured Excel (`.xlsx`) files.
It is designed for **non-technical end users** via a simple desktop application
(Windows `.exe` and macOS `.app`).

---

## What this tool does

- Reads AppRiver invoice PDFs
- Extracts:
  - Client
  - Product description
  - Term (Monthly / Annual)
  - Period
  - Quantity
  - Subtotal
  - Total
  - Total Current Charges Detail
- Outputs a formatted Excel file (EUR currency)
- Handles:
  - Renewals
  - Activations
  - Adjustments (highlighted in red)
- Works fully offline

---

## Repository contents

.
├── main.py # Application source code
├── README.md # This file

yaml
Copy code

No virtual environments, binaries, or build artifacts are committed.

---

## Prerequisites

### Windows
- Windows 10 / 11
- Python **3.11** (64-bit recommended)
- Internet access (for initial dependency install)

### macOS
- macOS 12+
- Python **3.10+** (`python3`)
- Xcode Command Line Tools installed

---

## Build instructions

> IMPORTANT:
> - Windows builds must be done on **Windows**
> - macOS builds must be done on **macOS**
> - Cross-compilation is **not supported**

---

## Windows – Build `.exe`

### 1. Create and activate virtual environment

`python -m venv .venv
powershell -ExecutionPolicy Bypass -NoProfile
.\.venv\Scripts\Activate.ps1`

### 2. Install dependencies
`python -m pip install --upgrade pip
python -m pip install pdfplumber openpyxl pyinstaller`

### 3. Build executable
`python -m PyInstaller --onefile --windowed --name "PDF_to_Excel" main.py`

### 4. Result
`dist/PDF_to_Excel.exe`

This is the only file to distribute to end users.


## Running the application (end users)
- Double-click the application

- Click Select PDF and Convert

- Choose an AppRiver invoice PDF

- Excel file is generated next to the PDF


For enterprise deployment, the app can be code-signed and notarized.

# Security / Endpoint Protection
On corporate machines:

*The executable may need to be whitelisted (e.g. ThreatLocker / Defender / AppLocker)*

Prefer path-based allow rules for easier updates

### Updating the application

Any change to main.py requires:
Rebuilding the .exe (Windows)
Distribute the new binary and replace the old one.

Notes
The application runs fully offline

No user data is sent anywhere

Excel output uses EUR currency formatting

Adjustment rows are highlighted in red
