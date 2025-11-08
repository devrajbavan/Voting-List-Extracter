# 🗳️ Voter Details Extraction API & Script (Marathi OCR)

This repository contains two integrated Python-based tools for **extracting structured voter details (Marathi + English)** from scanned sheets of voter ID cards.  
It supports both **standalone local script execution (`app.py`)** and **API-based automation (`API.py`)** built with **FastAPI**.

---

## 📋 Project Overview

### 1️⃣ `app.py` — Standalone Script
- Input: Large image (`voters.jpg`) containing multiple voter cards in a grid layout (default: 3×10).
- Automatically:
  - Crops each card from the sheet.
  - Extracts text using Tesseract OCR (`mar+eng`).
  - Parses details like voter name, relative name, house number, age, and gender.
  - Crops the voter’s face photo from each card.
  - Generates a clean, formatted Excel sheet with all extracted data and thumbnails.
- Cleans up all temporary files automatically.

### 2️⃣ `API.py` — REST API using FastAPI
- Provides an HTTP interface for uploading voter sheet images.
- Automatically:
  - Crops all cards.
  - Extracts voter details in parallel.
  - Generates and returns a downloadable Excel file.
  - Handles temporary file cleanup asynchronously.
- Perfect for web integration or automation pipelines.

---

## 🧩 Features

- **Marathi + English OCR** (via Tesseract)
- **Parallel OCR processing** (multi-core optimized)
- **Face extraction** from each card
- **Auto Excel generation** with proper fonts (`Mangal`)
- **Asynchronous cleanup**
- **Web API endpoints for integration**

---

## ⚙️ System Requirements

| Component | Description |
|------------|-------------|
| **OS** | Windows / Linux (tested on Windows 10, Ubuntu 22.04) |
| **Python** | ≥ 3.9 |
| **Tesseract OCR** | Installed and accessible via system PATH |
| **RAM** | Minimum 8 GB recommended for large image sheets |
| **Processor** | Multi-core CPU for faster OCR (uses ProcessPoolExecutor) |

---

## 🧰 Required Tools & Libraries

Install all Python dependencies using:

```bash
pip install fastapi uvicorn pillow pytesseract openpyxl
Additional Setup
Install Tesseract OCR

Windows default path:
C:\Program Files\Tesseract-OCR\tesseract.exe

Linux installation:

bash
Copy code
sudo apt update
sudo apt install tesseract-ocr tesseract-ocr-mar
Verify installation

bash
Copy code
tesseract --version
Ensure Marathi language pack is available:

bash
Copy code
tesseract --list-langs
Should display: mar, eng

🧠 Directory Structure
graphql
Copy code
project/
│
├── API.py                # FastAPI-based OCR API
├── app.py                # Standalone batch OCR + Excel generator
├── uploads/              # Temporary upload folder (auto-created)
├── generated_excels/     # Folder for generated Excel files
├── temp/                 # Temporary cropped image storage
├── voters.jpg            # Example source image (input)
└── README.md
🚀 1. Running app.py (Standalone Script)
🔧 Step-by-Step Setup
Place your source sheet image as voters.jpg in the project directory.

Must contain multiple voter cards arranged in a grid (default: 3×10).

Open app.py and confirm this configuration:

python
Copy code
IMG_SOURCE = "voters.jpg"
DEFAULT_COLS = 3
DEFAULT_ROWS = 10
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
Run the script:

bash
Copy code
python app.py
After execution:

voter_data.xlsx will be generated in the same folder.

All cropped temporary files will be deleted automatically.

Console output will confirm OCR progress, Excel save path, and cleanup status.

🧾 Example Output
pgsql
Copy code
Cropped 30 cards into 'temp'.
Excel saved: voter_data.xlsx
Processed: 30 voters. First serial: 9
🧹 Cleaned up temporary folder: temp
🌐 2. Running API.py (FastAPI Application)
🔧 Step-by-Step Setup
Ensure all dependencies are installed.

Start the API server:

bash
Copy code
uvicorn API:app --host 0.0.0.0 --port 8000 --reload
or simply:

bash
Copy code
python API.py
Access FastAPI interactive docs:

arduino
Copy code
http://127.0.0.1:8000/docs
Upload your .jpg or .png voter sheet using the /process-voters/ endpoint.

✅ Process Flow
Image uploaded → Cards cropped → OCR processed in parallel
→ Data parsed → Excel generated → Download link returned.

Example JSON response:

json
Copy code
{
  "status": "success",
  "file_name": "voter_data_63e7fa98b0a14b96b4c4.xlsx",
  "download_url": "/download/63e7fa98b0a14b96b4c4"
}
To download Excel:

arduino
Copy code
http://127.0.0.1:8000/download/63e7fa98b0a14b96b4c4
⚡ Performance Tips
For better OCR accuracy:

Use high-quality scanned images (≥ 300 DPI).

Ensure clear text and consistent lighting.

To tune face cropping:

Adjust FACE_LEFT_RATIO, FACE_TOP_RATIO, etc. in API.py.

To control grid layout:

Change DEFAULT_COLS and DEFAULT_ROWS as per your sheet format.

🧹 Automatic Cleanup
Temporary directories under /uploads are deleted after 10 seconds.

The temp/ directory (for app.py) is cleared after Excel generation.

🧾 Output Excel Format
S.No.	ID	Serial	मतदाराचे पूर्ण	पतीचे नाव / वडिलांचे नाव	घर क्रमांक	वय	लिंग	Face image

Each row represents one voter card extracted from the image grid.
Fonts are set to Mangal for proper Marathi rendering.

🛠️ Troubleshooting
Problem	Possible Cause	Fix
TesseractNotFoundError	Wrong Tesseract path	Update TESSERACT_CMD
Poor OCR accuracy	Low image quality	Increase resolution / preprocess contrast
API 500 error	Invalid image input	Ensure correct .jpg/.png upload
Missing Marathi text	Marathi language pack not installed	sudo apt install tesseract-ocr-mar

🧑‍💻 Developer Notes
Modular design allows integrating OCR, parsing, and Excel generation as reusable functions.

Can be extended to:

JSON-only APIs (without Excel)

Database integration

Frontend upload portals

🧾 License
This project is provided for educational and automation purposes.
You are free to modify and extend it for personal or organizational use.

Author: Devraj Bavan
Version: 1.0
Language Support: Marathi + English
Frameworks: FastAPI, OpenPyXL, Tesseract OCR

yaml
Copy code

---

Would you like me to **add a “Usage Workflow Diagram” (Markdown + ASCII art)** showing the end-to-end process (Image → OCR → Parse → Excel)? It makes the README more visual and clear for presentation/documentation purposes.






