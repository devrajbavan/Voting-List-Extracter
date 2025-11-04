# 🧾 Voter ID Data Extraction API (Marathi + English)

### ⚙️ Overview

This project is a **FastAPI-based OCR automation system** that processes scanned voter list sheets, automatically **splits them into individual voter cards**, performs **OCR text extraction** in Marathi and English, **cleans and structures** the extracted data, **crops voter photos**, and finally **generates a formatted Excel sheet** — where **each voter record is represented in a single row with its corresponding face image**.

All temporary files and folders are auto-deleted after processing, keeping the workspace clean and efficient.

---

## 🚀 Features

✅ Upload a full voter list image (Marathi + English text)  
✅ Automatically splits the sheet into individual voter cards  
✅ Performs OCR using Tesseract in parallel for speed  
✅ Extracts & cleans:
- Voter ID  
- Full name  
- Relative name (Father/Husband)  
- House number  
- Age  
- Gender  
✅ Crops voter photos and embeds them directly in Excel  
✅ Each voter = **1 row** in Excel (horizontal layout)  
✅ Automatically adjusts cell height to match image size  
✅ Auto-deletes temporary files after completion  

---

## 🧱 Tech Stack

| Component | Purpose |
|------------|----------|
| **FastAPI** | Web framework for API routes |
| **Uvicorn** | ASGI server to host the FastAPI app |
| **Pillow (PIL)** | Image cropping, enhancement, and resizing |
| **pytesseract** | OCR engine (supports Marathi + English) |
| **openpyxl** | Excel file creation, styling, and image embedding |
| **Regex (re)** | Text cleanup and field extraction |
| **BackgroundTasks (FastAPI)** | Async cleanup after response |
| **shutil / pathlib / uuid** | File management and cleanup utilities |
| **concurrent.futures** | Parallel OCR for performance boost |

---

## 🧩 Project Structure

📁 Voter-OCR-API
│
├── main.py # Complete FastAPI + OCR + Excel logic
├── requirements.txt # Dependencies list
├── README.md # Documentation (this file)
└── uploads/ # Temporary folder (auto-created & cleaned)

yaml
Copy code

---

## ⚙️ Setup Instructions

### 1️⃣ Clone the Repository
```bash
git clone https://github.com/devrajbavan/Voting-List-Extracter.git
cd Voting-List-Extracter
2️⃣ Create a Virtual Environment
bash
Copy code
python -m venv venv
# Activate
venv\Scripts\activate      # Windows
source venv/bin/activate   # Linux / macOS
3️⃣ Install Dependencies
bash
Copy code
pip install -r requirements.txt
4️⃣ Install Tesseract OCR
🔹 Windows:
Download and install from Tesseract OCR GitHub Releases.

Set the path in main.py:

python
Copy code
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
🔹 Ubuntu / Linux:
bash
Copy code
sudo apt update
sudo apt install tesseract-ocr tesseract-ocr-mar
🔹 macOS (Homebrew):
bash
Copy code
brew install tesseract
📦 requirements.txt
txt
Copy code
fastapi==0.115.2
uvicorn[standard]==0.30.1
python-multipart==0.0.9
pillow==10.4.0
pytesseract==0.3.13
openpyxl==3.1.5
▶️ Run the Application
bash
Copy code
python main.py
or

bash
Copy code
uvicorn main:app --reload
Then open in browser:

arduino
Copy code
http://127.0.0.1:8000/docs
Use the interactive Swagger UI to upload the voter list image and download the generated Excel report.

🧠 How It Works (Step-by-Step)
1️⃣ Upload

User uploads a scanned voter list image (.jpg / .png)

2️⃣ Split Image into Cards

Large sheet is divided into multiple small voter cards
via crop_all_cards_from_sheet_bytes()

3️⃣ Parallel OCR + Cleaning

Each card is preprocessed and OCR’d in parallel using ProcessPoolExecutor.
Extracted text is cleaned via regex-based functions:

python
Copy code
clean_voter_name(), clean_relative_name(), clean_age(), clean_house()
4️⃣ Face Extraction

Each voter’s photo is cropped using ratio-based coordinates via:

python
Copy code
crop_person_face_bytes()
5️⃣ Excel Report Generation

A tabular Excel report is generated using openpyxl:
Each voter = one row, with text + embedded photo.

6️⃣ Auto Cleanup

Temporary folders (uploads/<uuid>) are deleted asynchronously after 10 seconds using:

python
Copy code
cleanup_run_dir()
📊 API Endpoint Details
POST /process-voters/
Request:
file: image/* (.jpg, .jpeg, .png)

Response:
Returns a downloadable .xlsx Excel file.

Example using curl:

bash
Copy code
curl -X POST "http://127.0.0.1:8000/process-voters/" \
     -F "file=@/path/to/voters.jpg" \
     -o result.xlsx
🧹 Automatic Cleanup Logic
After generating the Excel file:

Each upload is stored under uploads/<uuid>/

A background task waits 10 seconds

Then deletes that directory safely using:

python
Copy code
shutil.rmtree(run_dir)
Ensures clean disk usage after each run.

🧠 Core Functions Overview
Function	Description
preprocess_for_ocr()	Enhances image before OCR
ocr_card_text_bytes()	Extracts raw text from in-memory card image
clean_*()	Cleans and normalizes Marathi/English OCR text
parse_card()	Extracts structured voter data from text
crop_all_cards_from_sheet_bytes()	Crops the main sheet into in-memory card images
crop_person_face_bytes()	Crops voter’s face image
generate_excel_from_cards()	Generates Excel with one voter per row and images auto-sized
cleanup_run_dir()	Deletes temporary directories
/process-voters/	Orchestrates the full workflow

🧩 Process Workflow Diagram
mermaid
Copy code
graph TD
A[📤 Upload Voter Sheet Image] --> B[🧩 Split into Individual Cards]
B --> C[⚙️ Parallel OCR + Text Cleaning]
C --> D[🖼️ Face Cropping]
D --> E[📊 Excel Generation (Row-wise Layout)]
E --> F[⬇️ File Download]
F --> G[🧹 Background Cleanup (10s Delay)]
📘 Example Excel Output
S.No.	ID	Serial	मतदाराचे पूर्ण:	पतीचे नाव / वडिलांचे नाव	घर क्रमांक :	वय :	लिंग :	Face image
9	XYZ12345 01/01/1990	9	राम शिंदे	गणेश शिंदे	६७	32	पुरुष	🖼️ (Auto-sized image)
10	XYZ12346 03/01/1988	10	सीमा शिंदे	राजेश शिंदे	८५	36	स्त्री	🖼️ (Auto-sized image)

🛡️ Notes & Warnings
⚠️ This API is for controlled environments — add authentication & rate limiting before public deployment.
⚠️ Adjust cropping ratios (FACE_*_RATIO) according to your voter card layout.
⚠️ OCR accuracy depends heavily on image clarity and proper Marathi training data.
⚠️ Ensure mar.traineddata is installed in your Tesseract directory.

🧰 Future Enhancements
🧠 Integrate OpenCV face detection for automatic face bounding

📦 Add ZIP upload support for batch sheets

⏱️ Include progress tracking & OCR metrics

☁️ Cloud integrations (Google Drive, Dropbox)

🐳 Dockerize for containerized deployment

👨‍💻 Author
Devraj Bavan
AI & Software Engineer | OCR, Computer Vision, Web Automation
📧 [Contact for collaborations or improvements]

🏁 License
Licensed under the MIT License — free for personal and commercial use.

markdown
Copy code

---

✅ **What’s Updated Here:**
- Reflects **row-wise Excel layout** (one record per row).
- Mentions **auto image resizing**.
- Notes **parallel OCR optimization**.
- Updated **workflow diagram** and **example output table**.
- Corrected folder names and consistent formatting for GitHub.

Would you like me to add a short **project badge section** (e.g., Python version, FastAPI version, license, etc.) at the top for GitHub visual appeal?