# 🧾 Voter ID Data Extraction API (Marathi + English)

### ⚙️ Overview

This project is a **FastAPI-based OCR automation system** that processes scanned voter list sheets, automatically **splits** them into individual voter cards, performs **text extraction (OCR)** in Marathi and English, **cleans & structures** the data, **extracts voter photos**, and finally **generates a formatted Excel sheet** containing all voter details with embedded thumbnails.

All temporary images and folders are automatically deleted after successful completion, ensuring a clean environment.

---

## 🚀 Features

✅ Upload a full voter-sheet image (Marathi + English text)  
✅ Automatically crops it into individual voter cards  
✅ Performs OCR on each card using Tesseract  
✅ Extracts & cleans:
- Voter ID  
- Full name  
- Relative name (Father/Husband)  
- House number  
- Age  
- Gender  
✅ Crops voter photos from each card  
✅ Embeds both text and image data into an **Excel report**  
✅ Auto-deletes temporary files after processing  

---

## 🧱 Tech Stack

| Component | Purpose |
|------------|----------|
| **FastAPI** | Web framework for API endpoints |
| **Uvicorn** | ASGI server to host FastAPI app |
| **Pillow (PIL)** | Image processing (crop, enhance, resize) |
| **pytesseract** | OCR engine (Marathi + English) |
| **openpyxl** | Excel workbook creation and image embedding |
| **Regex (re)** | Cleans OCR text and extracts structured info |
| **BackgroundTasks (FastAPI)** | Cleans temporary directories post-response |
| **shutil / pathlib / uuid** | File management and safe cleanup |

---

## 🧩 Project Structure

📁 Voter-OCR-API
│
├── main.py # Full API script (FastAPI + OCR + Excel)
├── requirements.txt # Dependencies list
├── README.md # Documentation (this file)
└── uploads/ # Temporary folder (auto-created & cleaned)

yaml
Copy code

---

## ⚙️ Setup Instructions

### 1️⃣ Clone the Repository
```bash
git clone https://github.com/<your-username>/voter-ocr-api.git
cd voter-ocr-api
2️⃣ Create a Virtual Environment
bash
Copy code
python -m venv venv
# Activate it
venv\Scripts\activate      # On Windows
source venv/bin/activate   # On macOS/Linux
3️⃣ Install Dependencies
bash
Copy code
pip install -r requirements.txt
4️⃣ Install Tesseract OCR
🔹 Windows:
Download and install from
👉 Tesseract OCR GitHub Releases

Then, set the correct path in main.py:

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
Then open your browser:

arduino
Copy code
http://127.0.0.1:8000/docs
You’ll see an interactive Swagger UI where you can upload the voter list image and download the processed Excel file.

🧠 How It Works (Step-by-Step)
1️⃣ Upload
User uploads the scanned voter list image (.jpg, .png).

2️⃣ Split Image into Cards
The script divides the large sheet into smaller images using:

python
Copy code
crop_all_cards_from_sheet()
3️⃣ OCR + Cleaning
Each card is processed using:

Pillow (for preprocessing)

pytesseract (for OCR)

Regex-based cleaners (clean_voter_name, clean_age, etc.)

4️⃣ Face Extraction
Each card’s photo area is cropped using ratio-based coordinates:

python
Copy code
crop_person_face()
5️⃣ Excel Report Generation
All cleaned data + images are compiled into an Excel file using openpyxl:

python
Copy code
generate_excel_from_cards()
6️⃣ Cleanup
Temporary directories are deleted asynchronously:

python
Copy code
cleanup_run_dir()
📊 API Endpoint Details
POST /process-voters/
Request:
File: image/* (.jpg, .jpeg, .png)

Response:
Excel file (.xlsx) ready for download

Example using curl:
bash
Copy code
curl -X POST "http://127.0.0.1:8000/process-voters/" \
     -F "file=@/path/to/voters.jpg" \
     -o result.xlsx
🧹 Automatic Cleanup Logic
After generating the Excel:

Each upload is stored under uploads/<uuid>/

A background task waits 10 seconds

Then removes the entire folder safely using:

python
Copy code
shutil.rmtree(run_dir)
So, disk usage stays clean even after multiple uploads.

🧠 Internals Overview
Function	Description
preprocess_for_ocr()	Enhances image before OCR
ocr_card_text()	Extracts raw text from card
clean_voter_name() / clean_relative_name()	Sanitizes Marathi names
clean_age() / clean_house() / clean_gender()	Converts and normalizes fields
parse_card()	Extracts structured voter data using regex
crop_all_cards_from_sheet()	Divides big sheet into card images
crop_person_face()	Crops voter’s face from each card
generate_excel_from_cards()	Creates the Excel output file
cleanup_run_dir()	Deletes temporary files asynchronously
/process-voters/ (FastAPI route)	Orchestrates the entire workflow

🧩 Process Workflow Diagram
mermaid
Copy code
graph TD
A[📤 Upload Voter Sheet Image] --> B[🧩 Split into Cards]
B --> C[🔍 OCR + Text Cleaning]
C --> D[🖼️ Face Cropping]
D --> E[📊 Excel Generation]
E --> F[⬇️ FileResponse Download]
F --> G[🧹 Background Cleanup (10s)]
📘 Example Output (Excel)
क्रमांक	मतदार ID	मतदाराचे पूर्ण नाव	पतीचे/वडिलांचे नाव	घर क्रमांक	वय	लिंग	छायाचित्र
9	XYZ12345 01/01/1990	राम शिंदे	गणेश शिंदे	६७	32	पुरुष	🖼️ (Image)

🛡️ Notes & Warnings
⚠️ Do not expose publicly — add authentication & rate limits before deployment.
⚠️ Adjust cropping ratios (FACE_*_RATIO) to match your card layout.
⚠️ OCR accuracy depends heavily on image clarity and Tesseract training data.
⚠️ Marathi (mar.traineddata) must be installed in your Tesseract path.

🧰 Future Enhancements
 Integrate OpenCV face detection (auto face area detection)

 Add support for bulk ZIP uploads

 Include real-time progress tracking

 Add Google Drive / Dropbox upload integration

 Dockerize the API for one-command deployment

👨‍💻 Author
Devraj Bavan

AI & Software Engineer | OCR, Computer Vision, and Web Automation
📧 [Contact for collaborations or improvements]

🏁 License
This project is licensed under the MIT License — free for personal and commercial use.