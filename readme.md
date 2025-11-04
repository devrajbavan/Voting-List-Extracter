🧾 Voter ID Data Extraction API (Marathi + English)
⚙️ Overview

This project is a FastAPI-based OCR automation system that processes scanned voter list sheets, automatically splits them into individual voter cards, performs parallel OCR (Marathi + English) for faster processing, cleans & structures the extracted text, crops voter faces, and finally generates a tabular Excel sheet — where each voter record occupies a single row with auto-sized face thumbnails.

All temporary images and folders are automatically deleted after successful completion, ensuring a clean environment.

🚀 Features

✅ Upload a full voter-sheet image (Marathi + English text)
✅ Automatically crops it into individual voter cards (in-memory, no disk saves)
✅ Performs parallel OCR on all cards to boost speed
✅ Extracts & cleans:

Voter ID

Full name

Relative name (Father/Husband)

House number

Age

Gender
✅ Crops voter photos from each card
✅ Generates a single-row-per-voter Excel sheet with properly aligned and visible images
✅ Automatically adjusts cell height and image size for each voter
✅ Cleans up temporary directories asynchronously after processing

🧱 Tech Stack
Component	Purpose
FastAPI	Web framework for API endpoints
Uvicorn	ASGI server to host FastAPI app
Pillow (PIL)	Image cropping, enhancement, resizing
pytesseract	OCR engine (Marathi + English)
openpyxl	Excel workbook creation and image embedding
Regex (re)	Text cleaning and data extraction
ProcessPoolExecutor	Parallel OCR execution for faster performance
BackgroundTasks (FastAPI)	Automatic cleanup of temporary folders
pathlib / shutil / uuid	File and directory management
🧩 Project Structure

📁 Voter-OCR-API
│
├── main.py – Complete FastAPI + OCR + Excel logic
├── requirements.txt – Dependencies list
├── README.md – Documentation (this file)
└── uploads/ – Temporary runtime folder (auto-created & cleaned)

⚙️ Setup Instructions
1️⃣ Clone the Repository
git clone https://github.com/devrajbavan/Voting-List-Extracter.git
cd voter-ocr-api

2️⃣ Create a Virtual Environment
python -m venv venv
# Activate it
venv\Scripts\activate      # On Windows

3️⃣ Install Dependencies
pip install -r requirements.txt

4️⃣ Install Tesseract OCR

🔹 Windows:
Download and install from
👉 Tesseract OCR GitHub Releases

Then set the correct path in main.py:

TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


🔹 Ubuntu / Linux:

sudo apt update
sudo apt install tesseract-ocr tesseract-ocr-mar


🔹 macOS (Homebrew):

brew install tesseract

📦 requirements.txt
fastapi==0.115.2
uvicorn[standard]==0.30.1
python-multipart==0.0.9
pillow==10.4.0
pytesseract==0.3.13
openpyxl==3.1.5

▶️ Run the Application
python main.py


or

uvicorn main:app --reload


Then open in your browser:

http://127.0.0.1:8000/docs


You’ll see an interactive Swagger UI where you can upload a voter list image and download the processed Excel file.

🧠 How It Works (Step-by-Step)
1️⃣ Upload

User uploads a scanned voter list image (.jpg, .jpeg, .png).

2️⃣ Split Image into Cards

The sheet is divided into smaller voter card regions in-memory using:

crop_all_cards_from_sheet_bytes()

3️⃣ Parallel OCR + Cleaning

Each card undergoes preprocessing and OCR extraction using:

Pillow (for image enhancement)

pytesseract (for OCR in Marathi + English)

Regex cleaners (for names, gender, etc.)

ocr_card_text_bytes()
parse_card()

4️⃣ Face Extraction

Each card’s photo is cropped using fixed ratios:

crop_person_face_bytes()

5️⃣ Excel Report Generation

All cleaned text + face images are compiled horizontally into one Excel sheet:

generate_excel_from_cards()


Each row = one voter record.
Row heights auto-adjust to match image size so all faces are clearly visible.

6️⃣ Cleanup

Temporary folders are deleted asynchronously:

cleanup_run_dir()

📊 API Endpoint Details
POST /process-voters/

Request:
file: image/* (.jpg, .jpeg, .png)

Response:
Returns a downloadable Excel file (.xlsx)

Example (curl):

curl -X POST "http://127.0.0.1:8000/process-voters/" \
     -F "file=@/path/to/voters.jpg" \
     -o result.xlsx

🧹 Automatic Cleanup Logic

After generating the Excel:

Uploads stored under uploads/<uuid>/

Background task waits 10 seconds

Then safely deletes the folder via:

shutil.rmtree(run_dir)


Keeps disk clean even after multiple uploads.

🧠 Internals Overview
Function	Description
preprocess_for_ocr()	Lightweight image enhancement before OCR
ocr_card_text_bytes()	OCR text extraction from in-memory images
clean_voter_name() / clean_relative_name()	Name cleanup (Marathi support)
clean_age() / clean_house() / clean_gender()	Numeric and gender normalization
parse_card()	Extracts structured voter data from OCR text
crop_all_cards_from_sheet_bytes()	Divides sheet image into card buffers
crop_person_face_bytes()	Crops voter’s face region (in-memory)
generate_excel_from_cards()	Creates final Excel with visible images & auto row sizing
cleanup_run_dir()	Deletes temporary data asynchronously
/process-voters/	Main API route coordinating the workflow
🧩 Process Workflow Diagram
graph TD
A[📤 Upload Voter Sheet Image] --> B[🧩 Split into Cards (in-memory)]
B --> C[⚙️ Parallel OCR + Text Cleaning]
C --> D[🖼️ Face Cropping]
D --> E[📊 Excel Generation (1 Row per Voter)]
E --> F[⬇️ FileResponse Download]
F --> G[🧹 Background Cleanup (10s)]

📘 Example Output (Excel)
S.No.	ID	Serial	मतदाराचे पूर्ण:	पतीचे नाव / वडिलांचे नाव	घर क्रमांक :	वय :	लिंग :	Face image
9	XYZ12345 01/01/1990	9	राम शिंदे	गणेश शिंदे	६७	32	पुरुष	🖼️ (Visible Image)
10	ABC78945 03/02/1988	10	सीमा पाटील	राजेश पाटील	३५६	३८	स्त्री	🖼️ (Visible Image)
🛡️ Notes & Warnings

⚠️ Add authentication & rate limits before public deployment.
⚠️ Tune cropping ratios (FACE_*_RATIO) for your card layout.
⚠️ OCR accuracy depends on image clarity and traineddata quality.
⚠️ Marathi (mar.traineddata) must exist in your Tesseract path.

🧰 Future Enhancements

Integrate OpenCV face detection (auto detect faces)

Add support for ZIP uploads (batch voter lists)

Include progress tracking via WebSocket

Add Google Drive / Dropbox output integration

Dockerize API for one-command deployment

👨‍💻 Author

Devraj Bavan
AI & Software Engineer | OCR, Computer Vision, and Web Automation
📧 [Contact for collaborations or improvements]

🏁 License

This project is licensed under the MIT License — free for personal and commercial use.