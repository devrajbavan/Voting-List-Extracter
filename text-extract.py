# text-extract.py - Clean OCR + Exact Excel layout (Marathi)
import shutil
import os
import re
import time
from pathlib import Path
from PIL import Image, ImageEnhance
import pytesseract
from openpyxl import Workbook
from openpyxl.styles import Font
from concurrent.futures import ProcessPoolExecutor
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

# ---------------- CONFIG / CONSTANTS ----------------
MARATHI_FONT = Font(name="Mangal", size=11)
MARATHI_BOLD = Font(name="Mangal", size=11, bold=True)

TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

IMG_SOURCE   = "voters.jpg"                # source big sheet
IMG_DIR      = Path("temp")                # folder for cropped cards and faces
EXCEL_OUT    = "voter_data.xlsx"
START_SERIAL = 9

TESS_LANG = "mar+eng"

DEFAULT_COLS = 3
DEFAULT_ROWS = 10

THUMB_W, THUMB_H = 80, 90  # thumbnail pixel size used in excel

DEV_DIGITS = "०१२३४५६७८९"
ENG_DIGITS = "0123456789"

# ---------------- IMAGE CROP / PREPROCESS ----------------

def crop_all_cards():
    """
    Crop the main sheet (IMG_SOURCE) into card_{r}_{c}.png under IMG_DIR.
    Skips cropping if the target PNGs already exist.
    """
    IMG_DIR.mkdir(exist_ok=True)
    existing = list(IMG_DIR.glob("card_*.png"))
    if existing:
        print(f"'{IMG_DIR}' already contains cropped cards – skipping crop ({len(existing)} found).")
        return sorted(existing)

    big = Image.open(IMG_SOURCE)
    w, h = big.size
    cols, rows = DEFAULT_COLS, DEFAULT_ROWS
    cw, rh = w // cols, h // rows
    saved = []
    for r in range(rows):
        for c in range(cols):
            box = (c * cw, r * rh, (c + 1) * cw, (r + 1) * rh)
            card = big.crop(box)
            fn = IMG_DIR / f"card_{r+1:02d}_{c+1:02d}.png"
            card.save(fn, "PNG")
            saved.append(fn)
    print(f"Cropped {len(saved)} cards into '{IMG_DIR}'.")
    return sorted(saved)

def preprocess(img_path: Path) -> Image.Image:
    """
    Lighter preprocessing: convert to L, conditional resize, optional enhancement.
    Returns PIL.Image (grayscale) ready for OCR.
    """
    img = Image.open(img_path).convert("L")
    # apply enhancements if image is small or faint
    # keep them moderate to avoid expensive ops
    if img.width < 350:
        img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)
    # optional contrast/sharpness could be enabled conditionally
    img = ImageEnhance.Contrast(img).enhance(1.5)
    img = ImageEnhance.Sharpness(img).enhance(1.2)
    return img

def ocr_card(img_path: Path) -> str:
    """
    Performs OCR on a Path to a card image and returns extracted text.
    """
    try:
        img = preprocess(img_path)
        txt = pytesseract.image_to_string(img, lang=TESS_LANG, config="--psm 6")
        return txt
    except Exception as e:
        print(f"⚠️ OCR failed for {img_path}: {e}")
        return ""

# ---------------- CLEANERS & PARSERS ----------------

def clean_voter_name(raw: str) -> str:
    s = re.sub(r'[|¦\\\/<>]', ' ', raw)
    s = re.sub(r'\s+[A-Za-z]{1,3}\s*', ' ', s)
    s = re.sub(r'[=z&*]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    parts = s.split()
    if len(parts) > 4:
        s = ' '.join(parts[:4])
    return s

def clean_relative_name(raw: str) -> str:
    s = re.sub(r'[|¦\\\/<>]', ' ', raw)
    s = re.sub(r'\s+[A-Za-z]{1,3}\s*', ' ', s)
    s = re.sub(r'[=z&*]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    parts = s.split()
    if len(parts) > 3:
        s = ' '.join(parts[:3])
    return s

def clean_house(text: str) -> str:
    text = re.sub(r'[^\d०-९NAna/]', '', text)
    text = text.replace('NA', 'NA').strip()
    if not text or text == 'NA':
        return 'NA'
    m = re.search(r'[\d०-९]+', text)
    return m.group(0) if m else 'NA'

def clean_age(text: str) -> str:
    m = re.search(r'[\d०-९]+', text)
    if not m:
        return ""
    age = m.group(0)
    age = "".join(ENG_DIGITS[DEV_DIGITS.index(d)] if d in DEV_DIGITS else d for d in age)
    return age

def clean_gender(code: str) -> tuple:
    code = (code or "").strip().lower()
    if 'जी' in code or 'स्त्री' in code or 'स्री' in code:
        return "स्त्री", "महिला"
    return "पु", "पुरुष"

def parse_card(text: str) -> dict:
    """
    Parse OCR text and return structured fields.
    """
    data = {
        "ID": "", "VoterName": "", "RelationLabel": "", "RelationName": "",
        "House": "NA", "Age": "", "GenderCode": "", "GenderFull": ""
    }
    if not text:
        return data

    lines = [l.strip() for l in text.split('\n') if l.strip()]

    # ID (common on first line)
    if lines:
        m = re.search(r"([A-Z0-9]{5,})\s+(\d{1,2}/\d{1,2}/\d{2,4})", lines[0])
        if m:
            data["ID"] = f"{m.group(1)} {m.group(2)}"
            lines = lines[1:]

    # Voter name (preferred pattern)
    voter_pat = re.compile(r"(?:मतदार(?:ाचे)?\s*(?:पूर्ण\s*)?)\s*[:ः]?\s*(.+)", re.I)
    for i, line in enumerate(lines):
        m = voter_pat.search(line)
        if m:
            data["VoterName"] = clean_voter_name(m.group(1))
            del lines[i]
            break
    else:
        # fallback: any 'नाव' line not containing house/gender/age
        for i, line in enumerate(lines):
            if re.search(r"(पती|वडिल|आई|सासू|पत्नी|सून)", line):
                continue
            if 'नाव' in line and not re.search(r"(घर|लिंग|वय)", line):
                parts = re.split(r'[:ः]', line, maxsplit=1)
                candidate = parts[1] if len(parts) > 1 else re.split(r'नाव', line, maxsplit=1)[-1]
                data["VoterName"] = clean_voter_name(candidate)
                del lines[i]
                break

    # Relative
    rel_pat_pati   = re.compile(r"पतीचे\s*नाव\s*[:ः]?\s*(.+)", re.I)
    rel_pat_vadil  = re.compile(r"वडिल(?:ांचे|े)\s*नाव\s*[:ः]?\s*(.+)", re.I)
    for i, line in enumerate(lines):
        m_p = rel_pat_pati.search(line)
        if m_p:
            data["RelationLabel"] = "पतीचे नाव"
            data["RelationName"] = clean_relative_name(m_p.group(1))
            del lines[i]
            break
        m_v = rel_pat_vadil.search(line)
        if m_v:
            data["RelationLabel"] = "वडिलांचे नाव"
            data["RelationName"] = clean_relative_name(m_v.group(1))
            del lines[i]
            break

    # House
    m = re.search(r"घर\s*क्रमांक\s*[:ः]?\s*(.+)", text, re.I)
    if m:
        data["House"] = clean_house(m.group(1))

    # Age
    m = re.search(r"वय\s*[:ः]?\s*([\d०-९]+)", text, re.I)
    if m:
        data["Age"] = clean_age(m.group(1))

    # Gender
    m = re.search(r"लिंग\s*[:ः]?\s*([^\s\r\n]+)", text, re.I)
    if m:
        code_raw = m.group(1)
        data["GenderCode"], data["GenderFull"] = clean_gender(code_raw)
    else:
        data["GenderCode"], data["GenderFull"] = "पु", "पुरुष"

    return data

# ---------------- FACE CROP ----------------

def crop_person_photo(card_path: Path) -> Path:
    """
    Crop the voter photo (top-right corner) from a card image and return the saved path.
    """
    img = Image.open(card_path)
    w, h = img.size

    # tuned ratios (can be adjusted)
    left   = int(w * 0.78)
    top    = int(h * 0.30)
    right  = int(w * 0.98)
    bottom = int(h * 0.85)

    left, top = max(0, left), max(0, top)
    right, bottom = min(w, right), min(h, bottom)
    if right <= left or bottom <= top:
        return None

    face = img.crop((left, top, right, bottom))
    face = ImageEnhance.Contrast(face).enhance(1.2)
    face = ImageEnhance.Sharpness(face).enhance(1.3)

    out_path = card_path.parent / f"face_{card_path.stem}.png"
    face.save(out_path, "PNG")
    return out_path

# ---------------- EXCEL GENERATION ----------------

def generate_excel_from_cards(start_serial: int = START_SERIAL):
    """
    Main pipeline: crop (if needed), OCR (parallel), parse, create Excel (one row per voter),
    insert thumbnail images and auto-adjust row heights, then cleanup images.
    """
    # 1) Ensure cards exist (crop if necessary)
    card_files = crop_all_cards()  # returns sorted list of Path objects in IMG_DIR

    if not card_files:
        raise FileNotFoundError("No card images found and source crop failed.")

    # 2) OCR in parallel (preserve order using map)
    cpu_count = max(1, os.cpu_count() or 1)
    with ProcessPoolExecutor(max_workers=cpu_count) as executor:
        ocr_results = list(executor.map(ocr_card, card_files))

    # 3) Create workbook & header
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = [
        "S.No.", "ID", "Serial", "मतदाराचे पूर्ण:",
        "पतीचे नाव / वडिलांचे नाव", "घर क्रमांक :",
        "वय :", "लिंग :", "Face image"
    ]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(1, col_idx, header)
        cell.font = MARATHI_BOLD

    # 4) Fill rows
    row = 2
    serial = start_serial
    processed = 0

    for idx, card_path in enumerate(card_files):
        txt = ocr_results[idx] if idx < len(ocr_results) else ""
        d = parse_card(txt)

        # populate text columns
        ws.cell(row, 1, idx + 1).font = MARATHI_FONT               # S.No.
        ws.cell(row, 2, d.get("ID", "")).font = MARATHI_FONT       # ID
        ws.cell(row, 3, serial).font = MARATHI_FONT                # Serial
        ws.cell(row, 4, d.get("VoterName", "")).font = MARATHI_FONT
        ws.cell(row, 5, d.get("RelationName", "")).font = MARATHI_FONT
        ws.cell(row, 6, d.get("House", "")).font = MARATHI_FONT
        ws.cell(row, 7, d.get("Age", "")).font = MARATHI_FONT
        ws.cell(row, 8, d.get("GenderFull", "पुरुष")).font = MARATHI_FONT

        # crop face and embed image (if crop produced a file)
        face_path = None
        try:
            face_path = crop_person_photo(card_path)
        except Exception as e:
            face_path = None
            print(f"⚠️ Face crop failed for {card_path}: {e}")

        if face_path and face_path.exists():
            try:
                img_for_excel = XLImage(str(face_path))
                img_for_excel.width, img_for_excel.height = THUMB_W, THUMB_H
                cell_ref = f"I{row}"
                ws.add_image(img_for_excel, cell_ref)
                # set this row height to match image (pixels -> points approx)
                ws.row_dimensions[row].height = img_for_excel.height * 0.75
            except Exception as e:
                print(f"⚠️ Adding image failed for row {row}: {e}")

        row += 1
        serial += 1
        processed += 1

    # 5) Adjust column widths (image column I wider)
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == 'I':
            ws.column_dimensions['I'].width = max(18, int(THUMB_W * 0.14) + 2)
            continue
        max_len = max((len(str(c.value)) for c in col if c.value), default=0)
        ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

    wb.save(EXCEL_OUT)
    print(f"\nExcel saved: {EXCEL_OUT}")
    print(f"Processed: {processed} voters. First serial: {START_SERIAL}")

    # Cleanup temp images
    cleanup_images()

# ---------------- CLEANUP ----------------

def cleanup_images():
    """Delete temporary folder with cropped cards and faces."""
    try:
        if IMG_DIR.exists():
            # small delay so any file handles are released
            time.sleep(0.1)
            shutil.rmtree(IMG_DIR)
            print(f"🧹 Cleaned up temporary folder: {IMG_DIR}")
    except Exception as e:
        print(f"⚠️ Cleanup failed: {e}")

# ---------------- RUN ----------------

if __name__ == "__main__":
    # Basic sanity checks
    if not Path(IMG_SOURCE).exists():
        print(f"ERROR: Source image '{IMG_SOURCE}' not found. Place voters.jpg in script folder.")
    else:
        generate_excel_from_cards()
