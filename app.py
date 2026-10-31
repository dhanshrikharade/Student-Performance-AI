import io
import re
import shutil
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import fitz
except ImportError:
    fitz = None

try:
    import pytesseract
    from PIL import Image
except ImportError:
    pytesseract = None
    Image = None

try:
    from docx import Document
except ImportError:
    Document = None


st.set_page_config(
    page_title="Student Performance AI",
    page_icon="📊",
    layout="wide"
)

TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

if pytesseract and Path(TESSERACT_PATH).exists():
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


# ============================================================
# STYLING
# ============================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

.stApp {
    background:
      radial-gradient(circle at 10% 0%, rgba(91,141,239,.12), transparent 32%),
      radial-gradient(circle at 95% 5%, rgba(147,97,226,.12), transparent 30%),
      var(--background-color);
    font-family: Inter, sans-serif;
}

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
    max-width: 1450px;
}

div[data-testid="stMetric"] {
    background: var(--secondary-background-color);
    border: 1px solid rgba(120,120,140,.16);
    padding: 16px;
    border-radius: 15px;
}

div.stButton > button,
div.stDownloadButton > button {
    border-radius: 10px;
    font-weight: 600;
    min-height: 2.7rem;
}

.hero {
    padding: 2rem 2.2rem;
    border-radius: 22px;
    color: white;
    margin-bottom: 1.2rem;
    background: linear-gradient(120deg, #173b70 0%, #345fbd 54%, #7650b9 100%);
    box-shadow: 0 12px 35px rgba(37,75,145,.2);
}

.hero h1 {
    font-size: clamp(2rem, 4vw, 3rem);
    margin: .3rem 0 .5rem;
    color: white;
}

.hero p {
    font-size: 1.05rem;
    opacity: .92;
    max-width: 850px;
}

.hero-tag {
    font-size: .78rem;
    letter-spacing: .15em;
    font-weight: 800;
    opacity: .85;
}

.section-card {
    padding: 1rem 1.1rem;
    border: 1px solid rgba(120,120,140,.2);
    border-radius: 15px;
    background: var(--secondary-background-color);
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# CLEANING HELPERS
# ============================================================

def clean_cell(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = str(value).replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", text).strip()


def norm(value):
    return re.sub(r"[^a-z0-9]+", "", clean_cell(value).lower())


def dataframe_from_rows(rows):
    if not rows:
        return pd.DataFrame()

    cleaned = [
        [clean_cell(v) for v in list(row)]
        for row in rows
        if row is not None
    ]

    if not cleaned:
        return pd.DataFrame()

    width = max((len(row) for row in cleaned), default=0)

    if width == 0:
        return pd.DataFrame()

    padded = [
        row + [""] * (width - len(row))
        for row in cleaned
    ]

    return pd.DataFrame(padded)


def to_mark(value):
    """
    Convert marks to numbers.

    Blank remains missing.
    AB/Absent becomes zero.
    Text that is not a mark remains missing.
    """
    text = clean_cell(value).upper().replace(",", "")

    if not text:
        return None

    if text in {"AB", "ABSENT", "A", "NA", "N/A", "-", "--"}:
        return 0.0

    # Avoid treating ordinary text as a mark.
    match = re.fullmatch(r"\s*(-?\d+(?:\.\d+)?)\s*", text)

    if not match:
        return None

    try:
        number = float(match.group(1))
        if 0 <= number <= 1000:
            return number
    except ValueError:
        pass

    return None


def dataframe_from_text(text):
    """
    Read tab-separated or clearly spaced text.

    This is a fallback for PDF/TXT/OCR content.
    """
    rows = []

    for line in text.splitlines():
        line = line.strip()

        if not line:
            continue

        if "\t" in line:
            parts = [p.strip() for p in line.split("\t")]
        elif re.search(r"\s{2,}", line):
            parts = re.split(r"\s{2,}", line)
        else:
            continue

        if len(parts) >= 3:
            rows.append(parts)

    return dataframe_from_rows(rows)


# ============================================================
# OCR
# ============================================================

def extract_ocr_image(image):
    if not pytesseract or not Image:
        raise RuntimeError(
            "OCR packages are missing. Install them using: "
            "pip install pytesseract pillow"
        )

    if not shutil.which("tesseract") and not Path(TESSERACT_PATH).exists():
        raise RuntimeError(
            "Tesseract OCR was not found. Install Tesseract OCR "
            "or correct TESSERACT_PATH in app.py."
        )

    return pytesseract.image_to_string(image, config="--psm 6")


# ============================================================
# FILE EXTRACTION
# ============================================================

def extract_from_file(uploaded_file):
    """
    Returns:
        tables: list of (source label, raw DataFrame)
        extracted text: text available from the source
    """
    name = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()

    tables = []
    text_parts = []

    # ---------------- EXCEL ----------------
    if name.endswith((".xlsx", ".xls", ".xlsm")):
        try:
            workbook = pd.ExcelFile(io.BytesIO(raw))
        except Exception as exc:
            raise ValueError(
                f"Could not open the Excel workbook: {exc}. "
                "For .xlsx files install openpyxl; for old .xls files "
                "install xlrd."
            )

        for sheet_name in workbook.sheet_names:
            try:
                df = pd.read_excel(
                    workbook,
                    sheet_name=sheet_name,
                    header=None,
                    dtype=object
                )
            except Exception as exc:
                text_parts.append(
                    f"Could not read sheet '{sheet_name}': {exc}"
                )
                continue

            if df is None or df.empty:
                continue

            # Keep all non-empty records, including rows with blank cells.
            df = df.dropna(how="all").reset_index(drop=True)
            df = df.loc[:, ~df.apply(
                lambda col: col.map(clean_cell).eq("").all()
            )]

            if not df.empty:
                tables.append((f"Excel sheet: {sheet_name}", df))

        if not tables:
            raise ValueError(
                "No non-empty worksheets were found in the Excel file."
            )

        return tables, "\n".join(text_parts)

    # ---------------- CSV ----------------
    if name.endswith(".csv"):
        decoded = None

        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
            try:
                decoded = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                continue

        if decoded is None:
            raise ValueError("Could not decode this CSV file.")

        try:
            df = pd.read_csv(
                io.StringIO(decoded),
                header=None,
                sep=None,
                engine="python",
                dtype=object
            )
        except Exception:
            df = pd.read_csv(
                io.StringIO(decoded),
                header=None,
                dtype=object
            )

        if not df.empty:
            tables.append(("CSV data", df))

        return tables, decoded

    # ---------------- WORD ----------------
    if name.endswith(".docx"):
        if not Document:
            raise RuntimeError(
                "Install Word support: pip install python-docx"
            )

        doc = Document(io.BytesIO(raw))

        # Read every table in the Word document.
        for table_num, table in enumerate(doc.tables, start=1):
            rows = [
                [cell.text for cell in row.cells]
                for row in table.rows
            ]

            df = dataframe_from_rows(rows)

            if not df.empty:
                tables.append((f"Word table {table_num}", df))

        paragraphs = [
            p.text.strip()
            for p in doc.paragraphs
            if p.text.strip()
        ]

        text = "\n".join(paragraphs)

        if text:
            text_parts.append(text)
            paragraph_df = dataframe_from_text(text)

            if not paragraph_df.empty:
                tables.append(("Word paragraph text", paragraph_df))

        return tables, "\n".join(text_parts)

    # ---------------- PDF ----------------
    if name.endswith(".pdf"):
        page_texts = {}

        if pdfplumber:
            with pdfplumber.open(io.BytesIO(raw)) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    page_text = page.extract_text() or ""
                    page_texts[page_num] = page_text

                    if page_text.strip():
                        text_parts.append(
                            f"PDF page {page_num}\n{page_text}"
                        )

                    try:
                        page_tables = page.extract_tables() or []
                    except Exception:
                        page_tables = []

                    for table_num, rows in enumerate(page_tables, start=1):
                        df = dataframe_from_rows(rows)

                        if not df.empty:
                            tables.append((
                                f"PDF page {page_num}, table {table_num}",
                                df
                            ))

        # Try OCR on pages with no useful extracted text.
        if fitz and Image and pytesseract:
            pdf_doc = fitz.open(stream=raw, filetype="pdf")

            try:
                for page_num, page in enumerate(pdf_doc, start=1):
                    existing_text = page_texts.get(page_num, "")

                    if existing_text.strip():
                        continue

                    pix = page.get_pixmap(
                        matrix=fitz.Matrix(2, 2),
                        alpha=False
                    )

                    image = Image.open(
                        io.BytesIO(pix.tobytes("png"))
                    )

                    ocr_text = extract_ocr_image(image)

                    if ocr_text.strip():
                        text_parts.append(
                            f"PDF page {page_num} OCR\n{ocr_text}"
                        )

                        ocr_df = dataframe_from_text(ocr_text)

                        if not ocr_df.empty:
                            tables.append((
                                f"PDF page {page_num}, OCR",
                                ocr_df
                            ))
            finally:
                pdf_doc.close()

        all_text = "\n".join(text_parts)

        if not tables and all_text:
            df = dataframe_from_text(all_text)
            if not df.empty:
                tables.append(("PDF text fallback", df))

        if not tables and not all_text:
            raise ValueError(
                "No readable PDF text or tables were found. "
                "If this is a scanned PDF, install Tesseract OCR."
            )

        return tables, all_text

    # ---------------- IMAGE ----------------
    if name.endswith((
        ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"
    )):
        if not Image:
            raise RuntimeError("Install image support: pip install pillow")

        text = extract_ocr_image(Image.open(io.BytesIO(raw)))
        df = dataframe_from_text(text)

        if not df.empty:
            tables.append(("Image OCR", df))

        return tables, text

    # ---------------- TXT ----------------
    if name.endswith(".txt"):
        text = raw.decode("utf-8", errors="replace")
        df = dataframe_from_text(text)

        if not df.empty:
            tables.append(("Text file", df))

        return tables, text

    raise ValueError(
        "Unsupported file type. Use XLSX, XLS, CSV, DOCX, PDF, TXT, or an image."
    )


# ============================================================
# HEADER DETECTION
# ============================================================

ROLL_WORDS = (
    "rollno", "rollnumber", "roll", "seatno", "seatnumber",
    "enrollmentno", "enrolmentno", "studentid",
    "registrationno", "regno", "prn"
)

NAME_WORDS = (
    "name", "studentname", "nameofstudent", "candidatename"
)

IGNORE_WORDS = (
    "total", "percentage", "percent", "average", "rank",
    "result", "grade", "remark", "remarks", "attendance",
    "srno", "serialno", "maxmarks", "minmarks",
    "outof", "obtainedmarks"
)


def is_roll_header(value):
    value_norm = norm(value)
    return any(word in value_norm for word in ROLL_WORDS)


def is_name_header(value):
    value_norm = norm(value)
    return any(word in value_norm for word in NAME_WORDS)


def is_subject_header(value):
    value_norm = norm(value)

    if not value_norm:
        return False

    if any(word in value_norm for word in IGNORE_WORDS):
        return False

    if is_roll_header(value) or is_name_header(value):
        return False

    return (
        bool(re.search(r"[a-z]", value_norm))
        and len(value_norm) <= 40
    )


def detect_header(df):
    """
    Find the likely header in the first 100 rows.
    Roll number is optional; student name plus subject columns is sufficient.
    """
    if df is None or df.empty:
        return None

    sample = df.head(100).fillna("")
    best = None
    best_score = -1

    for row_idx in range(len(sample)):
        values = [
            clean_cell(v)
            for v in sample.iloc[row_idx].tolist()
        ]

        roll_cols = [
            i for i, value in enumerate(values)
            if is_roll_header(value)
        ]

        name_cols = [
            i for i, value in enumerate(values)
            if is_name_header(value)
        ]

        subject_cols = [
            (i, value or f"Subject {i + 1}")
            for i, value in enumerate(values)
            if is_subject_header(value)
        ]

        score = (
            (4 if roll_cols else 0)
            + (4 if name_cols else 0)
            + min(len(subject_cols), 8)
        )

        if name_cols and subject_cols:
            score += 5

        if roll_cols and name_cols and subject_cols:
            score += 2

        if score > best_score:
            best_score = score
            best = (
                row_idx,
                roll_cols,
                name_cols,
                subject_cols
            )

    if best and best_score >= 5 and best[3] and (best[1] or best[2]):
        row_idx, roll_cols, name_cols, subject_cols = best

        return (
            row_idx,
            roll_cols[0] if roll_cols else None,
            name_cols[0] if name_cols else None,
            subject_cols
        )

    return None


def is_summary_or_header_row(name, roll=""):
    name_norm = norm(name)
    roll_norm = norm(roll)

    if not name_norm and not roll_norm:
        return True

    blocked = (
        "total", "grandtotal", "average", "avg", "percentage",
        "percent", "result", "remarks", "remark", "summary",
        "maximum", "minimum", "maxmarks", "minmarks",
        "passcount", "failcount", "classaverage", "subjectwise"
    )

    if any(word in name_norm for word in blocked):
        return True

    if name_norm in {
        "name", "studentname", "rollno", "seatno",
        "srno", "serialno"
    }:
        return True

    return False


# ============================================================
# NORMALIZE TABLES
# ============================================================

def normalize_student_table(raw_df):
    if raw_df is None or raw_df.empty:
        return None, None

    df = raw_df.copy().fillna("")
    df = df.apply(lambda col: col.map(clean_cell))

    header = detect_header(df)

    if not header:
        return None, None

    header_idx, roll_col, name_col, subject_cols = header

    if not subject_cols:
        return None, None

    data = df.iloc[header_idx + 1:].copy()

    result = pd.DataFrame(index=data.index)

    if roll_col is not None:
        result["Roll No"] = data.iloc[:, roll_col].map(clean_cell)
    else:
        result["Roll No"] = ""

    if name_col is not None:
        result["Name"] = data.iloc[:, name_col].map(clean_cell)
    else:
        result["Name"] = ""

    for col_idx, subject in subject_cols:
        subject_name = clean_cell(subject) or f"Subject {col_idx + 1}"

        base = subject_name
        suffix = 2
        label = subject_name

        while label in result.columns:
            label = f"{base} ({suffix})"
            suffix += 1

        result[label] = data.iloc[:, col_idx].map(to_mark)

    subject_names = list(result.columns[2:])

    has_marks = result[subject_names].notna().any(axis=1)

    valid_identity = []

    for _, row in result.iterrows():
        name_value = clean_cell(row["Name"])
        roll_value = clean_cell(row["Roll No"])

        name_ok = (
            bool(name_value)
            and not name_value.replace(".", "", 1).isdigit()
        )

        roll_ok = bool(roll_value)

        valid_identity.append(
            (name_ok or roll_ok)
            and not is_summary_or_header_row(name_value, roll_value)
        )

    result = result[
        has_marks
        & pd.Series(valid_identity, index=result.index)
    ].reset_index(drop=True)

    return (result if not result.empty else None), header


def normalize_continuation(raw_df, template_columns):
    """
    Handle tables/pages that omit their repeated header.
    Assumes column order matches the recognized table.
    """
    if raw_df is None or raw_df.empty or not template_columns:
        return None

    df = raw_df.copy().fillna("")
    output = []
    expected = len(template_columns)

    for _, row in df.iterrows():
        values = [clean_cell(v) for v in row.tolist()]

        if not any(values):
            continue

        if any(is_roll_header(v) for v in values):
            continue

        if any(is_name_header(v) for v in values):
            continue

        values = (values + [""] * expected)[:expected]

        roll = values[0] if expected > 0 else ""
        name = values[1] if expected > 1 else ""

        marks = [
            to_mark(v)
            for v in values[2:]
        ]

        if not any(mark is not None for mark in marks):
            continue

        if is_summary_or_header_row(name, roll):
            continue

        if not name and not roll:
            continue

        record = {
            "Roll No": roll,
            "Name": name
        }

        for col, mark in zip(template_columns[2:], marks):
            record[col] = mark

        output.append(record)

    if not output:
        return None

    return pd.DataFrame(output, columns=template_columns)


# ============================================================
# MERGE EXTRACTED TABLES
# ============================================================

def merge_student_tables(raw_tables):
    recognized = []
    messages = []
    template_columns = None

    # Pass 1: tables with their own headers.
    for source, raw_df in raw_tables:
        standardized, _ = normalize_student_table(raw_df)

        if standardized is not None and not standardized.empty:
            recognized.append((source, standardized))

            if (
                template_columns is None
                or len(standardized.columns) > len(template_columns)
            ):
                template_columns = list(standardized.columns)

    # Pass 2: headerless continuation tables.
    if template_columns:
        recognized_sources = {
            source for source, _ in recognized
        }

        for source, raw_df in raw_tables:
            if source in recognized_sources:
                continue

            continuation = normalize_continuation(
                raw_df,
                template_columns
            )

            if continuation is not None and not continuation.empty:
                recognized.append((source, continuation))
                messages.append(
                    f"{source}: treated as a continuation table. "
                    "Please verify its column alignment."
                )

    # Fallback for tables where headers were not detected.
    if not recognized:
        fallback = []

        for source, raw_df in raw_tables:
            for _, row in raw_df.fillna("").iterrows():
                cells = [clean_cell(v) for v in row.tolist()]

                if len(cells) < 3:
                    continue

                mark_positions = [
                    i for i, value in enumerate(cells)
                    if to_mark(value) is not None
                ]

                if not mark_positions or mark_positions[0] < 1:
                    continue

                first_mark = mark_positions[0]

                identity = " ".join(
                    cells[:first_mark]
                ).strip()

                marks = [
                    to_mark(value)
                    for value in cells[first_mark:]
                ]

                if not identity or not any(m is not None for m in marks):
                    continue

                match = re.match(
                    r"^([A-Za-z0-9/-]+)\s+(.+)$",
                    identity
                )

                if match:
                    roll, student_name = match.group(1), match.group(2)
                else:
                    roll, student_name = "", identity

                record = {
                    "Roll No": roll,
                    "Name": student_name
                }

                for i, mark in enumerate(marks):
                    record[f"Subject {i + 1}"] = mark

                fallback.append(record)

        if fallback:
            fallback_df = pd.DataFrame(fallback).fillna(pd.NA)

            messages.append(
                "Fallback text detection was used. "
                "Check names, marks, and column alignment manually."
            )

            return fallback_df, messages

        return None, messages

    # Combine all subject columns without dropping extra subjects.
    all_columns = ["Roll No", "Name"]

    for _, frame in recognized:
        for col in frame.columns:
            if col not in all_columns:
                all_columns.append(col)

    aligned = []

    for source, frame in recognized:
        frame = frame.copy()

        for col in all_columns:
            if col not in frame.columns:
                frame[col] = None

        frame = frame[all_columns]
        aligned.append(frame)

        messages.append(
            f"{source}: {len(frame)} student row(s) extracted."
        )

    combined = pd.concat(aligned, ignore_index=True)

    combined["Roll No"] = combined["Roll No"].map(clean_cell)
    combined["Name"] = combined["Name"].map(clean_cell)

    # Only remove exact duplicate copies based on identity AND all marks.
    # Same-name students with different marks are retained.
    # Different roll numbers are retained.
    before = len(combined)

    def duplicate_key(row):
        roll = norm(row.get("Roll No", ""))
        student_name = norm(row.get("Name", ""))

        marks = tuple(
            "" if pd.isna(row.get(col)) else str(row.get(col))
            for col in all_columns[2:]
        )

        identity = (
            ("roll", roll)
            if roll
            else ("name", student_name)
        )

        return identity + (marks,)

    keys = combined.apply(duplicate_key, axis=1)
    combined = combined.loc[
        ~keys.duplicated()
    ].reset_index(drop=True)

    removed = before - len(combined)

    if removed:
        messages.append(
            f"Removed {removed} exact duplicate record copy/copies."
        )

    return combined, messages


# ============================================================
# APP STATE
# ============================================================

with st.sidebar:
    st.markdown("## 🎓 Student AI")
    st.caption("Academic performance workspace")

    page = st.radio(
        "NAVIGATION",
        [
            "Home",
            "Upload & Extract",
            "Review Records",
            "Analytics & Ranking",
            "Export Results"
        ]
    )

    st.divider()
    st.markdown("### ⚙️ Analysis settings")

    max_marks = st.number_input(
        "Maximum marks per subject",
        min_value=1,
        max_value=1000,
        value=25,
        step=1,
        help="Set the maximum possible marks for each subject."
    )

    st.caption("AB / Absent is treated as 0. Blank marks remain missing.")
    st.divider()
    st.caption("Always verify extracted student records before using results.")


st.markdown("""
<div class="hero">
  <div class="hero-tag">STUDENT ANALYTICS PLATFORM</div>
  <h1>Student Performance AI</h1>
  <p>Upload academic records, review extracted student data, analyse marks,
  compare performance, and export class rankings.</p>
</div>
""", unsafe_allow_html=True)


if "raw_tables" not in st.session_state:
    st.session_state.raw_tables = []

if "extracted_text" not in st.session_state:
    st.session_state.extracted_text = ""

if "marks_df" not in st.session_state:
    st.session_state.marks_df = None

if "extracted" not in st.session_state:
    st.session_state.extracted = False

if "extract_messages" not in st.session_state:
    st.session_state.extract_messages = []

if "source_name" not in st.session_state:
    st.session_state.source_name = ""

if "extract_error" not in st.session_state:
    st.session_state.extract_error = ""


# ============================================================
# TOTALS, PERCENTAGE AND RANKING
# ============================================================

def get_analysed_df():
    base = st.session_state.get("marks_df")

    if not isinstance(base, pd.DataFrame) or base.empty:
        return None

    df = base.copy()

    if "Roll No" not in df.columns:
        df.insert(0, "Roll No", [""] * len(df))

    if "Name" not in df.columns:
        df.insert(1, "Name", [""] * len(df))

    df["Roll No"] = df["Roll No"].map(clean_cell)
    df["Name"] = df["Name"].map(clean_cell)

    metadata_columns = {
        "Roll No",
        "Name",
        "Rank",
        "Total",
        "Subjects Counted",
        "Percentage"
    }

    subject_cols = [
        col for col in df.columns
        if col not in metadata_columns
    ]

    if not subject_cols:
        return None

    # Convert marks consistently.
    for col in subject_cols:
        df[col] = df[col].map(to_mark)

    # Keep students who have a name OR a roll number.
    df = df[
        df["Name"].str.strip().ne("")
        | df["Roll No"].str.strip().ne("")
    ].copy()

    if df.empty:
        return None

    # Total from available numeric marks only.
    df["Total"] = df[subject_cols].sum(
        axis=1,
        min_count=1
    )

    df["Subjects Counted"] = df[subject_cols].notna().sum(axis=1)

    possible_marks = df["Subjects Counted"] * float(max_marks)

    df["Percentage"] = (
        df["Total"] / possible_marks * 100
    ).where(
        df["Subjects Counted"] > 0
    ).round(2)

    # Competition ranking: 1, 1, 3.
    # Rank is based on total marks only.
    df["Rank"] = (
        df["Total"]
        .rank(method="min", ascending=False)
        .astype("Int64")
    )

    # Sort for display; sorting doesn't change rank.
    df = df.sort_values(
        ["Rank", "Name", "Roll No"],
        ascending=[True, True, True],
        na_position="last",
        kind="stable"
    ).reset_index(drop=True)

    # Rank first in the displayed table.
    cols = list(df.columns)

    if "Rank" in cols:
        cols.remove("Rank")
        cols.insert(0, "Rank")
        df = df[cols]

    return df


# ============================================================
# UPLOAD PAGE
# ============================================================

def display_upload():
    st.subheader("📤 Upload academic records")

    st.write(
        "Upload the complete marksheet. The system scans supported tables "
        "and student records from the file."
    )

    uploaded_file = st.file_uploader(
        "Choose a marksheet file",
        type=[
            "pdf", "docx", "xlsx", "xls", "xlsm",
            "csv", "txt", "png", "jpg", "jpeg",
            "tif", "tiff", "bmp", "webp"
        ],
        key="main_uploader"
    )

    if uploaded_file is not None:
        st.info(f"Selected file: **{uploaded_file.name}**")

        if st.button(
            "🔍 Extract all tables and students",
            type="primary",
            use_container_width=True
        ):
            with st.spinner("Scanning file and extracting student records..."):
                try:
                    raw_tables, extracted_text = extract_from_file(uploaded_file)

                    combined_df, messages = merge_student_tables(raw_tables)

                    st.session_state.raw_tables = raw_tables
                    st.session_state.extracted_text = extracted_text
                    st.session_state.extract_messages = messages
                    st.session_state.source_name = uploaded_file.name
                    st.session_state.marks_df = combined_df
                    st.session_state.extracted = (
                        combined_df is not None
                        and not combined_df.empty
                    )

                    st.session_state.extract_error = (
                        ""
                        if st.session_state.extracted
                        else (
                            "No student marks table was detected. "
                            "Check the file layout and verify that the marks "
                            "and student names are readable."
                        )
                    )

                except Exception as exc:
                    st.session_state.extracted = False
                    st.session_state.extract_error = str(exc)
                    st.session_state.marks_df = None

            st.rerun()

    if st.session_state.extract_error:
        st.error(st.session_state.extract_error)

    if (
        st.session_state.extracted
        and isinstance(st.session_state.marks_df, pd.DataFrame)
    ):
        st.success(
            f"Extracted {len(st.session_state.marks_df)} student row(s) "
            f"from {len(st.session_state.raw_tables)} table/text block(s)."
        )

        for message in st.session_state.extract_messages:
            st.caption("• " + message)

        st.info(
            "Next: open **Review Records** and compare the extracted student "
            "count with your original marksheet."
        )
    else:
        st.caption(
            "Tip: Check the extracted count against the original file. "
            "Scanned documents may require OCR software."
        )


# ============================================================
# HOME
# ============================================================

if page == "Home":
    st.subheader("Welcome to your academic analysis workspace")

    st.write(
        "Use the left-side navigation to move between pages. "
        "Your extracted records remain available while you switch pages."
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Current source",
        "Loaded" if st.session_state.source_name else "Not loaded"
    )

    c2.metric(
        "Student records",
        len(st.session_state.marks_df)
        if isinstance(st.session_state.marks_df, pd.DataFrame)
        else 0
    )

    c3.metric(
        "Analysis status",
        "Ready" if st.session_state.extracted else "Waiting for file"
    )

    st.markdown("### How to use the system")

    steps = [
        ("01", "Upload & Extract", "Upload a marksheet and extract student records."),
        ("02", "Review Records", "Inspect and correct student names and marks."),
        ("03", "Analytics & Ranking", "View class metrics, top students, and charts."),
        ("04", "Export Results", "Download the processed ranking table as CSV.")
    ]

    for number, title, description in steps:
        st.markdown(f"""
        <div class="section-card" style="margin:.5rem 0;">
          <b>{number} · {title}</b><br>
          <span style="opacity:.8">{description}</span>
        </div>
        """, unsafe_allow_html=True)

    st.warning(
        "This application calculates and analyses marks. "
        "It does not claim to use a trained machine-learning model."
    )


# ============================================================
# UPLOAD & EXTRACT
# ============================================================

elif page == "Upload & Extract":
    display_upload()


# ============================================================
# REVIEW RECORDS
# ============================================================

elif page == "Review Records":
    st.subheader("🧾 Review extracted student records")

    if (
        not st.session_state.extracted
        or not isinstance(st.session_state.marks_df, pd.DataFrame)
    ):
        st.warning("No student data is loaded yet. Open Upload & Extract first.")

    else:
        st.caption(f"Source file: {st.session_state.source_name}")

        st.success(
            f"{len(st.session_state.marks_df)} student row(s) currently loaded. "
            "Compare this number with the original marksheet."
        )

        for message in st.session_state.extract_messages:
            st.caption("• " + message)

        edited_df = st.data_editor(
            st.session_state.marks_df,
            use_container_width=True,
            num_rows="dynamic",
            key="student_records_editor",
            column_config={
                "Roll No": st.column_config.TextColumn("Roll No"),
                "Name": st.column_config.TextColumn("Name")
            }
        )

        if st.button("💾 Save reviewed records", type="primary"):
            clean_df = edited_df.copy()

            if "Roll No" not in clean_df.columns:
                clean_df.insert(0, "Roll No", [""] * len(clean_df))

            if "Name" not in clean_df.columns:
                clean_df.insert(1, "Name", [""] * len(clean_df))

            clean_df["Roll No"] = clean_df["Roll No"].map(clean_cell)
            clean_df["Name"] = clean_df["Name"].map(clean_cell)

            for col in clean_df.columns:
                if col not in ("Roll No", "Name"):
                    clean_df[col] = clean_df[col].map(to_mark)

            clean_df = clean_df[
                clean_df["Roll No"].map(clean_cell).ne("")
                | clean_df["Name"].map(clean_cell).ne("")
            ].reset_index(drop=True)

            st.session_state.marks_df = clean_df
            st.success(
                f"Saved {len(clean_df)} reviewed student record(s)."
            )
            st.rerun()


# ============================================================
# ANALYTICS & RANKING
# ============================================================

elif page == "Analytics & Ranking":
    st.subheader("📈 Class analytics and student ranking")

    df = (
        get_analysed_df()
        if st.session_state.extracted
        else None
    )

    if df is None or df.empty:
        st.warning("Upload a file and review the records first.")

    else:
        metadata_columns = {
            "Rank", "Roll No", "Name", "Total",
            "Subjects Counted", "Percentage"
        }

        subject_cols = [
            col for col in df.columns
            if col not in metadata_columns
        ]

        a, b, c, d = st.columns(4)

        a.metric("Students", len(df))

        average_percentage = df["Percentage"].mean()
        b.metric(
            "Class average",
            f"{average_percentage:.2f}%"
            if pd.notna(average_percentage)
            else "N/A"
        )

        highest_percentage = df["Percentage"].max()
        c.metric(
            "Highest percentage",
            f"{highest_percentage:.2f}%"
            if pd.notna(highest_percentage)
            else "N/A"
        )

        d.metric("Subjects", len(subject_cols))

        st.markdown("### 🏆 Top 3 students")

        display_cols = (
            ["Rank", "Roll No", "Name"]
            + subject_cols
            + ["Total", "Percentage"]
        )

        st.dataframe(
            df.head(3)[display_cols],
            use_container_width=True,
            hide_index=True
        )

        st.markdown("### Full class ranking")

        search = st.text_input(
            "Search by student name or roll number",
            key="analytics_search"
        )

        filtered = df.copy()

        if search.strip():
            filtered = filtered[
                filtered["Name"].str.contains(
                    search,
                    case=False,
                    na=False,
                    regex=False
                )
                | filtered["Roll No"].str.contains(
                    search,
                    case=False,
                    na=False,
                    regex=False
                )
            ]

        st.dataframe(
            filtered[display_cols],
            use_container_width=True,
            hide_index=True
        )

        st.markdown("### Subject-wise performance")

        if subject_cols:
            averages = df[subject_cols].mean(
                numeric_only=True
            ).round(2)

            averages_df = averages.rename(
                "Average Marks"
            ).reset_index()

            averages_df.columns = [
                "Subject",
                "Average Marks"
            ]

            left, right = st.columns(2)

            with left:
                fig_avg = px.bar(
                    averages_df,
                    x="Subject",
                    y="Average Marks",
                    title="Average marks by subject",
                    text_auto=".2f"
                )

                st.plotly_chart(
                    fig_avg,
                    use_container_width=True
                )

            with right:
                fig_pct = px.bar(
                    df,
                    x="Name",
                    y="Percentage",
                    hover_data=["Roll No", "Total"],
                    title="Student percentage"
                )

                fig_pct.update_layout(
                    xaxis_title="Student",
                    yaxis_title="Percentage (%)",
                    xaxis_tickangle=-45
                )

                st.plotly_chart(
                    fig_pct,
                    use_container_width=True
                )

        st.caption(
            "Rank is calculated from total available marks. "
            "Verify extracted records before using the results."
        )


# ============================================================
# EXPORT RESULTS
# ============================================================

elif page == "Export Results":
    st.subheader("⬇️ Export student performance results")

    df = (
        get_analysed_df()
        if st.session_state.extracted
        else None
    )

    if df is None or df.empty:
        st.warning(
            "No processed records are available. "
            "Upload and review student records first."
        )

    else:
        metadata_columns = {
            "Rank", "Roll No", "Name", "Total",
            "Subjects Counted", "Percentage"
        }

        subject_cols = [
            col for col in df.columns
            if col not in metadata_columns
        ]

        export_cols = (
            ["Rank", "Roll No", "Name"]
            + subject_cols
            + ["Total", "Subjects Counted", "Percentage"]
        )

        export_df = df[export_cols].copy()

        st.success(
            f"Ready to export {len(export_df)} student record(s)."
        )

        st.dataframe(
            export_df,
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            "⬇️ Download ranking as CSV",
            data=export_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="student_performance_ranking.csv",
            mime="text/csv",
            use_container_width=True
        )