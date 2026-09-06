import os
import re
import shutil
import difflib
import tkinter as tk

from tkinter import ttk, filedialog, messagebox

from datetime import datetime

import pandas as pd
import openpyxl

# PDF libraries
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


# ============================================================
# CONFIGURATION
# ============================================================

MASTER_FILE = "Job_Application_Tracker.xlsx"
BACKUP_FOLDER = "Backups"

os.makedirs(BACKUP_FOLDER, exist_ok=True)


# ============================================================
# GLOBAL DATA
# ============================================================

data = []


# ============================================================
# NORMALIZE TEXT
# ============================================================

def normalize_text(value):

    if value is None:
        return ""

    if pd.isna(value):
        return ""

    text = str(value)

    text = text.replace("\n", " ")
    text = text.replace("\r", " ")
    text = text.replace("\t", " ")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def normalize_column_name(value):

    text = normalize_text(value).lower()

    # Common Excel variations
    replacements = {
        "&": " and ",
        "/": " ",
        "\\": " ",
        "-": " ",
        "_": " ",
        ".": " ",
        ":": " ",
        "(": " ",
        ")": " "
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ============================================================
# COLUMN ALIASES
# ============================================================

COLUMN_ALIASES = {

    "Company": [
        "company",
        "company name",
        "comapany name",
        "comapany",
        "employer",
        "employer name",
        "organization",
        "organisation",
        "organization name",
        "company/organization",
        "firm",
        "business name",
        "recruiter company"
    ],

    "Job Title": [
        "job",
        "job title",
        "job role",
        "role",
        "position",
        "position applied",
        "position applied for",
        "designation",
        "designation applied",
        "job designation",
        "job profile",
        "post",
        "post name",
        "vacancy",
        "opening",
        "job opening"
    ],

    "Location": [
        "location",
        "job location",
        "work location",
        "city",
        "job city",
        "place",
        "place of work",
        "office location",
        "preferred location",
        "workplace"
    ],

    "Application Date": [
        "date",
        "application date",
        "date applied",
        "applied date",
        "applied on",
        "application submitted",
        "submitted date",
        "submission date",
        "applied",
        "application day"
    ],

    "Source": [
        "source",
        "portal",
        "job portal",
        "website",
        "platform",
        "application source",
        "applied through",
        "found through",
        "recruitment platform"
    ],

    "Status": [
        "status",
        "application status",
        "current status",
        "result",
        "application result",
        "select reject",
        "select/reject",
        "selected",
        "selection status",
        "result status",
        "stage",
        "application stage"
    ],

    "Job URL": [
        "url",
        "job url",
        "job link",
        "link",
        "application link",
        "website link",
        "job website",
        "career link"
    ],

    "Notes": [
        "notes",
        "note",
        "remark",
        "remarks",
        "comment",
        "comments",
        "description",
        "details",
        "feedback"
    ]
}


# ============================================================
# FUZZY COLUMN MATCHING
# ============================================================

def similarity(a, b):

    a = normalize_column_name(a)
    b = normalize_column_name(b)

    if not a or not b:
        return 0

    if a == b:
        return 1.0

    if a in b or b in a:
        return 0.90

    return difflib.SequenceMatcher(None, a, b).ratio()


def detect_column(column_name):

    column_name = normalize_column_name(column_name)

    best_field = None
    best_score = 0

    for field, aliases in COLUMN_ALIASES.items():

        for alias in aliases:

            score = similarity(column_name, alias)

            if score > best_score:
                best_score = score
                best_field = field

    # Don't accept extremely weak matches
    if best_score >= 0.65:
        return best_field, best_score

    return None, 0


# ============================================================
# DETECT HEADER ROW
# ============================================================

def detect_header_row(raw_df):

    best_row = None
    best_score = 0

    for row_number in range(min(len(raw_df), 30)):

        row = raw_df.iloc[row_number]

        values = [
            normalize_column_name(x)
            for x in row.tolist()
            if normalize_text(x)
        ]

        if not values:
            continue

        matched_fields = set()

        for value in values:

            field, score = detect_column(value)

            if field and score >= 0.70:
                matched_fields.add(field)

        score = len(matched_fields)

        # Company + Job Title is especially important
        if "Company" in matched_fields:
            score += 2

        if "Job Title" in matched_fields:
            score += 2

        if score > best_score:

            best_score = score
            best_row = row_number

    return best_row


# ============================================================
# MAP COLUMNS
# ============================================================

def map_columns(df):

    mapping = {}

    used = set()

    for column in df.columns:

        field, score = detect_column(column)

        if field and field not in used:

            mapping[field] = column
            used.add(field)

    return mapping


# ============================================================
# CLEAN DATAFRAME
# ============================================================

def clean_dataframe(df):

    # Remove completely empty rows
    df = df.dropna(how="all")

    # Remove completely empty columns
    df = df.dropna(axis=1, how="all")

    # Convert column names
    df.columns = [
        normalize_text(c)
        for c in df.columns
    ]

    return df


# ============================================================
# CONVERT VALUE TO DATE
# ============================================================

def clean_date(value):

    if value is None:
        return ""

    if pd.isna(value):
        return ""

    # Timestamp
    if isinstance(value, pd.Timestamp):

        return value.strftime("%Y-%m-%d")

    text = normalize_text(value)

    if not text:
        return ""

    # Try pandas
    try:

        date_value = pd.to_datetime(
            text,
            errors="coerce",
            dayfirst=True
        )

        if not pd.isna(date_value):

            return date_value.strftime("%Y-%m-%d")

    except Exception:
        pass

    # Search date inside text
    patterns = [

        r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",

        r"\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b",

        r"\b\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}\b",

        r"\b[A-Za-z]{3,9}\s+\d{1,2},\s+\d{4}\b"
    ]

    for pattern in patterns:

        match = re.search(pattern, text)

        if match:

            try:

                date_value = pd.to_datetime(
                    match.group(),
                    errors="coerce",
                    dayfirst=True
                )

                if not pd.isna(date_value):

                    return date_value.strftime("%Y-%m-%d")

            except Exception:
                pass

    return text


# ============================================================
# NORMALIZE STATUS
# ============================================================

def normalize_status(value):

    text = normalize_text(value)

    if not text:
        return "Applied"

    lower = text.lower()

    if any(x in lower for x in [
        "reject",
        "rejected",
        "not selected",
        "unsuccessful",
        "declined"
    ]):
        return "Rejected"

    if any(x in lower for x in [
        "select",
        "selected",
        "shortlisted",
        "short list"
    ]):
        return "Selected"

    if any(x in lower for x in [
        "interview",
        "assessment",
        "test",
        "technical round",
        "hr round"
    ]):
        return text

    if any(x in lower for x in [
        "offer",
        "offered"
    ]):
        return "Offer"

    if any(x in lower for x in [
        "withdraw",
        "withdrawn"
    ]):
        return "Withdrawn"

    return text


# ============================================================
# DUPLICATE DETECTION
# ============================================================

def normalize_company_for_duplicate(company):

    text = normalize_text(company).lower()

    # Remove common company suffixes
    suffixes = [
        "private limited",
        "pvt ltd",
        "pvt. ltd.",
        "limited",
        "ltd",
        "llp",
        "inc",
        "inc.",
        "corporation",
        "corp"
    ]

    for suffix in suffixes:

        text = text.replace(suffix, "")

    text = re.sub(r"[^a-z0-9]", "", text)

    return text


def normalize_role_for_duplicate(role):

    text = normalize_text(role).lower()

    text = re.sub(r"[^a-z0-9]", "", text)

    return text


def is_duplicate(company, role):

    company_key = normalize_company_for_duplicate(company)
    role_key = normalize_role_for_duplicate(role)

    for record in data:

        old_company = normalize_company_for_duplicate(
            record.get("Company", "")
        )

        old_role = normalize_role_for_duplicate(
            record.get("Job Title", "")
        )

        if (
            company_key == old_company
            and role_key == old_role
        ):
            return True

    return False


# ============================================================
# SAVE DATA
# ============================================================

def save_data():

    try:

        df = pd.DataFrame(data)

        columns = [
            "Company",
            "Job Title",
            "Location",
            "Application Date",
            "Source",
            "Status",
            "Job URL",
            "Notes"
        ]

        for column in columns:

            if column not in df.columns:
                df[column] = ""

        df = df[columns]

        df.to_excel(
            MASTER_FILE,
            index=False,
            engine="openpyxl"
        )

    except PermissionError:

        messagebox.showerror(
            "Excel File Open",
            "Please close Job_Application_Tracker.xlsx "
            "before saving."
        )

    except Exception as e:

        messagebox.showerror(
            "Save Error",
            str(e)
        )


# ============================================================
# LOAD EXISTING DATA
# ============================================================

def load_existing_data():

    global data

    data = []

    if not os.path.exists(MASTER_FILE):
        return

    try:

        df = pd.read_excel(
            MASTER_FILE,
            engine="openpyxl"
        )

        df = clean_dataframe(df)

        for _, row in df.iterrows():

            record = {

                "Company": normalize_text(
                    row.get("Company", "")
                ),

                "Job Title": normalize_text(
                    row.get("Job Title", "")
                ),

                "Location": normalize_text(
                    row.get("Location", "")
                ),

                "Application Date": normalize_text(
                    row.get("Application Date", "")
                ),

                "Source": normalize_text(
                    row.get("Source", "")
                ),

                "Status": normalize_text(
                    row.get("Status", "")
                ),

                "Job URL": normalize_text(
                    row.get("Job URL", "")
                ),

                "Notes": normalize_text(
                    row.get("Notes", "")
                )
            }

            if record["Company"]:

                data.append(record)

    except Exception as e:

        print("Could not load existing database:", e)


# ============================================================
# BACKUP
# ============================================================

def create_backup():

    if not os.path.exists(MASTER_FILE):
        return

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    backup_file = os.path.join(
        BACKUP_FOLDER,
        f"Job_Application_Tracker_{timestamp}.xlsx"
    )

    shutil.copy2(
        MASTER_FILE,
        backup_file
    )


# ============================================================
# IMPORT ONE DATAFRAME
# ============================================================

def import_dataframe(df, source_name):

    global data

    df = clean_dataframe(df)

    if df.empty:
        return 0, 0

    mapping = map_columns(df)

    print("\n--------------------------------")
    print("SOURCE:", source_name)
    print("COLUMN MAPPING:")
    print(mapping)
    print("--------------------------------")

    # Company is absolutely necessary
    if "Company" not in mapping:

        return 0, 0

    imported = 0
    duplicates = 0

    company_column = mapping["Company"]

    role_column = mapping.get("Job Title")

    location_column = mapping.get("Location")

    date_column = mapping.get("Application Date")

    source_column = mapping.get("Source")

    status_column = mapping.get("Status")

    url_column = mapping.get("Job URL")

    notes_column = mapping.get("Notes")

    for _, row in df.iterrows():

        company = normalize_text(
            row.get(company_column, "")
        )

        # Ignore invalid rows
        if not company:
            continue

        # Ignore rows that are actually header text
        if normalize_column_name(company) in [
            "company",
            "company name",
            "comapany name",
            "employer",
            "organization"
        ]:
            continue

        role = ""

        if role_column:

            role = normalize_text(
                row.get(role_column, "")
            )

        # If no role column exists, try to continue
        if not role:

            role = "Not specified"

        location = ""

        if location_column:

            location = normalize_text(
                row.get(location_column, "")
            )

        application_date = ""

        if date_column:

            application_date = clean_date(
                row.get(date_column, "")
            )

        source = source_name

        if source_column:

            extracted_source = normalize_text(
                row.get(source_column, "")
            )

            if extracted_source:
                source = extracted_source

        status = "Applied"

        if status_column:

            status = normalize_status(
                row.get(status_column, "")
            )

        url = ""

        if url_column:

            url = normalize_text(
                row.get(url_column, "")
            )

        notes = ""

        if notes_column:

            notes = normalize_text(
                row.get(notes_column, "")
            )

        # Duplicate
        if is_duplicate(company, role):

            duplicates += 1
            continue

        record = {

            "Company": company,

            "Job Title": role,

            "Location": location,

            "Application Date": application_date,

            "Source": source,

            "Status": status,

            "Job URL": url,

            "Notes": notes
        }

        data.append(record)

        imported += 1

    return imported, duplicates


# ============================================================
# EXCEL IMPORT
# ============================================================

def import_excel_file(file_path):

    total_imported = 0
    total_duplicates = 0
    sheets_processed = 0

    try:

        # ----------------------------------------------------
        # CSV
        # ----------------------------------------------------

        if file_path.lower().endswith(".csv"):

            df = pd.read_csv(
                file_path,
                header=None
            )

            header_row = detect_header_row(df)

            if header_row is not None:

                df = pd.read_csv(
                    file_path,
                    header=header_row
                )

            imported, duplicates = import_dataframe(
                df,
                os.path.basename(file_path)
            )

            return imported, duplicates, 1

        # ----------------------------------------------------
        # EXCEL
        # ----------------------------------------------------

        excel_file = pd.ExcelFile(file_path)

        for sheet_name in excel_file.sheet_names:

            try:

                # First read without header
                raw_df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    header=None
                )

                if raw_df.empty:
                    continue

                header_row = detect_header_row(
                    raw_df
                )

                if header_row is None:

                    # Try normal Excel structure
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet_name
                    )

                else:

                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet_name,
                        header=header_row
                    )

                imported, duplicates = import_dataframe(
                    df,
                    f"{os.path.basename(file_path)} - {sheet_name}"
                )

                total_imported += imported
                total_duplicates += duplicates

                sheets_processed += 1

            except Exception as e:

                print(
                    f"Sheet '{sheet_name}' error:",
                    e
                )

        return (
            total_imported,
            total_duplicates,
            sheets_processed
        )

    except Exception as e:

        raise Exception(
            f"Excel import failed:\n\n{e}"
        )


# ============================================================
# IMPORT EXCEL BUTTON
# ============================================================

def import_excel():

    file_path = filedialog.askopenfilename(

        title="Select Excel / CSV File",

        filetypes=[

            (
                "Excel Files",
                "*.xlsx *.xls *.xlsm *.xlsb"
            ),

            (
                "CSV Files",
                "*.csv"
            ),

            (
                "All Files",
                "*.*"
            )
        ]
    )

    if not file_path:
        return

    try:

        create_backup()

        imported, duplicates, sheets = \
            import_excel_file(file_path)

        save_data()

        refresh_table()

        messagebox.showinfo(

            "Excel Import Complete",

            f"File:\n{os.path.basename(file_path)}\n\n"

            f"Sheets processed: {sheets}\n"

            f"New applications: {imported}\n"

            f"Duplicates skipped: {duplicates}\n\n"

            f"Total applications: {len(data)}"
        )

    except Exception as e:

        messagebox.showerror(
            "Excel Import Error",
            str(e)
        )


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

def extract_pdf_text(file_path):

    if fitz is None:

        raise Exception(
            "PyMuPDF is not installed.\n\n"
            "Run:\n"
            "pip install pymupdf"
        )

    text = ""

    document = fitz.open(file_path)

    for page in document:

        text += page.get_text(
            "text"
        )

        text += "\n"

    document.close()

    return text


# ============================================================
# PDF TABLE EXTRACTION
# ============================================================

def extract_pdf_tables(file_path):

    tables = []

    if pdfplumber is None:
        return tables

    try:

        with pdfplumber.open(file_path) as pdf:

            for page in pdf.pages:

                page_tables = page.extract_tables()

                if page_tables:

                    for table in page_tables:

                        if table:

                            tables.append(table)

    except Exception as e:

        print(
            "PDF table extraction error:",
            e
        )

    return tables


# ============================================================
# PDF TEXT -> RECORDS
# ============================================================

def extract_company_from_text(text):

    patterns = [

        r"company\s*(?:name)?\s*[:\-]\s*(.+)",

        r"employer\s*[:\-]\s*(.+)",

        r"organization\s*[:\-]\s*(.+)",

        r"organisation\s*[:\-]\s*(.+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:

            value = match.group(1).strip()

            value = value.split("\n")[0]

            return value[:200]

    return ""


def extract_role_from_text(text):

    patterns = [

        r"job\s*title\s*[:\-]\s*(.+)",

        r"position\s*[:\-]\s*(.+)",

        r"role\s*[:\-]\s*(.+)",

        r"designation\s*[:\-]\s*(.+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:

            value = match.group(1).strip()

            value = value.split("\n")[0]

            return value[:200]

    return ""


def extract_location_from_text(text):

    patterns = [

        r"location\s*[:\-]\s*(.+)",

        r"city\s*[:\-]\s*(.+)",

        r"job\s*location\s*[:\-]\s*(.+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:

            value = match.group(1).strip()

            value = value.split("\n")[0]

            return value[:150]

    return ""


def extract_date_from_text(text):

    patterns = [

        r"(?:application\s*)?date\s*[:\-]\s*([^\n]+)",

        r"applied\s*(?:on)?\s*[:\-]\s*([^\n]+)",

        r"submitted\s*(?:on)?\s*[:\-]\s*([^\n]+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:

            return clean_date(
                match.group(1)
            )

    # General date search
    date_patterns = [

        r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",

        r"\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b",

        r"\b\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}\b"
    ]

    for pattern in date_patterns:

        match = re.search(
            pattern,
            text
        )

        if match:

            return clean_date(
                match.group()
            )

    return ""


# ============================================================
# IMPORT PDF
# ============================================================

def import_pdf():

    file_path = filedialog.askopenfilename(

        title="Select PDF",

        filetypes=[
            ("PDF Files", "*.pdf"),
            ("All Files", "*.*")
        ]
    )

    if not file_path:
        return

    try:

        create_backup()

        imported = 0
        duplicates = 0

        # ----------------------------------------------------
        # FIRST TRY TABLES
        # ----------------------------------------------------

        tables = extract_pdf_tables(
            file_path
        )

        for table in tables:

            if len(table) < 2:
                continue

            try:

                df = pd.DataFrame(
                    table[1:],
                    columns=table[0]
                )

                a, b = import_dataframe(
                    df,
                    os.path.basename(file_path)
                )

                imported += a
                duplicates += b

            except Exception as e:

                print(
                    "Table import failed:",
                    e
                )

        # ----------------------------------------------------
        # THEN TEXT
        # ----------------------------------------------------

        text = extract_pdf_text(
            file_path
        )

        # Try structured Company/Role text
        company = extract_company_from_text(
            text
        )

        role = extract_role_from_text(
            text
        )

        location = extract_location_from_text(
            text
        )

        application_date = extract_date_from_text(
            text
        )

        # If structured data exists
        if company:

            if not role:
                role = "Not specified"

            if not is_duplicate(
                company,
                role
            ):

                data.append({

                    "Company": company,

                    "Job Title": role,

                    "Location": location,

                    "Application Date":
                        application_date,

                    "Source":
                        "PDF Import",

                    "Status":
                        "Applied",

                    "Job URL":
                        "",

                    "Notes":
                        "Imported from PDF"
                })

                imported += 1

            else:

                duplicates += 1

        # ----------------------------------------------------
        # SAVE
        # ----------------------------------------------------

        save_data()

        refresh_table()

        messagebox.showinfo(

            "PDF Import Complete",

            f"File:\n{os.path.basename(file_path)}\n\n"

            f"New applications: {imported}\n"

            f"Duplicates skipped: {duplicates}\n\n"

            f"Total applications: {len(data)}"
        )

    except Exception as e:

        messagebox.showerror(
            "PDF Import Error",
            str(e)
        )


# ============================================================
# MANUAL ENTRY
# ============================================================

def add_application():

    company = company_entry.get().strip()

    role = role_entry.get().strip()

    location = location_entry.get().strip()

    date = date_entry.get().strip()

    source = source_entry.get().strip()

    status = status_combo.get().strip()

    url = url_entry.get().strip()

    notes = notes_entry.get().strip()

    if not company:

        messagebox.showwarning(
            "Missing Company",
            "Please enter company name."
        )

        return

    if not role:

        messagebox.showwarning(
            "Missing Job Title",
            "Please enter job title."
        )

        return

    if is_duplicate(
        company,
        role
    ):

        messagebox.showwarning(
            "Duplicate",
            "This company + role already exists."
        )

        return

    data.append({

        "Company": company,

        "Job Title": role,

        "Location": location,

        "Application Date":
            date,

        "Source":
            source,

        "Status":
            status or "Applied",

        "Job URL":
            url,

        "Notes":
            notes
    })

    save_data()

    refresh_table()

    clear_entries()

    messagebox.showinfo(
        "Added",
        "Application added successfully."
    )


def clear_entries():

    for entry in [
        company_entry,
        role_entry,
        location_entry,
        date_entry,
        source_entry,
        url_entry,
        notes_entry
    ]:

        entry.delete(
            0,
            tk.END
        )

    status_combo.set(
        "Applied"
    )


# ============================================================
# SEARCH
# ============================================================

def search_data():

    query = search_entry.get().strip().lower()

    refresh_table(query)


# ============================================================
# REFRESH TABLE
# ============================================================

def refresh_table(search_query=""):

    for item in tree.get_children():

        tree.delete(item)

    for index, record in enumerate(data):

        combined = " ".join(
            str(v)
            for v in record.values()
        ).lower()

        if (
            search_query
            and search_query not in combined
        ):
            continue

        tree.insert(

            "",

            tk.END,

            iid=str(index),

            values=(

                index + 1,

                record.get(
                    "Company",
                    ""
                ),

                record.get(
                    "Job Title",
                    ""
                ),

                record.get(
                    "Location",
                    ""
                ),

                record.get(
                    "Application Date",
                    ""
                ),

                record.get(
                    "Source",
                    ""
                ),

                record.get(
                    "Status",
                    ""
                ),

                record.get(
                    "Job URL",
                    ""
                ),

                record.get(
                    "Notes",
                    ""
                )
            )
        )

    update_dashboard()


# ============================================================
# DELETE SELECTED
# ============================================================

def delete_selected():

    selected = tree.selection()

    if not selected:
        return

    if not messagebox.askyesno(
        "Confirm",
        "Delete selected application?"
    ):
        return

    indexes = sorted(
        [
            int(x)
            for x in selected
        ],
        reverse=True
    )

    for index in indexes:

        if 0 <= index < len(data):

            del data[index]

    save_data()

    refresh_table()


# ============================================================
# UPDATE STATUS
# ============================================================

def update_status():

    selected = tree.selection()

    if not selected:
        return

    index = int(
        selected[0]
    )

    if index >= len(data):
        return

    status_window = tk.Toplevel(
        root
    )

    status_window.title(
        "Update Status"
    )

    status_window.geometry(
        "300x150"
    )

    tk.Label(
        status_window,
        text="Select new status:"
    ).pack(
        pady=10
    )

    combo = ttk.Combobox(

        status_window,

        values=[
            "Applied",
            "Under Review",
            "Shortlisted",
            "Assessment",
            "Interview",
            "Selected",
            "Offer",
            "Rejected",
            "Withdrawn"
        ],

        state="readonly"
    )

    combo.pack(
        pady=5
    )

    combo.set(
        data[index].get(
            "Status",
            "Applied"
        )
    )

    def save_status():

        data[index][
            "Status"
        ] = combo.get()

        save_data()

        refresh_table()

        status_window.destroy()

    tk.Button(

        status_window,

        text="Save",

        command=save_status

    ).pack(
        pady=10
    )


# ============================================================
# DASHBOARD
# ============================================================

def update_dashboard():

    total = len(data)

    applied = 0
    interviews = 0
    selected = 0
    rejected = 0
    offers = 0

    for record in data:

        status = record.get(
            "Status",
            ""
        ).lower()

        if "reject" in status:
            rejected += 1

        elif "offer" in status:
            offers += 1

        elif "select" in status:
            selected += 1

        elif "interview" in status:
            interviews += 1

        else:
            applied += 1

    dashboard_label.config(

        text=(
            f"TOTAL: {total}    |    "
            f"APPLIED: {applied}    |    "
            f"INTERVIEW: {interviews}    |    "
            f"SELECTED: {selected}    |    "
            f"OFFER: {offers}    |    "
            f"REJECTED: {rejected}"
        )
    )


# ============================================================
# OPEN EXCEL
# ============================================================

def open_excel():

    save_data()

    if os.path.exists(
        MASTER_FILE
    ):

        os.startfile(
            os.path.abspath(
                MASTER_FILE
            )
        )

    else:

        messagebox.showinfo(
            "Not Found",
            "Excel database does not exist yet."
        )


# ============================================================
# EXPORT FILTERED DATA
# ============================================================

def export_data():

    file_path = filedialog.asksaveasfilename(

        title="Export Applications",

        defaultextension=".xlsx",

        filetypes=[
            ("Excel Files", "*.xlsx"),
            ("CSV Files", "*.csv")
        ]
    )

    if not file_path:
        return

    query = search_entry.get().strip().lower()

    records = []

    for record in data:

        combined = " ".join(
            str(v)
            for v in record.values()
        ).lower()

        if (
            not query
            or query in combined
        ):

            records.append(record)

    df = pd.DataFrame(
        records
    )

    if file_path.lower().endswith(
        ".csv"
    ):

        df.to_csv(
            file_path,
            index=False
        )

    else:

        df.to_excel(
            file_path,
            index=False
        )

    messagebox.showinfo(
        "Export Complete",
        f"Exported {len(records)} applications."
    )


# ============================================================
# GUI
# ============================================================

root = tk.Tk()

root.title(
    "Smart Job Application Tracker"
)

root.geometry(
    "1500x800"
)


# ============================================================
# TITLE
# ============================================================

title_label = tk.Label(

    root,

    text="SMART JOB APPLICATION TRACKER",

    font=(
        "Arial",
        20,
        "bold"
    )
)

title_label.pack(
    pady=10
)


# ============================================================
# DASHBOARD
# ============================================================

dashboard_label = tk.Label(

    root,

    text="TOTAL: 0",

    font=(
        "Arial",
        12,
        "bold"
    )
)

dashboard_label.pack(
    pady=5
)


# ============================================================
# MANUAL ENTRY FRAME
# ============================================================

entry_frame = ttk.LabelFrame(

    root,

    text="Add Application Manually"
)

entry_frame.pack(

    fill="x",

    padx=10,

    pady=10
)


fields = [

    ("Company", 0),

    ("Job Title", 1),

    ("Location", 2),

    ("Application Date", 3),

    ("Source", 4),

    ("Job URL", 5),

    ("Notes", 6)
]


entries = {}


for field_name, column in fields:

    tk.Label(

        entry_frame,

        text=field_name

    ).grid(

        row=0,

        column=column,

        padx=5,

        pady=5
    )

    entry = tk.Entry(

        entry_frame,

        width=20
    )

    entry.grid(

        row=1,

        column=column,

        padx=5,

        pady=5
    )

    entries[field_name] = entry


company_entry = entries["Company"]

role_entry = entries["Job Title"]

location_entry = entries["Location"]

date_entry = entries["Application Date"]

source_entry = entries["Source"]

url_entry = entries["Job URL"]

notes_entry = entries["Notes"]


tk.Label(
    entry_frame,
    text="Status"
).grid(
    row=0,
    column=7
)


status_combo = ttk.Combobox(

    entry_frame,

    values=[
        "Applied",
        "Under Review",
        "Shortlisted",
        "Assessment",
        "Interview",
        "Selected",
        "Offer",
        "Rejected",
        "Withdrawn"
    ],

    state="readonly",

    width=16
)

status_combo.grid(
    row=1,
    column=7,
    padx=5
)

status_combo.set(
    "Applied"
)


tk.Button(

    entry_frame,

    text="ADD",

    command=add_application,

    width=12

).grid(

    row=1,

    column=8,

    padx=10
)


# ============================================================
# TOOLBAR
# ============================================================

toolbar = tk.Frame(
    root
)

toolbar.pack(
    fill="x",
    padx=10,
    pady=5
)


tk.Button(

    toolbar,

    text="IMPORT EXCEL / CSV",

    command=import_excel,

    width=20

).pack(
    side="left",
    padx=5
)


tk.Button(

    toolbar,

    text="IMPORT PDF",

    command=import_pdf,

    width=15

).pack(
    side="left",
    padx=5
)


tk.Button(

    toolbar,

    text="OPEN MASTER EXCEL",

    command=open_excel,

    width=18

).pack(
    side="left",
    padx=5
)


tk.Button(

    toolbar,

    text="EXPORT",

    command=export_data,

    width=12

).pack(
    side="left",
    padx=5
)


tk.Button(

    toolbar,

    text="DELETE",

    command=delete_selected,

    width=12

).pack(
    side="left",
    padx=5
)


tk.Button(

    toolbar,

    text="UPDATE STATUS",

    command=update_status,

    width=15

).pack(
    side="left",
    padx=5
)


# ============================================================
# SEARCH
# ============================================================

search_frame = tk.Frame(
    root
)

search_frame.pack(
    fill="x",
    padx=10,
    pady=5
)


tk.Label(
    search_frame,
    text="Search:"
).pack(
    side="left"
)


search_entry = tk.Entry(
    search_frame,
    width=50
)

search_entry.pack(
    side="left",
    padx=5
)


tk.Button(

    search_frame,

    text="SEARCH",

    command=search_data

).pack(
    side="left"
)


tk.Button(

    search_frame,

    text="CLEAR",

    command=lambda: (
        search_entry.delete(
            0,
            tk.END
        ),
        refresh_table()
    )

).pack(
    side="left",
    padx=5
)


# ============================================================
# TABLE
# ============================================================

table_frame = tk.Frame(
    root
)

table_frame.pack(

    fill="both",

    expand=True,

    padx=10,

    pady=10
)


columns = (

    "No",

    "Company",

    "Job Title",

    "Location",

    "Application Date",

    "Source",

    "Status",

    "Job URL",

    "Notes"
)


tree = ttk.Treeview(

    table_frame,

    columns=columns,

    show="headings",

    selectmode="extended"
)


widths = {

    "No": 50,

    "Company": 180,

    "Job Title": 220,

    "Location": 130,

    "Application Date": 120,

    "Source": 130,

    "Status": 120,

    "Job URL": 250,

    "Notes": 250
}


for column in columns:

    tree.heading(
        column,
        text=column
    )

    tree.column(

        column,

        width=widths[column],

        anchor="w"
    )


scroll_y = ttk.Scrollbar(

    table_frame,

    orient="vertical",

    command=tree.yview
)


scroll_x = ttk.Scrollbar(

    table_frame,

    orient="horizontal",

    command=tree.xview
)


tree.configure(

    yscrollcommand=scroll_y.set,

    xscrollcommand=scroll_x.set
)


tree.pack(
    side="left",
    fill="both",
    expand=True
)


scroll_y.pack(
    side="right",
    fill="y"
)


scroll_x.pack(
    side="bottom",
    fill="x"
)


# ============================================================
# START APPLICATION
# ============================================================

load_existing_data()

refresh_table()

root.mainloop()