"""
parser.py — Unified Multi-Format Parsing Module for the AI Timetable Scheduler (v3).

v3 Changes:
  - Accepts CSV, XLSX, PDF, and DOCX files.
  - Auto-detects file type by extension.
  - Optional columns: credits, type (theory/lab), section (class/branch).
  - Credits → hours conversion: theory 1:1, lab 1:2.
  - If section column present → creates ClassSection objects per unique value.
  - Returns 'sections' list and per-section subject mapping in data dict.

Supported formats:
  CSV/XLSX (columnar):
    Required: subject, teacher, hours_per_week (or credits)
    Optional: credits, type, section

  DOCX/PDF (line-based):
    Format per line: "FacultyName - SubjectName - HoursPerWeek"
    Or: "FacultyName - SubjectName - Credits - Type"
"""

import os
import re
import pandas as pd
from .models import Subject, Teacher, ClassSection, Timetable


# ---------------------------------------------------------------------------
# Required and optional column sets
# ---------------------------------------------------------------------------

REQUIRED_COLUMNS = {'subject', 'teacher'}
OPTIONAL_COLUMNS = {'hours_per_week', 'credits', 'type', 'section'}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse_and_store(file_path: str) -> dict:
    """
    Parse any supported file and persist Subject + Teacher + ClassSection data.

    Auto-detects format by file extension:
      .csv  → pandas read_csv
      .xlsx → pandas read_excel (openpyxl)
      .docx → python-docx paragraph extraction
      .pdf  → pdfplumber text extraction

    Args:
        file_path: Absolute path to the uploaded file.

    Returns:
        dict with keys:
          'subjects'              — list of Subject ORM objects
          'teachers'              — list of Teacher ORM objects
          'sections'              — list of ClassSection ORM objects
          'subject_teacher_map'   — {subject_id: Teacher ORM object}
          'subject_hours_map'     — {subject_id: hours_per_week (int)}
          'section_subject_data'  — {section_name: [{'subject_id', 'teacher_id', 'hours'}]}
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext == '.csv':
        df = _load_csv(file_path)
    elif ext == '.xlsx':
        df = _load_xlsx(file_path)
    elif ext == '.docx':
        df = _load_docx(file_path)
    elif ext == '.pdf':
        df = _load_pdf(file_path)
    else:
        raise ValueError(
            f"Unsupported file format: '{ext}'. "
            f"Supported: .csv, .xlsx, .docx, .pdf"
        )

    return _process_dataframe(df)


# ---------------------------------------------------------------------------
# File Loaders
# ---------------------------------------------------------------------------

def _load_csv(path: str) -> pd.DataFrame:
    """Load and normalise a CSV file."""
    df = pd.read_csv(path)
    df.columns = [c.strip().lower().replace(' ', '_') for c in df.columns]
    return df


def _load_xlsx(path: str) -> pd.DataFrame:
    """Load and normalise an XLSX file using openpyxl backend."""
    df = pd.read_excel(path, engine='openpyxl')
    df.columns = [c.strip().lower().replace(' ', '_') for c in df.columns]
    return df


def _load_docx(path: str) -> pd.DataFrame:
    """
    Parse a DOCX file with line-based format.

    Expected format per paragraph/line:
      "FacultyName - SubjectName - HoursPerWeek"
      "FacultyName - SubjectName - Credits - Type"
      "FacultyName - SubjectName - HoursPerWeek - Type - Section"

    Lines not matching pattern are silently skipped.
    """
    from docx import Document

    doc = Document(path)
    rows = []
    for para in doc.paragraphs:
        line = para.text.strip()
        if not line or '-' not in line:
            continue
        parsed = _parse_text_line(line)
        if parsed:
            rows.append(parsed)

    if not rows:
        raise ValueError(
            "No valid entries found in DOCX file. "
            "Expected format: 'FacultyName - SubjectName - HoursPerWeek' per line."
        )

    return pd.DataFrame(rows)


def _load_pdf(path: str) -> pd.DataFrame:
    """
    Parse a PDF file with line-based format (same as DOCX).
    Uses pdfplumber for text extraction.
    """
    import pdfplumber

    rows = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                line = line.strip()
                if not line or '-' not in line:
                    continue
                parsed = _parse_text_line(line)
                if parsed:
                    rows.append(parsed)

    if not rows:
        raise ValueError(
            "No valid entries found in PDF file. "
            "Expected format: 'FacultyName - SubjectName - HoursPerWeek' per line."
        )

    return pd.DataFrame(rows)


def _parse_text_line(line: str) -> dict | None:
    """
    Parse a single text line from DOCX/PDF into a row dict.

    Supported formats:
      "Teacher - Subject - Hours"
      "Teacher - Subject - Hours - Type"
      "Teacher - Subject - Hours - Type - Section"
      "Teacher - Subject - Credits - Type"  (credits resolved by type)

    Returns dict or None if line doesn't match.
    """
    parts = [p.strip() for p in line.split('-')]
    if len(parts) < 3:
        return None

    teacher = parts[0]
    subject = parts[1]

    # Third part: hours or credits (must be a number)
    try:
        hours_or_credits = int(parts[2])
    except ValueError:
        return None

    row = {
        'teacher': teacher,
        'subject': subject,
        'hours_per_week': hours_or_credits,
    }

    # Optional 4th part: type (theory/lab)
    if len(parts) >= 4:
        stype = parts[3].strip().lower()
        if stype in ('theory', 'lab'):
            row['type'] = stype
        # If type is 'lab' and no explicit hours_per_week, treat 3rd part as credits
        if stype == 'lab' and hours_or_credits <= 3:
            # Heuristic: small numbers for labs are likely credits
            row['credits'] = hours_or_credits
            row['hours_per_week'] = hours_or_credits * 2

    # Optional 5th part: section
    if len(parts) >= 5:
        row['section'] = parts[4].strip()

    return row


# ---------------------------------------------------------------------------
# DataFrame Processing (shared across all formats)
# ---------------------------------------------------------------------------

def _process_dataframe(df: pd.DataFrame) -> dict:
    """
    Process a normalised DataFrame into ORM objects.

    Columns handled:
      Required: subject, teacher
      Required (one of): hours_per_week OR credits
      Optional: type (theory/lab), section (class/branch), credits
    """
    # --- Validate required columns ---
    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(
            f"File is missing required columns: {missing}. "
            f"Expected at minimum: subject, teacher, and either hours_per_week or credits."
        )

    # Must have at least one of hours_per_week or credits
    has_hours = 'hours_per_week' in df.columns
    has_credits = 'credits' in df.columns
    if not has_hours and not has_credits:
        raise ValueError(
            "File must contain either 'hours_per_week' or 'credits' column."
        )

    # Strip whitespace from string columns
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].str.strip()

    # Handle type column (default: theory)
    if 'type' not in df.columns:
        df['type'] = 'theory'
    else:
        df['type'] = df['type'].fillna('theory').str.lower()
        df.loc[~df['type'].isin(['theory', 'lab']), 'type'] = 'theory'

    # Handle section column (default: A)
    if 'section' not in df.columns:
        df['section'] = 'A'
    else:
        df['section'] = df['section'].fillna('A').astype(str).str.strip()

    # Handle credits → hours conversion
    if has_credits:
        df['credits'] = pd.to_numeric(df['credits'], errors='coerce').fillna(0).astype(int)

    if has_hours:
        df['hours_per_week'] = pd.to_numeric(df['hours_per_week'], errors='coerce').fillna(0).astype(int)
    else:
        df['hours_per_week'] = 0

    # Convert credits to hours where hours_per_week is 0
    mask_needs_conversion = df['hours_per_week'] <= 0
    if has_credits:
        lab_mask = mask_needs_conversion & (df['type'] == 'lab')
        theory_mask = mask_needs_conversion & (df['type'] == 'theory')
        df.loc[lab_mask, 'hours_per_week'] = df.loc[lab_mask, 'credits'] * 2
        df.loc[theory_mask, 'hours_per_week'] = df.loc[theory_mask, 'credits']

    # Validate: all rows must have hours_per_week > 0 after conversion
    invalid = df[df['hours_per_week'] <= 0]
    if not invalid.empty:
        bad_subjects = invalid['subject'].unique().tolist()
        raise ValueError(
            f"Subjects with 0 or missing hours: {bad_subjects}. "
            f"Provide either 'hours_per_week' or 'credits' for each subject."
        )

    # --- Clear old data ---
    Timetable.objects.all().delete()
    Subject.objects.all().delete()
    Teacher.objects.all().delete()
    ClassSection.objects.all().delete()

    # --- Create ClassSection objects ---
    section_names = sorted(df['section'].unique())
    section_map: dict = {}
    for sname in section_names:
        sec, _ = ClassSection.objects.get_or_create(name=sname)
        section_map[sname] = sec

    # --- Create Subject objects ---
    # Group by subject name, take max hours, first type
    agg_dict = {'hours_per_week': 'max', 'type': 'first'}
    if has_credits and 'credits' in df.columns:
        agg_dict['credits'] = 'max'

    subject_info = df.groupby('subject').agg(agg_dict).to_dict('index')


    subject_map: dict = {}  # name → Subject ORM
    for name, info in subject_info.items():
        subj = Subject(
            name=name,
            hours_per_week=info['hours_per_week'],
            subject_type=info.get('type', 'theory'),
        )
        if has_credits and info.get('credits', 0) > 0:
            subj.credits = info['credits']
        subj.save()
        subject_map[name] = subj

    # --- Create Teacher objects and link subjects ---
    teacher_map: dict = {}
    subject_teacher_map: dict = {}

    for _, row in df.iterrows():
        t_name = row['teacher']
        s_name = row['subject']

        if t_name not in teacher_map:
            teacher, _ = Teacher.objects.get_or_create(name=t_name)
            teacher_map[t_name] = teacher

        teacher = teacher_map[t_name]
        subj = subject_map[s_name]
        teacher.subjects.add(subj)
        subject_teacher_map[subj.id] = teacher

    # --- Build per-section subject data ---
    section_subject_data: dict = {}
    for section_name in section_names:
        section_df = df[df['section'] == section_name]
        items = []
        seen_subjects = set()
        for _, row in section_df.iterrows():
            subj = subject_map[row['subject']]
            if subj.id not in seen_subjects:
                items.append({
                    'subject_id': subj.id,
                    'teacher_id': subject_teacher_map[subj.id].id,
                    'hours': subj.hours_per_week,
                    'type': subj.subject_type,
                })
                seen_subjects.add(subj.id)
        section_subject_data[section_name] = items

    return {
        'subjects': list(subject_map.values()),
        'teachers': list(teacher_map.values()),
        'sections': list(section_map.values()),
        'subject_teacher_map': subject_teacher_map,
        'subject_hours_map': {s.id: s.hours_per_week for s in subject_map.values()},
        'section_subject_data': section_subject_data,
        'section_map': section_map,
    }
