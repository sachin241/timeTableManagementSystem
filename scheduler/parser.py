"""
parser.py — Unified Multi-Format Parsing Module for the AI Timetable Scheduler (v4)

v4 changes:
  - Keeps CSV / XLSX / DOCX / PDF support
  - Adds Room import support
  - Adds teacher max-hours import support
  - Keeps section-wise subject mapping
  - Supports optional preferred room / room type
  - More robust line parsing for DOCX/PDF
"""

import os
from collections.abc import Iterable

import pandas as pd

from .models import Subject, Teacher, ClassSection, Timetable, Room


REQUIRED_COLUMNS = {"subject", "teacher"}
OPTIONAL_COLUMNS = {
    "hours_per_week",
    "credits",
    "type",
    "section",
    "room",
    "room_type",
    "teacher_max_hours_per_week",
}


def parse_and_store(file_path: str, room_inventory: dict | list | None = None) -> dict:
    """
    Parse any supported file and persist Subject + Teacher + ClassSection + Room data.

    Supported formats:
      .csv  -> pandas read_csv
      .xlsx -> pandas read_excel
      .docx -> line-based paragraphs
      .pdf  -> line-based text extraction

    Returns:
        dict with keys:
          subjects
          teachers
          sections
          rooms
          subject_teacher_map
          subject_hours_map
          teacher_limit_map
          section_subject_data
          section_map
          room_map
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".csv":
        df = _load_csv(file_path)
    elif ext == ".xlsx":
        df = _load_xlsx(file_path)
    elif ext == ".docx":
        df = _load_docx(file_path)
    elif ext == ".pdf":
        df = _load_pdf(file_path)
    else:
        raise ValueError(
            f"Unsupported file format: '{ext}'. Supported: .csv, .xlsx, .docx, .pdf"
        )

    return _process_dataframe(df, room_inventory=room_inventory)


def _load_csv(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    return df


def _load_xlsx(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    return df


def _load_docx(path: str) -> pd.DataFrame:
    """
    DOCX line format:
      Teacher - Subject - Hours
      Teacher - Subject - Hours - Type
      Teacher - Subject - Hours - Type - Section
      Teacher - Subject - Hours - Type - Section - Room
    """
    from docx import Document

    doc = Document(path)
    rows = []

    for para in doc.paragraphs:
        line = para.text.strip()
        if not line or "-" not in line:
            continue
        parsed = _parse_text_line(line)
        if parsed:
            rows.append(parsed)

    if not rows:
        raise ValueError(
            "No valid entries found in DOCX file. "
            "Expected format: 'Teacher - Subject - Hours' per line."
        )

    return pd.DataFrame(rows)


def _load_pdf(path: str) -> pd.DataFrame:
    """
    PDF line format: same as DOCX.
    """
    import pdfplumber

    rows = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                line = line.strip()
                if not line or "-" not in line:
                    continue
                parsed = _parse_text_line(line)
                if parsed:
                    rows.append(parsed)

    if not rows:
        raise ValueError(
            "No valid entries found in PDF file. "
            "Expected format: 'Teacher - Subject - Hours' per line."
        )

    return pd.DataFrame(rows)


def _parse_text_line(line: str) -> dict | None:
    """
    Parse one DOCX/PDF line.

    Supported:
      Teacher - Subject - Hours
      Teacher - Subject - Hours - Type
      Teacher - Subject - Hours - Type - Section
      Teacher - Subject - Hours - Type - Section - Room

    Notes:
      - Third token must be numeric.
      - Type may be theory/lab.
      - Section and room are optional.
    """
    parts = [p.strip() for p in line.split("-")]
    parts = [p for p in parts if p]

    if len(parts) < 3:
        return None

    teacher = parts[0]
    subject = parts[1]

    try:
        hours_or_credits = int(parts[2])
    except ValueError:
        return None

    row = {
        "teacher": teacher,
        "subject": subject,
        "hours_per_week": hours_or_credits,
    }

    # Optional type
    idx = 3
    if len(parts) > idx:
        stype = parts[idx].lower()
        if stype in ("theory", "lab"):
            row["type"] = stype
            idx += 1

    # Optional section
    if len(parts) > idx:
        row["section"] = parts[idx]
        idx += 1

    # Optional room
    if len(parts) > idx:
        row["room"] = parts[idx]
        idx += 1

    return row


def _normalize_room_inventory(room_inventory: dict | list | None) -> list[dict]:
    """
    Normalize room data from the UI into a list of room specs.

    Accepted shapes:
      - {"room_names": "...", "room_count": 2, "lab_room_names": "..."}
      - [{"name": "Room 101", "is_lab": False}, ...]
      - ["Room 101", "Room 102"]
    """
    if not room_inventory:
        return []

    specs: list[dict] = []

    if isinstance(room_inventory, dict):
        room_names = room_inventory.get("room_names", "")
        lab_room_names = room_inventory.get("lab_room_names", "")
        room_count = room_inventory.get("room_count", 0)

        lab_name_set = {
            item.strip().lower()
            for item in str(lab_room_names).split(",")
            if str(item).strip()
        }

        if isinstance(room_names, str):
            explicit_names = [item.strip() for item in room_names.split(",") if item.strip()]
        elif isinstance(room_names, Iterable):
            explicit_names = [str(item).strip() for item in room_names if str(item).strip()]
        else:
            explicit_names = []

        for name in explicit_names:
            specs.append({"name": name, "is_lab": name.lower() in lab_name_set or "lab" in name.lower()})

        try:
            count = int(room_count or 0)
        except (TypeError, ValueError):
            count = 0

        for idx in range(1, count + 1):
            generated_name = f"Room {idx}"
            specs.append({"name": generated_name, "is_lab": False})

        return specs

    if isinstance(room_inventory, Iterable) and not isinstance(room_inventory, (str, bytes)):
        for item in room_inventory:
            if isinstance(item, str):
                name = item.strip()
                if name:
                    specs.append({"name": name, "is_lab": "lab" in name.lower()})
            elif isinstance(item, dict):
                name = str(item.get("name", "")).strip()
                if not name:
                    continue
                specs.append({"name": name, "is_lab": bool(item.get("is_lab", False))})

    return specs


def _process_dataframe(df: pd.DataFrame, room_inventory: dict | list | None = None) -> dict:
    """
    Normalize the dataframe and write records to DB.
    """

    # Normalize column names again in case source had mixed formatting
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]

    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(
            f"File is missing required columns: {missing}. "
            f"Expected at minimum: subject, teacher."
        )

    has_hours = "hours_per_week" in df.columns
    has_credits = "credits" in df.columns

    if not has_hours and not has_credits:
        raise ValueError("File must contain either 'hours_per_week' or 'credits' column.")

    # Clean string columns
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

    # Default type
    if "type" not in df.columns:
        df["type"] = "theory"
    else:
        df["type"] = df["type"].replace("", "theory").str.lower()
        df.loc[~df["type"].isin(["theory", "lab"]), "type"] = "theory"

    # Default section
    if "section" not in df.columns:
        df["section"] = "A"
    else:
        df["section"] = df["section"].replace("", "A").astype(str).str.strip()

    # Optional room
    if "room" not in df.columns:
        df["room"] = ""
    else:
        df["room"] = df["room"].fillna("").astype(str).str.strip()

    # Optional room_type
    if "room_type" not in df.columns:
        df["room_type"] = ""
    else:
        df["room_type"] = df["room_type"].fillna("").astype(str).str.strip().str.lower()

    # Optional teacher max hours
    if "teacher_max_hours_per_week" in df.columns:
        df["teacher_max_hours_per_week"] = pd.to_numeric(
            df["teacher_max_hours_per_week"], errors="coerce"
        ).fillna(0).astype(int)
    else:
        df["teacher_max_hours_per_week"] = 0

    # Numeric conversion
    if has_credits:
        df["credits"] = pd.to_numeric(df["credits"], errors="coerce").fillna(0).astype(int)

    if has_hours:
        df["hours_per_week"] = pd.to_numeric(df["hours_per_week"], errors="coerce").fillna(0).astype(int)
    else:
        df["hours_per_week"] = 0

    # Convert credits to hours when needed
    if has_credits:
        needs_conversion = df["hours_per_week"] <= 0
        lab_mask = needs_conversion & (df["type"] == "lab")
        theory_mask = needs_conversion & (df["type"] == "theory")

        df.loc[lab_mask, "hours_per_week"] = df.loc[lab_mask, "credits"] * 2
        df.loc[theory_mask, "hours_per_week"] = df.loc[theory_mask, "credits"]

    invalid = df[df["hours_per_week"] <= 0]
    if not invalid.empty:
        bad_subjects = invalid["subject"].unique().tolist()
        raise ValueError(
            f"Subjects with 0 or missing hours: {bad_subjects}. "
            f"Provide either hours_per_week or credits for each subject."
        )

    # Clear old data
    Timetable.objects.all().delete()
    Subject.objects.all().delete()
    Teacher.objects.all().delete()
    ClassSection.objects.all().delete()
    Room.objects.all().delete()

    # Sections
    section_names = sorted(df["section"].dropna().unique().tolist())
    section_map = {}
    for sname in section_names:
        sec, _ = ClassSection.objects.get_or_create(name=sname)
        section_map[sname] = sec

    # Rooms
    room_map = {}
    room_specs = {}

    for rname in sorted([r for r in df["room"].unique().tolist() if r]):
        room_rows = df[df["room"] == rname]
        room_type_values = room_rows["room_type"].dropna().tolist()
        room_specs[rname] = any(v == "lab" for v in room_type_values)

    for spec in _normalize_room_inventory(room_inventory):
        name = spec["name"]
        if not name:
            continue
        room_specs[name] = room_specs.get(name, False) or bool(spec.get("is_lab", False))

    for rname in sorted(room_specs.keys()):
        room, created = Room.objects.get_or_create(
            name=rname,
            defaults={"is_lab": room_specs[rname]},
        )
        if not created and room_specs[rname] and not room.is_lab:
            room.is_lab = True
            room.save()

        room_map[rname] = room

    # Subjects
    agg_dict = {"hours_per_week": "max", "type": "first"}
    if has_credits:
        agg_dict["credits"] = "max"

    subject_info = df.groupby("subject").agg(agg_dict).to_dict("index")

    subject_map = {}
    for name, info in subject_info.items():
        subj = Subject(
            name=name,
            hours_per_week=int(info["hours_per_week"]),
            subject_type=info.get("type", "theory"),
        )
        if has_credits and info.get("credits", 0) > 0:
            subj.credits = int(info["credits"])
        subj.save()
        subject_map[name] = subj

    # Teachers and subject links
    teacher_map = {}
    subject_teacher_map = {}
    teacher_limit_map = {}

    for _, row in df.iterrows():
        t_name = row["teacher"]
        s_name = row["subject"]

        if not t_name or not s_name:
            continue

        teacher, _ = Teacher.objects.get_or_create(name=t_name)
        subj = subject_map[s_name]

        teacher.subjects.add(subj)

        # Store max hours if provided
        limit = int(row.get("teacher_max_hours_per_week", 0) or 0)
        if limit > 0:
            if teacher.max_hours_per_week != limit:
                teacher.max_hours_per_week = limit
                teacher.save()

        teacher_map[t_name] = teacher

        # Keep first mapped teacher for the subject
        if subj.id not in subject_teacher_map:
            subject_teacher_map[subj.id] = teacher

        teacher_limit_map[teacher.id] = teacher.max_hours_per_week

    # Section-wise subject data
    section_subject_data = {}
    for section_name in section_names:
        section_df = df[df["section"] == section_name]
        items = []
        seen_subjects = set()

        for _, row in section_df.iterrows():
            subj = subject_map[row["subject"]]
            if subj.id in seen_subjects:
                continue

            teacher = row.get("teacher", "").strip()
            teacher_obj = teacher_map.get(teacher)
            if teacher_obj is None:
                teacher_obj, _ = Teacher.objects.get_or_create(name=teacher)
                teacher_map[teacher] = teacher_obj
                teacher_limit_map[teacher_obj.id] = teacher_obj.max_hours_per_week

            room_name = row.get("room", "").strip()
            room_obj = room_map.get(room_name) if room_name else None

            items.append(
                {
                    "subject_id": subj.id,
                    "teacher_id": teacher_obj.id,
                    "hours": int(subj.hours_per_week),
                    "type": subj.subject_type,
                    "room_id": room_obj.id if room_obj else None,
                    "room_name": room_name or None,
                }
            )
            seen_subjects.add(subj.id)

        section_subject_data[section_name] = items

    return {
        "subjects": list(subject_map.values()),
        "teachers": list(teacher_map.values()),
        "sections": list(section_map.values()),
        "rooms": list(room_map.values()),
        "subject_teacher_map": subject_teacher_map,
        "subject_hours_map": {s.id: s.hours_per_week for s in subject_map.values()},
        "teacher_limit_map": teacher_limit_map,
        "section_subject_data": section_subject_data,
        "section_map": section_map,
        "room_map": room_map,
    }
