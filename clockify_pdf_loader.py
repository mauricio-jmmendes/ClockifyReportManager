"""
Load Clockify Detailed report data from PDF exports.

Clockify PDF exports render each time entry as a multi-line block:
  1. Date + description start + duration + user
  2. Description continuation (optional, may repeat)
  3. Start/end time range
  4. Project line (client/project/tags)
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime

import pandas as pd
import pdfplumber

# Output schema expected by the rest of the converter
DETAILED_COLUMNS = [
    "Project",
    "Client",
    "Description",
    "User",
    "Tags",
    "Start Date",
    "Start Time",
    "End Date",
    "End Time",
    "Duration (h)",
]

DATE_PATTERN = re.compile(r"^\d{2}/\d{2}/\d{4}$")
DURATION_PATTERN = re.compile(r"^\d+:\d{2}:\d{2}$")
TIME_RANGE_PATTERN = re.compile(r"^(\d+:\d{2}:\d{2})\s*-\s*(\d+:\d{2}:\d{2})$")
DATE_RANGE_PATTERN = re.compile(
    r"(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})"
)
FOOTER_PATTERN = re.compile(r"created with clockify", re.IGNORECASE)
HEADER_LABELS = {"date", "description", "duration", "user", "project", "client", "tags"}

# Default column boundaries from Clockify PDF layout
DEFAULT_COLUMN_BOUNDS = {
    "date": (0, 80),
    "description": (80, 280),
    "duration": (280, 400),
    "user": (400, 1000),
}


@dataclass
class ParsedLine:
    date: str
    description: str
    duration: str
    user: str
    top: float


def _normalize_duration(value: str) -> str:
    """Convert Clockify duration strings to HH:MM:SS."""
    if not value or not str(value).strip():
        return "00:00:00"

    value = str(value).strip()

    if DATE_PATTERN.match(value):
        return "00:00:00"

    if DURATION_PATTERN.match(value):
        parts = value.split(":")
        return f"{int(parts[0]):02d}:{parts[1]}:{parts[2]}"

    if re.match(r"^\d+:\d{2}$", value):
        hours, minutes = value.split(":")
        return f"{int(hours):02d}:{minutes}:00"

    decimal_match = re.match(r"^(\d+(?:[.,]\d+)?)$", value)
    if decimal_match:
        decimal_hours = float(decimal_match.group(1).replace(",", "."))
        total_seconds = int(round(decimal_hours * 3600))
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    return value


def _parse_date(value: str):
    """Parse dd/mm/yyyy dates from PDF text."""
    if not value or not str(value).strip():
        return pd.NA

    value = str(value).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return value


def _parse_time(value: str):
    if not value or not str(value).strip():
        return pd.NA
    return _normalize_duration(str(value).strip())


def _parse_project_line(line: str) -> tuple[str, str | float, str | float]:
    """Split a Clockify project footer line into project, client, and tags."""
    tag_match = re.search(r"\[([^\]]+)\]\s*$", line)
    tag = tag_match.group(1).strip() if tag_match else pd.NA
    core = line[: tag_match.start()].strip() if tag_match else line.strip()

    parts = [part.strip() for part in core.split(" - ") if part.strip()]
    if len(parts) >= 2:
        project = " - ".join(parts[1:]).rstrip(" -")
        return project, parts[0], tag

    return core, pd.NA, tag


def _is_footer_line(text: str) -> bool:
    return bool(FOOTER_PATTERN.search(text))


def _is_skippable_line(parsed: ParsedLine) -> bool:
    combined = " ".join(
        part for part in (parsed.date, parsed.description, parsed.duration, parsed.user) if part
    ).strip()
    if not combined:
        return True
    if _is_footer_line(combined):
        return True

    lowered = combined.lower()
    if lowered in {"detailed report", "total:"}:
        return True
    if DATE_RANGE_PATTERN.search(combined) and not parsed.date:
        return True

    labels = {part.strip().lower() for part in combined.split()}
    if labels and labels <= HEADER_LABELS:
        return True

    # Report total duration shown under the title
    if parsed.duration and DURATION_PATTERN.match(parsed.duration) and not parsed.date and not parsed.description:
        return True

    return False


def _parse_entry_start_line(words: list[dict]) -> ParsedLine:
    """Parse the first line of an entry where description can span into later columns."""
    words = sorted(words, key=lambda item: item["x0"])
    date_parts: list[str] = []
    user_parts: list[str] = []
    desc_parts: list[str] = []
    duration = ""

    for word in words:
        text = word["text"]
        if word["x0"] < DEFAULT_COLUMN_BOUNDS["date"][1]:
            date_parts.append(text)
        elif word["x0"] >= DEFAULT_COLUMN_BOUNDS["user"][0]:
            user_parts.append(text)
        elif DURATION_PATTERN.match(text):
            duration = text
        else:
            desc_parts.append(text)

    return ParsedLine(
        date=" ".join(date_parts).strip(),
        description=" ".join(desc_parts).strip(),
        duration=duration.strip(),
        user=" ".join(user_parts).strip(),
        top=words[0]["top"],
    )


def _parse_followup_line(words: list[dict]) -> ParsedLine:
    """Parse continuation, time-range, and project lines using fixed layout bounds."""
    words = sorted(words, key=lambda item: item["x0"])
    user = " ".join(
        word["text"]
        for word in words
        if word["x0"] >= DEFAULT_COLUMN_BOUNDS["user"][0]
    ).strip()

    duration_words = [
        word
        for word in words
        if DEFAULT_COLUMN_BOUNDS["duration"][0] <= word["x0"] < DEFAULT_COLUMN_BOUNDS["duration"][1]
    ]
    duration = " ".join(word["text"] for word in duration_words).strip()

    if TIME_RANGE_PATTERN.match(duration):
        return ParsedLine("", "", duration, user, words[0]["top"])

    description = " ".join(
        word["text"]
        for word in words
        if DEFAULT_COLUMN_BOUNDS["description"][0] <= word["x0"] < DEFAULT_COLUMN_BOUNDS["user"][0]
        and not DURATION_PATTERN.match(word["text"])
    ).strip()

    return ParsedLine("", description, "", user, words[0]["top"])


def _split_line_words(words: list[dict], bounds: dict[str, tuple[float, float]]) -> ParsedLine:
    buckets = {"date": [], "description": [], "duration": [], "user": []}

    for word in sorted(words, key=lambda item: item["x0"]):
        center = (word["x0"] + word["x1"]) / 2
        for name, (start, end) in bounds.items():
            if start <= center < end:
                buckets[name].append(word["text"])
                break

    return ParsedLine(
        date=" ".join(buckets["date"]).strip(),
        description=" ".join(buckets["description"]).strip(),
        duration=" ".join(buckets["duration"]).strip(),
        user=" ".join(buckets["user"]).strip(),
        top=words[0]["top"],
    )


def _column_bounds_from_header(header_words: list[dict]) -> dict[str, tuple[float, float]]:
    positions: dict[str, float] = {}
    for word in sorted(header_words, key=lambda item: item["x0"]):
        label = word["text"].strip().lower()
        if label in {"date", "description", "duration", "user"}:
            positions[label] = word["x0"]

    if len(positions) < 2:
        return DEFAULT_COLUMN_BOUNDS

    ordered = sorted(positions.items(), key=lambda item: item[1])
    bounds: dict[str, tuple[float, float]] = {}

    for index, (name, x_start) in enumerate(ordered):
        lower = 0 if index == 0 else (ordered[index - 1][1] + x_start) / 2
        upper = (
            x_start + 500
            if index == len(ordered) - 1
            else (x_start + ordered[index + 1][1]) / 2
        )
        bounds[name] = (lower, upper)

    return bounds


def _find_header_bounds(words: list[dict]) -> dict[str, tuple[float, float]] | None:
    rows: dict[float, list[dict]] = {}
    for word in words:
        rows.setdefault(round(word["top"], 1), []).append(word)

    for top in sorted(rows.keys()):
        row_words = rows[top]
        labels = {word["text"].strip().lower() for word in row_words}
        if len(labels & HEADER_LABELS) >= 2:
            return _column_bounds_from_header(row_words)

    return None


def _is_valid_entry_start(line: ParsedLine) -> bool:
    return (
        bool(line.date and DATE_PATTERN.match(line.date))
        and bool(line.duration and DURATION_PATTERN.match(line.duration))
        and bool(line.user)
    )


def _group_words_into_lines(words: list[dict], bounds: dict[str, tuple[float, float]]) -> list[ParsedLine]:
    rows: dict[float, list[dict]] = {}
    for word in words:
        if word["top"] >= 760 or _is_footer_line(word.get("text", "")):
            continue
        rows.setdefault(round(word["top"], 1), []).append(word)

    lines = []
    for row_words in rows.values():
        has_date = any(
            DATE_PATTERN.match(word["text"])
            for word in row_words
            if word["x0"] < DEFAULT_COLUMN_BOUNDS["date"][1]
        )
        parsed = (
            _parse_entry_start_line(row_words)
            if has_date
            else _parse_followup_line(row_words)
        )
        if not _is_skippable_line(parsed):
            lines.append(parsed)

    return sorted(lines, key=lambda line: line.top)


def _split_into_blocks(lines: list[ParsedLine]) -> list[list[ParsedLine]]:
    blocks: list[list[ParsedLine]] = []
    current: list[ParsedLine] = []

    for line in lines:
        if _is_valid_entry_start(line):
            if current:
                blocks.append(current)
            current = [line]
        elif current:
            current.append(line)

    if current:
        blocks.append(current)

    return blocks


def _parse_block(block: list[ParsedLine]) -> dict:
    entry: dict = {}
    description_parts: list[str] = []
    desc_only_lines: list[str] = []

    first = block[0]
    entry["Start Date"] = first.date
    entry["Duration (h)"] = first.duration
    entry["User"] = first.user
    if first.description:
        description_parts.append(first.description)

    for line in block[1:]:
        if line.duration and TIME_RANGE_PATTERN.match(line.duration):
            start_time, end_time = TIME_RANGE_PATTERN.match(line.duration).groups()
            entry["Start Time"] = start_time
            entry["End Time"] = end_time
        elif line.description and not line.duration and not line.date and not line.user:
            desc_only_lines.append(line.description)

    if desc_only_lines:
        description_parts.extend(desc_only_lines[:-1])
        project, client, tag = _parse_project_line(desc_only_lines[-1])
        entry["Project"] = project
        entry["Client"] = client
        entry["Tags"] = tag

    entry["Description"] = " ".join(description_parts).strip()
    return entry


def _inherit_missing_fields(records: list[dict]) -> list[dict]:
    """Fill project metadata for entries split across PDF pages."""
    for index, record in enumerate(records):
        if record.get("Project") not in (None, "", "General"):
            continue

        for other_index in (index + 1, index - 1):
            if not 0 <= other_index < len(records):
                continue
            other = records[other_index]
            if other.get("Project") in (None, "", "General"):
                continue
            if other.get("Start Date") != record.get("Start Date"):
                continue

            record["Project"] = other["Project"]
            if _has_value(other.get("Client")):
                record["Client"] = other["Client"]
            if _has_value(other.get("Tags")):
                record["Tags"] = other["Tags"]
            break

    return records


def _has_value(value) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        return value != ""
    return not pd.isna(value)


def _entry_to_record(entry: dict) -> dict:
    record = {column: pd.NA for column in DETAILED_COLUMNS}

    for key, value in entry.items():
        if key in record and _has_value(value):
            record[key] = value

    if pd.notna(record["Duration (h)"]):
        record["Duration (h)"] = _normalize_duration(record["Duration (h)"])

    if pd.notna(record["Start Date"]):
        record["Start Date"] = _parse_date(record["Start Date"])
        record["End Date"] = record["Start Date"]

    if pd.notna(record["Start Time"]):
        record["Start Time"] = _parse_time(record["Start Time"])

    if pd.notna(record["End Time"]):
        record["End Time"] = _parse_time(record["End Time"])

    if not _has_value(record["Project"]):
        record["Project"] = "General"

    return record


def _extract_entries_from_page(page, bounds: dict[str, tuple[float, float]]) -> list[dict]:
    words = page.extract_words()
    lines = _group_words_into_lines(words, bounds)
    blocks = _split_into_blocks(lines)
    return [_parse_block(block) for block in blocks if _is_valid_entry_start(block[0])]


def load_detailed_data_from_pdf(pdf_file: str) -> pd.DataFrame:
    """Load Clockify Detailed report rows from a PDF export."""
    all_entries: list[dict] = []
    bounds = DEFAULT_COLUMN_BOUNDS

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            page_words = [
                word for word in page.extract_words()
                if word["top"] < 760 and not _is_footer_line(word.get("text", ""))
            ]
            header_bounds = _find_header_bounds(page_words)
            if header_bounds:
                bounds = header_bounds

            all_entries.extend(_extract_entries_from_page(page, bounds))

    if not all_entries:
        return pd.DataFrame(columns=DETAILED_COLUMNS)

    records = [_entry_to_record(entry) for entry in all_entries]
    records = _inherit_missing_fields(records)
    return pd.DataFrame(records, columns=DETAILED_COLUMNS)


def parse_date_range_from_pdf_text(pdf_file: str) -> tuple[str | None, str | None]:
    """Extract dd/mm/yyyy date range from PDF header text."""
    with pdfplumber.open(pdf_file) as pdf:
        if not pdf.pages:
            return None, None
        text = pdf.pages[0].extract_text() or ""
        match = DATE_RANGE_PATTERN.search(text)
        if match:
            return match.group(1), match.group(2)
    return None, None
