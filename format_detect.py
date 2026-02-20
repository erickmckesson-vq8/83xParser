"""Auto-detect healthcare file format from content."""

import json
import re


def detect_format(content):
    """Detect the healthcare data format of the given content.

    Returns one of:
        'x12'    - EDI X12 (835, 837, 270, 271, etc.)
        'hl7v2'  - HL7 v2.x pipe-delimited
        'fhir'   - FHIR JSON or XML
        'cda'    - HL7 v3 / CDA XML
        'ncpdp'  - NCPDP Telecommunications
        'csv'    - Delimited text (CSV/TSV)
        None     - Unknown format
    """
    if not content or not content.strip():
        return None

    # Preserve control characters for NCPDP detection
    cleaned = content.lstrip("\ufeff")
    stripped = cleaned.strip()

    # --- NCPDP: check BEFORE stripping since control chars may be significant ---
    # NCPDP uses 0x1C (FS), 0x1D (GS), 0x1E (RS) as delimiters
    if cleaned and cleaned[0] in ("\x1c", "\x1d", "\x1e"):
        return "ncpdp"
    if "\x1c" in content[:200] or "\x1d" in content[:200] or "\x1e" in content[:200]:
        # Contains NCPDP control characters
        return "ncpdp"
    # NCPDP can also start with a 6-digit BIN number followed by version
    if re.match(r"^\d{6}(51|D0)", stripped):
        return "ncpdp"

    # --- X12: starts with ISA ---
    if stripped[:3].upper() == "ISA" and len(stripped) >= 106:
        return "x12"

    # --- HL7 v2.x: starts with MSH, FHS, or BHS followed by field separator ---
    if re.match(r"^(MSH|FHS|BHS)[|^]", stripped):
        return "hl7v2"
    # Also handle batch files that may have FHS/BHS before MSH
    if "\nMSH|" in stripped or "\rMSH|" in stripped:
        return "hl7v2"

    # --- JSON-based formats ---
    if stripped.startswith("{") or stripped.startswith("["):
        try:
            data = json.loads(stripped)
            if isinstance(data, dict):
                if "resourceType" in data:
                    return "fhir"
                # Could be a FHIR Bundle
                if data.get("resourceType") == "Bundle":
                    return "fhir"
            elif isinstance(data, list) and len(data) > 0:
                if isinstance(data[0], dict) and "resourceType" in data[0]:
                    return "fhir"
            return "csv"  # Generic JSON treated as structured data
        except (json.JSONDecodeError, ValueError):
            pass

    # --- XML-based formats ---
    if stripped.startswith("<") or stripped.startswith("<?xml"):
        lower = stripped[:2000].lower()
        # FHIR XML
        if "fhir" in lower or 'xmlns="http://hl7.org/fhir"' in stripped[:2000]:
            return "fhir"
        # CDA
        if "clinicaldocument" in lower or "urn:hl7-org:v3" in lower:
            return "cda"
        # HL7 v3 (non-CDA)
        if "urn:hl7-org:v3" in lower:
            return "cda"  # Treat as CDA/v3
        # Generic XML — treat as CSV/structured
        return "csv"

    # --- CSV/TSV: contains commas or tabs in a structured pattern ---
    lines = stripped.split("\n", 5)
    if len(lines) >= 2:
        # Check if it looks like delimited data
        for delim in [",", "\t", "|"]:
            counts = [line.count(delim) for line in lines[:5] if line.strip()]
            if counts and min(counts) >= 1 and max(counts) - min(counts) <= 2:
                return "csv"

    return None


def detect_x12_type(content):
    """For X12 content, detect the transaction set type from ST segment.

    Returns the ST01 value (e.g., '835', '837', '270', '271') or None.
    """
    stripped = content.lstrip("\ufeff").strip()
    if len(stripped) < 106:
        return None

    element_sep = stripped[3]
    segment_term = stripped[105]

    segments = stripped.split(segment_term)
    for seg in segments:
        seg = seg.strip().replace("\n", "").replace("\r", "")
        parts = seg.split(element_sep)
        if parts[0].upper() == "ST" and len(parts) > 1:
            return parts[1]
    return None
