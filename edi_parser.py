"""Base EDI X12 parser — handles delimiter detection and segment splitting."""

import re


class EDIFile:
    """Represents a parsed EDI X12 file with segments and metadata."""

    def __init__(self, raw_content):
        self.raw = raw_content.strip()
        self.element_sep = None
        self.sub_element_sep = None
        self.segment_term = None
        self.segments = []
        self._detect_delimiters()
        self._split_segments()

    def _detect_delimiters(self):
        # Strip BOM and leading whitespace
        content = self.raw.lstrip("\ufeff").lstrip()

        if not content.upper().startswith("ISA"):
            raise ValueError("File does not appear to be a valid EDI X12 file (missing ISA segment)")

        if len(content) < 106:
            raise ValueError("File is too short to contain a valid ISA segment")

        # ISA segment is always exactly 106 characters (positions 0-105):
        #   Position 3:   element separator
        #   Position 104: sub-element (component) separator
        #   Position 105: segment terminator
        self.element_sep = content[3]
        self.sub_element_sep = content[104]
        self.segment_term = content[105]

    def _split_segments(self):
        content = self.raw.lstrip("\ufeff").strip()
        # Split on segment terminator, then strip whitespace/newlines from each
        raw_segments = content.split(self.segment_term)
        for seg in raw_segments:
            seg = seg.strip().replace("\n", "").replace("\r", "")
            if seg:
                self.segments.append(seg)

    def get_elements(self, segment_str):
        """Split a segment string into its elements."""
        return segment_str.split(self.element_sep)

    def get_sub_elements(self, element_str):
        """Split a composite element into sub-elements."""
        return element_str.split(self.sub_element_sep)

    def get_transaction_type(self):
        """Return '835', '837', or None based on the ST segment."""
        for seg in self.segments:
            elements = self.get_elements(seg)
            if elements[0].upper() == "ST":
                code = elements[1] if len(elements) > 1 else ""
                if code == "835":
                    return "835"
                elif code == "837":
                    return "837"
        return None

    def get_transactions(self):
        """Yield lists of segments for each ST..SE transaction."""
        current = []
        inside = False
        for seg in self.segments:
            elements = self.get_elements(seg)
            seg_id = elements[0].upper()
            if seg_id == "ST":
                inside = True
                current = [seg]
            elif seg_id == "SE":
                if inside:
                    current.append(seg)
                    yield current
                inside = False
                current = []
            elif inside:
                current.append(seg)


def safe_float(val, default=0.0):
    """Convert string to float, returning default on failure."""
    try:
        return float(val) if val else default
    except (ValueError, TypeError):
        return default


def safe_int(val, default=0):
    """Convert string to int, returning default on failure."""
    try:
        return int(val) if val else default
    except (ValueError, TypeError):
        return default


def format_edi_date(date_str):
    """Convert CCYYMMDD or YYMMDD date string to MM/DD/YYYY."""
    if not date_str:
        return ""
    date_str = date_str.strip()
    if len(date_str) == 8:
        return f"{date_str[4:6]}/{date_str[6:8]}/{date_str[0:4]}"
    elif len(date_str) == 6:
        year = int(date_str[0:2])
        century = "20" if year < 50 else "19"
        return f"{date_str[2:4]}/{date_str[4:6]}/{century}{date_str[0:2]}"
    return date_str
