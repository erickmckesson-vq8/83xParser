"""Parser for HL7 v3 / CDA (Clinical Document Architecture) XML documents."""

import re
import xml.etree.ElementTree as ET


# CDA namespace
NS = {"cda": "urn:hl7-org:v3", "sdtc": "urn:hl7-org:sdtc"}


def parse_cda(content):
    """Parse a CDA XML document and return sheet data.

    Returns:
        list of sheet dicts
    """
    try:
        root = ET.fromstring(content.encode("utf-8") if isinstance(content, str) else content)
    except ET.ParseError:
        # Try stripping namespace-heavy content
        try:
            cleaned = re.sub(r'xmlns[^"]*="[^"]*"', '', content)
            root = ET.fromstring(cleaned.encode("utf-8") if isinstance(cleaned, str) else cleaned)
        except ET.ParseError:
            return []

    # Detect namespace
    ns = ""
    tag = root.tag
    if "}" in tag:
        ns = tag.split("}")[0] + "}"

    sheets = []

    # --- Document Info ---
    doc_info = _parse_document_info(root, ns)
    if doc_info:
        sheets.append(doc_info)

    # --- Patient Demographics ---
    patient_info = _parse_patient(root, ns)
    if patient_info:
        sheets.append(patient_info)

    # --- Authors ---
    author_info = _parse_authors(root, ns)
    if author_info:
        sheets.append(author_info)

    # --- Document Sections ---
    sections = _parse_sections(root, ns)
    if sections:
        sheets.append(sections)

    # --- Structured Entries (if present) ---
    entries = _parse_entries(root, ns)
    if entries:
        sheets.extend(entries)

    return sheets


def _find(elem, path, ns):
    """Find element using namespace-aware path."""
    if ns:
        # Convert simple path to namespace-qualified
        parts = path.split("/")
        ns_path = "/".join(f"{ns}{p}" if p and not p.startswith("{") and not p.startswith(".")
                           else p for p in parts)
        result = elem.find(ns_path)
        if result is not None:
            return result
    return elem.find(path)


def _findall(elem, path, ns):
    """Find all elements using namespace-aware path."""
    if ns:
        parts = path.split("/")
        ns_path = "/".join(f"{ns}{p}" if p and not p.startswith("{") and not p.startswith(".")
                           else p for p in parts)
        result = elem.findall(ns_path)
        if result:
            return result
    return elem.findall(path)


def _get_text(elem):
    """Get all text content from an element and its children."""
    if elem is None:
        return ""
    texts = []
    if elem.text and elem.text.strip():
        texts.append(elem.text.strip())
    for child in elem:
        child_text = _get_text(child)
        if child_text:
            texts.append(child_text)
        if child.tail and child.tail.strip():
            texts.append(child.tail.strip())
    return " ".join(texts)


def _get_attr(elem, attr, default=""):
    """Get attribute value from element."""
    if elem is None:
        return default
    return elem.get(attr, default)


def _parse_name(name_elem, ns):
    """Parse a CDA name element (given/family)."""
    if name_elem is None:
        return ""
    given_elems = _findall(name_elem, "given", ns)
    family_elem = _find(name_elem, "family", ns)

    given = " ".join(_get_text(g) for g in given_elems) if given_elems else ""
    family = _get_text(family_elem) if family_elem is not None else ""

    if family and given:
        return f"{family}, {given}"
    return family or given or _get_text(name_elem)


def _parse_addr(addr_elem, ns):
    """Parse a CDA addr element."""
    if addr_elem is None:
        return ""
    parts = []
    for tag in ["streetAddressLine", "city", "state", "postalCode", "country"]:
        elem = _find(addr_elem, tag, ns)
        if elem is not None:
            text = _get_text(elem)
            if text:
                parts.append(text)
    if parts:
        return ", ".join(parts)
    return _get_text(addr_elem)


def _parse_telecom(telecom_elems):
    """Parse telecom elements into a string."""
    if not telecom_elems:
        return ""
    values = []
    for t in telecom_elems:
        val = t.get("value", "")
        use = t.get("use", "")
        if val:
            val = val.replace("tel:", "").replace("mailto:", "")
            values.append(f"{val} ({use})" if use else val)
    return "; ".join(values)


def _parse_document_info(root, ns):
    """Extract document-level metadata."""
    headers = ["Field", "Value"]
    rows = []

    # Title
    title = _find(root, "title", ns)
    if title is not None:
        rows.append(["Document Title", _get_text(title)])

    # Type/Code
    code = _find(root, "code", ns)
    if code is not None:
        code_val = code.get("code", "")
        display = code.get("displayName", "")
        rows.append(["Document Type", f"{code_val} — {display}" if display else code_val])

    # Effective time
    eff = _find(root, "effectiveTime", ns)
    if eff is not None:
        rows.append(["Document Date", eff.get("value", "")])

    # Confidentiality
    conf = _find(root, "confidentialityCode", ns)
    if conf is not None:
        rows.append(["Confidentiality", conf.get("displayName", conf.get("code", ""))])

    # Language
    lang = _find(root, "languageCode", ns)
    if lang is not None:
        rows.append(["Language", lang.get("code", "")])

    # Document ID
    doc_id = _find(root, "id", ns)
    if doc_id is not None:
        rows.append(["Document ID", doc_id.get("extension", doc_id.get("root", ""))])

    # Custodian
    custodian = _find(root, "custodian", ns)
    if custodian is not None:
        org_name = _find(custodian, ".//name", ns)
        if org_name is not None:
            rows.append(["Custodian", _get_text(org_name)])

    if not rows:
        return None
    return {"name": "CDA Document Info", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_patient(root, ns):
    """Extract patient demographics from recordTarget."""
    record_targets = _findall(root, "recordTarget", ns)
    if not record_targets:
        return None

    headers = ["Patient Name", "Date of Birth", "Gender",
               "Address", "Phone/Email", "MRN", "SSN", "Race", "Ethnicity"]
    rows = []

    for rt in record_targets:
        patient_role = _find(rt, "patientRole", ns)
        if patient_role is None:
            continue

        patient = _find(patient_role, "patient", ns)

        # IDs
        ids = _findall(patient_role, "id", ns)
        mrn = ssn = ""
        for id_elem in ids:
            ext = id_elem.get("extension", "")
            root_oid = id_elem.get("root", "")
            if "2.16.840.1.113883.4.1" in root_oid:  # SSN OID
                ssn = ext
            elif ext:
                mrn = ext

        # Name
        name = ""
        if patient is not None:
            name_elem = _find(patient, "name", ns)
            name = _parse_name(name_elem, ns)

        # DOB
        dob = ""
        if patient is not None:
            birth = _find(patient, "birthTime", ns)
            if birth is not None:
                dob = birth.get("value", "")

        # Gender
        gender = ""
        if patient is not None:
            gender_elem = _find(patient, "administrativeGenderCode", ns)
            if gender_elem is not None:
                gender = gender_elem.get("displayName", gender_elem.get("code", ""))

        # Address
        addr_elem = _find(patient_role, "addr", ns)
        address = _parse_addr(addr_elem, ns)

        # Telecom
        telecom_elems = _findall(patient_role, "telecom", ns)
        telecom = _parse_telecom(telecom_elems)

        # Race/Ethnicity
        race = ethnicity = ""
        if patient is not None:
            race_elem = _find(patient, "raceCode", ns)
            if race_elem is not None:
                race = race_elem.get("displayName", race_elem.get("code", ""))
            eth_elem = _find(patient, "ethnicGroupCode", ns)
            if eth_elem is not None:
                ethnicity = eth_elem.get("displayName", eth_elem.get("code", ""))

        rows.append([name, dob, gender, address, telecom, mrn, ssn, race, ethnicity])

    if not rows:
        return None
    return {"name": "CDA Patient", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_authors(root, ns):
    """Extract author information."""
    authors = _findall(root, "author", ns)
    if not authors:
        return None

    headers = ["Author Name", "Author Time", "Organization"]
    rows = []

    for author in authors:
        time_elem = _find(author, "time", ns)
        time_val = time_elem.get("value", "") if time_elem is not None else ""

        assigned = _find(author, "assignedAuthor", ns)
        name = ""
        org = ""
        if assigned is not None:
            person = _find(assigned, "assignedPerson", ns)
            if person is not None:
                name_elem = _find(person, "name", ns)
                name = _parse_name(name_elem, ns)

            org_elem = _find(assigned, "representedOrganization", ns)
            if org_elem is not None:
                org_name = _find(org_elem, "name", ns)
                org = _get_text(org_name) if org_name is not None else ""

        rows.append([name, time_val, org])

    if not rows:
        return None
    return {"name": "CDA Authors", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_sections(root, ns):
    """Extract document sections (title + narrative text)."""
    # Find structuredBody
    body = _find(root, ".//structuredBody", ns)
    if body is None:
        body = root  # Try root directly

    components = _findall(body, "component", ns)
    if not components:
        # Try deeper nesting
        components = _findall(root, ".//component/structuredBody/component", ns)

    headers = ["Section Title", "Section Code", "Narrative Text (Preview)"]
    rows = []

    for comp in components:
        section = _find(comp, "section", ns)
        if section is None:
            continue

        title_elem = _find(section, "title", ns)
        title = _get_text(title_elem) if title_elem is not None else ""

        code_elem = _find(section, "code", ns)
        code_val = ""
        if code_elem is not None:
            c = code_elem.get("code", "")
            d = code_elem.get("displayName", "")
            code_val = f"{c} — {d}" if d else c

        text_elem = _find(section, "text", ns)
        narrative = ""
        if text_elem is not None:
            narrative = _get_text(text_elem)
            # Truncate long narratives for the preview
            if len(narrative) > 500:
                narrative = narrative[:500] + "..."

        if title or code_val:
            rows.append([title, code_val, narrative])

    if not rows:
        return None
    return {"name": "CDA Sections", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_entries(root, ns):
    """Extract structured entries from sections (medications, problems, etc.)."""
    sheets = []

    body = _find(root, ".//structuredBody", ns)
    if body is None:
        return sheets

    components = _findall(body, "component", ns)
    if not components:
        components = _findall(root, ".//component/structuredBody/component", ns)

    for comp in components:
        section = _find(comp, "section", ns)
        if section is None:
            continue

        title_elem = _find(section, "title", ns)
        section_title = _get_text(title_elem) if title_elem is not None else "Unknown Section"

        entries = _findall(section, "entry", ns)
        if not entries:
            continue

        # Extract entry data generically
        entry_rows = []
        for entry in entries:
            row_data = _extract_entry_data(entry, ns)
            if row_data:
                entry_rows.append(row_data)

        if entry_rows:
            # Use consistent columns across all entries
            all_keys = []
            for row in entry_rows:
                for k in row:
                    if k not in all_keys:
                        all_keys.append(k)

            headers = all_keys
            rows = [[row.get(k, "") for k in all_keys] for row in entry_rows]

            sheet_name = f"CDA {section_title}"
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]
            sheets.append({"name": sheet_name, "headers": headers,
                           "rows": rows, "currency_cols": []})

    return sheets


def _extract_entry_data(entry, ns):
    """Extract key data from a CDA entry element."""
    data = {}

    # Look for common clinical act patterns
    for act_type in ["act", "observation", "substanceAdministration",
                     "procedure", "encounter", "organizer", "supply"]:
        act = _find(entry, act_type, ns)
        if act is not None:
            # Code
            code_elem = _find(act, "code", ns)
            if code_elem is not None:
                data["Code"] = code_elem.get("code", "")
                data["Display Name"] = code_elem.get("displayName", "")
                data["Code System"] = code_elem.get("codeSystemName", "")

            # Status
            status = _find(act, "statusCode", ns)
            if status is not None:
                data["Status"] = status.get("code", "")

            # Effective time
            eff_time = _find(act, "effectiveTime", ns)
            if eff_time is not None:
                val = eff_time.get("value", "")
                low = _find(eff_time, "low", ns)
                high = _find(eff_time, "high", ns)
                if val:
                    data["Date"] = val
                if low is not None:
                    data["Start Date"] = low.get("value", "")
                if high is not None:
                    data["End Date"] = high.get("value", "")

            # Value (for observations)
            value_elem = _find(act, "value", ns)
            if value_elem is not None:
                val = value_elem.get("value", "")
                unit = value_elem.get("unit", "")
                display = value_elem.get("displayName", "")
                code_val = value_elem.get("code", "")
                if val and unit:
                    data["Value"] = f"{val} {unit}"
                elif display:
                    data["Value"] = display
                elif code_val:
                    data["Value"] = code_val
                elif val:
                    data["Value"] = val

            # Text
            text_elem = _find(act, "text", ns)
            if text_elem is not None:
                text = _get_text(text_elem)
                if text and len(text) < 200:
                    data["Text"] = text

            break  # Only process first matching act type

    return data if data else None
