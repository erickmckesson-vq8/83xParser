"""Parser for FHIR (Fast Healthcare Interoperability Resources) — JSON and XML."""

import json
import re


def parse_fhir(content):
    """Parse FHIR content (JSON) and return sheet data.

    Handles both individual resources and Bundles.

    Returns:
        list of sheet dicts
    """
    data = _load_fhir(content)
    if data is None:
        return []

    # Collect all resources
    resources = []
    if isinstance(data, dict):
        if data.get("resourceType") == "Bundle":
            for entry in data.get("entry", []):
                res = entry.get("resource", entry)
                if isinstance(res, dict) and "resourceType" in res:
                    resources.append(res)
        elif "resourceType" in data:
            resources.append(data)
    elif isinstance(data, list):
        for item in data:
            if isinstance(item, dict) and "resourceType" in item:
                resources.append(item)

    if not resources:
        return []

    # Group by resource type
    by_type = {}
    for r in resources:
        rt = r.get("resourceType", "Unknown")
        by_type.setdefault(rt, []).append(r)

    sheets = []

    # --- Resource Summary sheet ---
    summary_headers = ["Resource Type", "Count"]
    summary_rows = [[rt, len(items)] for rt, items in sorted(by_type.items())]
    sheets.append({"name": "FHIR Summary", "headers": summary_headers,
                    "rows": summary_rows, "currency_cols": []})

    # --- Specific resource parsers ---
    parsers = {
        "Patient": _parse_patients,
        "Encounter": _parse_encounters,
        "Observation": _parse_observations,
        "Condition": _parse_conditions,
        "Procedure": _parse_procedures,
        "MedicationRequest": _parse_medication_requests,
        "MedicationDispense": _parse_medication_dispenses,
        "Claim": _parse_claims,
        "ExplanationOfBenefit": _parse_eobs,
        "Coverage": _parse_coverages,
        "DiagnosticReport": _parse_diagnostic_reports,
        "AllergyIntolerance": _parse_allergies,
        "Immunization": _parse_immunizations,
        "Practitioner": _parse_practitioners,
        "Organization": _parse_organizations,
    }

    for resource_type, items in by_type.items():
        parser = parsers.get(resource_type)
        if parser:
            sheet = parser(items)
            if sheet:
                sheets.append(sheet)
        else:
            # Generic fallback — show ID and key fields
            sheet = _parse_generic(resource_type, items)
            if sheet:
                sheets.append(sheet)

    return sheets


def _load_fhir(content):
    """Load FHIR content from JSON string or XML."""
    content = content.strip()

    # JSON
    if content.startswith("{") or content.startswith("["):
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            return None

    # XML — convert to dict (simplified)
    if content.startswith("<") or content.startswith("<?xml"):
        return _parse_fhir_xml(content)

    return None


def _parse_fhir_xml(xml_content):
    """Minimal FHIR XML parser — extracts resourceType and basic fields."""
    try:
        import xml.etree.ElementTree as ET
        # Strip FHIR namespace for easier parsing
        cleaned = re.sub(r'\s+xmlns[^"]*"[^"]*"', '', xml_content)
        cleaned = re.sub(r'\s+xmlns=["\'][^"\']*["\']', '', cleaned)
        root = ET.fromstring(cleaned)

        def elem_to_dict(elem):
            d = {}
            # Check for 'value' attribute (FHIR primitive pattern)
            if "value" in elem.attrib:
                return elem.attrib["value"]
            for child in elem:
                tag = child.tag.split("}")[-1]  # Remove namespace
                val = elem_to_dict(child)
                if tag in d:
                    if not isinstance(d[tag], list):
                        d[tag] = [d[tag]]
                    d[tag].append(val)
                else:
                    d[tag] = val
            if not d and elem.text and elem.text.strip():
                return elem.text.strip()
            if not d:
                return elem.attrib if elem.attrib else ""
            d.update({k: v for k, v in elem.attrib.items() if k != "value"})
            return d

        result = elem_to_dict(root)
        if isinstance(result, dict):
            result["resourceType"] = root.tag.split("}")[-1]
        return result
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Helper extractors
# ---------------------------------------------------------------------------

def _get_id(resource):
    return resource.get("id", "")


def _get_name(resource, path="name"):
    """Extract human name from FHIR HumanName."""
    names = resource.get(path, [])
    if isinstance(names, dict):
        names = [names]
    if not names:
        return ""
    name = names[0] if isinstance(names, list) else names
    if isinstance(name, str):
        return name
    text = name.get("text")
    if text:
        return text
    family = name.get("family", "")
    given = name.get("given", [])
    if isinstance(given, list):
        given = " ".join(given)
    return f"{family}, {given}".strip(", ")


def _get_coding(codeable_concept, join=True):
    """Extract code and display from a CodeableConcept."""
    if not codeable_concept:
        return ("", "") if not join else ""
    if isinstance(codeable_concept, str):
        return ("", codeable_concept) if not join else codeable_concept

    text = codeable_concept.get("text", "")
    codings = codeable_concept.get("coding", [])
    if isinstance(codings, dict):
        codings = [codings]
    if codings:
        code = codings[0].get("code", "")
        display = codings[0].get("display", text)
        if join:
            return f"{code} — {display}" if code and display else (display or code or text)
        return (code, display or text)
    if join:
        return text
    return ("", text)


def _get_reference_display(ref):
    """Extract display or reference string."""
    if not ref:
        return ""
    if isinstance(ref, str):
        return ref
    return ref.get("display", "") or ref.get("reference", "")


def _get_period(period):
    """Extract start and end from a Period."""
    if not period:
        return "", ""
    return period.get("start", ""), period.get("end", "")


def _get_address(resource, path="address"):
    """Extract address text."""
    addrs = resource.get(path, [])
    if isinstance(addrs, dict):
        addrs = [addrs]
    if not addrs:
        return ""
    addr = addrs[0]
    if isinstance(addr, str):
        return addr
    text = addr.get("text")
    if text:
        return text
    lines = addr.get("line", [])
    if isinstance(lines, str):
        lines = [lines]
    parts = lines + [addr.get("city", ""), addr.get("state", ""), addr.get("postalCode", "")]
    return ", ".join(p for p in parts if p)


def _get_telecom(resource, system_filter=None):
    """Extract telecom value."""
    telecoms = resource.get("telecom", [])
    if isinstance(telecoms, dict):
        telecoms = [telecoms]
    for t in telecoms:
        if isinstance(t, str):
            return t
        if system_filter and t.get("system") != system_filter:
            continue
        return t.get("value", "")
    return ""


# ---------------------------------------------------------------------------
# Resource-specific parsers
# ---------------------------------------------------------------------------

def _parse_patients(items):
    headers = ["ID", "Name", "Date of Birth", "Gender", "Address",
               "Phone", "Email", "Marital Status", "MRN", "SSN"]
    rows = []
    for r in items:
        # Extract identifiers
        mrn = ssn = ""
        for ident in _ensure_list(r.get("identifier", [])):
            type_coding = _get_coding(ident.get("type"), join=False)
            code = type_coding[0] if isinstance(type_coding, tuple) else ""
            val = ident.get("value", "")
            if code in ("MR", "MRN") or "medical" in str(ident.get("type", "")).lower():
                mrn = val
            elif code == "SS" or "social" in str(ident.get("type", "")).lower():
                ssn = val
            elif not mrn:
                mrn = val

        rows.append([
            _get_id(r), _get_name(r), r.get("birthDate", ""),
            r.get("gender", ""), _get_address(r),
            _get_telecom(r, "phone"), _get_telecom(r, "email"),
            _get_coding(r.get("maritalStatus")), mrn, ssn,
        ])
    return {"name": "FHIR Patients", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_encounters(items):
    headers = ["ID", "Status", "Class", "Type", "Subject",
               "Start", "End", "Reason", "Participant", "Location",
               "Service Provider", "Diagnosis Codes"]
    rows = []
    for r in items:
        enc_class = r.get("class", {})
        if isinstance(enc_class, dict):
            enc_class = enc_class.get("display", "") or enc_class.get("code", "")

        types = _ensure_list(r.get("type", []))
        type_str = "; ".join(_get_coding(t) for t in types)

        start, end = _get_period(r.get("period"))

        reasons = _ensure_list(r.get("reasonCode", []))
        reason_str = "; ".join(_get_coding(rc) for rc in reasons)

        participants = _ensure_list(r.get("participant", []))
        part_str = "; ".join(_get_reference_display(p.get("individual")) for p in participants)

        locations = _ensure_list(r.get("location", []))
        loc_str = "; ".join(_get_reference_display(l.get("location")) for l in locations)

        dx_list = _ensure_list(r.get("diagnosis", []))
        dx_str = "; ".join(_get_coding(d.get("condition") if isinstance(d.get("condition"), dict)
                           and "coding" in d.get("condition", {}) else d.get("condition"))
                           for d in dx_list)
        if not dx_str:
            dx_str = "; ".join(_get_reference_display(d.get("condition")) for d in dx_list)

        rows.append([
            _get_id(r), r.get("status", ""), enc_class, type_str,
            _get_reference_display(r.get("subject")),
            start, end, reason_str, part_str, loc_str,
            _get_reference_display(r.get("serviceProvider")), dx_str,
        ])
    return {"name": "FHIR Encounters", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_observations(items):
    headers = ["ID", "Status", "Category", "Code", "Display",
               "Value", "Units", "Reference Range",
               "Subject", "Effective Date", "Issued"]
    rows = []
    for r in items:
        categories = _ensure_list(r.get("category", []))
        cat_str = "; ".join(_get_coding(c) for c in categories)

        code, display = _get_coding(r.get("code"), join=False)

        # Value can be in many forms
        value = ""
        units = ""
        vq = r.get("valueQuantity")
        if vq:
            value = vq.get("value", "")
            units = vq.get("unit", "") or vq.get("code", "")
        elif "valueString" in r:
            value = r["valueString"]
        elif "valueCodeableConcept" in r:
            value = _get_coding(r["valueCodeableConcept"])
        elif "valueBoolean" in r:
            value = str(r["valueBoolean"])
        elif "valueInteger" in r:
            value = r["valueInteger"]

        # Reference range
        ref_ranges = _ensure_list(r.get("referenceRange", []))
        range_str = ""
        if ref_ranges:
            rr = ref_ranges[0]
            low = rr.get("low", {}).get("value", "")
            high = rr.get("high", {}).get("value", "")
            if low and high:
                range_str = f"{low} - {high}"
            elif rr.get("text"):
                range_str = rr["text"]

        effective = r.get("effectiveDateTime", "")
        if not effective:
            eff_period = r.get("effectivePeriod", {})
            effective = eff_period.get("start", "") if eff_period else ""

        rows.append([
            _get_id(r), r.get("status", ""), cat_str, code, display,
            value, units, range_str,
            _get_reference_display(r.get("subject")),
            effective, r.get("issued", ""),
        ])
    return {"name": "FHIR Observations", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_conditions(items):
    headers = ["ID", "Clinical Status", "Verification Status",
               "Code", "Display", "Subject",
               "Onset", "Abatement", "Recorded Date", "Severity"]
    rows = []
    for r in items:
        code, display = _get_coding(r.get("code"), join=False)
        onset = r.get("onsetDateTime", "") or r.get("onsetString", "")
        abate = r.get("abatementDateTime", "") or r.get("abatementString", "")
        rows.append([
            _get_id(r),
            _get_coding(r.get("clinicalStatus")),
            _get_coding(r.get("verificationStatus")),
            code, display,
            _get_reference_display(r.get("subject")),
            onset, abate, r.get("recordedDate", ""),
            _get_coding(r.get("severity")),
        ])
    return {"name": "FHIR Conditions", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_procedures(items):
    headers = ["ID", "Status", "Code", "Display", "Subject",
               "Performed Start", "Performed End", "Performer", "Location", "Reason"]
    rows = []
    for r in items:
        code, display = _get_coding(r.get("code"), join=False)
        perf = r.get("performedDateTime", "")
        perf_start = perf_end = ""
        if not perf:
            pp = r.get("performedPeriod", {})
            perf_start, perf_end = _get_period(pp)
        else:
            perf_start = perf

        performers = _ensure_list(r.get("performer", []))
        perf_str = "; ".join(_get_reference_display(p.get("actor")) for p in performers)

        reasons = _ensure_list(r.get("reasonCode", []))
        reason_str = "; ".join(_get_coding(rc) for rc in reasons)

        rows.append([
            _get_id(r), r.get("status", ""), code, display,
            _get_reference_display(r.get("subject")),
            perf_start, perf_end, perf_str,
            _get_reference_display(r.get("location")), reason_str,
        ])
    return {"name": "FHIR Procedures", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_medication_requests(items):
    headers = ["ID", "Status", "Intent", "Medication", "Subject",
               "Requester", "Dosage Instructions", "Authored On",
               "Reason", "Dispense Quantity"]
    rows = []
    for r in items:
        med = _get_coding(r.get("medicationCodeableConcept")) or \
              _get_reference_display(r.get("medicationReference"))

        dosages = _ensure_list(r.get("dosageInstruction", []))
        dosage_str = "; ".join(d.get("text", "") for d in dosages if d.get("text"))

        reasons = _ensure_list(r.get("reasonCode", []))
        reason_str = "; ".join(_get_coding(rc) for rc in reasons)

        disp_req = r.get("dispenseRequest", {})
        disp_qty = ""
        if disp_req:
            qty = disp_req.get("quantity", {})
            if qty:
                disp_qty = f"{qty.get('value', '')} {qty.get('unit', '')}".strip()

        rows.append([
            _get_id(r), r.get("status", ""), r.get("intent", ""),
            med, _get_reference_display(r.get("subject")),
            _get_reference_display(r.get("requester")),
            dosage_str, r.get("authoredOn", ""),
            reason_str, disp_qty,
        ])
    return {"name": "FHIR Medications", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_medication_dispenses(items):
    headers = ["ID", "Status", "Medication", "Subject",
               "Performer", "Quantity", "Days Supply", "When Handed Over"]
    rows = []
    for r in items:
        med = _get_coding(r.get("medicationCodeableConcept")) or \
              _get_reference_display(r.get("medicationReference"))
        qty = r.get("quantity", {})
        qty_str = f"{qty.get('value', '')} {qty.get('unit', '')}".strip() if qty else ""
        ds = r.get("daysSupply", {})
        ds_str = f"{ds.get('value', '')} {ds.get('unit', 'days')}".strip() if ds else ""

        performers = _ensure_list(r.get("performer", []))
        perf_str = "; ".join(_get_reference_display(p.get("actor")) for p in performers)

        rows.append([
            _get_id(r), r.get("status", ""), med,
            _get_reference_display(r.get("subject")),
            perf_str, qty_str, ds_str,
            r.get("whenHandedOver", ""),
        ])
    return {"name": "FHIR Dispenses", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_claims(items):
    headers = ["ID", "Status", "Type", "Use", "Patient",
               "Provider", "Priority", "Total",
               "Diagnosis Codes", "Item Count"]
    rows = []
    for r in items:
        total = r.get("total", {})
        total_str = total.get("value", "") if isinstance(total, dict) else total

        dx_list = _ensure_list(r.get("diagnosis", []))
        dx_str = "; ".join(_get_coding(d.get("diagnosisCodeableConcept"))
                           for d in dx_list if d.get("diagnosisCodeableConcept"))

        items_list = _ensure_list(r.get("item", []))

        rows.append([
            _get_id(r), r.get("status", ""), _get_coding(r.get("type")),
            r.get("use", ""), _get_reference_display(r.get("patient")),
            _get_reference_display(r.get("provider")),
            _get_coding(r.get("priority")), total_str,
            dx_str, len(items_list),
        ])
    return {"name": "FHIR Claims", "headers": headers, "rows": rows, "currency_cols": [8]}


def _parse_eobs(items):
    headers = ["ID", "Status", "Type", "Use", "Patient",
               "Provider", "Outcome",
               "Total (Submitted)", "Total (Benefit)",
               "Payment Amount", "Payment Date",
               "Diagnosis Codes", "Item Count"]
    rows = []
    for r in items:
        totals = _ensure_list(r.get("total", []))
        submitted = benefit = ""
        for t in totals:
            cat = _get_coding(t.get("category"), join=False)
            cat_code = cat[0] if isinstance(cat, tuple) else ""
            amount = t.get("amount", {})
            val = amount.get("value", "") if isinstance(amount, dict) else ""
            if "submitted" in cat_code.lower() or "submitted" in str(cat).lower():
                submitted = val
            elif "benefit" in cat_code.lower() or "benefit" in str(cat).lower():
                benefit = val
            elif not submitted:
                submitted = val

        payment = r.get("payment", {})
        pay_amount = ""
        pay_date = ""
        if payment:
            pa = payment.get("amount", {})
            pay_amount = pa.get("value", "") if isinstance(pa, dict) else ""
            pay_date = payment.get("date", "")

        dx_list = _ensure_list(r.get("diagnosis", []))
        dx_str = "; ".join(_get_coding(d.get("diagnosisCodeableConcept"))
                           for d in dx_list if d.get("diagnosisCodeableConcept"))

        items_list = _ensure_list(r.get("item", []))

        rows.append([
            _get_id(r), r.get("status", ""), _get_coding(r.get("type")),
            r.get("use", ""), _get_reference_display(r.get("patient")),
            _get_reference_display(r.get("provider")),
            r.get("outcome", ""),
            submitted, benefit, pay_amount, pay_date,
            dx_str, len(items_list),
        ])
    return {"name": "FHIR EOBs", "headers": headers, "rows": rows,
            "currency_cols": [8, 9, 10]}


def _parse_coverages(items):
    headers = ["ID", "Status", "Type", "Subscriber", "Beneficiary",
               "Payor", "Group ID", "Group Name", "Plan",
               "Period Start", "Period End"]
    rows = []
    for r in items:
        class_list = _ensure_list(r.get("class", []))
        group_id = group_name = plan = ""
        for cls in class_list:
            cls_type = _get_coding(cls.get("type"), join=False)
            cls_code = cls_type[0] if isinstance(cls_type, tuple) else ""
            if cls_code == "group" or "group" in str(cls_type).lower():
                group_id = cls.get("value", "")
                group_name = cls.get("name", "")
            elif cls_code == "plan" or "plan" in str(cls_type).lower():
                plan = cls.get("value", "") or cls.get("name", "")

        payors = _ensure_list(r.get("payor", []))
        payor_str = "; ".join(_get_reference_display(p) for p in payors)

        start, end = _get_period(r.get("period"))

        rows.append([
            _get_id(r), r.get("status", ""), _get_coding(r.get("type")),
            _get_reference_display(r.get("subscriber")),
            _get_reference_display(r.get("beneficiary")),
            payor_str, group_id, group_name, plan,
            start, end,
        ])
    return {"name": "FHIR Coverage", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_diagnostic_reports(items):
    headers = ["ID", "Status", "Category", "Code", "Display",
               "Subject", "Effective Date", "Issued",
               "Performer", "Result Count", "Conclusion"]
    rows = []
    for r in items:
        categories = _ensure_list(r.get("category", []))
        cat_str = "; ".join(_get_coding(c) for c in categories)
        code, display = _get_coding(r.get("code"), join=False)
        effective = r.get("effectiveDateTime", "")
        if not effective:
            ep = r.get("effectivePeriod", {})
            effective = ep.get("start", "") if ep else ""

        performers = _ensure_list(r.get("performer", []))
        perf_str = "; ".join(_get_reference_display(p) for p in performers)

        results = _ensure_list(r.get("result", []))

        rows.append([
            _get_id(r), r.get("status", ""), cat_str, code, display,
            _get_reference_display(r.get("subject")),
            effective, r.get("issued", ""),
            perf_str, len(results), r.get("conclusion", ""),
        ])
    return {"name": "FHIR Diagnostic Reports", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_allergies(items):
    headers = ["ID", "Clinical Status", "Verification Status", "Type",
               "Category", "Criticality", "Code", "Display",
               "Patient", "Onset", "Recorded Date", "Reactions"]
    rows = []
    for r in items:
        code, display = _get_coding(r.get("code"), join=False)
        categories = _ensure_list(r.get("category", []))
        cat_str = ", ".join(categories) if categories else ""

        reactions = _ensure_list(r.get("reaction", []))
        reaction_strs = []
        for rx in reactions:
            manifestations = _ensure_list(rx.get("manifestation", []))
            m_str = "; ".join(_get_coding(m) for m in manifestations)
            severity = rx.get("severity", "")
            if m_str:
                reaction_strs.append(f"{m_str} ({severity})" if severity else m_str)
        reaction_str = "; ".join(reaction_strs)

        rows.append([
            _get_id(r),
            _get_coding(r.get("clinicalStatus")),
            _get_coding(r.get("verificationStatus")),
            r.get("type", ""), cat_str, r.get("criticality", ""),
            code, display,
            _get_reference_display(r.get("patient")),
            r.get("onsetDateTime", ""), r.get("recordedDate", ""),
            reaction_str,
        ])
    return {"name": "FHIR Allergies", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_immunizations(items):
    headers = ["ID", "Status", "Vaccine Code", "Vaccine Name",
               "Patient", "Occurrence Date", "Lot Number",
               "Site", "Route", "Dose Quantity", "Performer"]
    rows = []
    for r in items:
        code, display = _get_coding(r.get("vaccineCode"), join=False)
        occurrence = r.get("occurrenceDateTime", "") or r.get("occurrenceString", "")

        dose = r.get("doseQuantity", {})
        dose_str = f"{dose.get('value', '')} {dose.get('unit', '')}".strip() if dose else ""

        performers = _ensure_list(r.get("performer", []))
        perf_str = "; ".join(_get_reference_display(p.get("actor")) for p in performers)

        rows.append([
            _get_id(r), r.get("status", ""), code, display,
            _get_reference_display(r.get("patient")),
            occurrence, r.get("lotNumber", ""),
            _get_coding(r.get("site")), _get_coding(r.get("route")),
            dose_str, perf_str,
        ])
    return {"name": "FHIR Immunizations", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_practitioners(items):
    headers = ["ID", "Name", "Gender", "Birth Date",
               "Identifier", "Qualification", "Phone", "Email", "Address"]
    rows = []
    for r in items:
        idents = _ensure_list(r.get("identifier", []))
        ident_str = "; ".join(f"{i.get('value', '')} ({_get_coding(i.get('type'))})"
                              for i in idents if i.get("value"))

        quals = _ensure_list(r.get("qualification", []))
        qual_str = "; ".join(_get_coding(q.get("code")) for q in quals)

        rows.append([
            _get_id(r), _get_name(r), r.get("gender", ""), r.get("birthDate", ""),
            ident_str, qual_str,
            _get_telecom(r, "phone"), _get_telecom(r, "email"),
            _get_address(r),
        ])
    return {"name": "FHIR Practitioners", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_organizations(items):
    headers = ["ID", "Name", "Type", "Identifier", "Phone", "Email",
               "Address", "Active"]
    rows = []
    for r in items:
        types = _ensure_list(r.get("type", []))
        type_str = "; ".join(_get_coding(t) for t in types)

        idents = _ensure_list(r.get("identifier", []))
        ident_str = "; ".join(f"{i.get('value', '')} ({_get_coding(i.get('type'))})"
                              for i in idents if i.get("value"))

        rows.append([
            _get_id(r), r.get("name", ""), type_str, ident_str,
            _get_telecom(r, "phone"), _get_telecom(r, "email"),
            _get_address(r), r.get("active", ""),
        ])
    return {"name": "FHIR Organizations", "headers": headers, "rows": rows, "currency_cols": []}


def _parse_generic(resource_type, items):
    """Fallback parser for unknown FHIR resource types."""
    headers = ["ID", "Resource Type", "Status", "Subject", "Date", "Key Data"]
    rows = []
    for r in items:
        status = r.get("status", "")
        subject = _get_reference_display(r.get("subject") or r.get("patient"))
        date = (r.get("date") or r.get("created") or r.get("issued") or
                r.get("effectiveDateTime") or r.get("authoredOn") or "")
        # Collect a few meaningful top-level string fields
        key_data = []
        for k, v in r.items():
            if k in ("resourceType", "id", "status", "meta", "text"):
                continue
            if isinstance(v, str) and v and len(v) < 100:
                key_data.append(f"{k}: {v}")
            if len(key_data) >= 5:
                break
        rows.append([
            _get_id(r), resource_type, status, subject, date,
            "; ".join(key_data),
        ])
    sheet_name = f"FHIR {resource_type}"
    if len(sheet_name) > 31:
        sheet_name = sheet_name[:31]
    return {"name": sheet_name, "headers": headers, "rows": rows, "currency_cols": []}


def _ensure_list(val):
    """Ensure value is a list."""
    if isinstance(val, list):
        return val
    if val is None:
        return []
    return [val]
