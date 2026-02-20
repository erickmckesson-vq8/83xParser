"""Generic parser for X12 transaction types other than 835/837."""

from edi_parser import EDIFile, safe_float, format_edi_date

# X12 transaction set names
TRANSACTION_NAMES = {
    "270": "Eligibility Inquiry",
    "271": "Eligibility Response",
    "276": "Claim Status Request",
    "277": "Claim Status Response",
    "278": "Health Care Services Review (Prior Auth)",
    "834": "Benefit Enrollment and Maintenance",
    "820": "Premium Payment",
    "835": "Remittance Advice",
    "837": "Health Care Claim",
    "999": "Implementation Acknowledgment",
    "997": "Functional Acknowledgment",
    "275": "Additional Information to Support a Health Care Claim",
    "274": "Health Care Provider Information",
}

# NM1 entity codes
ENTITY_CODES = {
    "03": "Dependent",
    "1P": "Provider",
    "2B": "Third-Party Administrator",
    "36": "Employer",
    "40": "Receiver",
    "41": "Submitter",
    "82": "Rendering Provider",
    "85": "Billing Provider",
    "87": "Pay-to Provider",
    "DQ": "Supervising Provider",
    "FA": "Facility",
    "IL": "Insured/Subscriber",
    "P3": "Primary Care Provider",
    "P4": "Prior Insurance Carrier",
    "P5": "Plan Sponsor",
    "PE": "Payee",
    "PR": "Payer",
    "QC": "Patient",
}

# Common EB (Eligibility/Benefit) codes
EB_INFO_CODES = {
    "1": "Active Coverage",
    "2": "Active - Full Risk Capitation",
    "3": "Active - Services Capitated",
    "4": "Active - Services Capitated to Primary Care Provider",
    "5": "Active - Pending Investigation",
    "6": "Inactive",
    "7": "Inactive - Pending Eligibility Update",
    "8": "Inactive - Pending Investigation",
    "A": "Co-Insurance",
    "B": "Co-Payment",
    "C": "Deductible",
    "CB": "Coverage Basis",
    "D": "Benefit Description",
    "E": "Exclusions",
    "F": "Limitations",
    "G": "Out of Pocket (Stop Loss)",
    "H": "Unlimited",
    "I": "Non-Covered",
    "J": "Cost Containment",
    "K": "Reserve",
    "L": "Primary Care Provider",
    "M": "Pre-existing Condition",
    "MC": "Managed Care Coordinator",
    "N": "Services Restricted to Following Provider",
    "O": "Not Deemed a Medical Necessity",
    "P": "Benefit Disclaimer",
    "Q": "Second Surgical Opinion Required",
    "R": "Other or Additional Payor",
    "S": "Prior Year(s) History",
    "T": "Card(s) Reported Lost/Stolen",
    "U": "Contact Following Entity for Information",
    "V": "Cannot Process",
    "W": "Other Source of Data",
    "X": "Health Care Facility",
    "Y": "Spend Down",
}

# Service type codes (common ones)
SERVICE_TYPE_CODES = {
    "1": "Medical Care",
    "2": "Surgical",
    "3": "Consultation",
    "4": "Diagnostic X-Ray",
    "5": "Diagnostic Lab",
    "6": "Radiation Therapy",
    "7": "Anesthesia",
    "8": "Surgical Assistance",
    "12": "Durable Medical Equipment Purchase",
    "14": "Renal Supplies in the Home",
    "18": "Durable Medical Equipment Rental",
    "23": "Diagnostic Dental",
    "24": "Periodontics",
    "25": "Restorative",
    "26": "Endodontics",
    "27": "Dental Crowns",
    "28": "Dental Accident",
    "30": "Health Benefit Plan Coverage",
    "32": "Plan Waiting Period",
    "33": "Chiropractic",
    "34": "Chiropractic Office Visits",
    "35": "Dental Care",
    "36": "Dental Crowns",
    "37": "Dental Accident",
    "38": "Orthodontics",
    "39": "Prosthodontics",
    "40": "Oral Surgery",
    "41": "Routine (Preventive) Dental",
    "42": "Home Health Care",
    "43": "Home Health Prescriptions",
    "44": "Home Health Visits",
    "45": "Hospice",
    "46": "Respite Care",
    "47": "Hospital",
    "48": "Hospital - Inpatient",
    "50": "Hospital - Outpatient",
    "51": "Hospital - Emergency Accident",
    "52": "Hospital - Emergency Medical",
    "53": "Hospital - Ambulatory Surgical",
    "54": "Long Term Care",
    "55": "Major Medical",
    "56": "Medically Related Transportation",
    "57": "Air Transportation",
    "58": "Cabulance",
    "59": "Licensed Ambulance",
    "60": "General Benefits",
    "61": "In-vitro Fertilization",
    "62": "MRI/CAT Scan",
    "63": "Donor Procedures",
    "64": "Acupuncture",
    "65": "Newborn Care",
    "66": "Pathology",
    "67": "Smoking Cessation",
    "68": "Well Baby Care",
    "69": "Maternity",
    "70": "Transplants",
    "71": "Audiology Exam",
    "72": "Inhalation Therapy",
    "73": "Diagnostic Medical",
    "74": "Private Duty Nursing",
    "75": "Prosthetic Device",
    "76": "Dialysis",
    "77": "Otological Exam",
    "78": "Chemotherapy",
    "79": "Allergy Testing",
    "80": "Immunizations",
    "81": "Routine Physical",
    "82": "Family Planning",
    "83": "Infertility",
    "84": "Abortion",
    "85": "AIDS",
    "86": "Emergency Services",
    "87": "Cancer",
    "88": "Pharmacy",
    "89": "Free Standing Prescription Drug",
    "90": "Mail Order Prescription Drug",
    "91": "Brand Name Prescription Drug",
    "92": "Generic Prescription Drug",
    "93": "Podiatry",
    "94": "Podiatry - Office Visits",
    "95": "Podiatry - Nursing Home Visits",
    "96": "Professional (Physician)",
    "97": "Anesthesiologist",
    "98": "Professional (Physician) Visit - Office",
    "99": "Professional (Physician) Visit - Inpatient",
    "A0": "Professional (Physician) Visit - Outpatient",
    "A1": "Professional (Physician) Visit - Nursing Home",
    "A2": "Professional (Physician) Visit - Skilled Nursing",
    "A3": "Professional (Physician) Visit - Home",
    "A4": "Psychiatric",
    "A5": "Psychiatric - Room and Board",
    "A6": "Psychotherapy",
    "A7": "Psychiatric - Inpatient",
    "A8": "Psychiatric - Outpatient",
    "A9": "Rehabilitation",
    "AB": "Rehabilitation - Inpatient",
    "AC": "Rehabilitation - Outpatient",
    "AD": "Occupational Therapy",
    "AE": "Physical Medicine",
    "AF": "Speech Therapy",
    "AG": "Skilled Nursing Care",
    "AH": "Skilled Nursing Care - Room and Board",
    "AI": "Substance Abuse",
    "AJ": "Alcoholism",
    "AK": "Drug Addiction",
    "AL": "Vision (Optometry)",
    "AM": "Frames",
    "AN": "Routine Exam (Optometry)",
    "AO": "Lenses",
    "AQ": "Nonmedically Necessary Physical",
    "AR": "Experimental Drug Therapy",
    "BA": "Independent Medical Evaluation",
    "BB": "Partial Hospitalization (Psychiatric)",
    "BC": "Day Care (Psychiatric)",
    "BD": "Cognitive Therapy",
    "BE": "Massage Therapy",
    "BF": "Pulmonary Rehabilitation",
    "BG": "Cardiac Rehabilitation",
    "BH": "Pediatric",
    "BI": "Nursery",
    "BJ": "Skin",
    "BK": "Orthopedic",
    "BL": "Cardiac",
    "BM": "Lymphatic",
    "BN": "Gastrointestinal",
    "BP": "Endocrine",
    "BQ": "Neurology",
    "BR": "Eye",
    "BS": "Invasive Procedures",
    "UC": "Urgent Care",
}


def parse_x12_generic(edi_file):
    """Parse any X12 file generically and return sheet data.

    Args:
        edi_file: An EDIFile instance

    Returns:
        list of sheet dicts
    """
    sheets = []

    for txn_segments in edi_file.get_transactions():
        txn_type = _get_txn_type(edi_file, txn_segments)
        txn_sheets = _parse_transaction(edi_file, txn_segments, txn_type)
        sheets.extend(txn_sheets)

    return sheets


def _get_txn_type(edi, segments):
    """Get the transaction type from ST segment."""
    for seg in segments:
        elements = edi.get_elements(seg)
        if elements[0].upper() == "ST" and len(elements) > 1:
            return elements[1]
    return "Unknown"


def _parse_transaction(edi, segments, txn_type):
    """Parse a single transaction and return sheets based on type."""
    txn_name = TRANSACTION_NAMES.get(txn_type, f"X12 {txn_type}")

    if txn_type in ("270", "271"):
        return _parse_270_271(edi, segments, txn_type)
    elif txn_type in ("276", "277"):
        return _parse_276_277(edi, segments, txn_type)
    elif txn_type == "834":
        return _parse_834(edi, segments)
    elif txn_type in ("278",):
        return _parse_278(edi, segments)
    elif txn_type in ("997", "999"):
        return _parse_997_999(edi, segments, txn_type)
    else:
        return _parse_raw_x12(edi, segments, txn_type)


# ---------------------------------------------------------------------------
# 270/271 — Eligibility Inquiry/Response
# ---------------------------------------------------------------------------

def _parse_270_271(edi, segments, txn_type):
    """Parse 270 (inquiry) or 271 (response) eligibility transactions."""
    is_response = (txn_type == "271")
    names = []
    eligibility_rows = []
    current_entity = None
    current_name = {}

    for seg_str in segments:
        elements = edi.get_elements(seg_str)
        seg_id = elements[0].upper()

        if seg_id == "HL":
            pass  # Hierarchy tracking (could enhance)

        elif seg_id == "NM1":
            entity = elements[1] if len(elements) > 1 else ""
            entity_type = elements[2] if len(elements) > 2 else ""
            last = elements[3] if len(elements) > 3 else ""
            first = elements[4] if len(elements) > 4 else ""
            id_qual = elements[8] if len(elements) > 8 else ""
            id_code = elements[9] if len(elements) > 9 else ""

            if entity_type == "1":
                name = f"{last}, {first}" if first else last
            else:
                name = last

            current_entity = entity
            current_name = {
                "entity_code": entity,
                "entity_type": ENTITY_CODES.get(entity, entity),
                "name": name,
                "id_qualifier": id_qual,
                "id": id_code,
            }
            names.append(current_name)

        elif seg_id == "DMG":
            dob = format_edi_date(elements[2]) if len(elements) > 2 else ""
            gender = elements[3] if len(elements) > 3 else ""
            if current_name:
                current_name["dob"] = dob
                current_name["gender"] = {"M": "Male", "F": "Female", "U": "Unknown"}.get(gender, gender)

        elif seg_id == "DTP":
            qualifier = elements[1] if len(elements) > 1 else ""
            date_val = elements[3] if len(elements) > 3 else ""
            if current_name:
                if qualifier == "291":
                    current_name["plan_date"] = format_edi_date(date_val)
                elif qualifier == "307":
                    current_name["eligibility_date"] = format_edi_date(date_val)

        elif seg_id == "EB" and is_response:
            info_code = elements[1] if len(elements) > 1 else ""
            coverage_level = elements[2] if len(elements) > 2 else ""
            service_type = elements[3] if len(elements) > 3 else ""
            plan_name = elements[4] if len(elements) > 4 else ""
            time_period = elements[5] if len(elements) > 5 else ""
            amount = safe_float(elements[6]) if len(elements) > 6 else ""
            pct = elements[7] if len(elements) > 7 else ""

            eligibility_rows.append({
                "info_code": info_code,
                "info_description": EB_INFO_CODES.get(info_code, info_code),
                "coverage_level": coverage_level,
                "service_type": service_type,
                "service_description": SERVICE_TYPE_CODES.get(service_type, service_type),
                "plan_name": plan_name,
                "time_period": time_period,
                "amount": amount,
                "percentage": pct,
                "entity_name": current_name.get("name", "") if current_name else "",
            })

        elif seg_id == "AAA":
            # Rejection/error
            valid = elements[1] if len(elements) > 1 else ""
            reject_code = elements[3] if len(elements) > 3 else ""
            follow_up = elements[4] if len(elements) > 4 else ""
            eligibility_rows.append({
                "info_code": "REJECT",
                "info_description": f"Reject: {reject_code}",
                "coverage_level": "",
                "service_type": "",
                "service_description": "",
                "plan_name": "",
                "time_period": "",
                "amount": "",
                "percentage": "",
                "entity_name": current_name.get("name", "") if current_name else "",
            })

    sheets = []

    # Names/Entities sheet
    if names:
        n_headers = ["Entity Type", "Name", "ID", "ID Qualifier",
                      "Date of Birth", "Gender", "Plan Date", "Eligibility Date"]
        n_rows = [[
            n.get("entity_type", ""), n.get("name", ""),
            n.get("id", ""), n.get("id_qualifier", ""),
            n.get("dob", ""), n.get("gender", ""),
            n.get("plan_date", ""), n.get("eligibility_date", ""),
        ] for n in names]
        label = "271 Entities" if is_response else "270 Entities"
        sheets.append({"name": label, "headers": n_headers,
                        "rows": n_rows, "currency_cols": []})

    # Eligibility/Benefits sheet (271 only)
    if eligibility_rows:
        e_headers = ["Entity Name", "Information Type", "Description",
                      "Coverage Level", "Service Type", "Service Description",
                      "Plan Name", "Time Period", "Amount", "Percentage"]
        e_rows = [[
            r["entity_name"], r["info_code"], r["info_description"],
            r["coverage_level"], r["service_type"], r["service_description"],
            r["plan_name"], r["time_period"], r["amount"], r["percentage"],
        ] for r in eligibility_rows]
        sheets.append({"name": "271 Benefits", "headers": e_headers,
                        "rows": e_rows, "currency_cols": [9]})

    return sheets


# ---------------------------------------------------------------------------
# 276/277 — Claim Status Request/Response
# ---------------------------------------------------------------------------

def _parse_276_277(edi, segments, txn_type):
    """Parse 276/277 claim status transactions."""
    is_response = (txn_type == "277")
    entries = []
    current = {}
    current_entity = ""

    for seg_str in segments:
        elements = edi.get_elements(seg_str)
        seg_id = elements[0].upper()

        if seg_id == "NM1":
            entity = elements[1] if len(elements) > 1 else ""
            last = elements[3] if len(elements) > 3 else ""
            first = elements[4] if len(elements) > 4 else ""
            id_code = elements[9] if len(elements) > 9 else ""

            name = f"{last}, {first}".strip(", ")
            entity_desc = ENTITY_CODES.get(entity, entity)

            if entity in ("IL", "QC"):
                current["patient_name"] = name
                current["patient_id"] = id_code
            elif entity in ("85", "1P"):
                current["provider_name"] = name
                current["provider_id"] = id_code
            elif entity == "PR":
                current["payer_name"] = name
                current["payer_id"] = id_code

        elif seg_id == "TRN":
            current["trace_number"] = elements[2] if len(elements) > 2 else ""

        elif seg_id == "REF":
            qualifier = elements[1] if len(elements) > 1 else ""
            value = elements[2] if len(elements) > 2 else ""
            if qualifier in ("EJ", "BLT", "D9"):
                current["claim_id"] = value
            elif qualifier == "1K":
                current["payer_claim_id"] = value

        elif seg_id == "DTP":
            qualifier = elements[1] if len(elements) > 1 else ""
            date_val = elements[3] if len(elements) > 3 else ""
            if qualifier == "472":
                current["service_date"] = format_edi_date(date_val)
            elif qualifier == "050":
                current["received_date"] = format_edi_date(date_val)

        elif seg_id == "AMT":
            qualifier = elements[1] if len(elements) > 1 else ""
            amount = safe_float(elements[2]) if len(elements) > 2 else ""
            if qualifier == "T3":
                current["total_charge"] = amount

        elif seg_id == "STC" and is_response:
            # Status information
            stc1 = elements[1] if len(elements) > 1 else ""
            parts = edi.get_sub_elements(stc1)
            cat_code = parts[0] if len(parts) > 0 else ""
            status_code = parts[1] if len(parts) > 1 else ""
            entity_code = parts[2] if len(parts) > 2 else ""

            current["status_category"] = cat_code
            current["status_code"] = status_code
            current["status_date"] = format_edi_date(elements[2]) if len(elements) > 2 else ""
            current["total_charge"] = safe_float(elements[4]) if len(elements) > 4 else current.get("total_charge", "")

        elif seg_id == "SE":
            if current:
                entries.append(current)
                current = {}

    if current:
        entries.append(current)

    if not entries:
        return []

    label = "277 Claim Status" if is_response else "276 Status Request"
    headers = ["Patient Name", "Patient ID", "Provider Name", "Provider ID",
               "Payer Name", "Claim ID", "Service Date", "Total Charge",
               "Trace Number"]
    if is_response:
        headers.extend(["Status Category", "Status Code", "Status Date", "Payer Claim ID"])

    rows = []
    for e in entries:
        row = [
            e.get("patient_name", ""), e.get("patient_id", ""),
            e.get("provider_name", ""), e.get("provider_id", ""),
            e.get("payer_name", ""), e.get("claim_id", ""),
            e.get("service_date", ""), e.get("total_charge", ""),
            e.get("trace_number", ""),
        ]
        if is_response:
            row.extend([
                e.get("status_category", ""), e.get("status_code", ""),
                e.get("status_date", ""), e.get("payer_claim_id", ""),
            ])
        rows.append(row)

    return [{"name": label, "headers": headers, "rows": rows,
             "currency_cols": [8] if not is_response else [8]}]


# ---------------------------------------------------------------------------
# 834 — Benefit Enrollment and Maintenance
# ---------------------------------------------------------------------------

def _parse_834(edi, segments):
    """Parse 834 enrollment transactions."""
    members = []
    current_member = {}
    sponsor_name = ""

    for seg_str in segments:
        elements = edi.get_elements(seg_str)
        seg_id = elements[0].upper()

        if seg_id == "N1":
            entity = elements[1] if len(elements) > 1 else ""
            name = elements[2] if len(elements) > 2 else ""
            if entity == "P5":  # Plan Sponsor
                sponsor_name = name
            elif entity == "IN":  # Insurer
                pass

        elif seg_id == "INS":
            if current_member:
                members.append(current_member)
            current_member = {
                "sponsor_name": sponsor_name,
                "subscriber_indicator": elements[1] if len(elements) > 1 else "",
                "relationship": elements[2] if len(elements) > 2 else "",
                "maintenance_type": elements[3] if len(elements) > 3 else "",
                "benefit_status": elements[5] if len(elements) > 5 else "",
            }

        elif seg_id == "NM1" and current_member is not None:
            entity = elements[1] if len(elements) > 1 else ""
            last = elements[3] if len(elements) > 3 else ""
            first = elements[4] if len(elements) > 4 else ""
            id_code = elements[9] if len(elements) > 9 else ""
            if entity == "IL":
                current_member["name"] = f"{last}, {first}".strip(", ")
                current_member["member_id"] = id_code

        elif seg_id == "DMG" and current_member is not None:
            current_member["dob"] = format_edi_date(elements[2]) if len(elements) > 2 else ""
            gender = elements[3] if len(elements) > 3 else ""
            current_member["gender"] = {"M": "Male", "F": "Female", "U": "Unknown"}.get(gender, gender)

        elif seg_id == "DTP" and current_member is not None:
            qualifier = elements[1] if len(elements) > 1 else ""
            date_val = elements[3] if len(elements) > 3 else ""
            if qualifier == "336":
                current_member["coverage_start"] = format_edi_date(date_val)
            elif qualifier == "337":
                current_member["coverage_end"] = format_edi_date(date_val)
            elif qualifier == "303":
                current_member["maintenance_effective"] = format_edi_date(date_val)

        elif seg_id == "HD" and current_member is not None:
            current_member["maintenance_code"] = elements[1] if len(elements) > 1 else ""
            current_member["plan_code"] = elements[3] if len(elements) > 3 else ""
            current_member["coverage_type"] = elements[5] if len(elements) > 5 else ""

        elif seg_id == "N3" and current_member is not None:
            current_member["address"] = elements[1] if len(elements) > 1 else ""

        elif seg_id == "N4" and current_member is not None:
            city = elements[1] if len(elements) > 1 else ""
            state = elements[2] if len(elements) > 2 else ""
            zip_code = elements[3] if len(elements) > 3 else ""
            current_member["city_state_zip"] = f"{city}, {state} {zip_code}".strip(", ")

    if current_member:
        members.append(current_member)

    if not members:
        return []

    headers = ["Name", "Member ID", "Date of Birth", "Gender",
               "Subscriber?", "Relationship", "Maintenance Type",
               "Benefit Status", "Plan Code", "Coverage Type",
               "Coverage Start", "Coverage End",
               "Address", "City/State/Zip", "Sponsor"]
    rows = [[
        m.get("name", ""), m.get("member_id", ""),
        m.get("dob", ""), m.get("gender", ""),
        m.get("subscriber_indicator", ""), m.get("relationship", ""),
        m.get("maintenance_type", ""), m.get("benefit_status", ""),
        m.get("plan_code", ""), m.get("coverage_type", ""),
        m.get("coverage_start", ""), m.get("coverage_end", ""),
        m.get("address", ""), m.get("city_state_zip", ""),
        m.get("sponsor_name", ""),
    ] for m in members]

    return [{"name": "834 Enrollment", "headers": headers, "rows": rows, "currency_cols": []}]


# ---------------------------------------------------------------------------
# 278 — Prior Authorization
# ---------------------------------------------------------------------------

def _parse_278(edi, segments):
    """Parse 278 prior authorization transactions."""
    entries = []
    current = {}

    for seg_str in segments:
        elements = edi.get_elements(seg_str)
        seg_id = elements[0].upper()

        if seg_id == "NM1":
            entity = elements[1] if len(elements) > 1 else ""
            last = elements[3] if len(elements) > 3 else ""
            first = elements[4] if len(elements) > 4 else ""
            id_code = elements[9] if len(elements) > 9 else ""
            name = f"{last}, {first}".strip(", ")

            if entity in ("IL", "QC"):
                current["patient_name"] = name
                current["patient_id"] = id_code
            elif entity in ("85", "1P"):
                current["provider_name"] = name
                current["provider_id"] = id_code
            elif entity == "PR":
                current["payer_name"] = name

        elif seg_id == "TRN":
            current["trace_number"] = elements[2] if len(elements) > 2 else ""

        elif seg_id == "UM":
            current["request_category"] = elements[1] if len(elements) > 1 else ""
            current["certification_type"] = elements[2] if len(elements) > 2 else ""
            current["service_type"] = elements[3] if len(elements) > 3 else ""

        elif seg_id == "HCR":
            current["decision"] = elements[1] if len(elements) > 1 else ""
            current["auth_number"] = elements[2] if len(elements) > 2 else ""

        elif seg_id == "SV1" or seg_id == "SV2":
            proc = elements[1] if len(elements) > 1 else ""
            sub = edi.get_sub_elements(proc)
            current["procedure_code"] = sub[1] if len(sub) > 1 else ""

        elif seg_id == "DTP":
            qualifier = elements[1] if len(elements) > 1 else ""
            date_val = elements[3] if len(elements) > 3 else ""
            if qualifier == "472":
                current["service_date"] = format_edi_date(date_val)

        elif seg_id == "HI":
            for i in range(1, len(elements)):
                if elements[i]:
                    parts = edi.get_sub_elements(elements[i])
                    code = parts[1] if len(parts) > 1 else ""
                    if code:
                        existing = current.get("diagnosis_codes", "")
                        current["diagnosis_codes"] = f"{existing}, {code}".strip(", ")

        elif seg_id == "SE":
            if current:
                entries.append(current)
                current = {}

    if current:
        entries.append(current)

    if not entries:
        return []

    headers = ["Patient Name", "Patient ID", "Provider Name", "Provider ID",
               "Payer Name", "Trace Number",
               "Request Category", "Certification Type", "Service Type",
               "Procedure Code", "Diagnosis Codes", "Service Date",
               "Decision", "Auth Number"]
    rows = [[
        e.get("patient_name", ""), e.get("patient_id", ""),
        e.get("provider_name", ""), e.get("provider_id", ""),
        e.get("payer_name", ""), e.get("trace_number", ""),
        e.get("request_category", ""), e.get("certification_type", ""),
        e.get("service_type", ""), e.get("procedure_code", ""),
        e.get("diagnosis_codes", ""), e.get("service_date", ""),
        e.get("decision", ""), e.get("auth_number", ""),
    ] for e in entries]

    return [{"name": "278 Prior Auth", "headers": headers, "rows": rows, "currency_cols": []}]


# ---------------------------------------------------------------------------
# 997/999 — Functional/Implementation Acknowledgment
# ---------------------------------------------------------------------------

def _parse_997_999(edi, segments, txn_type):
    """Parse 997/999 acknowledgment transactions."""
    acks = []

    for seg_str in segments:
        elements = edi.get_elements(seg_str)
        seg_id = elements[0].upper()

        if seg_id == "AK1":
            acks.append({
                "functional_id": elements[1] if len(elements) > 1 else "",
                "group_control": elements[2] if len(elements) > 2 else "",
            })
        elif seg_id == "AK9" or seg_id == "AK5":
            status = elements[1] if len(elements) > 1 else ""
            status_desc = {"A": "Accepted", "E": "Accepted with Errors",
                           "R": "Rejected", "P": "Partially Accepted"}.get(status, status)
            if acks:
                acks[-1]["status"] = status_desc
                if seg_id == "AK9":
                    acks[-1]["included_txns"] = elements[2] if len(elements) > 2 else ""
                    acks[-1]["received_txns"] = elements[3] if len(elements) > 3 else ""
                    acks[-1]["accepted_txns"] = elements[4] if len(elements) > 4 else ""
        elif seg_id == "IK5" or seg_id == "AK5":
            status = elements[1] if len(elements) > 1 else ""
            status_desc = {"A": "Accepted", "E": "Accepted with Errors",
                           "R": "Rejected"}.get(status, status)
            if acks:
                acks[-1]["txn_status"] = status_desc

    if not acks:
        return []

    label = "999 Acknowledgment" if txn_type == "999" else "997 Acknowledgment"
    headers = ["Functional ID", "Group Control #", "Status",
               "Included Txns", "Received Txns", "Accepted Txns"]
    rows = [[
        a.get("functional_id", ""), a.get("group_control", ""),
        a.get("status", a.get("txn_status", "")),
        a.get("included_txns", ""), a.get("received_txns", ""),
        a.get("accepted_txns", ""),
    ] for a in acks]

    return [{"name": label, "headers": headers, "rows": rows, "currency_cols": []}]


# ---------------------------------------------------------------------------
# Raw/Generic X12
# ---------------------------------------------------------------------------

def _parse_raw_x12(edi, segments, txn_type):
    """Generic fallback: parse all segments into a raw view."""
    txn_name = TRANSACTION_NAMES.get(txn_type, txn_type)

    headers = ["Segment ID", "Elements"]
    rows = []
    for seg_str in segments:
        elements = edi.get_elements(seg_str)
        seg_id = elements[0]
        rest = edi.element_sep.join(elements[1:]) if len(elements) > 1 else ""
        rows.append([seg_id, rest])

    sheet_name = f"X12 {txn_type} Segments"
    if len(sheet_name) > 31:
        sheet_name = sheet_name[:31]

    return [{"name": sheet_name, "headers": headers, "rows": rows, "currency_cols": []}]
