"""Parser for EDI 837 (Health Care Claim)."""

from edi_parser import EDIFile, safe_float, safe_int, format_edi_date

# Place of service code descriptions
PLACE_OF_SERVICE = {
    "01": "Pharmacy",
    "02": "Telehealth (other than patient home)",
    "03": "School",
    "04": "Homeless Shelter",
    "05": "Indian Health Service Free-Standing",
    "06": "Indian Health Service Provider-Based",
    "07": "Tribal 638 Free-Standing",
    "08": "Tribal 638 Provider-Based",
    "09": "Prison/Correctional Facility",
    "10": "Telehealth (patient home)",
    "11": "Office",
    "12": "Home",
    "13": "Assisted Living Facility",
    "14": "Group Home",
    "15": "Mobile Unit",
    "16": "Temporary Lodging",
    "17": "Walk-in Retail Health Clinic",
    "18": "Place of Employment/Worksite",
    "19": "Off-Campus Outpatient Hospital",
    "20": "Urgent Care Facility",
    "21": "Inpatient Hospital",
    "22": "On-Campus Outpatient Hospital",
    "23": "Emergency Room - Hospital",
    "24": "Ambulatory Surgical Center",
    "25": "Birthing Center",
    "26": "Military Treatment Facility",
    "31": "Skilled Nursing Facility",
    "32": "Nursing Facility",
    "33": "Custodial Care Facility",
    "34": "Hospice",
    "41": "Ambulance - Land",
    "42": "Ambulance - Air or Water",
    "49": "Independent Clinic",
    "50": "Federally Qualified Health Center",
    "51": "Inpatient Psychiatric Facility",
    "52": "Psychiatric Facility - Partial Hospitalization",
    "53": "Community Mental Health Center",
    "54": "Intermediate Care Facility/MR",
    "55": "Residential Substance Abuse Treatment",
    "56": "Psychiatric Residential Treatment Center",
    "57": "Non-Residential Substance Abuse Treatment",
    "60": "Mass Immunization Center",
    "61": "Comprehensive Inpatient Rehab Facility",
    "62": "Comprehensive Outpatient Rehab Facility",
    "65": "End-Stage Renal Disease Treatment Facility",
    "71": "State or Local Public Health Clinic",
    "72": "Rural Health Clinic",
    "81": "Independent Laboratory",
    "99": "Other Place of Service",
}

# NM1 entity identifier code descriptions
ENTITY_CODES = {
    "85": "Billing Provider",
    "87": "Pay-to Provider",
    "IL": "Subscriber/Insured",
    "QC": "Patient",
    "PR": "Payer",
    "DN": "Referring Provider",
    "82": "Rendering Provider",
    "77": "Service Facility",
    "DQ": "Supervising Provider",
    "DK": "Ordering Provider",
    "PE": "Payee",
    "71": "Attending Provider",
    "72": "Operating Provider",
}


def parse_837(edi_file):
    """Parse an 837 EDI file and return structured data.

    Args:
        edi_file: An EDIFile instance

    Returns:
        list of dicts with claim data
    """
    results = []

    for transaction_segments in edi_file.get_transactions():
        result = _parse_transaction(edi_file, transaction_segments)
        if result:
            results.append(result)

    return results


def _parse_transaction(edi, segments):
    """Parse a single 837 transaction (ST..SE)."""

    # We'll use the HL hierarchy to organize data
    # HL*seq*parent*level*child_code
    # level 20 = Information Source (Payer)
    # level 22 = Subscriber
    # level 23 = Patient/Dependent

    claims = []
    providers = []

    # Current context tracking
    billing_provider = {
        "name": "", "npi": "", "tax_id": "",
        "address_line": "", "city": "", "state": "", "zip": "",
    }
    current_subscriber = {
        "name": "", "member_id": "", "dob": "", "gender": "",
        "address_line": "", "city": "", "state": "", "zip": "",
        "payer_name": "", "payer_id": "",
    }
    current_patient = None  # If different from subscriber
    current_claim = None
    current_svc = None

    hl_level = None  # Current HL level code
    in_2310_loop = None  # Track NM1 context within claim
    sub_type = None  # 837P, 837I, 837D

    for seg_str in segments:
        elements = edi.get_elements(seg_str)
        seg_id = elements[0].upper()

        if seg_id == "ST":
            # Detect 837 subtype from ST03 (implementation reference)
            ref = elements[3] if len(elements) > 3 else ""
            if "222" in ref:
                sub_type = "837P"
            elif "223" in ref:
                sub_type = "837I"
            elif "224" in ref:
                sub_type = "837D"
            else:
                sub_type = "837"

        elif seg_id == "HL":
            # Save any pending service/claim
            if current_svc and current_claim:
                current_claim["service_lines"].append(current_svc)
                current_svc = None
            if current_claim:
                _finalize_claim(current_claim, billing_provider, current_subscriber, current_patient)
                claims.append(current_claim)
                current_claim = None

            hl_level = elements[3] if len(elements) > 3 else ""
            child_code = elements[4] if len(elements) > 4 else "0"

            if hl_level == "22":
                # New subscriber — reset
                current_subscriber = {
                    "name": "", "member_id": "", "dob": "", "gender": "",
                    "address_line": "", "city": "", "state": "", "zip": "",
                    "payer_name": "", "payer_id": "",
                    "group_number": "", "group_name": "",
                }
                current_patient = None
            elif hl_level == "23":
                # Patient is different from subscriber
                current_patient = {
                    "name": "", "dob": "", "gender": "",
                    "address_line": "", "city": "", "state": "", "zip": "",
                }

        elif seg_id == "SBR":
            # Subscriber information
            if len(elements) > 9:
                current_subscriber["group_number"] = elements[3] if len(elements) > 3 else ""
                current_subscriber["group_name"] = elements[4] if len(elements) > 4 else ""

        elif seg_id == "PAT":
            pass  # Patient relationship info if needed

        elif seg_id == "NM1":
            entity = elements[1] if len(elements) > 1 else ""
            entity_type = elements[2] if len(elements) > 2 else ""
            last_or_org = elements[3] if len(elements) > 3 else ""
            first = elements[4] if len(elements) > 4 else ""
            middle = elements[5] if len(elements) > 5 else ""
            id_qualifier = elements[8] if len(elements) > 8 else ""
            id_code = elements[9] if len(elements) > 9 else ""

            if entity_type == "1":  # Person
                name = f"{last_or_org}, {first}" + (f" {middle}" if middle else "")
            else:
                name = last_or_org

            if entity == "85":  # Billing Provider
                billing_provider["name"] = name
                if id_qualifier == "XX":
                    billing_provider["npi"] = id_code
            elif entity == "IL":  # Subscriber
                current_subscriber["name"] = name
                if id_qualifier == "MI":
                    current_subscriber["member_id"] = id_code
                elif id_code:
                    current_subscriber["member_id"] = id_code
            elif entity == "QC":  # Patient
                if current_patient is not None:
                    current_patient["name"] = name
                elif hl_level == "22":
                    # Patient same as subscriber (QC in subscriber loop)
                    pass
            elif entity == "PR":  # Payer
                current_subscriber["payer_name"] = name
                current_subscriber["payer_id"] = id_code
            elif entity == "82" and current_claim:  # Rendering Provider
                current_claim["rendering_provider_name"] = name
                current_claim["rendering_provider_npi"] = id_code
            elif entity == "DN" and current_claim:  # Referring Provider
                current_claim["referring_provider_name"] = name
                current_claim["referring_provider_npi"] = id_code
            elif entity == "77" and current_claim:  # Service Facility
                current_claim["service_facility_name"] = name
                current_claim["service_facility_npi"] = id_code

        elif seg_id == "N3":
            addr = elements[1] if len(elements) > 1 else ""
            addr2 = elements[2] if len(elements) > 2 else ""
            full_addr = f"{addr} {addr2}".strip() if addr2 else addr

            if hl_level == "20" or (hl_level is None and not current_claim):
                billing_provider["address_line"] = full_addr
            elif hl_level == "23" and current_patient:
                current_patient["address_line"] = full_addr
            elif hl_level == "22":
                current_subscriber["address_line"] = full_addr

        elif seg_id == "N4":
            city = elements[1] if len(elements) > 1 else ""
            state = elements[2] if len(elements) > 2 else ""
            zip_code = elements[3] if len(elements) > 3 else ""

            if hl_level == "20" or (hl_level is None and not current_claim):
                billing_provider["city"] = city
                billing_provider["state"] = state
                billing_provider["zip"] = zip_code
            elif hl_level == "23" and current_patient:
                current_patient["city"] = city
                current_patient["state"] = state
                current_patient["zip"] = zip_code
            elif hl_level == "22":
                current_subscriber["city"] = city
                current_subscriber["state"] = state
                current_subscriber["zip"] = zip_code

        elif seg_id == "REF":
            qualifier = elements[1] if len(elements) > 1 else ""
            value = elements[2] if len(elements) > 2 else ""
            if qualifier == "EI" and not current_claim:  # Employer ID
                billing_provider["tax_id"] = value
            elif qualifier == "1G" and current_claim:  # Prior Auth
                current_claim["prior_authorization"] = value
            elif qualifier == "G1" and current_claim:  # Prior Auth
                current_claim["prior_authorization"] = value
            elif qualifier == "D9" and current_claim:  # Claim ID
                current_claim["claim_original_ref"] = value

        elif seg_id == "DMG":
            dob = format_edi_date(elements[2]) if len(elements) > 2 else ""
            gender_code = elements[3] if len(elements) > 3 else ""
            gender = {"M": "Male", "F": "Female", "U": "Unknown"}.get(gender_code, gender_code)

            if hl_level == "23" and current_patient:
                current_patient["dob"] = dob
                current_patient["gender"] = gender
            elif hl_level == "22":
                current_subscriber["dob"] = dob
                current_subscriber["gender"] = gender

        elif seg_id == "CLM":
            # Save previous claim's last service
            if current_svc and current_claim:
                current_claim["service_lines"].append(current_svc)
                current_svc = None
            if current_claim:
                _finalize_claim(current_claim, billing_provider, current_subscriber, current_patient)
                claims.append(current_claim)

            claim_id = elements[1] if len(elements) > 1 else ""
            charge = safe_float(elements[2]) if len(elements) > 2 else 0.0

            # CLM05 is composite: facility_code:qualifier:frequency
            pos_composite = elements[5] if len(elements) > 5 else ""
            pos_parts = edi.get_sub_elements(pos_composite) if pos_composite else []
            facility_code = pos_parts[0] if len(pos_parts) > 0 else ""
            pos_desc = PLACE_OF_SERVICE.get(facility_code, facility_code)
            frequency = pos_parts[2] if len(pos_parts) > 2 else ""

            current_claim = {
                "claim_id": claim_id,
                "total_charge": charge,
                "place_of_service_code": facility_code,
                "place_of_service": pos_desc,
                "frequency_code": frequency,
                "provider_signature": elements[6] if len(elements) > 6 else "",
                "assignment_code": elements[7] if len(elements) > 7 else "",
                "benefits_assignment": elements[8] if len(elements) > 8 else "",
                "release_of_info": elements[9] if len(elements) > 9 else "",
                "diagnosis_codes": [],
                "service_lines": [],
                "service_date_from": "",
                "service_date_to": "",
                "admission_date": "",
                "discharge_date": "",
                "rendering_provider_name": "",
                "rendering_provider_npi": "",
                "referring_provider_name": "",
                "referring_provider_npi": "",
                "service_facility_name": "",
                "service_facility_npi": "",
                "prior_authorization": "",
                "claim_original_ref": "",
                # These get filled by _finalize_claim
                "billing_provider_name": "",
                "billing_provider_npi": "",
                "billing_provider_tax_id": "",
                "subscriber_name": "",
                "subscriber_id": "",
                "patient_name": "",
                "patient_dob": "",
                "patient_gender": "",
                "payer_name": "",
                "payer_id": "",
            }

        elif seg_id == "HI" and current_claim:
            # Diagnosis codes — each element is composite: qualifier:code
            for i in range(1, len(elements)):
                if not elements[i]:
                    continue
                parts = edi.get_sub_elements(elements[i])
                qualifier = parts[0] if len(parts) > 0 else ""
                code = parts[1] if len(parts) > 1 else ""
                if code:
                    dx_type = "Principal" if qualifier in ("ABK", "BK") else "Other"
                    current_claim["diagnosis_codes"].append({
                        "code": code,
                        "type": dx_type,
                        "qualifier": qualifier,
                    })

        elif seg_id == "DTP":
            qualifier = elements[1] if len(elements) > 1 else ""
            fmt = elements[2] if len(elements) > 2 else ""
            date_val = elements[3] if len(elements) > 3 else ""

            if fmt == "RD8" and "-" in date_val:
                # Date range: CCYYMMDD-CCYYMMDD
                parts = date_val.split("-")
                date_from = format_edi_date(parts[0])
                date_to = format_edi_date(parts[1]) if len(parts) > 1 else ""
            else:
                date_from = format_edi_date(date_val)
                date_to = ""

            if current_svc:
                if qualifier in ("472", "150", "151"):
                    current_svc["service_date_from"] = date_from
                    current_svc["service_date_to"] = date_to
            elif current_claim:
                if qualifier == "431":
                    current_claim["service_date_from"] = date_from
                    current_claim["service_date_to"] = date_to
                elif qualifier == "472":
                    current_claim["service_date_from"] = date_from
                    current_claim["service_date_to"] = date_to
                elif qualifier == "435":
                    current_claim["admission_date"] = date_from
                elif qualifier == "096":
                    current_claim["discharge_date"] = date_from

        elif seg_id == "SV1" and current_claim:
            # Professional service line
            if current_svc:
                current_claim["service_lines"].append(current_svc)

            proc_composite = elements[1] if len(elements) > 1 else ""
            sub = edi.get_sub_elements(proc_composite)
            proc_qualifier = sub[0] if len(sub) > 0 else ""
            proc_code = sub[1] if len(sub) > 1 else ""
            mod1 = sub[2] if len(sub) > 2 else ""
            mod2 = sub[3] if len(sub) > 3 else ""
            mod3 = sub[4] if len(sub) > 4 else ""
            mod4 = sub[5] if len(sub) > 5 else ""
            modifiers = ":".join(m for m in [mod1, mod2, mod3, mod4] if m)

            charge = safe_float(elements[2]) if len(elements) > 2 else 0.0
            unit_type = elements[3] if len(elements) > 3 else ""
            units = safe_float(elements[4]) if len(elements) > 4 else 0.0
            pos = elements[5] if len(elements) > 5 else ""
            dx_pointers = elements[7] if len(elements) > 7 else ""

            current_svc = {
                "line_number": len(current_claim["service_lines"]) + 1,
                "procedure_code": proc_code,
                "procedure_qualifier": proc_qualifier,
                "modifiers": modifiers,
                "charge_amount": charge,
                "unit_type": unit_type,
                "units": units,
                "place_of_service": PLACE_OF_SERVICE.get(pos, pos),
                "diagnosis_pointers": dx_pointers,
                "service_date_from": "",
                "service_date_to": "",
                "revenue_code": "",
                "ndc_code": "",
                "_claim_id": current_claim["claim_id"],
            }

        elif seg_id == "SV2" and current_claim:
            # Institutional service line
            if current_svc:
                current_claim["service_lines"].append(current_svc)

            revenue_code = elements[1] if len(elements) > 1 else ""
            proc_composite = elements[2] if len(elements) > 2 else ""
            sub = edi.get_sub_elements(proc_composite)
            proc_code = sub[1] if len(sub) > 1 else ""
            modifiers = ""

            charge = safe_float(elements[3]) if len(elements) > 3 else 0.0
            unit_type = elements[4] if len(elements) > 4 else ""
            units = safe_float(elements[5]) if len(elements) > 5 else 0.0

            current_svc = {
                "line_number": len(current_claim["service_lines"]) + 1,
                "procedure_code": proc_code,
                "procedure_qualifier": sub[0] if sub else "",
                "modifiers": modifiers,
                "charge_amount": charge,
                "unit_type": unit_type,
                "units": units,
                "place_of_service": "",
                "diagnosis_pointers": "",
                "service_date_from": "",
                "service_date_to": "",
                "revenue_code": revenue_code,
                "ndc_code": "",
                "_claim_id": current_claim["claim_id"],
            }

        elif seg_id == "LX" and current_claim:
            # Line counter — save any pending service
            if current_svc:
                current_claim["service_lines"].append(current_svc)
                current_svc = None

        elif seg_id == "LIN" and current_svc:
            # NDC info
            if len(elements) > 2:
                if elements[2] == "N4":
                    current_svc["ndc_code"] = elements[3] if len(elements) > 3 else ""

    # Save last claim/service
    if current_svc and current_claim:
        current_claim["service_lines"].append(current_svc)
    if current_claim:
        _finalize_claim(current_claim, billing_provider, current_subscriber, current_patient)
        claims.append(current_claim)

    return {
        "transaction_type": sub_type or "837",
        "billing_provider": billing_provider,
        "claims": claims,
    }


def _finalize_claim(claim, billing_provider, subscriber, patient):
    """Copy provider/subscriber/patient info into the claim dict."""
    claim["billing_provider_name"] = billing_provider.get("name", "")
    claim["billing_provider_npi"] = billing_provider.get("npi", "")
    claim["billing_provider_tax_id"] = billing_provider.get("tax_id", "")
    claim["subscriber_name"] = subscriber.get("name", "")
    claim["subscriber_id"] = subscriber.get("member_id", "")
    claim["payer_name"] = subscriber.get("payer_name", "")
    claim["payer_id"] = subscriber.get("payer_id", "")

    if patient:
        claim["patient_name"] = patient.get("name", "")
        claim["patient_dob"] = patient.get("dob", "")
        claim["patient_gender"] = patient.get("gender", "")
    else:
        claim["patient_name"] = subscriber.get("name", "")
        claim["patient_dob"] = subscriber.get("dob", "")
        claim["patient_gender"] = subscriber.get("gender", "")
