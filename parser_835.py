"""Parser for EDI 835 (Health Care Claim Payment/Remittance Advice)."""

from edi_parser import EDIFile, safe_float, safe_int, format_edi_date

# Claim status code descriptions
CLAIM_STATUS = {
    "1": "Processed as Primary",
    "2": "Processed as Secondary",
    "3": "Processed as Tertiary",
    "4": "Denied",
    "19": "Processed as Primary, Forwarded",
    "20": "Processed as Secondary, Forwarded",
    "21": "Processed as Tertiary, Forwarded",
    "22": "Reversal of Previous Payment",
    "23": "Not Our Claim, Forwarded",
    "25": "Reject",
}

# CAS group code descriptions
CAS_GROUP_CODES = {
    "CO": "Contractual Obligation",
    "CR": "Correction/Reversal",
    "OA": "Other Adjustment",
    "PI": "Payor Initiated Reduction",
    "PR": "Patient Responsibility",
}

# Payment method descriptions
PAYMENT_METHODS = {
    "ACH": "ACH (Electronic)",
    "CHK": "Check",
    "BOP": "Financial Institution Option",
    "FWT": "Federal Reserve Wire Transfer",
    "NON": "Non-Payment Data",
}

# Common Claim Adjustment Reason Codes (CARC)
REASON_CODES = {
    "1": "Deductible",
    "2": "Coinsurance",
    "3": "Copayment",
    "4": "Procedure not consistent with modifier or modifier missing",
    "5": "Procedure code inconsistent with place of service",
    "6": "Procedure/revenue code inconsistent with patient age",
    "9": "Diagnosis inconsistent with procedure",
    "10": "Diagnosis inconsistent with patient age",
    "11": "Diagnosis inconsistent with patient gender",
    "13": "Date of death precedes date of service",
    "14": "Date of birth follows date of service",
    "16": "Claim/service lacks information needed for adjudication",
    "18": "Exact duplicate claim/service",
    "19": "Expense incurred during lapse in coverage",
    "20": "Procedure/service not covered by benefits",
    "22": "Care may be covered by another payer",
    "23": "Payment adjusted (authorized amount)",
    "24": "Charges covered under capitation agreement",
    "26": "Expenses incurred prior to coverage",
    "27": "Expenses incurred after coverage terminated",
    "29": "Timely filing limit",
    "31": "Patient not eligible on date of service",
    "32": "Our records indicate patient is an inpatient",
    "33": "Equipment supply from another provider",
    "34": "Equipment supply already provided",
    "35": "Bundled/inclusive with another service",
    "39": "Revenue code and procedure code do not match",
    "40": "Charges do not meet qualifications for emergent/urgent care",
    "44": "Exceeds plan/payer limitation",
    "45": "Charges exceed fee schedule/maximum allowable",
    "49": "Non-covered (routine/preventive)",
    "50": "Non-covered unless condition coded/documented",
    "51": "Non-covered (pre-existing condition)",
    "53": "Services by an unauthorized provider",
    "54": "Multiple physicians/ambulance suppliers",
    "55": "Procedure/treatment not provided/utilized",
    "56": "Procedure/treatment not related to condition",
    "58": "Treatment was deemed experimental",
    "59": "Processed based on multiple/concurrent procedure rules",
    "66": "Blood deductible",
    "69": "Day outlier amount",
    "70": "Cost outlier — Loss exceeds threshold",
    "89": "Professional fee removed from DRG price",
    "90": "Ingredient cost adjustment",
    "94": "Plan procedures not followed",
    "95": "Plan not approved (non-network provider)",
    "96": "Non-covered charge(s)",
    "97": "Payment adjusted (processed information)",
    "100": "Payment made to patient/insured",
    "101": "Predetermination/preauthorization/precertification pricing",
    "102": "Major medical adjustment",
    "103": "Provider promotional discount",
    "104": "Managed care withholding",
    "105": "Tax withholding",
    "106": "Patient payment option/election not in effect",
    "107": "Claim/service denied (related to another covered service)",
    "108": "Rent/purchase guidelines",
    "109": "Claim/service not covered by this payer",
    "110": "Billing data correction",
    "111": "Not covered unless specific condition is met",
    "112": "Service not furnished directly to patient",
    "114": "Procedure/product not approved by FDA",
    "115": "Procedure postponed/canceled/delayed",
    "116": "Claim lacks required network authorization",
    "117": "Payment adjusted due to patient transfer",
    "118": "Benefit reduced for diagnostic test ordered by referring physician",
    "119": "Benefit maximum for this time period reached",
    "121": "Indemnification adjustment",
    "122": "Psychiatric reduction",
    "125": "Payment adjusted (submission/billing error)",
    "128": "Newborn's services covered by mother's claim",
    "129": "Prior processing information",
    "130": "Claim submission fee",
    "131": "Claim denied (specific clinical criteria not met)",
    "132": "Prematurity adjustment (institutional claims only)",
    "133": "Adjusted based on stop-loss provisions",
    "134": "Processed through global surgery package",
    "135": "Plan limitations (e.g., number of leaves per period)",
    "136": "Failure to obtain second surgical opinion",
    "137": "Regulatory surcharges/assessments/recovery",
    "138": "Appeal or review adjustment",
    "139": "Capital costs above/below threshold",
    "140": "Patient/insured health ID not on file",
    "142": "Monthly benefit maximum reached",
    "143": "Portion of payment deferred",
    "144": "Incentive adjustment",
    "146": "Diagnosis denied (mismatched procedures)",
    "147": "Provider contracted/negotiated rate expired",
    "148": "Information from another provider not provided",
    "149": "Lifetime benefit maximum reached",
    "150": "Payer deems this not a covered service",
    "151": "Payment adjusted based on plan requirements",
    "152": "Payer deems patient not eligible for this service",
    "153": "Prior auth/preservice decision",
    "154": "Bundled charges; separate reimbursement not allowed",
    "155": "Patient refused service/treatment",
    "157": "Service/procedure denied (not authorized referral)",
    "158": "Service not authorized on this date",
    "159": "Service at inappropriate level",
    "160": "Injury/illness related to work",
    "161": "Injury/illness related to auto accident",
    "162": "Injury/illness related to another party",
    "163": "Attachment/other documentation not received",
    "164": "Attachment/other documentation incomplete",
    "166": "These services not payable from this payer",
    "167": "Diagnosis not covered",
    "169": "Alternate benefit has been provided",
    "170": "Payment is denied when performed by this provider type",
    "171": "Payment adjusted based on jurisdiction regulations",
    "172": "Payment adjusted based on reason of review",
    "173": "Service claim paid toward contribution to premium",
    "174": "Service paid at zero rate for this payer",
    "175": "Payment adjusted: prescription coverage",
    "176": "Claims paid in full",
    "177": "Patient has not met the deductible",
    "178": "Previous payment reversed due to retroactive disenrollment",
    "179": "Services not provided by network/primary care providers",
    "180": "Service not furnished in a certified ASC",
    "181": "Procedure code billed is not correct/valid",
    "182": "Secondary payment on a non-covered service",
    "183": "Service denied (point of service requirement not met)",
    "184": "Purchased service charge exceeds acceptable limit",
    "185": "Dental code submitted is not appropriate",
    "186": "Level of care not appropriate",
    "187": "Non-covered consumer directed health plan",
    "188": "Procedure code/service inconsistent with provider type/specialty",
    "189": "Not otherwise classified procedure code",
    "190": "Payment adjusted: missing or invalid CPT/HCPCS",
    "192": "Non-standard adjustment code from payer",
    "193": "Original payment decision is maintained",
    "194": "Anesthesia performed by the operating doctor",
    "195": "Refund issued to insured/subscriber",
    "197": "Precertification/notification/authorization/utilization—denied",
    "198": "Claim/service adjusted: type of claim conflict",
    "199": "Revenue code and procedure code subject to NCCI edits",
    "200": "Expenses incurred during lapse/waiting period",
    "201": "Workers comp case settled; future bills not payable",
    "202": "Non-covered personal comfort/convenience item",
    "203": "Discontinued/reduced service",
    "204": "Service/equipment/drug not covered by this payer's plan",
    "205": "Pharmacy discount",
    "206": "National Drug Code (NDC) not covered",
    "207": "Dispensing fee adjustment",
    "208": "Claim denied — pharmacy not eligible",
    "209": "Reduced for home health prospective payment",
    "210": "Payment adjusted: pre-admission testing",
    "211": "National Drug Code (NDC) not eligible",
    "212": "Claim denied: global maternity fee",
    "213": "Payment adjusted: in-network/out-of-network status",
    "215": "Payment adjusted per auto no-fault processing",
    "216": "Payment adjusted based on plan limits",
    "219": "Payment adjusted: concurrent care reduction",
    "222": "Clinical trial reduction",
    "223": "Adjustment code for mandated federal/state/local law",
    "224": "Patient identification changed",
    "225": "Duplicate of an original claim processed as an adjustment",
    "226": "Information requested from patient/insured/responsible party",
    "227": "Information requested from provider",
    "228": "Denied — no response to repeated requests",
    "229": "Patient identification compromised",
    "230": "No available/qualified review organization",
    "231": "Mutually exclusive procedures per NCCI",
    "232": "Institutional claims: readmission reduction",
    "233": "This service line is out of balance",
    "234": "Service not rendered in a network facility",
    "235": "Sales tax",
    "236": "Pharmacy: claim cost exceeds pricing threshold",
    "237": "Pharmacy: claim gap claim not covered",
    "238": "Pharmacy: claim is below cost threshold",
    "239": "Claim span dates overlap previously adjudicated dates",
    "240": "Charges adjusted based on plan benefit design",
    "241": "Low income subsidy (LIS) co-pay amount",
    "242": "Services not provided by network pharmacy",
    "243": "Services not authorized by network/primary care providers",
    "244": "Payment reduced to zero due to litigation",
    "245": "Provider performance bonus",
    "246": "This service is not covered when performed by this provider",
    "247": "Non-payable code billed in a non-covered charge",
    "248": "Non-covered service: patient requested",
    "249": "Outpatient/ER claim with admission: processed as inpatient",
    "250": "Plan deeming: services covered by a prior payer",
    "251": "Service(s) adjusted due to prior payer's adjudication",
    "252": "Adjustment using a payer-determined fee schedule",
    "253": "Sequestration: mandated reduction to federal payment",
    "254": "Observation charges bundled with inpatient admission",
    "255": "Bundled or included procedure/service",
    "256": "Service subject to regulatory mandated discount",
    "257": "Service covered at reduced rate due to plan change",
    "258": "Claim adjusted: provider-assessed appeal rights",
    "260": "Requirement for additional processing",
    "261": "DRG weight-adjusted payment",
    "262": "Claim adjusted based on payer quality program",
    "263": "Adjustment for timeliness of claims processing",
    "264": "Adjusted per Value-Based Purchasing (VBP) program",
}


def parse_835(edi_file):
    """Parse an 835 EDI file and return structured data.

    Args:
        edi_file: An EDIFile instance

    Returns:
        dict with payment info, claims, service lines, and adjustments
    """
    results = []

    for transaction_segments in edi_file.get_transactions():
        result = _parse_transaction(edi_file, transaction_segments)
        if result:
            results.append(result)

    return results


def _parse_transaction(edi, segments):
    """Parse a single 835 transaction (ST..SE)."""
    payment = {
        "payment_amount": 0.0,
        "payment_method": "",
        "payment_date": "",
        "trace_number": "",
        "payer_name": "",
        "payer_id": "",
        "payee_name": "",
        "payee_id": "",
    }
    claims = []
    current_claim = None
    current_svc = None
    context = "header"  # header, claim, service

    for seg_str in segments:
        elements = edi.get_elements(seg_str)
        seg_id = elements[0].upper()

        if seg_id == "BPR":
            payment["payment_amount"] = safe_float(elements[2]) if len(elements) > 2 else 0.0
            if len(elements) > 4:
                method = elements[4]
                payment["payment_method"] = PAYMENT_METHODS.get(method, method)
            if len(elements) > 16:
                payment["payment_date"] = format_edi_date(elements[16])

        elif seg_id == "TRN":
            if len(elements) > 2:
                payment["trace_number"] = elements[2]

        elif seg_id == "N1":
            entity = elements[1] if len(elements) > 1 else ""
            name = elements[2] if len(elements) > 2 else ""
            id_code = elements[4] if len(elements) > 4 else ""
            if entity == "PR":  # Payer
                payment["payer_name"] = name
                payment["payer_id"] = id_code
            elif entity == "PE":  # Payee
                payment["payee_name"] = name
                payment["payee_id"] = id_code

        elif seg_id == "DTM":
            qualifier = elements[1] if len(elements) > 1 else ""
            date_val = format_edi_date(elements[2]) if len(elements) > 2 else ""
            # 405 = Production Date (payment date if BPR didn't have it)
            if qualifier == "405" and not payment["payment_date"]:
                payment["payment_date"] = date_val

        elif seg_id == "CLP":
            # Save previous claim
            if current_claim:
                if current_svc:
                    current_claim["service_lines"].append(current_svc)
                    current_svc = None
                claims.append(current_claim)

            current_claim = {
                "patient_control_number": elements[1] if len(elements) > 1 else "",
                "claim_status": "",
                "claim_status_code": elements[2] if len(elements) > 2 else "",
                "total_charge": safe_float(elements[3]) if len(elements) > 3 else 0.0,
                "payment_amount": safe_float(elements[4]) if len(elements) > 4 else 0.0,
                "patient_responsibility": safe_float(elements[5]) if len(elements) > 5 else 0.0,
                "filing_indicator": elements[6] if len(elements) > 6 else "",
                "payer_claim_number": elements[7] if len(elements) > 7 else "",
                "patient_last_name": "",
                "patient_first_name": "",
                "insured_last_name": "",
                "insured_first_name": "",
                "corrected_insured_last_name": "",
                "corrected_insured_first_name": "",
                "rendering_provider_last_name": "",
                "rendering_provider_first_name": "",
                "rendering_provider_npi": "",
                "claim_received_date": "",
                "claim_statement_from": "",
                "claim_statement_to": "",
                "coverage_expiration_date": "",
                "adjustments": [],
                "service_lines": [],
            }
            status_code = current_claim["claim_status_code"]
            current_claim["claim_status"] = CLAIM_STATUS.get(status_code, status_code)
            context = "claim"
            current_svc = None

        elif seg_id == "NM1" and current_claim:
            entity = elements[1] if len(elements) > 1 else ""
            last_name = elements[3] if len(elements) > 3 else ""
            first_name = elements[4] if len(elements) > 4 else ""
            id_code = elements[9] if len(elements) > 9 else ""
            if entity == "QC":  # Patient
                current_claim["patient_last_name"] = last_name
                current_claim["patient_first_name"] = first_name
            elif entity == "IL":  # Insured/Subscriber
                current_claim["insured_last_name"] = last_name
                current_claim["insured_first_name"] = first_name
            elif entity == "74":  # Corrected Insured
                current_claim["corrected_insured_last_name"] = last_name
                current_claim["corrected_insured_first_name"] = first_name
            elif entity == "82":  # Rendering Provider
                current_claim["rendering_provider_last_name"] = last_name
                current_claim["rendering_provider_first_name"] = first_name
                current_claim["rendering_provider_npi"] = id_code

        elif seg_id == "SVC" and current_claim:
            # Save previous service line
            if current_svc:
                current_claim["service_lines"].append(current_svc)

            # SVC01 is composite: qualifier:code[:modifier[:modifier...]]
            proc_composite = elements[1] if len(elements) > 1 else ""
            sub = edi.get_sub_elements(proc_composite)
            proc_qualifier = sub[0] if len(sub) > 0 else ""
            proc_code = sub[1] if len(sub) > 1 else ""
            modifiers = sub[2:] if len(sub) > 2 else []

            # SVC06 is the original submitted procedure (if different)
            orig_composite = elements[6] if len(elements) > 6 else ""
            orig_sub = edi.get_sub_elements(orig_composite) if orig_composite else []
            orig_code = orig_sub[1] if len(orig_sub) > 1 else ""

            current_svc = {
                "procedure_code": proc_code,
                "procedure_qualifier": proc_qualifier,
                "modifiers": ":".join(modifiers) if modifiers else "",
                "charge_amount": safe_float(elements[2]) if len(elements) > 2 else 0.0,
                "payment_amount": safe_float(elements[3]) if len(elements) > 3 else 0.0,
                "revenue_code": elements[4] if len(elements) > 4 else "",
                "units_paid": safe_float(elements[5]) if len(elements) > 5 else 0.0,
                "original_procedure_code": orig_code,
                "original_units": safe_float(elements[7]) if len(elements) > 7 else 0.0,
                "service_date": "",
                "adjustments": [],
                "remark_codes": [],
                # Carry forward claim info for flat service line view
                "_claim_id": current_claim["patient_control_number"],
            }
            context = "service"

        elif seg_id == "CAS":
            group_code = elements[1] if len(elements) > 1 else ""
            group_desc = CAS_GROUP_CODES.get(group_code, group_code)

            # CAS can have up to 6 adjustment triplets (reason, amount, quantity)
            i = 2
            while i < len(elements) and i + 1 < len(elements):
                reason = elements[i] if elements[i] else ""
                amount = safe_float(elements[i + 1]) if i + 1 < len(elements) else 0.0
                qty = safe_float(elements[i + 2]) if i + 2 < len(elements) else 0.0

                if not reason:
                    i += 3
                    continue

                adj = {
                    "group_code": group_code,
                    "group_description": group_desc,
                    "reason_code": reason,
                    "reason_description": REASON_CODES.get(reason, ""),
                    "amount": amount,
                    "quantity": qty,
                }

                if context == "service" and current_svc:
                    current_svc["adjustments"].append(adj)
                elif current_claim:
                    current_claim["adjustments"].append(adj)

                i += 3

        elif seg_id == "DTM" and current_claim:
            qualifier = elements[1] if len(elements) > 1 else ""
            date_val = format_edi_date(elements[2]) if len(elements) > 2 else ""
            if context == "service" and current_svc:
                if qualifier in ("472", "150", "151"):
                    current_svc["service_date"] = date_val
            elif context == "claim":
                if qualifier == "050":
                    current_claim["claim_received_date"] = date_val
                elif qualifier == "232":
                    current_claim["claim_statement_from"] = date_val
                elif qualifier == "233":
                    current_claim["claim_statement_to"] = date_val
                elif qualifier == "036":
                    current_claim["coverage_expiration_date"] = date_val

        elif seg_id == "AMT" and current_claim:
            pass  # Could capture additional monetary amounts if needed

        elif seg_id == "LQ" and current_svc:
            qualifier = elements[1] if len(elements) > 1 else ""
            code = elements[2] if len(elements) > 2 else ""
            if qualifier in ("HE", "RX"):
                current_svc["remark_codes"].append(code)

    # Don't forget the last claim/service
    if current_claim:
        if current_svc:
            current_claim["service_lines"].append(current_svc)
        claims.append(current_claim)

    return {
        "transaction_type": "835",
        "payment": payment,
        "claims": claims,
    }
