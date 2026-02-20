"""Parser for NCPDP Telecommunications Standard (pharmacy claims)."""


# NCPDP delimiters (non-printable control characters)
SEGMENT_SEP = "\x1e"  # Record Separator (RS)
GROUP_SEP = "\x1d"    # Group Separator (GS)
FIELD_SEP = "\x1c"    # File Separator (FS)

# Transaction codes
TRANSACTION_CODES = {
    "E1": "Eligibility Verification",
    "B1": "Billing",
    "B2": "Reversal",
    "B3": "Rebill",
    "P1": "Prior Auth Request Billing",
    "P2": "Prior Auth Reversal",
    "P3": "Prior Auth Inquiry",
    "P4": "Prior Auth Request Only",
    "N1": "Information Reporting",
    "N2": "Information Reporting Reversal",
    "N3": "Information Reporting Rebill",
    "C1": "Controlled Substance Reporting",
    "C2": "Controlled Substance Reporting Reversal",
    "C3": "Controlled Substance Reporting Rebill",
}

# Segment names (AM codes)
SEGMENT_NAMES = {
    "AM01": "Patient",
    "AM02": "Pharmacy Provider",
    "AM03": "Prescriber",
    "AM04": "Insurance",
    "AM05": "COB/Other Payments",
    "AM06": "Workers Compensation",
    "AM07": "Claim",
    "AM08": "DUR/PPS",
    "AM09": "Coupon",
    "AM10": "Compound",
    "AM11": "Pricing",
    "AM12": "Prior Authorization",
    "AM13": "Clinical",
    "AM14": "Additional Documentation",
    "AM15": "Facility",
    "AM16": "Narrative",
    "AM20": "Response Message",
    "AM21": "Response Status",
    "AM22": "Response Claim",
    "AM23": "Response Pricing",
    "AM24": "Response DUR/PPS",
    "AM25": "Response Insurance",
    "AM26": "Response Prior Authorization",
}

# Common NCPDP field codes and their human-readable names
FIELD_NAMES = {
    # Header fields
    "A1": "BIN Number",
    "A2": "Version/Release Number",
    "A3": "Transaction Code",
    "A4": "Processor Control Number",
    "A9": "Transaction Count",
    "A6": "Service Provider ID Qualifier",
    "A7": "Service Provider ID",
    "A5": "Date of Service",
    "AK": "Software Vendor/Certification ID",
    # Patient segment
    "CA": "Patient ID Qualifier",
    "CB": "Patient ID",
    "CC": "Date of Birth",
    "CD": "Patient Gender Code",
    "CE": "Patient First Name",
    "CF": "Patient Last Name",
    "CG": "Patient Street Address",
    "CH": "Patient City",
    "CI": "Patient State",
    "CJ": "Patient Zip Code",
    "CK": "Patient Phone Number",
    "CL": "Patient Location Code",
    "CM": "Employer ID",
    "CN": "Smoker/Non-Smoker Code",
    "CX": "Patient Email Address",
    "CY": "Patient Residence",
    # Insurance segment
    "C1": "Group ID",
    "C2": "Cardholder ID",
    "C3": "Person Code",
    "C6": "Patient Relationship Code",
    "C8": "Other Coverage Code",
    "C9": "Eligibility Clarification Code",
    "CC": "Cardholder First Name",
    "CD": "Cardholder Last Name",
    # Claim segment
    "D1": "Date of Service",
    "D2": "Prescription/Service Reference # Qualifier",
    "D3": "Fill Number",
    "D4": "DAW/Product Selection Code",
    "D5": "Compound Code",
    "D6": "Number of Refills Authorized",
    "D7": "Product/Service ID Qualifier",
    "D8": "Dispensing Status",
    "D9": "Date Prescription Written",
    "DA": "Number of Refills Authorized",
    "DB": "Prescription Origin Code",
    "DC": "Submission Clarification Code",
    "DD": "Quantity Prescribed",
    "DE": "Other Coverage Code",
    "DF": "Unit of Measure",
    "DG": "Pharmacy Service Type",
    "DJ": "Product/Service ID",
    "DK": "Prescription/Service Reference Number",
    "DQ": "Usual & Customary Charge",
    "DR": "Ingredient Cost Submitted",
    "DS": "Dispensing Fee Submitted",
    "DT": "Patient Paid Amount Submitted",
    "DU": "Days Supply",
    "DV": "Gross Amount Due",
    "DW": "Basis of Cost Determination",
    "DX": "Quantity Dispensed",
    "DY": "Level of Service",
    "DZ": "Reason for Service Code",
    # Prescriber segment
    "DB": "Prescriber ID Qualifier",
    "DR": "Prescriber ID",
    "2E": "Prescriber Last Name",
    "2F": "Prescriber First Name",
    "2G": "Prescriber Street Address",
    "2H": "Prescriber City",
    "2J": "Prescriber State",
    "2K": "Prescriber Zip Code",
    # Pharmacy Provider segment
    "B1": "Provider ID Qualifier",
    "B2": "Provider ID",
    # Pricing segment
    "HA": "Ingredient Cost Paid",
    "HB": "Dispensing Fee Paid",
    "HC": "Tax Exempt Indicator",
    "HD": "Patient Sales Tax Amount",
    "HE": "Flat Sales Tax Amount Paid",
    "HF": "Percentage Sales Tax Amount Paid",
    "HG": "Percentage Sales Tax Rate Paid",
    "HH": "Percentage Sales Tax Basis Paid",
    "HJ": "Incentive Amount Paid",
    "HK": "Professional Service Fee Paid",
    "HN": "Other Amount Paid",
    "HP": "Patient Pay Amount",
    # Response fields
    "AN": "Response Transaction Code",
    "F3": "Authorization Number",
    "F4": "Reject Code",
    "F5": "Reject Count",
    "F6": "Approved Message Code",
    "F9": "Additional Message Information",
    "FA": "Additional Message Qualifier",
    "FB": "Additional Message Count",
    "FC": "Remaining Benefit Amount",
    "FD": "Accumulated Deductible Amount",
    "FE": "Remaining Deductible Amount",
    "FH": "Plan ID",
    "FI": "Network Reimbursement ID",
    "FJ": "Payer ID Qualifier",
    "FK": "Payer ID",
}


def parse_ncpdp(content):
    """Parse NCPDP content and return sheet data.

    Returns:
        list of sheet dicts
    """
    transactions = _split_transactions(content)
    if not transactions:
        return []

    all_headers = []
    all_claims = []

    for txn in transactions:
        header, claim_groups = _parse_transaction(txn)
        if header:
            all_headers.append(header)
        all_claims.extend(claim_groups)

    sheets = []

    # --- Transaction Headers sheet ---
    if all_headers:
        h_headers = ["BIN Number", "Version", "Transaction Code", "Transaction Name",
                      "Processor Control Number", "Service Provider ID",
                      "Date of Service", "Transaction Count"]
        h_rows = []
        for h in all_headers:
            txn_code = h.get("transaction_code", "")
            h_rows.append([
                h.get("bin_number", ""),
                h.get("version", ""),
                txn_code,
                TRANSACTION_CODES.get(txn_code, txn_code),
                h.get("processor_control_number", ""),
                h.get("service_provider_id", ""),
                h.get("date_of_service", ""),
                h.get("transaction_count", ""),
            ])
        sheets.append({"name": "NCPDP Transactions", "headers": h_headers,
                        "rows": h_rows, "currency_cols": []})

    # --- Claims/Fields sheet ---
    if all_claims:
        # Collect all field names across all claims
        all_field_names = []
        for claim in all_claims:
            for k in claim:
                if k not in all_field_names:
                    all_field_names.append(k)

        # Use human-readable names
        c_headers = [FIELD_NAMES.get(k, k) for k in all_field_names]
        c_rows = [[claim.get(k, "") for k in all_field_names] for claim in all_claims]

        # Identify currency columns
        currency_fields = {"DQ", "DR", "DS", "DT", "DV", "HA", "HB", "HD", "HE",
                            "HF", "HJ", "HK", "HN", "HP", "FC", "FD", "FE"}
        currency_cols = [i + 1 for i, k in enumerate(all_field_names) if k in currency_fields]

        sheets.append({"name": "NCPDP Claims", "headers": c_headers,
                        "rows": c_rows, "currency_cols": currency_cols})

    # --- All Fields (Raw) sheet ---
    # Flatten all fields with their codes for reference
    if all_claims:
        raw_headers = ["Field Code", "Field Name", "Value"]
        raw_rows = []
        for i, claim in enumerate(all_claims):
            for code, value in claim.items():
                if value:
                    raw_rows.append([code, FIELD_NAMES.get(code, code), value])
            if i < len(all_claims) - 1:
                raw_rows.append(["---", "--- New Claim ---", "---"])
        sheets.append({"name": "NCPDP Raw Fields", "headers": raw_headers,
                        "rows": raw_rows, "currency_cols": []})

    return sheets


def _split_transactions(content):
    """Split NCPDP content into individual transactions."""
    # NCPDP can use control characters or be a single transaction
    if GROUP_SEP in content:
        return content.split(GROUP_SEP)
    return [content]


def _parse_transaction(txn_content):
    """Parse a single NCPDP transaction.

    Returns:
        (header_dict, list_of_claim_dicts)
    """
    # Determine if this uses control character delimiters
    if FIELD_SEP in txn_content:
        return _parse_control_char_format(txn_content)
    elif SEGMENT_SEP in txn_content:
        return _parse_control_char_format(txn_content)
    else:
        # Try fixed-position header parsing
        return _parse_fixed_format(txn_content)


def _parse_control_char_format(content):
    """Parse NCPDP with control character delimiters."""
    header = {}
    claims = []
    current_claim = {}

    # Split into segments
    segments = content.split(SEGMENT_SEP)

    for seg in segments:
        if not seg.strip():
            continue

        # Split fields
        fields = seg.split(FIELD_SEP)

        for field in fields:
            if len(field) < 2:
                continue

            # First 2 characters are the field identifier
            field_id = field[:2].upper()
            field_value = field[2:].strip()

            if not field_value:
                continue

            # Header fields
            if field_id in ("A1", "A2", "A3", "A4", "A9", "A6", "A7", "A5", "AK"):
                if field_id == "A1":
                    header["bin_number"] = field_value
                elif field_id == "A2":
                    header["version"] = field_value
                elif field_id == "A3":
                    header["transaction_code"] = field_value
                elif field_id == "A4":
                    header["processor_control_number"] = field_value
                elif field_id == "A9":
                    header["transaction_count"] = field_value
                elif field_id == "A7":
                    header["service_provider_id"] = field_value
                elif field_id == "A5":
                    header["date_of_service"] = field_value

            # Claim/segment fields
            current_claim[field_id] = field_value

    if current_claim:
        claims.append(current_claim)

    return header, claims


def _parse_fixed_format(content):
    """Parse NCPDP with fixed-position header."""
    header = {}
    claims = [{}]

    content = content.strip()
    if len(content) < 10:
        return header, []

    # Try to detect if it starts with a 6-digit BIN
    if content[:6].isdigit():
        header["bin_number"] = content[:6]
        header["version"] = content[6:8]
        header["transaction_code"] = content[8:10]
        if len(content) > 20:
            header["processor_control_number"] = content[10:20].strip()
        if len(content) > 21:
            header["transaction_count"] = content[20]
        if len(content) > 37:
            header["service_provider_id"] = content[23:38].strip()
        if len(content) > 45:
            header["date_of_service"] = content[38:46]

        # Remaining content after header
        remaining = content[46:] if len(content) > 46 else ""
        if remaining:
            claims[0]["raw_data"] = remaining[:200]
    else:
        # Unknown format — store as raw
        claims[0]["raw_data"] = content[:500]

    return header, claims if any(claims[0].values()) else []
