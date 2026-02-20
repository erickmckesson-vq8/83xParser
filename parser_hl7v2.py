"""Parser for HL7 v2.x pipe-delimited (ER7) messages."""

import re

# Patient class descriptions
PATIENT_CLASS = {
    "I": "Inpatient",
    "O": "Outpatient",
    "E": "Emergency",
    "P": "Preadmit",
    "R": "Recurring Patient",
    "B": "Obstetrics",
    "N": "Not Applicable",
    "U": "Unknown",
}

# Gender descriptions
GENDER = {"M": "Male", "F": "Female", "U": "Unknown", "O": "Other", "A": "Ambiguous"}

# Diagnosis type
DX_TYPE = {
    "A": "Admitting",
    "F": "Final",
    "W": "Working",
    "D": "Discharge",
}

# Order control codes
ORDER_CONTROL = {
    "NW": "New Order",
    "OK": "Order Accepted",
    "CA": "Cancel",
    "OC": "Order Canceled",
    "DC": "Discontinue",
    "HD": "Hold",
    "RL": "Release Hold",
    "SC": "Status Changed",
    "SN": "Send Number",
    "XO": "Change Order",
    "XR": "Changed as Requested",
    "RE": "Observations to Follow",
    "RU": "Replacement Unsolicited",
    "CH": "Child Order",
    "PA": "Parent Order",
}

# Result status
RESULT_STATUS = {
    "C": "Correction",
    "D": "Deleted",
    "F": "Final",
    "I": "Specimen In Lab",
    "O": "Order Received",
    "P": "Preliminary",
    "R": "Results Entered",
    "S": "Partial",
    "X": "Canceled",
    "U": "Results Unavailable",
    "W": "Post Original as Wrong",
}


class HL7Message:
    """Represents a single parsed HL7 v2.x message."""

    def __init__(self, raw, field_sep="|", comp_sep="^", rep_sep="~",
                 esc_char="\\", sub_sep="&", seg_sep="\r"):
        self.field_sep = field_sep
        self.comp_sep = comp_sep
        self.rep_sep = rep_sep
        self.esc_char = esc_char
        self.sub_sep = sub_sep
        self.segments = []
        self._parse(raw, seg_sep)

    def _parse(self, raw, seg_sep):
        # Normalize line endings
        text = raw.replace("\r\n", "\r").replace("\n", "\r")
        for line in text.split("\r"):
            line = line.strip()
            if not line:
                continue
            self.segments.append(line)

    def get_segments(self, seg_id):
        """Return all segments matching the given ID (e.g., 'PID', 'OBX')."""
        results = []
        for seg in self.segments:
            if seg.startswith(seg_id + self.field_sep) or seg == seg_id:
                results.append(seg)
        return results

    def get_field(self, segment_str, field_num):
        """Get a field value from a segment string. Field 0 = segment ID.

        For MSH, field 1 = field separator, field 2 = encoding characters,
        field 3 = MSH-3, etc.
        """
        parts = segment_str.split(self.field_sep)
        if parts[0] == "MSH":
            # MSH is special: MSH-1 is the field separator itself
            if field_num == 0:
                return "MSH"
            elif field_num == 1:
                return self.field_sep
            elif field_num == 2:
                return parts[1] if len(parts) > 1 else ""
            else:
                idx = field_num - 1
                return parts[idx] if idx < len(parts) else ""
        else:
            return parts[field_num] if field_num < len(parts) else ""

    def get_component(self, field_val, comp_num):
        """Get a component from a field value. Component 1 = first component."""
        if not field_val:
            return ""
        parts = field_val.split(self.comp_sep)
        idx = comp_num - 1
        return parts[idx] if idx < len(parts) else ""

    def get_repetitions(self, field_val):
        """Split a field into its repetitions."""
        if not field_val:
            return []
        return field_val.split(self.rep_sep)


def parse_hl7v2(content):
    """Parse HL7 v2.x content (single or batch) and return sheet data.

    Args:
        content: raw HL7 v2.x text content

    Returns:
        list of sheet dicts: [{"name": ..., "headers": [...], "rows": [...], "currency_cols": [...]}]
    """
    messages = _split_messages(content)
    parsed_msgs = []

    for raw_msg in messages:
        msg = _parse_single(raw_msg)
        if msg:
            parsed_msgs.append(msg)

    if not parsed_msgs:
        return []

    return _build_sheets(parsed_msgs)


def _split_messages(content):
    """Split content into individual HL7 messages (handle batches)."""
    # Normalize line endings
    text = content.replace("\r\n", "\r").replace("\n", "\r")

    # Split on MSH segments
    messages = []
    current = []
    for line in text.split("\r"):
        line = line.strip()
        if not line:
            continue
        # Skip batch headers/trailers
        if line.startswith("FHS") or line.startswith("BHS") or \
           line.startswith("BTS") or line.startswith("FTS"):
            continue
        if line.startswith("MSH"):
            if current:
                messages.append("\r".join(current))
            current = [line]
        elif current:
            current.append(line)

    if current:
        messages.append("\r".join(current))

    return messages


def _parse_single(raw_msg):
    """Parse a single HL7 v2.x message into a structured dict."""
    # Detect delimiters from MSH
    if not raw_msg.startswith("MSH"):
        return None

    field_sep = raw_msg[3] if len(raw_msg) > 3 else "|"
    enc_chars = raw_msg.split(field_sep)[1] if field_sep in raw_msg else "^~\\&"
    comp_sep = enc_chars[0] if len(enc_chars) > 0 else "^"
    rep_sep = enc_chars[1] if len(enc_chars) > 1 else "~"
    esc_char = enc_chars[2] if len(enc_chars) > 2 else "\\"
    sub_sep = enc_chars[3] if len(enc_chars) > 3 else "&"

    msg = HL7Message(raw_msg, field_sep, comp_sep, rep_sep, esc_char, sub_sep)

    result = {"_msg": msg}

    # --- MSH fields ---
    msh_segs = msg.get_segments("MSH")
    if msh_segs:
        msh = msh_segs[0]
        result["sending_app"] = msg.get_component(msg.get_field(msh, 3), 1)
        result["sending_facility"] = msg.get_component(msg.get_field(msh, 4), 1)
        result["receiving_app"] = msg.get_component(msg.get_field(msh, 5), 1)
        result["receiving_facility"] = msg.get_component(msg.get_field(msh, 6), 1)
        result["message_datetime"] = _format_hl7_datetime(msg.get_field(msh, 7))
        msg_type_field = msg.get_field(msh, 9)
        result["message_type"] = msg.get_component(msg_type_field, 1)
        result["trigger_event"] = msg.get_component(msg_type_field, 2)
        result["message_structure"] = msg.get_component(msg_type_field, 3)
        result["message_control_id"] = msg.get_field(msh, 10)
        result["version"] = msg.get_component(msg.get_field(msh, 12), 1)
    else:
        return None

    # --- PID (Patient) ---
    pid_segs = msg.get_segments("PID")
    pids = []
    for pid in pid_segs:
        patient = {}
        # PID-3: Patient Identifier List
        pid3 = msg.get_field(pid, 3)
        reps = msg.get_repetitions(pid3)
        ids = []
        for rep in reps:
            id_val = msg.get_component(rep, 1)
            id_type = msg.get_component(rep, 5)
            if id_val:
                ids.append(f"{id_val} ({id_type})" if id_type else id_val)
        patient["patient_ids"] = "; ".join(ids)
        patient["patient_id"] = msg.get_component(pid3, 1) if pid3 else ""

        # PID-5: Patient Name
        pid5 = msg.get_field(pid, 5)
        last = msg.get_component(pid5, 1)
        first = msg.get_component(pid5, 2)
        middle = msg.get_component(pid5, 3)
        patient["patient_name"] = f"{last}, {first}" + (f" {middle}" if middle else "")

        # PID-7: DOB
        patient["dob"] = _format_hl7_date(msg.get_field(pid, 7))

        # PID-8: Gender
        gender = msg.get_field(pid, 8)
        patient["gender"] = GENDER.get(gender, gender)

        # PID-11: Address
        pid11 = msg.get_field(pid, 11)
        addr_parts = [
            msg.get_component(pid11, 1),  # Street
            msg.get_component(pid11, 3),  # City
            msg.get_component(pid11, 4),  # State
            msg.get_component(pid11, 5),  # Zip
        ]
        patient["address"] = ", ".join(p for p in addr_parts if p)

        # PID-13: Home Phone
        patient["home_phone"] = msg.get_component(msg.get_field(pid, 13), 1)

        # PID-18: Account Number
        patient["account_number"] = msg.get_component(msg.get_field(pid, 18), 1)

        # PID-19: SSN
        patient["ssn"] = msg.get_field(pid, 19)

        pids.append(patient)
    result["patients"] = pids

    # --- PV1 (Patient Visit) ---
    pv1_segs = msg.get_segments("PV1")
    visits = []
    for pv1 in pv1_segs:
        visit = {}
        pclass = msg.get_field(pv1, 2)
        visit["patient_class"] = PATIENT_CLASS.get(pclass, pclass)

        # PV1-3: Location
        pv1_3 = msg.get_field(pv1, 3)
        visit["location"] = msg.get_component(pv1_3, 1)
        visit["room"] = msg.get_component(pv1_3, 2)
        visit["bed"] = msg.get_component(pv1_3, 3)

        # PV1-7: Attending Doctor
        pv1_7 = msg.get_field(pv1, 7)
        visit["attending_doctor"] = _format_xcn(msg, pv1_7)

        # PV1-8: Referring Doctor
        pv1_8 = msg.get_field(pv1, 8)
        visit["referring_doctor"] = _format_xcn(msg, pv1_8)

        # PV1-10: Hospital Service
        visit["hospital_service"] = msg.get_field(pv1, 10)

        # PV1-14: Admit Source
        visit["admit_source"] = msg.get_field(pv1, 14)

        # PV1-17: Admitting Doctor
        pv1_17 = msg.get_field(pv1, 17)
        visit["admitting_doctor"] = _format_xcn(msg, pv1_17)

        # PV1-19: Visit Number
        visit["visit_number"] = msg.get_component(msg.get_field(pv1, 19), 1)

        # PV1-36: Discharge Disposition
        visit["discharge_disposition"] = msg.get_field(pv1, 36)

        # PV1-44/45: Admit/Discharge dates
        visit["admit_date"] = _format_hl7_datetime(msg.get_field(pv1, 44))
        visit["discharge_date"] = _format_hl7_datetime(msg.get_field(pv1, 45))

        visits.append(visit)
    result["visits"] = visits

    # --- ORC + OBR (Orders) ---
    orc_segs = msg.get_segments("ORC")
    obr_segs = msg.get_segments("OBR")
    orders = []

    # Pair ORC with OBR by position
    max_orders = max(len(orc_segs), len(obr_segs))
    for i in range(max_orders):
        order = {}
        if i < len(orc_segs):
            orc = orc_segs[i]
            ctrl = msg.get_field(orc, 1)
            order["order_control"] = f"{ctrl} ({ORDER_CONTROL.get(ctrl, '')})" if ctrl else ""
            order["placer_order_num"] = msg.get_component(msg.get_field(orc, 2), 1)
            order["filler_order_num"] = msg.get_component(msg.get_field(orc, 3), 1)
            order["order_status"] = msg.get_field(orc, 5)
            order["order_datetime"] = _format_hl7_datetime(msg.get_field(orc, 9))
            order["ordering_provider"] = _format_xcn(msg, msg.get_field(orc, 12))

        if i < len(obr_segs):
            obr = obr_segs[i]
            if not order.get("placer_order_num"):
                order["placer_order_num"] = msg.get_component(msg.get_field(obr, 2), 1)
            if not order.get("filler_order_num"):
                order["filler_order_num"] = msg.get_component(msg.get_field(obr, 3), 1)

            obr4 = msg.get_field(obr, 4)
            order["service_id"] = msg.get_component(obr4, 1)
            order["service_name"] = msg.get_component(obr4, 2)

            order["observation_datetime"] = _format_hl7_datetime(msg.get_field(obr, 7))
            if not order.get("ordering_provider"):
                order["ordering_provider"] = _format_xcn(msg, msg.get_field(obr, 16))
            order["results_datetime"] = _format_hl7_datetime(msg.get_field(obr, 22))

            status = msg.get_field(obr, 25)
            order["result_status"] = f"{status} ({RESULT_STATUS.get(status, '')})" if status else ""

            obr31 = msg.get_field(obr, 31)
            order["reason"] = msg.get_component(obr31, 2) or msg.get_component(obr31, 1)

        # Fill missing keys
        for key in ["order_control", "placer_order_num", "filler_order_num",
                     "order_status", "order_datetime", "ordering_provider",
                     "service_id", "service_name", "observation_datetime",
                     "results_datetime", "result_status", "reason"]:
            order.setdefault(key, "")

        orders.append(order)
    result["orders"] = orders

    # --- OBX (Observations/Results) ---
    obx_segs = msg.get_segments("OBX")
    observations = []
    for obx in obx_segs:
        obs = {}
        obs["set_id"] = msg.get_field(obx, 1)
        obs["value_type"] = msg.get_field(obx, 2)

        obx3 = msg.get_field(obx, 3)
        obs["observation_id"] = msg.get_component(obx3, 1)
        obs["observation_name"] = msg.get_component(obx3, 2)
        obs["observation_sub_id"] = msg.get_field(obx, 4)

        # OBX-5: Value (may have repetitions)
        obx5 = msg.get_field(obx, 5)
        # For coded entries, get the text component
        if obs["value_type"] in ("CE", "CWE", "CNE"):
            obs["observation_value"] = msg.get_component(obx5, 2) or msg.get_component(obx5, 1)
        else:
            obs["observation_value"] = obx5

        obx6 = msg.get_field(obx, 6)
        obs["units"] = msg.get_component(obx6, 1)
        obs["reference_range"] = msg.get_field(obx, 7)
        obs["abnormal_flags"] = msg.get_field(obx, 8)

        status = msg.get_field(obx, 11)
        obs["result_status"] = RESULT_STATUS.get(status, status)

        obs["observation_datetime"] = _format_hl7_datetime(msg.get_field(obx, 14))
        observations.append(obs)
    result["observations"] = observations

    # --- DG1 (Diagnoses) ---
    dg1_segs = msg.get_segments("DG1")
    diagnoses = []
    for dg1 in dg1_segs:
        dx = {}
        dx["set_id"] = msg.get_field(dg1, 1)
        dg1_3 = msg.get_field(dg1, 3)
        dx["diagnosis_code"] = msg.get_component(dg1_3, 1)
        dx["diagnosis_description"] = msg.get_component(dg1_3, 2) or msg.get_field(dg1, 4)
        dx["diagnosis_datetime"] = _format_hl7_datetime(msg.get_field(dg1, 5))
        dx_type = msg.get_field(dg1, 6)
        dx["diagnosis_type"] = DX_TYPE.get(dx_type, dx_type)
        diagnoses.append(dx)
    result["diagnoses"] = diagnoses

    # --- IN1 (Insurance) ---
    in1_segs = msg.get_segments("IN1")
    insurance = []
    for in1 in in1_segs:
        ins = {}
        ins["set_id"] = msg.get_field(in1, 1)
        in1_2 = msg.get_field(in1, 2)
        ins["plan_id"] = msg.get_component(in1_2, 1)
        ins["plan_name"] = msg.get_component(in1_2, 2)

        in1_3 = msg.get_field(in1, 3)
        ins["company_id"] = msg.get_component(in1_3, 1)

        in1_4 = msg.get_field(in1, 4)
        ins["company_name"] = msg.get_component(in1_4, 1) or in1_4

        ins["group_number"] = msg.get_field(in1, 8)
        ins["group_name"] = msg.get_component(msg.get_field(in1, 9), 1) or msg.get_field(in1, 9)
        ins["plan_effective_date"] = _format_hl7_date(msg.get_field(in1, 12))
        ins["plan_expiration_date"] = _format_hl7_date(msg.get_field(in1, 13))

        in1_16 = msg.get_field(in1, 16)
        ins["insured_name"] = _format_xpn(msg, in1_16)

        ins["policy_number"] = msg.get_field(in1, 36)
        insurance.append(ins)
    result["insurance"] = insurance

    # --- AL1 (Allergies) ---
    al1_segs = msg.get_segments("AL1")
    allergies = []
    for al1 in al1_segs:
        allergy = {}
        allergy["set_id"] = msg.get_field(al1, 1)

        type_code = msg.get_field(al1, 2)
        allergy["allergen_type"] = {"DA": "Drug", "FA": "Food", "EA": "Environmental",
                                     "MA": "Miscellaneous", "LA": "Pollen",
                                     "AA": "Animal"}.get(type_code, type_code)

        al1_3 = msg.get_field(al1, 3)
        allergy["allergen"] = msg.get_component(al1_3, 2) or msg.get_component(al1_3, 1)

        severity = msg.get_field(al1, 4)
        allergy["severity"] = {"SV": "Severe", "MO": "Moderate", "MI": "Mild",
                                "U": "Unknown"}.get(severity, severity)

        allergy["reaction"] = msg.get_field(al1, 5)
        allergies.append(allergy)
    result["allergies"] = allergies

    # --- FT1 (Financial Transactions) ---
    ft1_segs = msg.get_segments("FT1")
    financials = []
    for ft1 in ft1_segs:
        fin = {}
        fin["set_id"] = msg.get_field(ft1, 1)
        fin["transaction_date"] = _format_hl7_datetime(msg.get_field(ft1, 4))
        fin["transaction_type"] = msg.get_field(ft1, 6)

        ft1_7 = msg.get_field(ft1, 7)
        fin["transaction_code"] = msg.get_component(ft1_7, 1)
        fin["transaction_description"] = msg.get_component(ft1_7, 2)

        fin["quantity"] = msg.get_field(ft1, 10)
        fin["amount_extended"] = _safe_float(msg.get_field(ft1, 11))
        fin["amount_unit"] = _safe_float(msg.get_field(ft1, 12))

        ft1_25 = msg.get_field(ft1, 25)
        fin["procedure_code"] = msg.get_component(ft1_25, 1)
        fin["procedure_description"] = msg.get_component(ft1_25, 2)

        ft1_19 = msg.get_field(ft1, 19)
        fin["diagnosis_code"] = msg.get_component(ft1_19, 1)
        financials.append(fin)
    result["financials"] = financials

    # --- SCH (Scheduling) ---
    sch_segs = msg.get_segments("SCH")
    schedules = []
    for sch in sch_segs:
        sched = {}
        sched["placer_appt_id"] = msg.get_component(msg.get_field(sch, 1), 1)
        sched["filler_appt_id"] = msg.get_component(msg.get_field(sch, 2), 1)

        sch7 = msg.get_field(sch, 7)
        sched["appointment_reason"] = msg.get_component(sch7, 2) or msg.get_component(sch7, 1)

        sch8 = msg.get_field(sch, 8)
        sched["appointment_type"] = msg.get_component(sch8, 2) or msg.get_component(sch8, 1)

        # SCH-11: Timing (start^end^duration)
        sch11 = msg.get_field(sch, 11)
        reps = msg.get_repetitions(sch11)
        if reps:
            first_rep = reps[0]
            sched["start_datetime"] = _format_hl7_datetime(msg.get_component(first_rep, 4))
            sched["end_datetime"] = _format_hl7_datetime(msg.get_component(first_rep, 5)) if msg.get_component(first_rep, 5) else ""
            sched["duration"] = msg.get_component(first_rep, 3)

        sched["filler_status"] = msg.get_field(sch, 25)
        schedules.append(sched)
    result["schedules"] = schedules

    return result


def _build_sheets(parsed_msgs):
    """Build Excel sheet data from parsed messages."""
    sheets = []

    # --- Messages sheet (always present) ---
    msg_headers = [
        "Message Type", "Trigger Event", "Message Date/Time",
        "Message Control ID", "Version",
        "Sending Application", "Sending Facility",
        "Receiving Application", "Receiving Facility",
    ]
    msg_rows = []
    for m in parsed_msgs:
        msg_rows.append([
            m.get("message_type", ""),
            m.get("trigger_event", ""),
            m.get("message_datetime", ""),
            m.get("message_control_id", ""),
            m.get("version", ""),
            m.get("sending_app", ""),
            m.get("sending_facility", ""),
            m.get("receiving_app", ""),
            m.get("receiving_facility", ""),
        ])
    sheets.append({"name": "HL7 Messages", "headers": msg_headers, "rows": msg_rows, "currency_cols": []})

    # --- Patients sheet ---
    all_patients = []
    for m in parsed_msgs:
        mcid = m.get("message_control_id", "")
        for p in m.get("patients", []):
            all_patients.append([mcid] + [
                p.get("patient_name", ""),
                p.get("patient_id", ""),
                p.get("patient_ids", ""),
                p.get("dob", ""),
                p.get("gender", ""),
                p.get("address", ""),
                p.get("home_phone", ""),
                p.get("account_number", ""),
                p.get("ssn", ""),
            ])
    if all_patients:
        sheets.append({
            "name": "HL7 Patients",
            "headers": ["Message Control ID", "Patient Name", "Patient ID",
                         "All Patient IDs", "Date of Birth", "Gender",
                         "Address", "Home Phone", "Account Number", "SSN"],
            "rows": all_patients,
            "currency_cols": [],
        })

    # --- Visits sheet ---
    all_visits = []
    for m in parsed_msgs:
        mcid = m.get("message_control_id", "")
        for v in m.get("visits", []):
            all_visits.append([mcid] + [
                v.get("patient_class", ""),
                v.get("location", ""),
                v.get("room", ""),
                v.get("bed", ""),
                v.get("attending_doctor", ""),
                v.get("referring_doctor", ""),
                v.get("admitting_doctor", ""),
                v.get("hospital_service", ""),
                v.get("visit_number", ""),
                v.get("admit_date", ""),
                v.get("discharge_date", ""),
                v.get("discharge_disposition", ""),
                v.get("admit_source", ""),
            ])
    if all_visits:
        sheets.append({
            "name": "HL7 Visits",
            "headers": ["Message Control ID", "Patient Class", "Location",
                         "Room", "Bed", "Attending Doctor", "Referring Doctor",
                         "Admitting Doctor", "Hospital Service", "Visit Number",
                         "Admit Date", "Discharge Date", "Discharge Disposition",
                         "Admit Source"],
            "rows": all_visits,
            "currency_cols": [],
        })

    # --- Orders sheet ---
    all_orders = []
    for m in parsed_msgs:
        mcid = m.get("message_control_id", "")
        for o in m.get("orders", []):
            all_orders.append([mcid] + [
                o.get("order_control", ""),
                o.get("placer_order_num", ""),
                o.get("filler_order_num", ""),
                o.get("order_status", ""),
                o.get("service_id", ""),
                o.get("service_name", ""),
                o.get("ordering_provider", ""),
                o.get("order_datetime", ""),
                o.get("observation_datetime", ""),
                o.get("results_datetime", ""),
                o.get("result_status", ""),
                o.get("reason", ""),
            ])
    if all_orders:
        sheets.append({
            "name": "HL7 Orders",
            "headers": ["Message Control ID", "Order Control", "Placer Order #",
                         "Filler Order #", "Order Status", "Service ID",
                         "Service Name", "Ordering Provider", "Order Date/Time",
                         "Observation Date/Time", "Results Date/Time",
                         "Result Status", "Reason"],
            "rows": all_orders,
            "currency_cols": [],
        })

    # --- Observations/Results sheet ---
    all_obs = []
    for m in parsed_msgs:
        mcid = m.get("message_control_id", "")
        for o in m.get("observations", []):
            all_obs.append([mcid] + [
                o.get("set_id", ""),
                o.get("observation_id", ""),
                o.get("observation_name", ""),
                o.get("observation_value", ""),
                o.get("units", ""),
                o.get("reference_range", ""),
                o.get("abnormal_flags", ""),
                o.get("result_status", ""),
                o.get("observation_datetime", ""),
                o.get("value_type", ""),
            ])
    if all_obs:
        sheets.append({
            "name": "HL7 Results",
            "headers": ["Message Control ID", "Set ID", "Observation ID",
                         "Observation Name", "Value", "Units",
                         "Reference Range", "Abnormal Flags", "Result Status",
                         "Observation Date/Time", "Value Type"],
            "rows": all_obs,
            "currency_cols": [],
        })

    # --- Diagnoses sheet ---
    all_dx = []
    for m in parsed_msgs:
        mcid = m.get("message_control_id", "")
        for d in m.get("diagnoses", []):
            all_dx.append([mcid] + [
                d.get("set_id", ""),
                d.get("diagnosis_code", ""),
                d.get("diagnosis_description", ""),
                d.get("diagnosis_type", ""),
                d.get("diagnosis_datetime", ""),
            ])
    if all_dx:
        sheets.append({
            "name": "HL7 Diagnoses",
            "headers": ["Message Control ID", "Set ID", "Diagnosis Code",
                         "Diagnosis Description", "Diagnosis Type", "Diagnosis Date"],
            "rows": all_dx,
            "currency_cols": [],
        })

    # --- Insurance sheet ---
    all_ins = []
    for m in parsed_msgs:
        mcid = m.get("message_control_id", "")
        for i in m.get("insurance", []):
            all_ins.append([mcid] + [
                i.get("set_id", ""),
                i.get("plan_id", ""),
                i.get("plan_name", ""),
                i.get("company_id", ""),
                i.get("company_name", ""),
                i.get("group_number", ""),
                i.get("group_name", ""),
                i.get("insured_name", ""),
                i.get("policy_number", ""),
                i.get("plan_effective_date", ""),
                i.get("plan_expiration_date", ""),
            ])
    if all_ins:
        sheets.append({
            "name": "HL7 Insurance",
            "headers": ["Message Control ID", "Set ID", "Plan ID", "Plan Name",
                         "Company ID", "Company Name", "Group Number", "Group Name",
                         "Insured Name", "Policy Number",
                         "Plan Effective Date", "Plan Expiration Date"],
            "rows": all_ins,
            "currency_cols": [],
        })

    # --- Allergies sheet ---
    all_al = []
    for m in parsed_msgs:
        mcid = m.get("message_control_id", "")
        for a in m.get("allergies", []):
            all_al.append([mcid] + [
                a.get("set_id", ""),
                a.get("allergen_type", ""),
                a.get("allergen", ""),
                a.get("severity", ""),
                a.get("reaction", ""),
            ])
    if all_al:
        sheets.append({
            "name": "HL7 Allergies",
            "headers": ["Message Control ID", "Set ID", "Allergen Type",
                         "Allergen", "Severity", "Reaction"],
            "rows": all_al,
            "currency_cols": [],
        })

    # --- Financial Transactions sheet ---
    all_fin = []
    for m in parsed_msgs:
        mcid = m.get("message_control_id", "")
        for f in m.get("financials", []):
            all_fin.append([mcid] + [
                f.get("transaction_date", ""),
                f.get("transaction_type", ""),
                f.get("transaction_code", ""),
                f.get("transaction_description", ""),
                f.get("quantity", ""),
                f.get("amount_extended", ""),
                f.get("amount_unit", ""),
                f.get("procedure_code", ""),
                f.get("procedure_description", ""),
                f.get("diagnosis_code", ""),
            ])
    if all_fin:
        sheets.append({
            "name": "HL7 Financial",
            "headers": ["Message Control ID", "Transaction Date", "Transaction Type",
                         "Transaction Code", "Description", "Quantity",
                         "Amount (Extended)", "Amount (Unit)",
                         "Procedure Code", "Procedure Description", "Diagnosis Code"],
            "rows": all_fin,
            "currency_cols": [7, 8],
        })

    # --- Scheduling sheet ---
    all_sch = []
    for m in parsed_msgs:
        mcid = m.get("message_control_id", "")
        for s in m.get("schedules", []):
            all_sch.append([mcid] + [
                s.get("placer_appt_id", ""),
                s.get("filler_appt_id", ""),
                s.get("appointment_reason", ""),
                s.get("appointment_type", ""),
                s.get("start_datetime", ""),
                s.get("end_datetime", ""),
                s.get("duration", ""),
                s.get("filler_status", ""),
            ])
    if all_sch:
        sheets.append({
            "name": "HL7 Scheduling",
            "headers": ["Message Control ID", "Placer Appointment ID",
                         "Filler Appointment ID", "Appointment Reason",
                         "Appointment Type", "Start Date/Time", "End Date/Time",
                         "Duration", "Filler Status"],
            "rows": all_sch,
            "currency_cols": [],
        })

    return sheets


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _format_hl7_date(val):
    """Convert YYYYMMDD to MM/DD/YYYY."""
    if not val or len(val) < 8:
        return val or ""
    return f"{val[4:6]}/{val[6:8]}/{val[0:4]}"


def _format_hl7_datetime(val):
    """Convert YYYYMMDDHHMMSS to MM/DD/YYYY HH:MM:SS."""
    if not val:
        return ""
    val = val.split("+")[0].split("-")[0]  # Strip timezone offset
    if len(val) >= 12:
        return f"{val[4:6]}/{val[6:8]}/{val[0:4]} {val[8:10]}:{val[10:12]}"
    elif len(val) >= 8:
        return f"{val[4:6]}/{val[6:8]}/{val[0:4]}"
    return val


def _format_xcn(msg, field_val):
    """Format an XCN (Extended Composite ID Number and Name) field."""
    if not field_val:
        return ""
    id_num = msg.get_component(field_val, 1)
    last = msg.get_component(field_val, 2)
    first = msg.get_component(field_val, 3)
    if last and first:
        name = f"{last}, {first}"
    elif last:
        name = last
    else:
        name = id_num
    return name


def _format_xpn(msg, field_val):
    """Format an XPN (Extended Person Name) field."""
    if not field_val:
        return ""
    last = msg.get_component(field_val, 1)
    first = msg.get_component(field_val, 2)
    middle = msg.get_component(field_val, 3)
    name = f"{last}, {first}" if first else last
    if middle:
        name += f" {middle}"
    return name


def _safe_float(val):
    """Convert to float or return empty string."""
    try:
        return float(val) if val else ""
    except (ValueError, TypeError):
        return val or ""
