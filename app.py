"""Flask web app — Healthcare EDI and API Parser."""

import os
import uuid

from flask import Flask, render_template, request, send_file, jsonify

from edi_parser import EDIFile
from format_detect import detect_format, detect_x12_type
from parser_835 import parse_835
from parser_837 import parse_837
from parser_hl7v2 import parse_hl7v2
from parser_fhir import parse_fhir
from parser_cda import parse_cda
from parser_ncpdp import parse_ncpdp
from parser_x12_generic import parse_x12_generic
from excel_writer import write_combined_excel

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB limit

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


FORMAT_LABELS = {
    "x12": "X12/EDI",
    "hl7v2": "HL7 v2.x",
    "fhir": "FHIR",
    "cda": "CDA/HL7v3",
    "ncpdp": "NCPDP",
    "csv": "Delimited Text",
}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("files")
    if not files or all(not f.filename for f in files):
        return jsonify({"error": "No files provided"}), 400

    all_835 = []   # (filename, parsed_transactions)
    all_837 = []   # (filename, parsed_transactions)
    other = []     # (filename, format_name, sheets_list)
    errors = []

    for file in files:
        if not file.filename:
            continue

        # Read file content
        try:
            raw = file.read()
            try:
                content = raw.decode("utf-8")
            except UnicodeDecodeError:
                content = raw.decode("latin-1")
        except Exception as e:
            errors.append(f"{file.filename}: Could not read file — {e}")
            continue

        if not content.strip():
            errors.append(f"{file.filename}: File is empty")
            continue

        # Detect format
        fmt = detect_format(content)

        if fmt is None:
            errors.append(f"{file.filename}: Could not detect file format")
            continue

        try:
            if fmt == "x12":
                _handle_x12(file.filename, content, all_835, all_837, other, errors)
            elif fmt == "hl7v2":
                sheets = parse_hl7v2(content)
                if sheets:
                    other.append((file.filename, "HL7 v2.x", sheets))
                else:
                    errors.append(f"{file.filename}: No HL7 v2 messages found")
            elif fmt == "fhir":
                sheets = parse_fhir(content)
                if sheets:
                    other.append((file.filename, "FHIR", sheets))
                else:
                    errors.append(f"{file.filename}: No FHIR resources found")
            elif fmt == "cda":
                sheets = parse_cda(content)
                if sheets:
                    other.append((file.filename, "CDA", sheets))
                else:
                    errors.append(f"{file.filename}: Could not parse CDA document")
            elif fmt == "ncpdp":
                sheets = parse_ncpdp(content)
                if sheets:
                    other.append((file.filename, "NCPDP", sheets))
                else:
                    errors.append(f"{file.filename}: Could not parse NCPDP data")
            elif fmt == "csv":
                sheets = _parse_csv(content)
                if sheets:
                    other.append((file.filename, "Delimited", sheets))
                else:
                    errors.append(f"{file.filename}: Could not parse delimited data")
            else:
                errors.append(f"{file.filename}: Unsupported format '{fmt}'")
        except Exception as e:
            errors.append(f"{file.filename}: Parse error — {e}")

    if not all_835 and not all_837 and not other:
        msg = "No valid healthcare files found."
        if errors:
            msg += " Errors: " + "; ".join(errors)
        return jsonify({"error": msg}), 400

    # Generate combined Excel
    if len(files) == 1 and files[0].filename:
        base_name = os.path.splitext(files[0].filename)[0]
        output_filename = f"{base_name}_parsed.xlsx"
    else:
        output_filename = "combined_parsed.xlsx"

    output_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}_{output_filename}")

    try:
        write_combined_excel(all_835, all_837, output_path, other_formats=other)
    except Exception as e:
        return jsonify({"error": f"Failed to generate Excel: {e}"}), 500

    response = send_file(
        output_path,
        as_attachment=True,
        download_name=output_filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    @response.call_on_close
    def cleanup():
        try:
            os.remove(output_path)
        except OSError:
            pass

    return response


def _handle_x12(filename, content, all_835, all_837, other, errors):
    """Route X12 files to the appropriate parser."""
    try:
        edi = EDIFile(content)
    except ValueError as e:
        errors.append(f"{filename}: {e}")
        return
    except Exception as e:
        errors.append(f"{filename}: Failed to parse X12 — {e}")
        return

    txn_type = edi.get_transaction_type()

    if txn_type == "835":
        parsed = parse_835(edi)
        all_835.append((filename, parsed))
    elif txn_type == "837":
        parsed = parse_837(edi)
        all_837.append((filename, parsed))
    elif txn_type:
        # Other X12 types (270, 271, 276, 277, 278, 834, etc.)
        sheets = parse_x12_generic(edi)
        if sheets:
            other.append((filename, f"X12 {txn_type}", sheets))
        else:
            errors.append(f"{filename}: No data found in X12 {txn_type}")
    else:
        errors.append(f"{filename}: Could not determine X12 transaction type")


def _parse_csv(content):
    """Parse CSV/TSV/delimited content into sheets."""
    import csv
    import io

    content = content.strip()
    if not content:
        return []

    # Detect delimiter
    sniffer = csv.Sniffer()
    try:
        dialect = sniffer.sniff(content[:4000])
    except csv.Error:
        dialect = csv.excel  # Default to comma

    reader = csv.reader(io.StringIO(content), dialect)
    rows_data = list(reader)

    if not rows_data:
        return []

    # Check if first row looks like headers (non-numeric, unique)
    first_row = rows_data[0]
    has_header = sniffer.has_header(content[:4000]) if len(rows_data) > 1 else True

    if has_header:
        headers = first_row
        rows = rows_data[1:]
    else:
        headers = [f"Column {i+1}" for i in range(len(first_row))]
        rows = rows_data

    # Pad short rows
    max_cols = len(headers)
    padded_rows = []
    for row in rows:
        if len(row) < max_cols:
            row = row + [""] * (max_cols - len(row))
        padded_rows.append(row[:max_cols])

    return [{"name": "Data", "headers": headers, "rows": padded_rows, "currency_cols": []}]


if __name__ == "__main__":
    print("=" * 56)
    print("  Healthcare EDI and API Parser")
    print("  Formats: X12, HL7 v2, FHIR, CDA, NCPDP, CSV")
    print("  Open http://127.0.0.1:5000 in your browser")
    print("=" * 56)
    app.run(debug=True, port=5000)
