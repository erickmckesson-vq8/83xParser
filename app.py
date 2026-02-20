"""Flask web app for 83x Parser — drag-and-drop EDI 835/837 to Excel converter."""

import os
import uuid
import tempfile

from flask import Flask, render_template, request, send_file, jsonify

from edi_parser import EDIFile
from parser_835 import parse_835
from parser_837 import parse_837
from excel_writer import write_combined_excel

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB limit

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("files")
    if not files or all(not f.filename for f in files):
        return jsonify({"error": "No files provided"}), 400

    all_835 = []  # list of (filename, parsed_transactions)
    all_837 = []
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

        # Parse EDI
        try:
            edi = EDIFile(content)
        except ValueError as e:
            errors.append(f"{file.filename}: {e}")
            continue
        except Exception as e:
            errors.append(f"{file.filename}: Failed to parse — {e}")
            continue

        txn_type = edi.get_transaction_type()
        if txn_type == "835":
            parsed = parse_835(edi)
            all_835.append((file.filename, parsed))
        elif txn_type == "837":
            parsed = parse_837(edi)
            all_837.append((file.filename, parsed))
        else:
            errors.append(f"{file.filename}: Unsupported transaction type "
                          f"'{txn_type or 'unknown'}' (expected 835 or 837)")

    if not all_835 and not all_837:
        msg = "No valid 835/837 files found."
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
        write_combined_excel(all_835, all_837, output_path)
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


if __name__ == "__main__":
    print("=" * 50)
    print("  83x Parser — EDI 835/837 to Excel")
    print("  Open http://127.0.0.1:5000 in your browser")
    print("=" * 50)
    app.run(debug=True, port=5000)
