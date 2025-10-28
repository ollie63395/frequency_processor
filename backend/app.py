from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import re

app = Flask(__name__)
CORS(app)  # allow requests from browser (localhost)

# limit upload size to 100MB
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

def excel_col_to_index(col_letter: str) -> int:
    """
    Convert Excel column letters (A, B, ..., Z, AA, AB, ...) to 0-based index.
    A -> 0, B -> 1, ..., Z -> 25, AA -> 26, etc.
    """
    col_letter = col_letter.strip().upper()
    total = 0
    for ch in col_letter:
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"Invalid column letter: {col_letter}")
        total = total * 26 + (ord(ch) - ord('A') + 1)
    return total - 1  # make it 0-based

@app.route("/analyze", methods=["POST"])
def analyze():
    """
    Expects multipart/form-data with:
      - file: the uploaded Excel file
      - columns: a string like "F, G, H, I"
    Returns JSON:
      {
        "F": {"a": 2, "b": 1},
        "G": {...},
        ...
      }
    """
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    excel_file = request.files["file"]
    columns_raw = request.form.get("columns", "")

    # Read entire sheet with NO header, so row0 = Excel row1
    # We'll treat row0 as header and ignore it manually.
    try:
        # figure out which columns to read first
        col_letters = re.split(r"[,\s;]+", columns_raw.strip())
        col_letters = [c for c in col_letters if c]
        col_indexes = [excel_col_to_index(c) for c in col_letters]

        df = pd.read_excel(
            excel_file,
            header=None,
            dtype=str,
            usecols=col_indexes,  # only load required columns
            engine="openpyxl"
        )
    except Exception as e:
        return jsonify({"error": f"Failed to read Excel: {e}"}), 400

    # Clean column list like "F, G, H" -> ["F","G","H"]
    # allow spaces, semicolons, etc.
    col_letters = re.split(r"[,\s;]+", columns_raw.strip())
    col_letters = [c for c in col_letters if c]  # remove empty

    result = {}

    for letter in col_letters:
        try:
            col_idx = excel_col_to_index(letter)
        except ValueError as ve:
            result[letter] = {"_error": str(ve)}
            continue

        if col_idx >= df.shape[1]:
            result[letter] = {"_error": f"Column {letter} not in file"}
            continue

        # take rows starting from index 1 to skip the header row (Excel row1)
        col_series = df.iloc[1:, col_idx]

        # drop NaN/None and strip whitespace
        col_series = col_series.dropna().map(lambda x: str(x).strip())

        # count frequencies
        freqs = col_series.value_counts().to_dict()

        result[letter] = freqs

    return jsonify(result), 200

@app.route("/")
def serve_frontend():
    return send_from_directory("static", "index.html")

# add catch-all route to serve other static files
@app.route("/<path:path>")
def serve_static_files(path):
    return send_from_directory("static", path)

# add a health check to confirm Flask is alive on Render
@app.route("/health")
def health():
    return "OK", 200

if __name__ == "__main__":
    # Flask dev server on http://localhost:5000
    app.run(host="0.0.0.0", port=5000, debug=True)
