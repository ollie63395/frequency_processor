from werkzeug.middleware.proxy_fix import ProxyFix
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import re

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True)

# limit upload size to 100MB
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
app.config["PREFERRED_URL_SCHEME"] = "https"
app.config["UPLOAD_FOLDER"] = "/tmp"
app.config["ENV"] = "production"
app.config["DEBUG"] = False

app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1)

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

    print("File received:", excel_file.filename)
    print("File size (approx):", len(excel_file.read()))
    excel_file.seek(0)

    # Parse column letters
    col_letters = re.split(r"[,\s;]+", columns_raw.strip())
    col_letters = [c.strip().upper() for c in col_letters if c]

    try:
        # Build Excel-style range string, e.g. "F:I" for F–I
        if len(col_letters) == 1:
            usecols_str = col_letters[0]
        else:
            # Pandas understands "F,H,I" or "F:I"
            usecols_str = ",".join(col_letters)

        df = pd.read_excel(
            excel_file,
            header=None,
            dtype=str,
            usecols=usecols_str,   # <<— use Excel letters
            engine="openpyxl"
        )

        print("Shape read:", df.shape)
    except Exception as e:
        return jsonify({"error": f"Failed to read Excel: {e}"}), 400

    # Clean column list like "F, G, H" -> ["F","G","H"]
    # allow spaces, semicolons, etc.
    col_letters = re.split(r"[,\s;]+", columns_raw.strip())
    col_letters = [c for c in col_letters if c]  # remove empty

    result = {}

    for i, letter in enumerate(col_letters):

        # Each letter corresponds to one of the loaded columns, 0..len(usecols)-1
        if i >= df.shape[1]:
            result[letter] = {"_error": f"Column {letter} not found after reading"}
            continue

        # take rows starting from index 1 to skip the header row (Excel row1)
        col_series = df.iloc[1:, i]
        col_series = col_series.dropna().map(lambda x: str(x).strip())
        result[letter] = col_series.value_counts().to_dict()

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
