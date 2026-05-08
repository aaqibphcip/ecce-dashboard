from flask import Flask, jsonify, send_file
from flask_cors import CORS
import openpyxl

app = Flask(__name__)
CORS(app)

# -----------------------------------------------
EXCEL_FILE = "23834 schools.xlsx"   # your Excel filename
SHEET_NAME = "Report"               # sheet with the district summary
# -----------------------------------------------

@app.route("/")
def home():
    return send_file("ecce_punjab_dashboard.html")

@app.route("/data")
def get_data():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        ws = wb[SHEET_NAME]

        data = []
        header_row = None
        header_col = None

        # Find the header row by scanning for "District"
        for row in ws.iter_rows(values_only=True):
            for i, cell in enumerate(row):
                if str(cell).strip() == "District":
                    header_row = list(row)
                    header_col = i  # column index where District is found
                    break
            if header_row:
                break

        if not header_row:
            return jsonify({"error": "Could not find District header in Report sheet"}), 400

        # Extract headers starting from the District column
        headers = [str(h).strip() if h is not None else "" for h in header_row[header_col:]]

        # Now read all data rows
        found_header = False
        for row in ws.iter_rows(values_only=True):
            row_slice = row[header_col:]

            # Skip until we pass the header row
            if not found_header:
                if str(row_slice[0]).strip() == "District":
                    found_header = True
                continue

            # Stop at empty or Total row
            if not any(row_slice):
                continue
            if str(row_slice[0]).strip().lower() in ('total', 'grand total', 'totals', '', 'none'):
                continue

            record = dict(zip(headers, row_slice))
            data.append(record)

        return jsonify(data)

    except FileNotFoundError:
        return jsonify({"error": f"File '{EXCEL_FILE}' not found. Make sure it is in the same folder as server.py"}), 404
    except KeyError:
        return jsonify({"error": f"Sheet '{SHEET_NAME}' not found."}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("=" * 55)
    print("  ECCE Dashboard Server")
    print(f"  Reading: {EXCEL_FILE}  |  Sheet: {SHEET_NAME}")
    print("  Open your browser and go to:")
    print("  http://localhost:5000")
    print("  Press Ctrl+C to stop the server.")
    print("=" * 55)
    app.run(port=5000, debug=False)
