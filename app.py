import os
from flask import Flask, request, send_file
import pandas as pd
import io

app = Flask(__name__)

@app.route('/', methods=['GET'])  # Health check route
def home():
    return "CSV/XLSX Processing API is running!"

@app.route('/process', methods=['POST'])  # Main API endpoint
def process_file():
    try:
        if 'file' not in request.files:
            return {"error": "No file part in the request"}, 400

        uploaded_file = request.files['file']

        if uploaded_file.filename == '':
            return {"error": "No selected file"}, 400

        filename = uploaded_file.filename.lower()

        # Step 1: Read into DataFrame depending on file type
        if filename.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        elif filename.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        else:
            return {"error": "Unsupported file format. Please upload a .csv or .xlsx file."}, 400

        # Step 2: Check if "Consultores[Mail]" column exists
        if "Consultores[Mail]" not in df.columns:
            return {"error": "Column 'Consultores[Mail]' not found."}, 400

        # Step 3: Get unique values in "Consultores[Mail]"
        unique_consultores = df["Consultores[Mail]"].dropna().unique()

        # Step 4: Create an Excel file with multiple sheets
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for consultor in unique_consultores:
                filtered_data = df[df["Consultores[Mail]"] == consultor]
                if not filtered_data.empty:
                    # Excel sheet names must be <= 31 characters
                    sheet_name = str(consultor)[:31]
                    filtered_data.to_excel(writer, sheet_name=sheet_name, index=False)

        output.seek(0)

        # Step 5: Return the processed XLSX file
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="processed_data.xlsx"
        )

    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
