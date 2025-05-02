import os
from flask import Flask, request, send_file
import pandas as pd
import io

app = Flask(__name__)

@app.route('/', methods=['GET'])  # Health check route
def home():
    return "CSV Processing API is running!"

@app.route('/process', methods=['POST'])  # Main API endpoint
def process_csv_to_xlsx():
    try:
        if 'file' not in request.files:
            return {"error": "No file part in the request"}, 400

        csv_file = request.files['file']

        if csv_file.filename == '':
            return {"error": "No selected file"}, 400

        # Step 1: Convert CSV to DataFrame
        df = pd.read_csv(csv_file)

        # Step 2: Check if "Oportunidades[Responsavel]" column exists
        if "Oportunidades[Responsavel]" not in df.columns:
            return {"error": "Column 'Oportunidades[Responsavel]' not found."}, 400
        
        # Step 3: Get unique "Responsavel" values
        unique_responsaveis = df["Oportunidades[Responsavel]"].dropna().unique()
        
        # Step 4: Create an Excel file with multiple sheets
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for responsavel in unique_responsaveis:
                # Filter data per "Responsavel"
                filtered_data = df[df["Oportunidades[Responsavel]"] == responsavel]
                if not filtered_data.empty:
                    filtered_data.to_excel(writer, sheet_name=str(responsavel), index=False)
        
        output.seek(0)

        # Step 5: Return the processed XLSX file
        return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name="processed_data.xlsx")

    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))  # Use Railway's assigned port
    app.run(host='0.0.0.0', port=port)
