import os
from flask import Flask, request, send_file
import pandas as pd
import io
import base64

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process_file():
    try:
        data = request.get_json()

        filename = data.get("filename")
        filedata = data.get("filedata")

        if not filename or not filedata:
            return {"error": "Missing filename or filedata"}, 400

        # Decode base64 to binary
        file_bytes = base64.b64decode(filedata)
        file_stream = io.BytesIO(file_bytes)

        # Choose read method based on extension
        if filename.endswith('.csv'):
            df = pd.read_csv(file_stream)
        elif filename.endswith('.xlsx'):
            df = pd.read_excel(file_stream, engine='openpyxl')
        else:
            return {"error": "Unsupported file type"}, 400

        # Check required column
        if "Consultores[Mail]" not in df.columns:
            return {"error": "Column 'Consultores[Mail]' not found."}, 400

        # Process
        unique_consultores = df["Consultores[Mail]"].dropna().unique()
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for consultor in unique_consultores:
                sheet_name = str(consultor)[:31]
                filtered = df[df["Consultores[Mail]"] == consultor]
                filtered.to_excel(writer, sheet_name=sheet_name, index=False)

        output.seek(0)

        return send_file(output,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True,
                         download_name="processed_data.xlsx")

    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
