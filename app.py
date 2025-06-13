import os
from flask import Flask, request, send_file
import pandas as pd
import io
import base64
import zipfile

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

        # Read based on file extension
        if filename.endswith('.csv'):
            df = pd.read_csv(file_stream)
        elif filename.endswith('.xlsx'):
            df = pd.read_excel(file_stream, engine='openpyxl')
        else:
            return {"error": "Unsupported file type"}, 400

        # Check required column
        if "Consultores[Mail]" not in df.columns:
            return {"error": "Column 'Consultores[Mail]' not found."}, 400

        unique_consultores = df["Consultores[Mail]"].dropna().unique()

        # Create ZIP archive in memory
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for consultor in unique_consultores:
                filtered = df[df["Consultores[Mail]"] == consultor]

                # Create Excel file for each consultor
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    filtered.to_excel(writer, index=False)

                excel_buffer.seek(0)
                # Safe filename: remove problematic characters
                safe_filename = f"{str(consultor).replace('/', '_')[:50]}.xlsx"
                zip_file.writestr(safe_filename, excel_buffer.read())

        zip_buffer.seek(0)

        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='consultores_files.zip'
        )

    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
