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
        segmentation_column = data.get("segmentation_column")

        if not filename or not filedata:
            return {"error": "Missing filename or filedata"}, 400

        if not segmentation_column:
            return {"error": "Missing segmentation_column"}, 400

        # Decode base64 to binary
        file_bytes = base64.b64decode(filedata)
        file_stream = io.BytesIO(file_bytes)

        # Read file based on extension
        if filename.endswith('.csv'):
            df = pd.read_csv(file_stream)
        elif filename.endswith('.xlsx'):
            df = pd.read_excel(file_stream, engine='openpyxl')
        else:
            return {"error": "Unsupported file type"}, 400

        # Validate the segmentation column
        if segmentation_column not in df.columns:
            return {"error": f"Column '{segmentation_column}' not found."}, 400

        unique_values = df[segmentation_column].dropna().unique()

        # Create ZIP archive in memory
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for value in unique_values:
                filtered = df[df[segmentation_column] == value]

                # Create Excel file for each value
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    filtered.to_excel(writer, index=False)

                excel_buffer.seek(0)
                safe_name = str(value).replace('/', '_').replace('@', '_at_')[:50]
                file_name = f"{safe_name}.xlsx"
                zip_file.writestr(file_name, excel_buffer.read())

        zip_buffer.seek(0)

        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='segmented_files.zip'
        )

    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
