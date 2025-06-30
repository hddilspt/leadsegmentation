import os
from flask import Flask, request, send_file
import pandas as pd
import io
import base64
import zipfile
import xml.etree.ElementTree as ET
import openpyxl

app = Flask(__name__)

def fallback_parse_xlsx(file_stream):
    """Reads the first worksheet of a broken .xlsx file without relying on styles."""
    file_stream.seek(0)
    with zipfile.ZipFile(file_stream) as z:
        ns = {
            'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

        # Load shared strings
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as s:
                tree = ET.parse(s)
                root = tree.getroot()
                for si in root.findall('.//main:t', ns):
                    shared_strings.append(si.text or '')

        # Get rel ID of first sheet from workbook.xml
        with z.open('xl/workbook.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            sheet_elem = root.find('.//main:sheets/main:sheet', ns)
            if sheet_elem is None:
                raise ValueError("No sheets found in workbook.")
            rel_id = sheet_elem.attrib.get(f'{{{ns["r"]}}}id')

        # Parse workbook.xml.rels WITHOUT namespaces
        rel_target = None
        with z.open('xl/_rels/workbook.xml.rels') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            for rel in root.findall('Relationship'):
                if rel.attrib.get('Id') == rel_id:
                    rel_target = rel.attrib.get('Target')
                    break

        if not rel_target:
            raise ValueError(f"Could not find worksheet target from rels (rel_id: {rel_id}).")

        # Normalize and clean path
        if rel_target.startswith('/'):
            sheet_path = rel_target[1:]
        else:
            sheet_path = os.path.normpath(f"xl/{rel_target}").replace("\\", "/")

        # Read and parse worksheet
        with z.open(sheet_path) as s:
            tree = ET.parse(s)
            root = tree.getroot()
            rows = []
            for row in root.findall('.//main:row', ns):
                values = []
                for c in row.findall('main:c', ns):
                    value = ''
                    cell_type = c.attrib.get('t')
                    v = c.find('main:v', ns)
                    if v is not None:
                        if cell_type == 's':
                            idx = int(v.text)
                            value = shared_strings[idx] if idx < len(shared_strings) else ''
                        else:
                            value = v.text
                    values.append(value)
                rows.append(values)

    if not rows:
        raise ValueError("No data found in Excel sheet.")
    return pd.DataFrame(rows[1:], columns=rows[0])


def robust_read_excel(file_stream):
    """Safely read Excel files with fallback if styles cause a crash."""
    try:
        return pd.read_excel(file_stream, engine='openpyxl')
    except IndexError as e:
        if 'list index out of range' in str(e):
            return fallback_parse_xlsx(file_stream)
        else:
            raise


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
            df = robust_read_excel(file_stream)
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
