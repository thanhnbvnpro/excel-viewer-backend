from flask import Flask, request, jsonify, send_from_directory, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/list-files')
def list_files():
    files = [f for f in os.listdir(UPLOAD_FOLDER) if allowed_file(f)]
    return jsonify(files)

@app.route('/list-sheets')
def list_sheets():
    file = request.args.get('file')
    try:
        file_path = os.path.join(UPLOAD_FOLDER, file)
        xl = pd.ExcelFile(file_path)
        return jsonify(xl.sheet_names)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/excel-data')
def get_excel_data():
    file = request.args.get('file')
    sheet = request.args.get('sheet')
    search = request.args.get('search', '').lower()
    try:
        file_path = os.path.join(UPLOAD_FOLDER, file)
        df = pd.read_excel(file_path, sheet_name=sheet)
        if search:
            df = df[df.apply(lambda row: row.astype(str).str.lower().str.contains(search).any(), axis=1)]
        return df.to_json(orient='records', force_ascii=False)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/export-csv')
def export_csv():
    file = request.args.get('file')
    sheet = request.args.get('sheet')
    try:
        file_path = os.path.join(UPLOAD_FOLDER, file)
        df = pd.read_excel(file_path, sheet_name=sheet)
        temp_csv = os.path.join(UPLOAD_FOLDER, 'export_temp.csv')
        df.to_csv(temp_csv, index=False, encoding='utf-8-sig')
        return send_file(temp_csv, as_attachment=True, download_name='export.csv')
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files['file']
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        save_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(save_path)
        return jsonify({"message": "File uploaded successfully", "filename": filename})
    return jsonify({"error": "Invalid file type"}), 400


@app.route('/')
def home():
    return '✅ Hello from Flask on Render!'

if __name__ == '__main__':
    app.run(host='http://127.0.0.1/excel-viewer-backend/', port=5000)
