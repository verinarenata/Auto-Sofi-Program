from flask import Flask, request, render_template, jsonify, send_from_directory, send_file
import os
import tempfile
import pandas as pd
import win32com.client as win32
import re
import uuid

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['GENERATED_EXCEL_FOLDER'] = 'generated_files_sofi_excel'
app.config['GENERATED_PDF_FOLDER'] = 'generated_files_sofi_pdf'

# Ensure the directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_EXCEL_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_PDF_FOLDER'], exist_ok=True)

uploaded_file_path = None

@app.route('/')
def index():
    # return render_template('index.html')
    return send_file('index.html')

# Route to serve CSS files
# @app.route('/templates/<path:path>')
# def serve_templates(path):
#     return send_from_directory('templates', path)

@app.route('/upload', methods=['POST'])
def upload_file():
    global uploaded_file_path
    if 'file' not in request.files:
        return jsonify(message='No file part'), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify(message='No selected file'), 400
    if file:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        uploaded_file_path = filepath
        return jsonify(message='File uploaded successfully!'), 200

@app.route('/generate_sofi', methods=['POST'])
def generate_sofi():
    global uploaded_file_path
    if not uploaded_file_path:
        return jsonify(message='No file uploaded'), 400

    df = pd.read_excel(uploaded_file_path, sheet_name="Form Office")
    df.columns = df.columns.str.replace('\n', '')

    import pythoncom
    pythoncom.CoInitialize()

    for i in df.columns:
        if 'SOFI : 1.' in i:
            df.rename(columns={i: "SOFI 1"}, inplace=True)
        if 'SOFI : 2.' in i:
            df.rename(columns={i: "SOFI 2"}, inplace=True)
        if 'KATEGORI INOVASI' in i:
            df.rename(columns={i: "KATEGORI INOVASI"}, inplace=True)

    # df['SOFI 1'] = df['SOFI 1'] + '\n'
    # df['SOFI 2'] = df['SOFI 2'] + '\n'
    # df['SOFI 1'].fillna("", inplace=True)
    # df['SOFI 2'].fillna("", inplace=True)
    # df['Aplikasi selain Microsoft Office 365'].fillna("", inplace=True)
    # df['Microsoft Power Apps'].fillna("", inplace=True)
    # df['Microsoft Office 365 Basic'].fillna("", inplace=True)

    df['SOFI 1'] = df['SOFI 1'].fillna("")
    df['SOFI 2'] = df['SOFI 2'].fillna("")
    df['Aplikasi selain Microsoft Office 365'] = df['Aplikasi selain Microsoft Office 365'].fillna("")
    df['Microsoft Power Apps'] = df['Microsoft Power Apps'].fillna("")
    df['Microsoft Office 365 Basic'] = df['Microsoft Office 365 Basic'].fillna("")

    def gabung_kolom(row):
        if row['Aplikasi selain Microsoft Office 365'] != "":
            return row['Aplikasi selain Microsoft Office 365']
        elif row['Microsoft Power Apps'] != "":
            return row['Microsoft Power Apps']
        elif row['Microsoft Office 365 Basic'] != "":
            return row['Microsoft Office 365 Basic']
        else:
            return ''
    
    df['NAMA TIM'] = df.apply(gabung_kolom, axis=1)

    df_grouped = df.groupby('NAMA TIM')

    excel = win32.Dispatch("Excel.Application")
    excel.DisplayAlerts = False

    file_links = []

    try:
        for nama_tim, df_i in df_grouped:
            sheets = excel.Workbooks.Open(os.path.join(os.getcwd(), 'SOFI_PRESENTATION_NEW.xlsx'))
            sheet = sheets.Worksheets[1]

            sofi_1_filtered = [item for item in df_i['SOFI 1'].sum().split("\n") if len(item) > 1]
            sofi_2_filtered = [item for item in df_i['SOFI 2'].sum().split("\n") if len(item) > 1]
            
            jumlah_penilaian_juri = df_grouped.size()[nama_tim]

            sheet.Range('D3').value = " : " + nama_tim
            sheet.Range('D4').value = " : " + df_i['JUDUL INOVASI'][df_i['JUDUL INOVASI'].index[0]]

            sheet.Range('I9').value = df_i['PENILAIAN (PLAN) : PENETAPAN AKTIFITAS'].sum() / jumlah_penilaian_juri
            sheet.Range('I10').value = df_i['PENILAIAN (PLAN) : PROSES PEMECAHAN MASALAH & PERBAIKAN'].sum() / jumlah_penilaian_juri
            sheet.Range('I11').value = df_i['PENILAIAN (PLAN) : SOLUSI'].sum() / jumlah_penilaian_juri
            sheet.Range('I12').value = df_i['PENILAIAN (DO) : TINGKAT KESULITAN'].sum() / jumlah_penilaian_juri
            sheet.Range('I13').value = df_i['PENILAIAN (DO) : MUTU PROSES PELAKSANAAN'].sum() / jumlah_penilaian_juri
            sheet.Range('I14').value = df_i['PENILAIAN (CHECK) : KETEPATAN & KELENGKAPAN EVALUASI'].sum() / jumlah_penilaian_juri
            sheet.Range('I15').value = df_i['PENILAIAN (CHECK) : DAMPAK HASIL'].sum() / jumlah_penilaian_juri
            sheet.Range('I16').value = df_i['PENILAIAN (ACTION) : STANDARISASI'].sum() / jumlah_penilaian_juri
            sheet.Range('I17').value = df_i['PENILAIAN  : MUTU MAKALAH'].sum() / jumlah_penilaian_juri
            sheet.Range('I18').value = df_i['PENILAIAN  : MUTU PRESENTASI'].sum() / jumlah_penilaian_juri

            if sheet.Range('I19').value >= 800:
                sheet.Range('D5').value = " : LOLOS KE TAHAP PRESENTASI"
            else:
                sheet.Range('D5').value = " : TIDAK LOLOS KE TAHAP PRESENTASI"

            num_row_to_add = len(sofi_1_filtered)
            row_index = 23

            if num_row_to_add > 0:
                sheet.Range('A23').value = 'a.'
                sheet.Range('C23').value = sofi_1_filtered[0]
                for i in range(1, num_row_to_add):
                    row_now = row_index + i
                    sheet.Rows(row_now).Insert()
                    sheet.Range(f"C{row_now}:I{row_now}").Merge()
                    sheet.Range(f"A{row_now}").value = f"{chr(97+i)}."
                    sheet.Range(f"C{row_now}").value = sofi_1_filtered[i]

            num_row_to_add_2 = len(sofi_2_filtered)

            if num_row_to_add > 0:
                row_index_2 = 25 + num_row_to_add - 1
            else:
                row_index_2 = 25 + num_row_to_add

            if num_row_to_add_2 > 0:
                sheet.Range(f'A{row_index_2}').value = 'a.'
                sheet.Range(f'C{row_index_2}').value = sofi_2_filtered[0]
                for i in range(1, num_row_to_add_2):
                    row_now = row_index_2 + i
                    sheet.Rows(row_now).Insert()
                    sheet.Range(f"C{row_now}:I{row_now}").Merge()
                    sheet.Range(f"A{row_now}").value = f"{chr(97+i)}."
                    sheet.Range(f"C{row_now}").value = sofi_2_filtered[i]

            row_index_3 = 27 + num_row_to_add + num_row_to_add_2 - 2
            sheet.Range(f'C{row_index_3}').value = df_i['BENEFIT ( REAL / POTENSIAL ) - FINANCIAL ATAU NON FINANCIAL'][df_i['BENEFIT ( REAL / POTENSIAL ) - FINANCIAL ATAU NON FINANCIAL'].index[0]]

            #unique_id = str(uuid.uuid4())
            downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
            excel_filename = f"SOFI_{nama_tim}.xlsx"
            pdf_filename = f"SOFI_{nama_tim}.pdf"

            # Create SOFI_PDF and SOFI_EXCEL folders if they don't exist
            pdf_folder_path = os.path.join(downloads_folder, "SOFI_PDF")
            excel_folder_path = os.path.join(downloads_folder, "SOFI_EXCEL")
            if not os.path.exists(pdf_folder_path):
                os.makedirs(pdf_folder_path)
            if not os.path.exists(excel_folder_path):
                os.makedirs(excel_folder_path)

            # Construct full file paths with folder names
            pdf_filepath = os.path.join(pdf_folder_path, pdf_filename)
            excel_filepath = os.path.join(excel_folder_path, excel_filename)

            sheets.SaveAs(excel_filepath)

            # Hide other sheets before exporting to PDF
            for i in range(1, sheets.Worksheets.Count + 1):
                if i != 2:  # assuming sheet index 2 is the one you want to keep visible
                    sheets.Worksheets(i).Visible = False

            sheets.ExportAsFixedFormat(0, pdf_filepath)

            # Unhide all sheets after exporting
            for i in range(1, sheets.Worksheets.Count + 1):
                sheets.Worksheets(i).Visible = True

            sheets.Close()

            file_links.append({
                'excel': excel_filename,
                'pdf': pdf_filename
            })
    finally:
        excel.Application.Quit()

    return jsonify(links=file_links, message='SO-FI generated successfully!')

# @app.route('/download/<file_type>/<filename>')
# def download_file(file_type, filename):
#     if file_type not in ['excel', 'pdf']:
#         return jsonify(message='Invalid file type'), 400

#     folder = app.config['GENERATED_EXCEL_FOLDER'] if file_type == 'excel' else app.config['GENERATED_PDF_FOLDER']
#     return send_from_directory(folder, filename)

if __name__ == "__main__":
    app.run(debug=True)
