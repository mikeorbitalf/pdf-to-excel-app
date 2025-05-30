import os
import pdfplumber
import pandas as pd
from flask import Flask, request, redirect, send_file, render_template_string

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# HTML template
HTML = '''
<!doctype html>
<title>PDF to Excel</title>
<h1>Upload PDF to Convert to Excel</h1>
<form method=post enctype=multipart/form-data>
  <input type=file name=file accept=".pdf">
  <input type=submit value=Upload>
</form>
{% if download_link %}
  <p><a href="{{ download_link }}">Download Excel File</a></p>
{% endif %}
'''

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        pdf_file = request.files['file']
        if pdf_file and pdf_file.filename.endswith('.pdf'):
            pdf_path = os.path.join(UPLOAD_FOLDER, pdf_file.filename)
            pdf_file.save(pdf_path)

            # Process PDF
            all_data = []
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        all_data.extend(table)

            if all_data:
                df = pd.DataFrame(all_data[1:], columns=all_data[0])
                excel_path = os.path.join(UPLOAD_FOLDER, 'output.xlsx')
                df.to_excel(excel_path, index=False)
                return render_template_string(HTML, download_link='/download')
            else:
                return 'No tables found in the PDF.'

    return render_template_string(HTML, download_link=None)

@app.route('/download')
def download_file():
    return send_file(os.path.join(UPLOAD_FOLDER, 'output.xlsx'), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
