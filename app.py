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
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>PDF to Excel Converter</title>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background: #f4f6f8;
      margin: 0;
      padding: 0;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      height: 100vh;
    }
    h1 {
      color: #333;
      margin-bottom: 20px;
    }
    form {
      background: white;
      padding: 30px 40px;
      border-radius: 10px;
      box-shadow: 0 0 20px rgba(0,0,0,0.1);
      display: flex;
      flex-direction: column;
      align-items: center;
    }
    input[type="file"] {
      margin-bottom: 20px;
    }
    input[type="submit"] {
      background: #4CAF50;
      color: white;
      border: none;
      padding: 10px 20px;
      border-radius: 5px;
      cursor: pointer;
      font-size: 16px;
    }
    input[type="submit"]:hover {
      background: #45a049;
    }
    a {
      display: inline-block;
      margin-top: 20px;
      color: #2196F3;
      text-decoration: none;
      font-weight: bold;
    }
    a:hover {
      text-decoration: underline;
    }
  </style>
</head>
<body>
  <h1>PDF to Excel Converter</h1>
  <form method="post" enctype="multipart/form-data">
    <input type="file" name="file" accept=".pdf" required>
    <input type="submit" value="Upload and Convert">
  </form>
  {% if download_link %}
    <a href="{{ download_link }}">⬇️ Download Excel File</a>
  {% endif %}
</body>
</html>
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
