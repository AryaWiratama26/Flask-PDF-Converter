from flask import Flask, render_template, request, send_file
import os
from pdf2docx import Converter

app = Flask(__name__)
UPLOAD_FOL = 'uploads'
app.config['UPLOAD_FOL'] = UPLOAD_FOL

if not os.path.exists(UPLOAD_FOL):
    os.makedirs(UPLOAD_FOL)

# Fungsi konvert pdf ke word
def pdf_to_word(filename):
    file_word = os.path.splitext(filename)[0] + '.docx'
    
    cv = Converter(filename)
    cv.convert(file_word, start=0, end=None)  
    cv.close()
    
    return send_file(file_word, as_attachment=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        pilihan_konversi = request.form.get('pilihan_konversi')

        # Simpan file ke uploads
        filename = os.path.join(app.config['UPLOAD_FOL'], file.filename)
        file.save(filename)

        if pilihan_konversi == 'pdfToWord':
            return pdf_to_word(filename)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, port=8000)