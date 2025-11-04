from flask import Flask, render_template, request, send_file
import os
from converter import convert_pdf_to_ppt  # your function

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['pdf']
    file_path = "temp.pdf"
    output_path = "converted.pptx"

    file.save(file_path)
    convert_pdf_to_ppt(file_path, output_path)

    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
