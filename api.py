import os, flask
from uuid import uuid4
from flask import Flask, render_template, request, jsonify, send_from_directory, send_file
from werkzeug.utils import secure_filename
from docx2pdf import convert

import win32com.client
import pythoncom


app = flask.Flask(__name__)
app.config["DEBUG"] = True


@app.route('/', methods=['GET'])
def home():
    return "<h1>Distant Reading Archive</h1><p>This site is a prototype API for distant reading of science fiction novels.</p>"

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    upload_id = str(uuid4())
    os.makedirs(os.path.join('request', upload_id), exist_ok=True)
    save_path = os.path.join('request', upload_id, secure_filename(file.filename))
    file.save(save_path)
    xl=win32com.client.Dispatch("Word.Application",pythoncom.CoInitialize())
    convert(save_path)
    #return save_path.replace('docx','pdf')
    return send_file(save_path.replace('docx','pdf'))

app.run()
