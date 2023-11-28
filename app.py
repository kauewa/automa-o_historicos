import io
import os
import webview
import pandas as pd
from flask import Flask, redirect, request, render_template, flash, session, send_file
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from src.utils import gerar_historicos, gerar_historicos_fund

load_dotenv()
app = Flask(__name__, template_folder='./templates', static_folder='./static')
app.secret_key = os.getenv('SECRET_KEY')

webview.create_window('Automação Históricos', app)



@app.route('/', methods=['GET'])
def home():
    return redirect('/medio')

@app.route('/medio', methods=['GET', 'POST'])
def medio():
    if request.method == 'POST':
        print('post')
        
        file_relacao = request.files['file-upload-relacao']
        file_1 = request.files['file-upload-1']
        file_2 = request.files['file-upload-2']
        file_3 = request.files['file-upload-3']
        files = [file_relacao, file_1, file_2, file_3]
        d_frames = []
        for file in files:
            filename = secure_filename(file.filename)
            file_extension = os.path.splitext(filename)[1]
            file.seek(0, os.SEEK_END)  
            file.seek(0)
            print(file_extension)
            if file_extension == '.xlsx':
                df = pd.read_excel(file)
                d_frames.append(df)
        file_zip = gerar_historicos(d_frames[0],d_frames[1],d_frames[2],d_frames[3])
        file_zip_io = io.BytesIO(file_zip)
        flash('Arquivos enviados com sucesso!', 'INFO')
        return send_file(file_zip_io, mimetype='application/zip',
                         as_attachment=True, download_name='arquivos.zip')
    return render_template('medio.html')

@app.route('/fundamental', methods=['GET', 'POST'])
def fundamental():
    if request.method == 'POST':
        print('post')
        
        file_relacao = request.files['file-upload-relacao']
        file_1 = request.files['file-upload-1']
        file_2 = request.files['file-upload-2']
        file_3 = request.files['file-upload-3']
        file_4 = request.files['file-upload-4']
        files = [file_relacao, file_1, file_2, file_3, file_4]
        d_frames = []
        for file in files:
            filename = secure_filename(file.filename)
            file_extension = os.path.splitext(filename)[1]
            file.seek(0, os.SEEK_END)  
            file.seek(0)
            print(file_extension)
            if file_extension == '.xlsx':
                df = pd.read_excel(file)
                d_frames.append(df)
        file_zip = gerar_historicos_fund(d_frames[0],d_frames[1],d_frames[2],d_frames[3], d_frames[4])
        file_zip_io = io.BytesIO(file_zip)
        flash('Arquivos enviados com sucesso!', 'INFO')
        return send_file(file_zip_io, mimetype='application/zip',
                         as_attachment=True, download_name='arquivos.zip')
    return render_template('fundamental.html')

if __name__ == '__main__':
    webview.start()