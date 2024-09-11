from flask import Flask, render_template, request, redirect, url_for, send_file
import sqlite3
import pandas as pd
from data_cleaning import process_excel
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file:
        file_path = 'uploaded_file.xlsx'
        file.save(file_path)
        process_excel(file_path, 'database.sqlite')
        return redirect(url_for('data'))

@app.route('/download', methods=['POST'])
def download_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file:
        file_path = 'uploaded_file.xlsx'
        output_path = 'processed_file.xlsx'
        file.save(file_path)
        process_excel(file_path, 'database.sqlite', output_path)
        return send_file(output_path, as_attachment=True)

@app.route('/data')
def data():
    conn = sqlite3.connect('database.sqlite')
    df = pd.read_sql_query("SELECT * FROM dados", conn)
    conn.close()
    table_html = df.to_html(classes='data', index=False)
    return render_template('data.html', table=table_html)

@app.route('/pix')
def arquivo_pix():
    return

if __name__ == '__main__':
    app.run(host='192.168.0.77', port=5000, debug=True)
