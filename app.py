from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify
import sqlite3
import pandas as pd
from services.data_cleaning import process_excel
from services.pixtxt import processar_lancamentos
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

# opção para subir no banco de dados
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

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(filename, as_attachment=True)

@app.route('/data', methods=['GET', 'POST'])
def data():
    conn = sqlite3.connect('database.sqlite')
    query = "SELECT * FROM dados"
    filters = []

    if request.method == 'POST':
        filter_value = request.form.get('filter_field')
        if filter_value:
            query += " WHERE coluna2 LIKE ?"
            filters.append(f"%{filter_value}%")
    df = pd.read_sql_query(query, conn, params=filters)
    conn.close()
    table_html = df.to_html(classes='data', index=False)
    return render_template('data.html', table=table_html)

@app.route('/pix', methods=['GET', 'POST'])
def arquivo_pix():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)

        file_path = 'uploaded_file.xlsx'
        file.save(file_path)

        # Carrega o arquivo e obtém os nomes das planilhas
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names

        # Renderiza um formulário para escolher a planilha
        return render_template('pix_select_sheet.html', sheet_names=sheet_names, file_path=file_path)

    return render_template('pix.html')

@app.route('/process_pix', methods=['POST'])
def process_pix():
    file_path = request.form.get('file_path')
    sheet_name = request.form.get('sheet_name')
    data = request.form.get('data')
    tipo = request.form.get('tipo')  # Captura o tipo de arquivo (Pix ou Avulso)

    if file_path and sheet_name and data and tipo:
        conteudo = processar_lancamentos(file_path, sheet_name, data, tipo)

        # Salva o conteúdo em um arquivo .txt para download
        output_path = 'arquivo_pix.txt'
        with open(output_path, 'w') as f:
            for linha in conteudo:
                f.write(linha[0] + "\n")

        return send_file(output_path, as_attachment=True)

    return redirect(url_for('arquivo_pix'))

@app.route('/import_data', methods=['GET', 'POST'])
def import_data():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)

        file_path = 'uploaded_file.xlsx'
        file.save(file_path)

        # Renderiza um formulário para escolher a planilha
        return render_template('import_data.html', file_path=file_path)

    return render_template('import_data.html')

@app.route('/process_data', methods=['POST'])
def process_data():
    file_path = request.form.get('file_path')
    tipo = request.form.get('tipo')  # Captura o tipo de arquivo (Razao ou Lote)
    enviar_bd = 'enviar_bd' in request.form  # Verifica se o checkbox foi marcado

    if file_path and tipo:
        try:
            # Caminho do banco de dados
            db_path = 'database.sqlite'
            # Caminho do arquivo de saída em Excel
            output_path = 'processed_file.xlsx'

            # Chama a função para processar o arquivo e contar as linhas importadas
            linhas_importadas = process_excel(file_path, db_path, output_path, enviar_bd=enviar_bd)

            # Mensagem de retorno
            mensagem = f"Arquivo processado com sucesso. {'Linhas importadas para o banco de dados: ' + str(linhas_importadas) if enviar_bd else 'Nenhuma linha importada, apenas o arquivo foi salvo.'}"

            return jsonify(success=True, message=mensagem, download_url=url_for('download_file', filename=output_path))

        except Exception as e:
            return jsonify(success=False, message=f"Erro ao processar o arquivo: {str(e)}")

    return jsonify(success=False, message="Ocorreu um problema: faltam informações necessárias.")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
