from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify, make_response,session
import sqlite3
import pandas as pd
from services.data_cleaning import process_excel, process_excel_varias_contas
from services.pixtxt import processar_lancamentos
import os
from dotenv import load_dotenv


# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

app = Flask(__name__)


# Defina a chave secreta a partir da variável de ambiente
app.secret_key = os.getenv('SECRET_KEY')

if not app.secret_key:
    raise RuntimeError("A chave secreta não foi configurada. Verifique as variáveis de ambiente.")

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

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(filename, as_attachment=True)

@app.route('/data', methods=['GET', 'POST'])
def data():
    conn = sqlite3.connect('database.sqlite')
    query = "SELECT * FROM dados"
    filters = []
    where_clauses = []

    # Lista de colunas válidas para ordenação
    valid_columns = ['id', 'data', 'historico', 'contra_partida', 'lote', 'lancamento', 'd', 'c', 'dc', 'conta', 'conciliado']

    # Captura os parâmetros de ordenação da solicitação
    order_by = request.args.get('order_by', 'id')  # Define a coluna 'id' como padrão
    order_direction = request.args.get('order_direction', 'asc')  # 'asc' é o padrão

    # Valida o parâmetro de ordenação para evitar injeção de SQL
    if order_by not in valid_columns:
        order_by = 'id'  # Redefine para a coluna padrão se o parâmetro for inválido

    # Valida a direção da ordenação
    if order_direction.lower() not in ['asc', 'desc']:
        order_direction = 'asc'  # Redefine para ascendente se o parâmetro for inválido

    # Adiciona a cláusula ORDER BY à consulta SQL
    query += f" ORDER BY {order_by} {order_direction.upper()}"

    if request.method == 'POST':
        filter_value = request.form.get('filter_field')
        filter_conta = request.form.get('filter_conta')

        if filter_value:
            where_clauses.append("historico LIKE ?")
            filters.append(f"%{filter_value}%")

        if filter_conta:
            where_clauses.append("conta LIKE ?")
            filters.append(f"%{filter_conta}%")

        if where_clauses:
            query = query.replace('ORDER BY', f"WHERE {' AND '.join(where_clauses)} ORDER BY")

    df = pd.read_sql_query(query, conn, params=filters)
    conn.close()

    if df.empty:
        return render_template('data.html', data_list=[], columns=[], initial_sum=0)

    # Calcular a soma inicial dos valores conciliados
    df['d'] = pd.to_numeric(df['d'], errors='coerce').fillna(0)
    df['c'] = pd.to_numeric(df['c'], errors='coerce').fillna(0)
    initial_sum = (df[df['conciliada'] == 1]['d'].sum() - df[df['conciliada'] == 1]['c'].sum())

    data_list = df.to_dict(orient='records')
    columns = df.columns.tolist()

    return render_template('data.html', data_list=data_list, columns=columns, initial_sum=initial_sum)



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

        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names

        return render_template('pix_select_sheet.html', sheet_names=sheet_names, file_path=file_path)

    return render_template('pix.html')

@app.route('/process_pix', methods=['POST'])
def process_pix():
    file_path = request.form.get('file_path')
    sheet_name = request.form.get('sheet_name')
    data = request.form.get('data')
    tipo = request.form.get('tipo')

    if file_path and sheet_name and data and tipo:
        conteudo = processar_lancamentos(file_path, sheet_name, data, tipo)

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

        return render_template('import_data.html', file_path=file_path)

    return render_template('import_data.html')

@app.route('/process_data', methods=['POST'])
def process_data():
    file_path = request.form.get('file_path')
    tipo = request.form.get('tipo')  # Captura o tipo de arquivo (Razao ou Lote)
    subtipo = request.form.get('subtipo')  # Captura o subtipo (Conta Unitaria ou Varias Contas)
    enviar_bd = 'enviar_bd' in request.form

    # Prints de depuração para verificar os valores recebidos
    print(f"Tipo recebido: {tipo}")
    print(f"Subtipo recebido: {subtipo}")
    print(f"Enviar para o banco de dados: {enviar_bd}")

    if file_path and tipo:
        try:
            db_path = 'database.sqlite'
            output_path = 'processed_file.xlsx'

            if tipo == "Razao" and subtipo == "varias":
                print("Chamando a função process_excel_varias_contas")
                linhas_importadas = process_excel_varias_contas(file_path, db_path, output_path, enviar_bd=enviar_bd)
            else:
                print("Chamando a função process_excel")
                linhas_importadas = process_excel(file_path, db_path, output_path, enviar_bd=enviar_bd)

            mensagem = f"Arquivo processado com sucesso. {'Linhas importadas para o banco de dados: ' + str(linhas_importadas) if enviar_bd else 'Nenhuma linha importada, apenas o arquivo foi salvo.'}"

            return jsonify(success=True, message=mensagem, download_url=url_for('download_file', filename=output_path))

        except Exception as e:
            return jsonify(success=False, message=f"Erro ao processar o arquivo: {str(e)}")

    return jsonify(success=False, message="Ocorreu um problema: faltam informações necessárias.")

import pandas as pd
from flask import send_file

@app.route('/download_filtered_data', methods=['POST'])
def download_filtered_data():
    filter_value = request.form.get('filter_field')
    filter_conta = request.form.get('filter_conta')

    conn = sqlite3.connect('database.sqlite')
    query = "SELECT * FROM dados"
    filters = []
    where_clauses = []

    if filter_value:
        where_clauses.append("historico LIKE ?")
        filters.append(f"%{filter_value}%")

    if filter_conta:
        where_clauses.append("conta LIKE ?")
        filters.append(f"%{filter_conta}%")

    if where_clauses:
        query += f" WHERE {' AND '.join(where_clauses)}"

    df = pd.read_sql_query(query, conn, params=filters)
    conn.close()

    if df.empty:
        return "Nenhum dado encontrado com os filtros aplicados."

    # Salvar o DataFrame em um arquivo Excel com formatação para linhas destacadas
    output_path = 'filtered_data.xlsx'
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='FilteredData')
        workbook = writer.book
        worksheet = writer.sheets['FilteredData']

        # Formatação para as células destacadas
        highlighted_format = workbook.add_format({'bg_color': 'yellow'})

        # Destacar as células de linhas que têm a coluna 'conciliada' igual a True ou 1
        for row_num, value in enumerate(df['conciliada'], start=1):  # Começa em 1 para pular o cabeçalho
            if value == 1 or value is True:
                for col_num in range(len(df.columns)):
                    worksheet.write(row_num, col_num, df.iloc[row_num - 1, col_num], highlighted_format)

    return send_file(output_path, as_attachment=True)


@app.route('/download_non_conciliated', methods=['POST'])
def download_non_conciliated():
    filter_value = request.form.get('filter_field')
    filter_conta = request.form.get('filter_conta')

    conn = sqlite3.connect('database.sqlite')
    query = "SELECT * FROM dados WHERE conciliada = 0"
    filters = []
    where_clauses = []

    if filter_value:
        where_clauses.append("historico LIKE ?")
        filters.append(f"%{filter_value}%")

    if filter_conta:
        where_clauses.append("conta LIKE ?")
        filters.append(f"%{filter_conta}%")

    if where_clauses:
        query += f" AND {' AND '.join(where_clauses)}"

    df = pd.read_sql_query(query, conn, params=filters)
    conn.close()

    if df.empty:
        return "Nenhum dado encontrado com os filtros aplicados."

    # Salvar o DataFrame em um arquivo Excel
    output_path = 'non_conciliated_data.xlsx'
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='NonConciliatedData')
        workbook = writer.book
        worksheet = writer.sheets['NonConciliatedData']

        # Formatação opcional pode ser adicionada aqui se necessário

    return send_file(output_path, as_attachment=True)


@app.route('/save_conciliation', methods=['POST'])
def save_conciliation():
    data = request.get_json()
    selected_ids = data.get('ids', [])  # IDs das linhas que estão selecionadas
    all_ids = data.get('all_ids', [])  # IDs de todas as linhas filtradas ou mostradas

    if not all_ids:
        return jsonify(success=False, message="Nenhuma linha foi encontrada para conciliação.")

    if not isinstance(all_ids, list) or not isinstance(selected_ids, list):
        return jsonify(success=False, message="Formato de dados inválido para IDs.")

    try:
        with sqlite3.connect('database.sqlite') as conn:
            cursor = conn.cursor()

            # Atualizar todas as linhas exibidas para conciliada = 0 (não selecionadas)
            if all_ids:
                query_reset = f"UPDATE dados SET conciliada = 0 WHERE id IN ({','.join('?' * len(all_ids))})"
                cursor.execute(query_reset, all_ids)

            # Atualizar apenas as linhas selecionadas para conciliada = 1
            if selected_ids:
                query_set = f"UPDATE dados SET conciliada = 1 WHERE id IN ({','.join('?' * len(selected_ids))})"
                cursor.execute(query_set, selected_ids)

            conn.commit()

        return jsonify(success=True, message="Conciliação salva com sucesso.")
    except Exception as e:
        return jsonify(success=False, message=f"Erro ao salvar a conciliação: {str(e)}")



@app.route('/remove_conciliation', methods=['POST'])
def remove_conciliation():
    data = request.get_json()
    filter_field = data.get('filter_field', '').strip()
    filter_conta = data.get('filter_conta', '').strip()

    conn = sqlite3.connect('database.sqlite')
    cursor = conn.cursor()

    try:
        if filter_field or filter_conta:
            # Remover a conciliação das linhas que correspondem aos filtros
            where_clauses = []
            params = []

            if filter_field:
                where_clauses.append("historico LIKE ?")
                params.append(f"%{filter_field}%")

            if filter_conta:
                where_clauses.append("conta LIKE ?")
                params.append(f"%{filter_conta}%")

            query = f"UPDATE dados SET conciliada = 0 WHERE {' AND '.join(where_clauses)}"
            cursor.execute(query, params)
        else:
            # Se não houver filtro, remover a conciliação de todas as linhas
            cursor.execute("UPDATE dados SET conciliada = 0")

        conn.commit()  # Commit temporário para refletir a mudança visualmente, mas sem salvar
        conn.close()

        return jsonify(success=True, message="Conciliação removida com sucesso.")
    except Exception as e:
        conn.close()
        return jsonify(success=False, message=f"Erro ao remover a conciliação: {str(e)}")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
