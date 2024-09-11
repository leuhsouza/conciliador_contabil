import os
from flask import Flask, render_template, request, redirect
import pandas as pd

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def limpar_arquivo_original(file_path, output_path):
    df = pd.read_excel(file_path, header=None)
    colunas_necessarias = [2, 3, 10, 11, 12, 13, 14]
    df_limpo = df.iloc[:, colunas_necessarias]
    df_limpo.columns = ['Coluna_C', 'Coluna_D', 'Coluna_K', 'Coluna_L', 'Coluna_M', 'Coluna_N', 'Coluna_O']
    df_limpo.to_excel(output_path, index=False)

def conciliar_dados(file_path):
    df_conta = pd.read_excel(file_path)
    coluna_debito = 'Coluna_N'
    coluna_credito = 'Coluna_O'
    df_conta['Conciliado'] = False
    debitos = df_conta[df_conta[coluna_debito].notna()]
    creditos = df_conta[df_conta[coluna_credito].notna()]

    for index, debito in debitos.iterrows():
        correspondente = creditos[
            (creditos[coluna_credito] == debito[coluna_debito]) &
            (creditos['Coluna_C'] == debito['Coluna_C']) &
            (creditos['Conciliado'] == False)
        ]
        if not correspondente.empty:
            df_conta.loc[debito.name, 'Conciliado'] = True
            df_conta.loc[correspondente.index, 'Conciliado'] = True

    for index, credito in creditos.iterrows():
        correspondente = debitos[
            (debitos[coluna_debito] == credito[coluna_credito]) &
            (debitos['Coluna_C'] == credito['Coluna_C']) &
            (debitos['Conciliado'] == False)
        ]
        if not correspondente.empty:
            df_conta.loc[credito.name, 'Conciliado'] = True
            df_conta.loc[correspondente.index, 'Conciliado'] = True

    return df_conta

def gerar_html_tabela(df):
    html = '<table class="data">'
    html += '<tr><th>Lote</th><th>Item</th><th>Débito</th><th>Crédito</th><th>Conciliado</th></tr>'
    for index, row in df.iterrows():
        color = 'lightgreen' if row['Conciliado'] else 'lightcoral'
        html += f'<tr style="background-color: {color};">'
        html += f'<td>{row.get("Coluna_C", "")}</td>'
        html += f'<td>{row.get("Coluna_D", "")}</td>'
        html += f'<td>{row.get("Coluna_N", "")}</td>'
        html += f'<td>{row.get("Coluna_O", "")}</td>'
        html += f'<td>{"Sim" if row["Conciliado"] else "Não"}</td>'
        html += '</tr>'
    html += '</table>'
    return html

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file:
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)
            
            # Gerar arquivo limpo
            arquivo_limpo_path = file_path.replace('.xlsx', '_limpo.xlsx')
            limpar_arquivo_original(file_path, arquivo_limpo_path)
            
            # Conciliar dados do arquivo limpo
            df_conciliado = conciliar_dados(arquivo_limpo_path)
            html_tabela = gerar_html_tabela(df_conciliado)
            return render_template('index.html', html_tabela=html_tabela)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
