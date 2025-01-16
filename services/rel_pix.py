import pandas as pd
from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename

# Defina o caminho fixo para o arquivo de contas contábeis
contabil_file_path = r"Códigos.xlsx"

def select_file(prompt="Selecione o arquivo"):
    Tk().withdraw()
    print(prompt)
    filename = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    return filename

def select_save_file(prompt="Salvar arquivo como"):
    Tk().withdraw()
    filename = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")], title=prompt)
    return filename

def criar_tabela_dinamica(df, output_file):
    # Gerar colunas "Descrição" e "Nome" para a tabela dinâmica
    df['Descrição'] = df.apply(lambda x: f"{int(x['Conta'])} {x['ContaDescricao']}|Contábil: {int(x['Contabil']) if pd.notna(x['Contabil']) else 747}", axis=1)
    df['Nome'] = df.apply(lambda x: f"{x['Documento']} {x['Matricula']} {x['Nome']}", axis=1)

    # Criar a tabela dinâmica
    tabela_dinamica = pd.pivot_table(
        df,
        index=['Conta', 'Descrição', 'Situacao', 'Nome'],
        values=['Valor', 'Juros', 'ValorPag'],
        aggfunc='sum',
        margins=False
    )

    # Adicionar linhas de subtotais
    subtotais_descricao = tabela_dinamica.groupby(level=['Conta', 'Descrição']).sum()
    subtotais_descricao.index = pd.MultiIndex.from_tuples(
        [(conta, desc, 'Subtotal', '') for conta, desc in subtotais_descricao.index]
    )
    subtotais_descricao.index.names = tabela_dinamica.index.names

    subtotais_situacao = tabela_dinamica.groupby(level=['Conta', 'Descrição', 'Situacao']).sum()
    subtotais_situacao.index = pd.MultiIndex.from_tuples(
        [(conta, desc, sit, 'Subtotal') for conta, desc, sit in subtotais_situacao.index]
    )
    subtotais_situacao.index.names = tabela_dinamica.index.names

    # Concatenar a tabela com os subtotais
    tabela_completa = pd.concat([tabela_dinamica, subtotais_situacao, subtotais_descricao]).sort_index()

    # Calcular e adicionar os totais das contas "8888" e "9999" ao final
    totais_8888_9999 = pd.DataFrame({
        'Valor': [df.loc[df['Conta'] == '8888', 'Valor'].sum(), df.loc[df['Conta'] == '9999', 'Valor'].sum()],
        'Juros': [0, 0],
        'ValorPag': [df.loc[df['Conta'] == '8888', 'ValorPag'].sum(), df.loc[df['Conta'] == '9999', 'ValorPag'].sum()]
    }, index=pd.MultiIndex.from_tuples([
        (8888, 'Operações|Contábil: 8888', 'Subtotal', ''),
        (9999, 'Taxa|Contábil: 9999', 'Subtotal', '')
    ], names=['Conta', 'Descrição', 'Situacao', 'Nome']))

    # Concatenar totais "8888" e "9999" ao final da tabela
    tabela_completa = pd.concat([tabela_completa, totais_8888_9999])

    # Calcular o total geral com a soma padrão da tabela dinâmica
    totais_gerais = pd.DataFrame(tabela_dinamica.sum()).T
    totais_gerais.index = pd.Index([('Total Geral', '', '', '')])
    totais_gerais.index.names = tabela_dinamica.index.names

    # Adicionar manualmente os valores de 8888 e 9999 ao "Total Geral"
    totais_gerais[['Valor', 'Juros', 'ValorPag']] += totais_8888_9999[['Valor', 'Juros', 'ValorPag']].sum()

    # Concatenar o total geral à tabela completa
    tabela_completa = pd.concat([tabela_completa, totais_gerais])

    # Selecionar colunas finais
    tabela_completa = tabela_completa[['Valor', 'Juros', 'ValorPag']]

    # Salvar a tabela dinâmica com subtotais no arquivo de saída
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
        tabela_completa.to_excel(writer, sheet_name='Relatório')

    print(f"Tabela dinâmica com subtotais salva com sucesso em: {output_file}")

def process_excel_pix():
    input_file = select_file("Selecione o arquivo principal:")
    if not input_file:
        print("Nenhum arquivo selecionado.")
        return

    try:
        df_contabil = pd.read_excel(contabil_file_path, sheet_name="Planilha2", dtype={'Conta': str})
        df_contabil['Conta'] = df_contabil['Conta'].str.strip()
        df_contabil['Descricao'] = df_contabil['Descricao'].astype(str).str.strip()
        df_contabil['Chave'] = df_contabil['Conta'] + ' ' + df_contabil['Descricao']
    except FileNotFoundError:
        print(f"Arquivo de contas contábeis não encontrado no caminho: {contabil_file_path}")
        return

    workbook = load_workbook(input_file, data_only=True)
    sheets = workbook.sheetnames
    print("Planilhas disponíveis:")
    for idx, sheet in enumerate(sheets):
        print(f"{idx + 1}. {sheet}")

    sheet_num = int(input("Digite o número da planilha que você deseja usar: "))
    if sheet_num < 1 or sheet_num > len(sheets):
        print("Número da planilha inválido.")
        return

    sheet_name = sheets[sheet_num - 1]
    df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)

    operacoes_rows = df.iloc[:, 7].str.contains('operações', case=False, na=False)
    taxa_rows = df.iloc[:, 7].str.contains('Taxa', case=False, na=False)

    operacoes_value = df.loc[operacoes_rows, df.columns[8]].values[0] if operacoes_rows.any() else None
    txt_operacoes = df.loc[operacoes_rows, df.columns[7]].values[0] if operacoes_rows.any() else None
    taxa_value = df.loc[taxa_rows, df.columns[8]].values[0] if taxa_rows.any() else None
    txt_taxa = df.loc[taxa_rows, df.columns[7]].values[0] if taxa_rows.any() else None

    first_value_index = df[df.iloc[:, 0].notna()].index[0]
    df.columns = df.iloc[first_value_index]
    df = df.iloc[first_value_index + 1:].reset_index(drop=True)

    if 'Documento' in df.columns:
        df['Documento'] = df['Documento'].astype(str).str.replace(r'^990000', '', regex=True)

    if 'TipoConta' in df.columns:
        df = df.drop('TipoConta', axis=1)

    ultima_linha = df.dropna(how='all').index[-1]
    df.at[ultima_linha + 1, ('Valor', 'ValorPag')] = operacoes_value
    df.at[ultima_linha + 2, ('Valor', 'ValorPag')] = taxa_value
    df.at[ultima_linha + 1, ('ContaDescricao')] = txt_operacoes
    df.at[ultima_linha + 2, ('ContaDescricao')] = txt_taxa
    df.at[ultima_linha + 1, ('Conta')] = '8888'
    df.at[ultima_linha + 2, ('Conta')] = '9999'

    df['Juros'] = df['ValorPag'] - df['Valor']
    df['Juros'] = df['Juros'].clip(lower=0)
    colunas = list(df.columns)
    colunas.insert(9, colunas.pop(colunas.index('Juros')))
    df = df[colunas]

    df.insert(3, 'Contabil', None)

    contas_especiais = ['0101', '0102', '0103', '0104', '0105', '0106']
    for i, row in df.iterrows():
        conta = str(row['Conta']).strip()
        situacao = str(row.get('Situacao', '')).strip()

        if conta in contas_especiais:
            chave = f"{conta} {situacao}"
            conta_contabil = df_contabil[df_contabil['Chave'] == chave]['Contabil'].values
        else:
            conta_contabil = df_contabil[df_contabil['Conta'] == conta]['Contabil'].values

        df.at[i, 'Contabil'] = conta_contabil[0] if conta_contabil.size > 0 else 747

    df = df.sort_values("Conta")

    output_file = select_save_file("Salvar arquivo como:")
    if not output_file:
        print("Nenhum local de salvamento selecionado.")
        return

    # Salva a planilha principal sem incluir 8888 e 9999
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("Arquivo salvo com sucesso em:", output_file)
    
    # Gera a tabela dinâmica incluindo 8888 e 9999
    criar_tabela_dinamica(df, output_file)

if __name__ == "__main__":
    process_excel_pix()
