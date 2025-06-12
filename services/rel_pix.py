import pandas as pd
from openpyxl import load_workbook

contabil_file_path = r"Códigos.xlsx"


def criar_tabela_dinamica(df, output_file):
    # Diagnóstico: ver colunas disponíveis e os primeiros valores
    print("\n📌 Colunas disponíveis:", list(df.columns))
    print("📌 Amostra de ValorPago:", df['ValorPago'].dropna().head())

    # Criar coluna 'Descrição'
    df['Descrição'] = df.apply(
        lambda x: f"{int(x['Conta'])} {x['ContaDescricao']}|Contábil: {int(x['Contabil']) if pd.notna(x['Contabil']) else 747}",
        axis=1
    )

    # Criar coluna 'Nome'
    df['Nome'] = df.apply(lambda x: f"{x['Documento']} {x['Matricula']} {x['Nome']}", axis=1)

    # Criar tabela dinâmica com as 3 colunas: Valor, Juros e ValorPago
    tabela_dinamica = pd.pivot_table(
        df,
        index=['Conta', 'Descrição', 'Situacao', 'Nome'],
        values=['Valor', 'Juros', 'ValorPago'],
        aggfunc='sum'
    )

    # Subtotais por Conta + Descrição
    subtotais_descricao = tabela_dinamica.groupby(level=['Conta', 'Descrição']).sum()
    subtotais_descricao.index = pd.MultiIndex.from_tuples(
        [(conta, desc, 'Subtotal', '') for conta, desc in subtotais_descricao.index]
    )
    subtotais_descricao.index.names = tabela_dinamica.index.names

    # Subtotais por Conta + Descrição + Situação
    subtotais_situacao = tabela_dinamica.groupby(level=['Conta', 'Descrição', 'Situacao']).sum()
    subtotais_situacao.index = pd.MultiIndex.from_tuples(
        [(conta, desc, sit, 'Subtotal') for conta, desc, sit in subtotais_situacao.index]
    )
    subtotais_situacao.index.names = tabela_dinamica.index.names

    # Concatenar tudo
    tabela_completa = pd.concat([tabela_dinamica, subtotais_situacao, subtotais_descricao]).sort_index()

    # Adicionar totais de contas especiais
    totais_8888_9999 = pd.DataFrame({
        'Valor': [df.loc[df['Conta'] == '8888', 'Valor'].sum(), df.loc[df['Conta'] == '9999', 'Valor'].sum()],
        'Juros': [0, 0],
        'ValorPago': [df.loc[df['Conta'] == '8888', 'ValorPago'].sum(), df.loc[df['Conta'] == '9999', 'ValorPago'].sum()]
    }, index=pd.MultiIndex.from_tuples([
        (8888, 'Operações|Contábil: 8888', 'Subtotal', ''),
        (9999, 'Taxa|Contábil: 9999', 'Subtotal', '')
    ], names=['Conta', 'Descrição', 'Situacao', 'Nome']))

    tabela_completa = pd.concat([tabela_completa, totais_8888_9999])

    # Totais gerais
    totais_gerais = pd.DataFrame(tabela_dinamica.sum()).T
    totais_gerais.index = pd.Index([('Total Geral', '', '', '')])
    totais_gerais.index.names = tabela_dinamica.index.names

    totais_gerais[['Valor', 'Juros', 'ValorPago']] += totais_8888_9999[['Valor', 'Juros', 'ValorPago']].sum()

    tabela_completa = pd.concat([tabela_completa, totais_gerais])

    # Garantir a ordem e presença das colunas
    for col in ['Valor', 'Juros', 'ValorPago']:
        if col not in tabela_completa.columns:
            tabela_completa[col] = 0.0

    tabela_completa = tabela_completa[['Valor', 'Juros', 'ValorPago']]

    # Escrever no arquivo Excel
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        tabela_completa.to_excel(writer, sheet_name='Relatório')

    print(f"\n✅ Tabela dinâmica com subtotais salva com sucesso em: {output_file}")


def process_excel_pix(input_file, sheet_name, output_file):
    try:
        if not input_file:
            raise ValueError("Caminho do arquivo não fornecido.")
        if not sheet_name:
            raise ValueError("Nome da planilha não fornecido.")
        
        workbook = load_workbook(input_file, data_only=True)
        sheets = workbook.sheetnames
        
        if sheet_name not in sheets:
            raise ValueError(f"A planilha '{sheet_name}' não foi encontrada no arquivo '{input_file}'.")

        df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)

        operacoes_rows = df.iloc[:, 1].str.contains('operações', case=False, na=False)
        taxa_rows = df.iloc[:, 1].str.contains('taxa', case=False, na=False)

        operacoes_value, txt_operacoes = None, None
        taxa_value, txt_taxa = None, None

        if operacoes_rows.any():
            operacoes_value = df.loc[operacoes_rows, df.columns[2]].values[0]
            txt_operacoes = df.loc[operacoes_rows, df.columns[1]].values[0]

        if taxa_rows.any():
            taxa_value = df.loc[taxa_rows, df.columns[2]].values[0]
            txt_taxa = df.loc[taxa_rows, df.columns[1]].values[0]

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

        
        # Função de limpeza
        def limpar_valor(valor):
            if isinstance(valor, str):
                return valor.replace('R$', '').replace('.', '').replace(',', '.').strip()
            return valor

        df['Valor'] = df['Valor'].apply(limpar_valor)
        df['ValorPag'] = df['ValorPag'].apply(limpar_valor)
        df['Juros'] = df['Juros'].apply(limpar_valor)

        # Conversão segura para float
        df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce')
        df['ValorPag'] = pd.to_numeric(df['ValorPag'], errors='coerce')
        df['Juros'] = pd.to_numeric(df['Juros'], errors='coerce')

        # Apenas calcula os Juros se ele estiver vazio (nan ou 0)
        juros_recalculado = df['ValorPag'] - df['Valor']
        df['Juros'] = df['Juros'].combine_first(juros_recalculado).fillna(0).clip(lower=0)


        colunas = list(df.columns)
        colunas.insert(9, colunas.pop(colunas.index('Juros')))
        df = df[colunas]

        df.insert(3, 'Contabil', None)

        df_contabil = pd.read_excel(contabil_file_path, sheet_name="Planilha2", dtype={'Conta': str})
        df_contabil['Conta'] = df_contabil['Conta'].str.strip()
        df_contabil['Descricao'] = df_contabil['Descricao'].astype(str).str.strip()
        df_contabil['Chave'] = df_contabil['Conta'] + ' ' + df_contabil['Descricao']

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

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Arquivo processado salvo em: {output_file}")
    
    except Exception as e:
        raise ValueError(f"Erro ao processar o arquivo: {e}")

    criar_tabela_dinamica(df, output_file)
