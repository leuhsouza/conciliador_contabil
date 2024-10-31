import pandas as pd
import os

def obter_data(data_str):
    # Verifica se a data está no formato DDMMAAAA
    while not (len(data_str) == 8 and data_str.isdigit()):
        raise ValueError("Data inválida. Use o formato DDMMAAAA.")
    return data_str

def gerar_primeira_linha(data):
    primeira_linha = "100" + "00" + "07" + "02" + data
    return primeira_linha

def obter_valor(row):
    contas_especificas = {
        '3530','3538','3540','1078','1039','1080','1118','1083','3035',
        '1102','1035','409','421','433','1100','323','330','329','331',
        '469','3098','486','3522','1088','1092','1112', '3548', '3550',
    }
    if row['ValorPag'] < row['Valor'] or row['Conta'] in contas_especificas:
        return row['ValorPag']
    else:
        return row['Valor']

def gerar_linhas_contabeis(df, tipo):
    lancamentos = []
    historico = '0413' if tipo == 'Pix' else '0408'  # Define o histórico com base no tipo de arquivo
    
    # Calcula o valor a considerar em cada linha
    df['ValorConsiderado'] = df.apply(obter_valor, axis=1)
    
    # Filtra e agrupa as contas para tratamento
    contas_separadas = df[df['Conta'].isin(['101', '102', '106'])]
    contas_agrupadas = df[~df['Conta'].isin(['101', '102', '106'])]
    
    contas_agrupadas['Conta'] = contas_agrupadas['Conta'].apply(lambda x: int(x) if str(x).isdigit() else x)
    contas_agrupadas = contas_agrupadas.sort_values(by='Conta', key=lambda x: pd.to_numeric(x, errors='coerce'))
    
    grupos_separados = contas_separadas.groupby(['Conta', 'Situacao'])
    grupos_agrupados = contas_agrupadas.groupby(['Conta'])
    
    numero_lancamento = 1

    # Processa os grupos separados
    for (conta, situacao), grupo in grupos_separados:
        valor_total = grupo['ValorConsiderado'].sum()
        if conta in ['8888', '9999']:
            valor_total = abs(valor_total)
        valor_formatado = f"{abs(valor_total):016.2f}".replace('.', '')
        conta_formatada = f"{int(grupo['Contabil'].iloc[0]):05}"
        complemento = ""
        complemento = complemento[:255]
        
        # Usa o valor de histórico de acordo com o tipo selecionado
        linha = f"200{numero_lancamento:02}{valor_formatado}C{conta_formatada}{historico}{complemento:<255}"
        lancamentos.append((linha, conta, situacao, valor_total, complemento))
        
        numero_lancamento += 1

    # Processa os grupos agrupados
    for conta, grupo in grupos_agrupados:
        valor_total = abs(grupo['ValorConsiderado'].sum())
        if conta in ['8888', '9999']:
            valor_total = abs(valor_total)
        conta_formatada = f"{int(grupo['Contabil'].iloc[0]):05}"
        valor_formatado = f"{valor_total:016.2f}".replace('.', '')
        complemento = "/".join(grupo['Documento'].astype(str).str.replace('.0', '', regex=False))
        complemento = complemento[:255]
        
        # Usa o valor de histórico de acordo com o tipo selecionado
        linha = f"200{numero_lancamento:02}{valor_formatado}C{conta_formatada}{historico}{complemento:<255}"
        lancamentos.append((linha, conta, '', valor_total, complemento))
        
        numero_lancamento += 1
    
    return lancamentos

def gerar_ultima_linha(lancamentos):
    total_linhas = len(lancamentos) + 2
    valor_total = sum(l[3] for l in lancamentos)
    valor_total_formatado = f"{valor_total:014.2f}".replace('.', '')
    ultima_linha = f"300{total_linhas:02}{(total_linhas -1):04}{valor_total_formatado}{valor_total_formatado}"
    return ultima_linha

def gerar_lancamento_debito_total(lancamentos):
    total_creditos = sum(l[3] for l in lancamentos)
    valor_formatado = f"{abs(total_creditos):016.2f}".replace('.', '')
    linha = f"200{len(lancamentos) + 1:02}{valor_formatado}D{'30991'}{'0409'}{'Total de Créditos':<255}"
    return (linha, '00000', '0409', total_creditos, 'Total de Créditos')

def salvar_relatorio_excel(lancamentos, caminho_completo):
    df_relatorio = pd.DataFrame(lancamentos, columns=['Linha', 'Conta', 'Situacao', 'Valor', 'Complemento'])
    df_relatorio.to_excel(caminho_completo, index=False)
    return caminho_completo

def processar_lancamentos(arquivo_path, sheet_name, data, tipo):
    # Formata a data de acordo com o formato desejado
    data_formatada = obter_data(data)
    
    # Carrega e prepara o DataFrame
    df = pd.read_excel(arquivo_path, sheet_name=sheet_name)
    df.rename(columns=lambda x: x.strip(), inplace=True)
    df = df[['Conta', 'Situacao', 'Valor', 'ValorPag', 'Contabil', 'Documento']]
    df.columns = ['Conta', 'Situacao', 'Valor', 'ValorPag', 'Contabil', 'Documento']
    df['Documento'] = df['Documento'].astype(str).str.replace('.0', '', regex=False)
    df['Conta'] = df['Conta'].astype(str).str.replace('.0', '', regex=False)
    df['Contabil'] = df['Contabil'].astype(str).str.replace('.0', '', regex=False)
    
    # Gera a primeira linha do lançamento contábil
    primeira_linha = gerar_primeira_linha(data_formatada)

    # Passa o tipo para a função gerar_linhas_contabeis
    linhas_contabeis = gerar_linhas_contabeis(df, tipo)
    
    # Calcula o lançamento de débito total e a última linha
    lancamento_debito_total = gerar_lancamento_debito_total(linhas_contabeis)
    ultima_linha = gerar_ultima_linha(linhas_contabeis)
    
    # Cria o conteúdo final para ser salvo
    conteudo = [(primeira_linha, '', '', 0, '')] + linhas_contabeis + [lancamento_debito_total] + [(ultima_linha, '', '', 0, '')]
    return conteudo

