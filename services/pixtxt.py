import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog
import os

def escolher_arquivo():
    root = tk.Tk()
    root.withdraw()
    arquivo_path = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivo Excel", "*.xlsx *.xls")]
    )
    return arquivo_path

def escolher_planilha(arquivo_path):
    xls = pd.ExcelFile(arquivo_path)
    planilhas = xls.sheet_names
    
    print("Planilhas disponíveis:")
    for idx, planilha in enumerate(planilhas):
        print(f"{idx+1}. {planilha}")
    
    while True:
        escolha = simpledialog.askinteger("Escolha da Planilha", "Escolha o número da planilha:")
        if 1 <= escolha <= len(planilhas):
            planilha_escolhida = planilhas[escolha - 1]
            break
        else:
            print("Escolha inválida. Tente novamente.")
    
    return planilha_escolhida

def obter_data():
    data = simpledialog.askstring("Data do Lançamento", "Insira a data no formato DDMMAAAA:")
    while not (len(data) == 8 and data.isdigit()):
        data = simpledialog.askstring("Data do Lançamento", "Formato inválido. Insira a data no formato DDMMAAAA:")
    return data

def gerar_primeira_linha(data):
    primeira_linha = "100" + "00" + "07" + "02" + data
    return primeira_linha

#considerar a conta como '409' sem o zero
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

def gerar_linhas_contabeis(df):
    lancamentos = []
    df['ValorConsiderado'] = df.apply(obter_valor, axis=1)
    
    # Filtra as contas que devem ser tratadas separadamente por situação
    contas_separadas = df[df['Conta'].isin(['101', '102', '106'])]
    contas_agrupadas = df[~df['Conta'].isin(['101', '102', '106'])]

    
    def convert_to_int(value):
        try:
            return int(value)
        except ValueError:
            return value
        
    contas_agrupadas['Conta'] = contas_agrupadas['Conta'].apply(convert_to_int)
    contas_agrupadas = contas_agrupadas.sort_values(by='Conta', key=lambda x: pd.to_numeric(x, errors='coerce'))
    # Agrupa as contas separadas por Conta e Situacao
    grupos_separados = contas_separadas.groupby(['Conta', 'Situacao'])
    # Agrupa as demais contas apenas por Conta
    grupos_agrupados = contas_agrupadas.groupby(['Conta'])
    for nome_grupo, grupo in grupos_agrupados:
        print(f"Grupo: {nome_grupo}")
        print(grupo)
    
    # Inicializa o número do lançamento
    numero_lancamento = 1

    # Processa os grupos separados
    for (conta, situacao), grupo in grupos_separados:
        valor_total = grupo['ValorConsiderado'].sum()
        # Garante que valores das contas 8888 e 9999 sejam positivos
        if conta in ['8888', '9999']:
            valor_total = abs(valor_total)
        valor_formatado = f"{abs(valor_total):016.2f}".replace('.', '')

        # Formata a conta com zero à esquerda
        conta_formatada = f"{int(grupo['Contabil'].iloc[0]):05}"

        # Verifica se a conta é uma das específicas para o complemento
        complemento = ""
        complemento = complemento[:255]
        
        linha = f"200{numero_lancamento:02}{valor_formatado}C{conta_formatada}{'0408'}{complemento:<255}"
        lancamentos.append((linha, conta, situacao, valor_total, complemento))
        
        numero_lancamento += 1

    # Processa os grupos agrupados
    for conta, grupo in grupos_agrupados:
        valor_total = abs(grupo['ValorConsiderado'].sum())
        
        # Garante que valores das contas 8888 e 9999 sejam positivos
        if conta in ['8888', '9999']:
            valor_total = abs(valor_total)

        conta_formatada = f"{int(grupo['Contabil'].iloc[0]):05}"
        valor_formatado = f"{valor_total:016.2f}".replace('.', '')
        # Verifica se a conta é uma das específicas para o complemento
        complemento = "/".join(grupo['Documento'].astype(str).str.replace('.0', '', regex=False))
        
        complemento = complemento[:255]
        
        linha = f"200{numero_lancamento:02}{valor_formatado}C{conta_formatada}{'0408'}{complemento:<255}"
        lancamentos.append((linha, conta, '', valor_total, complemento))
        
        numero_lancamento += 1
    
    return lancamentos

def gerar_ultima_linha(lancamentos):
    total_linhas = len(lancamentos) + 2  # Adiciona 1 por causa da linha inicial "100..."
    valor_total = sum(l[3] for l in lancamentos)  # Soma todos os valores dos lançamentos
    valor_total_formatado = f"{valor_total:014.2f}".replace('.', '')
    ultima_linha = f"300{total_linhas:02}{(total_linhas -1):04}{valor_total_formatado}{valor_total_formatado}"
    return ultima_linha

def gerar_lancamento_debito_total(lancamentos):
    total_creditos = sum(l[3] for l in lancamentos)
    valor_formatado = f"{abs(total_creditos):016.2f}".replace('.', '')
    linha = f"200{len(lancamentos) + 1:02}{valor_formatado}D{'30991'}{'0409'}{'Total de Créditos':<255}"
    return (linha, '00000', '0409', total_creditos, 'Total de Créditos')

def salvar_arquivo_txt(conteudo):
    root = tk.Tk()
    root.withdraw()
    diretorio = filedialog.askdirectory(title="Selecione o diretório para salvar o arquivo")
    nome_arquivo = simpledialog.askstring("Nome do Arquivo", "Insira o nome do arquivo (sem extensão):")
    caminho_completo = os.path.join(diretorio, nome_arquivo + ".txt")
    
    with open(caminho_completo, 'w') as file:
        for linha in conteudo:
            file.write(linha[0] + "\n")
    
    print(f"Arquivo salvo em: {caminho_completo}")

def salvar_relatorio_excel(lancamentos):
    root = tk.Tk()
    root.withdraw()
    diretorio = filedialog.askdirectory(title="Selecione o diretório para salvar o relatório")
    nome_arquivo = simpledialog.askstring("Nome do Arquivo", "Insira o nome do relatório (sem extensão):")
    caminho_completo = os.path.join(diretorio, nome_arquivo + ".xlsx")
    
    # Cria DataFrame a partir dos lançamentos
    df_relatorio = pd.DataFrame(lancamentos, columns=['Linha', 'Conta', 'Situacao', 'Valor', 'Complemento'])
    df_relatorio.to_excel(caminho_completo, index=False)
    
    print(f"Relatório salvo em: {caminho_completo}")

def main():
    arquivo_path = escolher_arquivo()
    if not arquivo_path:
        print("Nenhum arquivo selecionado.")
        return
    
    planilha_escolhida = escolher_planilha(arquivo_path)
    data = obter_data()
    primeira_linha = gerar_primeira_linha(data)
    
    df = pd.read_excel(arquivo_path, sheet_name=planilha_escolhida)
    df.rename(columns=lambda x: x.strip(), inplace=True)

    # Lista as colunas do DataFrame para inspeção
    #print("Colunas disponíveis no DataFrame:")
    #print(df.columns)

    # Ajuste os nomes das colunas conforme necessário
    df = df[['Conta', 'Situacao', 'Valor', 'ValorPag', 'Contábil', 'Documento']]
    df.columns = ['Conta', 'Situacao', 'Valor', 'ValorPag', 'Contabil', 'Documento']
    
    # Converte as colunas Documento, Conta e Contabil para string e remove ".0"
    df['Documento'] = df['Documento'].astype(str).str.replace('.0', '', regex=False)
    df['Conta'] = df['Conta'].astype(str).str.replace('.0', '', regex=False)
    df['Contabil'] = df['Contabil'].astype(str).str.replace('.0', '', regex=False)
    
    linhas_contabeis = gerar_linhas_contabeis(df)
    lancamento_debito_total = gerar_lancamento_debito_total(linhas_contabeis)
    ultima_linha = gerar_ultima_linha(linhas_contabeis)
    
    conteudo = [(primeira_linha, '', '', 0, '')] + linhas_contabeis + [lancamento_debito_total] + [(ultima_linha, '', '', 0, '')]  # Adiciona a primeira linha, linhas contábeis, lançamento de débito e a última linha
    
    salvar_arquivo_txt(conteudo)
    
    # Salva o relatório em Excel com os dados das linhas contábeis geradas
    salvar_relatorio_excel(linhas_contabeis + [lancamento_debito_total])
    
    print("Processo concluído com sucesso!")

if __name__ == "__main__":
    main()
