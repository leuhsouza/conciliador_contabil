import pandas as pd
import tkinter as tk
from tkinter import filedialog

def escolher_arquivo():
    """Abre uma janela para selecionar um arquivo de entrada."""
    root = tk.Tk()
    root.withdraw()
    arquivo_path = filedialog.askopenfilename(
        title="Selecione o arquivo para importação",
        filetypes=[("Arquivo XLS", "*.xls *.xlsx")]
    )
    return arquivo_path

def mover_colunas_para_esquerda(df):
    """Move os dados de uma planilha para a esquerda, a partir da coluna B, quando a condição for atendida."""
    # Encontrar a linha onde a coluna C tem o valor "Numero" e a coluna I está vazia
    condicao = (df.iloc[:, 2] == 'Numero') & (df.iloc[:, 8].isna())

    if condicao.any():
        # Encontrar o índice da primeira linha que atende à condição
        indice_inicio = df[condicao].index[0]

        # Encontrar o índice da última linha com valor na coluna C
        indice_fim = df.iloc[:, 2].last_valid_index()

        # Recuar o índice do cabeçalho em uma coluna
        indice_cabecalho = indice_inicio - 6
        print(indice_cabecalho)

        # Ajustar o valor da célula do cabeçalho
        df.iloc[indice_cabecalho:indice_cabecalho + 1, 6:16] = df.iloc[indice_cabecalho:indice_cabecalho + 1, 7:17].values

        # Mover os dados para a esquerda (até a coluna B)
        df.iloc[indice_inicio:indice_fim + 1, 8:16] = df.iloc[indice_inicio:indice_fim + 1, 9:17].values

        # Remover a última coluna, que agora está vazia
        df = df.iloc[:, :-1]


    # Retorna o DataFrame modificado
    return df

def process_excel(file_path, output_path):
    """Processa o arquivo Excel, excluindo linhas e formatando o arquivo."""
    # Carregar o arquivo Excel sem cabeçalhos
    df = pd.read_excel(file_path, header=None)

    # Chamar a função para mover colunas para a esquerda e obter o DataFrame modificado
    df = mover_colunas_para_esquerda(df)

    # Inicializar índice atual
    current_index = 0

    while current_index < len(df):
        # Encontrar o próximo índice onde a coluna H tem um valor
        index_to_delete = df.iloc[current_index:, 6].first_valid_index()

        if index_to_delete is None:
            break

        # Calcular o intervalo de linhas a serem deletadas (7 linhas a partir do índice encontrado)
        rows_to_delete = range(index_to_delete, index_to_delete + 7)

        # Remover as linhas do DataFrame
        df.drop(rows_to_delete, inplace=True, errors='ignore')

        # Atualizar o índice atual para continuar a busca
        #current_index = index_to_delete + 8
        df.reset_index(drop=True, inplace=True)

    # Remover linhas completamente em branco
    df.dropna(how='all', inplace=True)

    # Remover a linha que contém "Total:" na coluna M (índice 12)
    df = df[df.iloc[:, 12] != 'Total:']

    # Remover colunas completamente em branco
    df.dropna(axis=1, how='all', inplace=True)

    # Conferir linha a linha começando da última linha com valor na coluna F (índice 5)
    for i in range(len(df) - 1, 0, -1):
        print(df.iloc[i,5])
        if pd.notna(df.iloc[i, 5]) and pd.isna(df.iloc[i, 1]):
            # Concatenar os valores da linha atual com a anterior na coluna F
            df.iloc[i-1, 5] = str(df.iloc[i-1, 5]) + " " + str(df.iloc[i, 5])
            df.reset_index(drop=True, inplace=True)
            # Remover a linha atual
            df.drop(i, inplace=True, errors='ignore')

    # Resetar o índice do DataFrame após as operações
    df.reset_index(drop=True, inplace=True)

    # Salvar o DataFrame modificado em um novo arquivo Excel sem cabeçalhos
    df.to_excel(output_path, index=False, header=False)

def main():
    """Função principal para escolha de arquivo e processamento."""
    input_file = escolher_arquivo()
    if not input_file:
        print("Nenhum arquivo selecionado")
        return
    
    output_file = filedialog.asksaveasfilename(
        title="Salvar arquivo processado como",
        defaultextension=".xlsx",
        filetypes=[("Arquivo XLSX", "*.xlsx")]
    )
    if not output_file:
        print("Nenhum local de salvamento selecionado")
        return

    process_excel(input_file, output_file)
    print(f"Arquivo processado salvo como {output_file}")

if __name__ == "__main__":
    main()
