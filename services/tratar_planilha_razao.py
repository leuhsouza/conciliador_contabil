import pandas as pd
import tkinter as tk
from tkinter import filedialog

def escolher_arquivo():
    root = tk.Tk()
    root.withdraw()
    arquivo_path = filedialog.askopenfilename(
        title="Selecione o arquivo para importação",
        filetypes=[("Arquivo XLS", "*.xls *.xlsx")]
    )
    return arquivo_path

def process_excel(file_path, output_path):
    # Carregar o arquivo Excel sem cabeçalhos
    df = pd.read_excel(file_path, header=None)

    # Inicializar índice atual
    current_index = 0

    while current_index < len(df):
        # Encontrar o próximo índice onde a coluna H tem um valor
        index_to_delete = df.iloc[current_index:, 7].first_valid_index()

        if index_to_delete is None:
            break

        # Calcular o intervalo de linhas a serem deletadas
        rows_to_delete = range(index_to_delete, index_to_delete + 8)

        # Remover as linhas do DataFrame
        df.drop(rows_to_delete, inplace=True, errors='ignore')

        # Atualizar o índice atual para continuar a busca
        df.reset_index(drop=True, inplace=True)

    # Remover linhas em branco
    df.dropna(how='all', inplace=True)

    # Remover a linha que contém "Total:" na coluna M (índice 12)
    df = df[df.iloc[:, 12] != 'Total:']

    # Remover colunas em branco
    df.dropna(axis=1, how='all', inplace=True)

    # Conferir linha a linha começando da última linha com valor na coluna B
    for i in range(len(df) - 1, 0, -1):
        if pd.notna(df.iloc[i, 1]) and pd.isna(df.iloc[i, 3]):
            df.iloc[i-1, 1] = str(df.iloc[i-1, 1]) + " " + str(df.iloc[i, 1])
            print(df.iloc[i,1])
            df.reset_index(drop=True, inplace=True)
            df.drop(i, inplace=True, errors='ignore')

    # Resetar o índice do DataFrame após as operações
    df.reset_index(drop=True, inplace=True)

    # Salvar o DataFrame modificado em um novo arquivo Excel sem cabeçalhos
    df.to_excel(output_path, index=False, header=False)

def main():
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
