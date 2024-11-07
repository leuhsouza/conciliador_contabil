import pandas as pd
import tkinter as tk
from tkinter import filedialog
import re

def escolher_arquivo():
    root = tk.Tk()
    root.withdraw()
    arquivo_path = filedialog.askopenfilename(
        title="Selecione o arquivo para importação",
        filetypes=[("Arquivo XLS", "*.xls *.xlsx")]
    )
    return arquivo_path

def salvar_dataframe_intermediario(df, message="Salvar DataFrame intermediário como"):
    output_path = filedialog.asksaveasfilename(
        title=message,
        defaultextension=".xlsx",
        filetypes=[("Arquivo XLSX", "*.xlsx")]
    )
    if output_path:
        df.to_excel(output_path, index=False, header=False)
        print(f"DataFrame intermediário salvo como {output_path}")
    else:
        print("Nenhum local de salvamento selecionado.")

def process_excel(file_path, output_path):
    # Carregar o arquivo Excel sem cabeçalhos
    df = pd.read_excel(file_path, header=None)

    # Inicializar índice atual para verificar a coluna H
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

        # Resetar o índice após a remoção e reiniciar o loop
        df.reset_index(drop=True, inplace=True)

    # Inicializar índice atual para verificar a coluna E
    current_index = 0

    while current_index < len(df):
        # Verificar a primeira ocorrência da string "CLUBE CURITIBANO" na coluna E
        index_to_delete = df.iloc[current_index:, 4].apply(lambda x: "CLUBE CURITIBANO" in str(x)).idxmax()

        if index_to_delete == 0 and "CLUBE CURITIBANO" not in str(df.iloc[0, 4]):
            break

        # Calcular o intervalo de linhas a serem deletadas
        rows_to_delete = range(index_to_delete, index_to_delete + 8)

        # Remover as linhas do DataFrame
        df.drop(rows_to_delete, inplace=True, errors='ignore')

        # Resetar o índice após a remoção e reiniciar o loop
        df.reset_index(drop=True, inplace=True)
        current_index = 0  # Reiniciar o loop do início após a deleção

    df.dropna(how='all', inplace=True)
    df = df[df.iloc[:, 12] != 'Total:']
    df.dropna(axis=1, how='all', inplace=True)

    # Salvar o DataFrame intermediário após a remoção de colunas
    salvar_dataframe_intermediario(df, "Salvar DataFrame após a remoção de colunas como")

    # Verificar e unir linhas com base na coluna A (índice 0)
    for i in range(len(df) - 1, 0, -1):
        if pd.notna(df.iloc[i, 1]) and pd.isna(df.iloc[i, 0]):  # Verificando coluna A
            df.iloc[i-1, 1] = str(df.iloc[i-1, 1]) + " " + str(df.iloc[i, 1])
            df.reset_index(drop=True, inplace=True)
            df.drop(i, inplace=True, errors='ignore')

    df.reset_index(drop=True, inplace=True)

    # Inicializar a variável de conta contábil
    conta_contábil = None

    # Preencher a coluna 'Conta Contábil' com base na coluna de índice 4
    for index, row in df.iterrows():
        if pd.notna(row[4]) and isinstance(row[4], str):
            # Remover múltiplos espaços em branco
            row[4] = re.sub(r'\s+', ' ', row[4]).strip()
            # Verificar se a string contém "Red.:" e extrair a conta contábil de 5 dígitos mais o dígito após o hífen
            if row[4].startswith("Red.:"):
                match = re.search(r"Red\.\:\s(\d{4})-(\d)", row[4])
                if match:
                    conta_contábil = match.group(1) + match.group(2)  # Concatena o valor de 4 dígitos com o dígito após o hífen
                    print(f"Conta contábil encontrada: {conta_contábil} na linha {index}")
        # Adicionar a conta contábil em uma nova coluna na última posição do DataFrame
        df.at[index, 'Conta Contábil'] = conta_contábil

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
