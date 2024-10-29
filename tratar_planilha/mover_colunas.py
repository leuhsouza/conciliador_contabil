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

def mover_colunas_para_esquerda(file_path, output_path):
    """Move os dados de uma planilha para a esquerda, da coluna J até a Q para I até a P."""
    df = pd.read_excel(file_path, header=None)

    # Encontrar a linha onde a coluna C tem o valor "Numero" e a coluna I está vazia
    condicao = (df.iloc[:, 2] == 'Numero') & (df.iloc[:, 8].isna())

    if condicao.any():
        # Encontrar o índice da primeira linha que atende à condição
        indice_inicio = df[condicao].index[0]

        # Encontrar o índice da última linha com valor na coluna J (coluna 9)
        indice_fim = df.iloc[:, 2].last_valid_index()

        # Mover os dados das colunas J:Q (9:16) para I:P (8:15)
        df.iloc[indice_inicio:indice_fim+1, 8:16] = df.iloc[indice_inicio:indice_fim+1, 9:17].values
    
        # Remover a última coluna, que agora está vazia
        df = df.iloc[:, :-1]

        # Salvar o DataFrame modificado
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

    mover_colunas_para_esquerda(input_file, output_file)
    print(f"Arquivo processado salvo como {output_file}")

if __name__ == "__main__":
    main()
