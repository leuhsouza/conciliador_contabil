import pandas as pd

def obter_data(data_str):
    # Verifica se a data está no formato DDMMAAAA
    while not (len(data_str) == 8 and data_str.isdigit()):
        raise ValueError("Data inválida. Use o formato DDMMAAAA.")
    return data_str

def gerar_primeira_linha(data):
    primeira_linha = "100" + "00" + "07" + "02" + data
    return primeira_linha