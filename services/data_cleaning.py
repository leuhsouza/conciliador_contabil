import pandas as pd
import sqlite3
import re

def process_excel(file_path, db_path, output_path=None, enviar_bd=False):
    df = pd.read_excel(file_path, header=None)

    # Capturar a conta contábil no início, verificando a coluna E (índice 4)
    conta_contabil = None
    for index, row in df.iterrows():
        if pd.notna(row[4]) and isinstance(row[4], str):  # Verifica a coluna E
            row[4] = re.sub(r'\s+', ' ', row[4]).strip()
            if "Red.:" in row[4]:
                match = re.search(r"Red\.\:\s(\d{4})-(\d)", row[4])
                if match:
                    conta_contabil = match.group(1) + match.group(2)
                    print(f"Conta contábil capturada: {conta_contabil} na linha {index}")
                    break

    if conta_contabil:
        print("Conta contábil foi capturada com sucesso.")
    else:
        print("Nenhuma conta contábil foi capturada.")

    # Processamento dos dados conforme a lógica existente
    current_index = 0
    while current_index < len(df):
        index_to_delete = df.iloc[current_index:, 7].first_valid_index()
        if index_to_delete is None:
            break
        rows_to_delete = range(index_to_delete, index_to_delete + 8)
        df.drop(rows_to_delete, inplace=True, errors='ignore')
        df.reset_index(drop=True, inplace=True)

    df.dropna(how='all', inplace=True)
    df = df[df.iloc[:, 12] != 'Total:']
    df.dropna(axis=1, how='all', inplace=True)

    for i in range(len(df) - 1, 0, -1):
        if pd.notna(df.iloc[i, 1]) and pd.isna(df.iloc[i, 3]):
            df.iloc[i-1, 1] = str(df.iloc[i-1, 1]) + " " + str(df.iloc[i, 1])
            df.reset_index(drop=True, inplace=True)
            df.drop(i, inplace=True, errors='ignore')

    df.reset_index(drop=True, inplace=True)

    # Adicionar a conta contábil em todas as linhas para o DataFrame do banco de dados
    if conta_contabil:
        df['Conta'] = conta_contabil  # Adicionar a coluna 'Conta'

    # Criar o DataFrame para o banco de dados sem a coluna de saldo (assumindo que a coluna de saldo seja a coluna 7)
    df_db = df.drop(df.columns[7], axis=1)  # Remove a coluna de saldo

    # Verificar a quantidade de colunas após a remoção da coluna de saldo
    print(f"Número de colunas em df_db: {df_db.shape[1]}")
    print("Colunas do DataFrame para o banco de dados:", df_db.columns)

    # Ajustar o DataFrame para ter exatamente 9 colunas (excluindo saldo e incluindo conta)
    if df_db.shape[1] != 9:
        raise ValueError(f"O DataFrame para o banco de dados não tem 9 colunas. Ele tem {df_db.shape[1]} colunas.")

    linhas_importadas = 0

    if enviar_bd:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS dados (
                id INTEGER PRIMARY KEY,
                data TEXT,
                historico TEXT,
                contra_partida TEXT,
                lote TEXT,
                lancamento TEXT,
                d TEXT,
                c TEXT,
                dc TEXT,
                conta TEXT,
                UNIQUE(lancamento, lote)
            )
        ''')

        cursor.execute('SELECT MAX(id) FROM dados')
        max_id = cursor.fetchone()[0]
        next_id = (max_id + 1) if max_id is not None else 1

        for row in df_db.itertuples(index=False, name=None):
            print(f"Tentando inserir linha: {row}")
            cursor.execute('''
                INSERT OR IGNORE INTO dados (id, data, historico, contra_partida, lote, lancamento, d, c, dc, conta)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (next_id, *row))
            print(f"Rowcount após inserção: {cursor.rowcount}")
            if cursor.rowcount > 0:
                next_id += 1
                linhas_importadas += 1

        conn.commit()
        conn.close()

    if output_path:
        # Salvar o DataFrame completo com a coluna de saldo
        df.to_excel(output_path, index=False, header=False)

    return linhas_importadas
