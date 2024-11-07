import pandas as pd
import sqlite3

def process_excel(file_path, db_path, output_path=None):
    df = pd.read_excel(file_path, header=None)

    current_index = 0
    while current_index < len(df):
        index_to_delete = df.iloc[current_index:, 7].first_valid_index()
        if index_to_delete is None:
            break
        rows_to_delete = range(index_to_delete, index_to_delete + 8)
        df.drop(rows_to_delete, inplace=True, errors='ignore')
        df.reset_index(drop=True, inplace=True)
        #current_index = index_to_delete + 9

    df.dropna(how='all', inplace=True)
    df = df[df.iloc[:, 12] != 'Total:']
    df.dropna(axis=1, how='all', inplace=True)

    for i in range(len(df) - 1, 0, -1):
        if pd.notna(df.iloc[i, 1]) and pd.isna(df.iloc[i, 3]):
            df.iloc[i-1, 1] = str(df.iloc[i-1, 1]) + " " + str(df.iloc[i, 1])
            df.reset_index(drop=True, inplace=True)
            df.drop(i, inplace=True, errors='ignore')



    # Garantir que o DataFrame tenha exatamente 9 colunas
    if df.shape[1] < 9:
        for _ in range(9 - df.shape[1]):
            df[f'coluna{df.shape[1] + 1}'] = None
    elif df.shape[1] > 9:
        df = df.iloc[:, :9]

    # Remover linhas que são completamente vazias
    df.dropna(how='all', inplace=True)

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS dados (
            id INTEGER PRIMARY KEY,
            coluna1 TEXT, coluna2 TEXT, coluna3 TEXT, coluna4 TEXT,
            coluna5 TEXT, coluna6 TEXT, coluna7 TEXT, coluna8 TEXT,
            coluna9 TEXT,
            UNIQUE(coluna4, coluna5)
        )
    ''')

    # Obter o próximo ID disponível
    cursor.execute('SELECT MAX(id) FROM dados')
    max_id = cursor.fetchone()[0]
    next_id = (max_id + 1) if max_id is not None else 1

    for row in df.itertuples(index=False, name=None):
        cursor.execute('''
            INSERT OR IGNORE INTO dados (id, coluna1, coluna2, coluna3, coluna4, coluna5, coluna6, coluna7, coluna8, coluna9)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (next_id, *row))
        if cursor.rowcount > 0:  # Se a linha foi inserida
            next_id += 1

    conn.commit()
    conn.close()

    if output_path:
        df.to_excel(output_path, index=False, header=False)
