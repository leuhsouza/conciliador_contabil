import pandas as pd
import sqlite3
import re



def ensure_conciliated_column(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Verifica se a coluna 'conciliada' já existe na tabela 'dados'
    cursor.execute("PRAGMA table_info(dados)")
    columns = cursor.fetchall()
    column_names = [column[1] for column in columns]
    
    if 'conciliada' not in column_names:
        cursor.execute('ALTER TABLE dados ADD COLUMN conciliada BOOLEAN DEFAULT 0')
        print("Coluna 'conciliada' criada com sucesso.")
    else:
        print("Coluna 'conciliada' já existe.")
    
    conn.commit()
    conn.close()

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

    # Remove colunas totalmente vazias, exceto a coluna de índice 10
    for col in df.columns:
        if col not in [10, 13] and df[col].isna().all():
            df.drop(columns=col, inplace=True)


    for i in range(len(df) - 1, 0, -1):
        if pd.notna(df.iloc[i, 1]) and pd.isna(df.iloc[i, 3]):
            df.iloc[i-1, 1] = str(df.iloc[i-1, 1]) + " " + str(df.iloc[i, 1])
            df.reset_index(drop=True, inplace=True)
            df.drop(i, inplace=True, errors='ignore')

    df.reset_index(drop=True, inplace=True)

    # Adicionar a conta contábil em todas as linhas para o DataFrame do banco de dados
    if conta_contabil:
        df['Conta'] = conta_contabil  # Adicionar a coluna 'Conta'

# Converte a coluna de índice 2 (data) para o formato esperado pelo SQLite
    if 2 in df.columns:
        df[2] = pd.to_datetime(df[2], format='%d/%m/%Y', errors='coerce')  # Converte para datetime
        df[2] = df[2].dt.strftime('%Y-%m-%d')  # Formata para o padrão YYYY-MM-DD

    df.reset_index(drop=True, inplace=True)
    pd.set_option('display.max_columns', None)
    print(df.head(20))

    # Criar o DataFrame para o banco de dados sem a coluna de saldo (assumindo que a coluna de saldo seja a coluna 7)
    df_db = df.drop(df.columns[7], axis=1)  # Remove a coluna de saldo e a coluna D/C
    df_db.drop(df_db.columns[7], axis=1, inplace=True) # Remove a coluna de D/C #Após a ultima deleção o indice do D/C é 7

    # Verificar os itens do DataFrame
    pd.set_option('display.max_columns', None)
    print(df_db.head(20))

    # Verificar a quantidade de colunas após a remoção da coluna de saldo
    print(f"Número de colunas em df_db: {df_db.shape[1]}")
    print("Colunas do DataFrame para o banco de dados:", df_db.columns)

    print(df_db.columns)
    #preparando df para importação   
 

    # Ajustar o DataFrame para ter exatamente 8 colunas (excluindo saldo e incluindo conta)
    if df_db.shape[1] != 8:
        raise ValueError(f"O DataFrame para o banco de dados não tem 8 colunas. Ele tem {df_db.shape[1]} colunas.")

    linhas_importadas = 0
    
    if enviar_bd:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS dados (
                id INTEGER PRIMARY KEY,
                data DATE,
                historico TEXT,
                contra_partida TEXT,
                lote TEXT,
                lancamento TEXT,
                d TEXT,
                c TEXT,
                conta TEXT,
                conciliada BOOLEAN DEFAULT 0,
                UNIQUE(lancamento, lote)
            )
        ''')

        cursor.execute('SELECT MAX(id) FROM dados')
        max_id = cursor.fetchone()[0]
        next_id = (max_id + 1) if max_id is not None else 1

        for row in df_db.itertuples(index=False, name=None):
            print(f"Tentando inserir linha: {row}")
            cursor.execute('''
                INSERT OR IGNORE INTO dados (id, data, historico, contra_partida, lote, lancamento, d, c, conta, conciliada)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?,  ?, ?)
            ''', (next_id, *row, 0))  # Adiciona 0 para a coluna 'conciliada'
            print(f"Rowcount após inserção: {cursor.rowcount}")
            if cursor.rowcount > 0:
                next_id += 1
                linhas_importadas += 1

        conn.commit()
        conn.close()

    if output_path:
        # Salvar o DataFrame completo com a coluna de saldo
        df.to_excel(output_path, index=False, header=False)

    ensure_conciliated_column(db_path)

    return linhas_importadas

def process_excel_varias_contas(file_path, db_path=None, output_path=None, enviar_bd=False):
    # Carregar o arquivo Excel sem cabeçalhos
    df = pd.read_excel(file_path, header=None)

    # Inicializar índice atual para verificar a coluna H
    current_index = 0
    while current_index < len(df):
        index_to_delete = df.iloc[current_index:, 7].first_valid_index()
        if index_to_delete is None:
            break
        rows_to_delete = range(index_to_delete, index_to_delete + 6)
        df.drop(rows_to_delete, inplace=True, errors='ignore')
        df.reset_index(drop=True, inplace=True)

    # Inicializar índice atual para verificar a coluna E
    current_index = 0
    while current_index < len(df):
        index_to_delete = df.iloc[current_index:, 4].apply(lambda x: "CLUBE CURITIBANO" in str(x)).idxmax()
        if index_to_delete == 0 and "CLUBE CURITIBANO" not in str(df.iloc[0, 4]):
            break
        rows_to_delete = range(index_to_delete, index_to_delete + 6)
        df.drop(rows_to_delete, inplace=True, errors='ignore')
        df.reset_index(drop=True, inplace=True)
        current_index = 0

    df.dropna(how='all', inplace=True)
    df = df[df.iloc[:, 12] != 'Total:']
    df = df[df.iloc[:, 14] != 'Transporte da página anterior:']
    df.dropna(axis=1, how='all', inplace=True)

    # Verificar e unir linhas com base na coluna A (índice 0)
    for i in range(len(df) - 1, 0, -1):
        if pd.notna(df.iloc[i, 1]) and pd.isna(df.iloc[i, 0]):
            df.iloc[i-1, 1] = str(df.iloc[i-1, 1]) + " " + str(df.iloc[i, 1])
            df.reset_index(drop=True, inplace=True)
            df.drop(i, inplace=True, errors='ignore')

    df.reset_index(drop=True, inplace=True)

    # Capturar e adicionar a conta contábil em cada linha do DataFrame
    conta_contabil = None
    indices_para_remover = []  # Lista para armazenar os índices das linhas a serem removidas

# Converte a coluna de índice 2 (data) para o formato esperado pelo SQLite
    if 2 in df.columns:
        df[2] = pd.to_datetime(df[2], format='%d/%m/%Y', errors='coerce')  # Converte para datetime
        df[2] = df[2].dt.strftime('%Y-%m-%d')  # Formata para o padrão YYYY-MM-DD

    for index, row in df.iterrows():
        if pd.notna(row[4]) and isinstance(row[4], str):
            row[4] = re.sub(r'\s+', ' ', row[4]).strip()
            if "Red.:" in row[4]:
                match = re.search(r"Red\.\:\s(\d{4})-(\d)", row[4])
                if match:
                    conta_contabil = match.group(1) + match.group(2)
                    print(f"Conta contábil encontrada: {conta_contabil} na linha {index}")
                    indices_para_remover.append(index)  # Adiciona o índice para remoção
        df.at[index, 'Conta'] = conta_contabil  # Adiciona a conta contábil à linha

    # Mostrar os números das linhas deletadas e as colunas A e B
    print("Linhas que serão removidas (número da linha, coluna A, coluna B):")
    for idx in indices_para_remover:
        print(f"Linha {idx}: A = {df.iloc[idx, 0]}, B = {df.iloc[idx, 1]}")

    # Remover as linhas identificadas do DataFrame para o banco de dados
    df_db = df.drop(indices_para_remover).copy()

    print(df_db.columns)
    # Preparar DataFrame para o banco de dados
    df_db.drop(df_db.columns[2], axis=1, inplace=True)  # Remove a coluna de índice 4
    df_db.drop(df_db.columns[7], axis=1, inplace=True)  # Remove a coluna de saldo
    # ver sobre o datraframe está excluindo no dataframe de exportação.
    #df_db.drop(df_db.columns[7], axis=1, inplace=True) # Remove a coluna de D/C

    #verificar os istens do dataframe
    pd.set_option('display.max_columns', None)
    print(df.head(20))
    # Verificar a quantidade de colunas após a remoção da coluna de saldo
    print(f"Número de colunas em df_db: {df_db.shape[1]}")

    # Verificar se o DataFrame para o banco de dados tem linhas
    print(f"Número de linhas em df_db: {len(df_db)}")
    if len(df_db) == 0:
        print("O DataFrame para o banco de dados está vazio. Nenhuma linha para importar.")
        return 0  # Retorna 0 para indicar que nenhuma linha foi importada

    if enviar_bd and db_path:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS dados (
                id INTEGER PRIMARY KEY,
                data DATE,
                historico TEXT,
                contra_partida TEXT,
                lote TEXT,
                lancamento TEXT,
                d TEXT,
                c TEXT,
                conta TEXT,
                conciliada BOOLEAN DEFAULT 0,
                UNIQUE(lancamento, lote)
            )
        ''')

        cursor.execute('SELECT MAX(id) FROM dados')
        max_id = cursor.fetchone()[0]
        next_id = (max_id + 1) if max_id is not None else 1

        linhas_importadas = 0
        for row in df_db.itertuples(index=False, name=None):
            cursor.execute('''
                INSERT OR IGNORE INTO dados (id, data, historico, contra_partida, lote, lancamento, d, c, conta, conciliada)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (next_id, *row, 0)) # Adiciona 0 para a coluna 'conciliada'
            if cursor.rowcount > 0:
                next_id += 1
                linhas_importadas += 1

        conn.commit()
        conn.close()
        print(f"Total de linhas importadas: {linhas_importadas}")
        return linhas_importadas

    if output_path:
        df_output = df#.drop('Conta', axis=1)
        df_output.to_excel(output_path, index=False, header=False)
        print(f"Arquivo processado salvo como {output_path}")

    ensure_conciliated_column(db_path)


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

def process_lote(file_path, db_path=None, output_path=None, enviar_bd=False):
    #Processa o arquivo Excel, excluindo linhas e formatando o arquivo.
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

    if output_path:
        df_output = df#.drop('Conta', axis=1)
        df_output.to_excel(output_path, index=False, header=False)
    print(f"Arquivo processado salvo como {output_path}")