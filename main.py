import pandas as pd
import os

def create_excel_tables(input_csv, output_excel):
    """Cria um arquivo Excel com uma aba para cada questão usando apenas pandas"""

    # Ler o arquivo CSV
    try:
        df = pd.read_csv(input_csv, encoding='utf-8')
    except:
        try:
            df = pd.read_csv(input_csv, encoding='latin-1')
        except Exception as e:
            print(f"Erro ao ler o arquivo CSV: {e}")
            return

    # Criar um escritor Excel
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        # Para cada coluna (questão) no DataFrame
        for column in df.columns:
            # Criar um DataFrame com apenas essa coluna, removendo linhas vazias
            question_df = df[[column]].dropna(how='all')

            # Se houver dados, escrever em uma nova aba
            if not question_df.empty:
                # Limitar o nome da aba a 31 caracteres e remover caracteres inválidos
                sheet_name = column[:31]
                invalid_chars = ['\\', '/', '?', '*', ':', '[', ']']
                for char in invalid_chars:
                    sheet_name = sheet_name.replace(char, ' ')

                # Escrever no Excel
                question_df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Ajustar a largura das colunas
                worksheet = writer.sheets[sheet_name]
                worksheet.set_column('A:A', 50)  # Ajuste a largura conforme necessário

    print(f"Arquivo Excel criado com sucesso: {output_excel}")

# Configurações
input_csv = "Abaetetuba(PA).csv"
output_excel = "Respostas_Abaetetuba.xlsx"

# Executar a função principal
if os.path.exists(input_csv):
    create_excel_tables(input_csv, output_excel)
else:
    print(f"Arquivo de entrada não encontrado: {input_csv}")
