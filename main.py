import pandas as pd
import os
import re

def clean_sheet_name(name):
    """Clean Excel sheet name for Excel rules"""
    invalid_chars = ['\\', '/', '?', '*', ':', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, ' ')
    return name[:31]

def unique_sheet_name(name, used_names):
    """Ensure sheet name is unique"""
    base = clean_sheet_name(name)
    sheet_name = base
    counter = 1

    while sheet_name in used_names:
        suffix = f"_{counter}"
        max_len = 31 - len(suffix)
        sheet_name = clean_sheet_name(base[:max_len] + suffix)
        counter += 1

    used_names.add(sheet_name)
    return sheet_name

def extract_question_base(column_name):
    """
    Extract main question text before '/'.
    Example:
    '3. Pergunta / opção' -> '3. Pergunta'
    """
    if '/' in column_name:
        return column_name.split('/')[0].strip()
    return column_name.strip()

def is_question_column(column_name):
    """
    Returns True if column starts with a question number like:
    1.
    3.
    7.1
    30.1
    """
    return bool(re.match(r'^\d+(?:\.\d+)?\s*\.', str(column_name).strip()))

def create_excel_tables_all_columns(input_csv, output_excel):
    # Read CSV
    try:
        df = pd.read_csv(input_csv, encoding='utf-8')
    except UnicodeDecodeError:
        df = pd.read_csv(input_csv, encoding='latin-1')

    grouped_columns = {}
    standalone_columns = []
    
    # Separate grouped question columns from standalone columns
    for column in df.columns:

        if is_question_column(column):
            base_question = extract_question_base(column)
            grouped_columns.setdefault(base_question, []).append(column)
        else:
            standalone_columns.append(column)

    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        used_sheet_names = set()

        # 1) Write one tab for each standalone column
        for column in standalone_columns:
            col_df = df[[column]].dropna(how='all')

            if not col_df.empty:
                sheet_name = unique_sheet_name(column, used_sheet_names)
                col_df.to_excel(writer, sheet_name=sheet_name, index=False)

                worksheet = writer.sheets[sheet_name]
                max_len = max(
                    len(str(column)),
                    col_df[column].astype(str).map(len).max() if not col_df[column].empty else 0
                )
                worksheet.set_column(0, 0, min(max_len + 2, 50))

        # 2) Write grouped question tabs
        for question, columns in grouped_columns.items():
            group_df = df[columns].dropna(how='all')

            if not group_df.empty:
                sheet_name = unique_sheet_name(question, used_sheet_names)
                group_df.to_excel(writer, sheet_name=sheet_name, index=False)

                worksheet = writer.sheets[sheet_name]
                for i, col in enumerate(group_df.columns):
                    max_len = max(
                        len(str(col)),
                        group_df[col].astype(str).map(len).max() if not group_df[col].empty else 0
                    )
                    worksheet.set_column(i, i, min(max_len + 2, 50))

    print(f"Arquivo Excel criado com sucesso: {output_excel}")


# Configurações
input_csv = "Peconheiros_2026_RMB.csv"
output_excel = "Mapa final - RMB - Peconheiros.xlsx"

# Executar
if os.path.exists(input_csv):
    create_excel_tables_all_columns(input_csv, output_excel)
else:
    print(f"Arquivo de entrada não encontrado: {input_csv}")