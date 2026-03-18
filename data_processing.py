import pandas as pd
import matplotlib.pyplot as plt
import os
import re
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter

# 1. Configurações Iniciais
caminho_arquivo_entrada = 'resultado_do_script_do_prince.xlsx' #Mude o nome dessa linha para o nome do arquivo original
caminho_arquivo_saida = 'resultado_completo_graficos.xlsx'

# Lista para guardar o nome das imagens temporárias e apagá-las no final
imagens_temporarias = []
nomes_abas_usados = []

try:
    # Lendo todas as abas do Excel de uma vez. 
    # sheet_name=None retorna um dicionário onde: chave = Nome da Aba, valor = DataFrame (os dados)
    dfs = pd.read_excel(caminho_arquivo_entrada, sheet_name=None)
    
    print(f"Lendo o arquivo '{caminho_arquivo_entrada}'...")
    print(f"Foram encontradas {len(dfs)} abas. Iniciando processamento...\n")

    # Inicializa o escritor do Excel
    with pd.ExcelWriter(caminho_arquivo_saida, engine='openpyxl') as writer:
        
        for nome_aba_original, df in dfs.items():
            # Se a aba estiver vazia, pula para a próxima
            if df.empty or len(df.columns) == 0:
                print(f"Aba '{nome_aba_original}' está vazia. Ignorando.")
                continue

            # Assumimos que a coluna a ser contada é a PRIMEIRA coluna da aba (índice 0)
            # Se a sua coluna alvo estiver em outro lugar, você pode alterar o índice abaixo
            nome_coluna_alvo = df.columns[0]
            
            # --- Regras do Excel para Nomes de Abas ---
            # 1. Máximo de 31 caracteres
            # 2. Não pode conter os caracteres: \ / * ? : [ ]
            novo_nome_aba = re.sub(r'[\\/*?:\[\]]', '', str(nome_coluna_alvo))
            novo_nome_aba = novo_nome_aba[:30] # Limita a 30 caracteres
            
            # Evita erro caso duas abas acabem com o exato mesmo nome
            if novo_nome_aba in nomes_abas_usados:
                novo_nome_aba = f"{novo_nome_aba}_{len(nomes_abas_usados)}"
            nomes_abas_usados.append(novo_nome_aba)

            print(f"Processando: '{nome_aba_original}' -> Criando aba '{novo_nome_aba}'")

            # 1. Escreve os DADOS ORIGINAIS na nova aba (começando na coluna A)
            df.to_excel(writer, sheet_name=novo_nome_aba, index=False)
            
            # 2. Realiza a contagem (CONT.SE)
            contagem = df[nome_coluna_alvo].value_counts()
            porcentagem = df[nome_coluna_alvo].value_counts(normalize=True) * 100

            tabela_resumo = pd.DataFrame({
                'Quantidade': contagem,
                '%': porcentagem.round(2)
            })

            tabela_resumo.loc['Total'] = [contagem.sum(), porcentagem.sum().round(2)]

            # Calcula onde colocar a tabela de resumo (ao lado dos dados originais)
            # Se a tabela original tem 1 coluna, coloca a tabela na coluna 3 (C)
            coluna_inicio_resumo = df.shape[1] + 1 
            
            # Escreve a tabela de resumo dando um espaço das colunas originais
            tabela_resumo.to_excel(writer, sheet_name=novo_nome_aba, startcol=coluna_inicio_resumo, index_label='Respostas')

            # 3. Gerar o Gráfico de Pizza
            plt.figure(figsize=(7, 5))
            dados_grafico = tabela_resumo.drop('Total')
            plt.pie(dados_grafico['Quantidade'], 
                    labels=dados_grafico.index, 
                    autopct='%1.1f%%', 
                    startangle=140, 
                    colors=plt.cm.Set3.colors,
                    pctdistance=0.85)

            # Define o título do gráfico com o nome da coluna
            plt.title(f'{nome_coluna_alvo}', fontsize=12)
            plt.tight_layout()

            # Salva imagem temporária
            caminho_img = f'temp_grafico_{len(nomes_abas_usados)}.png'
            plt.savefig(caminho_img, dpi=120)
            plt.close()
            imagens_temporarias.append(caminho_img)

            # 4. Inserir a imagem e ajustar a formatação na planilha
            worksheet = writer.sheets[novo_nome_aba]
            img = Image(caminho_img)
            
            # Pula a largura da tabela de resumo para não colocar a imagem por cima do texto
            coluna_imagem_idx = coluna_inicio_resumo + 4 
            celula_imagem = f"{get_column_letter(coluna_imagem_idx)}2" # Ex: G2, H2, etc
            worksheet.add_image(img, celula_imagem)

            # Ajustar a largura das colunas do resumo para ficar bonito
            # get_column_letter(1) -> 'A', get_column_letter(2) -> 'B', etc.
            worksheet.column_dimensions[get_column_letter(coluna_inicio_resumo + 1)].width = 45 # Coluna 'Respostas'
            worksheet.column_dimensions[get_column_letter(coluna_inicio_resumo + 2)].width = 12 # Coluna 'Quantidade'
            worksheet.column_dimensions[get_column_letter(coluna_inicio_resumo + 3)].width = 10 # Coluna '%'

    # 5. Limpeza Final (apagar imagens temporárias do PC)
    for img_path in imagens_temporarias:
        if os.path.exists(img_path):
            os.remove(img_path)

    print(f"\nSucesso! O arquivo '{caminho_arquivo_saida}' foi gerado com todas as abas, tabelas e gráficos ao lado.")

except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")