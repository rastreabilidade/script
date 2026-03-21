import pandas as pd
import matplotlib.pyplot as plt
import os
import re
import textwrap
import numpy as np
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter

# 1. Configurações Iniciais
caminho_arquivo_entrada = 'Mapa final - Castanhal - Produtores.xlsx' # Mude para o nome do arquivo original
caminho_arquivo_saida = 'Testefinal.xlsx' # Nome do arquivo que será salvo
limite_caracteres_titulo = 75 # Limite de letras no título antes de quebrar a linha

# Lista para guardar o nome das imagens temporárias e apagá-las no final
imagens_temporarias = []
nomes_abas_usados = []

try:
    # Lendo todas as abas do Excel de uma vez. 
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
            nome_coluna_alvo = df.columns[0]
            
            # --- Regras do Excel para Nomes de Abas ---
            novo_nome_aba = re.sub(r'[\\/*?:\[\]]', '', str(nome_coluna_alvo))
            novo_nome_aba = novo_nome_aba[:30] # Limita a 30 caracteres
            
            # Evita erro caso duas abas acabem com o exato mesmo nome
            if novo_nome_aba in nomes_abas_usados:
                novo_nome_aba = f"{novo_nome_aba}_{len(nomes_abas_usados)}"
            nomes_abas_usados.append(novo_nome_aba)

            print(f"Processando: '{nome_aba_original}' -> Criando aba '{novo_nome_aba}'")

            # 1. Escreve os DADOS ORIGINAIS na nova aba
            df.to_excel(writer, sheet_name=novo_nome_aba, index=False)
            
            # 2. Realiza a contagem e porcentagem
            contagem = df[nome_coluna_alvo].value_counts()
            porcentagem = df[nome_coluna_alvo].value_counts(normalize=True) * 100

            tabela_resumo = pd.DataFrame({
                'Quantidade': contagem,
                '%': porcentagem.round(2)
            })

            tabela_resumo.loc['Total'] = [contagem.sum(), porcentagem.sum().round(2)]

            # Calcula onde colocar a tabela de resumo (ao lado dos dados originais)
            coluna_inicio_resumo = df.shape[1] + 1 
            
            # Escreve a tabela de resumo
            tabela_resumo.to_excel(writer, sheet_name=novo_nome_aba, startcol=coluna_inicio_resumo, index_label='Respostas')

           # 3. Gerar o Gráfico de Pizza Misto (Dentro e Fora)
            # AUMENTAMOS a imagem (14 de largura, 8 de altura) para dar espaço de sobra para a legenda
            fig, ax = plt.subplots(figsize=(14, 8), subplot_kw=dict(aspect="equal"))
            dados_grafico = tabela_resumo.drop('Total', errors='ignore')
            
            # Pega o total de respostas
            total = dados_grafico['Quantidade'].sum()
            
            # Limite para decidir quem fica dentro e quem vai para fora
            limite_percentual = 2.0 
            
            # 3.1 Função: Só escreve DENTRO se for MAIOR que o limite
            def formato_porcentagem_dentro(pct):
                return ('%1.1f%%' % pct) if pct > limite_percentual else ''

            # Cria a pizza
            wedges, texts, autotexts = ax.pie(
                dados_grafico['Quantidade'], 
                autopct=formato_porcentagem_dentro, 
                startangle=140, 
                colors=plt.cm.Set3.colors,
                pctdistance=0.75
            )

 # --- 3.2 Rótulos externos ao redor da pizza ---
            for i, p in enumerate(wedges):
                valor = dados_grafico['Quantidade'].iloc[i]
                pct = (valor / total) * 100

                if pct <= limite_percentual:
                    ang = (p.theta2 - p.theta1) / 2.0 + p.theta1

                    x = np.cos(np.deg2rad(ang))
                    y = np.sin(np.deg2rad(ang))

                    # ponto onde a linha sai da pizza
                    x_seta = 1.00 * x
                    y_seta = 1.00 * y

                    # texto um pouco mais para fora, na mesma direção
                    x_text = 1.25 * x
                    y_text = 1.25 * y

                    ha = "left" if x >= 0 else "right"

                    ax.annotate(
                        f"{pct:.1f}%",
                        xy=(x_seta, y_seta),
                        xytext=(x_text, y_text),
                        horizontalalignment=ha,
                        fontsize=9,
                        va="center",
                        arrowprops=dict(
                            arrowstyle="-",
                            color="black",
                            lw=0.6
                        )
                    )


            # --- LEGENDA LATERAL ORGANIZADA ---
            numero_de_respostas = len(dados_grafico)
            max_por_coluna = 40
            colunas_dinamicas = (numero_de_respostas - 1) // max_por_coluna + 1

            ax.legend(
                wedges,
                dados_grafico.index,
                loc="center left",
                bbox_to_anchor=(1.15, 0.5),
                frameon=False,
                fontsize=9,
                ncol=colunas_dinamicas
            )

            # --- TÍTULO DO GRÁFICO ---
            titulo_formatado = textwrap.fill(str(nome_coluna_alvo), width=limite_caracteres_titulo)
            ax.set_title(titulo_formatado, fontsize=12, pad=35)

            # Salva imagem temporária
            caminho_img = f'temp_grafico_{len(nomes_abas_usados)}.png'
            plt.savefig(caminho_img, dpi=120, bbox_inches='tight', pad_inches=0.2)
            plt.close()
            imagens_temporarias.append(caminho_img)

            # 4. Inserir a imagem e ajustar a formatação na planilha
            worksheet = writer.sheets[novo_nome_aba]
            img = Image(caminho_img)
            
            # Pula a largura da tabela de resumo para inserir a imagem
            coluna_imagem_idx = coluna_inicio_resumo + 4 
            celula_imagem = f"{get_column_letter(coluna_imagem_idx)}2"
            worksheet.add_image(img, celula_imagem)

            # Ajustar a largura das colunas do resumo
            worksheet.column_dimensions[get_column_letter(coluna_inicio_resumo + 1)].width = 45 # Coluna 'Respostas'
            worksheet.column_dimensions[get_column_letter(coluna_inicio_resumo + 2)].width = 12 # Coluna 'Quantidade'
            worksheet.column_dimensions[get_column_letter(coluna_inicio_resumo + 3)].width = 10 # Coluna '%'

    # 5. Limpeza Final (apagar imagens temporárias do PC)
    for img_path in imagens_temporarias:
        if os.path.exists(img_path):
            os.remove(img_path)

    print(f"\nSucesso! O arquivo '{caminho_arquivo_saida}' foi gerado com todas as abas, tabelas e gráficos configurados.")

except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")