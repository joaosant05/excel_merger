import pandas as pd
import os

def carregar_arquivos_excel(diretorio):
    """
    Carrega todos os arquivos Excel (.xlsx) de um diretório.
    """
    arquivos_excel = [f for f in os.listdir(diretorio) if f.endswith('.xlsx')]
    return arquivos_excel

def processar_arquivos(diretorio):
    """
    Processa cada arquivo Excel do diretório e cria o novo layout especificado.
    """
    dados_combinados = []

    for arquivo in carregar_arquivos_excel(diretorio):
        
        df = pd.read_excel(os.path.join(diretorio, arquivo))

        # Nome do arquivo sem a extensão
        parent_code = os.path.splitext(arquivo)[0]

        # Monta o novo DataFrame com o layout especificado
        novo_layout = pd.DataFrame({
            'Position': df['Posição'],
            'Name_PT': df['Nome'],
            'Code': df['Código'],
            'Rev': df['Revisão'],
            'Is_Assembly': 0,
            'Update_Existing': 1,
            'Parent_Code': parent_code,
            'Parent_Rev': '',
            'Quantity': df['Quantidade'],
            'Specifications': df['Especificações'],
            'Buyable': 1,
            'Draw_File': '',
            'Image_File': ''  
        })

        # Adiciona ao conjunto de dados combinados
        dados_combinados.append(novo_layout)

    # Concatena todos os DataFrames em um único DataFrame
    df_final = pd.concat(dados_combinados, ignore_index=True)

    # Salva o DataFrame final em um novo arquivo Excel
    df_final.to_excel('saida_combinada.xlsx', index=False)

# Diretório onde estão os arquivos Excel
diretorio = r'C:\Users\joaog\Desktop\exceis'

# Processa os arquivos e gera o arquivo combinado
processar_arquivos(diretorio)
