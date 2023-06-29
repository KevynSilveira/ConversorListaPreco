import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askdirectory
from Messagebox import messagebox_feito
import os

def converter_arquivo_excel_sesi(caminho_arquivo):
    # Carregar o arquivo CSV usando o separador ";"
    df = pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')

    # Renomear as colunas
    df.columns = ['EAN', 'PRODUTO', 'PRC_LIQ_IMP', '% DESCONTO', 'DATA_LISTA', 'DATA_LISTA', 'iNDEFINIDO']

    # Exibindo diálogo para seleção do diretório de salvamento
    root = Tk()
    root.withdraw()
    caminho_salvar = askdirectory(title='Selecionar diretório de salvamento')

    if caminho_salvar:
        nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0]
        caminho_salvar = os.path.join(caminho_salvar, nome_arquivo + '.xlsx')
        df.to_excel(caminho_salvar, index=False)

        print("Arquivo convertido com sucesso e salvo como Excel!")
        messagebox_feito()
    else:
        print("Operação cancelada pelo usuário.")