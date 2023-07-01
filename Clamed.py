import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askdirectory
from Messagebox import messagebox_done
import os


def converter_arquivo_excel_clamed(caminho_arquivo):
    colunas = [(6, 14), (14, 23), (23, 28), (28, 33), (33, 38), (38, 43), (43, 52), (52, 61), (61, 64), (64, 84),
               (90, 99)]
    nomes_colunas = ['Cod prod', 'Prç unit', '% desc', '% rep ICMS', '% ICMS', '% red ICMS',
                     'Prç BC ICMS ST', 'ICMS ST', 'Pz Pag', 'EAN', 'ICMS ST FP']

    dados = []

    with open(caminho_arquivo, 'r') as arquivo:
        linhas = arquivo.readlines()
        for linha in linhas:
            registro = [linha[inicio:fim].strip() for inicio, fim in colunas]
            dados.append(registro)

    df = pd.DataFrame(dados, columns=nomes_colunas)

    # Exibindo diálogo para seleção do diretório de salvamento
    root = Tk()
    root.withdraw()
    caminho_salvar = askdirectory(title='Selecionar diretório de salvamento')

    if caminho_salvar:
        nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0]
        caminho_salvar = os.path.join(caminho_salvar, nome_arquivo + '.xlsx')
        df.to_excel(caminho_salvar, index=False)
        print("Arquivo convertido com sucesso e salvo como Excel!")
        messagebox_done()
    else:
        print("Operação cancelada pelo usuário.")


