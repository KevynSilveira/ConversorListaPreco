from tkinter import filedialog
import tkinter as tk
import customtkinter
from Clamed import converter_arquivo_excel_clamed
from Sesi import converter_arquivo_excel_sesi


def criar_interface():
    # Criar a janela principal
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("dark-blue")

    janela = customtkinter.CTk()
    janela.geometry("230x250")
    janela.title("Conversor")
    janela.resizable(False, False)
    janela.iconbitmap('Neosul.ico')

    label_selecione_layout = customtkinter.CTkLabel(master=janela, text="Selecione o layout")
    label_selecione_layout.place(x=60, y=10)

    # Função para selecionar o arquivo
    def selecionar_arquivo():
        # Abrir o diálogo de seleção de arquivo
        file_path = filedialog.askopenfilename()

        # Atualizar a entrada de texto com o caminho do arquivo selecionado
        if file_path:
            caminho_arquivo.delete(0, tk.END)
            caminho_arquivo.insert(tk.END, file_path)

    # Função para converter o arquivo
    def converter():
        # Obter o caminho do arquivo da entrada de texto
        file_path = caminho_arquivo.get()
        print(file_path)

        # Verificar se foi selecionado um arquivo
        if not file_path:
            print("Nenhum arquivo selecionado.")
            return

        # Obter o tipo de conversão selecionado
        tipo_conversao = combobox.get()

        print(tipo_conversao)

        # Chamar a função de conversão correspondente ao tipo selecionado
        if tipo_conversao == "Clamed":
            converter_arquivo_excel_clamed(file_path)
        elif tipo_conversao == "Sesi":
            converter_arquivo_excel_sesi(file_path)

    opcoes_layout = ["Clamed", "Sesi"]

    # Criar os widgets da interface

    label_selecione_layout = customtkinter.CTkLabel(master=janela, text="Selecione o layout")
    label_selecione_layout.place(x=60, y=10)

    combobox = customtkinter.CTkComboBox(master=janela, values=opcoes_layout, width=200)
    combobox.place(x=15, y=40)

    caminho_arquivo = customtkinter.CTkEntry(master=janela, width=200)
    caminho_arquivo.place(x=15, y=130)

    botao_selecionar_arquivo = customtkinter.CTkButton(master=janela, text="Selecionar arquivo", command=selecionar_arquivo)
    botao_selecionar_arquivo.place(x=50, y=170)

    botao_converter = customtkinter.CTkButton(master=janela, text="Converter", command=converter, width=200)
    botao_converter.place(x=15, y=210)

    # Iniciar o loop da interface
    janela.mainloop()


# Chamar a função para criar a interface
if __name__ == "__main__":
    criar_interface()
