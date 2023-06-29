import customtkinter
import ctypes

def center(frame):
    # Obtém a largura e altura da tela principal
    width_frame = frame.winfo_screenwidth()
    height_frame = frame.winfo_screenheight()

    # Calcula a posição x e y para centralizar o frame
    pos_x = (width_frame - frame.winfo_width()) // 2
    pos_y = (height_frame - frame.winfo_height()) // 2

    # Define a posição do frame
    frame.geometry(f"+{pos_x}+{pos_y}")

def remove_minimize(frame):
    # Obter o identificador da janela
    identifier_frame = ctypes.windll.user32.GetParent(frame.winfo_id())

    # Alterar o estilo da janela para remover o botão de minimizar
    style = ctypes.windll.user32.GetWindowLongPtrW(identifier_frame, ctypes.c_int(-16))
    style = style & ~0x00020000  # WS_MINIMIZEBOX
    ctypes.windll.user32.SetWindowLongPtrW(identifier_frame, ctypes.c_int(-16), style)

def remove_screen(frame):
    frame.destroy()

def messagebox_done():
    #Cria uma massage box de processo concluido com sucesso
    messagebox = customtkinter.CTk()
    messagebox.geometry("250x100")
    messagebox.title("Operação concluida!")

    #tira os botoes maximizar e minimizar
    messagebox.resizable(False, False)
    remove_minimize(messagebox)
    center(messagebox)

    #Adiciona ao frame a logo da Neosul
    messagebox.iconbitmap('Neosul.ico')

    # Criar uma label e atribuir o ícone
    label_selecione_layout = customtkinter.CTkLabel(master=messagebox, text="Arquivo salvo como Excel!")
    label_selecione_layout.place(x=50, y=30)

    botao_ok = customtkinter.CTkButton(master=messagebox, text="OK", width=40, height=20, command=lambda: remove_screen(messagebox))
    botao_ok.place(x=205, y=75)

    messagebox.mainloop()

def messagebox_error():
    messagebox = customtkinter.CTk()
    messagebox.geometry("250x100")
    messagebox.title("Erro!")

    #tira os botoes maximizar e minimizar
    messagebox.resizable(False, False)
    remove_minimize(messagebox)
    center(messagebox)

    #Adiciona ao frame a logo da Neosul
    messagebox.iconbitmap('Neosul.ico')

    # Criar uma label e atribuir o ícone
    lbl_select_file = customtkinter.CTkLabel(master=messagebox, text="Por favor, selecione um arquivo!")
    lbl_select_file.place(x=30, y=30)

    botao_ok = customtkinter.CTkButton(master=messagebox, text="OK", width=40, height=20, command=lambda: remove_screen(messagebox))
    botao_ok.place(x=205, y=75)

    messagebox.mainloop()
