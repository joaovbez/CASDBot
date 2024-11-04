import os
import time
from tkinter import *
from tkinter import filedialog
from PIL import Image, ImageTk
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import urllib.parse

# Classe da interface gráfica do CASDbot + comandos 
class Application:
    # Interface gráfica e associação com os comandos
    def __init__(self, master=None):
        
        # Interface Geral        
        master.title("CASDbot v2024")
        master.geometry("600x400")  # Aumentei a altura para garantir que os botões não fiquem cortados
        master.configure(bg='#3192b3')  # Define cor de fundo
        self.master = master

        # Configura o layout da janela principal
        self.master.columnconfigure(0, weight=1)  # Permite que a janela expanda horizontalmente de forma proporcional
        self.master.rowconfigure(0, weight=1)     # Permite que a janela expanda verticalmente de forma proporcional

        # Adicionando widgets para a interface
        self.widget1 = Frame(master, bg="#3192b3")
        self.widget1.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)  # Centraliza o conteúdo
        self.widget1.grid_columnconfigure(0, weight=1)

        self.text1 = Label(self.widget1, text="CASDbot", anchor=CENTER, font=("MontSerrat Bold", 18), fg="white", bg="#3192b3")
        self.text1.grid(row=0, column=0, pady=(10, 5))  # Espaçamento ajustado

        self.text2 = Label(self.widget1, text="Leia o tutorial antes de utilizar!", anchor=CENTER, font=("MontSerrat", 14), fg="white", bg="#3192b3")
        self.text2.grid(row=1, column=0, pady=(5, 20))  # Espaçamento ajustado

        # Botão de selecionar arquivo
        self.button_file = Button(self.widget1, text="Escolher Arquivo", font=("Montserrat Bold", 10), command=self.select_file, fg="white", bg="#f9b342")
        self.button_file.grid(row=2, column=0, pady=(5, 10))  # Botão posicionado de forma relativa, com espaçamento

        # Botão de enviar mensagens
        self.button_message = Button(self.widget1, text="Enviar Mensagens", font=("Montserrat Bold", 10), command=self.send_message, fg="white", bg="#f9b342")
        self.button_message.grid(row=3, column=0, pady=(5, 10))  # Garantido que o botão esteja totalmente visível

        # Botão de status será adicionado dinamicamente após o envio de mensagens
        self.widget_export = None
        self.button_export = None

        # Rodapé (informações de contato)
        self.widget2 = Frame(master, bg="#3192b3")
        self.widget2.grid(row=5, column=0, sticky="nsew", pady=(50, 10))  # Ajustando posição
        self.widget2.grid_columnconfigure(0, weight=1)

        self.text3 = Label(self.widget2, text="Qualquer dúvida, contate o Fóton - T26 - (85) 98413-2943", anchor=CENTER, font=("MontSerrat Bold", 8), fg="white", bg="#3192b3")
        self.text3.grid(row=0, column=0)

        self.df = None

    # Função de mensagem de erro personalizada
    def custom_error_message(self, title, message):
        error_window = Toplevel(self.master)
        error_window.title(title)
        error_window.geometry("400x200")  # Aumentei o tamanho para caber mais texto
        error_window.configure(bg='#ffe6e6')  # Fundo claro para destacar o erro

        # Centralizar a janela de erro na tela
        error_window.update_idletasks()
        x = (self.master.winfo_screenwidth() - error_window.winfo_reqwidth()) // 2
        y = (self.master.winfo_screenheight() - error_window.winfo_reqheight()) // 2
        error_window.geometry(f"+{x}+{y}")  # Define a posição da janela no centro da tela

        label_title = Label(error_window, text=title, font=("Arial Bold", 16), fg="red", bg='#ffe6e6')
        label_title.pack(pady=10)

        # O 'wraplength' quebra o texto automaticamente em múltiplas linhas se necessário
        label_message = Label(error_window, text=message, font=("Arial", 12), bg='#ffe6e6', wraplength=350)
        label_message.pack(pady=5)

        button_close = Button(error_window, text="Fechar", command=error_window.destroy, font=("Arial", 10, "bold"), bg="#ff3333", fg="white")
        button_close.pack(pady=10)

        error_window.transient(self.master)
        error_window.grab_set()
        self.master.wait_window(error_window)

    # Função para adicionar botão de status na interface (aparece apenas após envio de mensagens)
    def button_status(self):
        if not self.button_export:
            self.widget_export = Frame(self.master, bg="#3192b3")
            self.widget_export.grid(row=4, column=0, pady=(10, 20))  # Posicionado logo abaixo do botão de enviar mensagens
            self.widget_export.grid_columnconfigure(0, weight=1)
            self.button_export = Button(self.widget_export, text="Ver Status das Mensagens", font=("Montserrat Bold", 10), command=self.export_file, fg="white", bg="#f9b342")
            self.button_export.grid(row=0, column=0, pady=10)

    # Função para selecionar arquivo
    def select_file(self):
        filepath = filedialog.askopenfilename(title="Escolha o arquivo padronizado", filetypes=[("Arquivos Excel", "*xlsx")])
        if filepath:
            try:
                self.df = pd.read_excel(filepath, engine='openpyxl')
                if 'Status' not in self.df.columns:
                    self.df['Status'] = ""  # Inicializa a coluna 'Status' com strings vazias

                required_columns = ["Número", "Mensagem"]
                for col in required_columns:
                    if col not in self.df.columns:
                        raise ValueError(f"Coluna '{col}' não encontrada no arquivo Excel.")

                self.df['Status'] = self.df['Status'].astype(object)  # Converte para tipo object

            except FileNotFoundError as e:
                self.custom_error_message("Erro", f"Erro ao abrir o arquivo: {e}")
            except ValueError as e:
                self.custom_error_message("Erro", str(e))
        else:
            self.custom_error_message("Erro", "Nenhum arquivo foi selecionado")

    # Função para enviar as mensagens
    def send_message(self):
        if self.df is not None:
            total_contacts = range(self.df.shape[0])
            error = 0
            self.driver = webdriver.Chrome()
            for i in total_contacts:
                number = self.df.at[i, 'Número']
                message = self.df.at[i, 'Mensagem']
                message = message.replace("\n", "%0A")
                # Codificar a mensagem corretamente para ser usada na URL
                message_encoded = urllib.parse.quote(message)

                try:
                    url = f"https://web.whatsapp.com/send?phone={number}&text={message}"
                    self.driver.get(url)

                    WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.XPATH, "//span[@data-icon='send']")))
                    time.sleep(1.5)
                    send_button = self.driver.find_element(By.XPATH, "//span[@data-icon='send']")
                    send_button.click()    
                    time.sleep(3)                            
                    self.df.at[i, 'Status'] = 'Mensagem Enviada'                    
                except Exception:
                    self.df.at[i, 'Status'] = 'Erro ao enviar'
                    error = 1

            # Fecha o navegador após o envio das mensagens
            self.driver.quit()
            # Mostra o botão de status após o envio
            self.button_status()
            if error == 0:
                self.custom_error_message("Sucesso", "Envio das mensagens concluído!")
            else:
                self.custom_error_message("Erro", "Envio concluído, mas ocorreram alguns erros. Verifique a coluna Status.")
            
        
        else:
            self.custom_error_message("Erro", "Você não anexou o arquivo corretamente")
            return
        

    # Função para exportar o arquivo atualizado
    def export_file(self):
        if self.df is not None:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Salvar como")
            if save_path:
                try:
                    self.df.to_excel(save_path, index=False, engine="openpyxl")
                    self.custom_error_message("Sucesso", "Arquivo salvo com sucesso")
                except Exception as e:
                    self.custom_error_message("Erro", f"Erro ao salvar arquivo: {e}")
        else:
            self.custom_error_message("Erro", "Nenhum status disponível para exportar")


root = Tk() 
Application(root)
root.mainloop()


