import os
import time
import logging
import threading
import re
from typing import Optional, Dict, Any
from dataclasses import dataclass
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, WebDriverException
import pandas as pd
import urllib.parse

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('casdbot.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

@dataclass
class Config:
    """Configurações da aplicação"""
    WAIT_TIMEOUT: int = 60
    SEND_DELAY: float = 1.5
    POST_SEND_DELAY: float = 3.0
    WINDOW_WIDTH: int = 700
    WINDOW_HEIGHT: int = 500
    PRIMARY_COLOR: str = '#3192b3'
    ACCENT_COLOR: str = '#f9b342'
    ACCENT_HOVER_COLOR: str = '#e6a23c'
    ERROR_COLOR: str = '#ffe6e6'

class ModernButton:
    """Classe para criar botões modernos com bordas arredondadas e hover"""
    
    def __init__(self, parent, text, command, **kwargs):
        self.parent = parent
        self.text = text
        self.command = command
        self.config = Config()
        
        # Configurações padrão
        self.bg_color = kwargs.get('bg', self.config.ACCENT_COLOR)
        self.fg_color = kwargs.get('fg', 'white')
        self.font = kwargs.get('font', ("Montserrat Bold", 12))
        self.padx = kwargs.get('padx', 25)
        self.pady = kwargs.get('pady', 12)
        self.state = kwargs.get('state', 'normal')
        
        self.create_button()
    
    def create_button(self):
        """Cria o botão moderno"""
        # Frame principal do botão
        self.button_frame = tk.Frame(
            self.parent,
            bg=self.bg_color,
            relief="flat",
            bd=0,
            highlightthickness=0
        )
        
        # Label do botão
        self.button_label = tk.Label(
            self.button_frame,
            text=self.text,
            font=self.font,
            fg=self.fg_color,
            bg=self.bg_color,
            padx=self.padx,
            pady=self.pady,
            cursor="hand2"
        )
        self.button_label.pack()
        
        # Bindings para hover e clique
        self.button_frame.bind("<Button-1>", self._on_click)
        self.button_label.bind("<Button-1>", self._on_click)
        
        self.button_frame.bind("<Enter>", self._on_enter)
        self.button_label.bind("<Enter>", self._on_enter)
        
        self.button_frame.bind("<Leave>", self._on_leave)
        self.button_label.bind("<Leave>", self._on_leave)
        
        # Configurar estado inicial
        if self.state == 'disabled':
            self.disable()
    
    def _on_click(self, event):
        """Handler para clique do botão"""
        if self.state == 'normal' and self.command:
            self.command()
    
    def _on_enter(self, event):
        """Handler para hover enter"""
        if self.state == 'normal':
            self.button_frame.configure(bg=self.config.ACCENT_HOVER_COLOR)
            self.button_label.configure(bg=self.config.ACCENT_HOVER_COLOR)
    
    def _on_leave(self, event):
        """Handler para hover leave"""
        if self.state == 'normal':
            self.button_frame.configure(bg=self.bg_color)
            self.button_label.configure(bg=self.bg_color)
    
    def pack(self, **kwargs):
        """Pack do frame do botão"""
        return self.button_frame.pack(**kwargs)
    
    def grid(self, **kwargs):
        """Grid do frame do botão"""
        return self.button_frame.grid(**kwargs)
    
    def configure(self, **kwargs):
        """Configura propriedades do botão"""
        if 'state' in kwargs:
            self.state = kwargs['state']
            if self.state == 'disabled':
                self.disable()
            else:
                self.enable()
        
        if 'text' in kwargs:
            self.button_label.configure(text=kwargs['text'])
    
    def disable(self):
        """Desabilita o botão"""
        self.state = 'disabled'
        self.button_frame.configure(bg='#cccccc')
        self.button_label.configure(bg='#cccccc', fg='#666666')
    
    def enable(self):
        """Habilita o botão"""
        self.state = 'normal'
        self.button_frame.configure(bg=self.bg_color)
        self.button_label.configure(bg=self.bg_color, fg=self.fg_color)

class WhatsAppSender:
    """Classe responsável pelo envio de mensagens via WhatsApp Web"""
    
    def __init__(self):
        self.driver: Optional[webdriver.Chrome] = None
        self.config = Config()
        
    def setup_driver(self) -> bool:
        """Configura e inicializa o driver do Chrome"""
        try:
            chrome_options = Options()
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-plugins")
            chrome_options.add_argument("--disable-images")
            chrome_options.add_argument("--disable-javascript")
            chrome_options.add_argument("--disable-web-security")
            chrome_options.add_argument("--allow-running-insecure-content")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            logger.info("Driver do Chrome inicializado com sucesso")
            return True
            
        except Exception as e:
            logger.error(f"Erro ao inicializar driver: {e}")
            return False
    
    def validate_phone_number(self, number: str) -> bool:
        """Valida formato do número de telefone"""
        # Remove caracteres não numéricos
        clean_number = re.sub(r'[^\d]', '', str(number))
        
        # Verifica se tem entre 10 e 15 dígitos (incluindo código do país)
        if len(clean_number) < 10 or len(clean_number) > 15:
            return False
            
        return True
    
    def send_single_message(self, number: str, message: str) -> Dict[str, Any]:
        """Envia uma única mensagem"""
        result = {
            'success': False,
            'status': 'Erro desconhecido',
            'error': None
        }
        
        try:
            # Valida número
            if not self.validate_phone_number(number):
                result['status'] = 'Número inválido'
                return result
            
            # Limpa e codifica a mensagem
            clean_message = str(message).strip()
            if not clean_message:
                result['status'] = 'Mensagem vazia'
                return result
                
            # Codifica a mensagem para URL
            message_test = message.replace("\n", "%0A")
            print(message_test)
            message_encoded = urllib.parse.quote(clean_message)
            print(message_encoded)
            print(number)
            # Constrói URL
            url = f"https://web.whatsapp.com/send?phone={number}&text={message_encoded}"
            
            # Navega para a página
            self.driver.get(url)
            
            # Aguarda carregamento da página
            wait = WebDriverWait(self.driver, self.config.WAIT_TIMEOUT)
            send_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='send']"))
            )
            
            # Aguarda um pouco antes de clicar
            time.sleep(self.config.SEND_DELAY)
            
            # Clica no botão enviar
            send_button.click()
            
            # Aguarda após envio
            time.sleep(self.config.POST_SEND_DELAY)
            
            result['success'] = True
            result['status'] = 'Mensagem Enviada'
            logger.info(f"Mensagem enviada com sucesso para {number}")
            
        except TimeoutException:
            result['status'] = 'Timeout - Página não carregou'
            result['error'] = 'TimeoutException'
            logger.warning(f"Timeout ao enviar mensagem para {number}")
            
        except WebDriverException as e:
            result['status'] = f'Erro do navegador: {str(e)[:50]}'
            result['error'] = 'WebDriverException'
            logger.error(f"Erro do WebDriver para {number}: {e}")
            
        except Exception as e:
            result['status'] = f'Erro inesperado: {str(e)[:50]}'
            result['error'] = type(e).__name__
            logger.error(f"Erro inesperado ao enviar para {number}: {e}")
            
        return result
    
    def close_driver(self):
        """Fecha o driver de forma segura"""
        if self.driver:
            try:
                self.driver.quit()
                logger.info("Driver fechado com sucesso")
            except Exception as e:
                logger.error(f"Erro ao fechar driver: {e}")
            finally:
                self.driver = None

class ExcelHandler:
    """Classe responsável pelo manuseio de arquivos Excel"""
    
    @staticmethod
    def load_excel(filepath: str) -> Optional[pd.DataFrame]:
        """Carrega arquivo Excel com validações"""
        try:
            df = pd.read_excel(filepath, engine='openpyxl')
            
            # Verifica colunas obrigatórias
            required_columns = ["Número", "Mensagem"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                raise ValueError(f"Colunas obrigatórias não encontradas: {', '.join(missing_columns)}")
            
            # Inicializa coluna Status se não existir
            if 'Status' not in df.columns:
                df['Status'] = ""
            
            # Converte Status para object para permitir strings
            df['Status'] = df['Status'].astype(object)
            
            # Remove linhas vazias
            df = df.dropna(subset=['Número', 'Mensagem'])
            
            logger.info(f"Arquivo carregado com sucesso: {len(df)} linhas válidas")
            return df
            
        except FileNotFoundError:
            logger.error(f"Arquivo não encontrado: {filepath}")
            raise
        except ValueError as e:
            logger.error(f"Erro de validação: {e}")
            raise
        except Exception as e:
            logger.error(f"Erro ao carregar arquivo: {e}")
            raise
    
    @staticmethod
    def save_excel(df: pd.DataFrame, filepath: str) -> bool:
        """Salva DataFrame em arquivo Excel"""
        try:
            df.to_excel(filepath, index=False, engine="openpyxl")
            logger.info(f"Arquivo salvo com sucesso: {filepath}")
            return True
        except Exception as e:
            logger.error(f"Erro ao salvar arquivo: {e}")
            return False

class ProgressDialog:
    """Dialog de progresso para operações longas"""
    
    def __init__(self, parent, title="Processando..."):
        self.parent = parent
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("400x150")
        self.dialog.configure(bg=Config.PRIMARY_COLOR)
        
        # Centraliza o dialog
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.center_dialog()
        
        # Widgets
        self.label = tk.Label(
            self.dialog, 
            text="Iniciando...", 
            font=("Arial", 12),
            bg=Config.PRIMARY_COLOR,
            fg="white"
        )
        self.label.pack(pady=20)
        
        self.progress = ttk.Progressbar(
            self.dialog, 
            mode='indeterminate',
            length=300
        )
        self.progress.pack(pady=10)
        self.progress.start()
        
        self.cancel_button = ModernButton(
            self.dialog,
            text="Cancelar",
            command=self.cancel,
            font=("Arial", 10, "bold"),
            padx=20,
            pady=8
        )
        self.cancel_button.pack(pady=10)
        
        self.cancelled = False
        
    def center_dialog(self):
        """Centraliza o dialog na tela"""
        self.dialog.update_idletasks()
        x = (self.parent.winfo_screenwidth() - self.dialog.winfo_reqwidth()) // 2
        y = (self.parent.winfo_screenheight() - self.dialog.winfo_reqheight()) // 2
        self.dialog.geometry(f"+{x}+{y}")
    
    def update_text(self, text: str):
        """Atualiza o texto do dialog"""
        self.label.config(text=text)
        self.dialog.update()
    
    def cancel(self):
        """Cancela a operação"""
        self.cancelled = True
        self.dialog.destroy()
    
    def close(self):
        """Fecha o dialog"""
        self.dialog.destroy()
    

class CASDbotGUI:
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.config = Config()
        self.whatsapp_sender = WhatsAppSender()
        self.df: Optional[pd.DataFrame] = None
        self.progress_dialog: Optional[ProgressDialog] = None
        
        self.setup_gui()
        
    def setup_gui(self):
        self.root.title("CASDbot v2025")
        self.root.geometry(f"{self.config.WINDOW_WIDTH}x{self.config.WINDOW_HEIGHT}")
        self.root.configure(bg=self.config.PRIMARY_COLOR)
        self.root.resizable(True, True)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        main_frame = tk.Frame(self.root, bg=self.config.PRIMARY_COLOR)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        main_frame.columnconfigure(0, weight=1)
        
        title_label = tk.Label(
            main_frame,
            text="CASDbot",
            font=("Montserrat Bold", 24),
            fg="white",
            bg=self.config.PRIMARY_COLOR
        )
        title_label.grid(row=0, column=0, pady=(10, 5))
        
        subtitle_label = tk.Label(
            main_frame,
            text="Um script simples para enviar mensagens em massa no WhatsApp. Por favor, leia o tutorial antes de usar",
            font=("Montserrat", 12),
            fg="white",
            bg=self.config.PRIMARY_COLOR,
            wraplength=self.config.WINDOW_WIDTH - 60, 
            justify="center"
        )
        subtitle_label.grid(row=1, column=0, pady=(5, 10), sticky="ew")
        main_frame.grid_columnconfigure(0, weight=1)
        
        button_frame = tk.Frame(main_frame, bg=self.config.PRIMARY_COLOR)
        button_frame.grid(row=2, column=0, pady=20)
        
        self.select_file_btn = ModernButton(
            button_frame,
            text="📁 Escolher Arquivo Excel",
            command=self.select_file,
            font=("Montserrat Bold", 12),
            padx=25,
            pady=12
        )
        self.select_file_btn.pack(pady=8)
        
        self.send_messages_btn = ModernButton(
            button_frame,
            text="📤 Enviar Mensagens",
            command=self.send_messages,
            font=("Montserrat Bold", 12),
            padx=25,
            pady=12,
            state="disabled"
        )
        self.send_messages_btn.pack(pady=8)
        
        self.export_btn = ModernButton(
            button_frame,
            text="💾 Exportar Status",
            command=self.export_file,
            font=("Montserrat Bold", 12),
            padx=25,
            pady=12,
            state="disabled"
        )
        self.export_btn.pack(pady=8)
        
        status_frame = tk.Frame(main_frame, bg=self.config.PRIMARY_COLOR)
        status_frame.grid(row=3, column=0, pady=20, sticky="ew")
        status_frame.columnconfigure(0, weight=1)
        
        self.status_label = tk.Label(
            status_frame,
            text="Nenhum arquivo selecionado",
            font=("Montserrat", 10),
            fg="white",
            bg=self.config.PRIMARY_COLOR
        )
        self.status_label.grid(row=0, column=0)
                
        footer_frame = tk.Frame(self.root, bg=self.config.PRIMARY_COLOR)
        footer_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        footer_frame.columnconfigure(0, weight=1)
        
        footer_label = tk.Label(
            footer_frame,
            text="Qualquer dúvida, contate o Fóton - T26 - (85) 98413-2943",
            font=("Montserrat Bold", 9),
            fg="white",
            bg=self.config.PRIMARY_COLOR
        )
        footer_label.grid(row=0, column=0)
        
    def select_file(self):
        filepath = filedialog.askopenfilename(
            title="Escolha o arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        
        if filepath:
            try:
                self.df = ExcelHandler.load_excel(filepath)
                self.status_label.config(text=f"Arquivo carregado: {Path(filepath).name} ({len(self.df)} mensagens)")
                self.send_messages_btn.enable()
                self.export_btn.enable()
                messagebox.showinfo("Sucesso", f"Arquivo carregado com sucesso!\n{len(self.df)} mensagens encontradas.")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar arquivo:\n{str(e)}")
                logger.error(f"Erro ao selecionar arquivo: {e}")
    
    def send_messages(self):
        if self.df is None:
            messagebox.showerror("Erro", "Nenhum arquivo carregado!")
            return
        
        self.send_messages_btn.disable()
        self.select_file_btn.disable()
        
        thread = threading.Thread(target=self._send_messages_thread)
        thread.daemon = True
        thread.start()
    
    def _send_messages_thread(self):
        """Thread para envio de mensagens"""
        try:
            # Cria dialog de progresso
            self.progress_dialog = ProgressDialog(self.root, "Enviando Mensagens...")
            
            # Inicializa driver
            if not self.whatsapp_sender.setup_driver():
                raise Exception("Falha ao inicializar navegador")
            
            total_messages = len(self.df)
            success_count = 0
            error_count = 0
            current_message = 0
            
            for index, row in self.df.iterrows():
                if self.progress_dialog.cancelled:
                    break
                
                current_message += 1
                
                progress_text = f"Enviando mensagem {current_message} de {total_messages}..."
                self.progress_dialog.update_text(progress_text)
                
                result = self.whatsapp_sender.send_single_message(
                    str(row['Número']), 
                    str(row['Mensagem']),                    
                )
                
                self.df.at[index, 'Status'] = result['status']
                
                if result['success']:
                    success_count += 1
                else:
                    error_count += 1
            
            self.whatsapp_sender.close_driver()
            
            if self.progress_dialog:
                self.progress_dialog.close()
            
            self._show_send_result(success_count, error_count, total_messages)
            
        except Exception as e:
            logger.error(f"Erro durante envio: {e}")
            self.whatsapp_sender.close_driver()
            
            if self.progress_dialog:
                self.progress_dialog.close()
            
            messagebox.showerror("Erro", f"Erro durante envio:\n{str(e)}")
            
        finally:
            self.root.after(0, self._reenable_buttons)
    
    def _show_send_result(self, success_count: int, error_count: int, total_count: int):
        message = f"Envio concluído!\n\n"
        message += f"Total de mensagens: {total_count}\n"
        message += f"Enviadas com sucesso: {success_count}\n"
        message += f"Erros: {error_count}"
        
        if error_count > 0:
            messagebox.showwarning("Envio Concluído", message)
        else:
            messagebox.showinfo("Sucesso", message)
    
    def _reenable_buttons(self):
        self.send_messages_btn.enable()
        self.select_file_btn.enable()
    
    def export_file(self):
        if self.df is None:
            messagebox.showerror("Erro", "Nenhum arquivo carregado!")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Salvar arquivo com status"
        )
        
        if filepath:
            if ExcelHandler.save_excel(self.df, filepath):
                messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso!")
            else:
                messagebox.showerror("Erro", "Erro ao salvar arquivo!")
    

def main():

    try:
        root = tk.Tk()
        app = CASDbotGUI(root)
        root.mainloop()
    except Exception as e:
        logger.error(f"Erro na aplicação principal: {e}")
        messagebox.showerror("Erro Fatal", f"Erro na aplicação:\n{str(e)}")

if __name__ == "__main__":
    main() 
