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
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, WebDriverException
import pandas as pd
import urllib.parse

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
        self.bg_color = kwargs.get('bg', self.config.ACCENT_COLOR)
        self.fg_color = kwargs.get('fg', 'white')
        self.font = kwargs.get('font', ("Montserrat Bold", 12))
        self.padx = kwargs.get('padx', 25)
        self.pady = kwargs.get('pady', 12)
        self.state = kwargs.get('state', 'normal')
        self.create_button()

    def create_button(self):
        self.button_frame = tk.Frame(
            self.parent,
            bg=self.bg_color,
            relief="flat",
            bd=0,
            highlightthickness=0
        )
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
        self.button_frame.bind("<Button-1>", self._on_click)
        self.button_label.bind("<Button-1>", self._on_click)
        self.button_frame.bind("<Enter>", self._on_enter)
        self.button_label.bind("<Enter>", self._on_enter)
        self.button_frame.bind("<Leave>", self._on_leave)
        self.button_label.bind("<Leave>", self._on_leave)
        if self.state == 'disabled':
            self.disable()

    def _on_click(self, event):
        if self.state == 'normal' and self.command:
            self.command()

    def _on_enter(self, event):
        if self.state == 'normal':
            self.button_frame.configure(bg=self.config.ACCENT_HOVER_COLOR)
            self.button_label.configure(bg=self.config.ACCENT_HOVER_COLOR)

    def _on_leave(self, event):
        if self.state == 'normal':
            self.button_frame.configure(bg=self.bg_color)
            self.button_label.configure(bg=self.bg_color)

    def pack(self, **kwargs):
        return self.button_frame.pack(**kwargs)

    def grid(self, **kwargs):
        return self.button_frame.grid(**kwargs)

    def configure(self, **kwargs):
        if 'state' in kwargs:
            self.state = kwargs['state']
            if self.state == 'disabled':
                self.disable()
            else:
                self.enable()
        if 'text' in kwargs:
            self.button_label.configure(text=kwargs['text'])

    def disable(self):
        self.state = 'disabled'
        self.button_frame.configure(bg='#cccccc')
        self.button_label.configure(bg='#cccccc', fg='#666666')

    def enable(self):
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
            # removido: chrome_options.add_argument("--disable-javascript")
            chrome_options.add_argument("--disable-web-security")
            chrome_options.add_argument("--allow-running-insecure-content")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)

            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.execute_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )
            logger.info("Driver do Chrome inicializado com sucesso")
            return True
        except Exception as e:
            logger.error(f"Erro ao inicializar driver: {e}")
            return False

    def validate_phone_number(self, number: str) -> bool:
        clean_number = re.sub(r'[^\d]', '', str(number))
        return 10 <= len(clean_number) <= 15

    def _dismiss_whatsapp_update_popup(self):
        buttons = self.driver.find_elements(
            By.XPATH,
            "//*[@id='app']/div/span[2]/div/div/div/div/div/div/div[2]/div/button/div/div"
        )
        if buttons:
            try:
                buttons[0].click()
                logger.info("Popup de atualização fechado")
                time.sleep(1)
            except Exception:
                pass

    def send_single_message(self, number: str, message: str) -> Dict[str, Any]:
        result = {
            'success': False,
            'status': 'Erro desconhecido',
            'error': None
        }

        clean_number = re.sub(r'\D', '', str(number))
        if not self.validate_phone_number(clean_number):
            result['status'] = 'Número inválido'
            return result

        clean_message = str(message).strip()
        if not clean_message:
            result['status'] = 'Mensagem vazia'
            return result

        message_encoded = urllib.parse.quote(clean_message)
        url = f"https://web.whatsapp.com/send?phone={clean_number}&text={message_encoded}"

        try:
            self.driver.get(url)

            # espera curta para o botão de enviar
            wait = WebDriverWait(self.driver, 15)

            sent = False

            # 1) Tenta focar o COMPOSER (campo de digitação) no rodapé e mandar ENTER
            try:
                composer = wait.until(EC.element_to_be_clickable((
                    By.XPATH,
                    # só procura dentro do footer da conversa, evitando a caixa de busca
                    "//footer//div[@contenteditable='true' and @role='textbox']"
                )))
                # garante foco real no campo de escrita
                self.driver.execute_script("arguments[0].focus();", composer)
                time.sleep(0.2)
                composer.send_keys(Keys.ENTER)
                sent = True
            except TimeoutException:
                pass

            # 2) Se ainda não enviou, clica no botão de enviar do rodapé (PT/EN)
            if not sent:
                try:
                    send_btn = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//footer//button[@aria-label='Enviar' or @aria-label='Send' or @title='Send']"
                    )))
                    try:
                        send_btn.click()
                    except Exception:
                        self.driver.execute_script("arguments[0].click();", send_btn)
                    sent = True
                except TimeoutException:
                    # fallback extra: ícone de 'send'
                    try:
                        icon_btn = wait.until(EC.element_to_be_clickable((
                            By.XPATH,
                            "//footer//*[@data-icon='send']/ancestor::button[1]"
                        )))
                        try:
                            icon_btn.click()
                        except Exception:
                            self.driver.execute_script("arguments[0].click();", icon_btn)
                        sent = True
                    except TimeoutException:
                        pass

            if not sent:
                raise TimeoutException("Composer/botão de enviar não disponível")


            time.sleep(self.config.POST_SEND_DELAY)

        except TimeoutException:
            result['status'] = 'Timeout – botão enviar não apareceu'
            return result

        except WebDriverException as e:
            result['status'] = f'Erro WebDriver: {e.msg[:50]}'
            return result

        except Exception as e:
            result['status'] = str(e)
            return result

        result['success'] = True
        result['status'] = 'Mensagem Enviada'
        return result

    def close_driver(self):
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
        try:
            df = pd.read_excel(filepath, engine='openpyxl')
            required_columns = ["Número", "Mensagem"]
            missing_columns = [c for c in required_columns if c not in df.columns]
            if missing_columns:
                raise ValueError(f"Colunas obrigatórias não encontradas: {', '.join(missing_columns)}")
            if 'Status' not in df.columns:
                df['Status'] = ""
            df['Status'] = df['Status'].astype(object)
            df = df.dropna(subset=['Número', 'Mensagem'])
            logger.info(f"Arquivo carregado com sucesso: {len(df)} linhas válidas")
            return df
        except Exception:
            logger.error(f"Erro ao carregar arquivo: {filepath}")
            raise

    @staticmethod
    def save_excel(df: pd.DataFrame, filepath: str) -> bool:
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
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.center_dialog()
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
        self.dialog.update_idletasks()
        x = (self.parent.winfo_screenwidth() - self.dialog.winfo_reqwidth()) // 2
        y = (self.parent.winfo_screenheight() - self.dialog.winfo_reqheight()) // 2
        self.dialog.geometry(f"+{x}+{y}")

    def update_text(self, text: str):
        self.label.config(text=text)
        self.dialog.update()

    def cancel(self):
        self.cancelled = True
        self.dialog.destroy()

    def close(self):
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
            text=("Um script simples para enviar mensagens em massa no WhatsApp. "
                  "Por favor, leia o tutorial antes de usar"),
            font=("Montserrat", 12),
            fg="white",
            bg=self.config.PRIMARY_COLOR,
            wraplength=self.config.WINDOW_WIDTH - 60,
            justify="center"
        )
        subtitle_label.grid(row=1, column=0, pady=(5, 10), sticky="ew")

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
                self.status_label.config(
                    text=f"Arquivo carregado: {Path(filepath).name} ({len(self.df)} mensagens)"
                )
                self.send_messages_btn.enable()
                self.export_btn.enable()
                messagebox.showinfo(
                    "Sucesso",
                    f"Arquivo carregado com sucesso!\n{len(self.df)} mensagens encontradas."
                )
            except Exception as e:
                messagebox.showerror("Erro ao carregar arquivo", str(e))
                logger.error(f"Erro ao selecionar arquivo: {e}")

    def send_messages(self):
        if self.df is None:
            messagebox.showerror("Erro", "Nenhum arquivo carregado!")
            return
        self.send_messages_btn.disable()
        self.select_file_btn.disable()
        thread = threading.Thread(target=self._send_messages_thread, daemon=True)
        thread.start()

    def _send_messages_thread(self):
        try:
            self.progress_dialog = ProgressDialog(self.root, "Enviando Mensagens...")
            sender = self.whatsapp_sender

            if not sender.setup_driver():
                raise Exception("Falha ao inicializar navegador")

            # --- Abre o WhatsApp Web root e descarta o popup uma só vez ---
            driver = sender.driver
            driver.get("https://web.whatsapp.com")
            WebDriverWait(driver, self.config.WAIT_TIMEOUT).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='grid']"))
            )
            sender._dismiss_whatsapp_update_popup()

            total = len(self.df)
            success_count = 0

            for idx, row in self.df.iterrows():
                if self.progress_dialog.cancelled:
                    break

                # limpa e prepara dados
                num_raw = str(row['Número'])
                msg_raw = str(row['Mensagem'])
                clean_number = re.sub(r'\D', '', num_raw)
                clean_message = msg_raw.strip()

                self.progress_dialog.update_text(
                    f"Enviando {idx+1}/{total} para {clean_number}…"
                )

                result = sender.send_single_message(clean_number, clean_message)
                self.df.at[idx, 'Status'] = result['status']

                if result['success']:
                    success_count += 1

            sender.close_driver()
            self.progress_dialog.close()
            self._show_send_result(success_count, total - success_count, total)

        except Exception as e:
            logger.error(f"Erro durante envio: {e}")
            sender.close_driver()
            if self.progress_dialog:
                self.progress_dialog.close()
            messagebox.showerror("Erro durante envio", str(e))
        finally:
            self.root.after(0, self._reenable_buttons)

    def _show_send_result(self, success_count: int, error_count: int, total_count: int):
        message = (
            f"Envio concluído!\n\n"
            f"Total de mensagens: {total_count}\n"
            f"Enviadas com sucesso: {success_count}\n"
            f"Erros: {error_count}"
        )
        if error_count:
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
                messagebox.showerror("Erro ao salvar", "Erro ao salvar arquivo!")

def main():
    try:
        root = tk.Tk()
        app = CASDbotGUI(root)
        root.mainloop()
    except Exception as e:
        logger.error(f"Erro na aplicação principal: {e}")
        messagebox.showerror("Erro Fatal", str(e))

if __name__ == "__main__":
    main()
