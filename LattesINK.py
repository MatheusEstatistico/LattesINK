import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


class lattesink:
    def __init__(self):
        self.janela = tk.Tk()
        self.janela.title('Lattes ink')
        self.janela.iconbitmap('')
        self.janela.geometry('800x480')
        self.janela.resizable(width=False, height=False)
        self.janela.configure(bg='#faffff')

        # Estado interno
        self.ids_carregados = []
        self._captura_ativa = False

        # Frame principal
        self.frame_principal = tk.Frame(self.janela, bg='#faffff')
        self.frame_principal.pack(expand=True, fill='both', padx=20, pady=20)

        # Título
        tk.Label(
            self.frame_principal,
            text='Verificador de Currículos Lattes',
            font=('Arial', 14, 'bold'),
            bg='#faffff'
        ).pack(pady=(0, 15))

        # Seção 1: Pasta de destino
        tk.Label(
            self.frame_principal,
            text='1. Selecione a pasta de destino',
            font=('Arial', 10, 'bold'),
            bg='#faffff',
            anchor='w'
        ).pack(fill='x')

        frame_caminho = tk.Frame(self.frame_principal, bg='white', relief='solid', borderwidth=1)
        frame_caminho.pack(fill='x', pady=(5, 12))

        tk.Label(frame_caminho, text="📁", font=("Arial", 18), bg="white").pack(side="left", padx=10, pady=8)

        self.pasta_destino = tk.StringVar(value="Nenhuma pasta selecionada")
        self.label_caminho = tk.Label(
            frame_caminho,
            textvariable=self.pasta_destino,
            wraplength=1000,
            justify="left",
            bg="white",
            fg="gray",
            font=("Arial", 9, 'bold')
        )
        self.label_caminho.pack(side="left", padx=5, pady=8, fill="x", expand=True)

        tk.Button(
            frame_caminho,
            text="Pasta",
            width=9,
            command=self.escolher_pasta,
            bg="#4CAF50", fg="white",
            font=("Arial", 9, "bold"),
            padx=8, pady=4,
            cursor="hand2", relief='flat'
        ).pack(side="right", padx=10, pady=8)

        # Seção 2: Arquivo de IDs
        tk.Label(
            self.frame_principal,
            text='2. Envie o arquivo com os Lattes IDs (.txt)',
            font=('Arial', 10, 'bold'),
            bg='#faffff',
            anchor='w'
        ).pack(fill='x')

        frame_ids = tk.Frame(self.frame_principal, bg='white', relief='solid', borderwidth=1)
        frame_ids.pack(fill='x', pady=(5, 12))

        tk.Label(frame_ids, text="📄", font=("Arial", 18), bg="white").pack(side="left", padx=10, pady=8)

        self.ids_status_var = tk.StringVar(value="Nenhum arquivo selecionado")
        self.label_ids_status = tk.Label(
            frame_ids,
            textvariable=self.ids_status_var,
            justify="left",
            bg="white",
            fg="gray",
            font=("Arial", 9, 'bold')
        )
        self.label_ids_status.pack(side="left", padx=5, pady=8, fill="x", expand=True)

        tk.Button(
            frame_ids,
            text="Lattes ID's",
            width=9,
            command=self.escolher_ids,
            bg="#4CAF50", fg="white",
            font=("Arial", 9, "bold"),
            padx=8, pady=4,
            cursor="hand2", relief='flat'
        ).pack(side="right", padx=10, pady=8)

        # Seção 3: Progresso
        tk.Label(
            self.frame_principal,
            text='3. Progresso da captura',
            font=('Arial', 10, 'bold'),
            bg='#faffff',
            anchor='w'
        ).pack(fill='x')

        frame_progresso = tk.Frame(self.frame_principal, bg='#faffff')
        frame_progresso.pack(fill='x', pady=(5, 12))

        self.progresso_var = tk.DoubleVar(value=0)
        self.barra_progresso = ttk.Progressbar(
            frame_progresso,
            variable=self.progresso_var,
            maximum=100,
            length=560
        )
        self.barra_progresso.pack(side='left', fill='x', expand=True)

        self.label_progresso = tk.Label(
            frame_progresso,
            text="0 / 0",
            font=("Arial", 9, "bold"),
            bg='#faffff',
            width=8
        )
        self.label_progresso.pack(side='left', padx=(8, 0))

        # Log de status
        self.log_var = tk.StringVar(value="Aguardando início...")
        tk.Label(
            self.frame_principal,
            textvariable=self.log_var,
            font=("Arial", 9),
            bg='#faffff',
            fg='#333333',
            anchor='w'
        ).pack(fill='x', pady=(0, 10))

        # Botão iniciar
        self.btn_iniciar = tk.Button(
            self.frame_principal,
            text="Iniciar Captura",
            command=self.iniciar_captura,
            bg="#583f20", fg="white",
            font=("Arial", 10, "bold"),
            padx=10, pady=8,
            cursor="arrow", relief='flat',
            state="disabled"
        )
        self.btn_iniciar.pack()

    # Métodos de UI

    def escolher_pasta(self):
        pasta = filedialog.askdirectory(
            title='Escolha a pasta onde os arquivos serão salvos',
            initialdir=os.path.expanduser('~')
        )
        if pasta:
            self.pasta_destino.set(pasta)
            self.label_caminho.configure(fg="black")
            self._atualizar_btn_iniciar()

    def escolher_ids(self):
        pasta = self.pasta_destino.get()
        if pasta == "Nenhuma pasta selecionada" or not os.path.exists(pasta):
            messagebox.showwarning("Atenção", "Selecione a pasta de destino antes de carregar os IDs.")
            return

        arquivo = filedialog.askopenfilename(
            title='Selecione o arquivo com os Lattes IDs',
            initialdir=os.path.expanduser('~'),
            filetypes=[("Arquivo de texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        if not arquivo:
            return

        try:
            with open(arquivo, 'r', encoding='utf-8') as f:
                linhas = f.readlines()
        except Exception as e:
            messagebox.showerror("Erro ao ler arquivo", f"Não foi possível ler o arquivo:\n{e}")
            return

        ids = [linha.strip() for linha in linhas if linha.strip()]
        if not ids:
            messagebox.showwarning("Arquivo vazio", "O arquivo selecionado não contém nenhum ID.")
            return

        self.ids_carregados = ids
        self.ids_status_var.set(f"✅  {len(ids)} ID(s) carregado(s)  —  {os.path.basename(arquivo)}")
        self.label_ids_status.configure(fg="#2e7d32")
        self.label_progresso.configure(text=f"0 / {len(ids)}")
        self._atualizar_btn_iniciar()

    def _atualizar_btn_iniciar(self):
        pasta_ok = (
            self.pasta_destino.get() != "Nenhuma pasta selecionada"
            and os.path.exists(self.pasta_destino.get())
        )
        ids_ok = len(self.ids_carregados) > 0
        if pasta_ok and ids_ok:
            self.btn_iniciar.configure(state="normal", cursor="hand2")
        else:
            self.btn_iniciar.configure(state="disabled", cursor="arrow")

    def _atualizar_log(self, mensagem):
        """Atualiza o log de status na thread principal."""
        self.log_var.set(mensagem)

    def _atualizar_progresso(self, atual, total):
        """Atualiza a barra e o contador na thread principal."""
        self.progresso_var.set((atual / total) * 100)
        self.label_progresso.configure(text=f"{atual} / {total}")

    # Captura

    def iniciar_captura(self):
        """Dispara a captura em uma thread separada para não travar a janela."""
        if self._captura_ativa:
            return

        self._captura_ativa = True
        self.btn_iniciar.configure(state="disabled", cursor="arrow", text="⏳  Capturando...")
        self.progresso_var.set(0)

        thread = threading.Thread(target=self._executar_capturas, daemon=True)
        thread.start()

    def _executar_capturas(self):
        """Roda em background: itera os IDs e chama o Selenium para cada um."""
        total = len(self.ids_carregados)
        pasta = self.pasta_destino.get()
        erros = []

        for i, lattes_id in enumerate(self.ids_carregados, start=1):
            self.janela.after(0, self._atualizar_log, f"🌐  [{i}/{total}]  Abrindo navegador para ID: {lattes_id}")
            sucesso = self._screenshot_com_captcha_manual(lattes_id, pasta)

            if not sucesso:
                erros.append(lattes_id)

            self.janela.after(0, self._atualizar_progresso, i, total)

        # Finaliza na thread principal
        self.janela.after(0, self._finalizar_captura, erros, total)

    def _screenshot_com_captcha_manual(self, lattes_id, pasta_destino):
        """
        Abre o navegador visível, aguarda o usuário resolver o CAPTCHA
        manualmente e salva o screenshot na pasta de destino.
        Retorna True em caso de sucesso, False em caso de erro.
        """
        options = Options()
        options.add_argument("--window-size=1080,720")

        url = f'https://lattes.cnpq.br/{lattes_id}'
        caminho_arquivo = os.path.join(pasta_destino, f'{lattes_id}.png')
        driver = webdriver.Chrome(options=options)

        try:
            driver.get(url)
            self.janela.after(0, self._atualizar_log,
                f"⏳  [{lattes_id}]  Aguardando resolução do CAPTCHA (20s)...")

            time.sleep(20) # Tempo para o usuário resolver o CAPTCHA manualmente

            # Confirma que a página carregou após o CAPTCHA
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
            except Exception:
                self.janela.after(0, self._atualizar_log,
                    f"⚠️  [{lattes_id}]  Não foi possível confirmar o carregamento da página.")

            time.sleep(3)  # Aguarda carregamento completo

            driver.save_screenshot(caminho_arquivo)
            self.janela.after(0, self._atualizar_log,
                f"📸  [{lattes_id}]  Screenshot salvo em: {caminho_arquivo}")
            return True

        except Exception as e:
            self.janela.after(0, self._atualizar_log,
                f"❌  [{lattes_id}]  Erro: {e}")
            return False

        finally:
            driver.quit()

    def _finalizar_captura(self, erros, total):
        """Chamado na thread principal ao terminar todos os IDs."""
        self._captura_ativa = False
        self.btn_iniciar.configure(state="normal", cursor="hand2", text="▶  Iniciar Captura")

        if erros:
            self._atualizar_log(f"⚠️  Concluído com erros em {len(erros)} ID(s): {', '.join(erros)}")
            messagebox.showwarning(
                "Captura concluída com erros",
                f"{total - len(erros)} de {total} capturas concluídas com sucesso.\n\n"
                f"IDs com erro:\n" + "\n".join(erros)
            )
        else:
            self._atualizar_log(f"✅  Todos os {total} screenshots foram salvos com sucesso!")
            messagebox.showinfo(
                "Captura concluída",
                f"Todos os {total} screenshots foram salvos com sucesso em:\n{self.pasta_destino.get()}"
            )



if __name__ == "__main__":
    app = lattesink()
    app.janela.mainloop()
