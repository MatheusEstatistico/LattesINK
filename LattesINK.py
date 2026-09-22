"""
Lattes Ink: verificador de currículos Lattes (interface gráfica).

A coleta em si (navegador, extração, planilha) fica em coletor.py,
que precisa estar na mesma pasta deste arquivo.
"""

import json
import os
import queue
import sys
import threading
import tkinter as tk
from collections import Counter
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import coletor


# Cores e fonte
COR_FUNDO = "#faffff"
COR_CAIXA = "white"
COR_PRIMARIA = "#00aca0"
COR_INICIAR = "#4CAF50"
COR_CANCELAR = "#d65548"
COR_EXPORTAR = "#6ab8f7"
COR_TEXTO = "#333333"
COR_APAGADO = "gray"
COR_SUCESSO = "#2e7d32"
COR_AVISO = "#c0392b"
FONTE = "Arial"

ARQUIVO_CONFIG = os.path.join(os.path.expanduser("~"), ".lattesink.json")
INTERVALO_FILA_MS = 100  # de quanto em quanto tempo a janela lê as mensagens da captura


def caminho_recurso(nome):
    """Caminho de um arquivo que fica ao lado do programa (funciona também com PyInstaller)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nome)


def carregar_config():
    try:
        with open(ARQUIVO_CONFIG, encoding="utf-8") as f:
            dados = json.load(f)
        return dados if isinstance(dados, dict) else {}
    except (OSError, ValueError):
        return {}


def salvar_config(config):
    try:
        with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except OSError:
        pass  # não lembrar a pasta não é motivo para interromper nada


def identificar_app_no_windows():
    """
    Rodando pelo python.exe, o Windows agrupa a janela como "Python" e mostra o
    ícone do Python na barra de tarefas. Com um ID próprio, ele usa o ícone da janela.
    Precisa ser chamado antes de criar a janela.
    """
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("LattesInk.Verificador")
        except Exception:
            pass


def encurtar(texto, limite=50):
    """Corta textos longos, só colocando reticências quando realmente cortou."""
    return texto if len(texto) <= limite else texto[:limite - 1].rstrip() + "…"


class LattesInk:
    def __init__(self):
        identificar_app_no_windows()
        self.janela = tk.Tk()
        self.janela.title("Lattes ink")
        self._definir_icone()
        self.janela.geometry("800x660")
        self.janela.resizable(width=False, height=False)
        self.janela.configure(bg=COR_FUNDO)
        self.janela.protocol("WM_DELETE_WINDOW", self._ao_fechar)

        # Estado interno
        self.config = carregar_config()
        self.pasta_destino = None        # None = nenhuma pasta escolhida
        self.ids = []                    # IDs válidos, sem repetição
        self.registros = []              # coletor.Registro de cada ID já processado
        self.pasta_execucao = None       # subpasta desta captura (captura_AAAA-MM-DD_HHMM)
        self.arquivo_log = None          # log.txt dentro da subpasta
        self.pode_continuar = False      # True depois de cancelar com IDs pendentes
        self.captura_ativa = False
        self.cancelado = threading.Event()
        self.coletor = None
        self.thread = None
        self.fila = queue.Queue()        # mensagens da thread de captura para a janela

        self._montar_interface()

        pasta_salva = self.config.get("pasta_destino")
        if pasta_salva and os.path.isdir(pasta_salva):
            self._definir_pasta(pasta_salva, lembrar=False)
            self._log(f"📁  Usando a última pasta escolhida: {pasta_salva}")
        else:
            self._log("Pronto. Selecione a pasta de destino e o arquivo com os IDs.")

        self._atualizar_botoes()
        self._id_fila = self.janela.after(INTERVALO_FILA_MS, self._verificar_fila)

    # ------------------------------------------------------------------
    # Montagem da janela
    # ------------------------------------------------------------------

    def _definir_icone(self):
        try:
            # "default=" aplica o ícone também às outras janelas (ex.: seleção de coluna)
            self.janela.iconbitmap(default=caminho_recurso("lattes.ico"))
        except tk.TclError:
            pass  # sem o arquivo (ou fora do Windows) o programa abre sem ícone

    def _criar_botao(self, pai, texto, comando, cor, tamanho=10, **opcoes):
        return tk.Button(
            pai, text=texto, command=comando,
            bg=cor, fg="white", activebackground=cor, activeforeground="white",
            font=(FONTE, tamanho, "bold"), relief="flat", cursor="hand2",
            padx=opcoes.pop("padx", 10), pady=opcoes.pop("pady", 8),
            **opcoes,
        )

    @staticmethod
    def _ativar(botao, ativo):
        botao.configure(state="normal" if ativo else "disabled",
                        cursor="hand2" if ativo else "arrow")

    def _titulo_secao(self, pai, texto):
        tk.Label(pai, text=texto, font=(FONTE, 10, "bold"), bg=COR_FUNDO,
                 anchor="w").pack(fill="x")

    def _caixa_arquivo(self, pai, icone, texto_inicial, texto_botao, comando):
        """Caixa branca com ícone, texto e botão à direita. Retorna (label, botão)."""
        caixa = tk.Frame(pai, bg=COR_CAIXA, relief="solid", borderwidth=1)
        caixa.pack(fill="x", pady=(5, 12))

        tk.Label(caixa, text=icone, font=(FONTE, 18), bg=COR_CAIXA).pack(side="left", padx=10, pady=8)
        botao = self._criar_botao(caixa, texto_botao, comando, COR_PRIMARIA,
                                  tamanho=9, width=9, padx=8, pady=4)
        botao.pack(side="right", padx=10, pady=8)
        label = tk.Label(caixa, text=texto_inicial, wraplength=560, justify="left",
                         anchor="w", bg=COR_CAIXA, fg=COR_APAGADO, font=(FONTE, 9, "bold"))
        label.pack(side="left", padx=5, pady=8, fill="x", expand=True)
        return label, botao

    def _montar_interface(self):
        principal = tk.Frame(self.janela, bg=COR_FUNDO)
        principal.pack(expand=True, fill="both", padx=20, pady=16)

        tk.Label(principal, text="Verificador de Currículos Lattes",
                 font=(FONTE, 14, "bold"), bg=COR_FUNDO).pack(pady=(0, 12))

        # 1. Pasta de destino
        self._titulo_secao(principal, "1. Selecione a pasta de destino")
        self.label_pasta, self.btn_pasta = self._caixa_arquivo(
            principal, "📁", "Nenhuma pasta selecionada", "Pasta", self.escolher_pasta)

        # 2. Arquivo de IDs
        self._titulo_secao(principal, "2. Envie o arquivo com os Lattes IDs (.txt, .xlsx ou .csv)")
        self.label_ids, self.btn_ids = self._caixa_arquivo(
            principal, "📄", "Nenhum arquivo selecionado", "Lattes ID's", self.escolher_ids)

        # 3. Progresso
        self._titulo_secao(principal, "3. Progresso da captura")
        frame_progresso = tk.Frame(principal, bg=COR_FUNDO)
        frame_progresso.pack(fill="x", pady=(5, 8))

        self.progresso_var = tk.DoubleVar(value=0)
        ttk.Progressbar(frame_progresso, variable=self.progresso_var,
                        maximum=100).pack(side="left", fill="x", expand=True)
        self.label_progresso = tk.Label(frame_progresso, text="0 / 0", font=(FONTE, 9, "bold"),
                                        bg=COR_FUNDO, width=9)
        self.label_progresso.pack(side="left", padx=(8, 0))

        # Log com histórico (somente leitura)
        self.caixa_log = ScrolledText(principal, height=10, font=(FONTE, 9), wrap="word",
                                      bg=COR_CAIXA, fg=COR_TEXTO, relief="solid",
                                      borderwidth=1, state="disabled")
        self.caixa_log.pack(fill="both", expand=True, pady=(0, 12))

        # Botões de ação
        frame_botoes = tk.Frame(principal, bg=COR_FUNDO)
        frame_botoes.pack()

        self.btn_iniciar = self._criar_botao(frame_botoes, "Iniciar Captura",
                                             self.iniciar_captura, COR_INICIAR)
        self.btn_iniciar.pack(side="left", padx=(0, 8))
        self.btn_cancelar = self._criar_botao(frame_botoes, "Cancelar",
                                              self.cancelar_captura, COR_CANCELAR)
        self.btn_cancelar.pack(side="left", padx=(0, 8))
        self.btn_exportar = self._criar_botao(frame_botoes, "Exportar Excel",
                                              self.exportar_para_excel, COR_EXPORTAR)
        self.btn_exportar.pack(side="left")

    def _centralizar(self, janela, largura, altura):
        self.janela.update_idletasks()
        x = self.janela.winfo_rootx() + (self.janela.winfo_width() - largura) // 2
        y = self.janela.winfo_rooty() + (self.janela.winfo_height() - altura) // 2
        janela.geometry(f"{largura}x{altura}+{max(x, 0)}+{max(y, 0)}")

    # ------------------------------------------------------------------
    # Estado da interface
    # ------------------------------------------------------------------

    def _pendentes(self):
        feitos = {r.id_lattes for r in self.registros}
        return [i for i in self.ids if i not in feitos]

    def _atualizar_botoes(self):
        """Único lugar que decide quais botões ficam ativos."""
        capturando = self.captura_ativa
        cancelando = capturando and self.cancelado.is_set()

        # Durante a captura não dá para trocar a pasta nem os IDs
        self._ativar(self.btn_pasta, not capturando)
        self._ativar(self.btn_ids, not capturando)
        self._ativar(self.btn_exportar, not capturando and bool(self.registros))
        self._ativar(self.btn_cancelar, capturando and not cancelando)
        self.btn_cancelar.configure(text="Cancelando..." if cancelando else "Cancelar")

        if capturando:
            self.btn_iniciar.configure(text="⏳  Capturando...")
            self._ativar(self.btn_iniciar, False)
            return

        pendentes = self._pendentes()
        if self.pode_continuar and pendentes:
            self.btn_iniciar.configure(text=f"▶  Continuar ({len(pendentes)} restantes)")
        else:
            self.btn_iniciar.configure(text="Iniciar Captura")
        self._ativar(self.btn_iniciar, bool(self.pasta_destino and self.ids))

    def _atualizar_progresso(self):
        total, feitos = len(self.ids), len(self.registros)
        self.progresso_var.set(feitos / total * 100 if total else 0)
        self.label_progresso.configure(text=f"{feitos} / {total}")

    def _encerrar_execucao(self):
        """Esquece a captura anterior: a próxima começa do zero, em uma subpasta nova."""
        self.pode_continuar = False
        self.pasta_execucao = None
        self.arquivo_log = None

    def _log(self, mensagem):
        """Acrescenta uma linha ao log da janela e ao log.txt da captura."""
        agora = datetime.now()
        self.caixa_log.configure(state="normal")
        self.caixa_log.insert(tk.END, f"[{agora:%H:%M:%S}]  {mensagem}\n")
        self.caixa_log.see(tk.END)
        self.caixa_log.configure(state="disabled")

        if self.arquivo_log:
            try:
                with open(self.arquivo_log, "a", encoding="utf-8") as f:
                    f.write(f"[{agora:%d/%m/%Y %H:%M:%S}]  {mensagem}\n")
            except OSError:
                pass

    # ------------------------------------------------------------------
    # 1. Pasta de destino
    # ------------------------------------------------------------------

    def escolher_pasta(self):
        pasta = filedialog.askdirectory(
            title="Escolha a pasta onde os arquivos serão salvos",
            initialdir=self.pasta_destino or os.path.expanduser("~"),
        )
        if pasta:
            self._definir_pasta(pasta)

    def _definir_pasta(self, pasta, lembrar=True):
        if pasta != self.pasta_destino:
            self._encerrar_execucao()
        self.pasta_destino = pasta
        self.label_pasta.configure(text=pasta, fg="black")
        if lembrar:
            self.config["pasta_destino"] = pasta
            salvar_config(self.config)
        self._atualizar_botoes()

    # ------------------------------------------------------------------
    # 2. Arquivo de IDs
    # ------------------------------------------------------------------

    def escolher_ids(self):
        if not self.pasta_destino or not os.path.isdir(self.pasta_destino):
            messagebox.showwarning("Atenção", "Selecione a pasta de destino antes de carregar os IDs.")
            return

        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo com os Lattes IDs",
            initialdir=self.config.get("pasta_ids") or os.path.expanduser("~"),
            filetypes=[
                ("Todos os formatos suportados", "*.txt *.xlsx *.csv"),
                ("Arquivo de texto", "*.txt"),
                ("Planilha Excel", "*.xlsx"),
                ("Arquivo CSV", "*.csv"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if not arquivo:
            return

        self.config["pasta_ids"] = os.path.dirname(arquivo)
        salvar_config(self.config)

        extensao = os.path.splitext(arquivo)[1].lower()
        if extensao == ".txt":
            try:
                valores = coletor.ler_ids_txt(arquivo)
            except Exception as erro:
                messagebox.showerror("Erro ao ler arquivo", f"Não foi possível ler o arquivo:\n{erro}")
                return
            self._aplicar_ids(valores, arquivo)
        elif extensao in (".xlsx", ".csv"):
            self._abrir_seletor_coluna(arquivo)
        else:
            messagebox.showerror("Formato não suportado", "Use arquivos .txt, .xlsx ou .csv.")

    def _aplicar_ids(self, valores, arquivo, parent=None):
        """Valida os valores lidos e, se o usuário concordar, passa a usá-los. Retorna True se aplicou."""
        parent = parent or self.janela
        ids, invalidos, duplicados = coletor.normalizar_ids(valores)

        if not ids:
            messagebox.showwarning(
                "Nenhum ID válido",
                "Não encontrei nenhum ID Lattes (16 dígitos) nos dados selecionados.",
                parent=parent)
            return False

        if invalidos or duplicados:
            partes = [f"{len(ids)} ID(s) válido(s) encontrado(s)."]
            if duplicados:
                partes.append(f"{duplicados} ID(s) repetido(s) foram removidos.")
            if invalidos:
                amostra = "\n".join(f"  • {encurtar(v, 60)}" for v in invalidos[:10])
                if len(invalidos) > 10:
                    amostra += f"\n  … e mais {len(invalidos) - 10}"
                partes.append(f"{len(invalidos)} valor(es) não são IDs Lattes válidos "
                              f"e serão ignorados:\n{amostra}")
            if not messagebox.askokcancel(
                    "Conferir IDs", "\n\n".join(partes) + "\n\nContinuar com os IDs válidos?",
                    parent=parent):
                return False

        self.ids = ids
        self.registros = []
        self._encerrar_execucao()

        texto = f"✅  {len(ids)} ID(s) carregado(s)  —  {os.path.basename(arquivo)}"
        ignorados = len(invalidos) + duplicados
        if ignorados:
            texto += f"  ({ignorados} ignorado(s))"
        self.label_ids.configure(text=texto, fg=COR_SUCESSO)
        self._log(f"📄  {len(ids)} ID(s) carregado(s) de {os.path.basename(arquivo)}")
        self._atualizar_progresso()
        self._atualizar_botoes()
        return True

    def _abrir_seletor_coluna(self, arquivo):
        """Janela modal para escolher qual coluna do .xlsx/.csv contém os IDs."""
        try:
            colunas = coletor.ler_colunas(arquivo)
        except Exception as erro:
            messagebox.showerror("Erro ao ler arquivo", f"Não foi possível ler o arquivo:\n{erro}")
            return
        if not colunas:
            messagebox.showwarning("Arquivo vazio", "O arquivo não contém dados.")
            return

        validos = [len(coletor.normalizar_ids(c.valores)[0]) for c in colunas]

        modal = tk.Toplevel(self.janela)
        modal.title("Selecionar coluna dos IDs")
        modal.resizable(False, False)
        modal.configure(bg=COR_FUNDO)
        modal.transient(self.janela)
        self._centralizar(modal, 460, 420)
        modal.grab_set()  # bloqueia a janela principal enquanto a modal estiver aberta

        tk.Label(modal, text="Selecione a coluna que contém os Lattes IDs:",
                 font=(FONTE, 10, "bold"), bg=COR_FUNDO).pack(pady=(18, 6), padx=16, anchor="w")

        frame_lista = tk.Frame(modal, bg=COR_FUNDO)
        frame_lista.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        barra = tk.Scrollbar(frame_lista)
        barra.pack(side="right", fill="y")
        lista = tk.Listbox(frame_lista, yscrollcommand=barra.set, font=(FONTE, 10),
                           selectmode="single", activestyle="dotbox", height=7)
        for coluna, qtd in zip(colunas, validos):
            lista.insert(tk.END, f"{coluna.nome}   ({qtd} ID(s) válido(s))")
        lista.pack(side="left", fill="both", expand=True)
        barra.config(command=lista.yview)

        # Já deixa selecionada a coluna com mais IDs válidos
        melhor = max(range(len(colunas)), key=lambda i: validos[i])
        lista.selection_set(melhor)
        lista.activate(melhor)
        lista.see(melhor)

        lbl_previa = tk.Label(modal, font=(FONTE, 8), fg="#555", bg=COR_FUNDO,
                              anchor="w", justify="left", wraplength=420)
        lbl_previa.pack(padx=16, fill="x")
        lbl_alerta = tk.Label(modal, font=(FONTE, 8, "bold"), fg=COR_AVISO, bg=COR_FUNDO,
                              anchor="w", justify="left", wraplength=420)
        lbl_alerta.pack(padx=16, fill="x")
        tk.Label(modal, font=(FONTE, 8), fg="#555", bg=COR_FUNDO, anchor="w", justify="left",
                 wraplength=420,
                 text="Dica: no Excel, formate a coluna dos IDs como Texto antes de digitar ou "
                      "colar. Como Número, o Excel guarda só 15 dígitos e troca o último "
                      "dígito do ID por zero.").pack(padx=16, pady=(6, 0), fill="x")

        def coluna_selecionada():
            selecao = lista.curselection()
            return colunas[selecao[0]] if selecao else None

        def atualizar_previa(*_):
            coluna = coluna_selecionada()
            if coluna is None:
                return
            amostra = [v for v in coluna.valores if v][:5]
            lbl_previa.configure(text="Prévia: " + (",  ".join(amostra) if amostra else "(sem dados)"))
            lbl_alerta.configure(
                text="⚠️  Esta coluna tem IDs salvos como número no Excel: "
                     "o último dígito pode estar errado." if coluna.tem_numeros_grandes else "")

        def confirmar(*_):
            coluna = coluna_selecionada()
            if coluna is None:
                messagebox.showwarning("Atenção", "Selecione uma coluna.", parent=modal)
                return
            if coluna.tem_numeros_grandes and not messagebox.askyesno(
                    "IDs possivelmente errados",
                    "Esta coluna tem IDs salvos como número. O Excel guarda só 15 dígitos, "
                    "então o último dígito pode ter virado zero e o programa abriria o "
                    "currículo errado.\n\nO ideal é formatar a coluna como Texto e colar os "
                    "IDs de novo.\n\nUsar esta coluna mesmo assim?",
                    icon="warning", parent=modal):
                return
            if self._aplicar_ids(coluna.valores, arquivo, parent=modal):
                modal.destroy()

        lista.bind("<<ListboxSelect>>", atualizar_previa)
        lista.bind("<Double-Button-1>", confirmar)
        atualizar_previa()

        frame_botoes = tk.Frame(modal, bg=COR_FUNDO)
        frame_botoes.pack(pady=(10, 14))
        self._criar_botao(frame_botoes, "Confirmar", confirmar, COR_INICIAR,
                          padx=12, pady=6).pack(side="left", padx=(0, 10))
        self._criar_botao(frame_botoes, "Cancelar", modal.destroy, COR_CANCELAR,
                          padx=12, pady=6).pack(side="left")

    # ------------------------------------------------------------------
    # 3. Captura
    # ------------------------------------------------------------------

    def iniciar_captura(self):
        """Inicia (ou retoma) a captura em uma thread separada para não travar a janela."""
        if self.captura_ativa or not self.ids or not self.pasta_destino:
            return
        if not os.path.isdir(self.pasta_destino):
            messagebox.showerror("Pasta não encontrada",
                                 f"A pasta de destino não existe mais:\n{self.pasta_destino}")
            return

        retomando = self.pode_continuar and self.pasta_execucao and os.path.isdir(self.pasta_execucao)
        if retomando:
            pendentes = self._pendentes()
            self._log(f"▶  Retomando a captura: faltam {len(pendentes)} ID(s).")
        else:
            try:
                self.pasta_execucao = coletor.criar_pasta_execucao(self.pasta_destino)
            except OSError as erro:
                messagebox.showerror("Erro", f"Não foi possível criar a pasta da captura:\n{erro}")
                return
            self.registros = []
            pendentes = list(self.ids)
            self.arquivo_log = os.path.join(self.pasta_execucao, coletor.NOME_LOG)
            self._log(f"📁  Nova captura de {len(pendentes)} ID(s) em: {self.pasta_execucao}")

        self.pode_continuar = False
        self.cancelado.clear()
        self.captura_ativa = True
        self.coletor = coletor.ColetorLattes(self.cancelado,
                                             log=lambda msg: self.fila.put(("log", msg)))
        self._atualizar_progresso()
        self._atualizar_botoes()

        # Tudo o que a thread precisa vai por argumento: ela não lê nada do Tkinter
        self.thread = threading.Thread(
            target=self._executar_capturas,
            args=(pendentes, self.pasta_execucao, self.coletor, len(self.registros), len(self.ids)),
            daemon=True,
        )
        self.thread.start()

    def _executar_capturas(self, pendentes, pasta_execucao, coletor_, ja_feitos, total):
        """
        Roda em segundo plano. Não toca em nenhum widget: tudo o que a janela
        precisa saber vai pela fila (self.fila), lida na thread principal.
        """
        caminho_csv = os.path.join(pasta_execucao, coletor.NOME_CSV)
        motivo, detalhe = "concluido", ""
        try:
            for posicao, lattes_id in enumerate(pendentes, start=ja_feitos + 1):
                if self.cancelado.is_set():
                    motivo = "cancelado"
                    break
                self.fila.put(("log", f"🌐  [{posicao}/{total}]  Abrindo o currículo {lattes_id}"))
                registro = coletor_.capturar(lattes_id, pasta_execucao)
                try:
                    coletor.acrescentar_csv(registro, caminho_csv)
                except OSError as erro:
                    self.fila.put(("log", "⚠️  Não foi possível gravar o CSV de backup "
                                          f"(ele está aberto no Excel?): {erro}"))
                self.fila.put(("registro", registro))
        except coletor.CapturaCancelada:
            motivo = "cancelado"
        except coletor.ErroNavegador as erro:
            motivo, detalhe = "erro", f"Não foi possível abrir o Chrome:\n{erro}"
        except Exception as erro:
            motivo, detalhe = "erro", f"Erro inesperado: {coletor.resumir_erro(erro)}"
        finally:
            coletor_.fechar()
            self.fila.put(("fim", (motivo, detalhe)))

    def cancelar_captura(self):
        if not self.captura_ativa or self.cancelado.is_set():
            return
        self.cancelado.set()
        self._atualizar_botoes()
        self._log("🛑  Cancelamento solicitado. Fechando o navegador...")
        # Fechar o Chrome interrompe na hora o que a thread estiver esperando.
        # É feito em outra thread para a janela não congelar enquanto ele fecha.
        if self.coletor:
            threading.Thread(target=self.coletor.fechar, daemon=True).start()

    # --- mensagens vindas da thread ------------------------------------

    def _verificar_fila(self):
        self._processar_fila()
        self._id_fila = self.janela.after(INTERVALO_FILA_MS, self._verificar_fila)

    def _processar_fila(self, incluir_fim=True):
        while True:
            try:
                tipo, dado = self.fila.get_nowait()
            except queue.Empty:
                return
            if tipo == "log":
                self._log(dado)
            elif tipo == "registro":
                self._receber_registro(dado)
            elif tipo == "fim" and incluir_fim:
                self._ao_terminar(*dado)

    def _receber_registro(self, r):
        self.registros.append(r)
        self._atualizar_progresso()
        posicao = f"[{len(self.registros)}/{len(self.ids)}]"
        if r.status == coletor.STATUS_OK:
            self._log(f"✅  {posicao}  {encurtar(r.nome)} — atualizado em "
                      f"{r.ultima_atualizacao:%d/%m/%Y}")
        elif r.status == coletor.STATUS_SEM_DATA:
            self._log(f"⚠️  {posicao}  {encurtar(r.nome)} — data de atualização não encontrada")
        elif r.status == coletor.STATUS_CAPTCHA:
            self._log(f"⏰  {posicao}  {r.id_lattes} — CAPTCHA não resolvido a tempo")
        else:
            self._log(f"❌  {posicao}  {r.id_lattes} — {r.observacao}")

    def _ao_terminar(self, motivo, detalhe):
        """Chamado na thread principal quando a captura termina, é cancelada ou falha."""
        self.captura_ativa = False
        self.coletor = None
        self.cancelado.clear()
        planilha = self._salvar_automatico()
        self.pode_continuar = motivo != "concluido" and bool(self._pendentes())
        self._atualizar_botoes()

        local = f"\n\nPlanilha e prints salvos em:\n{self.pasta_execucao}" if planilha else ""
        dica_continuar = ("\n\nClique em \"Continuar\" para retomar de onde parou."
                          if self.pode_continuar else "")
        processados = f"{len(self.registros)} de {len(self.ids)} ID(s) processados"

        if motivo == "cancelado":
            self._log(f"🛑  Captura cancelada: {processados}.")
            messagebox.showinfo("Captura cancelada",
                                f"A captura foi cancelada.\n{processados}.{local}{dica_continuar}")
        elif motivo == "erro":
            self._log(f"❌  Captura interrompida: {detalhe}")
            messagebox.showerror("Captura interrompida",
                                 f"{detalhe}\n\n{processados}.{local}{dica_continuar}")
        else:
            self._mostrar_resumo(local)

    def _mostrar_resumo(self, local):
        total = len(self.registros)
        ok = Counter(r.status for r in self.registros)[coletor.STATUS_OK]
        problemas = [r for r in self.registros if r.status != coletor.STATUS_OK]
        desatualizados = sum(1 for r in self.registros
                             if (r.dias_desde_atualizacao or 0) > coletor.LIMITE_DIAS_DESATUALIZADO)

        linhas = [f"{ok} de {total} currículo(s) capturado(s) com sucesso."]
        if desatualizados:
            linhas.append(f"{desatualizados} sem atualização há mais de "
                          f"{coletor.LIMITE_DIAS_DESATUALIZADO} dias (destacados na planilha).")

        if problemas:
            lista = "\n".join(f"  • {r.id_lattes}: {r.status}" for r in problemas[:15])
            if len(problemas) > 15:
                lista += f"\n  … e mais {len(problemas) - 15}"
            linhas.append(f"Com problema:\n{lista}")
            self._log(f"⚠️  Captura concluída: {ok} de {total} OK, {len(problemas)} com problema.")
            messagebox.showwarning("Captura concluída com problemas", "\n\n".join(linhas) + local)
        else:
            self._log(f"✅  Captura concluída: todos os {total} currículos capturados.")
            messagebox.showinfo("Captura concluída", "\n\n".join(linhas) + local)

    # ------------------------------------------------------------------
    # Planilha
    # ------------------------------------------------------------------

    def _salvar_automatico(self):
        """Salva a planilha na subpasta da captura. Retorna o caminho, ou None se não salvou."""
        if not self.registros or not self.pasta_execucao:
            return None
        caminho = os.path.join(self.pasta_execucao, coletor.NOME_PLANILHA)
        try:
            coletor.salvar_planilha(self.registros, caminho)
        except Exception as erro:
            self._log(f"⚠️  Não foi possível salvar a planilha automaticamente "
                      f"({coletor.resumir_erro(erro)}). Use \"Exportar Excel\".")
            return None
        self._log(f"💾  Planilha salva: {caminho}")
        return caminho

    def exportar_para_excel(self):
        """Salva uma cópia da planilha em outro lugar."""
        if not self.registros:
            messagebox.showwarning("Sem dados", "Nenhum dado foi coletado ainda. Execute a captura primeiro.")
            return

        caminho = filedialog.asksaveasfilename(
            title="Salvar arquivo Excel",
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
            initialdir=self.pasta_execucao or self.pasta_destino,
            initialfile=coletor.NOME_PLANILHA,
        )
        if not caminho:
            return

        try:
            coletor.salvar_planilha(self.registros, caminho)
        except Exception as erro:
            messagebox.showerror("Erro ao exportar", f"Erro ao salvar arquivo Excel:\n{erro}")
            self._log(f"❌  Erro ao exportar Excel: {erro}")
            return

        self._log(f"✅  Dados exportados para: {caminho}")
        messagebox.showinfo("Exportação concluída",
                            f"Dados exportados com sucesso para:\n{caminho}\n\n"
                            f"Total de registros: {len(self.registros)}")

    # ------------------------------------------------------------------
    # Fechar a janela
    # ------------------------------------------------------------------

    def _ao_fechar(self):
        if self.captura_ativa:
            if not messagebox.askyesno(
                    "Captura em andamento",
                    "Há uma captura em andamento.\n\nDeseja cancelar e sair? "
                    "Os dados já capturados serão salvos na planilha.",
                    icon="warning"):
                return
            self.cancelado.set()
            if self.coletor:
                self.coletor.fechar()
            if self.thread:
                self.thread.join(timeout=5)
            self._processar_fila(incluir_fim=False)
            self._salvar_automatico()

        self.janela.after_cancel(self._id_fila)
        self.janela.destroy()


def main():
    app = LattesInk()
    app.janela.mainloop()


if __name__ == "__main__":
    main()