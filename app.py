"""
Envio de Contracheques — interface para o RH.

Aba "Envio":          1) Conferir  ->  2) revisar a lista  ->  3) Enviar.
Aba "Configurações":  pastas, forma de envio (Outlook / SMTP / Simulado), mensagem e senha do PDF.

Nada é enviado na conferência. O envio só considera itens "Pronto".
"""

from __future__ import annotations

import dataclasses
import html
from collections import Counter
import logging
import os
import queue
import subprocess
import tempfile
import webbrowser
import sys
import threading
import tkinter as tk
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

# Bibliotecas (todas em Python puro) ficam na subpasta lib/ — nada é instalado no Windows
BASE = Path(__file__).resolve().parent
# Isola o programa das bibliotecas instaladas no Python da máquina (site-packages).
# Usa só a biblioteca padrão + a pasta lib/. Evita conflito com versões instaladas e com
# pacotes compilados (numpy, lxml...) que o Controle de Aplicativo Inteligente pode bloquear.
sys.path[:] = [p for p in sys.path if "site-packages" not in p.lower() and "dist-packages" not in p.lower()]
for _caminho in (str(BASE / "lib"), str(BASE)):
    if _caminho in sys.path:
        sys.path.remove(_caminho)
sys.path[:0] = [str(BASE), str(BASE / "lib")]

import cofre  # noqa: E402
import identidade  # noqa: E402
import nucleo  # noqa: E402
from nucleo import Situacao  # noqa: E402

# Paleta e fontes
AZUL = "#1F4E79"
AZUL_CLARO = "#2E75B6"
FUNDO = "#F2F4F7"
BRANCO = "#FFFFFF"
TINTA = "#1C2430"          # texto principal
CINZA_TEXTO = "#5E6876"    # texto secundário
LINHA = "#DDE2E8"          # divisórias e bordas
FONTE = ("Segoe UI", 10)
FONTE_NEGRITO = ("Segoe UI", 10, "bold")

COMP_AUTO = "Automática (lida do recibo)"
ORIGENS = {
    "individual": "PDFs individuais (um arquivo por colaborador)",
    "unico": "PDF único da contabilidade (separar automaticamente)",
}


def _misturar(cor: str, alvo: str, fator: float) -> str:
    """Clareia/escurece uma cor hex misturando com outra (0 = cor original, 1 = alvo)."""
    a = [int(cor[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(alvo[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * fator):02X}" for x, y in zip(a, b))


def aplicar_tema(cfg) -> None:
    """Cores da empresa (seção [tema] do config.ini) aplicadas à tela inteira."""
    global AZUL, AZUL_CLARO, FUNDO, DESTAQUE, AZUL_SUAVE
    AZUL = cfg.cor_primaria
    DESTAQUE = cfg.cor_destaque
    FUNDO = cfg.cor_fundo
    AZUL_CLARO = _misturar(cfg.cor_primaria, "#FFFFFF", 0.18)
    AZUL_SUAVE = _misturar(cfg.cor_primaria, "#FFFFFF", 0.88)


DESTAQUE = "#1FA2B8"
AZUL_SUAVE = "#E3E8EF"

# Fundo da linha (suave) e cor do indicador (forte) por situação
CORES_SITUACAO = {
    Situacao.PRONTO: "#E8F4EC",
    Situacao.REVISAR: "#FFF4DA",
    Situacao.ERRO: "#FBE8EA",
    Situacao.PENDENTE: "#F3F4F6",
    Situacao.JA_ENVIADO: "#EFF1F4",
    Situacao.ENVIADO: "#E6EFF8",
    Situacao.TESTE: "#EFEBF8",
    Situacao.SIMULADO: "#EFEBF8",
}
PONTOS_SITUACAO = {
    Situacao.PRONTO: "#2F855A",
    Situacao.REVISAR: "#B7791F",
    Situacao.ERRO: "#C53030",
    Situacao.PENDENTE: "#8A94A3",
    Situacao.JA_ENVIADO: "#A0AAB8",
    Situacao.ENVIADO: "#2B6CB0",
    Situacao.TESTE: "#6B46C1",
    Situacao.SIMULADO: "#6B46C1",
}
ORDEM_SITUACAO = [Situacao.ERRO, Situacao.REVISAR, Situacao.PRONTO, Situacao.PENDENTE,
                  Situacao.JA_ENVIADO, Situacao.ENVIADO, Situacao.TESTE, Situacao.SIMULADO]

MODOS = {
    "outlook": "Outlook desta máquina (recomendado, não precisa de senha)",
    "smtp": "Servidor de e-mail com login e senha (SMTP)",
    "simulado": "Simulado: não envia, só gera os arquivos para conferência",
}
SEGURANCAS = {"STARTTLS (porta 587)": "starttls", "SSL (porta 465)": "ssl"}
FORMATOS_SENHA = {
    "4 últimos dígitos do CPF": "{cpf4}",
    "4 últimos dígitos do CPF + dia e mês de nascimento": "{cpf4}{nasc_ddmm}",
}


class HandlerFila(logging.Handler):
    """Leva as mensagens do log para a caixa de texto da tela."""

    def __init__(self, fila: queue.Queue):
        super().__init__()
        self.fila = fila
        self.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))

    def emit(self, record):
        self.fila.put(("log", self.format(record)))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.base = BASE
        self.arq_config = self.base / "config.ini"
        self.arq_usuario = self.base / nucleo.ARQUIVO_CONFIG_USUARIO
        nucleo.configurar_logging(self.base)

        self.fila: queue.Queue = queue.Queue()
        nucleo.log.addHandler(HandlerFila(self.fila))

        self.itens: list[nucleo.ItemEnvio] = []
        self.cancelar = threading.Event()
        self.ocupado = False
        self.competencia_conferida = ""

        self.logo_topo = None
        try:
            self._recarregar_cfg()
            aplicar_tema(self.cfg)
            self.logo_topo = identidade.carregar_logo(self._caminho_logo(), altura_max=52)
            icone = identidade.carregar_logo(self._caminho_logo(), altura_max=64)
            if icone is not None:
                self._icone = icone
                self.iconphoto(True, icone)
        except Exception as exc:
            self.withdraw()
            messagebox.showerror("Configuração inválida", f"Erro nas configurações:\n\n{exc}")
            raise SystemExit(1)

        # --- variáveis da aba Envio
        self.var_pasta = tk.StringVar(value=self.cfg.pasta_pdfs)
        self.var_planilha = tk.StringVar(value=self.cfg.planilha)
        self.var_competencia = tk.StringVar(value=COMP_AUTO)
        self.var_origem = tk.StringVar(value=ORIGENS["individual"])
        self.var_pdf_unico = tk.StringVar()
        self.var_teste = tk.BooleanVar(value=True)  # começa em teste de propósito
        self.var_email_teste = tk.StringVar(value=self.cfg.email_teste)
        self.var_conferir_nome = tk.BooleanVar(value=self.cfg.conferir_nome)
        self.var_status = tk.StringVar(value="Selecione a pasta e a planilha e clique em Conferir.")
        self.var_resumo = tk.StringVar(value="")
        self.var_selo = tk.StringVar()

        # --- variáveis da aba Configurações
        self.c_pasta = tk.StringVar()
        self.c_planilha = tk.StringVar()
        self.c_modo = tk.StringVar()
        self.c_remetente_outlook = tk.StringVar()
        self.c_smtp_host = tk.StringVar()
        self.c_smtp_porta = tk.StringVar()
        self.c_smtp_seg = tk.StringVar()
        self.c_smtp_usuario = tk.StringVar()
        self.c_smtp_senha = tk.StringVar()
        self.c_smtp_remetente = tk.StringVar()
        self.c_empresa = tk.StringVar()
        self.c_assunto = tk.StringVar()
        self.c_email_teste = tk.StringVar()
        self.c_formato = tk.StringVar()
        self.c_pausa = tk.StringVar()
        self.c_mover = tk.BooleanVar()
        self.c_conferir = tk.BooleanVar()
        self.var_status_senha = tk.StringVar()
        self.var_status_config = tk.StringVar()

        self.title(f"Envio de Contracheques — {self.cfg.nome_empresa}")
        self.geometry("1240x800")
        self.minsize(980, 640)
        self.configure(bg=FUNDO)
        self._estilos()
        self._montar_tela()
        self._carregar_formulario()
        self._atualizar_selo()
        self.protocol("WM_DELETE_WINDOW", self._ao_fechar)
        self.after(100, self._processar_fila)

        nucleo.log.info("Programa iniciado. Modo de envio: %s | Senha do PDF: %s | Criptografia: %s",
                        self.cfg.modo_envio, nucleo.descricao_senha(self.cfg.formato_senha),
                        nucleo.PROVEDOR_CRIPTO)

        if not self.arq_usuario.exists():
            # Primeira execução: começa pelas configurações
            self.abas.select(self.aba_config)
            self.var_status_config.set("Primeiro uso: preencha as configurações e clique em Salvar.")

    # ================================================================== identidade
    def _caminho_logo(self) -> Path:
        return self.base / self.cfg.logo if hasattr(self, "cfg") else self.base / "logo.png"

    def _recarregar_cfg(self):
        """Lê a configuração e, se houver logo, usa as cores da marca na tela e no e-mail."""
        self.cfg = nucleo.carregar_config(self.arq_config, self.arq_usuario)
        if self.cfg.cores_da_logo:
            cores = identidade.cores_da_logo(self.base / self.cfg.logo)
            if cores:
                self.cfg.cor_primaria, self.cfg.cor_destaque = cores

    # ================================================================== visual
    def _estilos(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", font=FONTE, background=FUNDO, foreground=TINTA)
        s.configure("Card.TFrame", background=BRANCO)
        s.configure("Card.TLabelframe", background=BRANCO, padding=14)
        s.configure("Card.TLabelframe.Label", background=BRANCO, foreground=AZUL, font=FONTE_NEGRITO)
        s.configure("Card.TLabel", background=BRANCO, foreground=TINTA)
        s.configure("Dica.TLabel", background=BRANCO, foreground=CINZA_TEXTO, font=("Segoe UI", 9))
        s.configure("Ok.TLabel", background=BRANCO, foreground="#2E7D32", font=("Segoe UI", 9, "bold"))
        s.configure("Card.TCheckbutton", background=BRANCO)
        s.configure("Card.TRadiobutton", background=BRANCO)
        s.configure("Resumo.TLabel", background=FUNDO, foreground=AZUL, font=FONTE_NEGRITO)
        s.configure("Status.TLabel", background=FUNDO, foreground=CINZA_TEXTO)
        s.configure("TEntry", padding=5)
        s.configure("TNotebook", background=FUNDO, borderwidth=0, tabmargins=(0, 0, 0, 0))
        s.configure("TNotebook.Tab", padding=(20, 9), font=FONTE, background=FUNDO, foreground=CINZA_TEXTO,
                    borderwidth=0)
        s.map("TNotebook.Tab", background=[("selected", BRANCO), ("active", AZUL_SUAVE)],
              foreground=[("selected", AZUL)], font=[("selected", FONTE_NEGRITO)])
        s.configure("Painel.TFrame", background=BRANCO, bordercolor=LINHA, relief="solid", borderwidth=1)
        s.configure("Primario.TButton", background=AZUL, foreground=BRANCO, padding=(16, 8),
                    font=FONTE_NEGRITO, borderwidth=0)
        s.map("Primario.TButton", background=[("disabled", "#A9B7C6"), ("active", AZUL_CLARO)])
        s.configure("Sucesso.TButton", background="#2E7D32", foreground=BRANCO, padding=(16, 8),
                    font=FONTE_NEGRITO, borderwidth=0)
        s.map("Sucesso.TButton", background=[("disabled", "#A5C9A7"), ("active", "#388E3C")])
        s.configure("Secundario.TButton", padding=(12, 6), background=BRANCO, foreground=AZUL,
                    bordercolor=LINHA, lightcolor=BRANCO, darkcolor=BRANCO)
        s.map("Secundario.TButton", background=[("active", AZUL_SUAVE), ("disabled", FUNDO)],
              foreground=[("disabled", "#A0AAB8")])
        s.configure("Treeview", rowheight=28, font=FONTE, background=BRANCO, fieldbackground=BRANCO,
                    bordercolor=LINHA, foreground=TINTA)
        s.configure("Treeview.Heading", font=FONTE_NEGRITO, background=AZUL_SUAVE, foreground=AZUL)
        s.map("Treeview", background=[("selected", AZUL_CLARO)], foreground=[("selected", BRANCO)])
        s.configure("Destaque.TButton", background=DESTAQUE, foreground=BRANCO, padding=(14, 6),
                    font=FONTE_NEGRITO, borderwidth=0)
        s.map("Destaque.TButton", background=[("disabled", "#B8C4CC"), ("active", _misturar(DESTAQUE, "#000000", 0.15))])
        s.configure("Marcador.TButton", padding=(8, 3), font=("Segoe UI", 9))
        s.configure("Titulo.TLabel", background=BRANCO, foreground=AZUL, font=("Segoe UI Semibold", 11))
        s.configure("Horizontal.TProgressbar", troughcolor=AZUL_SUAVE, background=DESTAQUE, thickness=10)

    def _montar_tela(self):
        # Cabeçalho branco: qualquer versão da logo (colorida) fica legível; a cor da marca entra na faixa
        topo = tk.Frame(self, bg=BRANCO)
        topo.pack(fill="x")
        interno = tk.Frame(topo, bg=BRANCO)
        interno.pack(fill="x", padx=26, pady=14)
        if self.logo_topo is not None:
            tk.Label(interno, image=self.logo_topo, bg=BRANCO).pack(side="left")
        else:
            # Sem arquivo de logo: marca nominativa na cor da empresa
            tk.Label(interno, text=self.cfg.nome_empresa.split()[0], bg=BRANCO, fg=AZUL,
                     font=("Segoe UI Semibold", 22)).pack(side="left")
        tk.Frame(interno, bg=LINHA, width=1).pack(side="left", fill="y", padx=20, pady=4)
        textos = tk.Frame(interno, bg=BRANCO)
        textos.pack(side="left")
        tk.Label(textos, text="Envio de contracheques", bg=BRANCO, fg=TINTA,
                 font=("Segoe UI Semibold", 16)).pack(anchor="w")
        tk.Label(textos, text=f"Recursos Humanos, {self.cfg.nome_empresa}", bg=BRANCO, fg=CINZA_TEXTO,
                 font=("Segoe UI", 10)).pack(anchor="w")
        self.lbl_selo = tk.Label(interno, textvariable=self.var_selo, font=("Segoe UI", 9, "bold"),
                                 padx=14, pady=6)
        self.lbl_selo.pack(side="right")
        tk.Frame(self, bg=AZUL, height=3).pack(fill="x")
        tk.Frame(self, bg=LINHA, height=1).pack(fill="x")

        self.abas = ttk.Notebook(self)
        self.abas.pack(fill="both", expand=True, padx=18, pady=(14, 16))
        self.aba_envio = ttk.Frame(self.abas, padding=(8, 10))
        self.aba_mensagem = ttk.Frame(self.abas, padding=(8, 10))
        self.aba_config = ttk.Frame(self.abas, padding=(8, 10))
        self.abas.add(self.aba_envio, text="Envio")
        self.abas.add(self.aba_mensagem, text="Mensagem do e-mail")
        self.abas.add(self.aba_config, text="Configurações")
        self._montar_aba_envio(self.aba_envio)
        self._montar_aba_mensagem(self.aba_mensagem)
        self._montar_aba_config(self.aba_config)

    # ------------------------------------------------------------------ aba Envio
    def _montar_aba_envio(self, corpo):
        etapas = tk.Frame(corpo, bg=FUNDO)
        etapas.pack(fill="x", pady=(0, 10))
        self.etapas = []
        for n, texto in enumerate(("Arquivos e planilha", "Conferência", "Envio"), start=1):
            if n > 1:
                tk.Frame(etapas, bg=LINHA, height=2, width=48).pack(side="left", padx=10)
            bloco = tk.Frame(etapas, bg=FUNDO)
            bloco.pack(side="left")
            numero = tk.Label(bloco, text=str(n), width=2, font=("Segoe UI Semibold", 10), padx=4, pady=1)
            numero.pack(side="left")
            rotulo = tk.Label(bloco, text=texto, bg=FUNDO, font=("Segoe UI", 10), padx=8)
            rotulo.pack(side="left")
            self.etapas.append((numero, rotulo))
        self._marcar_etapa(1)

        card = ttk.Frame(corpo, style="Painel.TFrame", padding=18)
        card.pack(fill="x")
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="Origem dos contracheques", style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        cmb_origem = ttk.Combobox(card, textvariable=self.var_origem, values=list(ORIGENS.values()),
                                  state="readonly", width=52)
        cmb_origem.grid(row=0, column=1, sticky="w", padx=8)
        cmb_origem.bind("<<ComboboxSelected>>", lambda _e: self._mostrar_origem())

        # Linha do PDF único (só aparece na origem "PDF único")
        self.lbl_unico = ttk.Label(card, text="PDF da contabilidade", style="Card.TLabel")
        self.ent_unico = ttk.Entry(card, textvariable=self.var_pdf_unico)
        self.frm_unico_botoes = ttk.Frame(card, style="Card.TFrame")
        ttk.Button(self.frm_unico_botoes, text="Selecionar…", style="Secundario.TButton",
                   command=self._escolher_pdf_unico).pack(side="left")
        self.btn_separar = ttk.Button(self.frm_unico_botoes, text="Separar e renomear", style="Destaque.TButton",
                                      command=self._separar)
        self.btn_separar.pack(side="left", padx=(8, 0))

        self.lbl_pasta = ttk.Label(card, text="Pasta dos PDFs", style="Card.TLabel")
        self.lbl_pasta.grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.var_pasta).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(card, text="Selecionar…", style="Secundario.TButton",
                   command=lambda: self._escolher_pasta(self.var_pasta, True)).grid(row=2, column=2, sticky="w")

        ttk.Label(card, text="Planilha de colaboradores", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.var_planilha).grid(row=3, column=1, sticky="ew", padx=8)
        ttk.Button(card, text="Selecionar…", style="Secundario.TButton",
                   command=lambda: self._escolher_planilha(self.var_planilha, True)).grid(row=3, column=2, sticky="w")

        linha = ttk.Frame(card, style="Card.TFrame")
        linha.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        ttk.Label(linha, text="Competência", style="Card.TLabel").pack(side="left")
        ttk.Combobox(linha, textvariable=self.var_competencia, width=26,
                     values=[COMP_AUTO, nucleo.competencia_padrao()]).pack(side="left", padx=(8, 24))
        ttk.Checkbutton(linha, text="Conferir código e nome no recibo", variable=self.var_conferir_nome,
                        style="Card.TCheckbutton").pack(side="left", padx=(0, 24))
        ttk.Checkbutton(linha, text="Modo teste, enviar tudo para:", variable=self.var_teste,
                        style="Card.TCheckbutton", command=self._atualizar_botoes).pack(side="left")
        self.ent_teste = ttk.Entry(linha, textvariable=self.var_email_teste, width=32)
        self.ent_teste.pack(side="left", padx=8)

        self.lbl_dica = ttk.Label(card, style="Dica.TLabel", wraplength=1100, justify="left")
        self.lbl_dica.grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))
        self.card_envio = card
        self._atualizar_dica()
        self._mostrar_origem()

        botoes = ttk.Frame(corpo, padding=(0, 12))
        botoes.pack(fill="x")
        self.btn_conferir = ttk.Button(botoes, text="Conferir arquivos", style="Primario.TButton",
                                       command=self._conferir)
        self.btn_conferir.pack(side="left")
        self.btn_enviar = ttk.Button(botoes, text="Enviar contracheques", style="Sucesso.TButton",
                                     command=self._enviar, state="disabled")
        self.btn_enviar.pack(side="left", padx=8)
        self.btn_cancelar = ttk.Button(botoes, text="Parar", style="Secundario.TButton",
                                       command=self._pedir_cancelamento, state="disabled")
        self.btn_cancelar.pack(side="left")
        ttk.Button(botoes, text="Abrir histórico", style="Secundario.TButton",
                   command=self._abrir_logs).pack(side="right")
        self.frm_contadores = tk.Frame(botoes, bg=FUNDO)
        self.frm_contadores.pack(side="right", padx=16)

        quadro = ttk.Frame(corpo)
        quadro.pack(fill="both", expand=True)
        colunas = ("situacao", "codigo", "nome", "tipo", "competencia", "email", "arquivo", "detalhe")
        self.tree = ttk.Treeview(quadro, columns=colunas, show="headings", selectmode="browse", height=8)
        for col, titulo, largura, estica in [
            ("situacao", "Situação", 110, False), ("codigo", "Código", 60, False),
            ("nome", "Nome", 200, True), ("tipo", "Tipo", 130, False), ("competencia", "Compet.", 75, False),
            ("email", "E-mail", 200, True), ("arquivo", "Arquivo", 220, True), ("detalhe", "Observação", 300, True),
        ]:
            self.tree.heading(col, text=titulo)
            self.tree.column(col, width=largura, stretch=estica, anchor="w")
        for situacao, cor in CORES_SITUACAO.items():
            self.tree.tag_configure(situacao, background=cor)
        barra = ttk.Scrollbar(quadro, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=barra.set)
        self.tree.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._duplo_clique)
        self.tree.bind("<Button-3>", self._menu_contexto)
        self.menu_item = tk.Menu(self, tearoff=0)
        self.menu_item.add_command(label="Abrir PDF para conferência", command=self._abrir_pdf_selecionado)
        self.menu_item.add_command(label="Liberar para envio", command=self._liberar_selecionado)
        self.menu_item.add_command(label="Mostrar na pasta", command=self._mostrar_na_pasta)

        rodape = ttk.Frame(corpo, padding=(0, 10, 0, 0))
        rodape.pack(fill="x", side="bottom", before=quadro)
        self.progresso = ttk.Progressbar(rodape, mode="determinate")
        self.progresso.pack(fill="x")
        ttk.Label(rodape, textvariable=self.var_status, style="Status.TLabel").pack(anchor="w", pady=(4, 4))
        self.txt_log = tk.Text(rodape, height=4, font=("Consolas", 9), bg="#FBFCFD", fg="#333",
                               relief="flat", highlightthickness=1, highlightbackground="#DDE3EA")
        self.txt_log.pack(fill="x")
        self.txt_log.configure(state="disabled")
        self._atualizar_botoes()

    def _mostrar_origem(self):
        unico = self.var_origem.get() == ORIGENS["unico"]
        if unico:
            self.lbl_unico.grid(row=1, column=0, sticky="w", pady=4)
            self.ent_unico.grid(row=1, column=1, sticky="ew", padx=8)
            self.frm_unico_botoes.grid(row=1, column=2, sticky="w")
            self.lbl_pasta.configure(text="Pasta dos PDFs separados")
        else:
            for w in (self.lbl_unico, self.ent_unico, self.frm_unico_botoes):
                w.grid_forget()
            self.lbl_pasta.configure(text="Pasta dos PDFs")

    def _escolher_pdf_unico(self):
        arq = filedialog.askopenfilename(title="PDF da contabilidade com todos os contracheques",
                                         filetypes=[("PDF", "*.pdf")])
        if arq:
            self.var_pdf_unico.set(arq)

    def _separar(self):
        origem_txt = self.var_pdf_unico.get().strip()
        if not origem_txt or not Path(origem_txt).is_file():
            messagebox.showwarning("PDF da contabilidade", "Selecione o PDF único enviado pela contabilidade.")
            return
        origem = Path(origem_txt)
        self._ocupar(True, "Lendo o PDF da contabilidade…")
        self.progresso.configure(mode="determinate", value=0, maximum=100)
        cfg = self.cfg

        def progresso(n, total):
            self.fila.put(("progresso_simples", n, total, f"Separando página {n} de {total}…"))

        def tarefa():
            try:
                resultado = nucleo.separar_pdf_unico(origem, cfg, limpar_pasta=False, progresso=progresso)
                self.fila.put(("separado", resultado))
            except Exception as exc:
                nucleo.log.exception("Falha ao separar o PDF único")
                self.fila.put(("erro", "Não foi possível separar o PDF", str(exc)))

        threading.Thread(target=tarefa, daemon=True).start()

    def _pos_separacao(self, r: "nucleo.ResultadoSeparacao"):
        # PDFs que já estavam na pasta e não vieram desta separação (ex.: rodada anterior)
        sobras = sorted(set(r.pasta.glob("*.pdf")) - set(r.arquivos))
        if sobras and messagebox.askyesno(
            "Arquivos antigos na pasta",
            f"A pasta {r.pasta.name} já tinha {len(sobras)} PDF(s) que não vieram deste arquivo.\n\n"
            "Remover esses arquivos antigos para não misturar com esta folha?\n"
            "(A subpasta Enviados não é alterada.)",
        ):
            for antigo in sobras:
                antigo.unlink(missing_ok=True)
            nucleo.log.info("%d PDF(s) antigos removidos de %s", len(sobras), r.pasta)

        self.var_pasta.set(str(r.pasta))
        if len(r.competencias) == 1:
            self.var_competencia.set(r.competencias[0])
        texto = (f"{r.paginas} página(s) lidas → {len(r.arquivos)} contracheque(s) separados.\n"
                 f"Competência: {', '.join(r.competencias) or 'não identificada'}\n"
                 f"Tipo: {', '.join(r.tipos) or 'não identificado'}\n\nPasta: {r.pasta}")
        if r.problemas:
            texto += "\n\nATENÇÃO — páginas não separadas:\n• " + "\n• ".join(r.problemas[:12])
            if len(r.problemas) > 12:
                texto += f"\n… e mais {len(r.problemas) - 12} (veja o log)."
            for pbm in r.problemas:
                nucleo.log.warning(pbm)
        (messagebox.showwarning if r.problemas else messagebox.showinfo)("Separação concluída", texto)
        self._ocupar(False, "Separação concluída. Conferindo…")
        self._conferir()

    # ------------------------------------------------------------------ aba Mensagem
    def _montar_aba_mensagem(self, corpo):
        corpo.columnconfigure(0, weight=3)
        corpo.columnconfigure(1, weight=1)
        corpo.rowconfigure(0, weight=1)

        editor = ttk.Frame(corpo, style="Painel.TFrame", padding=18)
        editor.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        editor.columnconfigure(1, weight=1)
        editor.rowconfigure(3, weight=1)
        ttk.Label(editor, text="Assunto", style="Titulo.TLabel").grid(row=0, column=0, sticky="w")
        self.ent_assunto = ttk.Entry(editor, textvariable=self.c_assunto)
        self.ent_assunto.grid(row=0, column=1, sticky="ew", padx=(10, 0))
        self.ent_assunto.bind("<FocusIn>", lambda _e: setattr(self, "_foco_editor", "assunto"))
        ttk.Label(editor, text="Texto do e-mail", style="Titulo.TLabel").grid(row=2, column=0, columnspan=2,
                                                                             sticky="w", pady=(16, 6))
        quadro = ttk.Frame(editor, style="Card.TFrame")
        quadro.grid(row=3, column=0, columnspan=2, sticky="nsew")
        quadro.columnconfigure(0, weight=1)
        quadro.rowconfigure(0, weight=1)
        self.txt_mensagem = tk.Text(quadro, wrap="word", font=("Segoe UI", 11), undo=True, padx=12, pady=10,
                                    relief="flat", highlightthickness=1, highlightbackground="#DDE3EA",
                                    highlightcolor=DESTAQUE)
        self.txt_mensagem.grid(row=0, column=0, sticky="nsew")
        barra = ttk.Scrollbar(quadro, orient="vertical", command=self.txt_mensagem.yview)
        barra.grid(row=0, column=1, sticky="ns")
        self.txt_mensagem.configure(yscrollcommand=barra.set)
        self.txt_mensagem.bind("<FocusIn>", lambda _e: setattr(self, "_foco_editor", "texto"))
        self._foco_editor = "texto"
        ttk.Label(editor, style="Dica.TLabel",
                  text="Linha em branco separa parágrafos. Use **texto** para negrito. "
                       "Os campos entre { } são trocados pelos dados de cada colaborador.").grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        acoes = ttk.Frame(editor, style="Card.TFrame")
        acoes.grid(row=5, column=0, columnspan=2, sticky="w", pady=(14, 0))
        ttk.Button(acoes, text="Salvar mensagem", style="Primario.TButton",
                   command=self._salvar_mensagem).pack(side="left")
        ttk.Button(acoes, text="Pré-visualizar no navegador", style="Destaque.TButton",
                   command=self._previsualizar).pack(side="left", padx=8)
        ttk.Button(acoes, text="Restaurar texto padrão", style="Secundario.TButton",
                   command=self._restaurar_mensagem).pack(side="left")
        self.var_status_msg = tk.StringVar()
        ttk.Label(acoes, textvariable=self.var_status_msg, style="Card.TLabel").pack(side="left", padx=12)

        lateral = ttk.Frame(corpo, style="Painel.TFrame", padding=18)
        lateral.grid(row=0, column=1, sticky="nsew")
        ttk.Label(lateral, text="Inserir campo", style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(lateral, text="Clique para inserir onde está o cursor.", style="Dica.TLabel").pack(anchor="w",
                                                                                                    pady=(0, 10))
        for marcador, descricao in nucleo.MARCADORES.items():
            ttk.Button(lateral, text=descricao, style="Marcador.TButton",
                       command=lambda m=marcador: self._inserir_marcador(m)).pack(fill="x", pady=2)
        ttk.Label(lateral, style="Dica.TLabel", wraplength=230, justify="left",
                  text="Importante: mantenha o campo 'Como abrir o PDF' no texto — é ele que diz ao "
                       "colaborador qual é a senha.").pack(anchor="w", pady=(14, 0))
        self._carregar_mensagem()

    def _carregar_mensagem(self):
        self.txt_mensagem.delete("1.0", "end")
        self.txt_mensagem.insert("1.0", nucleo.ler_mensagem(self.base))
        self.txt_mensagem.edit_reset()

    def _inserir_marcador(self, marcador: str):
        if getattr(self, "_foco_editor", "texto") == "assunto":
            self.ent_assunto.insert("insert", marcador)
            self.ent_assunto.focus_set()
        else:
            self.txt_mensagem.insert("insert", marcador)
            self.txt_mensagem.focus_set()

    def _restaurar_mensagem(self):
        if messagebox.askyesno("Mensagem", "Substituir o texto atual pelo texto padrão?"):
            self.txt_mensagem.delete("1.0", "end")
            self.txt_mensagem.insert("1.0", nucleo.MENSAGEM_PADRAO)
            self.c_assunto.set(nucleo.Config.assunto)

    def _item_exemplo(self) -> "nucleo.ItemEnvio":
        prontos = [i for i in self.itens if i.colaborador and i.competencia]
        if prontos:
            return prontos[0]
        return nucleo.ItemEnvio("1234", "MARIA SILVA", "exemplo@empresa.com.br", None, Situacao.PRONTO,
                                competencia=nucleo.competencia_padrao(), tipo_folha="Folha Mensal")

    def _previsualizar(self):
        texto = self.txt_mensagem.get("1.0", "end")
        cfg = dataclasses.replace(self.cfg, assunto=self.c_assunto.get())
        item = self._item_exemplo()
        html_email = nucleo.montar_corpo(texto, item, cfg)
        assunto = nucleo.montar_assunto(item, cfg)
        cabecalho = (f'<div style="font-family:Segoe UI,Arial;background:#fff;border-bottom:1px solid #ddd;'
                     f'padding:12px 20px;font-size:13px;color:#333"><b>Pré-visualização</b> — exemplo com '
                     f'{html.escape(item.nome)}<br><b>Assunto:</b> {html.escape(assunto)}</div>')
        arquivo = Path(tempfile.gettempdir()) / "previa_contracheque.html"
        arquivo.write_text(html_email.replace(
            '<table role="presentation" width="100%"', cabecalho + '<table role="presentation" width="100%"', 1),
            encoding="utf-8")
        webbrowser.open(arquivo.as_uri())

    def _salvar_mensagem(self):
        texto = self.txt_mensagem.get("1.0", "end").strip()
        assunto = self.c_assunto.get().strip()
        if not texto or not assunto:
            messagebox.showwarning("Mensagem", "Assunto e texto não podem ficar vazios.")
            return
        if "{dica_senha}" not in texto and not messagebox.askyesno(
            "Mensagem", "O texto não tem o campo 'Como abrir o PDF' ({dica_senha}).\n"
                        "Sem ele o colaborador não saberá a senha. Salvar mesmo assim?"):
            return
        try:
            nucleo.salvar_mensagem(self.base, texto)
            nucleo.salvar_configuracoes(self.arq_usuario, {"email": {"assunto": assunto}})
            self._recarregar_cfg()
        except Exception as exc:
            messagebox.showerror("Mensagem", f"Não foi possível salvar:\n\n{exc}")
            return
        self.var_status_msg.set(f"Salvo às {datetime.now():%H:%M}.")
        nucleo.log.info("Mensagem do e-mail atualizada.")

    # ------------------------------------------------------------------ aba Configurações
    def _montar_aba_config(self, corpo):
        corpo.columnconfigure(0, weight=1, uniform="c")
        corpo.columnconfigure(1, weight=1, uniform="c")

        # ----- Arquivos
        arq = ttk.Labelframe(corpo, text="Arquivos padrão", style="Card.TLabelframe")
        arq.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))
        arq.columnconfigure(1, weight=1)
        ttk.Label(arq, text="Pasta dos PDFs", style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(arq, textvariable=self.c_pasta).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(arq, text="Selecionar…", style="Secundario.TButton",
                   command=lambda: self._escolher_pasta(self.c_pasta, False)).grid(row=0, column=2)
        ttk.Label(arq, text="Planilha", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(arq, textvariable=self.c_planilha).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(arq, text="Selecionar…", style="Secundario.TButton",
                   command=lambda: self._escolher_planilha(self.c_planilha, False)).grid(row=1, column=2)
        ttk.Label(arq, style="Dica.TLabel", wraplength=480, justify="left",
                  text="Já vêm preenchidas na aba Envio. Guarde a planilha em pasta com acesso restrito ao RH."
                  ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # ----- Mensagem e segurança
        msg = ttk.Labelframe(corpo, text="Mensagem e proteção do PDF", style="Card.TLabelframe")
        msg.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))
        msg.columnconfigure(1, weight=1)
        linhas = [("Nome da empresa", self.c_empresa), ("E-mail para testes", self.c_email_teste)]
        for i, (rotulo, var) in enumerate(linhas):
            ttk.Label(msg, text=rotulo, style="Card.TLabel").grid(row=i, column=0, sticky="w", pady=4)
            ttk.Entry(msg, textvariable=var).grid(row=i, column=1, columnspan=2, sticky="ew", padx=(8, 0))
        ttk.Label(msg, style="Dica.TLabel", wraplength=360, justify="left",
                  text="Assunto e texto do e-mail ficam na aba Mensagem do e-mail.").grid(
            row=3, column=1, columnspan=2, sticky="w", padx=8)
        ttk.Label(msg, text="Senha do PDF", style="Card.TLabel").grid(row=4, column=0, sticky="w", pady=(10, 4))
        ttk.Combobox(msg, textvariable=self.c_formato, values=list(FORMATOS_SENHA), state="readonly").grid(
            row=4, column=1, columnspan=2, sticky="ew", padx=(8, 0), pady=(10, 4))
        ttk.Label(msg, text="Pausa entre e-mails (s)", style="Card.TLabel").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Spinbox(msg, from_=0, to=30, increment=1, textvariable=self.c_pausa, width=6).grid(
            row=5, column=1, sticky="w", padx=8)
        ttk.Checkbutton(msg, text="Mover PDFs enviados para a subpasta Enviados", variable=self.c_mover,
                        style="Card.TCheckbutton").grid(row=6, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Checkbutton(msg, text="Conferir código e nome dentro do recibo (recomendado)", variable=self.c_conferir,
                        style="Card.TCheckbutton").grid(row=7, column=0, columnspan=3, sticky="w")

        # ----- Forma de envio
        env = ttk.Labelframe(corpo, text="Como os e-mails saem", style="Card.TLabelframe")
        env.grid(row=0, column=1, rowspan=2, sticky="nsew", pady=(0, 8))
        env.columnconfigure(0, weight=1)
        for i, (chave, texto) in enumerate(MODOS.items()):
            ttk.Radiobutton(env, text=texto, value=chave, variable=self.c_modo, style="Card.TRadiobutton",
                            command=self._mostrar_bloco_envio).grid(row=i, column=0, sticky="w", pady=2)

        # Outlook
        self.bloco_outlook = ttk.Frame(env, style="Card.TFrame", padding=(18, 10, 0, 0))
        self.bloco_outlook.columnconfigure(1, weight=1)
        ttk.Label(self.bloco_outlook, text="Conta que envia", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        self.cmb_contas = ttk.Combobox(self.bloco_outlook, textvariable=self.c_remetente_outlook)
        self.cmb_contas.grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(self.bloco_outlook, text="Buscar contas", style="Secundario.TButton",
                   command=self._buscar_contas_outlook).grid(row=0, column=2)
        ttk.Label(self.bloco_outlook, style="Dica.TLabel", wraplength=470, justify="left",
                  text="Vazio = conta padrão do Outlook. Pode ser uma caixa compartilhada (ex.: rh@empresa), "
                       "desde que este usuário tenha permissão 'Enviar como'. Use o Outlook CLÁSSICO, "
                       "aberto e logado — o 'novo Outlook' não aceita automação."
                  ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))

        # SMTP
        self.bloco_smtp = ttk.Frame(env, style="Card.TFrame", padding=(18, 10, 0, 0))
        self.bloco_smtp.columnconfigure(1, weight=1)
        campos = [("Servidor", self.c_smtp_host), ("Porta", self.c_smtp_porta),
                  ("Usuário (login)", self.c_smtp_usuario)]
        for i, (rotulo, var) in enumerate(campos):
            ttk.Label(self.bloco_smtp, text=rotulo, style="Card.TLabel").grid(row=i, column=0, sticky="w", pady=3)
            ttk.Entry(self.bloco_smtp, textvariable=var, width=12 if rotulo == "Porta" else 40).grid(
                row=i, column=1, sticky="w" if rotulo == "Porta" else "ew", padx=8)
        ttk.Label(self.bloco_smtp, text="Segurança", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=3)
        cmb_seg = ttk.Combobox(self.bloco_smtp, textvariable=self.c_smtp_seg, values=list(SEGURANCAS),
                               state="readonly", width=22)
        cmb_seg.grid(row=3, column=1, sticky="w", padx=8)
        cmb_seg.bind("<<ComboboxSelected>>", self._sugerir_porta)
        ttk.Label(self.bloco_smtp, text="Senha", style="Card.TLabel").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Entry(self.bloco_smtp, textvariable=self.c_smtp_senha, show="•").grid(row=4, column=1, sticky="ew", padx=8)
        ttk.Label(self.bloco_smtp, textvariable=self.var_status_senha, style="Ok.TLabel").grid(
            row=5, column=1, sticky="w", padx=8)
        ttk.Label(self.bloco_smtp, text="Remetente", style="Card.TLabel").grid(row=6, column=0, sticky="w", pady=3)
        ttk.Entry(self.bloco_smtp, textvariable=self.c_smtp_remetente).grid(row=6, column=1, sticky="ew", padx=8)
        ttk.Label(self.bloco_smtp, style="Dica.TLabel", text="Vazio = o próprio usuário.").grid(
            row=7, column=1, sticky="w", padx=8)
        acoes = ttk.Frame(self.bloco_smtp, style="Card.TFrame")
        acoes.grid(row=8, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(acoes, text="Testar conexão", style="Secundario.TButton",
                   command=self._testar_smtp).pack(side="left")
        ttk.Button(acoes, text="Apagar senha salva", style="Secundario.TButton",
                   command=self._apagar_senha).pack(side="left", padx=8)
        ttk.Label(self.bloco_smtp, style="Dica.TLabel", wraplength=470, justify="left",
                  text="A senha é guardada criptografada pelo Windows e só abre neste usuário e neste "
                       "computador. Gmail exige 'senha de app'. No Microsoft 365, login com senha "
                       "(SMTP AUTH) costuma estar desativado — prefira o modo Outlook."
                  ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(10, 0))

        self.bloco_simulado = ttk.Frame(env, style="Card.TFrame", padding=(18, 10, 0, 0))
        ttk.Label(self.bloco_simulado, style="Dica.TLabel", wraplength=470, justify="left",
                  text="Nada sai por e-mail. Os PDFs protegidos e uma cópia de cada e-mail vão para "
                       "logs\\simulacao. Use para conferir senhas e layout antes do primeiro envio real."
                  ).grid(row=0, column=0, sticky="w")

        # ----- Rodapé
        rodape = ttk.Frame(corpo)
        rodape.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        ttk.Button(rodape, text="Salvar configurações", style="Primario.TButton",
                   command=self._salvar_config).pack(side="left")
        ttk.Button(rodape, text="Desfazer alterações", style="Secundario.TButton",
                   command=self._carregar_formulario).pack(side="left", padx=8)
        ttk.Label(rodape, textvariable=self.var_status_config, style="Resumo.TLabel").pack(side="left", padx=12)

    def _mostrar_bloco_envio(self):
        for bloco in (self.bloco_outlook, self.bloco_smtp, self.bloco_simulado):
            bloco.grid_forget()
        bloco = {"outlook": self.bloco_outlook, "smtp": self.bloco_smtp}.get(self.c_modo.get(), self.bloco_simulado)
        bloco.grid(row=len(MODOS), column=0, sticky="nsew")

    def _sugerir_porta(self, _evento=None):
        seg = SEGURANCAS.get(self.c_smtp_seg.get(), "starttls")
        if self.c_smtp_porta.get().strip() in ("", "465", "587"):
            self.c_smtp_porta.set("465" if seg == "ssl" else "587")

    def _carregar_formulario(self):
        c = self.cfg
        self.c_pasta.set(c.pasta_pdfs)
        self.c_planilha.set(c.planilha)
        self.c_modo.set(c.modo_envio)
        self.c_remetente_outlook.set(c.remetente)
        self.c_smtp_host.set(c.smtp_host)
        self.c_smtp_porta.set(str(c.smtp_porta))
        self.c_smtp_seg.set(next((k for k, v in SEGURANCAS.items() if v == c.smtp_seguranca), "STARTTLS (porta 587)"))
        self.c_smtp_usuario.set(c.smtp_usuario)
        self.c_smtp_senha.set("")
        self.c_smtp_remetente.set(c.smtp_remetente)
        self.c_empresa.set(c.nome_empresa)
        self.c_assunto.set(c.assunto)
        self.c_email_teste.set(c.email_teste)
        self.c_formato.set(next((k for k, v in FORMATOS_SENHA.items() if v == c.formato_senha),
                                list(FORMATOS_SENHA)[0]))
        self.c_pausa.set(str(int(c.pausa_segundos)) if float(c.pausa_segundos).is_integer() else str(c.pausa_segundos))
        self.c_mover.set(c.mover_enviados)
        self.c_conferir.set(c.conferir_nome)
        self._atualizar_status_senha()
        self._mostrar_bloco_envio()
        self.var_status_config.set("")

    def _atualizar_status_senha(self):
        self.var_status_senha.set("✓ Senha salva (deixe em branco para manter)" if cofre.tem_senha(self.base) else "")

    def _cfg_do_formulario(self) -> nucleo.Config:
        """Monta uma Config com o que está digitado (sem salvar) — usada para validar e testar."""
        try:
            porta = int(self.c_smtp_porta.get().strip() or 0)
        except ValueError:
            raise ValueError("Porta SMTP deve ser um número.")
        try:
            pausa = float(self.c_pausa.get().replace(",", ".") or 0)
        except ValueError:
            raise ValueError("Pausa deve ser um número de segundos.")
        return dataclasses.replace(
            self.cfg,
            pasta_pdfs=self.c_pasta.get().strip(),
            planilha=self.c_planilha.get().strip(),
            modo_envio=self.c_modo.get(),
            remetente=self.c_remetente_outlook.get().strip(),
            smtp_host=self.c_smtp_host.get().strip(),
            smtp_porta=porta,
            smtp_seguranca=SEGURANCAS.get(self.c_smtp_seg.get(), "starttls"),
            smtp_usuario=self.c_smtp_usuario.get().strip(),
            smtp_remetente=self.c_smtp_remetente.get().strip(),
            nome_empresa=self.c_empresa.get().strip() or "Empresa",
            assunto=self.c_assunto.get().strip() or nucleo.Config.assunto,
            email_teste=self.c_email_teste.get().strip().lower(),
            formato_senha=FORMATOS_SENHA.get(self.c_formato.get(), "{cpf4}"),
            pausa_segundos=max(0.0, pausa),
            mover_enviados=self.c_mover.get(),
            conferir_nome=self.c_conferir.get(),
        )

    def _validar(self, c: nucleo.Config) -> list[str]:
        erros = []
        if c.email_teste and not nucleo.RE_EMAIL.match(c.email_teste):
            erros.append("E-mail para testes inválido.")
        if c.modo_envio == "outlook" and c.remetente and not nucleo.RE_EMAIL.match(c.remetente):
            erros.append("Conta do Outlook inválida.")
        if c.modo_envio == "smtp":
            if not c.smtp_host:
                erros.append("Informe o servidor SMTP.")
            if not (1 <= c.smtp_porta <= 65535):
                erros.append("Porta SMTP inválida.")
            if not nucleo.RE_EMAIL.match(c.smtp_usuario or ""):
                erros.append("Usuário SMTP deve ser um e-mail.")
            if c.smtp_remetente and not nucleo.RE_EMAIL.match(c.smtp_remetente):
                erros.append("Remetente inválido.")
            if not self.c_smtp_senha.get() and not cofre.tem_senha(self.base):
                erros.append("Informe a senha do e-mail.")
            if not cofre.disponivel():
                erros.append("Guardar senha só funciona no Windows.")
        return erros

    def _salvar_config(self):
        try:
            nova = self._cfg_do_formulario()
        except ValueError as exc:
            messagebox.showwarning("Configurações", str(exc))
            return
        erros = self._validar(nova)
        if erros:
            messagebox.showwarning("Configurações", "Corrija antes de salvar:\n\n• " + "\n• ".join(erros))
            return
        try:
            if nova.modo_envio == "smtp" and self.c_smtp_senha.get():
                cofre.salvar_senha(self.base, self.c_smtp_senha.get())
            nucleo.salvar_configuracoes(self.arq_usuario, {
                "geral": {"formato_senha": nova.formato_senha, "pausa_segundos": nova.pausa_segundos,
                          "mover_enviados": nova.mover_enviados, "conferir_conteudo_pdf": nova.conferir_nome},
                "email": {"modo_envio": nova.modo_envio, "remetente": nova.remetente,
                          "assunto": nova.assunto, "nome_empresa": nova.nome_empresa},
                "smtp": {"host": nova.smtp_host, "porta": nova.smtp_porta, "seguranca": nova.smtp_seguranca,
                         "usuario": nova.smtp_usuario, "remetente": nova.smtp_remetente},
                "ultimo_uso": {"pasta_pdfs": nova.pasta_pdfs, "planilha": nova.planilha,
                               "email_teste": nova.email_teste},
            })
            self._recarregar_cfg()
        except Exception as exc:
            nucleo.log.exception("Falha ao salvar configurações")
            messagebox.showerror("Configurações", f"Não foi possível salvar:\n\n{exc}")
            return

        self.c_smtp_senha.set("")
        self._atualizar_status_senha()
        # Leva os padrões para a aba Envio
        self.var_pasta.set(self.cfg.pasta_pdfs or self.var_pasta.get())
        self.var_planilha.set(self.cfg.planilha or self.var_planilha.get())
        self.var_email_teste.set(self.cfg.email_teste or self.var_email_teste.get())
        self.var_conferir_nome.set(self.cfg.conferir_nome)
        self._invalidar_conferencia()
        self._atualizar_selo()
        self.var_status_config.set(f"Configurações salvas às {datetime.now():%H:%M}.")
        nucleo.log.info("Configurações salvas. Modo de envio: %s", self.cfg.modo_envio)

    def _testar_smtp(self):
        try:
            teste = self._cfg_do_formulario()
        except ValueError as exc:
            messagebox.showwarning("Testar conexão", str(exc))
            return
        teste = dataclasses.replace(teste, modo_envio="smtp")
        erros = [e for e in self._validar(teste) if "senha" not in e.lower() or not self.c_smtp_senha.get()]
        if erros:
            messagebox.showwarning("Testar conexão", "\n".join(erros))
            return
        senha_digitada = self.c_smtp_senha.get()
        self.var_status_config.set("Testando conexão…")

        def tarefa():
            try:
                senha = senha_digitada or cofre.ler_senha(self.base)
                self.fila.put(("info", "Conexão OK", nucleo.testar_smtp(teste, senha)))
                self.fila.put(("status_config", "Conexão SMTP testada com sucesso."))
            except Exception as exc:
                self.fila.put(("erro_simples", "Falha na conexão", str(exc)))
                self.fila.put(("status_config", "Falha no teste de conexão."))

        threading.Thread(target=tarefa, daemon=True).start()

    def _buscar_contas_outlook(self):
        self.var_status_config.set("Consultando o Outlook…")

        def tarefa():
            try:
                self.fila.put(("contas", nucleo.listar_contas_outlook()))
            except Exception as exc:
                self.fila.put(("erro_simples", "Outlook",
                               f"Não consegui falar com o Outlook. Ele precisa ser o Outlook clássico, "
                               f"aberto e logado.\n\n{exc}"))
                self.fila.put(("status_config", ""))

        threading.Thread(target=tarefa, daemon=True).start()

    def _apagar_senha(self):
        if cofre.tem_senha(self.base) and messagebox.askyesno("Senha", "Apagar a senha de e-mail salva?"):
            cofre.apagar_senha(self.base)
            self._atualizar_status_senha()
            nucleo.log.info("Senha SMTP apagada pelo usuário.")

    def _atualizar_selo(self):
        c = self.cfg
        if c.modo_envio == "simulado":
            self.var_selo.set("Modo simulado: nenhum e-mail sai")
            self.lbl_selo.configure(bg="#FFF4DA", fg="#7A5200")
        elif c.modo_envio == "outlook":
            self.var_selo.set(f"Enviando pelo Outlook: {c.remetente or 'conta padrão'}")
            self.lbl_selo.configure(bg="#E8F4EC", fg="#22643F")
        else:
            self.var_selo.set(f"Enviando por SMTP: {c.smtp_remetente or c.smtp_usuario}")
            self.lbl_selo.configure(bg=AZUL_SUAVE, fg=AZUL)

    def _atualizar_dica(self):
        if self.cfg.identificar_por == "nome":
            dica = ("O arquivo deve começar pelo nome do colaborador igual à planilha "
                    "(ex.: Maria_Silva_Recibo_de_Salário_09_2026.pdf). Código e nome do recibo precisam bater.")
        else:
            dica = "O arquivo deve começar pelo código do colaborador (ex.: 1234_MARIA.pdf)."
        self.lbl_dica.configure(text=dica + " Itens amarelos: confira e libere com duplo clique.")

    # ================================================================== ações da aba Envio
    def _escolher_pasta(self, var: tk.StringVar, invalidar: bool):
        pasta = filedialog.askdirectory(title="Pasta com os contracheques em PDF", initialdir=var.get() or None)
        if pasta:
            var.set(pasta)
            if invalidar:
                self._invalidar_conferencia()

    def _escolher_planilha(self, var: tk.StringVar, invalidar: bool):
        arq = filedialog.askopenfilename(title="Planilha de colaboradores", filetypes=[("Excel", "*.xlsx *.xlsm")])
        if arq:
            var.set(arq)
            if invalidar:
                self._invalidar_conferencia()

    def _invalidar_conferencia(self):
        self.itens = []
        self.tree.delete(*self.tree.get_children())
        self._preencher_tabela()

    def _conferir(self):
        pasta_txt, planilha_txt = self.var_pasta.get().strip(), self.var_planilha.get().strip()
        competencia_tela = self.var_competencia.get().strip()
        if not pasta_txt or not planilha_txt:
            messagebox.showwarning("Faltam dados", "Selecione a pasta dos PDFs e a planilha.")
            return
        if competencia_tela == COMP_AUTO:
            competencia = ""
        elif nucleo.competencia_valida(competencia_tela):
            competencia = competencia_tela
        else:
            messagebox.showwarning("Competência inválida",
                                   "Escolha 'Automática' ou informe a competência no formato MM/AAAA.")
            return
        pasta, planilha = Path(pasta_txt), Path(planilha_txt)
        try:
            nucleo.salvar_configuracoes(self.arq_usuario, {"ultimo_uso": {
                "pasta_pdfs": pasta_txt, "planilha": planilha_txt,
                "email_teste": self.var_email_teste.get().strip()}})
        except OSError:
            nucleo.log.warning("Não foi possível gravar o último uso.")

        self._ocupar(True, "Conferindo arquivos…")
        self.progresso.configure(mode="indeterminate")
        self.progresso.start(12)
        conferir_nome = self.var_conferir_nome.get()
        cfg = self.cfg

        def tarefa():
            try:
                itens = nucleo.montar_lote(planilha, pasta, cfg, competencia,
                                           self.base / "logs" / "historico_envios.csv", conferir_nome)
                self.fila.put(("lote", itens, competencia_tela))
            except Exception as exc:
                nucleo.log.exception("Falha na conferência")
                self.fila.put(("erro", "Não foi possível conferir", str(exc)))

        threading.Thread(target=tarefa, daemon=True).start()

    def _enviar(self):
        if self.var_competencia.get().strip() != self.competencia_conferida:
            messagebox.showwarning("Confira novamente",
                                   "A competência mudou depois da conferência. Clique em Conferir de novo.")
            return
        prontos = [i for i in self.itens if i.situacao == Situacao.PRONTO]
        if not prontos:
            messagebox.showinfo("Nada a enviar", "Nenhum item está como 'Pronto'.")
            return
        if self.cfg.modo_envio == "smtp" and not cofre.tem_senha(self.base):
            messagebox.showwarning("Senha do e-mail", "Informe a senha do e-mail na aba Configurações.")
            self.abas.select(self.aba_config)
            return

        email_teste = ""
        if self.var_teste.get():
            email_teste = self.var_email_teste.get().strip().lower()
            if not nucleo.RE_EMAIL.match(email_teste):
                messagebox.showwarning("E-mail de teste", "Informe um e-mail válido para o modo teste.")
                return

        revisar = sum(1 for i in self.itens if i.situacao == Situacao.REVISAR)
        erros = sum(1 for i in self.itens if i.situacao == Situacao.ERRO)
        destino = f"TODOS para {email_teste} (teste)" if email_teste else "cada colaborador"
        aviso = ""
        if revisar or erros:
            aviso = f"\n\nAtenção: {erros} com erro e {revisar} para revisar NÃO serão enviados."
        if self.cfg.modo_envio == "simulado":
            aviso += "\n\nModo SIMULADO: nada sai por e-mail, arquivos vão para logs\\simulacao."
        resumo_lote = Counter(f"{i.tipo_folha or 'Folha'} {i.competencia}" for i in prontos)
        descricao = "\n".join(f"• {qtd} × {chave}" for chave, qtd in resumo_lote.most_common())
        if not messagebox.askyesno(
            "Confirmar envio",
            f"Enviar {len(prontos)} contracheque(s) para {destino}?\n\n{descricao}{aviso}",
            icon="warning" if not email_teste else "question",
        ):
            return

        self.cancelar.clear()
        self._ocupar(True, "Enviando…")
        self.progresso.configure(mode="determinate", maximum=len(prontos), value=0)
        self.btn_cancelar.configure(state="normal")
        cfg = self.cfg

        def progresso(n, total, item):
            self.fila.put(("progresso", n, total, item))

        def tarefa():
            try:
                resumo = nucleo.enviar_lote(self.itens, cfg, self.base, email_teste, progresso, self.cancelar)
                self.fila.put(("fim_envio", resumo, bool(email_teste)))
            except Exception as exc:
                nucleo.log.exception("Falha no envio")
                self.fila.put(("erro", "Falha no envio", str(exc)))

        threading.Thread(target=tarefa, daemon=True).start()

    def _pedir_cancelamento(self):
        self.cancelar.set()
        self.var_status.set("Parando após o e-mail atual…")
        self.btn_cancelar.configure(state="disabled")

    def _item_da_linha(self, iid: str):
        return self.itens[int(iid)] if iid else None

    def _duplo_clique(self, evento):
        iid = self.tree.identify_row(evento.y)
        item = self._item_da_linha(iid)
        if not item:
            return
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self._abrir_pdf(item)
        if item.situacao == Situacao.REVISAR and not self.ocupado:
            # Dá tempo do leitor de PDF abrir antes da pergunta
            self.after(1200, lambda: self._liberar(item))

    def _menu_contexto(self, evento):
        iid = self.tree.identify_row(evento.y)
        item = self._item_da_linha(iid)
        if not item:
            return
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        tem_pdf = bool(item.arquivo and item.arquivo.exists())
        self.menu_item.entryconfigure(0, state="normal" if tem_pdf else "disabled")
        self.menu_item.entryconfigure(1, state="normal" if item.situacao == Situacao.REVISAR and not self.ocupado
                                      else "disabled")
        self.menu_item.entryconfigure(2, state="normal" if tem_pdf else "disabled")
        self.menu_item.tk_popup(evento.x_root, evento.y_root)

    def _selecionado(self):
        return self._item_da_linha(self.tree.focus())

    def _abrir_pdf_selecionado(self):
        item = self._selecionado()
        if item:
            self._abrir_pdf(item)

    def _liberar_selecionado(self):
        item = self._selecionado()
        if item:
            self._liberar(item)

    def _abrir_pdf(self, item):
        if not item.arquivo:
            messagebox.showinfo("Abrir PDF", "Este item não tem arquivo PDF.")
            return
        if not item.arquivo.exists():
            messagebox.showwarning("Abrir PDF", f"Arquivo não encontrado (já foi movido?):\n{item.arquivo}")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(item.arquivo)  # noqa: S606 — abre no leitor de PDF padrão
            else:
                subprocess.Popen(["xdg-open", str(item.arquivo)])
            nucleo.log.info("PDF aberto para conferência: %s", item.arquivo.name)
        except OSError as exc:
            messagebox.showerror("Abrir PDF", f"Não foi possível abrir o PDF:\n{exc}")

    def _mostrar_na_pasta(self):
        item = self._selecionado()
        if not item or not item.arquivo or not item.arquivo.exists():
            return
        if sys.platform.startswith("win"):
            subprocess.Popen(["explorer", "/select,", str(item.arquivo)])
        else:
            subprocess.Popen(["xdg-open", str(item.arquivo.parent)])

    def _liberar(self, item):
        if item.situacao != Situacao.REVISAR:
            return
        if messagebox.askyesno(
            "Liberar para envio",
            f"Você conferiu o PDF e confirma que ele pertence a:\n\n{item.nome}  (código {item.codigo})\n"
            f"{item.tipo_folha or 'Tipo não identificado'} — {item.competencia or 'competência não identificada'}"
            f"\n\nArquivo: {item.arquivo.name}",
        ):
            item.situacao = Situacao.PRONTO
            item.detalhe = "Liberado manualmente após conferência visual"
            nucleo.log.info("Liberado manualmente: %s | %s | %s", item.codigo, item.nome, item.arquivo.name)
            self._preencher_tabela()

    def _abrir_logs(self):
        pasta = self.base / "logs"
        pasta.mkdir(exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(pasta)  # noqa: S606
        else:
            subprocess.Popen(["xdg-open", str(pasta)])

    # ================================================================== estado da tela
    def _ocupar(self, ocupado: bool, status: str = ""):
        self.ocupado = ocupado
        if status:
            self.var_status.set(status)
        self._atualizar_botoes()

    def _atualizar_botoes(self):
        tem_prontos = any(i.situacao == Situacao.PRONTO for i in self.itens)
        self.btn_conferir.configure(state="disabled" if self.ocupado else "normal")
        if hasattr(self, "btn_separar"):
            self.btn_separar.configure(state="disabled" if self.ocupado else "normal")
        self.btn_enviar.configure(state="normal" if (tem_prontos and not self.ocupado) else "disabled")
        if not self.ocupado:
            self.btn_cancelar.configure(state="disabled")
        self.ent_teste.configure(state="normal" if self.var_teste.get() else "disabled")
        self.btn_enviar.configure(text="Enviar teste" if self.var_teste.get() else "Enviar contracheques")

    def _preencher_tabela(self):
        self.tree.delete(*self.tree.get_children())
        ordem = {s: i for i, s in enumerate(ORDEM_SITUACAO)}
        indices = sorted(range(len(self.itens)),
                         key=lambda k: (ordem.get(self.itens[k].situacao, 99), self.itens[k].nome))
        for k in indices:
            it = self.itens[k]
            self.tree.insert("", "end", iid=str(k), tags=(it.situacao,), values=(
                it.situacao, it.codigo, it.nome, it.tipo_folha, it.competencia, it.email,
                it.arquivo.name if it.arquivo else "", it.detalhe,
            ))
        contagem = Counter(it.situacao for it in self.itens)
        for w in self.frm_contadores.winfo_children():
            w.destroy()
        for situacao in ORDEM_SITUACAO:
            if not contagem.get(situacao):
                continue
            chip = tk.Frame(self.frm_contadores, bg=CORES_SITUACAO[situacao])
            chip.pack(side="left", padx=3)
            tk.Label(chip, text="●", bg=CORES_SITUACAO[situacao], fg=PONTOS_SITUACAO[situacao],
                     font=("Segoe UI", 9)).pack(side="left", padx=(8, 2), pady=3)
            tk.Label(chip, text=f"{contagem[situacao]} {situacao.lower()}", bg=CORES_SITUACAO[situacao],
                     fg=TINTA, font=("Segoe UI Semibold", 9)).pack(side="left", padx=(0, 9))
        if any(i.situacao in (Situacao.ENVIADO, Situacao.TESTE, Situacao.SIMULADO) for i in self.itens):
            self._marcar_etapa(4)
        elif self.itens:
            self._marcar_etapa(3 if contagem.get(Situacao.PRONTO) else 2)
        else:
            self._marcar_etapa(1)
        self._atualizar_botoes()

    def _marcar_etapa(self, atual: int):
        """1 = escolher arquivos, 2 = conferindo/revisar, 3 = pronto para enviar, 4 = concluído."""
        for n, (numero, rotulo) in enumerate(self.etapas, start=1):
            if n < atual:
                numero.configure(text="✓", bg=AZUL_SUAVE, fg=AZUL)
                rotulo.configure(fg=AZUL)
            elif n == atual:
                numero.configure(text=str(n), bg=AZUL, fg=BRANCO)
                rotulo.configure(fg=TINTA, font=("Segoe UI Semibold", 10))
            else:
                numero.configure(text=str(n), bg=LINHA, fg=CINZA_TEXTO)
                rotulo.configure(fg=CINZA_TEXTO, font=("Segoe UI", 10))

    def _escrever_log(self, texto: str):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", texto + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _processar_fila(self):
        try:
            while True:
                msg = self.fila.get_nowait()
                tipo = msg[0]
                if tipo == "log":
                    self._escrever_log(msg[1])
                elif tipo == "lote":
                    self.progresso.stop()
                    self.progresso.configure(mode="determinate", value=0)
                    self.itens, self.competencia_conferida = msg[1], msg[2]
                    self._ocupar(False, "Conferência concluída. Revise a lista antes de enviar.")
                    self._preencher_tabela()
                elif tipo == "progresso":
                    _, n, total, item = msg
                    self.progresso.configure(value=n)
                    self.var_status.set(f"{n}/{total} — {item.nome}: {item.situacao}")
                    self._preencher_tabela()
                elif tipo == "fim_envio":
                    _, resumo, teste = msg
                    self._ocupar(False, "Envio finalizado.")
                    self._preencher_tabela()
                    texto = f"Enviados: {resumo['enviados']}\nCom erro: {resumo['erros']}"
                    if resumo["cancelado"]:
                        texto += "\n\nO envio foi interrompido pelo usuário."
                    if teste:
                        texto += ("\n\nConfira a caixa de teste e abra alguns PDFs com a senha. "
                                  "Depois desmarque 'Modo teste', clique em Conferir e envie de verdade.")
                    messagebox.showinfo("Teste concluído" if teste else "Envio concluído", texto)
                elif tipo == "erro":
                    self.progresso.stop()
                    self.progresso.configure(mode="determinate", value=0)
                    self._ocupar(False, "Ocorreu um erro. Veja a mensagem.")
                    messagebox.showerror(msg[1], msg[2])
                elif tipo == "separado":
                    self.progresso.configure(value=0)
                    self._pos_separacao(msg[1])
                elif tipo == "progresso_simples":
                    _, n, total, texto = msg
                    self.progresso.configure(maximum=total, value=n)
                    self.var_status.set(texto)
                elif tipo == "erro_simples":
                    messagebox.showerror(msg[1], msg[2])
                elif tipo == "info":
                    messagebox.showinfo(msg[1], msg[2])
                elif tipo == "status_config":
                    self.var_status_config.set(msg[1])
                elif tipo == "contas":
                    contas = msg[1]
                    self.cmb_contas.configure(values=contas)
                    if contas and not self.c_remetente_outlook.get():
                        self.c_remetente_outlook.set(contas[0])
                    self.var_status_config.set(f"{len(contas)} conta(s) encontrada(s) no Outlook.")
        except queue.Empty:
            pass
        self.after(100, self._processar_fila)

    def _ao_fechar(self):
        if self.ocupado and not messagebox.askyesno(
            "Envio em andamento", "Há um processo em andamento. Deseja realmente fechar?"
        ):
            return
        self.cancelar.set()
        self.destroy()


def _registrar_falha_inicial() -> None:
    """Sem console (pythonw), um erro na abertura sumiria. Grava em log e avisa na tela."""
    (BASE / "logs").mkdir(exist_ok=True)
    caminho = BASE / "logs" / "erro_inicializacao.log"
    with open(caminho, "a", encoding="utf-8") as f:
        f.write(f"\n==== {datetime.now():%Y-%m-%d %H:%M:%S} ====\n{traceback.format_exc()}")
    try:
        raiz = tk.Tk()
        raiz.withdraw()
        messagebox.showerror("Envio de Contracheques", f"O programa não conseguiu abrir.\n\nDetalhes em:\n{caminho}")
        raiz.destroy()
    except Exception:
        pass


def main() -> None:
    try:
        App().mainloop()
    except SystemExit:
        raise
    except Exception:
        _registrar_falha_inicial()


if __name__ == "__main__":
    main()
