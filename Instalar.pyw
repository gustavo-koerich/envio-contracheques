"""
Instalador do Envio de Contracheques — duplo clique, uma vez, na máquina do RH.

Não instala nada no Windows, não pede administrador e não altera configuração de segurança:
confere o ambiente, desbloqueia os arquivos DESTA pasta, faz um autoteste e cria os atalhos.
"""

from __future__ import annotations

import gc
import os
import queue
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

BASE = Path(__file__).resolve().parent
# Isola o programa das bibliotecas instaladas no Python da máquina (site-packages).
# Usa só a biblioteca padrão + a pasta lib/. Evita conflito com versões instaladas e com
# pacotes compilados (numpy, lxml...) que o Controle de Aplicativo Inteligente pode bloquear.
sys.path[:] = [p for p in sys.path if "site-packages" not in p.lower() and "dist-packages" not in p.lower()]
sys.path[:0] = [str(BASE), str(BASE / "lib")]
LOG = BASE / "instalar.log"
VERSAO_MINIMA = (3, 11)
EXTENSOES_BINARIAS = {".exe", ".dll", ".pyd", ".so"}
NOME_ATALHO = "Envio de Contracheques.lnk"

AZUL, FUNDO, BRANCO, CINZA = "#1F4E79", "#F4F6F9", "#FFFFFF", "#5B6573"


class FalhaInstalacao(Exception):
    pass


def registrar(msg: str) -> None:
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


def pythonw() -> Path:
    exe = Path(sys.executable)
    candidato = exe.with_name("pythonw.exe") if exe.name.lower() != "pythonw.exe" else exe
    return candidato if candidato.exists() else exe


# ---------------------------------------------------------------- etapas
def conferir_python() -> str:
    v = sys.version_info
    if (v.major, v.minor) < VERSAO_MINIMA:
        raise FalhaInstalacao(f"É preciso Python {VERSAO_MINIMA[0]}.{VERSAO_MINIMA[1]} ou superior (python.org).")
    aviso = " (Microsoft Store — prefira o python.org)" if "WindowsApps" in sys.executable else ""
    return f"Python {v.major}.{v.minor}.{v.micro}{aviso}"


def conferir_pasta() -> str:
    if "\\Temp\\" in str(BASE) and ".zip" in str(BASE).lower():
        raise FalhaInstalacao("Está rodando de dentro do .zip. Use 'Extrair tudo' e rode de novo.")
    obrigatorios = ("app.py", "nucleo.py", "aes_puro.py", "cofre.py", "identidade.py", "config.ini",
                    "Executar.pyw", "lib/pypdf", "lib/openpyxl", "lib/et_xmlfile", "lib/comtypes")
    faltando = [o for o in obrigatorios if not (BASE / o).exists()]
    if faltando:
        raise FalhaInstalacao(
            "Faltam arquivos: " + ", ".join(faltando) + ". Baixe o envio_contracheques.zip "
            "completo, clique com o botão direito > Extrair tudo e rode o Instalar.pyw de dentro "
            "da pasta extraída.")
    teste = BASE / ".teste_escrita"
    try:
        teste.write_text("ok", encoding="utf-8")
        teste.unlink()
    except OSError:
        raise FalhaInstalacao(f"Sem permissão de escrita em {BASE}. Use C:\\RH\\EnvioContracheques.")
    return str(BASE)


def desbloquear_arquivos() -> str:
    """Remove a marca 'baixado da internet' (Zone.Identifier) só dos arquivos desta pasta."""
    if os.name != "nt":
        return "não é Windows — ignorado"
    total = 0
    for raiz, _, arquivos in os.walk(BASE):
        for nome in arquivos:
            try:
                os.remove(os.path.join(raiz, nome) + ":Zone.Identifier")
                total += 1
            except OSError:
                pass
    return f"{total} arquivo(s) desbloqueado(s)"


def conferir_lib() -> str:
    binarios = [p for p in (BASE / "lib").rglob("*") if p.suffix.lower() in EXTENSOES_BINARIAS]
    if binarios:
        for p in binarios:
            registrar(f"Binário inesperado: {p}")
        raise FalhaInstalacao("A pasta lib deveria ter só código Python. Recopie a pasta original.")
    return "somente código Python"


def carregar_modulos() -> None:
    """
    Importa tudo na thread da tela, ANTES das etapas: importação de pacote em thread
    secundária pode gerar 'partially initialized module'. Erro aqui fica registrado no log.
    """
    if os.name == "nt":
        import comtypes  # noqa: F401
    import nucleo  # noqa: F401
    import openpyxl  # noqa: F401
    import pypdf  # noqa: F401


def autoteste() -> str:
    import nucleo
    from pypdf import PdfReader, PdfWriter

    with tempfile.TemporaryDirectory() as tmp:
        origem, destino = Path(tmp) / "origem.pdf", Path(tmp) / "protegido.pdf"
        w = PdfWriter()
        w.add_blank_page(width=595, height=842)
        with open(origem, "wb") as f:
            w.write(f)
        nucleo.proteger_pdf(origem, destino, "1234", "Autoteste")
        if not PdfReader(str(destino)).decrypt("1234") or PdfReader(str(destino)).decrypt("9999"):
            raise FalhaInstalacao("A proteção do PDF não passou na verificação.")
    nucleo.carregar_config(BASE / "config.ini", BASE / nucleo.ARQUIVO_CONFIG_USUARIO)
    return f"PDF protegido com AES-256 ({nucleo.PROVEDOR_CRIPTO})"


def testar_outlook() -> str:
    if os.name != "nt":
        return "não é Windows — ignorado"
    import nucleo

    try:
        contas = nucleo.listar_contas_outlook()
    except Exception as exc:
        registrar(traceback.format_exc())
        return f"não conectou ({exc.__class__.__name__}). Se for usar Outlook: clássico, aberto e logado."
    return f"{len(contas)} conta(s): " + ", ".join(contas) if contas else "conectado, sem contas"


ICONE = BASE / "icone.ico"


_RESULTADO_ICONE: list = []   # preenchido na thread da tela (o Tk não pode ser usado em outra thread)


def etapa_icone() -> str:
    resultado = _RESULTADO_ICONE[0] if _RESULTADO_ICONE else "não executado"
    if isinstance(resultado, Exception):
        raise resultado
    return resultado


def gerar_icone() -> str:
    """icone.png (símbolo quadrado, se existir) tem prioridade; senão usa a logo.png."""
    import identidade

    for origem in (BASE / "icone.png", BASE / "logo.png"):
        if origem.is_file():
            if identidade.gerar_icone(origem, ICONE):
                return f"gerado a partir de {origem.name}"
            raise FalhaInstalacao(f"Não foi possível ler {origem.name}. Salve como PNG e tente de novo.")
    ICONE.unlink(missing_ok=True)
    return "logo.png não encontrada: atalho fica com o ícone padrão"


def criar_atalhos() -> str:
    if os.name != "nt":
        return "não é Windows — ignorado"
    import comtypes
    import comtypes.client

    alvo = pythonw()
    comtypes.CoInitialize()
    criados = []
    try:
        shell = comtypes.client.CreateObject("WScript.Shell", dynamic=True)
        # Área de Trabalho, Menu Iniciar e a própria pasta do programa (no lugar de clicar no Executar.pyw)
        destinos = [("Área de Trabalho", Path(str(shell.SpecialFolders("Desktop")))),
                    ("Menu Iniciar", Path(str(shell.SpecialFolders("Programs")))),
                    ("pasta do programa", BASE)]
        for rotulo, pasta in destinos:
            atalho = shell.CreateShortcut(str(pasta / NOME_ATALHO))
            atalho.TargetPath = str(alvo)
            atalho.Arguments = f'"{BASE / "Executar.pyw"}"'
            atalho.WorkingDirectory = str(BASE)
            atalho.Description = "Envio de contracheques protegidos por senha"
            if ICONE.is_file():
                atalho.IconLocation = f"{ICONE},0"
            atalho.Save()
            criados.append(rotulo)
            del atalho
        del shell
    finally:
        gc.collect()
        comtypes.CoUninitialize()
    return "criado em " + ", ".join(criados) + (" com a logo" if ICONE.is_file() else "")


ETAPAS = [
    ("Python", conferir_python, True),
    ("Pasta do programa", conferir_pasta, True),
    ("Desbloquear arquivos desta pasta", desbloquear_arquivos, True),
    ("Bibliotecas", conferir_lib, True),
    ("Autoteste de proteção de PDF", autoteste, True),
    ("Conexão com o Outlook (opcional)", testar_outlook, False),
    ("Ícone com a logo", etapa_icone, False),
    ("Atalhos", criar_atalhos, True),
]


# ---------------------------------------------------------------- tela
class Instalador(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Instalar o envio de contracheques")
        self.geometry("720x540")
        self.resizable(False, False)
        self.configure(bg=FUNDO)
        self.fila: queue.Queue = queue.Queue()
        self.var_outlook = tk.BooleanVar(value=True)

        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", font=("Segoe UI", 10), background=FUNDO)
        s.configure("Card.TFrame", background=BRANCO)
        s.configure("Card.TLabel", background=BRANCO)
        s.configure("Dica.TLabel", background=BRANCO, foreground=CINZA, font=("Segoe UI", 9))
        s.configure("Card.TCheckbutton", background=BRANCO)
        s.configure("Primario.TButton", background=AZUL, foreground=BRANCO, padding=(18, 8),
                    font=("Segoe UI", 10, "bold"), borderwidth=0)
        s.map("Primario.TButton", background=[("disabled", "#A9B7C6"), ("active", "#2E75B6")])

        # Mesma identidade do programa: logo à esquerda, faixa na cor da marca
        cor = AZUL
        self.logo = None
        try:
            import identidade
            self.logo = identidade.carregar_logo(BASE / "logo.png", altura_max=44)
            cores = identidade.cores_da_logo(BASE / "logo.png")
            if cores:
                cor = cores[0]
            if self.logo is not None:
                self.iconphoto(True, self.logo)
        except Exception:
            registrar(traceback.format_exc())
        s.configure("Primario.TButton", background=cor)
        topo = tk.Frame(self, bg=BRANCO)
        topo.pack(fill="x")
        if self.logo is not None:
            tk.Label(topo, image=self.logo, bg=BRANCO).pack(side="left", padx=(20, 0), pady=14)
            tk.Frame(topo, bg="#DDE2E8", width=1).pack(side="left", fill="y", padx=16, pady=14)
        textos = tk.Frame(topo, bg=BRANCO)
        textos.pack(side="left", padx=(0 if self.logo else 20, 0), pady=12)
        tk.Label(textos, text="Instalar o envio de contracheques", bg=BRANCO, fg="#1C2430",
                 font=("Segoe UI Semibold", 14)).pack(anchor="w")
        tk.Label(textos, text="Sem administrador e sem mudar a segurança do Windows.",
                 bg=BRANCO, fg=CINZA, font=("Segoe UI", 9)).pack(anchor="w")
        tk.Frame(self, bg=cor, height=3).pack(fill="x")

        card = ttk.Frame(self, style="Card.TFrame", padding=18)
        card.pack(fill="both", expand=True, padx=16, pady=16)
        self.linhas = []
        for i, (titulo, _, _) in enumerate(ETAPAS):
            icone = tk.Label(card, text="○", bg=BRANCO, fg=CINZA, font=("Segoe UI", 12), width=2)
            icone.grid(row=i, column=0, sticky="w", pady=3)
            ttk.Label(card, text=titulo, style="Card.TLabel", width=32).grid(row=i, column=1, sticky="w")
            detalhe = ttk.Label(card, text="", style="Dica.TLabel", wraplength=300)
            detalhe.grid(row=i, column=2, sticky="w")
            self.linhas.append((icone, detalhe))
        ttk.Checkbutton(card, text="Testar a conexão com o Outlook (deixe-o aberto)", variable=self.var_outlook,
                        style="Card.TCheckbutton").grid(row=len(ETAPAS), column=0, columnspan=3,
                                                        sticky="w", pady=(16, 0))
        self.var_msg = tk.StringVar(value="Clique em Instalar.")
        ttk.Label(card, textvariable=self.var_msg, style="Card.TLabel", wraplength=600,
                  font=("Segoe UI", 10, "bold")).grid(row=len(ETAPAS) + 1, column=0, columnspan=3,
                                                       sticky="w", pady=(14, 0))

        botoes = ttk.Frame(self, padding=(16, 0, 16, 16))
        botoes.pack(fill="x")
        self.btn_instalar = ttk.Button(botoes, text="Instalar", style="Primario.TButton", command=self._iniciar)
        self.btn_instalar.pack(side="left")
        self.btn_abrir = ttk.Button(botoes, text="Abrir o programa", command=self._abrir_programa, state="disabled")
        self.btn_abrir.pack(side="left", padx=8)
        ttk.Button(botoes, text="Fechar", command=self.destroy).pack(side="right")
        self.after(100, self._processar_fila)

    def _iniciar(self):
        self.btn_instalar.configure(state="disabled")
        try:
            carregar_modulos()
        except Exception as exc:
            registrar(f"[ERRO] Carregar bibliotecas: {exc}\n{traceback.format_exc()}")
            registrar("sys.path = " + " | ".join(sys.path))
            self.var_msg.set(f"Falha ao carregar as bibliotecas da pasta lib: {exc}. Detalhes em {LOG.name}.")
            self.btn_instalar.configure(state="normal")
            return
        _RESULTADO_ICONE.clear()
        try:
            _RESULTADO_ICONE.append(gerar_icone())
        except Exception as exc:
            _RESULTADO_ICONE.append(exc)
        self.var_msg.set("Instalando…")
        testar = self.var_outlook.get()
        threading.Thread(target=self._executar, args=(testar,), daemon=True).start()

    def _executar(self, testar_outlook_flag: bool):
        registrar(f"\n==== {datetime.now():%Y-%m-%d %H:%M:%S} ====")
        for i, (titulo, funcao, obrigatoria) in enumerate(ETAPAS):
            if funcao is testar_outlook and not testar_outlook_flag:
                self.fila.put(("etapa", i, "pulado", "não testado"))
                continue
            self.fila.put(("etapa", i, "rodando", ""))
            try:
                resultado = funcao()
                registrar(f"[OK] {titulo}: {resultado}")
                self.fila.put(("etapa", i, "ok", resultado))
            except Exception as exc:
                registrar(f"[ERRO] {titulo}: {exc}\n{traceback.format_exc()}")
                if obrigatoria:
                    self.fila.put(("etapa", i, "erro", str(exc)))
                    self.fila.put(("fim", False, f"Falhou em '{titulo}'. Detalhes em {LOG.name}."))
                    return
                self.fila.put(("etapa", i, "aviso", str(exc)))
        self.fila.put(("fim", True, "Pronto! Use o atalho 'Envio de Contracheques' na Área de Trabalho."))

    def _processar_fila(self):
        icones = {"rodando": ("◐", AZUL), "ok": ("✓", "#2E7D32"), "erro": ("✗", "#C62828"),
                  "aviso": ("!", "#B26A00"), "pulado": ("–", CINZA)}
        try:
            while True:
                msg = self.fila.get_nowait()
                if msg[0] == "etapa":
                    _, i, estado, texto = msg
                    simbolo, cor = icones[estado]
                    self.linhas[i][0].configure(text=simbolo, fg=cor)
                    self.linhas[i][1].configure(text=texto)
                elif msg[0] == "fim":
                    _, sucesso, texto = msg
                    self.var_msg.set(texto)
                    if sucesso:
                        self.btn_abrir.configure(state="normal")
                    else:
                        self.btn_instalar.configure(state="normal")
                        messagebox.showerror("Instalação", texto)
        except queue.Empty:
            pass
        self.after(100, self._processar_fila)

    def _abrir_programa(self):
        subprocess.Popen([str(pythonw()), str(BASE / "Executar.pyw")], cwd=str(BASE))
        self.destroy()


if __name__ == "__main__":
    try:
        Instalador().mainloop()
    except Exception:
        registrar(traceback.format_exc())
        raise
