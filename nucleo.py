"""
Núcleo do Envio de Contracheques.

Responsável por:
  - ler a planilha de colaboradores (CODIGO, NOME, EMAIL, CPF_FINAL, ...)
  - varrer a pasta de PDFs e casar cada arquivo com um colaborador pelo código
  - conferir se o nome do colaborador aparece dentro do PDF (trava contra troca de arquivo)
  - proteger cada PDF com senha (AES-256) e enviar por e-mail (Outlook, SMTP ou simulado)
  - registrar histórico para nunca reenviar o mesmo arquivo em duplicidade

Nenhuma senha é gravada em log ou histórico.
"""

from __future__ import annotations

import configparser
import csv
from collections import Counter
import hashlib
import html
import logging
import os
import re
import secrets
import shutil
import smtplib
import ssl
import sys
import tempfile
import threading
import time
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from email.message import EmailMessage
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Callable, Optional

from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter

import aes_puro
import cofre

log = logging.getLogger("contracheques")

# Sem o pacote "cryptography" (DLL sem assinatura), o AES do pypdf vem do aes_puro.py
PROVEDOR_CRIPTO = aes_puro.instalar_no_pypdf()


# ---------------------------------------------------------------------------
# Situações possíveis de cada item do lote
# ---------------------------------------------------------------------------
class Situacao:
    PRONTO = "Pronto"            # vai ser enviado
    REVISAR = "Revisar"          # precisa de liberação manual (duplo clique)
    PENDENTE = "Sem PDF"         # colaborador ativo sem arquivo na pasta
    ERRO = "Erro"                # bloqueado, não envia
    JA_ENVIADO = "Já enviado"    # mesmo arquivo já enviado antes
    ENVIADO = "Enviado"
    TESTE = "Enviado (teste)"
    SIMULADO = "Simulado"


COLUNAS_OBRIGATORIAS = ["CODIGO", "NOME", "EMAIL", "CPF_FINAL"]

# Apelidos aceitos no cabeçalho da planilha (depois de normalizado)
APELIDOS_CABECALHO = {
    "E_MAIL": "EMAIL",
    "EMAIL_COLABORADOR": "EMAIL",
    "CPF": "CPF_FINAL",
    "FINAL_CPF": "CPF_FINAL",
    "MATRICULA": "CODIGO",
    "COD": "CODIGO",
    "CODIGO_FOLHA": "CODIGO",
    "DATA_NASCIMENTO": "DATA_NASC",
    "NASCIMENTO": "DATA_NASC",
}

RE_EMAIL = re.compile(r"^[^@\s;,]+@[^@\s;,]+\.[^@\s;,]{2,}$")
RE_COMPETENCIA = re.compile(r"^(0[1-9]|1[0-2])/(\d{4})$")


class ErroPlanilha(Exception):
    """Problema estrutural na planilha (arquivo, aba ou colunas)."""


# ---------------------------------------------------------------------------
# Utilitários
# ---------------------------------------------------------------------------
def caminho_base() -> Path:
    """Pasta do executável (PyInstaller) ou do script."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def configurar_logging(base: Path) -> None:
    pasta_logs = base / "logs"
    pasta_logs.mkdir(parents=True, exist_ok=True)
    if any(isinstance(h, RotatingFileHandler) for h in log.handlers):
        return
    log.setLevel(logging.INFO)
    arquivo = RotatingFileHandler(
        pasta_logs / "envio.log", maxBytes=2_000_000, backupCount=10, encoding="utf-8"
    )
    arquivo.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
    log.addHandler(arquivo)


def normalizar_texto(texto: str) -> str:
    """Maiúsculo, sem acento, só letras/números e espaço simples."""
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^A-Z0-9 ]", " ", texto.upper())
    return re.sub(r"\s+", " ", texto).strip()


def normalizar_codigo(valor) -> str:
    """'00123', 123 e 123.0 viram '123' — o Excel costuma comer zeros à esquerda."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    texto = str(valor).strip()
    if texto.isdigit():
        texto = str(int(texto))
    return texto.upper()


def extrair_cpf4(valor) -> Optional[str]:
    """
    Aceita só os 4 dígitos (recomendado) ou o CPF completo.
    Célula numérica perde zero à esquerda (0123 -> 123), por isso o zfill.
    """
    if valor is None or str(valor).strip() == "":
        return None
    if isinstance(valor, (int, float)):
        digitos = str(int(valor))
        return digitos.zfill(4)[-4:] if len(digitos) <= 4 else digitos[-4:]
    digitos = re.sub(r"\D", "", str(valor))
    return digitos[-4:] if len(digitos) >= 4 else None


def extrair_ddmm(valor) -> Optional[str]:
    if valor is None or str(valor).strip() == "":
        return None
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%d%m")
    m = re.match(r"^\s*(\d{1,2})[/\-.](\d{1,2})[/\-.]\d{2,4}\s*$", str(valor))
    if not m:
        return None
    return f"{int(m.group(1)):02d}{int(m.group(2)):02d}"


def sha256_arquivo(caminho: Path) -> str:
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(65536), b""):
            h.update(bloco)
    return h.hexdigest()


def competencia_padrao() -> str:
    """Mês anterior — contracheque normalmente sai no início do mês seguinte."""
    hoje = date.today()
    mes, ano = (12, hoje.year - 1) if hoje.month == 1 else (hoje.month - 1, hoje.year)
    return f"{mes:02d}/{ano}"


def competencia_valida(texto: str) -> bool:
    return bool(RE_COMPETENCIA.match(texto.strip()))


def competencia_iso(texto: str) -> str:
    m = RE_COMPETENCIA.match(texto.strip())
    return f"{m.group(2)}-{m.group(1)}" if m else "sem-competencia"


# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------
@dataclass
class Config:
    identificar_por: str = "nome"                       # nome | codigo
    padrao_nome_arquivo: str = r"^(.+?)_"               # 1º grupo = nome do colaborador
    padrao_codigo: str = r"^(\d+)"                      # usado se identificar_por = codigo
    padrao_competencia_arquivo: str = r"(\d{2})[.\-/ ](\d{4})"
    filtro_arquivo: str = ""                            # regex; vazio = todos os PDFs
    conferir_nome: bool = True                          # conferência do conteúdo do PDF
    conferir_codigo_pdf: bool = True                    # só no modo genérico (sem regex_cabecalho)
    # Cabeçalho do recibo: grupo 1 = código, grupo 2 = nome (layout da folha atual)
    regex_cabecalho_pdf: str = r"(\d+)\s*\n\s*C[óo]digo\s*\n\s*([^\n]+?)\s*\n\s*Nome do Funcion"
    conferir_competencia_pdf: bool = True
    formato_senha: str = "{cpf4}"
    pausa_segundos: float = 2.0
    mover_enviados: bool = True
    modo_envio: str = "simulado"         # outlook | smtp | simulado
    remetente: str = ""                  # Outlook: conta/caixa que envia
    assunto: str = "Contracheque {tipo_folha} {competencia} - {nome}"
    nome_empresa: str = "Empresa Demo"
    # Identidade visual (tela e e-mail) — seção [tema] do config.ini
    cor_primaria: str = "#0E3B5C"
    cor_destaque: str = "#1FA2B8"
    cor_fundo: str = "#F2F4F7"
    logo: str = "logo.png"               # PNG ou GIF na pasta do programa
    cores_da_logo: bool = True           # extrai a cor principal da própria logo
    # SMTP (a senha fica no cofre do Windows, nunca aqui)
    smtp_host: str = ""
    smtp_porta: int = 587
    smtp_seguranca: str = "starttls"     # starttls | ssl
    smtp_usuario: str = ""
    smtp_remetente: str = ""
    # Último uso / padrões da tela
    pasta_pdfs: str = ""
    planilha: str = ""
    email_teste: str = ""


ARQUIVO_CONFIG_USUARIO = "configuracoes.ini"   # gravado pela tela de Configurações


def carregar_config(caminho: Path, caminho_usuario: Optional[Path] = None) -> Config:
    """
    config.ini = regras avançadas (padrões de arquivo, regex, layout do recibo).
    configuracoes.ini = o que o RH ajusta pela tela; sobrepõe o config.ini.
    """
    cfg = Config()
    p = configparser.ConfigParser(interpolation=None)
    arquivos = [str(c) for c in (caminho, caminho_usuario) if c is not None and Path(c).exists()]
    if not arquivos:
        log.warning("Nenhum arquivo de configuração encontrado — usando padrões.")
    p.read(arquivos, encoding="utf-8")

    def secao(nome):
        return p[nome] if p.has_section(nome) else {}

    g, e, sm, u, t = secao("geral"), secao("email"), secao("smtp"), secao("ultimo_uso"), secao("tema")
    sim = ("sim", "s", "true", "1")
    cfg.identificar_por = str(g.get("identificar_por", cfg.identificar_por)).strip().lower()
    cfg.padrao_nome_arquivo = g.get("padrao_nome_arquivo", cfg.padrao_nome_arquivo)
    cfg.padrao_codigo = g.get("padrao_codigo_arquivo", cfg.padrao_codigo)
    cfg.padrao_competencia_arquivo = g.get("padrao_competencia_arquivo", cfg.padrao_competencia_arquivo)
    cfg.filtro_arquivo = str(g.get("filtro_arquivo", "")).strip()
    cfg.conferir_nome = str(g.get("conferir_conteudo_pdf", "sim")).strip().lower() in sim
    cfg.conferir_codigo_pdf = str(g.get("conferir_codigo_no_pdf", "sim")).strip().lower() in sim
    cfg.regex_cabecalho_pdf = str(g.get("regex_cabecalho_pdf", cfg.regex_cabecalho_pdf)).strip()
    cfg.conferir_competencia_pdf = str(g.get("conferir_competencia_pdf", "sim")).strip().lower() in sim
    cfg.formato_senha = g.get("formato_senha", cfg.formato_senha).strip()
    cfg.pausa_segundos = float(g.get("pausa_segundos", cfg.pausa_segundos))
    cfg.mover_enviados = str(g.get("mover_enviados", "sim")).strip().lower() in sim
    cfg.modo_envio = str(e.get("modo_envio", cfg.modo_envio)).strip().lower()
    cfg.remetente = str(e.get("remetente", "")).strip()
    cfg.assunto = e.get("assunto", cfg.assunto)
    nome = str(e.get("nome_empresa", "")).strip()
    # Valores provisórios de versões anteriores não sobrescrevem o nome real
    if nome and normalizar_texto(nome) not in ("SUA EMPRESA", "EMPRESA"):
        cfg.nome_empresa = nome
    cfg.smtp_host = str(sm.get("host", "")).strip()
    cfg.smtp_porta = int(str(sm.get("porta", cfg.smtp_porta)).strip() or cfg.smtp_porta)
    cfg.smtp_seguranca = str(sm.get("seguranca", cfg.smtp_seguranca)).strip().lower()
    cfg.smtp_usuario = str(sm.get("usuario", "")).strip()
    cfg.smtp_remetente = str(sm.get("remetente", "")).strip()
    cfg.logo = str(t.get("logo", cfg.logo)).strip()
    cfg.cores_da_logo = str(t.get("cores_da_logo", "sim")).strip().lower() in sim
    for chave in ("cor_primaria", "cor_destaque", "cor_fundo"):
        valor = str(t.get(chave, getattr(cfg, chave))).strip()
        if re.fullmatch(r"#[0-9A-Fa-f]{6}", valor):
            setattr(cfg, chave, valor.upper())
    cfg.pasta_pdfs = str(u.get("pasta_pdfs", "")).strip()
    cfg.planilha = str(u.get("planilha", "")).strip()
    cfg.email_teste = str(u.get("email_teste", "")).strip()

    # Falha cedo se alguma regex estiver errada
    for regex in (cfg.padrao_nome_arquivo, cfg.padrao_codigo, cfg.padrao_competencia_arquivo,
                  cfg.filtro_arquivo, cfg.regex_cabecalho_pdf):
        re.compile(regex)
    if cfg.identificar_por not in ("nome", "codigo"):
        raise ValueError("identificar_por deve ser nome ou codigo.")
    if "{cpf4}" not in cfg.formato_senha and "{nasc_ddmm}" not in cfg.formato_senha:
        raise ValueError("formato_senha precisa conter {cpf4} e/ou {nasc_ddmm}.")
    if cfg.modo_envio not in ("outlook", "smtp", "simulado"):
        raise ValueError("modo_envio deve ser outlook, smtp ou simulado.")
    if cfg.smtp_seguranca not in ("starttls", "ssl"):
        raise ValueError("Segurança SMTP deve ser starttls ou ssl.")
    return cfg


def salvar_configuracoes(caminho: Path, valores: dict[str, dict[str, object]]) -> None:
    """Atualiza só as chaves informadas em configuracoes.ini (preserva o resto)."""
    p = configparser.ConfigParser(interpolation=None)
    if caminho.exists():
        p.read(caminho, encoding="utf-8")
    for nome_secao, chaves in valores.items():
        if not p.has_section(nome_secao):
            p.add_section(nome_secao)
        for chave, valor in chaves.items():
            if isinstance(valor, bool):
                valor = "sim" if valor else "nao"
            p[nome_secao][chave] = str(valor if valor is not None else "")
    temporario = caminho.with_suffix(".tmp")
    with open(temporario, "w", encoding="utf-8") as f:
        p.write(f)
    os.replace(temporario, caminho)  # gravação atômica: nunca deixa o arquivo pela metade


# ---------------------------------------------------------------------------
# Modelos de dados
# ---------------------------------------------------------------------------
@dataclass
class Colaborador:
    codigo: str
    nome: str
    email: str
    cpf4: Optional[str]
    nasc_ddmm: Optional[str]
    ativo: bool
    linha: int


@dataclass
class ItemEnvio:
    codigo: str
    nome: str
    email: str
    arquivo: Optional[Path]
    situacao: str
    detalhe: str = ""
    colaborador: Optional[Colaborador] = field(default=None, repr=False)
    sha256: str = ""
    competencia: str = ""      # MM/AAAA lida do recibo (ou do nome do arquivo)
    tipo_folha: str = ""       # Folha Mensal, Férias, 13º Salário...


# ---------------------------------------------------------------------------
# Leitura da planilha
# ---------------------------------------------------------------------------
def _chave_cabecalho(valor) -> str:
    chave = normalizar_texto(valor).replace(" ", "_")
    return APELIDOS_CABECALHO.get(chave, chave)


def ler_planilha(caminho: Path) -> tuple[dict[str, Colaborador], list[ItemEnvio]]:
    """Retorna (colaboradores válidos por código, itens de erro encontrados na planilha)."""
    if not caminho.exists():
        raise ErroPlanilha(f"Planilha não encontrada: {caminho}")
    try:
        wb = load_workbook(caminho, read_only=True, data_only=True)
    except PermissionError as exc:
        raise ErroPlanilha("A planilha está aberta no Excel. Feche e tente de novo.") from exc

    try:
        ws = wb["Colaboradores"] if "Colaboradores" in wb.sheetnames else wb.worksheets[0]
        linhas = ws.iter_rows(values_only=True)
        cabecalho = next(linhas, None)
        if not cabecalho:
            raise ErroPlanilha("A planilha está vazia.")

        idx = {_chave_cabecalho(c): i for i, c in enumerate(cabecalho) if c is not None}
        faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in idx]
        if faltando:
            raise ErroPlanilha("Colunas obrigatórias ausentes: " + ", ".join(faltando))

        brutos: list[Colaborador] = []
        erros: list[ItemEnvio] = []

        for n, row in enumerate(linhas, start=2):
            if row is None or all(v in (None, "") for v in row):
                continue

            def val(coluna):
                i = idx.get(coluna)
                return row[i] if i is not None and i < len(row) else None

            codigo = normalizar_codigo(val("CODIGO"))
            nome = re.sub(r"\s+", " ", str(val("NOME") or "")).strip()
            email = str(val("EMAIL") or "").strip().lower()
            ativo_txt = normalizar_texto(val("ATIVO") or "S")
            ativo = not ativo_txt.startswith("N")

            if not codigo:
                erros.append(ItemEnvio("", nome or "-", email, None, Situacao.ERRO,
                                       f"Linha {n}: CODIGO vazio"))
                continue

            brutos.append(Colaborador(
                codigo=codigo,
                nome=nome,
                email=email,
                cpf4=extrair_cpf4(val("CPF_FINAL")),
                nasc_ddmm=extrair_ddmm(val("DATA_NASC")),
                ativo=ativo,
                linha=n,
            ))
    finally:
        wb.close()

    # Código duplicado bloqueia todas as linhas daquele código (não dá para saber qual é o certo)
    contagem: dict[str, int] = {}
    for c in brutos:
        contagem[c.codigo] = contagem.get(c.codigo, 0) + 1

    colaboradores: dict[str, Colaborador] = {}
    for c in brutos:
        if contagem[c.codigo] > 1:
            erros.append(ItemEnvio(c.codigo, c.nome, c.email, None, Situacao.ERRO,
                                   f"Linha {c.linha}: código duplicado na planilha"))
        else:
            colaboradores[c.codigo] = c

    log.info("Planilha lida: %d colaboradores válidos, %d problemas.", len(colaboradores), len(erros))
    return colaboradores, erros


# ---------------------------------------------------------------------------
# PDFs
# ---------------------------------------------------------------------------
def chave_do_arquivo(arq: Path, cfg: Config) -> Optional[str]:
    """Extrai do nome do arquivo a chave de identificação (nome normalizado ou código)."""
    if cfg.identificar_por == "nome":
        m = re.search(cfg.padrao_nome_arquivo, unicodedata.normalize("NFC", arq.stem), re.IGNORECASE)
        if not m:
            return None
        chave = normalizar_texto(m.group(1) if m.groups() else m.group(0))
        return chave or None
    m = re.search(cfg.padrao_codigo, arq.stem)
    if not m:
        return None
    return normalizar_codigo(m.group(1) if m.groups() else m.group(0)) or None


def chave_do_colaborador(colab: "Colaborador", cfg: Config) -> str:
    return normalizar_texto(colab.nome) if cfg.identificar_por == "nome" else colab.codigo


def competencia_do_arquivo(arq: Path, cfg: Config) -> Optional[str]:
    """'Fulano_Recibo de Salário 09.2026' -> '09/2026'. None se não achar."""
    m = re.search(cfg.padrao_competencia_arquivo, arq.stem)
    if not m or len(m.groups()) < 2:
        return None
    return f"{int(m.group(1)):02d}/{m.group(2)}"


def varrer_pdfs(pasta: Path, cfg: Config) -> tuple[dict[str, list[Path]], list[Path], int]:
    """
    Retorna ({chave: [pdfs]}, [pdfs fora do padrão], qtd ignorados pelo filtro).
    Não entra em subpastas (a pasta Enviados fica de fora).
    """
    if not pasta.is_dir():
        raise FileNotFoundError(f"Pasta de PDFs não encontrada: {pasta}")
    filtro = re.compile(cfg.filtro_arquivo, re.IGNORECASE) if cfg.filtro_arquivo else None
    por_chave: dict[str, list[Path]] = {}
    fora_padrao: list[Path] = []
    ignorados = 0
    for arq in sorted(pasta.iterdir()):
        if not arq.is_file() or arq.suffix.lower() != ".pdf":
            continue
        if filtro and not filtro.search(unicodedata.normalize("NFC", arq.stem)):
            ignorados += 1
            continue
        chave = chave_do_arquivo(arq, cfg)
        if chave is None:
            fora_padrao.append(arq)
        else:
            por_chave.setdefault(chave, []).append(arq)
    if ignorados:
        log.info("%d PDF(s) ignorados pelo filtro_arquivo.", ignorados)
    return por_chave, fora_padrao, ignorados


MESES = {"JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6, "JULHO": 7,
         "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12}
RE_MES_ANO = re.compile(
    r"\b(janeiro|fevereiro|mar[çc]o|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)"
    r"\s+de\s+(\d{4})\b", re.IGNORECASE)


def cabecalhos_do_pdf(texto: str, regex: str) -> set[tuple[str, str]]:
    """
    Pares (código, nome normalizado) encontrados no cabeçalho do recibo.
    O recibo tem 2 vias; as duas precisam trazer o mesmo par.
    """
    return {(normalizar_codigo(c), normalizar_texto(n))
            for c, n in re.findall(regex, texto, re.IGNORECASE)}


def competencias_do_pdf(texto: str) -> set[str]:
    """'Setembro de 2026' -> {'09/2026'}."""
    return {f"{MESES[normalizar_texto(m)]:02d}/{a}" for m, a in RE_MES_ANO.findall(texto)}


NOMES_MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
               "setembro", "outubro", "novembro", "dezembro"]
# Linha imediatamente acima de "Setembro de 2026" = tipo da folha ("Folha Mensal", "Férias"...)
RE_TIPO_FOLHA = re.compile(
    r"([^\n]+?)[ \t]*\n[ \t]*(?:janeiro|fevereiro|mar[çc]o|abril|maio|junho|julho|agosto|setembro|"
    r"outubro|novembro|dezembro)\s+de\s+\d{4}", re.IGNORECASE)


def mes_extenso(competencia: str) -> str:
    """'09/2026' -> 'setembro de 2026'."""
    m = RE_COMPETENCIA.match(competencia or "")
    return f"{NOMES_MESES[int(m.group(1)) - 1]} de {m.group(2)}" if m else (competencia or "")


def tipo_amigavel(bruto: str) -> str:
    """Padroniza o tipo de folha impresso no recibo para o texto do e-mail."""
    n = normalizar_texto(bruto)
    if not n:
        return ""
    if re.search(r"\b13", n) or "DECIMO TERCEIRO" in n:
        if "ADIANT" in n or "1A PARC" in n or "PRIMEIRA" in n:
            return "13º Salário (1ª parcela)"
        if "2A PARC" in n or "SEGUNDA" in n:
            return "13º Salário (2ª parcela)"
        return "13º Salário"
    if "FERIAS" in n:
        return "Férias"
    if "ADIANT" in n:
        return "Adiantamento Salarial"
    if "COMPLEMENT" in n:
        return "Folha Complementar"
    if "RESCIS" in n:
        return "Rescisão"
    if "PLR" in n.split() or "PARTICIPACAO" in n:
        return "PLR"
    if "MENSAL" in n or "PAGAMENTO" in n:
        return "Folha Mensal"
    return re.sub(r"\s+", " ", bruto).strip()


def rotulo_documento(tipo: str) -> str:
    """Parte do nome do arquivo: 'Recibo de Salário', 'Recibo de Férias'..."""
    if not tipo or tipo == "Folha Mensal":
        return "Recibo de Salário"
    return f"Recibo de {tipo}"


def nome_proprio(nome: str) -> str:
    """'MARIA DA SILVA SOUZA' -> 'Maria da Silva Souza'."""
    minusculas = {"da", "de", "do", "das", "dos", "e"}
    partes = re.sub(r"\s+", " ", nome).strip().lower().split(" ")
    return " ".join(p if (i and p in minusculas) else p[:1].upper() + p[1:] for i, p in enumerate(partes))


def nome_arquivo_seguro(texto: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "-", texto).strip(" .")


@dataclass
class DadosRecibo:
    pares: set[tuple[str, str]]          # (código, nome normalizado)
    nomes: dict[str, str]                # nome normalizado -> nome como impresso
    competencias: set[str]               # {"09/2026"}
    tipos: set[str]                      # {"Folha Mensal"} (já padronizado)


def ler_dados_recibo(texto: str, cfg: Config) -> DadosRecibo:
    pares, nomes = set(), {}
    if cfg.regex_cabecalho_pdf:
        for cod, nome in re.findall(cfg.regex_cabecalho_pdf, texto, re.IGNORECASE):
            chave = normalizar_texto(nome)
            pares.add((normalizar_codigo(cod), chave))
            nomes.setdefault(chave, re.sub(r"\s+", " ", nome).strip())
    tipos = {tipo_amigavel(t) for t in RE_TIPO_FOLHA.findall(texto)} - {""}
    return DadosRecibo(pares, nomes, competencias_do_pdf(texto), tipos)


@dataclass
class ResultadoSeparacao:
    pasta: Path
    arquivos: list[Path]
    problemas: list[str]
    paginas: int
    competencias: list[str]
    tipos: list[str]


def pasta_saida_padrao(origem: Path, competencias: set[str], tipos: set[str]) -> Path:
    if len(competencias) == 1 and len(tipos) <= 1:
        tipo = next(iter(tipos), "Folha")
        return origem.parent / nome_arquivo_seguro(f"Contracheques {competencia_iso(next(iter(competencias)))} {tipo}")
    return origem.parent / nome_arquivo_seguro(f"Contracheques - {origem.stem}")


def separar_pdf_unico(
    origem: Path,
    cfg: Config,
    pasta_saida: Optional[Path] = None,
    limpar_pasta: bool = False,
    progresso: Optional[Callable[[int, int], None]] = None,
) -> ResultadoSeparacao:
    """
    Separa o PDF único da contabilidade em um PDF por colaborador e renomeia:
        'Maria Silva_Recibo de Salário 09.2026.pdf'
    Regras:
      - página com 1 colaborador no cabeçalho -> desse colaborador
      - página sem cabeçalho -> continuação da página anterior
      - página com 2 colaboradores diferentes -> NÃO separa (reportada como problema)
      - mesmo colaborador em páginas não seguidas -> juntadas no mesmo arquivo
    """
    reader = PdfReader(str(origem))
    if reader.is_encrypted:
        raise ValueError("O PDF da contabilidade está protegido por senha. Peça a versão sem senha.")
    total = len(reader.pages)
    grupos: dict[tuple[str, str, str, str], dict] = {}
    problemas: list[str] = []
    ultimo = None

    for i, pagina in enumerate(reader.pages):
        dados = ler_dados_recibo(pagina.extract_text() or "", cfg)
        if len(dados.pares) > 1:
            lista = ", ".join(f"{c} {dados.nomes.get(n, n)}" for c, n in sorted(dados.pares))
            problemas.append(f"Página {i + 1}: mais de um colaborador na mesma página ({lista}) — não separada")
            ultimo = None
        elif not dados.pares:
            if ultimo is None:
                problemas.append(f"Página {i + 1}: código/nome não encontrados — não separada")
            else:
                grupos[ultimo]["paginas"].append(i)
        else:
            codigo, nome_norm = next(iter(dados.pares))
            comp = next(iter(dados.competencias)) if len(dados.competencias) == 1 else ""
            tipo = next(iter(dados.tipos)) if len(dados.tipos) == 1 else ""
            chave = (codigo, nome_norm, comp, tipo)
            if chave not in grupos:
                grupos[chave] = {"nome": dados.nomes.get(nome_norm, nome_norm), "paginas": []}
            grupos[chave]["paginas"].append(i)
            ultimo = chave
            if not comp:
                problemas.append(f"Página {i + 1} ({codigo}): competência não identificada")
        if progresso:
            progresso(i + 1, total)

    competencias = {k[2] for k in grupos if k[2]}
    tipos = {k[3] for k in grupos if k[3]}
    pasta = pasta_saida or pasta_saida_padrao(origem, competencias, tipos)
    pasta.mkdir(parents=True, exist_ok=True)
    if limpar_pasta:
        for antigo in pasta.glob("*.pdf"):  # só arquivos soltos; subpastas (Enviados) ficam intactas
            antigo.unlink()

    # Nome do arquivo por colaborador; homônimos no mesmo lote ganham o código no nome
    contagem_nomes = Counter((nome_proprio(g["nome"]), k[2], k[3]) for k, g in grupos.items())
    arquivos = []
    for (codigo, _nome_norm, comp, tipo), g in grupos.items():
        nome = nome_proprio(g["nome"])
        sufixo_comp = f"{comp[:2]}.{comp[3:]}" if comp else "sem competencia"
        base_nome = f"{nome}_{rotulo_documento(tipo)} {sufixo_comp}"
        if contagem_nomes[(nome, comp, tipo)] > 1:
            base_nome += f" ({codigo})"
        destino = pasta / f"{nome_arquivo_seguro(base_nome)}.pdf"
        writer = PdfWriter()
        for n in g["paginas"]:
            writer.add_page(reader.pages[n])
        with open(destino, "wb") as f:
            writer.write(f)
        arquivos.append(destino)

    log.info("PDF único separado: %s | %d páginas -> %d arquivos | %d problema(s) | pasta %s",
             origem.name, total, len(arquivos), len(problemas), pasta)
    return ResultadoSeparacao(pasta, arquivos, problemas, total, sorted(competencias), sorted(tipos))


def codigo_no_texto(codigo: str, texto: str) -> bool:
    """Código como número isolado (aceita zeros à esquerda: 101 casa com 000101)."""
    if not codigo:
        return False
    return re.search(rf"(?<![0-9]){re.escape(codigo).lstrip('0') or '0'}(?![0-9])",
                     re.sub(r"(?<![0-9])0+(?=[0-9])", "", texto)) is not None


def texto_do_pdf(caminho: Path) -> str:
    reader = PdfReader(str(caminho))
    if reader.is_encrypted:
        raise ValueError("PDF de origem já está protegido por senha")
    return "\n".join((p.extract_text() or "") for p in reader.pages)


def nome_confere(nome: str, texto_pdf: str) -> bool:
    """
    Confere se o colaborador aparece no PDF.
    Aceita nome completo ou, se o sistema de folha abreviar, primeiro nome + último sobrenome
    como palavras inteiras.
    """
    alvo = normalizar_texto(nome)
    texto = normalizar_texto(texto_pdf)
    if not alvo or not texto:
        return False
    if f" {alvo} " in f" {texto} ":
        return True
    partes = [p for p in alvo.split() if len(p) > 2]
    if len(partes) < 2:
        return False
    palavras = set(texto.split())
    return partes[0] in palavras and partes[-1] in palavras


def montar_senha(formato: str, colab: Colaborador) -> str:
    if "{cpf4}" in formato and not colab.cpf4:
        raise ValueError("CPF_FINAL ausente ou inválido")
    if "{nasc_ddmm}" in formato and not colab.nasc_ddmm:
        raise ValueError("DATA_NASC ausente ou inválida")
    return formato.replace("{cpf4}", colab.cpf4 or "").replace("{nasc_ddmm}", colab.nasc_ddmm or "")


def descricao_senha(formato: str) -> str:
    """Texto amigável para o corpo do e-mail (nunca a senha em si)."""
    formas = {  # (primeira parte, parte seguinte)
        "{cpf4}": ("os 4 últimos dígitos do seu CPF", "dos 4 últimos dígitos do seu CPF"),
        "{nasc_ddmm}": ("o dia e o mês do seu nascimento (DDMM)", "do dia e mês do seu nascimento (DDMM)"),
    }
    partes = []
    for n, token in enumerate(re.findall(r"\{cpf4\}|\{nasc_ddmm\}|[^{}]+", formato)):
        primeira, seguinte = formas.get(token, (f'o texto "{token}"', f'do texto "{token}"'))
        partes.append(primeira if n == 0 else seguinte)
    return ", seguidos ".join(partes)


def proteger_pdf(origem: Path, destino: Path, senha: str, titulo: str) -> None:
    """Gera cópia protegida com AES-256 e valida que a senha abre o arquivo."""
    reader = PdfReader(str(origem))
    if reader.is_encrypted:
        raise ValueError("PDF de origem já está protegido por senha")
    writer = PdfWriter()
    writer.append(reader)
    writer.add_metadata({"/Title": titulo})
    writer.encrypt(
        user_password=senha,
        owner_password=secrets.token_urlsafe(24),  # senha de dono aleatória, ninguém precisa dela
        algorithm="AES-256",
    )
    with open(destino, "wb") as f:
        writer.write(f)

    # Trava de segurança: nunca envia um arquivo que não esteja protegido ou que a senha não abra
    teste = PdfReader(str(destino))
    if not teste.is_encrypted or not teste.decrypt(senha):
        raise RuntimeError("Falha na verificação da proteção do PDF")


# ---------------------------------------------------------------------------
# Histórico (CSV com ; para abrir direto no Excel)
# ---------------------------------------------------------------------------
CAMPOS_HISTORICO = ["data_hora", "competencia", "tipo_folha", "codigo", "nome", "email_destino",
                    "arquivo", "sha256", "situacao", "detalhe"]


def carregar_historico(caminho: Path) -> tuple[set[tuple[str, str]], set[tuple[str, str, str]]]:
    """Retorna ({(codigo, sha256) enviados}, {(competencia, tipo, codigo) enviados})."""
    por_arquivo, por_competencia = set(), set()
    # Lê o histórico atual e os antigos (formato anterior, sem tipo_folha): nada enviado é esquecido
    for arquivo in sorted(caminho.parent.glob("historico_envios*.csv")):
        with open(arquivo, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f, delimiter=";"):
                if r.get("situacao") == Situacao.ENVIADO:
                    por_arquivo.add((r.get("codigo") or "", r.get("sha256") or ""))
                    por_competencia.add((r.get("competencia") or "", normalizar_texto(r.get("tipo_folha") or ""),
                                         r.get("codigo") or ""))
    return por_arquivo, por_competencia


def registrar_historico(caminho: Path, item: ItemEnvio, email_destino: str) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    if caminho.exists():
        # Histórico antigo (sem a coluna tipo_folha) é preservado com outro nome
        with open(caminho, encoding="utf-8-sig") as f:
            cabecalho = f.readline().strip().split(";")
        if cabecalho != CAMPOS_HISTORICO:
            caminho.rename(caminho.with_name(f"historico_envios_ate_{datetime.now():%Y%m%d%H%M%S}.csv"))
    novo = not caminho.exists()
    with open(caminho, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS_HISTORICO, delimiter=";")
        if novo:
            w.writeheader()
        w.writerow({
            "data_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "competencia": item.competencia,
            "tipo_folha": item.tipo_folha,
            "codigo": item.codigo,
            "nome": item.nome,
            "email_destino": email_destino,
            "arquivo": item.arquivo.name if item.arquivo else "",
            "sha256": item.sha256,
            "situacao": item.situacao,
            "detalhe": item.detalhe,
        })


# ---------------------------------------------------------------------------
# Montagem do lote (conferência — não envia nada)
# ---------------------------------------------------------------------------
def _avaliar_arquivo(
    arq: Path,
    chave: str,
    colab: Optional[Colaborador],
    cfg: Config,
    competencia_sel: str,
    conferir_conteudo: bool,
    enviados_arquivo: set,
    enviados_competencia: set,
) -> ItemEnvio:
    """Aplica todas as regras a UM PDF. Não envia nada."""
    codigo = colab.codigo if colab else ("" if cfg.identificar_por == "nome" else chave)
    item = ItemEnvio(codigo, colab.nome if colab else "-", colab.email if colab else "", arq,
                     Situacao.PRONTO, "", colab)

    def bloquear(motivo: str) -> ItemEnvio:
        item.situacao, item.detalhe = Situacao.ERRO, motivo
        return item

    # --- leitura do recibo (competência e tipo vêm daqui mesmo sem cadastro, para exibir na tela)
    comp_arquivo = competencia_do_arquivo(arq, cfg)
    try:
        texto = texto_do_pdf(arq)
    except Exception as exc:  # PDF corrompido, protegido etc.
        return bloquear(f"PDF ilegível: {exc}")
    dados = ler_dados_recibo(texto, cfg)
    item.tipo_folha = next(iter(dados.tipos)) if len(dados.tipos) == 1 else ""
    if len(dados.competencias) > 1:
        return bloquear("O recibo traz mais de uma competência: " + ", ".join(sorted(dados.competencias)))
    comp_pdf = next(iter(dados.competencias), None)
    item.competencia = comp_pdf or comp_arquivo or ""

    # --- cadastro
    if not colab:
        return bloquear("Nome do arquivo não encontrado na planilha (confira grafia)"
                        if cfg.identificar_por == "nome" else "Código não cadastrado na planilha")
    if not colab.ativo:
        return bloquear("Colaborador marcado como inativo")
    if not RE_EMAIL.match(colab.email):
        return bloquear("E-mail vazio ou inválido")
    try:
        montar_senha(cfg.formato_senha, colab)
    except ValueError as exc:
        return bloquear(str(exc))

    # --- competência
    avisos: list[str] = []
    if comp_pdf and comp_arquivo and comp_pdf != comp_arquivo:
        return bloquear(f"Nome do arquivo diz {comp_arquivo}, mas o recibo é de {comp_pdf}")
    if not item.competencia:
        item.situacao = Situacao.REVISAR
        avisos.append("Competência não identificada")
    elif competencia_sel and item.competencia != competencia_sel:
        return bloquear(f"Recibo é de {item.competencia}, diferente da selecionada ({competencia_sel})")
    if len(dados.tipos) > 1:
        avisos.append("Mais de um tipo de folha no recibo: " + ", ".join(sorted(dados.tipos)))
    elif not item.tipo_folha:
        avisos.append("Tipo de folha não identificado")

    # --- código e nome impressos no recibo
    if conferir_conteudo:
        if not texto.strip():
            item.situacao = Situacao.REVISAR
            avisos.append("PDF sem texto (escaneado?)")
        elif cfg.regex_cabecalho_pdf:
            if not dados.pares:
                item.situacao = Situacao.REVISAR
                avisos.append("Código/nome não localizados no recibo (layout diferente?)")
            elif len(dados.pares) > 1:
                return bloquear("O PDF contém mais de um colaborador: "
                                + ", ".join(f"{c} {n}" for c, n in sorted(dados.pares)))
            else:
                cod_pdf, nome_pdf = next(iter(dados.pares))
                divergencias = []
                if cod_pdf != colab.codigo:
                    divergencias.append(f"código no recibo {cod_pdf} ≠ planilha {colab.codigo}")
                if nome_pdf != normalizar_texto(colab.nome):
                    divergencias.append(f"nome no recibo '{nome_pdf}' ≠ planilha")
                if divergencias:
                    return bloquear("; ".join(divergencias))
        else:
            if not nome_confere(colab.nome, texto):
                item.situacao = Situacao.REVISAR
                avisos.append("Nome não encontrado dentro do PDF")
            if cfg.conferir_codigo_pdf and not codigo_no_texto(colab.codigo, texto):
                item.situacao = Situacao.REVISAR
                avisos.append(f"Código {colab.codigo} não encontrado dentro do PDF")

    # --- histórico
    try:
        item.sha256 = sha256_arquivo(arq)
    except OSError as exc:
        return bloquear(f"Não foi possível ler o arquivo: {exc}")
    if (codigo, item.sha256) in enviados_arquivo:
        item.situacao, item.detalhe = Situacao.JA_ENVIADO, "Este mesmo arquivo já foi enviado"
        return item
    if (item.competencia, normalizar_texto(item.tipo_folha), codigo) in enviados_competencia:
        avisos.append("Reenvio: já houve envio desta competência e tipo com outro arquivo")

    if item.situacao == Situacao.REVISAR:
        avisos.append("abra (duplo clique) e libere se estiver correto")
    item.detalhe = " | ".join(avisos)
    return item


def montar_lote(
    planilha: Path,
    pasta_pdfs: Path,
    cfg: Config,
    competencia: str,
    caminho_historico: Path,
    conferir_conteudo: bool,
) -> list[ItemEnvio]:
    """
    competencia = "MM/AAAA" exige essa competência em todos os recibos;
    competencia = "" (automática) usa a competência lida de cada recibo.
    """
    colaboradores, itens = ler_planilha(planilha)
    por_chave, fora_padrao, _ = varrer_pdfs(pasta_pdfs, cfg)
    enviados_arquivo, enviados_competencia = carregar_historico(caminho_historico)

    # Índice chave -> colaborador. Homônimos são bloqueados (poderia ir para a pessoa errada).
    indice: dict[str, Colaborador] = {}
    homonimos: set[str] = set()
    for colab in colaboradores.values():
        chave = chave_do_colaborador(colab, cfg)
        if chave in indice:
            homonimos.add(chave)
        indice[chave] = colab
    for chave in homonimos:
        indice.pop(chave, None)
    for colab in colaboradores.values():
        if chave_do_colaborador(colab, cfg) in homonimos:
            itens.append(ItemEnvio(colab.codigo, colab.nome, colab.email, None, Situacao.ERRO,
                                   f"Linha {colab.linha}: nome repetido na planilha (homônimo)", colab))

    padrao_txt = "nome do colaborador" if cfg.identificar_por == "nome" else "código"
    for arq in fora_padrao:
        itens.append(ItemEnvio("", "-", "", arq, Situacao.ERRO,
                               f"Nome do arquivo fora do padrão ({padrao_txt} não encontrado)"))

    avaliados: list[ItemEnvio] = []
    for chave, arquivos in sorted(por_chave.items(), key=lambda kv: kv[0]):
        for arq in arquivos:
            if chave in homonimos:
                avaliados.append(ItemEnvio("", "-", "", arq, Situacao.ERRO,
                                           "Nome com homônimo na planilha — não é possível identificar"))
                continue
            avaliados.append(_avaliar_arquivo(arq, chave, indice.get(chave), cfg, competencia,
                                              conferir_conteudo, enviados_arquivo, enviados_competencia))

    # Dois PDFs para o mesmo colaborador + competência + tipo -> bloqueia todos
    ativos = (Situacao.PRONTO, Situacao.REVISAR)
    grupos = Counter((i.codigo, i.competencia, normalizar_texto(i.tipo_folha))
                     for i in avaliados if i.situacao in ativos)
    for it in avaliados:
        if it.situacao in ativos and grupos[(it.codigo, it.competencia, normalizar_texto(it.tipo_folha))] > 1:
            it.situacao = Situacao.ERRO
            it.detalhe = "Mais de um PDF para o mesmo colaborador, competência e tipo de folha"

    # Competência automática: arquivo de mês diferente da maioria vai para revisão
    if not competencia:
        comps = Counter(i.competencia for i in avaliados if i.situacao in ativos and i.competencia)
        if len(comps) > 1:
            maioria = comps.most_common(1)[0][0]
            for it in avaliados:
                if it.situacao in ativos and it.competencia and it.competencia != maioria:
                    it.situacao = Situacao.REVISAR
                    it.detalhe = " | ".join(x for x in (
                        f"Competência {it.competencia} diferente da maioria ({maioria})", it.detalhe,
                        "" if "libere" in it.detalhe else "abra (duplo clique) e libere se estiver correto") if x)
    itens.extend(avaliados)

    # Colaboradores ativos sem PDF (informativo)
    for colab in colaboradores.values():
        chave = chave_do_colaborador(colab, cfg)
        if colab.ativo and chave not in homonimos and chave not in por_chave:
            itens.append(ItemEnvio(colab.codigo, colab.nome, colab.email, None, Situacao.PENDENTE,
                                   "Nenhum PDF na pasta para este colaborador", colab))

    log.info("Conferência concluída: %d itens (%d prontos).",
             len(itens), sum(1 for i in itens if i.situacao == Situacao.PRONTO))
    return itens


# ---------------------------------------------------------------------------
# Envio
# ---------------------------------------------------------------------------
class EnviadorOutlook:
    """
    Usa o Outlook clássico já logado na máquina do RH. Nenhuma credencial no programa.
    Automação via comtypes (Python puro) — sem pywin32, que traz DLLs sem assinatura.
    """

    def __init__(self, remetente: str = ""):
        self.remetente = remetente
        self.app = None
        self.conta = None

    def __enter__(self):
        import comtypes
        import comtypes.client

        comtypes.CoInitialize()  # obrigatório: o envio roda numa thread separada da tela
        self._comtypes = comtypes
        try:
            self.app = comtypes.client.CreateObject("Outlook.Application", dynamic=True)
        except Exception as exc:
            comtypes.CoUninitialize()
            raise RuntimeError(
                "Não foi possível abrir o Outlook. Use o Outlook clássico (o 'novo Outlook' "
                "não aceita automação) e deixe-o aberto e logado."
            ) from exc
        if self.remetente:
            contas = self.app.Session.Accounts
            for i in range(1, contas.Count + 1):
                conta = contas.Item(i)
                if str(conta.SmtpAddress).lower() == self.remetente.lower():
                    self.conta = conta
                    break
            if self.conta is None:
                log.info("Conta %s não está no perfil do Outlook; usando 'Em nome de'.", self.remetente)
        return self

    def enviar(self, para: str, assunto: str, corpo_html: str, anexo: Path) -> None:
        mail = self.app.CreateItem(0)  # 0 = olMailItem
        mail.To = para
        mail.Subject = assunto
        mail.HTMLBody = corpo_html
        mail.Attachments.Add(str(anexo))
        if self.conta is not None:
            try:
                mail.SendUsingAccount = self.conta
            except Exception:
                log.warning("SendUsingAccount falhou; usando SentOnBehalfOfName.")
                mail.SentOnBehalfOfName = self.remetente
        elif self.remetente:
            # Caixa compartilhada (ex.: rh@empresa) — exige permissão "Enviar como"/"Em nome de"
            mail.SentOnBehalfOfName = self.remetente
        mail.Send()

    def __exit__(self, *exc):
        import gc

        self.conta = None
        self.app = None
        gc.collect()  # libera as referências COM antes de encerrar o COM da thread
        self._comtypes.CoUninitialize()
        return False


class EnviadorSMTP:
    """SMTP com usuário da tela de Configurações e senha do cofre do Windows (DPAPI)."""

    def __init__(self, cfg: Config, senha: str):
        if not all([cfg.smtp_host, cfg.smtp_usuario, senha]):
            raise RuntimeError("SMTP incompleto: informe servidor, usuário e senha em Configurações.")
        self.cfg = cfg
        self._senha = senha
        self.remetente = cfg.smtp_remetente or cfg.smtp_usuario
        self.cliente: Optional[smtplib.SMTP] = None

    def __enter__(self):
        self.cliente = conectar_smtp(self.cfg, self._senha)
        self._senha = ""  # não mantém a senha em memória além do necessário
        return self

    def enviar(self, para: str, assunto: str, corpo_html: str, anexo: Path) -> None:
        msg = EmailMessage()
        msg["From"] = self.remetente
        msg["To"] = para
        msg["Subject"] = assunto
        msg.set_content("Seu contracheque está em anexo. Abra este e-mail em um leitor com suporte a HTML.")
        msg.add_alternative(corpo_html, subtype="html")
        msg.add_attachment(anexo.read_bytes(), maintype="application", subtype="pdf",
                           filename=anexo.name)
        self.cliente.send_message(msg)

    def __exit__(self, *exc):
        try:
            if self.cliente:
                self.cliente.quit()
        except Exception:
            pass
        return False


def conectar_smtp(cfg: Config, senha: str) -> smtplib.SMTP:
    contexto = ssl.create_default_context()
    if cfg.smtp_seguranca == "ssl":
        cliente = smtplib.SMTP_SSL(cfg.smtp_host, cfg.smtp_porta, timeout=60, context=contexto)
    else:
        cliente = smtplib.SMTP(cfg.smtp_host, cfg.smtp_porta, timeout=60)
        cliente.starttls(context=contexto)
    try:
        cliente.login(cfg.smtp_usuario, senha)
    except smtplib.SMTPAuthenticationError as exc:
        cliente.close()
        detalhe = exc.smtp_error.decode(errors="ignore") if isinstance(exc.smtp_error, bytes) else str(exc)
        if "5.7.139" in detalhe or "5.7.30" in detalhe or "basic authentication" in detalhe.lower():
            raise RuntimeError(
                "O Microsoft 365 recusou login com senha (SMTP AUTH desativado). "
                "Use o modo Outlook ou peça ao administrador para liberar SMTP AUTH nesta caixa."
            ) from exc
        raise RuntimeError(f"Usuário ou senha recusados pelo servidor: {detalhe}") from exc
    return cliente


def testar_smtp(cfg: Config, senha: str) -> str:
    cliente = conectar_smtp(cfg, senha)
    try:
        cliente.noop()
    finally:
        try:
            cliente.quit()
        except Exception:
            pass
    return f"Conectado a {cfg.smtp_host}:{cfg.smtp_porta} como {cfg.smtp_usuario}"


def listar_contas_outlook() -> list[str]:
    """Contas configuradas no Outlook clássico desta máquina."""
    import gc

    import comtypes
    import comtypes.client

    comtypes.CoInitialize()
    try:
        app = comtypes.client.CreateObject("Outlook.Application", dynamic=True)
        contas = app.Session.Accounts
        enderecos = [str(contas.Item(i).SmtpAddress) for i in range(1, contas.Count + 1)]
        del contas, app
        return enderecos
    finally:
        gc.collect()
        comtypes.CoUninitialize()


class EnviadorSimulado:
    """Não envia nada: grava os PDFs protegidos e o e-mail em logs/simulacao para conferência."""

    def __init__(self, base: Path):
        self.pasta = base / "logs" / "simulacao" / datetime.now().strftime("%Y%m%d_%H%M%S")

    def __enter__(self):
        self.pasta.mkdir(parents=True, exist_ok=True)
        return self

    def enviar(self, para: str, assunto: str, corpo_html: str, anexo: Path) -> None:
        shutil.copy2(anexo, self.pasta / anexo.name)
        (self.pasta / f"{anexo.stem}.html").write_text(
            f"<p><b>Para:</b> {html.escape(para)}<br><b>Assunto:</b> {html.escape(assunto)}</p>{corpo_html}",
            encoding="utf-8",
        )

    def __exit__(self, *exc):
        return False


def criar_enviador(cfg: Config, base: Path):
    if cfg.modo_envio == "smtp":
        return EnviadorSMTP(cfg, cofre.ler_senha(base))
    if cfg.modo_envio == "simulado":
        return EnviadorSimulado(base)
    return EnviadorOutlook(cfg.remetente)


ARQUIVO_MENSAGEM = "mensagem_email.txt"
MARCADORES = {
    "{primeiro_nome}": "Primeiro nome",
    "{nome}": "Nome completo",
    "{tipo_folha}": "Tipo (Folha Mensal, Férias, 13º...)",
    "{mes_extenso}": "Mês por extenso (setembro de 2026)",
    "{competencia}": "Competência (09/2026)",
    "{dica_senha}": "Como abrir o PDF",
    "{empresa}": "Nome da empresa",
}
MENSAGEM_PADRAO = """Olá, {primeiro_nome}!

Segue em anexo o seu contracheque de **{tipo_folha}** referente a **{mes_extenso}**.

Para abrir o arquivo, use como senha **{dica_senha}**.

Em caso de dúvida, procure o RH.

Atenciosamente,
Recursos Humanos — {empresa}
"""


def ler_mensagem(base: Path) -> str:
    caminho = base / ARQUIVO_MENSAGEM
    return caminho.read_text(encoding="utf-8") if caminho.exists() else MENSAGEM_PADRAO


def salvar_mensagem(base: Path, texto: str) -> None:
    caminho = base / ARQUIVO_MENSAGEM
    temporario = caminho.with_suffix(".tmp")
    temporario.write_text(texto.strip() + "\n", encoding="utf-8")
    os.replace(temporario, caminho)


def _valores_marcadores(item: ItemEnvio, cfg: Config) -> dict[str, str]:
    primeiro = item.nome.split()[0].capitalize() if item.nome and item.nome != "-" else ""
    return {
        "{primeiro_nome}": primeiro,
        "{nome}": nome_proprio(item.nome) if item.nome != "-" else "",
        "{tipo_folha}": item.tipo_folha or "Folha",
        "{mes_extenso}": mes_extenso(item.competencia),
        "{competencia}": item.competencia,
        "{dica_senha}": descricao_senha(cfg.formato_senha),
        "{empresa}": cfg.nome_empresa,
    }


def aplicar_marcadores(texto: str, item: ItemEnvio, cfg: Config) -> str:
    for chave, valor in _valores_marcadores(item, cfg).items():
        texto = texto.replace(chave, valor)
    return re.sub(r"\s{2,}", " ", texto).strip() if "\n" not in texto else texto


def montar_assunto(item: ItemEnvio, cfg: Config) -> str:
    return aplicar_marcadores(cfg.assunto, item, cfg)


def _texto_para_html(texto: str) -> str:
    """Texto simples do editor -> HTML. Linha em branco = parágrafo; **texto** = negrito."""
    paragrafos = []
    for bloco in re.split(r"\n\s*\n", texto.strip()):
        linhas = [html.escape(l) for l in bloco.splitlines()]
        conteudo = "<br>".join(linhas)
        conteudo = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", conteudo)
        paragrafos.append(f'<p style="margin:0 0 14px 0;">{conteudo}</p>')
    return "\n".join(paragrafos)


def montar_corpo(modelo_texto: str, item: ItemEnvio, cfg: Config) -> str:
    """Layout em tabela (compatível com Outlook) nas cores da empresa."""
    corpo = _texto_para_html(aplicar_marcadores(modelo_texto, item, cfg))
    tipo = html.escape(item.tipo_folha or "Contracheque")
    periodo = html.escape(mes_extenso(item.competencia).capitalize())
    empresa = html.escape(cfg.nome_empresa)
    p, d, f = cfg.cor_primaria, cfg.cor_destaque, cfg.cor_fundo
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width"></head>
<body style="margin:0;padding:0;background:{f};">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:{f};padding:24px 0;">
<tr><td align="center">
  <table role="presentation" width="600" cellpadding="0" cellspacing="0"
         style="width:600px;max-width:100%;background:#FFFFFF;border-radius:8px;overflow:hidden;
                font-family:'Segoe UI',Arial,sans-serif;color:#1E2329;">
    <tr><td style="background:{p};padding:22px 28px;">
      <div style="font-size:13px;letter-spacing:2px;text-transform:uppercase;color:#FFFFFF;opacity:.85;">{empresa}</div>
      <div style="font-size:22px;font-weight:600;color:#FFFFFF;margin-top:4px;">Contracheque</div>
    </td></tr>
    <tr><td style="background:{d};height:4px;line-height:4px;font-size:0;">&nbsp;</td></tr>
    <tr><td style="padding:18px 28px 0 28px;">
      <table role="presentation" cellpadding="0" cellspacing="0"><tr>
        <td style="background:{f};border-left:4px solid {d};padding:10px 16px;font-size:13px;">
          <span style="color:#5B6573;">Tipo</span><br><strong style="font-size:15px;color:{p};">{tipo}</strong>
        </td>
        <td style="width:10px;"></td>
        <td style="background:{f};border-left:4px solid {d};padding:10px 16px;font-size:13px;">
          <span style="color:#5B6573;">Período</span><br><strong style="font-size:15px;color:{p};">{periodo}</strong>
        </td>
      </tr></table>
    </td></tr>
    <tr><td style="padding:22px 28px 8px 28px;font-size:15px;line-height:1.55;">
{corpo}
    </td></tr>
    <tr><td style="padding:14px 28px 22px 28px;border-top:1px solid #E6EAF0;font-size:11px;color:#7A8491;">
      Documento pessoal e confidencial. Não encaminhe a terceiros.<br>
      O RH nunca solicita senhas ou dados bancários por e-mail.
    </td></tr>
  </table>
</td></tr>
</table>
</body></html>"""


def mover_para_enviados(arquivo: Path, competencia: str) -> None:
    destino_dir = arquivo.parent / "Enviados" / competencia_iso(competencia)
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / arquivo.name
    if destino.exists():
        destino = destino_dir / f"{arquivo.stem}_{datetime.now():%Y%m%d%H%M%S}{arquivo.suffix}"
    shutil.move(str(arquivo), str(destino))


def enviar_lote(
    itens: list[ItemEnvio],
    cfg: Config,
    base: Path,
    email_teste: str = "",
    progresso: Optional[Callable[[int, int, ItemEnvio], None]] = None,
    cancelar: Optional[threading.Event] = None,
) -> dict:
    """Envia somente os itens com situação PRONTO. Cada item usa a SUA competência e tipo."""
    alvo = [i for i in itens if i.situacao == Situacao.PRONTO]
    total = len(alvo)
    resumo = {"enviados": 0, "erros": 0, "cancelado": False}
    if total == 0:
        return resumo

    modelo_texto = ler_mensagem(base)
    historico = base / "logs" / "historico_envios.csv"
    modo_teste = bool(email_teste)
    simulado = cfg.modo_envio == "simulado"

    log.info("Iniciando envio: %d itens | modo=%s%s", total, cfg.modo_envio,
             " | TESTE -> " + email_teste if modo_teste else "")

    with tempfile.TemporaryDirectory(prefix="contracheques_") as tmp, criar_enviador(cfg, base) as enviador:
        for n, item in enumerate(alvo, start=1):
            if cancelar is not None and cancelar.is_set():
                resumo["cancelado"] = True
                log.warning("Envio cancelado pelo usuário após %d itens.", n - 1)
                break

            tipo_arquivo = nome_arquivo_seguro(item.tipo_folha or "Contracheque").replace(" ", "_")
            anexo = Path(tmp) / f"Contracheque_{tipo_arquivo}_{competencia_iso(item.competencia)}_{item.codigo}.pdf"
            destino = email_teste if modo_teste else item.email
            try:
                senha = montar_senha(cfg.formato_senha, item.colaborador)
                proteger_pdf(item.arquivo, anexo, senha,
                             titulo=f"Contracheque {item.tipo_folha} {item.competencia} - {nome_proprio(item.nome)}")
                del senha

                assunto = montar_assunto(item, cfg)
                if modo_teste:
                    assunto = f"[TESTE p/ {item.email}] {assunto}"
                enviador.enviar(destino, assunto, montar_corpo(modelo_texto, item, cfg), anexo)

                if simulado:
                    item.situacao = Situacao.SIMULADO
                elif modo_teste:
                    item.situacao = Situacao.TESTE
                else:
                    item.situacao = Situacao.ENVIADO
                item.detalhe = f"Para {destino}"
                if item.situacao == Situacao.ENVIADO and cfg.mover_enviados:
                    # Falha ao mover NÃO pode virar "erro de envio": o e-mail já saiu.
                    try:
                        mover_para_enviados(item.arquivo, item.competencia)
                    except OSError as exc:
                        item.detalhe += f" | Aviso: não moveu o PDF ({exc})"
                        log.warning("Enviado, mas não moveu %s: %s", item.arquivo, exc)
                resumo["enviados"] += 1
                log.info("OK  %s | %s | %s %s | %s", item.codigo, item.nome, item.tipo_folha,
                         item.competencia, destino)
            except Exception as exc:
                item.situacao, item.detalhe = Situacao.ERRO, f"Falha no envio: {exc}"
                resumo["erros"] += 1
                log.exception("ERRO %s | %s", item.codigo, item.nome)
            finally:
                try:
                    anexo.unlink(missing_ok=True)
                except OSError:
                    pass

            registrar_historico(historico, item, destino)
            if progresso:
                progresso(n, total, item)
            if n < total:
                time.sleep(cfg.pausa_segundos)

    log.info("Envio finalizado: %s", resumo)
    return resumo
