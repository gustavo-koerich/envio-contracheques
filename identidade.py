"""
Identidade visual: carrega a logo e extrai dela as cores da marca.

Só usa o Tk (que vem no Python oficial) — sem Pillow, que traria binário sem assinatura.
O Tk 8.6 lê PNG e GIF nativamente.
"""

from __future__ import annotations

import math
import tkinter as tk
from collections import Counter
from pathlib import Path
from typing import Optional


def _rgb(img: tk.PhotoImage, x: int, y: int) -> tuple[int, int, int]:
    valor = img.get(x, y)
    if isinstance(valor, str):  # versões antigas do Tk devolvem "r g b"
        valor = tuple(int(v) for v in valor.split())
    return int(valor[0]), int(valor[1]), int(valor[2])


def _transparente(img: tk.PhotoImage, x: int, y: int) -> bool:
    try:
        return bool(img.transparency_get(x, y))
    except (AttributeError, tk.TclError):
        return False


def hex_de(rgb: tuple[float, float, float]) -> str:
    return "#" + "".join(f"{max(0, min(255, round(c))):02X}" for c in rgb)


def rgb_de(cor: str) -> tuple[int, int, int]:
    return tuple(int(cor[i:i + 2], 16) for i in (1, 3, 5))  # type: ignore[return-value]


def luminancia(cor: str) -> float:
    """Luminância relativa (WCAG)."""
    def canal(c: int) -> float:
        c = c / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = rgb_de(cor)
    return 0.2126 * canal(r) + 0.7152 * canal(g) + 0.0722 * canal(b)


def contraste(a: str, b: str) -> float:
    la, lb = sorted((luminancia(a), luminancia(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def ajustar_para_texto_branco(cor: str, minimo: float = 4.5) -> str:
    """Escurece a cor até o texto branco por cima atingir o contraste AA."""
    r, g, b = rgb_de(cor)
    for _ in range(40):
        atual = hex_de((r, g, b))
        if contraste(atual, "#FFFFFF") >= minimo:
            return atual
        r, g, b = r * 0.94, g * 0.94, b * 0.94
    return hex_de((r, g, b))


def _matiz(cor: tuple[int, int, int]) -> float:
    r, g, b = (c / 255 for c in cor)
    mx, mn = max(r, g, b), min(r, g, b)
    if mx == mn:
        return 0.0
    if mx == r:
        h = (g - b) / (mx - mn) % 6
    elif mx == g:
        h = (b - r) / (mx - mn) + 2
    else:
        h = (r - g) / (mx - mn) + 4
    return h * 60


def carregar_logo(caminho: Path, altura_max: int) -> Optional[tk.PhotoImage]:
    """Lê PNG/GIF e reduz por fator inteiro até caber na altura (Tk não redimensiona livremente)."""
    if not caminho.is_file():
        return None
    try:
        img = tk.PhotoImage(file=str(caminho))
    except tk.TclError:
        return None
    fator = math.ceil(img.height() / altura_max) if img.height() > altura_max else 1
    return img.subsample(fator, fator) if fator > 1 else img


def cores_da_logo(caminho: Path) -> Optional[tuple[str, str]]:
    """
    Retorna (primária, destaque) a partir das cores saturadas mais frequentes da logo.
    Ignora transparente, branco, preto e cinzas. Ambas ajustadas para texto branco legível.
    """
    if not caminho.is_file():
        return None
    try:
        img = tk.PhotoImage(file=str(caminho))
    except tk.TclError:
        return None
    w, h = img.width(), img.height()
    passo = max(1, int(max(w, h) / 90))
    baldes: Counter = Counter()
    somas: dict[tuple[int, int, int], list[float]] = {}
    for y in range(0, h, passo):
        for x in range(0, w, passo):
            if _transparente(img, x, y):
                continue
            r, g, b = _rgb(img, x, y)
            mx, mn = max(r, g, b), min(r, g, b)
            if mx - mn < 45 or mx < 40 or mn > 225:  # cinza, preto ou branco
                continue
            chave = (r // 24, g // 24, b // 24)
            baldes[chave] += 1
            acumulado = somas.setdefault(chave, [0.0, 0.0, 0.0])
            acumulado[0] += r
            acumulado[1] += g
            acumulado[2] += b
    if not baldes:
        return None

    def media(chave):
        n = baldes[chave]
        return tuple(v / n for v in somas[chave])

    ordenados = [media(k) for k, _ in baldes.most_common(12)]
    principal = ordenados[0]
    destaque = next((c for c in ordenados[1:] if min(abs(_matiz(c) - _matiz(principal)),
                                                     360 - abs(_matiz(c) - _matiz(principal))) > 35), None)
    cor_principal = ajustar_para_texto_branco(hex_de(principal))
    if destaque is None:
        # Sem segunda cor na marca: destaque = a própria cor da marca, mais clara
        r, g, b = principal
        destaque = (r + (255 - r) * 0.25, g + (255 - g) * 0.25, b + (255 - b) * 0.25)
    return cor_principal, ajustar_para_texto_branco(hex_de(destaque), 3.0)


def escalar(img: tk.PhotoImage, proporcao: float) -> tk.PhotoImage:
    """
    Redimensiona pela melhor fração zoom/subsample <= proporção (o Tk só escala por inteiros).
    Ex.: 0,49 vira 1/2 em vez de 1/3, aproveitando ao máximo o espaço.
    """
    if abs(proporcao - 1) < 0.01:
        return img
    melhor = (1, max(1, math.ceil(1 / proporcao)))
    for a in range(1, 9):
        for b in range(1, 17):
            if a / b <= proporcao and a / b > melhor[0] / melhor[1]:
                melhor = (a, b)
    a, b = melhor
    resultado = img.zoom(a, a) if a > 1 else img
    return resultado.subsample(b, b) if b > 1 else resultado


def gerar_icone(logo: Path, destino: Path) -> bool:
    """
    Gera um .ico (256, 128, 64, 32 e 16 px) a partir da logo, só com Tk + Python puro.
    A logo é centralizada num quadrado transparente; cada tamanho vai como PNG dentro do .ico
    (formato aceito pelo Windows desde o Vista). Retorna False se não houver logo legível.
    """
    import struct
    import tempfile

    if not logo.is_file():
        return False
    try:
        original = tk.PhotoImage(file=str(logo))
    except tk.TclError:
        return False

    lado = 256
    base = escalar(original, lado / max(original.width(), original.height()))

    quadro = tk.PhotoImage(width=lado, height=lado)  # começa 100% transparente
    x = (lado - base.width()) // 2
    y = (lado - base.height()) // 2
    quadro.tk.call(quadro, "copy", base, "-to", x, y)

    imagens = []
    with tempfile.TemporaryDirectory() as tmp:
        for divisor in (1, 2, 4, 8, 16):
            img = quadro if divisor == 1 else quadro.subsample(divisor, divisor)
            arquivo = Path(tmp) / f"{divisor}.png"
            img.write(str(arquivo), format="png")
            imagens.append((lado // divisor, arquivo.read_bytes()))

    cabecalho = struct.pack("<HHH", 0, 1, len(imagens))
    deslocamento = 6 + 16 * len(imagens)
    entradas, dados = b"", b""
    for tamanho, png in imagens:
        dim = 0 if tamanho >= 256 else tamanho  # 0 = 256 no formato ICO
        entradas += struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, 32, len(png), deslocamento + len(dados))
        dados += png
    destino.write_bytes(cabecalho + entradas + dados)
    return True
