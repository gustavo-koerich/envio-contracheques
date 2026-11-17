"""Abre o Envio de Contracheques (duplo clique). Roda com o pythonw.exe oficial, sem console."""
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
# Isola o programa das bibliotecas instaladas no Python da máquina (site-packages).
# Usa só a biblioteca padrão + a pasta lib/. Evita conflito com versões instaladas e com
# pacotes compilados (numpy, lxml...) que o Controle de Aplicativo Inteligente pode bloquear.
sys.path[:] = [p for p in sys.path if "site-packages" not in p.lower() and "dist-packages" not in p.lower()]
sys.path[:0] = [str(BASE / "lib"), str(BASE)]

import app  # noqa: E402

app.main()
