"""
Guarda a senha do e-mail (modo SMTP) criptografada pelo próprio Windows (DPAPI).

- A senha NUNCA fica em texto puro em arquivo, log ou config.
- O arquivo credencial.dat só pode ser aberto pelo MESMO usuário do Windows, na MESMA
  máquina. Copiado para outro computador ou outro login, ele não serve para nada.
- Usa só ctypes + crypt32.dll do Windows (assinada pela Microsoft) — nada de binário extra.
"""

from __future__ import annotations

import base64
import ctypes
import os
from pathlib import Path

ARQUIVO = "credencial.dat"
_ENTROPIA = b"EnvioContracheques::smtp::v1"   # segredo adicional amarrado a este programa
_CRYPTPROTECT_UI_FORBIDDEN = 0x01


class CofreIndisponivel(RuntimeError):
    """Fora do Windows não há DPAPI — o programa não guarda senha nesse caso."""


if os.name == "nt":
    from ctypes import wintypes

    class _DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    _crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    _crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(_DATA_BLOB), wintypes.LPCWSTR, ctypes.POINTER(_DATA_BLOB),
        ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_DATA_BLOB),
    ]
    _crypt32.CryptProtectData.restype = wintypes.BOOL
    _crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(_DATA_BLOB), ctypes.POINTER(wintypes.LPWSTR), ctypes.POINTER(_DATA_BLOB),
        ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_DATA_BLOB),
    ]
    _crypt32.CryptUnprotectData.restype = wintypes.BOOL
    _kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    _kernel32.LocalFree.restype = ctypes.c_void_p

    def _para_blob(dados: bytes):
        buffer = ctypes.create_string_buffer(dados, len(dados))
        blob = _DATA_BLOB(len(dados), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char)))
        return blob, buffer  # o buffer precisa continuar vivo enquanto o blob for usado

    def _do_blob(blob: "_DATA_BLOB") -> bytes:
        try:
            return ctypes.string_at(blob.pbData, blob.cbData)
        finally:
            _kernel32.LocalFree(ctypes.cast(blob.pbData, ctypes.c_void_p))

    def _proteger(dados: bytes) -> bytes:
        entrada, _b1 = _para_blob(dados)
        entropia, _b2 = _para_blob(_ENTROPIA)
        saida = _DATA_BLOB()
        if not _crypt32.CryptProtectData(ctypes.byref(entrada), "EnvioContracheques", ctypes.byref(entropia),
                                         None, None, _CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(saida)):
            raise ctypes.WinError(ctypes.get_last_error())
        return _do_blob(saida)

    def _desproteger(dados: bytes) -> bytes:
        entrada, _b1 = _para_blob(dados)
        entropia, _b2 = _para_blob(_ENTROPIA)
        saida = _DATA_BLOB()
        if not _crypt32.CryptUnprotectData(ctypes.byref(entrada), None, ctypes.byref(entropia),
                                           None, None, _CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(saida)):
            raise ctypes.WinError(ctypes.get_last_error())
        return _do_blob(saida)


def disponivel() -> bool:
    return os.name == "nt"


def salvar_senha(base: Path, senha: str) -> None:
    if not disponivel():
        raise CofreIndisponivel("Guardar senha só é suportado no Windows.")
    if not senha:
        raise ValueError("Senha vazia.")
    cifrado = _proteger(senha.encode("utf-8"))
    (base / ARQUIVO).write_text(base64.b64encode(cifrado).decode("ascii"), encoding="ascii")


def ler_senha(base: Path) -> str:
    caminho = base / ARQUIVO
    if not caminho.exists():
        raise FileNotFoundError("Nenhuma senha de e-mail salva. Abra Configurações e informe a senha.")
    if not disponivel():
        raise CofreIndisponivel("Guardar senha só é suportado no Windows.")
    try:
        dados = base64.b64decode(caminho.read_text(encoding="ascii"))
        return _desproteger(dados).decode("utf-8")
    except OSError as exc:
        raise RuntimeError(
            "A senha salva não pode ser lida por este usuário/computador. "
            "Informe a senha novamente em Configurações."
        ) from exc


def tem_senha(base: Path) -> bool:
    return (base / ARQUIVO).exists()


def apagar_senha(base: Path) -> None:
    (base / ARQUIVO).unlink(missing_ok=True)
