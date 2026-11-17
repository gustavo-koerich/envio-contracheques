"""Testes das travas que não dependem de e-mail: senha do PDF e proteção AES-256."""
import sys
import tempfile
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(RAIZ), str(RAIZ / "lib")]

from pypdf import PdfReader, PdfWriter  # noqa: E402

import nucleo  # noqa: E402


def colaborador(cpf4="1234", nasc="1503"):
    return nucleo.Colaborador("1001", "MARIA SILVA", "maria@empresa-demo.com.br", cpf4, nasc, True, 2)


class TestSenha(unittest.TestCase):
    def test_formatos(self):
        self.assertEqual(nucleo.montar_senha("{cpf4}", colaborador()), "1234")
        self.assertEqual(nucleo.montar_senha("{cpf4}{nasc_ddmm}", colaborador()), "12341503")

    def test_sem_dado_bloqueia(self):
        with self.assertRaises(ValueError):
            nucleo.montar_senha("{cpf4}", colaborador(cpf4=None))
        with self.assertRaises(ValueError):
            nucleo.montar_senha("{nasc_ddmm}", colaborador(nasc=None))

    def test_descricao_nunca_traz_a_senha(self):
        texto = nucleo.descricao_senha("{cpf4}{nasc_ddmm}")
        self.assertIn("CPF", texto)
        self.assertNotIn("1234", texto)


class TestProtecao(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp())
        self.origem = self.pasta / "recibo.pdf"
        w = PdfWriter()
        w.add_blank_page(width=300, height=300)
        with open(self.origem, "wb") as f:
            w.write(f)

    def test_pdf_protegido_so_abre_com_a_senha(self):
        destino = self.pasta / "protegido.pdf"
        nucleo.proteger_pdf(self.origem, destino, "1234", "Contracheque")
        r = PdfReader(str(destino))
        self.assertTrue(r.is_encrypted)
        self.assertFalse(r.decrypt("0000"))
        self.assertTrue(PdfReader(str(destino)).decrypt("1234"))

    def test_recusa_pdf_ja_protegido(self):
        destino = self.pasta / "protegido.pdf"
        nucleo.proteger_pdf(self.origem, destino, "1234", "Contracheque")
        with self.assertRaises(ValueError):
            nucleo.proteger_pdf(destino, self.pasta / "de-novo.pdf", "1234", "Contracheque")


if __name__ == "__main__":
    unittest.main()
