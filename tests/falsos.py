"""Dados de exemplo, um navegador falso e um coletor falso, usados pelos testes."""

import os
from datetime import date
from typing import ClassVar

from PIL import Image

import coletor


ID_A = "1234567890123456"
ID_B = "6543210987654321"
ID_C = "1111222233334444"

TEXTO_COM_DATA = "Maria da Silva\nÚltima atualização do currículo em 15/03/2024\nProdução bibliográfica"


class ElementoFalso:
    def __init__(self, texto):
        self.text = texto


class DriverFalso:
    """
    Imita a parte do WebDriver usada por coletor.ColetorLattes.

    `paginas` é o que o navegador mostra a cada verificação do coletor:
    "curriculo", "captcha", "outra" (nem currículo nem CAPTCHA) ou
    "carregando". A última página da sequência se repete, e a sequência
    recomeça a cada get().
    """

    current_window_handle = "aba-1"

    def __init__(self, nome="Maria da Silva", texto=TEXTO_COM_DATA, erro_ao_abrir=None,
                 paginas=("curriculo",), titulo="Currículo do Sistema de Currículos Lattes",
                 texto_outra="Currículo não encontrado"):
        self.nome = nome
        self.texto = texto
        self.erro_ao_abrir = erro_ao_abrir
        self.paginas = list(paginas)
        self.title = titulo
        self.texto_outra = texto_outra
        self.verificacoes = 0
        self.urls = []
        self.fechado = False
        self.window_handles = ["aba-1"]

    @property
    def pagina(self):
        indice = min(max(self.verificacoes - 1, 0), len(self.paginas) - 1)
        return self.paginas[indice]

    def get(self, url):
        if self.erro_ao_abrir is not None:
            raise self.erro_ao_abrir
        self.urls.append(url)
        self.verificacoes = 0

    def find_elements(self, by, seletor):
        if seletor == coletor.SELETOR_NOME and self.pagina == "curriculo":
            return [ElementoFalso(f"  {self.nome}  ")]
        if seletor == coletor.SELETOR_CAPTCHA and self.pagina == "captcha":
            return [ElementoFalso("")]
        return []

    def execute_script(self, script):
        if "readyState" in script:
            self.verificacoes += 1
            return "loading" if self.pagina == "carregando" else "complete"
        if "innerText" in script:
            return self.texto if self.pagina == "curriculo" else self.texto_outra
        raise AssertionError(f"script inesperado: {script}")

    def save_screenshot(self, caminho):
        Image.new("RGB", (640, 400), "white").save(caminho)
        return True

    def quit(self):
        self.fechado = True


class ColetorFalso:
    """
    Substitui coletor.ColetorLattes nos testes da janela, sem navegador.

    `roteiro` diz o status de cada ID ({id: status} ou {id: [status da 1ª
    tentativa, da 2ª, ...]}); IDs fora do roteiro dão OK. Com `bloquear`,
    cada captura espera o cancelamento (para testar o botão Cancelar).
    """

    roteiro: ClassVar[dict] = {}
    bloquear: ClassVar[bool] = False

    def __init__(self, cancelado, log=print, **_):
        self.cancelado = cancelado
        self.log = log
        self.fechado = False

    def capturar(self, lattes_id, pasta):
        if self.bloquear:
            self.cancelado.wait(10)
            raise coletor.CapturaCancelada()
        status = self.roteiro.get(lattes_id, coletor.STATUS_OK)
        if isinstance(status, list):
            status = status.pop(0) if len(status) > 1 else status[0]
        if status not in (coletor.STATUS_OK, coletor.STATUS_SEM_DATA):
            return coletor.Registro(lattes_id, status, observacao="falha simulada")
        caminho = os.path.join(pasta, f"{lattes_id}.png")
        Image.new("RGB", (64, 64), "white").save(caminho)
        data = date(2024, 3, 15) if status == coletor.STATUS_OK else None
        return coletor.Registro(lattes_id, status, nome=f"Pesquisador {lattes_id[-4:]}",
                                ultima_atualizacao=data, arquivo_print=caminho)

    def fechar(self):
        self.fechado = True
