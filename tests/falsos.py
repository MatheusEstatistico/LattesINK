"""Dados de exemplo e um navegador falso, usados pelos testes."""

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

    `nome=None` simula uma página sem currículo (por exemplo, a do CAPTCHA).
    """

    current_window_handle = "aba-1"

    def __init__(self, nome="Maria da Silva", texto=TEXTO_COM_DATA, erro_ao_abrir=None):
        self.nome = nome
        self.texto = texto
        self.erro_ao_abrir = erro_ao_abrir
        self.urls = []
        self.fechado = False
        self.window_handles = ["aba-1"]

    def get(self, url):
        if self.erro_ao_abrir is not None:
            raise self.erro_ao_abrir
        self.urls.append(url)

    def find_elements(self, by, seletor):
        if seletor == coletor.SELETOR_NOME and self.nome is not None:
            return [ElementoFalso(f"  {self.nome}  ")]
        return []

    def find_element(self, by, seletor):
        return ElementoFalso(self.texto)

    def execute_script(self, script):
        if "readyState" in script:
            return "complete"
        return self.texto

    def save_screenshot(self, caminho):
        Image.new("RGB", (640, 400), "white").save(caminho)
        return True

    def quit(self):
        self.fechado = True
