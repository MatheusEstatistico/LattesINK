"""Navegadores: quais estão instalados e como cada um é aberto. Nenhum teste abre um navegador."""

import os
import sys

import pytest
from falsos import ID_A
from selenium import webdriver

import coletor


# --- onde cada navegador está instalado --------------------------------------------

@pytest.mark.skipif(not sys.platform.startswith("linux"), reason="procura no PATH do Linux")
def test_encontra_os_navegadores_no_path_do_linux(monkeypatch, tmp_path):
    for nome in ("firefox-esr", "brave-browser"):
        arquivo = tmp_path / nome
        arquivo.write_text("#!/bin/sh\n")
        arquivo.chmod(0o755)
    monkeypatch.setenv("PATH", str(tmp_path))
    # o Chromium do snap é procurado por um caminho fixo, que pode existir na máquina dos testes
    monkeypatch.setitem(coletor.LOCAIS_DOS_NAVEGADORES["linux"], "chromium", ["chromium"])

    assert coletor.navegadores_instalados() == ["firefox", "brave"]
    assert coletor.caminho_do_navegador("firefox") == str(tmp_path / "firefox-esr")
    assert coletor.caminho_do_navegador("chrome") is None


def test_encontra_os_navegadores_nas_pastas_de_programas_do_windows(monkeypatch, tmp_path):
    monkeypatch.setattr(coletor.sys, "platform", "win32")
    pastas = {}
    for variavel in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        pastas[variavel] = str(tmp_path / variavel.replace("(", "").replace(")", ""))
        monkeypatch.setenv(variavel, pastas[variavel])

    def instalar(variavel, relativo):
        caminho = os.path.join(pastas[variavel], relativo)
        os.makedirs(os.path.dirname(caminho), exist_ok=True)
        open(caminho, "w").close()
        return caminho

    instalar("PROGRAMFILES(X86)", r"Microsoft\Edge\Application\msedge.exe")
    brave = instalar("LOCALAPPDATA", r"BraveSoftware\Brave-Browser\Application\brave.exe")

    assert coletor.navegadores_instalados() == ["edge", "brave"]
    assert coletor.caminho_do_navegador("brave") == brave


def test_todo_navegador_tem_onde_ser_procurado_em_cada_sistema():
    for locais in coletor.LOCAIS_DOS_NAVEGADORES.values():
        assert set(locais) == set(coletor.NAVEGADORES)


# --- como cada navegador é aberto ----------------------------------------------------

class DriverAberto:
    def set_window_size(self, largura, altura):
        self.tamanho = (largura, altura)

    def set_page_load_timeout(self, segundos):
        self.tempo_carregamento = segundos


@pytest.fixture
def abertos(monkeypatch):
    """Troca as classes do Selenium por uma que só anota qual foi usada e com que opções."""
    lista = []
    for classe in ("Chrome", "Edge", "Firefox"):
        def abrir(options, _classe=classe):
            lista.append((_classe, options))
            return DriverAberto()
        monkeypatch.setattr(webdriver, classe, abrir)
    monkeypatch.setattr(coletor, "caminho_do_navegador", lambda navegador: f"/opt/{navegador}/{navegador}")
    return lista


@pytest.mark.parametrize(("navegador", "classe", "opcoes", "executavel"), [
    ("chrome", "Chrome", webdriver.ChromeOptions, ""),
    ("edge", "Edge", webdriver.EdgeOptions, ""),
    ("firefox", "Firefox", webdriver.FirefoxOptions, ""),
    ("chromium", "Chrome", webdriver.ChromeOptions, "/opt/chromium/chromium"),
    ("brave", "Chrome", webdriver.ChromeOptions, "/opt/brave/brave"),
])
def test_abre_o_navegador_escolhido(abertos, cancelado, navegador, classe, opcoes, executavel):
    driver = coletor.ColetorLattes(cancelado, navegador=navegador)._criar_driver()

    ((usada, opcoes_usadas),) = abertos
    assert usada == classe
    assert type(opcoes_usadas) is opcoes
    # o Selenium acha sozinho o Chrome, o Edge e o Firefox; o Chromium e o Brave vão pelo caminho
    assert opcoes_usadas.binary_location == executavel
    assert driver.tamanho == (1080, 720)
    assert driver.tempo_carregamento == coletor.TEMPO_MAX_CARREGAMENTO


def test_navegador_nao_encontrado_tem_mensagem_clara(abertos, monkeypatch, cancelado):
    monkeypatch.setattr(coletor, "caminho_do_navegador", lambda navegador: None)
    c = coletor.ColetorLattes(cancelado, log=lambda _: None, navegador="brave")
    with pytest.raises(coletor.ErroNavegador) as erro:
        c.capturar(ID_A, "")
    assert str(erro.value) == "Não foi possível abrir o Brave:\nele não foi encontrado neste computador"
    assert abertos == []  # e não abre o Chrome no lugar dele
