"""ColetorLattes com um navegador falso: nenhum teste abre o Chrome."""

import os
import time
from datetime import date
from types import SimpleNamespace

import pytest
from falsos import ID_A, ID_B, DriverFalso
from selenium.common.exceptions import JavascriptException, NoSuchWindowException, WebDriverException

import coletor


def test_captura_bem_sucedida(novo_coletor, tmp_path):
    driver = DriverFalso()
    c = novo_coletor(driver)

    r = c.capturar(ID_A, str(tmp_path))

    assert driver.urls == [f"https://lattes.cnpq.br/{ID_A}"]
    assert r.status == coletor.STATUS_OK
    assert r.nome == "Maria da Silva"
    assert r.ultima_atualizacao == date(2024, 3, 15)
    assert r.arquivo_print == os.path.join(str(tmp_path), f"{ID_A}.png")
    assert os.path.isfile(r.arquivo_print)
    assert r.observacao == ""


def test_curriculo_sem_data(novo_coletor, tmp_path):
    r = novo_coletor(DriverFalso(texto="Currículo sem a linha de atualização")).capturar(ID_A, str(tmp_path))
    assert r.status == coletor.STATUS_SEM_DATA
    assert r.ultima_atualizacao is None
    assert os.path.isfile(r.arquivo_print)


def test_captcha_nao_resolvido_a_tempo(novo_coletor, tmp_path):
    c = novo_coletor(DriverFalso(paginas=["captcha"]), tempo_max_captcha=1)
    r = c.capturar(ID_A, str(tmp_path))
    assert r.status == coletor.STATUS_CAPTCHA
    assert r.observacao == "O currículo não abriu em 1 s"
    assert r.arquivo_print == ""
    assert not os.listdir(tmp_path)
    assert c.captcha_reconhecido


def test_resolver_o_captcha_abre_o_curriculo(novo_coletor, tmp_path):
    c = novo_coletor(DriverFalso(paginas=["captcha", "captcha", "carregando", "curriculo"]),
                     tempo_max_captcha=30)
    r = c.capturar(ID_A, str(tmp_path))
    assert r.status == coletor.STATUS_OK
    assert r.nome == "Maria da Silva"
    assert c.captcha_reconhecido


def test_cancelar_durante_a_espera(novo_coletor, cancelado, tmp_path):
    c = novo_coletor(DriverFalso(paginas=["captcha"]), tempo_max_captcha=60)
    cancelado.set()
    with pytest.raises(coletor.CapturaCancelada):
        c.capturar(ID_A, str(tmp_path))


# --- páginas que não são nem o currículo nem o CAPTCHA ---------------------------

def test_pagina_desconhecida_antes_de_reconhecer_o_captcha_espera_o_tempo_todo(novo_coletor, tmp_path):
    # Sem ter visto o CAPTCHA nesta captura, não dá para saber se a página é um
    # CAPTCHA que o seletor não reconhece: espera o tempo todo, como antes.
    driver = DriverFalso(paginas=["outra"], titulo="Erro", texto_outra="Currículo não encontrado")
    c = novo_coletor(driver, tempo_max_captcha=2, tempo_pagina_inesperada=0)
    inicio = time.monotonic()
    r = c.capturar(ID_A, str(tmp_path))
    assert time.monotonic() - inicio >= 2
    assert r.status == coletor.STATUS_CAPTCHA
    assert r.observacao == "O currículo não abriu em 2 s (página: Erro — Currículo não encontrado)"


def test_pagina_inesperada_depois_de_reconhecer_o_captcha_desiste_antes(novo_coletor, tmp_path):
    driver = DriverFalso(paginas=["captcha", "curriculo"])
    c = novo_coletor(driver, tempo_max_captcha=60, tempo_pagina_inesperada=1)
    assert c.capturar(ID_A, str(tmp_path)).status == coletor.STATUS_OK

    driver.paginas = ["outra"]
    driver.title = "Erro"
    inicio = time.monotonic()
    r = c.capturar(ID_B, str(tmp_path))
    assert time.monotonic() - inicio < 10           # e não os 60 s do CAPTCHA
    assert r.status == coletor.STATUS_PAGINA_INESPERADA
    assert r.observacao == "A página não mostrou o currículo nem o CAPTCHA: Erro — Currículo não encontrado"
    assert r.arquivo_print == ""


def test_pagina_que_nao_termina_de_carregar(novo_coletor, tmp_path):
    r = novo_coletor(DriverFalso(paginas=["carregando"]), tempo_max_captcha=1).capturar(ID_A, str(tmp_path))
    assert r.status == coletor.STATUS_ERRO
    assert r.observacao == "A página não terminou de carregar em 1 s"


def test_pagina_trocando_no_meio_da_verificacao(novo_coletor, tmp_path):
    class Trocando(DriverFalso):
        def execute_script(self, script):
            if "readyState" in script and self.verificacoes == 0:
                self.verificacoes += 1
                raise JavascriptException("javascript error: document unloaded while waiting for result")
            return super().execute_script(script)

    r = novo_coletor(Trocando(paginas=["captcha", "curriculo"]), tempo_max_captcha=30).capturar(
        ID_A, str(tmp_path))
    assert r.status == coletor.STATUS_OK


def test_descrever_pagina():
    assert coletor.descrever_pagina(DriverFalso(paginas=["outra"], titulo="", texto_outra="")) == (
        "página em branco")
    longa = coletor.descrever_pagina(DriverFalso(paginas=["outra"], titulo="Título",
                                                 texto_outra="palavra " * 100), limite=40)
    assert len(longa) == 40
    assert longa.startswith("Título — palavra palavra") and longa.endswith("…")


def test_erro_ao_abrir_a_pagina_vira_registro_de_erro(novo_coletor, tmp_path):
    driver = DriverFalso(erro_ao_abrir=WebDriverException("net::ERR_NAME_NOT_RESOLVED"))
    r = novo_coletor(driver).capturar(ID_A, str(tmp_path))
    assert r.status == coletor.STATUS_ERRO
    assert r.observacao == "net::ERR_NAME_NOT_RESOLVED"


def test_erro_sem_mensagem_nao_derruba_o_lote(novo_coletor, tmp_path):
    r = novo_coletor(DriverFalso(erro_ao_abrir=WebDriverException(" "))).capturar(ID_A, str(tmp_path))
    assert r.status == coletor.STATUS_ERRO
    assert r.observacao == "WebDriverException"


def test_abre_outro_navegador_se_o_usuario_fechou_o_anterior(novo_coletor, tmp_path):
    primeiro, segundo = DriverFalso(), DriverFalso()
    c = novo_coletor(primeiro, segundo)
    assert c.capturar(ID_A, str(tmp_path)).status == coletor.STATUS_OK

    class Fechado:
        @property
        def current_window_handle(self):
            raise WebDriverException("chrome not reachable")

        def quit(self):
            pass

    c.driver = Fechado()
    assert c.capturar(ID_A, str(tmp_path)).status == coletor.STATUS_OK
    assert c.driver is segundo
    assert any("navegador foi fechado" in m for m in c.mensagens)


def test_volta_para_outra_aba_se_a_aba_usada_foi_fechada(novo_coletor, tmp_path):
    driver = DriverFalso()
    c = novo_coletor(driver)
    c.capturar(ID_A, str(tmp_path))

    trocas = []

    class SemAba(DriverFalso):
        @property
        def current_window_handle(self):
            raise NoSuchWindowException("no such window")

    c.driver = SemAba()
    c.driver.window_handles = ["aba-2"]
    c.driver.switch_to = SimpleNamespace(window=trocas.append)
    assert c.capturar(ID_A, str(tmp_path)).status == coletor.STATUS_OK
    assert trocas == ["aba-2"]


def test_sem_chrome_levanta_erro_de_navegador(cancelado):
    c = coletor.ColetorLattes(cancelado, log=lambda _: None)

    def falha():
        raise WebDriverException("unable to obtain driver for chrome")

    c._criar_driver = falha
    with pytest.raises(coletor.ErroNavegador, match="unable to obtain driver"):
        c.capturar(ID_A, "")


def test_fechar_pode_ser_chamado_varias_vezes(novo_coletor, tmp_path):
    driver = DriverFalso()
    c = novo_coletor(driver)
    c.capturar(ID_A, str(tmp_path))
    c.fechar()
    c.fechar()
    assert driver.fechado
    assert c.driver is None
