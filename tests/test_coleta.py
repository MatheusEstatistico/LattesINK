"""ColetorLattes com um navegador falso: nenhum teste abre o Chrome."""

import os
from datetime import date
from types import SimpleNamespace

import pytest
from falsos import ID_A, DriverFalso
from selenium.common.exceptions import NoSuchWindowException, WebDriverException

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
    r = novo_coletor(DriverFalso(nome=None), tempo_max_captcha=1).capturar(ID_A, str(tmp_path))
    assert r.status == coletor.STATUS_CAPTCHA
    assert r.arquivo_print == ""
    assert not os.listdir(tmp_path)


def test_cancelar_durante_a_espera(novo_coletor, cancelado, tmp_path):
    c = novo_coletor(DriverFalso(nome=None), tempo_max_captcha=60)
    cancelado.set()
    with pytest.raises(coletor.CapturaCancelada):
        c.capturar(ID_A, str(tmp_path))


def test_erro_ao_abrir_a_pagina_vira_registro_de_erro(novo_coletor, tmp_path):
    driver = DriverFalso(erro_ao_abrir=WebDriverException("net::ERR_NAME_NOT_RESOLVED"))
    r = novo_coletor(driver).capturar(ID_A, str(tmp_path))
    assert r.status == coletor.STATUS_ERRO
    assert r.observacao == "net::ERR_NAME_NOT_RESOLVED"


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
