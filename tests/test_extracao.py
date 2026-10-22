"""Extração da data de atualização e resumo das mensagens de erro."""

from datetime import date

import pytest
from selenium.common.exceptions import (
    InvalidSessionIdException,
    NoSuchWindowException,
    TimeoutException,
)

import coletor


@pytest.mark.parametrize("texto", [
    "Última atualização do currículo em 15/03/2024",
    "ultima atualizacao do curriculo em 15/03/2024",
    "ÚLTIMA ATUALIZAÇÃO DO CURRÍCULO EM 15/03/2024",
    "Endereço para acessar este CV: ...\nÚltima atualização do currículo em\n15/03/2024\nResumo",
])
def test_extrai_a_data_de_atualizacao(texto):
    assert coletor.extrair_data_atualizacao(texto) == date(2024, 3, 15)


@pytest.mark.parametrize("texto", [
    "",
    "Currículo sem a linha de atualização",
    "Última atualização do currículo em 31/02/2024",  # data impossível
])
def test_sem_data_valida_retorna_none(texto):
    assert coletor.extrair_data_atualizacao(texto) is None


def test_resume_mensagem_do_selenium():
    erro = TimeoutException("timeout: Timed out receiving message from renderer: 59.5\n"
                            "  (Session info: chrome=141.0)\nStacktrace:\n#0 0x55d...")
    assert coletor.resumir_erro(erro) == "timeout: Timed out receiving message from renderer: 59.5"


def test_remove_o_link_de_documentacao():
    erro = Exception("no such element: Unable to locate element; For documentation on this error, "
                     "please visit: https://www.selenium.dev/documentation/...")
    assert coletor.resumir_erro(erro) == "no such element: Unable to locate element"


@pytest.mark.parametrize("classe", [InvalidSessionIdException, NoSuchWindowException])
def test_navegador_fechado_tem_mensagem_amigavel(classe):
    assert coletor.resumir_erro(classe("invalid session id")) == "O navegador foi fechado durante a captura"


def test_erro_sem_mensagem_usa_o_nome_da_classe():
    assert coletor.resumir_erro(KeyError()) == "KeyError"


def test_mensagem_longa_e_cortada():
    assert len(coletor.resumir_erro(Exception("x" * 500))) == 200
