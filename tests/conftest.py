"""Fixtures compartilhadas pelos testes."""

import threading

import pytest

import coletor


@pytest.fixture
def cancelado():
    return threading.Event()


@pytest.fixture
def novo_coletor(cancelado):
    """Cria um ColetorLattes que usa os navegadores falsos passados, em ordem."""

    def criar(*drivers, tempo_max_captcha=1, tempo_pagina_inesperada=30):
        mensagens = []
        fila = list(drivers)
        c = coletor.ColetorLattes(cancelado, log=mensagens.append, tempo_max_captcha=tempo_max_captcha,
                                  tempo_pagina_inesperada=tempo_pagina_inesperada)
        c._criar_driver = lambda: fila.pop(0)
        c.mensagens = mensagens
        return c

    return criar
