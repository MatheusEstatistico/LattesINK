"""Onde o programa guarda a configuração (última pasta usada etc.)."""

import json
import os

import pytest


pytest.importorskip("tkinter")

import LattesINK


@pytest.fixture
def caminhos(monkeypatch, tmp_path):
    novo = tmp_path / "config" / "config.json"
    antigo = tmp_path / ".lattesink.json"
    monkeypatch.setattr(LattesINK, "ARQUIVO_CONFIG", str(novo))
    monkeypatch.setattr(LattesINK, "ARQUIVO_CONFIG_ANTIGO", str(antigo))
    return novo, antigo


@pytest.mark.parametrize(("plataforma", "variaveis", "esperado"), [
    ("win32", {"APPDATA": os.path.join("C:", "Users", "ana", "AppData", "Roaming")},
     os.path.join("C:", "Users", "ana", "AppData", "Roaming", "LattesINK")),
    ("linux", {"XDG_CONFIG_HOME": os.path.join("/home", "ana", ".cfg")},
     os.path.join("/home", "ana", ".cfg", "lattesink")),
    ("linux", {}, os.path.join("~", ".config", "lattesink")),
    ("darwin", {}, os.path.join("~", "Library", "Application Support", "LattesINK")),
])
def test_pasta_padrao_de_cada_sistema(monkeypatch, plataforma, variaveis, esperado):
    monkeypatch.setattr(LattesINK.sys, "platform", plataforma)
    monkeypatch.setattr(LattesINK.os.path, "expanduser", lambda caminho: caminho)
    for nome in ("APPDATA", "XDG_CONFIG_HOME"):
        monkeypatch.delenv(nome, raising=False)
    for nome, valor in variaveis.items():
        monkeypatch.setenv(nome, valor)
    assert LattesINK.pasta_de_configuracao() == esperado


def test_salvar_e_carregar(caminhos):
    novo, _ = caminhos
    assert LattesINK.salvar_config({"pasta_destino": "D:/Capturas"})  # cria a pasta se preciso
    assert json.loads(novo.read_text(encoding="utf-8")) == {"pasta_destino": "D:/Capturas"}
    assert LattesINK.carregar_config() == {"pasta_destino": "D:/Capturas"}


def test_migra_a_configuracao_da_versao_anterior(caminhos):
    novo, antigo = caminhos
    antigo.write_text(json.dumps({"pasta_ids": "C:/ids"}), encoding="utf-8")

    assert LattesINK.carregar_config() == {"pasta_ids": "C:/ids"}
    assert json.loads(novo.read_text(encoding="utf-8")) == {"pasta_ids": "C:/ids"}
    assert not antigo.exists()


def test_configuracao_corrompida_e_ignorada(caminhos):
    novo, _ = caminhos
    novo.parent.mkdir()
    novo.write_text("{ isso não é json", encoding="utf-8")
    assert LattesINK.carregar_config() == {}
