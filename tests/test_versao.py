"""scripts/versao.py: versão do programa e notas de release para o GitHub Actions."""

import sys
from pathlib import Path

import pytest


sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

import versao


CHANGELOG = """# Changelog

## [Não lançado]

- algo em andamento

## [1.1.0] - 2026-11-01

### Corrigido
- coisa um

## [1.0.0] - 2026-09-22

- primeira versão

[1.1.0]: https://github.com/MatheusEstatistico/LattesINK/compare/v1.0.0...v1.1.0
"""


def test_le_a_versao_do_programa():
    pytest.importorskip("tkinter")
    import LattesINK

    assert versao.ler_versao() == LattesINK.__version__


def test_notas_da_versao(tmp_path):
    (tmp_path / "CHANGELOG.md").write_text(CHANGELOG, encoding="utf-8")
    assert versao.notas_da_versao("1.1.0", tmp_path) == "### Corrigido\n- coisa um\n"
    assert versao.notas_da_versao("1.0.0", tmp_path) == "- primeira versão\n"
    assert versao.notas_da_versao("2.0.0", tmp_path) is None


def test_sem_changelog(tmp_path):
    assert versao.notas_da_versao("1.0.0", tmp_path) is None
