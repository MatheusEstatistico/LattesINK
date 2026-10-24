"""
Versão do programa e notas de release, para o build e o GitHub Actions.

    python scripts/versao.py           -> imprime a versão (ex.: 1.1.0)
    python scripts/versao.py --notas   -> imprime a seção dessa versão no CHANGELOG.md
                                          (sai com erro se não houver)

A versão vem de __version__ em LattesINK.py, sem importar o programa
(que precisaria do Tkinter e das dependências instaladas).
"""

import ast
import os
import re
import sys


RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def ler_versao(raiz=RAIZ):
    with open(os.path.join(raiz, "LattesINK.py"), encoding="utf-8") as f:
        arvore = ast.parse(f.read())
    for no in arvore.body:
        if (isinstance(no, ast.Assign) and isinstance(no.value, ast.Constant)
                and any(isinstance(alvo, ast.Name) and alvo.id == "__version__" for alvo in no.targets)):
            return no.value.value
    raise ValueError("__version__ não encontrado em LattesINK.py")


def notas_da_versao(versao, raiz=RAIZ):
    """Texto da seção "## [versao]" do CHANGELOG.md (formato Keep a Changelog), ou None."""
    caminho = os.path.join(raiz, "CHANGELOG.md")
    if not os.path.isfile(caminho):
        return None
    with open(caminho, encoding="utf-8") as f:
        texto = f.read()
    secao = re.search(rf"^## \[?{re.escape(versao)}\]?.*?$(.*?)(?=^## |^\[[^\]]+\]: |\Z)",
                      texto, re.MULTILINE | re.DOTALL)
    if not secao or not secao.group(1).strip():
        return None
    return secao.group(1).strip() + "\n"


def main(argumentos):
    versao = ler_versao()
    if argumentos == ["--notas"]:
        notas = notas_da_versao(versao)
        if notas is None:
            print(f"CHANGELOG.md não tem uma seção para a versão {versao}", file=sys.stderr)
            return 1
        sys.stdout.write(notas)
    elif not argumentos:
        print(versao)
    else:
        print(__doc__, file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
