"""
Receita do PyInstaller para gerar o executável do LattesINK:

    python -m pip install ".[build]"
    pyinstaller --noconfirm LattesINK.spec

- Windows: um único arquivo, dist/LattesINK.exe
- Linux: a pasta dist/LattesINK/ (é ela que vai dentro do pacote .deb)

Ícone: coloque a imagem em assets/icone.png (de preferência quadrada, com
256x256 ou mais). O .ico do Windows, com todos os tamanhos, é gerado a
partir dela. Sem a imagem, o build funciona e o programa fica com o ícone
genérico do sistema.

Rode o build num ambiente virtual limpo: pacotes que estiverem instalados
por acaso (numpy, lxml...) podem acabar dentro do executável.
"""

import os
import platform
import re
import sys

from PIL import Image


RAIZ = os.path.abspath(SPECPATH)  # noqa: F821 - variáveis definidas pelo PyInstaller
PASTA_TEMPORARIA = os.path.abspath(workpath)  # noqa: F821
sys.path.insert(0, os.path.join(RAIZ, "scripts"))
from versao import ler_versao  # noqa: E402


VERSAO = ler_versao(RAIZ)
WINDOWS = sys.platform == "win32"


# --- Ícone ---------------------------------------------------------------------

dados = []
icone_do_exe = None
icone_png = os.path.join(RAIZ, "assets", "icone.png")
icone_antigo = next((caminho for caminho in (os.path.join(RAIZ, "assets", "lattes.ico"),
                                              os.path.join(RAIZ, "lattes.ico"))
                     if os.path.isfile(caminho)), None)

if os.path.isfile(icone_png):
    dados.append((icone_png, "assets"))
    if WINDOWS:
        os.makedirs(PASTA_TEMPORARIA, exist_ok=True)
        icone_do_exe = os.path.join(PASTA_TEMPORARIA, "icone.ico")
        with Image.open(icone_png) as imagem:
            imagem.save(icone_do_exe, sizes=[(n, n) for n in (16, 24, 32, 48, 64, 128, 256)])
elif icone_antigo:
    dados.append((icone_antigo, "."))
    icone_do_exe = icone_antigo
else:
    print("AVISO: assets/icone.png não encontrado; o executável vai sair sem ícone próprio.")


# --- O que não precisa ir no executável ------------------------------------------

EXCLUIR = [
    "numpy",      # o openpyxl usa se estiver instalado, mas o programa não precisa (~10 MB)
    "lxml",       # idem (~3,5 MB); sem ele o openpyxl usa o XML da biblioteca padrão
    "PIL._avif",  # suporte do Pillow a imagens AVIF (~4 MB no Windows)
]

# O Selenium traz o selenium-manager (que encontra o Chrome e o chromedriver)
# de todas as plataformas. O executável só precisa do da plataforma do build.
if WINDOWS:
    PLATAFORMAS_DO_SELENIUM = {"windows"}
elif sys.platform == "darwin":
    PLATAFORMAS_DO_SELENIUM = {"macos"}
elif platform.machine().lower() in ("aarch64", "arm64"):
    PLATAFORMAS_DO_SELENIUM = {"linux-arm64"}
else:
    PLATAFORMAS_DO_SELENIUM = {"linux-x86_64", "linux"}  # "linux" nas versões antigas do Selenium


def de_outra_plataforma(destino):
    """True para selenium/webdriver/common/<outra plataforma>/selenium-manager[.exe]."""
    partes = destino.replace("\\", "/").split("/")
    return (len(partes) == 5 and partes[:3] == ["selenium", "webdriver", "common"]
            and partes[4].startswith("selenium-manager") and partes[3] not in PLATAFORMAS_DO_SELENIUM)


# --- Build -----------------------------------------------------------------------

a = Analysis(  # noqa: F821
    [os.path.join(RAIZ, "LattesINK.py")],
    pathex=[RAIZ],
    datas=dados,
    excludes=EXCLUIR,
)
a.datas = [item for item in a.datas if not de_outra_plataforma(item[0])]
a.binaries = [item for item in a.binaries if not de_outra_plataforma(item[0])]

pyz = PYZ(a.pure)  # noqa: F821

if WINDOWS:
    from PyInstaller.utils.win32.versioninfo import (
        FixedFileInfo,
        StringFileInfo,
        StringStruct,
        StringTable,
        VarFileInfo,
        VarStruct,
        VSVersionInfo,
    )

    # Propriedades do .exe (botão direito > Propriedades > Detalhes)
    numeros = tuple(int(n) for n in (re.findall(r"\d+", VERSAO) + ["0"] * 4)[:4])
    informacoes = VSVersionInfo(
        ffi=FixedFileInfo(filevers=numeros, prodvers=numeros),
        kids=[
            StringFileInfo([StringTable("041604B0", [  # português do Brasil, Unicode
                StringStruct("CompanyName", "Matheus Fernandes"),
                StringStruct("FileDescription", "LattesINK - verificador de currículos Lattes"),
                StringStruct("FileVersion", VERSAO),
                StringStruct("InternalName", "LattesINK"),
                StringStruct("LegalCopyright", "Licença GPL-3.0"),
                StringStruct("OriginalFilename", "LattesINK.exe"),
                StringStruct("ProductName", "LattesINK"),
                StringStruct("ProductVersion", VERSAO),
            ])]),
            VarFileInfo([VarStruct("Translation", [0x0416, 1200])]),
        ],
    )

    exe = EXE(  # noqa: F821
        pyz, a.scripts, a.binaries, a.datas, [],
        name="LattesINK",
        console=False,
        upx=False,  # compactar com UPX aumenta os falsos positivos de antivírus
        icon=icone_do_exe or "NONE",  # "NONE": ícone genérico em vez do ícone do PyInstaller
        version=informacoes,
    )
else:
    exe = EXE(  # noqa: F821
        pyz, a.scripts, [],
        exclude_binaries=True,
        name="LattesINK",
        console=False,
        upx=False,
    )
    coll = COLLECT(exe, a.binaries, a.datas, name="LattesINK", upx=False)  # noqa: F821
