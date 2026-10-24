# LattesINK
Repositório dedicado ao desenvolvimento de uma aplicação simples para verificações de currículos Lattes para fins acadêmicos, sem nenhum propósito comercial ou lucrativo.

## Pyinstaller
Para gerar o executável, use a receita `LattesINK.spec` com o [*PyInstaller*](https://pyinstaller.org/en/stable/index.html), de preferência num ambiente virtual limpo:

```
python -m pip install ".[build]"
pyinstaller --noconfirm LattesINK.spec
```

No Windows sai o `dist/LattesINK.exe`; no Linux, a pasta `dist/LattesINK/`, que vira um pacote `.deb` com `bash packaging/linux/build_deb.sh`. O GitHub Actions faz isso a cada PR (os arquivos ficam em "Artifacts", na página da execução) e anexa os arquivos a cada release publicada.

## Antivirus 
É comum que antivirus testem falso positivo com arquivos desse tipo. Caso isso aconteça, pode-se autorizar o aplicativo nas configurações do antivirus. 

## reCAPTCHA
Esse aplicativo não pode passar pelo reCAPTCHA, então o usuário tem que resolver manualmente. Caso o usuário tenha acesso ao serviço [*Extrator Lattes*](https://www.gov.br/pt-br/servicos/obter-acesso-ao-extrator-da-plataforma-lattes), do próprio CNPq, é possível solucionar isso. O Extrator Lattes foi criado pela [*Fábrica de Software da UNIVILLE*](https://github.com/FabricadeSoftwareUNIVILLE/ExtratorLattesCNPq). 

## 

