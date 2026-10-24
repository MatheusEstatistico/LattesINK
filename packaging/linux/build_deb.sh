#!/usr/bin/env bash
# Monta o pacote .deb a partir de dist/LattesINK/, gerado no Linux por
#   pyinstaller --noconfirm LattesINK.spec
#
# Uso: bash packaging/linux/build_deb.sh [pasta de saída]   (padrão: dist)
# Imprime o caminho do .deb gerado.
set -euo pipefail

raiz="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
saida="$(realpath -m "${1:-$raiz/dist}")"
programa="$raiz/dist/LattesINK"

if [[ ! -x "$programa/LattesINK" ]]; then
    echo "Não encontrei $programa/LattesINK. Rode antes: pyinstaller --noconfirm LattesINK.spec" >&2
    exit 1
fi

versao="$(python3 "$raiz/scripts/versao.py")"
arquitetura="$(dpkg --print-architecture)"
nome="lattesink_${versao}_${arquitetura}"

temporaria="$(mktemp -d)"
trap 'rm -rf "$temporaria"' EXIT
pacote="$temporaria/$nome"

# Programa em /opt, comando "lattesink" no PATH e atalho no menu de aplicativos
install -d "$pacote/DEBIAN" "$pacote/opt" "$pacote/usr/bin" "$pacote/usr/share/doc/lattesink"
cp -a "$programa" "$pacote/opt/lattesink"
ln -s /opt/lattesink/LattesINK "$pacote/usr/bin/lattesink"
install -D -m 644 "$raiz/packaging/linux/lattesink.desktop" "$pacote/usr/share/applications/lattesink.desktop"
if [[ -f "$raiz/assets/icone.png" ]]; then
    install -D -m 644 "$raiz/assets/icone.png" "$pacote/usr/share/pixmaps/lattesink.png"
fi
install -m 644 "$raiz/LICENSE" "$pacote/usr/share/doc/lattesink/copyright"

cat > "$pacote/DEBIAN/control" <<CONTROLE
Package: lattesink
Version: $versao
Architecture: $arquitetura
Maintainer: Matheus Fernandes <14061352+MatheusEstatistico@users.noreply.github.com>
Installed-Size: $(du -sk --exclude=DEBIAN "$pacote" | cut -f1)
Section: science
Priority: optional
Homepage: https://github.com/MatheusEstatistico/LattesINK
Recommends: google-chrome-stable
Description: verificador de currículos Lattes
 Abre no Chrome o currículo Lattes de cada ID de uma lista (o CAPTCHA é
 resolvido por quem está usando), captura o nome do pesquisador, a data da
 última atualização e um print da página, e gera uma planilha Excel.
 .
 Sem o Google Chrome instalado, o Selenium baixa uma versão de testes do
 Chrome na primeira captura.
CONTROLE

mkdir -p "$saida"
dpkg-deb --root-owner-group --build "$pacote" "$saida/$nome.deb" >/dev/null
echo "$saida/$nome.deb"
