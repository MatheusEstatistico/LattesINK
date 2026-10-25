# Histórico de mudanças

As mudanças relevantes do LattesINK ficam registradas aqui. O formato segue o
[Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/), e as versões seguem o
[Versionamento Semântico](https://semver.org/lang/pt-BR/).

## [Não lançado]

### Adicionado

- Pacote `.deb` para Ubuntu, Debian e derivados, e um `.tar.gz` para outras distribuições Linux ([#6](https://github.com/MatheusEstatistico/LattesINK/issues/6)).
- Botão **Repetir falhas**, que captura de novo só os currículos que não deram certo.
- Botão **Abrir pasta**, que abre a pasta com os resultados.
- A captura é pausada depois de 3 currículos seguidos que não abriram (CAPTCHAs sem ninguém resolvendo, site fora do ar ou com erro), em vez de seguir pela lista inteira.
- Status **Página inesperada** para páginas que não são nem o currículo nem o CAPTCHA (ID inexistente, erro do site), com o título e o começo do texto da página na coluna Observação. Elas são abandonadas depois de 30 segundos.
- Erros inesperados aparecem numa mensagem e ficam registrados, com os detalhes técnicos, em `lattesink.log`.
- Versão do programa no título da janela e nas propriedades do `.exe`.
- Testes automatizados e verificação contínua (GitHub Actions) no Windows e no Linux.
- Geração automática do `.exe`, do `.deb` e do `.tar.gz` em cada PR, anexados automaticamente a cada release publicada.
- `CITATION.cff`, para citar o programa em trabalhos acadêmicos.

### Alterado

- O `.exe` ficou bem menor: ele levava numpy, lxml, o suporte a imagens AVIF e o selenium-manager de outros sistemas, nada disso usado pelo programa.
- Texto nítido em telas com escala de 125%, 150% etc. no Windows. A janela agora pode ser redimensionada.
- A configuração foi para a pasta padrão do sistema (no Windows, `%APPDATA%\LattesINK`). A configuração antiga é migrada automaticamente.
- A data de atualização é lida muito mais rápido em currículos longos.
- A planilha sai sempre na ordem do arquivo de IDs.
- Botões desativados ficam cinza.
- Quando o Chrome não abre, a mensagem sugere verificar se ele está instalado.

### Corrigido

- Um erro sem mensagem interrompia a captura de todos os IDs restantes.
- Arquivos `.txt` em UTF-16 (o "Texto Unicode" do Excel e o PowerShell gravam assim) não tinham nenhum ID reconhecido.
- Só o primeiro ID de cada linha era lido; os outros sumiam sem aviso.
- Só a aba ativa do `.xlsx` era lida.
- Página de erro do site era registrada como "CAPTCHA não resolvido", depois de 3 minutos de espera.

## [1.0.0] - 2026-09-22

Primeira versão publicada: interface gráfica para capturar o nome, a data da última atualização e um print com data e hora de uma lista de currículos Lattes (`.txt`, `.csv` ou `.xlsx`), gerando uma planilha formatada, um CSV de backup e um log. A captura pode ser cancelada e retomada.

[Não lançado]: https://github.com/MatheusEstatistico/LattesINK/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/MatheusEstatistico/LattesINK/releases/tag/v1.0.0
