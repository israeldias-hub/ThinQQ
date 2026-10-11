# ThinQQ

Suplemento do Excel (VBA, Windows) para criar gráficos de apresentação no estilo think-cell: colunas e barras
(empilhadas, agrupadas e 100%), cascata (ponte, com várias séries), linha, área, combinação, Mekko, Gantt, pizza e
rosca, com totais, rótulos vivos, conectores, setas de CAGR e de diferença, linha de valor, indicador 100% e mesma
escala. Os gráficos gerados são gráficos nativos do Excel, ligados aos dados por fórmulas: quem recebe o arquivo
não precisa ter o ThinQQ instalado.

Repositório: <https://github.com/israeldias-hub/ThinQQ>

## Recursos

**Fase 1 (1.0.0)**

- Aba própria na faixa de opções, com opções que mudam conforme o gráfico selecionado.
- Colunas empilhadas (com totais), agrupadas, 100% e linha; cascata com marcador `e` para totais e
  subtotais calculados e barras que cruzam o zero.
- Rótulos ligados às células, conectores entre colunas ou barras, paletas (Padrão ThinQQ, Tons de cinza,
  Cores do tema), formatos numéricos (inteiro, decimais, percentual, milhares, milhões) e espaço entre colunas.
- Setas de CAGR e de diferença que se reposicionam quando os dados ou o tamanho do gráfico mudam.
- Pasta de exemplos (ThinQQ > Exemplos) e tutorial (ThinQQ > Tutorial).
- Instalação por usuário, sem administrador. Atualização dentro do Excel, por manifesto (HTTP ou pasta),
  com conferência SHA-256 e troca do núcleo sem reiniciar o Excel.

**Fase 2 (1.1.0)**

- Faixa de opções em menus por família (Colunas, Barras, Cascata, Linha e área, Mais tipos) e, no grupo do gráfico
  selecionado, os menus Mais opções e Setas. Os controles só ficam ativos nos tipos que os aceitam.
- Barras horizontais (empilhadas, agrupadas e 100%) com totais no fim das barras, e Cascata (barras).
- Cascata com várias séries (até 6, com legenda) e "de cima para baixo" (`e` na primeira linha e um número na última).
- Área empilhada e 100%; Combinação (colunas empilhadas + última coluna como linha, com eixo secundário opcional);
  Legenda (Direita, Nenhuma, Nome no fim da linha).
- Mekko (%) e Mekko (unidades), Gantt (escala automática de semanas a anos, marcos, colchetes), Pizza e Rosca.
- Linha de valor (média), Indicador 100% (ligado por padrão nos 100% e no Mekko (%)), Mesma escala e Escala automática.
- 18 exemplos em duas planilhas (ThinQQ > Exemplos) e tutorial com uma seção por tipo.

### Galeria de gráficos (1.1.0)

| Menu | Tipos | Tabela de origem |
|---|---|---|
| Colunas | Colunas empilhadas, Colunas agrupadas, Colunas 100% | cabeçalho, categorias na 1ª coluna e uma coluna de valores por série |
| Barras | Barras empilhadas, Barras agrupadas, Barras 100% | a mesma das colunas |
| Cascata | Cascata (colunas), Cascata (barras) | etapa e 1 a 6 colunas de valores; `e` marca um total |
| Linha e área | Linha, Área empilhada, Área 100%, Combinação (colunas + linha) | a mesma das colunas; a combinação precisa de 2 ou mais colunas de valores |
| Mais tipos | Mekko (%), Mekko (unidades), Gantt, Pizza, Rosca | Mekko (unidades): 2ª coluna = largura; Gantt: atividade, início, fim; pizza e rosca: categorias e uma coluna de valores |

O tutorial (`docs/tutorial/tutorial.html`, aberto por ThinQQ > Tutorial) traz a tabela de dados de cada tipo, a tabela
do que cada tipo aceita, as mensagens do ThinQQ e os limites conhecidos.

## Antes de distribuir (checklist do responsável pela publicação)

**Esta lista é de ações suas, fora do código** (o repositório ainda não foi publicado; o `push` falhou por falta de credenciais). A primeira versão a ser publicada é a **1.1.0**: a 1.0.0 (Fase 1) nunca saiu.

- [ ] Publicar o conteúdo do repositório no GitHub (`israeldias-hub/ThinQQ`).
- [ ] Gerar os arquivos com `build/release.ps1` e publicar a release **v1.1.0** com os assets (ZIP, `ThinQQ.xlam`, `ThinQQ.core.xlam`, `manifest.json`, `tutorial.html`).
- [ ] Fazer commit de `release/manifest.json` (com URLs reais) e de `version.json` (já em `{ "loader": "1.1.0", "core": "1.1.0" }`) na `main`.
- [ ] **Privacidade antes de publicar** (o repositório ou o Pages): os arquivos versionados já foram limpos (caminhos de perfil do Windows, ids de máquina, PIDs do Excel do autor e e-mail corporativo) e `tests/ps/Test-Privacy.ps1` (rodado por `build/run-tests.ps1`) falha se algum voltar. O que sobra é o **histórico do Git**: o autor e o e-mail dos commits e o texto antigo dos arquivos continuam nos commits antigos. Para publicar só o estado atual, crie um repositório novo a partir da árvore de trabalho (`git archive HEAD`); para manter o histórico, reescreva-o antes de publicar (`git filter-repo --mailmap <arquivo>` troca nome e e-mail dos autores; `git filter-repo --replace-text <arquivo>` troca o texto antigo; instale o `git-filter-repo`, rode numa cópia e confira com `git log --format='%an <%ae>'` e `git grep` no histórico).
- [ ] Habilitar o GitHub Pages para o tutorial online (veja "Tutorial online (GitHub Pages)": pela pasta `/docs`, tudo o que está em `docs/` fica público, inclusive a especificação e o plano em `docs/superpowers`).
- [ ] Verificar a verificação de atualização a partir de um perfil limpo (usuário do Windows novo ou máquina sem o ThinQQ): instalar uma versão anterior (como a 1.1.0 é a primeira publicada, use um ensaio local com uma versão menor) e conferir que o aviso de atualização aparece e atualiza.
- [ ] Ativar autenticação em dois fatores na conta do GitHub (veja Segurança abaixo).
- [ ] Rodar o roteiro manual no Excel real (seção "Verificações manuais no Excel real"), a começar pela conferência de que a aba ThinQQ aparece.

## Segurança e confiança

- **A assinatura de código (certificado no projeto VBA) é da Fase 3.** Até lá, a integridade das atualizações vem só da conferência SHA-256 contra o manifesto.
- **Quem controla a branch `main` do repositório e os assets das releases controla o código que roda no Excel dos colegas** (o manifesto padrão vem de lá). Proteja a conta com autenticação em dois fatores e limite quem pode publicar; em empresas, use uma fonte de atualização interna.
- O SHA-256 é calculado em VBA puro (módulo `modSha256`), sem iniciar processos externos (nada de `certutil`).
- Downloads ficam em `%TEMP%\ThinQQ`, fora do local confiável; só depois de conferidos vão para `%LOCALAPPDATA%\ThinQQ`.

### O que o instalador altera

- Arquivos em `%LOCALAPPDATA%\ThinQQ` (loader na raiz, core e tutorial em `versions\<versão>`).
- `HKCU\Software\Microsoft\Office\16.0\Excel\Options`: valor `OPEN` (ou `OPEN1`, `OPEN2`...) que registra o suplemento.
- `HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ`: local confiável para `%LOCALAPPDATA%\ThinQQ\` **com subpastas**.
- `HKCU\Software\VB and VBA Program Settings\ThinQQ`: configurações do atualizador.
- Nada em `HKLM`. O `uninstall.ps1` desfaz tudo.

### Frequência das verificações

O valor `CheckIntervalHours` (horas; padrão 24) em `HKCU\Software\VB and VBA Program Settings\ThinQQ\Updater` controla o intervalo da verificação automática. Para praticamente desligá-la, grave um número muito grande (por exemplo `876000`); o botão de verificação manual continua funcionando sob demanda.

## Instalação (usuários)

Requisitos: Excel do Office 16.0 para Windows. Os testes automatizados rodaram só no Excel 2019 de 32 bits; o 2016, o 2021 e o Microsoft 365 são esperados, não testados. O código não faz chamadas à API do Windows (nenhuma instrução `Declare`), então é compatível com 32 e 64 bits por construção.

1. Baixe `ThinQQ-x.y.z.zip` da página de [Releases](https://github.com/israeldias-hub/ThinQQ/releases) e descompacte a pasta inteira.
2. Feche o Excel e dê dois cliques em `Instalar-ThinQQ.cmd`.
3. Abra o Excel: a aba ThinQQ aparece na faixa de opções.

A instalação manual (se o script for bloqueado), a desinstalação e a solução de problemas estão em
`installer/LEIA-ME.txt` (incluída no ZIP) e no tutorial (`docs/tutorial/tutorial.html`).

## Desenvolvimento

Pré-requisitos: Windows, Excel 2016+ (os `.xlam` são montados por automação COM), PowerShell 7 (`pwsh`),
Git e Node.js (só para o servidor HTTP de fixtures dos testes do atualizador).

No Excel, habilite uma vez: Arquivo > Opções > Central de Confiabilidade > Configurações da Central de
Confiabilidade > Configurações de Macro > "Confiar no acesso ao modelo de objeto do projeto VBA"
(`AccessVBOM=1`). O build confere isso e avisa se faltar.

```powershell
pwsh build/build.ps1 -Target all -Configuration Release          # monta dist/ThinQQ.xlam e dist/ThinQQ.core.xlam
pwsh build/build.ps1 -Target all -Configuration Release -Install # + instala em %LOCALAPPDATA%\ThinQQ (Excel fechado)
pwsh build/run-tests.ps1                                          # lint, compilação, testes VBA (core e loader), faixa, instalador, documentação
pwsh build/run-tests.ps1 -Target core -NoBuild                    # só o core, reaproveitando o build de teste
pwsh build/run-tests.ps1 -LoaderSwap                              # + teste de troca do loader em Excel real
pwsh build/run-tests.ps1 -Target core -Filter "TestBars|TestGantt" -TimeoutSec 240   # só os testes cujo Modulo.Teste contém o texto (pula as suítes PowerShell)
pwsh tests/ps/Test-Docs.ps1                                       # só a consistência da documentação (tutorial, README, CHANGELOG, faixa)
pwsh tests/ps/Test-Privacy.ps1                                    # só a privacidade: nenhum arquivo versionado com caminho de perfil, id de máquina ou e-mail corporativo
```

Os testes VBA ficam em `tests/vba/core` e `tests/vba/loader` (`Test*.bas`, importados pelo build sem registro manual).
No core, um arquivo por área: `TestKinds` (matriz de capacidades), `TestStacked`, `TestBars`, `TestWaterfall` e
`TestWaterfall2` (cascata, várias séries, horizontal), `TestArea` (área, combinação, legenda), `TestPie`, `TestMekko`,
`TestGantt`, `TestDecorations` (linha de valor, indicador 100%, mesma escala), além de `TestAnnotations`,
`TestGeometry`, `TestCalcSheet`, `TestCoreApi`, `TestSamples` e outros. `tests/ps/Test-Docs.ps1` confere que os rótulos do
tutorial existem na faixa, que a tabela "O que cada tipo aceita" bate com `Kinds_Supports` e que as mensagens citadas
existem no código.

Formatos numéricos: rótulos e células de números do bloco usam `NumberFormatLocal` com `Style_LocalizeCode` (ou
`Style_AssignNumberFormat`); `Range.NumberFormat =` direto só com código invariante ou copiado da origem. Exceções aceitas na
auditoria da 1.1.0: `"@"` (texto) em `modCalcSheet` e `TestWaterfall`; `"0.0%"` em `TestArea` e `"#,##0.00000"` em `TestChartData`;
o formato copiado da categoria de origem em `modChartMekko` e os de data da escala em `modChartGantt` (`dd/mm`, `mmm/yy`).

Os testes usam um Excel invisível próprio e pastas temporárias. Feche instâncias do Excel que tenham
`dist\ThinQQ.core.test.xlam` aberto (o arquivo fica travado). O build de teste (`*.test.xlam`) nunca é
publicado: o `release.ps1` recusa artefatos com módulos de teste.

## Publicar uma versão

1. Atualize `release/notes.txt` (vai para o campo `notes` do manifesto) e o `CHANGELOG.md`.
2. Rode o release (a partir da raiz do repositório, com o Excel fechado):

   ```powershell
   pwsh build/release.ps1 -Version x.y.z -BaseUrl https://github.com/israeldias-hub/ThinQQ/releases/download/vx.y.z
   # pasta de rede ou OneDrive sincronizado:
   pwsh build/release.ps1 -Version x.y.z -BaseUrl '\\servidor\ThinQQ' -PublishTo '\\servidor\ThinQQ'
   ```

   Parâmetros: `-Version` (x.y.z, vira a versão do core), `-LoaderVersion` (opcional, só quando o loader muda),
   `-BaseUrl` (base das URLs do manifesto), `-PublishTo` (copia os arquivos para uma pasta) e `-NotesFile`.
   O script recusa o tutorial-esqueleto, grava `version.json`, monta a configuração Release e gera
   `dist/manifest.json` (com os SHA-256) e `dist/ThinQQ-x.y.z.zip`.
3. No GitHub, crie a release `vx.y.z` e anexe como assets o ZIP e os arquivos soltos de `dist/`
   (`ThinQQ.xlam`, `ThinQQ.core.xlam`, `manifest.json`, `tutorial.html`): o atualizador baixa esses arquivos individualmente.
4. Copie `dist/manifest.json` para `release/manifest.json` e faça commit junto com `version.json` na
   branch `main` (a fonte padrão de atualização é o `release/manifest.json` da `main`). Só commite o
   manifesto com URLs reais, nunca o de um teste local.
5. Os usuários recebem o aviso na próxima abertura do Excel (no máximo uma verificação por dia).

Para ensaiar sem publicar, use uma cópia do repositório (o Excel do usuário pode manter `dist\ThinQQ.core.test.xlam` travado):
`pwsh build/release.ps1 -Version x.y.z -BaseUrl C:\temp\ensaio -PublishTo C:\temp\ensaio`, confira o ZIP e o
`manifest.json` (versões e os dois SHA-256) e apague a pasta. O ensaio não grava `release/manifest.json`.

Detalhes em `release/README.md`.

### Tutorial online (GitHub Pages)

O botão ThinQQ > Tutorial abre o `tutorial.html` da versão instalada; se ele não existir, abre
<https://israeldias-hub.github.io/ThinQQ/tutorial/tutorial.html>. Para essa página existir, habilite uma vez
em Settings > Pages > "Deploy from a branch", branch `main`, pasta `/docs`.

Atenção: o Pages por branch só serve a raiz (`/`) ou a pasta `/docs`, então mover o tutorial para outra subpasta de `docs/` (como `docs/tutorial-site`) **não** reduz o que é publicado: com `/docs`, tudo sob `docs/` vai ao ar, inclusive `docs/superpowers` (especificação e plano). Para publicar só o tutorial há duas saídas: (1) publicar por um workflow do GitHub Actions (Settings > Pages > Source: GitHub Actions) que envie apenas a pasta do tutorial, mantendo o caminho `tutorial/tutorial.html` para que `TUTORIAL_URL` em `src/loader/modules/modRibbon.bas` continue valendo; ou (2) tirar a especificação e os planos de `docs/` (por exemplo para `design/`) e manter o Pages em `/docs`. Em qualquer caso, confira antes o item de privacidade da checklist acima (o histórico do Git).

## Arquitetura

- **Loader** (`ThinQQ.xlam`): faixa de opções (`customUI14.xml`), inicialização, atualizador e configurações
  (registro em `HKCU\Software\VB and VBA Program Settings\ThinQQ\Updater`). Muda raramente.
- **Core** (`ThinQQ.core.xlam`): motor de gráficos e API pública `Core_*`. O loader abre o core mais novo de
  `%LOCALAPPDATA%\ThinQQ\versions\<x.y.z>\` e encaminha os comandos por `Application.Run`. Toda função `Core_*`
  trata os próprios erros.
- **Troca a quente:** uma atualização do core baixa o arquivo, confere o SHA-256 do manifesto, instala em uma
  pasta de versão nova e recarrega o core sem reiniciar o Excel. O loader é substituído pelo core (com backup
  `ThinQQ.xlam.bak`) em um segundo momento, para que nenhum código do loader esteja em execução.
- **Dados:** cada gráfico tem um bloco de fórmulas na planilha muito oculta `_ThinQQ`, ligado à tabela de origem
  (maior nos tipos novos: 5 linhas por categoria no Mekko e 40 linhas de escala no Gantt); séries auxiliares `tqq_*`
  (XY no grupo secundário nos gráficos horizontais) levam totais, rótulos e linhas de valor; as setas, o indicador 100% e
  os colchetes são formas dentro do gráfico, reposicionadas nos eventos de cálculo e de redimensionamento.
- **Capacidades:** `modChartKinds` guarda a matriz de recursos por tipo (`Kinds_Supports`); a faixa, os comandos de
  opção e a reconstrução ("Alterar dados") consultam a mesma matriz.
- Fonte em texto em `src/`; os `.xlam` são montados pelo build e não ficam no Git.

## Estrutura

```
build/       build, testes e release (PowerShell 7)
src/loader/  fonte do loader (módulos, formulário, faixa de opções em menus: customUI14.xml)
src/core/    fonte do core: um módulo por família de gráfico (modChartStacked: colunas, barras e linha;
             modChartWaterfall; modChartArea; modChartCombo; modChartMekko; modChartGantt; modChartPie),
             a matriz de capacidades (modChartKinds), decorações (modDecorations), anotações, geometria,
             rótulos, estilo, exemplos (modSamples), classes e o formulário de setas
tests/       testes VBA (TAP), testes PowerShell, servidor Node de fixtures
installer/   instalador e desinstalador por usuário, LEIA-ME.txt
release/     notas e manifesto publicado
docs/        tutorial (tutorial.html), especificações e planos
```

Especificação: `docs/superpowers/specs/2026-10-10-thinqq-design.md` (arquitetura, distribuição e atualização; Fase 1).
Plano da Fase 1: `docs/superpowers/plans/2026-10-10-thinqq-fase1.md`.

Fase 2 (1.1.0): adendo à especificação, com a pesquisa do think-cell, as técnicas nativas do Excel escolhidas para
cada tipo e a matriz de capacidades: `docs/superpowers/specs/2026-10-11-thinqq-fase2-design.md`. Plano de execução:
`docs/superpowers/plans/2026-10-11-thinqq-fase2.md`.

## Verificações manuais no Excel real

Os testes automáticos usam um Excel invisível; estas verificações precisam de uma sessão de Excel real (depois de `Instalar-ThinQQ.cmd` ou de `build.ps1 -Install`), com uma tabela 4x3 com cabeçalho para os itens 1 a 11 e as tabelas de exemplo de cada tipo (planilha "Exemplos 2") para os demais:

1. A aba ThinQQ aparece na faixa de opções, ao lado de Inserir, com os cinco menus de inserção (Colunas, Barras, Cascata, Linha e área, Mais tipos). Esta é a primeira conferência: sem a aba, nenhum outro item vale.
2. Com células selecionadas, só os cinco menus, Mais opções (com Mesma escala e Escala automática), Atualizar e o grupo ThinQQ ficam ativos; o resto do grupo "Gráfico selecionado" fica cinza.
3. ThinQQ > Colunas > Colunas empilhadas insere o gráfico à direita, selecionado; Totais e Rótulos pressionados; Totais e Conectores ligam e desligam.
4. Formato, Paleta e Espaço (que fica cinza em gráfico de linha); Colunas 100% mostra participações nos segmentos e nasce com "100%" à esquerda.
5. Setas > Seta CAGR abre o formulário e desenha o texto; mudar um valor atualiza; Seta diferença; Remover setas; redimensionar reposiciona.
6. Cascata (colunas) com `e`: barras, rótulos e conectores; clicar numa célula deixa cinza o resto do grupo (Mais opções e Atualizar continuam ativos); Sobre, Tutorial, Configurações, Exemplos.
7. Copiar um gráfico e apagar o original: a cópia mantém os dados; inserir linha na tabela de origem mostra o aviso uma vez.
8. Colunas 100% com Totais desligado: copiar e colar o gráfico, apagar o original e clicar em Atualizar; a cópia mantém as participações e continua certa depois de inserir outro gráfico.
9. Verificar atualizações e um ciclo real de atualização (core e, depois, loader); conferir `Workbooks("ThinQQ.core.xlam")`.
10. Instalar-ThinQQ.cmd e uninstall.ps1 no Excel real; conferir que nenhuma chave extra permanece.
11. Os ícones da faixa não aparecem em branco e são coerentes (Mesma escala, Eixo secundário (linha), Indicador 100%, Escala automática, Colchete (Gantt)).
12. Barras: inserir os três tipos a partir de uma tabela 3x3; totais no fim das barras; mudar um valor e ver os totais acompanharem; redimensionar; Seta CAGR e Seta diferença; Formato 1.234,5; Barras 100% com participações e conectores.
13. Cascata com várias séries (tabela Etapa, Loja, Online com `e` nas linhas de total): segmentos coloridos por série, legenda com 2 entradas, rótulos por segmento, conectores no líquido; mudar -200 para -100 e ver tudo acompanhar; `e` na primeira linha com um número na última: a primeira barra é calculada.
14. Cascata (barras) com a tabela Receita a Lucro do exemplo 9: rótulos à direita ou à esquerda, conectores verticais; redimensionar e mudar valores.
15. Área empilhada e Área 100% (rótulos no meio das faixas, totais acima, bordas brancas); Combinação com uma coluna de percentual (linha no eixo secundário, totais só das colunas); Legenda > Nome no fim da linha num gráfico de linha; Legenda > Nenhuma.
16. Mais opções > Linha de valor (média) num empilhado: linha tracejada com "Média N" à direita; mudar um valor move a linha.
17. Dois gráficos com escalas diferentes selecionados com Ctrl+clique > Mais opções > Mesma escala; depois Escala automática; com um só gráfico selecionado ou nenhum, vale para todos os da planilha e a mensagem informa quantos foram afetados (Escala automática com um selecionado age só nele); com dados acima do máximo fixo, o aviso aparece uma vez.
18. Pizza e Rosca de uma tabela de 3 linhas: Mais opções > Rótulos fora da pizza (nomes nas fatias, legenda some; arraste um rótulo para ver as linhas de chamada; na rosca o item fica cinza); Formato 12% mostra participações; Paleta; as setas ficam cinza.
19. Mekko (%) de uma tabela 3x3: larguras proporcionais aos totais, separadores brancos, participações nos segmentos, larguras no topo, nomes embaixo, "100%" à esquerda; mudar um valor redistribui as larguras ao vivo.
20. Mekko (unidades) com a coluna Largura e Linha de valor (média): média ponderada pela largura.
21. Gantt com 3 atividades (uma sem fim, que vira marco): escala semanal no topo com separadores; mudar o fim para o ano seguinte e clicar Atualizar passa a escala para meses; Setas > Colchete (Gantt); Espaço muda o espaçamento das linhas; a 4.ª coluna ("Responsável") aparece à direita das barras.
22. Alterar dados em barras, Mekko e Gantt mantém o tipo e as opções; numa cascata de 1 para 2 colunas de valores, Legenda passa a ficar ativa sem clicar em outro gráfico.
23. Estado dos controles por tipo: pizza (Totais, Conectores, Espaço e as três setas cinza; Rótulos fora da pizza ativo), rosca (Rótulos fora da pizza cinza), Gantt (Colchete (Gantt) ativo; Seta CAGR, Seta diferença, Formato e Legenda cinza), cascata de 1 série (Legenda cinza), Colunas 100% (Seta diferença cinza; Indicador 100% pressionado), linha (Totais, Conectores e Espaço cinza), combinação (Eixo secundário (linha) ativo), Mekko (%) (setas cinza).
24. Trocar a seleção entre gráficos de tipos diferentes da mesma planilha, só clicando: a faixa muda sem outro gesto.
25. ThinQQ > Exemplos abre uma pasta com duas planilhas; percorrer "Exemplos 2" e conferir cada gráfico contra o título.
26. ThinQQ > Tutorial abre o tutorial novo (cinco menus na maquete, uma seção por tipo) e ThinQQ > Sobre mostra Loader 1.1.0 e Core 1.1.0 depois da instalação pelo ZIP.

## Limitações conhecidas

Os detalhes de cada tipo estão na seção "Limitações" do tutorial.

- Adicionar ou remover linhas na tabela de origem exige o botão Alterar dados (o ThinQQ avisa uma vez por gráfico e por sessão quando detecta a mudança de tamanho).
- Linhas filtradas ou ocultas continuam plotadas; a linha de totais de uma Tabela do Excel entra como categoria; formato `.xls` não é suportado (use `.xlsx`, `.xlsm` ou `.xlsb`); colunas empilhadas com sinais misturados têm rótulo de total líquido no topo da parte positiva.
- Cópias de gráficos do ThinQQ e gráficos movidos para planilha de gráfico mantêm os dados, mas deixam de ser reconhecidos como gráficos do ThinQQ (sem botões nem setas automáticas).
- Somente Excel para Windows, Office 16.0; sem Mac nem Excel na web. A compatibilidade com Excel 64 bits é por construção (sem `Declare`), ainda não verificada em teste.
- Cascata com no máximo 6 séries. Pizza e rosca usam só a primeira coluna de valores; valores negativos, zero, texto e erros viram fatias vazias.
- A seta de diferença não existe em gráficos 100%, Mekko, pizza e Gantt; a seta CAGR não existe em Mekko, pizza e Gantt; o colchete só existe no Gantt.
- Mesma escala é uma escala fixa, não dinâmica: depois de mudar muito os dados, use de novo ou Escala automática; um ajuste manual do eixo é desfeito no próximo cálculo ou redimensionamento. Gráficos 100%, Mekko (%), pizza e Gantt não entram. Com um só gráfico selecionado (ou nenhum), a Mesma escala vale para todos os gráficos ThinQQ da planilha.
- A linha de valor só existe na média dos valores. Nome no fim da linha só em linha e combinação.
- Gantt: sem sombreamento de fins de semana, colunas de responsável e observação (só uma 4ª coluna, que vira o rótulo da barra), âncoras, setas de processo nem escala de dias; os separadores da escala são desenhados sobre as barras; o intervalo do eixo só é recalculado pelo ThinQQ (Atualizar ou eventos de cálculo).
- Mekko: sem setas; colunas estreitas ou muitas categorias deixam os rótulos e os nomes sobrepostos.
- O espaço reservado para setas, totais e legenda só é refeito pelos comandos do ThinQQ; trocar a Legenda, desligar o Indicador 100% ou a Linha de valor e usar Escala automática descartam um layout manual da área de plotagem.
- Os eventos que reposicionam as setas só disparam com o gráfico visível; Atualizar refaz o layout.
- Macros precisam estar habilitadas; políticas corporativas podem exigir local confiável ou assinatura digital.

## Licença

Nenhuma licença de código aberto foi escolhida: todos os direitos reservados. O repositório não tem arquivo `LICENSE`; para permitir o uso por terceiros, o autor precisa escolher e adicionar uma licença.
