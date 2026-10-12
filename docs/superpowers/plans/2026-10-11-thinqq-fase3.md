# ThinQQ — Plano de Implementação (Fase 3)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Entregar o ThinQQ 1.2.0 (core e loader): borboleta, dispersão e bolhas, rótulos com soma consistente (arredondamento por maior resto em fórmulas nativas) e "Arredondar intervalo", corte de eixo, cópia para o PowerPoint (imagem e gráfico vinculado), ícones próprios na faixa, scripts de certificado/assinatura e roteiro corporativo, pasta de demonstração gerada por script, planilha "Exemplos 3", documentação de instalação completa, tutorial/README/CHANGELOG e release 1.2.0.

**Architecture:** Mantém a das Fases 1 e 2: `ThinQQ.xlam` (loader: faixa, boot, atualizador) chama `Core_*` em `ThinQQ.core.xlam` via `Application.Run`; cada gráfico é um gráfico nativo do Excel ligado por fórmulas a um bloco da planilha muito oculta `_ThinQQ`, com séries auxiliares `tqq_*` e formas `tqq_ann_*` só para o que não é dado. Novidades estruturais: famílias `butterfly` e `scatter` (módulos próprios), colunas acrescentadas **à direita do bloco sob demanda** (`Block_Extend`) para o arredondamento e o corte de eixo (por isso as colunas Pos/VLine passam a ser localizadas pelo cabeçalho), módulo de automação tardia do PowerPoint, ícones PNG gerados no build e empacotados no `.xlam` do loader, scripts de certificado/assinatura em `tools/`, gerador da pasta de demonstração em `tools/`.

**Tech Stack:** VBA 7 (Excel 2016+ Windows, 32/64 bits, sem `Declare`, sem referências), customUI14, PowerShell 7 (build/testes/release/ferramentas) e 5.1 (instalador), Node.js (fixtures HTTP), Git.

**Spec:** `.superpowers/sdd/fase3-draft/design.md` (decisões, pesquisa e sondas medidas desta fase) sobre `docs/superpowers/specs/2026-10-11-thinqq-fase2-design.md` (matriz, blocos por família, faixa) e `docs/superpowers/specs/2026-10-10-thinqq-design.md` (arquitetura, distribuição, atualização). Ledger e relatórios da Fase 2: `.superpowers/sdd/2026-10-11-thinqq-fase2/` (progress.md, task-N-report.md: armadilhas medidas no Excel real).

**Estado de execução:** nenhuma tarefa concluída. Branch: a indicada pelo controlador no prompt de cada implementador (prevista `feat/fase3`, criada pelo controlador a partir de `feat/fase2`); o implementador **não** cria nem troca de branch. As regras de ambiente ficam em `.superpowers/sdd/2026-10-11-thinqq-fase2/implementer-rules.md` (valem para a Fase 3).

## Global Constraints

- **Ambiente:** Excel 2019 32-bit pt-BR no Windows 11; PowerPoint 16.0 instalado; testes em Excel invisível próprio. O VBA compila sob demanda: o build roda o lint (`build/lib/Vba-Lint.ps1`) e uma compilação completa; erro de compilação => saída 3; travamento morto pelo watchdog => saída 2. Fontes UTF-8 com **somente caracteres cp1252** nos arquivos VBA (nada de `→`, `≥`, `…`, aspas tipográficas fora do cp1252, emojis), sem ` _` perdido, `Option Explicit`, declarações antes dos procedimentos, `Attribute VB_Name` na primeira linha de cada `.bas`.
- **Testes:** `Public Sub Test_x()` sem parâmetros em `tests/vba/{core,loader}/Test*.bas`, executados por um despachante gerado (nunca `Application.Run`), via `pwsh build/run-tests.ps1` (suíte completa ~9–10 min ao fim desta fase, estimativa do pré-flight: o `-TimeoutSec 900` dos comandos basta; parâmetros `-Target core|loader|all`, `-NoBuild`, `-TimeoutSec`, `-Filter "Modulo.Teste|outro"` (pula as suítes PowerShell), `-LoaderSwap`). Asserts: `AssertTrue/False/Equal/Almost`, `NewTempWorkbook`, `CloseTemp`, `TempDir`, `ForceLayout`, `RepoPath`. Cada módulo de teste novo define os próprios `Fixture`, `Teardown` (`Watchers_Stop` + `CloseTemp`) e `TqqNameCount` (são `Private` em `TestCoreApi`).
- **Instâncias do usuário:** os processos Excel já abertos quando a tarefa começa (anote `Get-Process excel | Select-Object Id, StartTime` no início) nunca são tocados nem encerrados; um deles pode manter `dist\ThinQQ.core.test.xlam` travado no repositório real. Por isso **builds e testes rodam numa CÓPIA**: `robocopy "<repo>" C:\temp\tqq-f3 /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS` (repetir antes de cada rodada; edite sempre o repositório real; commite a partir do repositório real; `cd` só com caminho absoluto). Nunca `build.ps1 -Install`; nunca tocar em `%LOCALAPPDATA%\ThinQQ`, no registro real do usuário (além do que os testes existentes do atualizador salvam/restauram) nem nas lojas de certificados além do que `Test-Signing.ps1` cria e remove com assunto único (em `My`, `CA` e `TrustedPublisher`). **PowerPoint:** nenhum teste ou script cria uma instância do PowerPoint se já existe um processo `POWERPNT` (é aplicativo de instância única: `CreateObject` anexaria ao do usuário); o que a tarefa criar, ela encerra pelo PID que anotou. **Disco C: com pouco espaço livre:** apague as suas pastas temporárias ao fim de cada rodada (`C:\temp\tqq-f3`, `C:\temp\tqq-f3-*`); nunca apague nada fora delas.
- **Código:** procedimentos `Core_*` nunca propagam erro (`Errors_Handle`/`Errors_HandleErr`); `MsgBox`/formulários/`InputBox` só fora do build `Test` (`Core_Notify`, `RxMsg`, `Loader_Msg`; os comandos que abrem `InputBox` saem antes dele quando `BUILD_CONFIG = "Test"`, como `Core_AskCategories`); fórmulas via `.Formula`/`.FormulaR1C1` em inglês; formatos de **rótulo** via `Style_AssignNumberFormat`/`NumberFormatLocal` (nunca `DataLabel.NumberFormat`), formatos de **célula** do bloco via `NumberFormatLocal` com `Style_LocalizeCode` para números e `Range.NumberFormat` em inglês só para datas e `General`; testes de texto renderizado via `DataLabel.Text` com expectativas derivadas de `Application.International`/`Format$`; textos de interface em pt-BR; identificadores em inglês com os prefixos existentes (`Core_`, `Stacked_`, `Area_`, `Geom_`, `Ann_`, `Labels_`, `Style_`, `Chart_`, `Util_`, `Registry_`, `CalcSheet_`, `Block_`, `Deco_`, `Samples_`, `Rx_`, `Loader_`) e os novos (`Bfly_`, `Scatter_`, `Round_`, `Break_`, `Ppt_`).
- **Todo tipo novo de gráfico entrega, obrigatoriamente:** `Insert` com rollback (`Util_TestHookAfterBlock` após alocar o bloco, `Util_TestHookAfterChart` após o gráfico existir), opções persistidas por `Registry_SetOption` (incl. `n`), rótulos ligados a células, recusas explícitas pela matriz (7001 anotações, 7002 opções), testes de geometria independentes das fórmulas de produção (regra de três sobre `PlotArea.Inside*` e `Axis.MinimumScale/MaximumScale` reais), exemplo em `Samples_Create`, seção no tutorial, entrada na faixa validada por `Test-Ribbon`/`TestRibbon` e cobertura em `Test-Docs`.
- **Armadilhas medidas (ledger da Fase 2 e sondas desta fase; não redescubra):** rótulos e `Format$` seguem a localidade; um ponto `NA()` com `DisplayBlanksAs = xlZero` mostra o rótulo no zero (use `xlNotPlotted` quando o rótulo deve sumir); a geometria de um `DataLabel` só atualiza após `Chart.Export` (`ForceLayout`); `Chart.Export` grava PNG de 0 byte para gráfico fora da janela do Excel invisível (ative a planilha e role até o gráfico antes de exportar, ou exporte só gráficos em `Left < 700`); ids do customUI não começam por dígito e filho de menu não leva `size` (a aba inteira some em silêncio); a rosca não aceita `DataLabels.Position`; `Date1904`; o `Format.Fill` de marcador XY não pega a cor (use `MarkerBackgroundColor/MarkerForegroundColor`); pontos XY do grupo secundário entram na escala automática do eixo primário; `Axes(xlCategory).CrossesAt = x0` desenha a linha vertical em x0 e `Axes(xlValue).CrossesAt = y0` a horizontal em y0; `MaximumScale` fixo corta colunas na borda do plot e empurra os rótulos nativos cortados para `Top = -4`; `PowerPoint.Application.Visible = msoFalse` é recusado (use `Presentations.Add(msoFalse)` nos testes); PowerShell: nomes de variável **não** diferenciam maiúsculas (`$D` e `$d` são a mesma) e literais hexadecimais que cabem em Int32 viram negativos (use `[uint64]`).
- **Armadilhas achadas no pré-flight desta fase (valem para todas as tarefas; não redescubra):** VBA: nada de `ElseIf` num `If` de uma linha (`If a Then x ElseIf b Then y` não compila: use o bloco `If ... ElseIf ... End If`); toda célula que um bloco usa, inclusive a linha de meta de uma coluna acrescentada por `Block_Extend`, fica **dentro** do nome `tqq_<id>` (uma célula fora do bloco não é limpa por `CalcSheet_DeleteBlock`, passa para o próximo bloco alocado nas mesmas linhas e faz `Block_Extend` recusar com 6014); `Series.Name` devolve o texto do nome, nunca a referência (procure séries de dados pelo índice, não por `Chart_FindSeries(cht, "=Dados!$B$1")`). PowerShell: nunca nomeie uma função com o nome de um alias embutido (`r`, `h`, `ls`, `cd`, `gc`, `gi`, `ni`...): o alias vence a função e a chamada falha; `2>&1` não captura `Write-Host` (use `*>&1`); uma função que devolve uma coleção COM (um `Range` de várias células) a desenrola no pipeline (use `return ,$rng`); num literal de várias linhas, `@(,@(a), @(b))` embrulha só a primeira linha num array a mais (escreva `@(@(a), @(b))`; `,@(...)` só dentro de `ForEach-Object` ou em `+= ,@(...)`); `Range.Value2` de várias linhas exige matriz 2-D (`object[,]`), não um array de arrays (grave linha a linha); variáveis de nível de script de um arquivo que outros fazem dot-source vazam para eles (monte tabelas dentro de `& { ... }` ou use nomes `$script:Tqq*`); não use `$args` como variável (é automática); um `.ps1` com acentos que precisa rodar no Windows PowerShell 5.1 vai em UTF-8 **com BOM** (como `installer/install.ps1`). Scripts PowerShell chamados pelo `run-tests`: corpo em `try/catch/finally`, `exit 0`/`exit 1` explícito no fim e, se podem lançar ou chamam outro script que deixa `$LASTEXITCODE` (um `exit 2` interno continua valendo no chamador), chamada num processo próprio (`& (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File ...`, como os testes do instalador). Lojas de certificados: um teste que instala algo nas lojas do usuário remove **exatamente** o que criou, pelo assunto único, em todas as lojas em que o item pode estar, inclusive a cópia que `New-SelfSignedCertificate` grava em `Cert:\CurrentUser\CA` (Autoridades de Certificação Intermediárias), e confere que nada sobrou. PowerPoint: em build `Test`, `Ppt_GetApp` só usa a instância injetada por `Ppt_SetAppForTest` (nunca anexa ao PowerPoint do usuário). Arquivos binários gerados (`.xlsx`) não são lidos pelo `Test-Privacy`: o teste do próprio arquivo confere que `docProps` não leva autor.
- **Commits:** pequenos, em português, prefixos `feat:`/`fix:`/`test:`/`docs:`/`build:`; trailer `Co-Authored-By:` **indicado no prompt de cada implementador**; "(+ trailer)" nos comandos de commit deste plano significa essa linha. **Não** fazer push, **não** criar nem trocar de branch.
- **Versões:** `version.json` passa a `{"loader":"1.2.0","core":"1.2.0"}` só na Tarefa 11. O build importa automaticamente todo `.bas` de `src/<alvo>/modules`, `.cls` de `classes`, `.json`+`.vba` de `forms` e todo `Test*.bas` de `tests/vba/<alvo>`.
- **Processo:** cada tarefa é executada por um implementador novo que só vê o texto da tarefa; tudo o que ele precisa está na própria tarefa. Cada tarefa termina com a suíte completa verde na cópia e um commit por tarefa (ou mais, se útil). Modelo sugerido por tarefa: **cheap** (edições mecânicas), **standard** (código com testes), **most-capable** (famílias novas, geometria, geração da demonstração).

## Review Focus

1. **Borboleta com um nome de categoria longo e um lado muito maior que o outro** (ex.: "Trinta a cinquenta e nove", 600 à esquerda e 60 à direita): o usuário espera o nome inteiro no vão central sem invadir as barras, a mesma escala dos dois lados e os totais no fim da barra direita. Teste em Tarefa 1 (`Test_Bfly_LongName_GapFitsAndAxisSymmetric`).
2. **Dispersão em que um ponto troca de grupo depois de inserido** e um ponto com texto em X: o usuário espera o ponto migrar de série (cor e legenda) sem rótulo fantasma, e o ponto inválido fora do gráfico. Teste em Tarefa 2 (`Test_Scatter_GroupChange_MovesPointAndLabel`).
3. **Colunas 100% com três participações de 33,3%** (e um formato `12%`): o usuário espera 34% + 33% + 33% = 100% (empate de restos: a primeira coluna recebe a unidade), não 99%; e com `12,3%`, 33,4/33,3/33,3; uma categoria vazia não mostra "1%". Teste em Tarefa 3 (`Test_Round_Stacked100_SharesSumTo100`).
4. **Corte de eixo numa coluna que cruza o limite** (segmento parte dentro, parte fora): o usuário espera a coluna cortada na borda, o total real acima, o rótulo do segmento atravessado na base, nenhum rótulo "órfão" no alto do gráfico e a marca de corte sobre a coluna; mudar os dados abaixo do limite remove a marca. Teste em Tarefa 4 (`Test_Break_SegmentCrossingLimit_LabelsAndMark`).
5. **Copiar para o PowerPoint com a pasta não salva e sem PowerPoint aberto**: o usuário espera a imagem num slide novo (abrindo o PowerPoint) e, no modo vinculado, uma mensagem clara pedindo para salvar. Testes em Tarefa 5 (`Test_Ppt_LinkedPrecheck_UnsavedWorkbook`, `Test_Ppt_RealPaste_WhenAvailable`).

---

## Estrutura de arquivos (novos e alterados)

```
src/core/modules/
├── modChartKinds.bas        ALT   20 tipos, 20 recursos, rótulos novos, Kinds_Insert com butterfly/scatter
├── modCalcSheet.bas         ALT   Block_FindCol, Block_Extend, Block_SourceRange
├── modChartStacked.bas      ALT   Stacked_PosCol/VLineCol por cabeçalho; Stacked_SetLabels com roundsum (T3); Stacked_SetTotals com LblPos (T4)
├── modChartArea.bas         ALT   Area_PosCol/VLineCol/SeriesCount por cabeçalho; Area_SetLabels com roundsum (T3)
├── modLabels.bas            ALT   famílias scatter/butterfly; RndD no formato (T3)
├── modChartStyle.bas        ALT   Style_PaletteRGB; Style_ApplyNumberFormat por família (scatter sai; butterfly -> Bfly_ApplyNumberFormat)
├── modChartButterfly.bas    NOVO  Bfly_* (T1)
├── modChartScatter.bas      NOVO  Scatter_* (T2)
├── modRounding.bas          NOVO  Round_* (T3)
├── modAxisBreak.bas         NOVO  Break_* (T4)
├── modPowerPoint.bas        NOVO  Ppt_* (T5)
├── modAnnotations.bas       ALT   Ann_GetCategoryData butterfly (T1); kind=break em Ann_Refresh, Ann_RemoveAll poupa as marcas, topo limitado ao corte (T4)
├── modDecorations.bas       ALT   Deco_SameScale ignora gráfico com vbreak (T4)
├── modCoreApi.bas           ALT   despachos novos, Core_RoundRange, Core_SetAxisBreak/At/RemoveAxisBreak, Core_CopyToPowerPoint, Core_SafeRefresh
├── modChartUtil.bas         ALT   Chart_EnsureRoomAfterChange chama Break_EnsureRoom (T4)
└── modSamples.bas           ALT   planilha "Exemplos 3" (T9)
src/loader/ribbon/customUI14.xml  ALT  (T8); src/loader/modules/modRibbon.bas ALT (T8)
build/lib/Ribbon-Icons.ps1   NOVO  gravador PNG + tabela de ícones (T6); build/lib/Build-Common.ps1 ALT (Add-RibbonToPackage -Icons); build/build.ps1 ALT (-RibbonIcons)
build/release.ps1            ALT   -SignPfx / -SignPassword (T7); build/run-tests.ps1 ALT (Test-Signing, Test-DemoWorkbook)
tools/New-ThinQQCert.ps1, tools/Trust-ThinQQPublisher.ps1, tools/Sign-ThinQQ.ps1   NOVOS (T7)
tools/DemoCoverage.ps1, tools/New-DemoWorkbook.ps1   NOVOS (T10); exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx NOVO (T10)
tests/vba/core/  TestButterfly.bas, TestScatter.bas, TestRounding.bas, TestAxisBreak.bas, TestPowerPoint.bas (NOVOS);
                 TestKinds, TestCalcSheet, TestCoreApi, TestSamples (ALT)
tests/vba/loader/TestRibbon.bas  ALT;  tests/ps/Test-Ribbon.ps1 ALT; tests/ps/Test-Docs.ps1 ALT
tests/ps/Test-Signing.ps1, tests/ps/Test-DemoWorkbook.ps1   NOVOS
docs/tutorial/tutorial.html, README.md, CHANGELOG.md, version.json, release/notes.txt, installer/LEIA-ME.txt  ALT (T0 matriz; T8 rótulos; T11 tudo)
```

## Como trabalhar (todas as tarefas)

```powershell
# 1) editar no repositório real; 2) espelhar para a cópia; 3) rodar lá
robocopy "<repo>" C:\temp\tqq-f3 /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS
pwsh C:\temp\tqq-f3\build\run-tests.ps1 -Target core -Filter "TestButterfly" -TimeoutSec 300     # iteração (RED/GREEN)
pwsh C:\temp\tqq-f3\build\run-tests.ps1 -TimeoutSec 900                                          # suíte completa antes de commitar
# saída 3 = erro de compilação (módulo:linha na tela); 2 = travou (watchdog); 1 = asserts; 0 = ok
Remove-Item C:\temp\tqq-f3 -Recurse -Force                                                      # ao terminar a tarefa (disco quase cheio)
```
Fixtures: cada módulo de teste tem a sua `Private Function Fixture(...)`; use `Teardown`/`Watchers_Stop` quando chamar `Core_*` (estado global dos observadores). Nunca deixe `Util_*ForTest` ligado em caminho de erro (padrão `On Error GoTo Cleanup` que reseta e repropaga).

---

### Tarefa 0: Matriz 1.2.0, utilitários de bloco e famílias novas nos pontos comuns (infraestrutura) — modelo: standard

**Files:**
- Modify: `src/core/modules/modChartKinds.bas`, `src/core/modules/modCalcSheet.bas`, `src/core/modules/modChartStacked.bas`, `src/core/modules/modChartArea.bas`, `src/core/modules/modLabels.bas`, `src/core/modules/modChartStyle.bas`, `docs/tutorial/tutorial.html` (só a matriz `#matriz`), `tests/ps/Test-Docs.ps1` (contagens 20/20 e limite do tutorial em 160 KB)
- Test: `tests/vba/core/TestKinds.bas`, `tests/vba/core/TestCalcSheet.bas`, `tests/vba/core/TestCoreApi.bas` (teste temporário de tipo não pronto)

**Interfaces:**
- Produces (`modChartKinds`): `KINDS` passa a ter 20 tipos (`...,doughnut,butterfly,scatter,bubble,`); `Kinds_IsHorizontal` inclui `butterfly`; `Kinds_Family`: `butterfly` → `"butterfly"`, `scatter`/`bubble` → `"scatter"`; `Kinds_Supports` com a matriz do design §6 (listas literais; `numfmt` vira lista explícita sem `gantt`, `scatter`, `bubble`; recursos novos `roundsum`, `break`, `trend`, `quad`); `Kinds_FeatureLabel`: `roundsum` → "Rótulos com soma consistente", `break` → "Corte de eixo", `trend` → "Linha de tendência", `quad` → "Quadrantes (médias)"; `Kinds_Insert`: `Case "butterfly", "scatter"` caem no `Case Else` (7200 "ainda não está disponível") até as Tarefas 1 e 2.
- Produces (`modCalcSheet`): `Block_FindCol(blk As Range, header As String) As Long` (coluna do corpo cujo cabeçalho — linha 1 do corpo — é `header`; 0 se não existe); `Block_Extend(wb As Workbook, id As String, extraCols As Long) As Range` (redimensiona o bloco `tqq_<id>` para `Columns.Count + extraCols`, limpa as células novas, reescreve o `RefersTo` do nome e devolve o bloco novo; erro 6014 se as células à direita não estiverem vazias); `Block_SourceRange(cht As Chart) As Range` (a tabela de origem gravada na meta, coluna C, resolvida na pasta do gráfico; `Nothing` se não resolver).
- Produces (`modChartStacked`): `Stacked_PosCol(blk)` = `Block_FindCol(blk, "Pos")` ou, se 0, `blk.Columns.Count - 1`; `Stacked_VLineCol(blk)` = `Block_FindCol(blk, "VLine")` ou `blk.Columns.Count`.
- Produces (`modChartArea`): `Area_PosCol`/`Area_VLineCol` idem; `Area_SeriesCount(blk)` = número de cabeçalhos `Val<k>` contíguos a partir da coluna 4 (reserva: `(Columns.Count - 5) \ 2` quando não há cabeçalho `Val1`); `Area_BuildBlock` grava todos os cabeçalhos `Val<k>` antes dos `Mid<k>` (`Area_MidCol` depende da contagem por cabeçalho).
- Produces (`modLabels`): `Labels_RefreshPositions` e `Labels_ApplyCellFormat` saem cedo na família `scatter` (a `butterfly` segue o caminho padrão: coluna 2 e `tqq_total`).
- Produces (`modChartStyle`): `Style_PaletteRGB(cht As Chart, v As Variant) As Long` (RGB de uma cor da paleta; na "tema", a cor do tema da pasta; igual ao `PaletteRGB` privado do Gantt, que fica como está); `Style_ApplyNumberFormat` sai cedo na família `scatter` e, na família `butterfly`, chama `Bfly_ApplyNumberFormat cht` (**só na Tarefa 1**, quando a função existir; nesta tarefa a família `butterfly` apenas sai cedo).
- Produces (tutorial, `#matriz`): 3 linhas novas (`data-kind="butterfly"` "Borboleta", `"scatter"` "Dispersão", `"bubble"` "Bolhas") nas duas tabelas existentes e uma terceira tabela `class="mx"` ("Arredondamento, corte e análise") com `data-f="roundsum"`, `"break"`, `"trend"`, `"quad"` e as 20 linhas. `Test-Docs`: `Check ($allKinds.Count -eq 20)`, `Check ($matrix.Count -eq 20)` e o limite do tutorial em 160 KB (`163840` bytes).

- [ ] **Passo 1: Testes (RED)** — em `tests/vba/core/TestKinds.bas` acrescente:

```vb
Public Sub Test_Kinds_Phase3_TypesAndFamilies()
    AssertTrue Kinds_IsValid("butterfly"): AssertTrue Kinds_IsValid("scatter"): AssertTrue Kinds_IsValid("bubble")
    AssertEqual "butterfly", Kinds_Family("butterfly")
    AssertEqual "scatter", Kinds_Family("scatter"): AssertEqual "scatter", Kinds_Family("bubble")
    AssertTrue Kinds_IsHorizontal("butterfly"): AssertFalse Kinds_IsHorizontal("scatter")
    AssertFalse Kinds_Is100("butterfly")
    AssertEqual 3, Kinds_ExpectedCategoryRefs("scatter", 3)
End Sub

Public Sub Test_Kinds_Phase3_Matrix()
    AssertTrue Kinds_Supports("butterfly", "totals"): AssertTrue Kinds_Supports("butterfly", "gap"): AssertTrue Kinds_Supports("butterfly", "samescale")
    AssertFalse Kinds_Supports("butterfly", "cagr"): AssertFalse Kinds_Supports("butterfly", "connectors"): AssertFalse Kinds_Supports("butterfly", "vline")
    AssertTrue Kinds_Supports("scatter", "trend"): AssertTrue Kinds_Supports("bubble", "quad"): AssertTrue Kinds_Supports("scatter", "legend")
    AssertFalse Kinds_Supports("scatter", "numfmt"): AssertFalse Kinds_Supports("bubble", "gap"): AssertFalse Kinds_Supports("scatter", "samescale")
    AssertFalse Kinds_Supports("stacked", "trend"): AssertFalse Kinds_Supports("stacked", "quad")
    AssertTrue Kinds_Supports("stacked", "roundsum"): AssertTrue Kinds_Supports("area100", "roundsum"): AssertFalse Kinds_Supports("clustered", "roundsum")
    AssertFalse Kinds_Supports("mekko", "roundsum"): AssertFalse Kinds_Supports("pie", "roundsum")
    AssertTrue Kinds_Supports("stacked", "break"): AssertTrue Kinds_Supports("barclustered", "break"): AssertFalse Kinds_Supports("stacked100", "break")
    AssertFalse Kinds_Supports("area", "break"): AssertFalse Kinds_Supports("waterfall", "break")
    AssertTrue Kinds_Supports("pie", "numfmt"): AssertFalse Kinds_Supports("gantt", "numfmt")
    AssertEqual "Rótulos com soma consistente", Kinds_FeatureLabel("roundsum")
    AssertEqual "Corte de eixo", Kinds_FeatureLabel("break")
    AssertEqual "Linha de tendência", Kinds_FeatureLabel("trend")
    AssertEqual "Quadrantes (médias)", Kinds_FeatureLabel("quad")
End Sub

Public Sub Test_Style_PaletteRGB()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set co = ws.ChartObjects.Add(10, 10, 100, 100)
    AssertEqual RGB(31, 56, 100), Style_PaletteRGB(co.Chart, Style_Palette("padrao")(0))
    AssertEqual RGB(64, 64, 64), Style_PaletteRGB(co.Chart, Style_Palette("cinza")(0))
    AssertEqual wb.Theme.ThemeColorScheme.Colors(msoThemeAccent1).RGB, Style_PaletteRGB(co.Chart, Style_Palette("tema")(0)), "tema: cor de destaque 1 da pasta"
    CloseTemp wb
End Sub
```

Em `tests/vba/core/TestCalcSheet.bas`:

```vb
Public Sub Test_Block_FindCol_Extend_SourceRange()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, blk As Range, blk2 As Range, cols As Long, src As Range, calc As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 10, 5): ws.Range("A3:C3").Value = Array(2023, 20, 6)
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Set cht = Stacked_Insert(ws, d, "stacked").Chart
    Set blk = Registry_Block(cht)
    cols = blk.Columns.Count
    AssertEqual cols - 1, Block_FindCol(blk, "Pos"): AssertEqual cols, Block_FindCol(blk, "VLine")
    AssertEqual 0, Block_FindCol(blk, "NaoExiste")
    AssertEqual cols - 1, Stacked_PosCol(blk): AssertEqual cols, Stacked_VLineCol(blk)
    Set blk2 = Block_Extend(wb, Registry_IdFromChart(cht), 3)
    AssertEqual cols + 3, blk2.Columns.Count
    AssertEqual cols + 3, Registry_Block(cht).Columns.Count, "o nome tqq_<id> cobre as colunas novas"
    AssertEqual blk.Row, blk2.Row: AssertEqual blk.Rows.Count, blk2.Rows.Count
    AssertEqual cols - 1, Stacked_PosCol(blk2), "Pos continua sendo achada pelo cabecalho depois da extensao"
    AssertEqual cols, Stacked_VLineCol(blk2)
    AssertEqual "", CStr(Block_Body(blk2).Cells(1, cols + 1).Value), "colunas novas vazias"
    Set src = Block_SourceRange(cht)
    AssertEqual "$A$1:$C$3", src.Address(True, True): AssertEqual "Dados", src.Worksheet.Name
    ' extensao recusada quando ha dado a direita (nunca sobrescreve)
    Set calc = wb.Worksheets(CALC_SHEET)
    calc.Cells(blk2.Row + 1, blk2.Column + blk2.Columns.Count).Value = "x"
    On Error Resume Next
    Block_Extend wb, Registry_IdFromChart(cht), 1
    AssertEqual vbObjectError + 6014, Err.Number, "Block_Extend com celula ocupada a direita"
    On Error GoTo 0
    CloseTemp wb
End Sub

Public Sub Test_Area_Cols_ByHeader_AfterExtend()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, blk As Range, p As Long, v As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 10, 5): ws.Range("A3:C3").Value = Array(2023, 20, 6)
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Set cht = Area_Insert(ws, d, "area").Chart
    Set blk = Registry_Block(cht)
    AssertEqual 2, Area_SeriesCount(blk): p = Area_PosCol(blk): v = Area_VLineCol(blk)
    AssertEqual "Mid1", Block_Body(blk).Cells(1, Area_MidCol(blk, 1)).Value: AssertEqual "Mid2", Block_Body(blk).Cells(1, Area_MidCol(blk, 2)).Value
    AssertEqual "Pos", Block_Body(blk).Cells(1, p).Value: AssertEqual "VLine", Block_Body(blk).Cells(1, v).Value
    Set blk = Block_Extend(wb, Registry_IdFromChart(cht), 5)
    AssertEqual 2, Area_SeriesCount(blk), "numero de series pelo cabecalho Val<k>, nao pela largura"
    AssertEqual p, Area_PosCol(blk): AssertEqual v, Area_VLineCol(blk)
    CloseTemp wb
End Sub
```

Em `tests/vba/core/TestCoreApi.bas` (temporário; a Tarefa 2 o remove quando os três tipos existirem):

```vb
' Temporario: a Tarefa 2 o remove ao implementar dispersao/bolhas (os ultimos tipos novos). Tipo valido ainda sem
' implementacao avisa e nao deixa bloco.
Public Sub Test_Core_Insert_NewKindsNotReady_NotifiesWithoutChart()
    Dim wb As Workbook, ws As Worksheet, k As Variant
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    For Each k In Array("butterfly", "scatter", "bubble")
        ws.Range("A1:C5").Select
        Core_ClearLastNotice
        Core_Insert CStr(k)
        AssertEqual 0, ws.ChartObjects.Count, CStr(k)
        AssertEqual 0, TqqNameCount(wb), CStr(k)
        AssertTrue InStr(Core_LastNotice(), "ainda não está disponível") > 0, CStr(k) & ": " & Core_LastNotice()
    Next
    Teardown wb
End Sub
```

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-f3\build\run-tests.ps1 -Target core -Filter "Phase3|PaletteRGB|Block_FindCol|Area_Cols|NewKindsNotReady" -TimeoutSec 300` → saída 3 (símbolos indefinidos).

- [ ] **Passo 3: Implementar `modChartKinds`** — substitua a constante e as funções indicadas:

```vb
Private Const KINDS As String = ",stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,gantt,pie,doughnut,butterfly,scatter,bubble,"

Public Function Kinds_IsHorizontal(kind As String) As Boolean
    Kinds_IsHorizontal = InList("barstacked,barclustered,barstacked100,barwaterfall,gantt,butterfly", kind)
End Function

Public Function Kinds_Family(kind As String) As String
    If InList("stacked,clustered,stacked100,line,barstacked,barclustered,barstacked100", kind) Then
        Kinds_Family = "stacked"
    ElseIf InList("waterfall,barwaterfall", kind) Then
        Kinds_Family = "waterfall"
    ElseIf InList("area,area100", kind) Then
        Kinds_Family = "area"
    ElseIf kind = "combo" Then
        Kinds_Family = "combo"
    ElseIf InList("mekko,mekkounits", kind) Then
        Kinds_Family = "mekko"
    ElseIf kind = "gantt" Then
        Kinds_Family = "gantt"
    ElseIf InList("pie,doughnut", kind) Then
        Kinds_Family = "pie"
    ElseIf kind = "butterfly" Then
        Kinds_Family = "butterfly"
    ElseIf InList("scatter,bubble", kind) Then
        Kinds_Family = "scatter"
    End If
End Function

' Matriz de capacidades (design da Fase 3, secao 6). False para tipo ou recurso desconhecido. Listas literais: tests/ps/Test-Docs.ps1 le
' estas linhas por expressao regular (Case "x": list = "..." | allKinds | Replace(allKinds, ",y", "")) para conferir a tabela do tutorial.
Public Function Kinds_Supports(kind As String, feature As String, Optional seriesCount As Long = 0) As Boolean
    Dim list As String, allKinds As String
    If Not Kinds_IsValid(kind) Then Exit Function
    allKinds = Mid$(KINDS, 2, Len(KINDS) - 2)
    Select Case feature
        Case "totals":     list = "stacked,stacked100,barstacked,barstacked100,area,area100,combo,mekko,mekkounits,butterfly"
        Case "labels":     list = allKinds
        Case "connectors": list = "stacked,stacked100,barstacked,barstacked100,waterfall,barwaterfall,combo"
        Case "numfmt":     list = "stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,pie,doughnut,butterfly"
        Case "palette":    list = allKinds
        Case "gap":        list = "stacked,clustered,stacked100,barstacked,barclustered,barstacked100,waterfall,barwaterfall,combo,gantt,butterfly"
        Case "legend":     list = Replace(allKinds, ",gantt", "")      ' waterfall/barwaterfall: so com varias series (abaixo)
        Case "lineend":    list = "line,combo"
        Case "ind100":     list = "stacked100,barstacked100,area100,mekko"
        Case "vline":      list = "stacked,clustered,line,barstacked,barclustered,area,combo,mekkounits"
        Case "secondary":  list = "combo"
        Case "pieout":     list = "pie"
        Case "cagr":       list = "stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo"
        Case "diff":       list = "stacked,clustered,line,waterfall,barstacked,barclustered,barwaterfall,area,combo"
        Case "bracket":    list = "gantt"
        Case "samescale":  list = "stacked,clustered,line,waterfall,barstacked,barclustered,barwaterfall,area,combo,mekkounits,butterfly"
        Case "roundsum":   list = "stacked,stacked100,barstacked,barstacked100,area,area100"
        Case "break":      list = "stacked,clustered,barstacked,barclustered"
        Case "trend":      list = "scatter,bubble"
        Case "quad":       list = "scatter,bubble"
        Case Else: Exit Function
    End Select
    Kinds_Supports = InList(list, kind)
    If Kinds_Supports And feature = "legend" And Kinds_Family(kind) = "waterfall" Then Kinds_Supports = (seriesCount > 1)
End Function
```

Em `Kinds_FeatureLabel` acrescente os quatro `Case`, cada um numa linha só e no formato das linhas existentes (o `Test-Docs` lê os rótulos com a expressão regular `Case "x":\s*Kinds_FeatureLabel = "..."`, de uma linha):

```vb
        Case "roundsum":   Kinds_FeatureLabel = "Rótulos com soma consistente"
        Case "break":      Kinds_FeatureLabel = "Corte de eixo"
        Case "trend":      Kinds_FeatureLabel = "Linha de tendência"
        Case "quad":       Kinds_FeatureLabel = "Quadrantes (médias)"
```
`Kinds_Insert` fica como está (as famílias novas caem no 7200).

- [ ] **Passo 4: Implementar `modCalcSheet`** (acrescentar ao fim):

```vb
' Coluna do corpo (1..Columns.Count) cujo cabecalho (linha 1 do corpo) e exatamente header; 0 se nao existe. As familias usam esta
' busca em vez da posicao: a Fase 3 acrescenta colunas A DIREITA do bloco sob demanda (Block_Extend), e "as duas ultimas colunas" deixam
' de ser Pos/VLine.
Public Function Block_FindCol(blk As Range, header As String) As Long
    Dim j As Long, hdr As Range
    If blk Is Nothing Then Exit Function
    If blk.Rows.Count < 2 Then Exit Function
    Set hdr = blk.Rows(2)
    For j = 1 To hdr.Cells.Count
        If CStr(hdr.Cells(1, j).Value) = header Then Block_FindCol = j: Exit Function
    Next
End Function

' Acrescenta extraCols colunas vazias a direita do bloco tqq_<id> (os blocos se empilham a partir da coluna A: a area a direita e livre) e
' reescreve o nome para cobri-las. Devolve o bloco novo. Recusa (6014) se alguma celula a direita estiver ocupada: nunca sobrescreve.
Public Function Block_Extend(wb As Workbook, id As String, extraCols As Long) As Range
    Dim blk As Range, ext As Range, newBlk As Range
    Set blk = CalcSheet_FindBlock(wb, id)
    If blk Is Nothing Then Err.Raise vbObjectError + 6015, "Block_Extend", "Bloco inexistente: " & id
    If extraCols < 1 Then Set Block_Extend = blk: Exit Function
    Set ext = blk.Offset(0, blk.Columns.Count).Resize(blk.Rows.Count, extraCols)
    If Application.WorksheetFunction.CountA(ext) > 0 Then Err.Raise vbObjectError + 6014, "Block_Extend", "Há dados à direita do bloco " & id & "; o bloco não pode ser estendido."
    ext.Clear
    Set newBlk = blk.Resize(blk.Rows.Count, blk.Columns.Count + extraCols)
    wb.Names(ID_PREFIX & id).RefersTo = "='" & CALC_SHEET & "'!" & newBlk.Address(True, True)
    Set Block_Extend = newBlk
End Function

' Tabela de origem do grafico (endereco gravado na meta, coluna C, sem o "="), resolvida na pasta do grafico; Nothing se nao resolver.
Public Function Block_SourceRange(cht As Chart) As Range
    Dim blk As Range, a As String, bang As Long, sh As String
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    a = CStr(blk.Cells(1, 3).Value)
    bang = InStrRev(a, "!")
    If bang < 2 Then Exit Function
    sh = Left$(a, bang - 1)
    If Left$(sh, 1) = "'" And Right$(sh, 1) = "'" Then sh = Replace(Mid$(sh, 2, Len(sh) - 2), "''", "'")
    Set Block_SourceRange = Registry_Workbook(cht).Worksheets(sh).Range(Mid$(a, bang + 1))
    Exit Function
Fail:
    Set Block_SourceRange = Nothing
End Function
```

- [ ] **Passo 5: Colunas por cabeçalho** — em `modChartStacked`:

```vb
' Coluna Pos (posicao da categoria) e coluna VLine (linha de valor): achadas pelo cabecalho (a Fase 3 pode acrescentar colunas a direita);
' sem cabecalho (bloco de versao anterior), as duas ultimas colunas, como antes.
Public Function Stacked_PosCol(blk As Range) As Long
    Stacked_PosCol = Block_FindCol(blk, "Pos")
    If Stacked_PosCol = 0 Then Stacked_PosCol = blk.Columns.Count - 1
End Function

Public Function Stacked_VLineCol(blk As Range) As Long
    Stacked_VLineCol = Block_FindCol(blk, "VLine")
    If Stacked_VLineCol = 0 Then Stacked_VLineCol = blk.Columns.Count
End Function
```

Em `modChartArea`:

```vb
' Numero de series de dados: cabecalhos Val1, Val2, ... contiguos a partir da coluna 4 (a Fase 3 estende o bloco a direita; a largura
' ja nao diz quantas series ha). Reserva (bloco sem cabecalho Val1): (Columns.Count - 5) \ 2, como antes.
Public Function Area_SeriesCount(blk As Range) As Long
    Dim j As Long, body As Range
    Set body = Block_Body(blk)
    For j = 4 To body.Columns.Count
        If CStr(body.Cells(1, j).Value) <> "Val" & (j - 3) Then Exit For
        Area_SeriesCount = Area_SeriesCount + 1
    Next
    If Area_SeriesCount = 0 Then Area_SeriesCount = (blk.Columns.Count - 5) \ 2
End Function

Public Function Area_PosCol(blk As Range) As Long
    Area_PosCol = Block_FindCol(blk, "Pos")
    If Area_PosCol = 0 Then Area_PosCol = blk.Columns.Count - 1
End Function

Public Function Area_VLineCol(blk As Range) As Long
    Area_VLineCol = Block_FindCol(blk, "VLine")
    If Area_VLineCol = 0 Then Area_VLineCol = blk.Columns.Count
End Function
```

Ainda em `modChartArea`, no `Area_BuildBlock`, troque o laço que grava `"Val" & j` e `"Mid" & j` na mesma volta por dois laços (com a contagem por cabeçalho, `Area_MidCol` chama `Area_SeriesCount`, que na mesma volta ainda não vê o `Val` seguinte: `Mid1` cairia na coluna do `Val2` e a coluna do `Mid1` ficaria sem cabeçalho):

```vb
    For j = 1 To ns: body.Cells(1, Area_ValCol(blk, j)).Value = "Val" & j: Next
    For j = 1 To ns: body.Cells(1, Area_MidCol(blk, j)).Value = "Mid" & j: Next
```

- [ ] **Passo 6: Famílias novas nos pontos comuns** — `modLabels`: em `Labels_RefreshPositions`, depois da linha do Gantt, `If Kinds_Family(kind) = "scatter" Then GoTo Done`; em `Labels_ApplyCellFormat`, depois da linha do Gantt, `If Kinds_Family(kind) = "scatter" Then Exit Sub`. `modChartStyle`: em `Style_ApplyNumberFormat`, logo após o ramo do Gantt, `If Kinds_Family(Registry_ChartType(cht)) = "scatter" Then Exit Sub` e `If Kinds_Family(Registry_ChartType(cht)) = "butterfly" Then Exit Sub` (a Tarefa 1 troca a saída da borboleta por `Bfly_ApplyNumberFormat cht: Exit Sub`); acrescente ao módulo:

```vb
' RGB de uma cor da paleta (Style_Palette): o proprio valor ou, na paleta "tema" (-1..-6 = destaque 1..6), a cor do tema da pasta do
' grafico. Para marcadores XY (dispersao, bolhas, marcos), cujo Format.Fill nao aceita a cor (medido na Fase 2).
Public Function Style_PaletteRGB(cht As Chart, v As Variant) As Long
    On Error GoTo Fail
    If v >= 0 Then
        Style_PaletteRGB = CLng(v)
    Else
        Style_PaletteRGB = Registry_Workbook(cht).Theme.ThemeColorScheme.Colors(msoThemeAccent1 + (-v - 1)).RGB
    End If
    Exit Function
Fail:
    Style_PaletteRGB = RGB(68, 114, 196)
End Function
```

- [ ] **Passo 7: Tutorial (matriz) e Test-Docs** — em `docs/tutorial/tutorial.html`, nas duas tabelas `class="mx"` de `#matriz`, acrescente após a linha `data-kind="doughnut"` as três linhas (células na ordem das colunas de cada tabela; `<td class="y">Sim</td>` ou `<td>—</td>`):
  - Tabela 1 (totals, labels, connectors, numfmt, palette, gap, legend, lineend): `butterfly` "Borboleta": Sim, Sim, —, Sim, Sim, Sim, Sim, —; `scatter` "Dispersão": —, Sim, —, —, Sim, —, Sim, —; `bubble` "Bolhas": —, Sim, —, —, Sim, —, Sim, —.
  - Tabela 2 (ind100, vline, secondary, pieout, cagr, diff, bracket, samescale): `butterfly`: —, —, —, —, —, —, —, Sim; `scatter`: todas —; `bubble`: todas —.
  - Tabela 3 nova, depois da segunda `</div>` e antes do `<ul>` (`<h4>Arredondamento, corte e análise</h4>` + `<div class="tw" tabindex="0" role="region" aria-label="Tabela do que cada tipo aceita: arredondamento, corte e análise (em telas estreitas, role para os lados)">`), cabeçalhos `data-f="roundsum"` "Rótulos com soma consistente", `data-f="break"` "Corte de eixo", `data-f="trend"` "Linha de tendência", `data-f="quad"` "Quadrantes (médias)" (cada um em `<span class="k">` até a Tarefa 8), 20 linhas: roundsum Sim em stacked, stacked100, barstacked, barstacked100, area, area100; break Sim em stacked, clustered, barstacked, barclustered; trend e quad Sim em scatter e bubble; o resto —. Os rótulos das linhas novas: "Borboleta", "Dispersão", "Bolhas" (os mesmos da faixa da Tarefa 8).
  - `tests/ps/Test-Docs.ps1`: `Check ($allKinds.Count -eq 20) "20 tipos..."` e `Check ($matrix.Count -eq 20) "20 recursos..."`; e troque `Check ($bytes -le 122880) "tutorial com até 120 KB ($bytes bytes)"` por `Check ($bytes -le 163840) "tutorial com até 160 KB ($bytes bytes)"` (o tutorial tem ~112 KB; esta tarefa e a Tarefa 8 acrescentam ~8 KB e o limite antigo estouraria antes da Tarefa 11). Os rótulos novos ("Rótulos com soma consistente" etc.) ainda não estão na faixa: o `Test-Docs` compara o texto de cada cabeçalho da matriz com `$featureLabel` (lido de `Kinds_FeatureLabel`), não com a faixa; já todo `<span class="ui">` entra na lista `$ui` e **é** conferido contra os rótulos da faixa (`$known`). Por isso use `<span class="k">` (sem a classe `ui`) nos rótulos das três linhas e dos quatro cabeçalhos novos. A Tarefa 8 troca para `ui` os que viram controles (as três linhas e os cabeçalhos `roundsum`, `trend` e `quad`); o cabeçalho `data-f="break"` ("Corte de eixo") continua `k`, porque a faixa só tem "Corte de eixo..." e "Remover corte de eixo".

- [ ] **Passo 8: GREEN e suíte completa** — filtro acima verde; depois `pwsh C:\temp\tqq-f3\build\run-tests.ps1 -TimeoutSec 900` verde (inclusive `Test-Docs`). `Test_Core_AllKinds_InsertTogglesOptionsRebuild` continua com os 17 tipos (as Tarefas 1 e 2 acrescentam os novos).

- [ ] **Passo 9: Commit** — `git add src/core/modules/modChartKinds.bas src/core/modules/modCalcSheet.bas src/core/modules/modChartStacked.bas src/core/modules/modChartArea.bas src/core/modules/modLabels.bas src/core/modules/modChartStyle.bas docs/tutorial/tutorial.html tests/ps/Test-Docs.ps1 tests/vba/core/TestKinds.bas tests/vba/core/TestCalcSheet.bas tests/vba/core/TestCoreApi.bas` e `git commit -m "feat: matriz 1.2.0 (borboleta, dispersao, bolhas; soma consistente, corte de eixo, tendencia, quadrantes), extensao de bloco e colunas por cabecalho"` (+ trailer).

**Verificação no Excel real:** nenhuma mudança visível (os tipos antigos continuam iguais); inserir Colunas empilhadas e Área e ligar a Linha de valor (média) continua funcionando (colunas Pos/VLine achadas pelo cabeçalho).

---
### Tarefa 1: Borboleta (`butterfly`) — modelo: most-capable

**Files:**
- Create: `src/core/modules/modChartButterfly.bas`
- Modify: `src/core/modules/modChartKinds.bas` (`Kinds_Insert`: `Case "butterfly"` → `Bfly_Insert(ws, data)`), `src/core/modules/modChartStyle.bas` (`Style_ApplyNumberFormat`: família `butterfly` → `Bfly_ApplyNumberFormat cht: Exit Sub`), `src/core/modules/modAnnotations.bas` (`Ann_GetCategoryData`: `Case "butterfly"`), `src/core/modules/modCoreApi.bas` (`ApplyToggle` totals/labels na família `butterfly`, `PALETTE_FAMILIES` + `Case "butterfly": Bfly_ApplyColors`, `Core_SafeRefresh` chama `Bfly_RefreshAfterDataChange`)
- Test: `tests/vba/core/TestButterfly.bas` (novo), `tests/vba/core/TestCoreApi.bas` (`FixtureFor` + lista de `Test_Core_AllKinds_InsertTogglesOptionsRebuild`; o teste temporário `Test_Core_Insert_NewKindsNotReady_NotifiesWithoutChart` passa a testar só `scatter` e `bubble`)

**Interfaces:**
- Consumes: `CalcSheet_AllocateBlock`, `Block_Body`, `Block_Ref`, `Chart_AddHelperSeries`, `Chart_SetupSecondaryAxes`, `Chart_HideHelperLegendEntries`, `Style_ApplyBase`, `Style_ApplySeriesColors`, `Style_ContrastFont`, `Style_NumberFormat`, `Style_AssignNumberFormat`, `Labels_LinkPoints`, `Labels_RefreshPositions`, `Labels_ApplyCellFormat`, `Labels_EstimateWidth`, `Stacked_EnsureRoomAfterChange`, `Geom_ChartValueToX`, `Registry_*`, `Util_*`, `Kinds_*`.
- Produces (`modChartButterfly`): constantes `BF_COL_CAT=1, BF_COL_TOTAL=2, BF_COL_TOTALPOS=3, BF_COL_A=4, BF_COL_B=5, BF_COL_GAP=6, BF_COL_NEGGAP=7, BF_COL_NEGA=8, BF_COL_CATPOS=9, BF_COL_POS=10`; `Bfly_BuildBlock(wb, id, data) As Range`; `Bfly_Insert(ws, data) As ChartObject` (6011 "A borboleta precisa de exatamente 2 colunas de valores (lado esquerdo e lado direito)." antes de alocar); `Bfly_SetLabels(cht, on)`; `Bfly_SetTotals(cht, on)`; `Bfly_ApplyColors(cht, paletteName)`; `Bfly_ApplyNumberFormat(cht)`; `Bfly_AbsFormat(fmt) As String` (pura: `sec0;sec0;`); `Bfly_RefreshGap(cht)`; `Bfly_RefreshAfterDataChange(cht)` (só na família; nunca propaga erro); `Bfly_SideSeries(cht, side As Long) As Series` (1 = esquerda, 2 = direita: a k-ésima série não auxiliar).
- Opções gravadas no Insert: `numfmt=int, palette=padrao, gap=50, n, legend=right, labels=1, totals=1`. Meta do bloco: `E = GapK` (número), `F = SymL`, `G = SymR` (fórmulas), `H:I = n/2`.
- Mensagens exatas: 6011 acima; recusas pela matriz (7001/7002) vêm dos textos existentes.
- Limitação conhecida (documentada na Tarefa 11; `Deco_SameScale` não muda nesta fase): `Mesma escala` entre uma borboleta e gráficos de outro tipo aplica a todos a união [mínimo, máximo]; a borboleta perde o centro e os outros ganham mínimo negativo. Use `Mesma escala` só entre borboletas (ou só entre os outros tipos).

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestButterfly.bas`:

```vb
Attribute VB_Name = "TestButterfly"
Option Explicit

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Faixa", "Mulheres", "Homens")
    ws.Range("A2:C2").Value = Array("0-14", 120, 130)
    ws.Range("A3:C3").Value = Array("15-29", 200, 210)
    ws.Range("A4:C4").Value = Array("30-59", 310, 290)
    ws.Range("A5:C5").Value = Array("60+", 90, 60)
    Set Fixture = ChartData_FromRange(ws.Range("A1:C5"))
End Function

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function TqqNameCount(wb As Workbook) As Long
    Dim nm As Name
    For Each nm In wb.Names
        If Left$(nm.Name, Len(ID_PREFIX)) = ID_PREFIX Then TqqNameCount = TqqNameCount + 1
    Next
End Function

' x (pt) de um valor do eixo de valores, por regra de tres sobre a area interna e a escala REAL (independe da producao).
Private Function XOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, xlPrimary)
    XOf = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (v - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
End Function

Public Sub Test_Bfly_Insert_SeriesBlockOptions()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range, body As Range, ax As Axis, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Bfly_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "butterfly", Registry_ChartType(cht)
    AssertEqual xlBarStacked, cht.ChartType
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 10, blk.Columns.Count: AssertEqual 5, body.Rows.Count
    AssertEqual "Pos", body.Cells(1, BF_COL_POS).Value: AssertEqual 3.5, body.Cells(2, BF_COL_POS).Value
    AssertEqual 120, body.Cells(2, BF_COL_A).Value: AssertEqual 130, body.Cells(2, BF_COL_B).Value
    AssertEqual 250, body.Cells(2, BF_COL_TOTAL).Value
    AssertEqual -120, body.Cells(2, BF_COL_NEGA).Value
    AssertTrue body.Cells(2, BF_COL_GAP).Value > 0, "vao positivo"
    AssertAlmost body.Cells(2, BF_COL_GAP).Value, body.Cells(5, BF_COL_GAP).Value, 0.000001, "mesmo vao em todas as linhas"
    AssertAlmost CDbl(body.Cells(2, BF_COL_GAP).Value) + 130, CDbl(body.Cells(2, BF_COL_TOTALPOS).Value), 0.000001, "TotalPos = Gap + B"
    AssertAlmost -CDbl(blk.Cells(1, 7).Value), CDbl(blk.Cells(1, 6).Value), 0.000001, "SymL = -SymR"
    AssertAlmost CDbl(body.Cells(2, BF_COL_GAP).Value) + 310, CDbl(blk.Cells(1, 7).Value), 0.000001, "SymR = Gap + maior valor"
    AssertEqual 2, CDbl(blk.Cells(1, 8).Value), "Y dos pontos de simetria = n/2"
    AssertEqual "tqq_gapL", cht.SeriesCollection(1).Name: AssertEqual "Mulheres", cht.SeriesCollection(2).Name
    AssertEqual "tqq_gapR", cht.SeriesCollection(3).Name: AssertEqual "Homens", cht.SeriesCollection(4).Name
    AssertEqual msoFalse, cht.SeriesCollection(1).Format.Fill.Visible: AssertEqual msoFalse, cht.SeriesCollection(3).Format.Fill.Visible
    For Each s In cht.SeriesCollection
        If s.Name = "tqq_sym" Or s.Name = "tqq_cat" Or s.Name = "tqq_total" Then AssertEqual xlSecondary, s.AxisGroup, s.Name
    Next
    AssertTrue Not Chart_FindSeries(cht, "tqq_sym") Is Nothing: AssertTrue Not Chart_FindSeries(cht, "tqq_cat") Is Nothing
    AssertTrue Not Chart_FindSeries(cht, "tqq_total") Is Nothing
    Set ax = cht.Axes(xlValue, xlPrimary)
    AssertTrue ax.MinimumScaleIsAuto And ax.MaximumScaleIsAuto, "escala automatica"
    AssertAlmost -ax.MaximumScale, ax.MinimumScale, 0.000001, "eixo simetrico"
    AssertEqual 4, cht.Axes(xlValue, xlSecondary).MaximumScale, "eixo Y secundario 0..n"
    AssertEqual xlTickLabelPositionNone, cht.Axes(xlCategory).TickLabelPosition, "nomes das categorias so no centro"
    AssertTrue cht.Axes(xlCategory).ReversePlotOrder
    AssertEqual 2, cht.Legend.LegendEntries.Count, "legenda so com Mulheres e Homens"
    AssertEqual "1", Registry_GetOption(cht, "totals"): AssertEqual "1", Registry_GetOption(cht, "labels")
    AssertEqual "50", Registry_GetOption(cht, "gap"): AssertEqual "4", Registry_GetOption(cht, "n"): AssertEqual "right", Registry_GetOption(cht, "legend")
    AssertTrue Util_ToDouble(blk.Cells(1, 5).Value) >= 0.15, "GapK >= minimo"
    CloseTemp wb
End Sub

Public Sub Test_Bfly_Geometry_CenterLabelsAndMirroredBars()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range, g As Double, pA As Point, pB As Point, lbl As DataLabel, i As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Bfly_Insert(d.Source.Worksheet, d).Chart
    Set body = Block_Body(Registry_Block(cht))
    ForceLayout cht
    g = CDbl(body.Cells(2, BF_COL_GAP).Value)
    For i = 1 To 4
        Set pA = Bfly_SideSeries(cht, 1).Points(i): Set pB = Bfly_SideSeries(cht, 2).Points(i)
        AssertAlmost XOf(cht, -g), pA.Left + pA.Width, 1.5, "fim da barra esquerda " & i & " em x(-gap)"
        AssertAlmost XOf(cht, -g - CDbl(body.Cells(i + 1, BF_COL_A).Value)), pA.Left, 1.5, "inicio da barra esquerda " & i
        AssertAlmost XOf(cht, g), pB.Left, 1.5, "inicio da barra direita " & i & " em x(gap)"
        Set lbl = Chart_FindSeries(cht, "tqq_cat").Points(i).DataLabel
        AssertAlmost XOf(cht, 0), lbl.Left + lbl.Width / 2, 1.5, "nome " & i & " centrado em x(0)"
        AssertTrue lbl.Top + lbl.Height / 2 > pA.Top - 1 And lbl.Top + lbl.Height / 2 < pA.Top + pA.Height + 1, "nome " & i & " na faixa da categoria"
        AssertEqual Util_CellLabel(body.Cells(i + 1, 1)), lbl.Text
        Set lbl = Chart_FindSeries(cht, "tqq_total").Points(i).DataLabel
        AssertTrue lbl.Left >= pB.Left + pB.Width - 1, "total " & i & " a direita da barra direita"
        AssertEqual Format$(body.Cells(i + 1, BF_COL_TOTAL).Value, "#,##0"), lbl.Text
    Next
    CloseTemp wb
End Sub

' Review Focus 1: nome longo e lados desiguais: o nome cabe no vao central e o eixo e simetrico.
Public Sub Test_Bfly_LongName_GapFitsAndAxisSymmetric()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, body As Range, g As Double, lbl As DataLabel, ax As Axis, gapPt As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A4").Value = "Trinta a cinquenta e nove": ws.Range("B4").Value = 600: ws.Range("C4").Value = 60
    Set cht = Bfly_Insert(ws, d).Chart
    Set body = Block_Body(Registry_Block(cht))
    ForceLayout cht
    g = CDbl(body.Cells(2, BF_COL_GAP).Value)
    gapPt = XOf(cht, g) - XOf(cht, -g)
    Set lbl = Chart_FindSeries(cht, "tqq_cat").Points(3).DataLabel
    AssertTrue gapPt >= lbl.Width + 6, "vao (" & gapPt & " pt) cabe o nome (" & lbl.Width & " pt)"
    AssertAlmost XOf(cht, 0), lbl.Left + lbl.Width / 2, 1.5, "nome longo centrado"
    Set ax = cht.Axes(xlValue, xlPrimary)
    AssertAlmost -ax.MaximumScale, ax.MinimumScale, 0.000001, "eixo simetrico com 600 x 60"
    AssertTrue ax.MaximumScale >= 600 + g, "o lado direito tambem alcanca 600 + vao"
    AssertEqual Format$(660, "#,##0"), Chart_FindSeries(cht, "tqq_total").Points(3).DataLabel.Text
    CloseTemp wb
End Sub

' O vao cresce no refresh quando um nome fica longo DEPOIS de inserir (Core_SafeRefresh -> Bfly_RefreshAfterDataChange -> Bfly_RefreshGap).
Public Sub Test_Bfly_NameChange_GapGrowsOnRefresh()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, body As Range, g0 As Double, g As Double, lbl As DataLabel
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Bfly_Insert(ws, d).Chart
    Set body = Block_Body(Registry_Block(cht))
    g0 = CDbl(body.Cells(2, BF_COL_GAP).Value)
    ws.Range("A4").Value = "Trinta a cinquenta e nove"
    Core_SafeRefresh cht
    ForceLayout cht
    g = CDbl(body.Cells(2, BF_COL_GAP).Value)
    AssertTrue g > g0, "o vao cresceu no refresh (" & g0 & " -> " & g & ")"
    Set lbl = Chart_FindSeries(cht, "tqq_cat").Points(3).DataLabel
    AssertTrue XOf(cht, g) - XOf(cht, -g) >= lbl.Width, "o nome longo cabe no vao depois do refresh (" & lbl.Width & " pt)"
    Teardown wb
End Sub

Public Sub Test_Bfly_Labels_AbsFormat_And_NumFmt()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, dec As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C5").Select
    Core_Insert "butterfly"
    Set cht = Core_ActiveChart()
    AssertEqual "butterfly", Registry_ChartType(cht)
    AssertEqual "120", Bfly_SideSeries(cht, 1).Points(1).DataLabel.Text, "lado esquerdo sem sinal"
    AssertEqual "130", Bfly_SideSeries(cht, 2).Points(1).DataLabel.Text
    AssertEqual xlLabelPositionInsideEnd, Bfly_SideSeries(cht, 1).DataLabels.Position
    AssertEqual "#,##0;#,##0;", Bfly_AbsFormat("#,##0;-#,##0;")
    AssertEqual "0.0%;0.0%;", Bfly_AbsFormat("0.0%;-0.0%;")
    Core_SetOption "numfmt", "dec1"
    dec = Application.International(xlDecimalSeparator)
    AssertEqual "120" & dec & "0", Bfly_SideSeries(cht, 1).Points(1).DataLabel.Text, "dec1 no lado esquerdo, ainda sem sinal"
    AssertEqual "250" & dec & "0", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "total com dec1"
    Core_SetToggle "labels", False
    AssertFalse Bfly_SideSeries(cht, 1).HasDataLabels: AssertFalse Bfly_SideSeries(cht, 2).HasDataLabels
    AssertTrue Chart_FindSeries(cht, "tqq_cat").Points(1).HasDataLabel, "nomes das categorias nao sao 'Rotulos'"
    Core_SetToggle "labels", True
    AssertEqual "120" & dec & "0", Bfly_SideSeries(cht, 1).Points(1).DataLabel.Text, "religar mantem o formato"
    Teardown wb
End Sub

Public Sub Test_Bfly_Totals_Legend_Palette_Gap()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C5").Select
    Core_Insert "butterfly"
    Set cht = Core_ActiveChart()
    Core_SetToggle "totals", False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing: AssertEqual "0", Registry_GetOption(cht, "totals")
    Core_SetToggle "totals", True
    AssertTrue Not Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual 2, cht.Legend.LegendEntries.Count, "religar os totais nao devolve auxiliares a legenda"
    AssertTrue Chart_FindSeries(cht, "tqq_total").DataLabels.Font.Bold
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), Bfly_SideSeries(cht, 1).Format.Fill.ForeColor.RGB
    AssertEqual RGB(127, 127, 127), Bfly_SideSeries(cht, 2).Format.Fill.ForeColor.RGB
    AssertEqual RGB(255, 255, 255), Bfly_SideSeries(cht, 1).DataLabels.Font.Color, "fonte por contraste no escuro"
    AssertEqual msoFalse, cht.SeriesCollection(1).Format.Fill.Visible, "o vao continua invisivel apos a paleta"
    Core_SetOption "gap", "80"
    AssertEqual 80, cht.ChartGroups(1).GapWidth
    Core_SetOption "legend", "none"
    AssertFalse cht.HasLegend
    Core_ClearLastNotice
    Core_SetOption "legend", "lineend"
    AssertEqual "Rótulos no fim da linha só existem em gráficos de linha e combinação.", Core_LastNotice()
    Teardown wb
End Sub

Public Sub Test_Bfly_Refusals_And_Rollback()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, en As Long, ed As String, nm As Name, coll As New Collection, skipped As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("D1").Value = "Outro": ws.Range("D2:D5").Value = 1
    ws.Range("A1:D5").Select
    Core_ClearLastNotice
    Core_Insert "butterfly"
    AssertEqual "A borboleta precisa de exatamente 2 colunas de valores (lado esquerdo e lado direito).", Core_LastNotice()
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    ws.Range("D1:D5").ClearContents
    Set cht = Bfly_Insert(ws, d).Chart
    On Error Resume Next
    Ann_AddCagr cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    AssertEqual vbObjectError + 7001, en: AssertEqual "A seta CAGR não está disponível para este tipo de gráfico.", ed
    Deco_SetValueLine cht, True
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    coll.Add cht
    AssertEqual 0, Deco_SameScale(coll, skipped): AssertEqual 0, skipped, "borboleta tem escala livre (um grafico so nao fixa)"
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Bfly_Insert ws, d
    en = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, en
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    Util_FailAfterChartForTest = True
    On Error Resume Next
    Bfly_Insert ws, d
    en = Err.Number
    On Error GoTo 0
    Util_FailAfterChartForTest = False
    AssertEqual vbObjectError + 9998, en
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    CloseTemp wb
End Sub

' Sem o ThinQQ (nenhum refresh): mudar um valor mantem o eixo simetrico e o vao proporcional (so formulas e series nativas).
Public Sub Test_Bfly_DataChange_LiveWithoutThinQQ()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, ax As Axis, body As Range, g0 As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Bfly_Insert(ws, d).Chart
    Set body = Block_Body(Registry_Block(cht)): Set ax = cht.Axes(xlValue, xlPrimary)
    g0 = CDbl(body.Cells(2, BF_COL_GAP).Value)
    ws.Range("B4").Value = 620
    cht.Refresh
    AssertAlmost -ax.MaximumScale, ax.MinimumScale, 0.000001, "eixo continua simetrico"
    AssertTrue ax.MaximumScale >= 620, "escala acompanha o dado novo"
    AssertTrue CDbl(body.Cells(2, BF_COL_GAP).Value) > g0, "vao cresce com o maximo (formula)"
    AssertEqual Format$(910, "#,##0"), Chart_FindSeries(cht, "tqq_total").Points(3).DataLabel.Text
    CloseTemp wb
End Sub
```

Em `tests/vba/core/TestCoreApi.bas`: em `FixtureFor`, antes do `Case Else`:

```vb
        Case "butterfly"
            ws.Range("A1:C1").Value = Array("Faixa", "Mulheres", "Homens")
            ws.Range("A2:C2").Value = Array("0-14", 120, 130): ws.Range("A3:C3").Value = Array("15-29", 200, 210)
            ws.Range("A4:C4").Value = Array("30-59", 310, 290)
            Set FixtureFor = ws.Range("A1:C4")
```
Na lista `kinds` de `Test_Core_AllKinds_InsertTogglesOptionsRebuild` acrescente `"butterfly"`. Em `Test_Core_Insert_NewKindsNotReady_NotifiesWithoutChart` troque `Array("butterfly", "scatter", "bubble")` por `Array("scatter", "bubble")`.

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-f3\build\run-tests.ps1 -Target core -Filter "TestButterfly|AllKinds|NewKindsNotReady" -TimeoutSec 400` → saída 3.

- [ ] **Passo 3: Implementar `modChartButterfly.bas`**:

```vb
Attribute VB_Name = "modChartButterfly"
Option Explicit
' Borboleta (kind butterfly, familia butterfly, horizontal): barras EMPILHADAS com dois lados espelhados. Series na ordem tqq_gapL (-Gap,
' invisivel), <lado esquerdo> (-A), tqq_gapR (+Gap, invisivel), <lado direito> (+B): os negativos empilham para a esquerda e os positivos
' para a direita na ordem das series (medido), o que deixa um vao central de 2*Gap onde ficam os nomes das categorias (auxiliar XY
' tqq_cat em x = 0, rotulos ligados a coluna 1, posicao central). Os totais (A + B) sao a auxiliar XY tqq_total no fim da barra direita
' (x = Gap + B). A escala do eixo de valores e AUTOMATICA e simetrica: a auxiliar XY tqq_sym tem dois pontos em x = -(Gap + max) e
' +(Gap + max) (linha de meta do bloco) e os pontos XY do grupo secundario entram na escala automatica do eixo primario (medido na sonda):
' sem o ThinQQ a simetria continua viva. O lado esquerdo e desenhado com valores negativos, mas os rotulos usam o formato "absoluto"
' (secao negativa igual a positiva: Bfly_AbsFormat).
' Bloco (corpo): linha 1 = cabecalhos; n linhas: 1 Categoria | 2 Total (A+B) | 3 TotalPos (Gap+B) | 4 A | 5 B | 6 Gap | 7 NegGap | 8 NegA |
' 9 CatPos (0) | 10 Pos (n - i + 0,5). Meta (linha 1 do bloco): E = GapK (numero: fracao do maximo que vira o vao, gravado por Bfly_RefreshGap),
' F = SymL, G = SymR (formulas), H:I = n/2 (Y dos pontos de simetria). Gap = GapK * MAX(A, B) em todas as linhas.
Public Const BF_COL_CAT As Long = 1
Public Const BF_COL_TOTAL As Long = 2
Public Const BF_COL_TOTALPOS As Long = 3
Public Const BF_COL_A As Long = 4
Public Const BF_COL_B As Long = 5
Public Const BF_COL_GAP As Long = 6
Public Const BF_COL_NEGGAP As Long = 7
Public Const BF_COL_NEGA As Long = 8
Public Const BF_COL_CATPOS As Long = 9
Public Const BF_COL_POS As Long = 10
Private Const BF_COLS As Long = 10
Private Const BF_META_GAPK As Long = 5
Private Const BF_META_SYML As Long = 6
Private Const BF_META_SYMR As Long = 7
Private Const BF_META_SYMY As Long = 8
Private Const BF_GAPK_MIN As Double = 0.15       ' vao minimo: 15% do maior valor de cada lado
Private Const BF_LABEL_MARGIN As Double = 8      ' respiro (pt) alem da largura estimada do maior nome
Private Const BF_ERR_COLS As String = "A borboleta precisa de exatamente 2 colunas de valores (lado esquerdo e lado direito)."
Private mGapBusy As Boolean                      ' Bfly_RefreshGap grava uma celula: o evento Calculate do grafico reentra em Core_SafeRefresh

Private Sub Bfly_CheckData(data As clsChartData)
    If data.SeriesCount <> 2 Then Err.Raise vbObjectError + 6011, "Bfly_Insert", BF_ERR_COLS
End Sub

Public Function Bfly_BuildBlock(wb As Workbook, id As String, data As clsChartData) As Range
    Dim blk As Range, body As Range, n As Long, i As Long, r As Long, sA As String, sB As String, a As String, b As String, g As String
    Dim allAB As String, gapK As String, hdr As Variant, k As Long
    Bfly_CheckData data
    n = data.CategoryCount
    Set blk = CalcSheet_AllocateBlock(wb, id, "butterfly", Mid$(ChartData_ExternalAddress(data.Source), 2), n + 1, BF_COLS)
    Set body = Block_Body(blk)
    hdr = Array("Categoria", "Total", "TotalPos", "A", "B", "Gap", "NegGap", "NegA", "CatPos", "Pos")
    For k = 0 To BF_COLS - 1
        body.Cells(1, k + 1).Value = hdr(k)
    Next
    allAB = body.Cells(2, BF_COL_A).Resize(n, 2).Address(True, True)       ' mesma planilha: sem nome de planilha
    gapK = blk.Cells(1, BF_META_GAPK).Address(True, True)
    For i = 1 To n
        r = i + 1
        sA = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 2)), 2)
        sB = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 3)), 2)
        a = body.Cells(r, BF_COL_A).Address(False, False): b = body.Cells(r, BF_COL_B).Address(False, False): g = body.Cells(r, BF_COL_GAP).Address(False, False)
        body.Cells(r, BF_COL_CAT).Formula = ChartData_ExternalAddress(data.Source.Cells(r, 1))
        body.Cells(r, BF_COL_A).Formula = "=IF(ISNUMBER(" & sA & "),MAX(0," & sA & "),0)"      ' negativo, texto e erro valem 0
        body.Cells(r, BF_COL_B).Formula = "=IF(ISNUMBER(" & sB & "),MAX(0," & sB & "),0)"
        body.Cells(r, BF_COL_TOTAL).Formula = "=" & a & "+" & b
        body.Cells(r, BF_COL_GAP).Formula = "=" & gapK & "*MAX(" & allAB & ")"
        body.Cells(r, BF_COL_TOTALPOS).Formula = "=" & g & "+" & b
        body.Cells(r, BF_COL_NEGGAP).Formula = "=-" & g
        body.Cells(r, BF_COL_NEGA).Formula = "=-" & a
        body.Cells(r, BF_COL_CATPOS).Formula = "=0"
        body.Cells(r, BF_COL_POS).Formula = "=" & (n - i) & ".5"           ' independe do separador decimal da localidade
    Next
    blk.Cells(1, BF_META_GAPK).Value = BF_GAPK_MIN
    blk.Cells(1, BF_META_SYML).Formula = "=-(" & body.Cells(2, BF_COL_GAP).Address(True, True) & "+MAX(" & allAB & "))"
    blk.Cells(1, BF_META_SYMR).Formula = "=" & body.Cells(2, BF_COL_GAP).Address(True, True) & "+MAX(" & allAB & ")"
    blk.Cells(1, BF_META_SYMY).Value = n / 2
    blk.Cells(1, BF_META_SYMY + 1).Value = n / 2
    Set Bfly_BuildBlock = blk
End Function

Private Function AddBar(cht As Chart, name As String, valueCells As Range, cats As String) As Series
    Dim s As Series
    Set s = cht.SeriesCollection.NewSeries
    s.Name = name
    s.Values = Block_Ref(valueCells)
    s.XValues = cats
    Set AddBar = s
End Function

Public Function Bfly_Insert(ws As Worksheet, data As clsChartData) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, anchor As Range, blk As Range, body As Range, n As Long, cats As String
    Dim en As Long, es As String, ed As String
    Bfly_CheckData data                              ' recusa (6011) antes de alocar: nada a desfazer
    On Error GoTo Fail
    id = Registry_NewId()
    Set blk = Bfly_BuildBlock(ws.Parent, id, data)
    Util_TestHookAfterBlock
    Set body = Block_Body(blk)
    n = data.CategoryCount
    cats = ChartData_ExternalAddress(data.CategoryCells())
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    AddBar cht, "tqq_gapL", body.Cells(2, BF_COL_NEGGAP).Resize(n, 1), cats
    AddBar cht, ChartData_ExternalAddress(data.SeriesNameCell(1)), body.Cells(2, BF_COL_NEGA).Resize(n, 1), cats
    AddBar cht, "tqq_gapR", body.Cells(2, BF_COL_GAP).Resize(n, 1), cats
    AddBar cht, ChartData_ExternalAddress(data.SeriesNameCell(2)), body.Cells(2, BF_COL_B).Resize(n, 1), cats
    cht.ChartType = xlBarStacked
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlZero
    Style_ApplyBase cht                              ' categorias invertidas (1a no topo), eixo de valores oculto
    Util_TestHookAfterChart
    cht.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionNone     ' os nomes ficam no vao central (tqq_cat)
    cht.ChartGroups(1).GapWidth = 50
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight
    cht.Legend.Font.Size = 9
    Registry_SetOption cht, "numfmt", "int"
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "gap", "50"
    Registry_SetOption cht, "n", CStr(n)
    Registry_SetOption cht, "legend", "right"
    Bfly_ApplyColors cht, "padrao"
    ' auxiliares: a simetria primeiro (uma linha: Chart_SetupSecondaryAxes a fixaria em 0..1), depois nomes e totais (n linhas)
    Chart_AddHelperSeries cht, "tqq_sym", blk.Cells(1, BF_META_SYML).Resize(1, 2), blk.Cells(1, BF_META_SYMY).Resize(1, 2)
    Bfly_SetCategoryLabels cht
    Bfly_SetLabels cht, True
    Bfly_SetTotals cht, True
    Chart_SetupSecondaryAxes cht, n
    Chart_HideHelperLegendEntries cht
    Bfly_RefreshGap cht
    Set Bfly_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' k-esima serie NAO auxiliar: 1 = lado esquerdo (valores negativos), 2 = lado direito. Nothing se nao existe.
Public Function Bfly_SideSeries(cht As Chart, side As Long) As Series
    Dim s As Series, k As Long
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            k = k + 1
            If k = side Then Set Bfly_SideSeries = s: Exit Function
        End If
    Next
End Function

' Nomes das categorias no vao central: auxiliar XY tqq_cat (x = CatPos = 0, y = Pos) com o rotulo de cada ponto ligado a celula da
' categoria (DataLabel.Formula; nunca Labels_LinkPoints, que gravaria um formato numerico na celula e mostraria um ano como 2.022).
Private Sub Bfly_SetCategoryLabels(cht As Chart)
    Dim s As Series, body As Range, n As Long, i As Long
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    Set s = Chart_FindSeries(cht, "tqq_cat")
    If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_cat", body.Cells(2, BF_COL_CATPOS).Resize(n, 1), body.Cells(2, BF_COL_POS).Resize(n, 1))
    For i = 1 To n
        With s.Points(i)
            .HasDataLabel = True
            .DataLabel.Formula = Block_Ref(body.Cells(i + 1, BF_COL_CAT))
            .DataLabel.Position = xlLabelPositionCenter
            .DataLabel.Font.Size = 9
            .DataLabel.Font.Bold = False
            .DataLabel.Font.Color = RGB(64, 64, 64)
        End With
    Next
End Sub

' Rotulos de valor nativos nas duas series de dados (InsideEnd, fonte por contraste); o lado esquerdo mostra o valor absoluto.
Public Sub Bfly_SetLabels(cht As Chart, turnOn As Boolean)
    Dim k As Long, s As Series
    For k = 1 To 2
        Set s = Bfly_SideSeries(cht, k)
        If Not s Is Nothing Then
            s.HasDataLabels = turnOn
            If turnOn Then
                With s.DataLabels
                    .ShowValue = True: .ShowSeriesName = False: .ShowCategoryName = False
                    .Position = xlLabelPositionInsideEnd
                    .Font.Size = 9
                    .Font.Color = Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
                End With
            End If
        End If
    Next
    Registry_SetOption cht, "labels", IIf(turnOn, "1", "0")
    If turnOn Then Bfly_ApplyNumberFormat cht
End Sub

' Formato "absoluto": a secao negativa igual a positiva (o lado esquerdo e plotado em negativo). Pura.
Public Function Bfly_AbsFormat(fmt As String) As String
    Dim sec As String
    sec = Split(fmt & ";", ";")(0)
    Bfly_AbsFormat = sec & ";" & sec & ";"
End Function

' Formato numerico (opcao numfmt): lado esquerdo absoluto, lado direito normal, celulas do Total (Labels_ApplyCellFormat) e os rotulos
' ligados dos totais, ponto a ponto.
Public Sub Bfly_ApplyNumberFormat(cht As Chart)
    Dim fmt As String, s As Series, i As Long
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    Set s = Bfly_SideSeries(cht, 1)
    If Not s Is Nothing Then If s.HasDataLabels Then Style_AssignNumberFormat s.DataLabels, Bfly_AbsFormat(fmt)
    Set s = Bfly_SideSeries(cht, 2)
    If Not s Is Nothing Then If s.HasDataLabels Then Style_AssignNumberFormat s.DataLabels, fmt
    Labels_ApplyCellFormat cht, fmt
    Set s = Chart_FindSeries(cht, "tqq_total")
    If Not s Is Nothing Then
        For i = 1 To s.Points.Count
            If s.Points(i).HasDataLabel Then Style_AssignNumberFormat s.Points(i).DataLabel, fmt
        Next
    End If
End Sub

' Totais: auxiliar XY tqq_total em (TotalPos = Gap + B, Pos) com o rotulo ligado ao Total (A + B), negrito, a direita; a folga a
' direita e a mesma das barras (Stacked_EnsureRoomAfterChange le a coluna 3 = TotalPos).
Public Sub Bfly_SetTotals(cht As Chart, turnOn As Boolean)
    Dim s As Series, body As Range, n As Long
    Set s = Chart_FindSeries(cht, "tqq_total")
    If Not turnOn Then
        If Not s Is Nothing Then s.Delete
        Registry_SetOption cht, "totals", "0"
        Exit Sub
    End If
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_total", body.Cells(2, BF_COL_TOTALPOS).Resize(n, 1), body.Cells(2, BF_COL_POS).Resize(n, 1))
    Chart_HideHelperLegendEntries cht
    Labels_LinkPoints s, body.Cells(2, BF_COL_TOTAL).Resize(n, 1), Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int")), , , True
    Labels_RefreshPositions cht
    s.DataLabels.Font.Bold = True
    Registry_SetOption cht, "totals", "1"
    Stacked_EnsureRoomAfterChange cht
End Sub

' Cores: a paleta nas duas series de dados (Style_ApplySeriesColors pula as auxiliares), os vaos sempre invisiveis e a fonte dos rotulos
' por contraste com a cor nova.
Public Sub Bfly_ApplyColors(cht As Chart, paletteName As String)
    Dim s As Series, k As Long
    Style_ApplySeriesColors cht, paletteName
    For Each s In cht.SeriesCollection
        If s.Name = "tqq_gapL" Or s.Name = "tqq_gapR" Then
            s.Format.Fill.Visible = msoFalse
            s.Format.Line.Visible = msoFalse
        End If
    Next
    For k = 1 To 2
        Set s = Bfly_SideSeries(cht, k)
        If Not s Is Nothing Then
            If s.HasDataLabels Then s.DataLabels.Font.Color = Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
        End If
    Next
End Sub

' Vao central: GapK (meta E) x maior valor. O vao em pontos e medido na escala REAL do eixo (Geom_ChartValueToX) e comparado com a largura
' estimada do maior nome (Labels_EstimateWidth, que superestima de proposito) + margem: cresce se nao cabe (ate 3 passadas) e encolhe se
' sobra mais de 60%. So grava quando muda mais de 2% (gravar recalcula e dispara o Calculate do grafico, que reentra aqui: guarda mGapBusy).
' Sem o ThinQQ, GapK fica fixo e o vao continua proporcional ao maximo. Nunca propaga erro.
Public Sub Bfly_RefreshGap(cht As Chart)
    Dim blk As Range, body As Range, n As Long, i As Long, w As Double, lw As Double, need As Double, k As Double, k0 As Double, g As Double, gapPt As Double, pass As Long
    If mGapBusy Then Exit Sub
    mGapBusy = True
    On Error GoTo Done
    If Kinds_Family(Registry_ChartType(cht)) <> "butterfly" Then GoTo Done
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    For i = 1 To n
        w = Labels_EstimateWidth(Util_CellLabel(body.Cells(i + 1, BF_COL_CAT)))
        If w > lw Then lw = w
    Next
    need = lw + BF_LABEL_MARGIN
    k = Util_ToDouble(blk.Cells(1, BF_META_GAPK).Value)
    If k < BF_GAPK_MIN Then k = BF_GAPK_MIN
    k0 = k
    For pass = 1 To 3
        g = Util_ToDouble(body.Cells(2, BF_COL_GAP).Value)
        If g <= 0 Then Exit For
        gapPt = Geom_ChartValueToX(cht, g) - Geom_ChartValueToX(cht, -g)
        If gapPt <= 0 Then Exit For
        If gapPt < need - 0.5 Then
            k = k * (need / gapPt) * 1.05
        ElseIf gapPt > need * 1.6 And k > BF_GAPK_MIN Then
            k = k * (need * 1.2 / gapPt)
            If k < BF_GAPK_MIN Then k = BF_GAPK_MIN
        Else
            Exit For
        End If
        If Abs(k - Util_ToDouble(blk.Cells(1, BF_META_GAPK).Value)) <= 0.02 * k Then Exit For
        blk.Cells(1, BF_META_GAPK).Value = k
    Next
Done:
    If Err.Number <> 0 Then Errors_Log "Bfly_RefreshGap" & vbTab & Err.Number & vbTab & Err.Description
    mGapBusy = False
End Sub

' O que na borboleta depende dos DADOS ou do tamanho e muda depois de inserir (Core_SafeRefresh): o vao central. Nunca propaga erro.
Public Sub Bfly_RefreshAfterDataChange(cht As Chart)
    On Error Resume Next
    If Kinds_Family(Registry_ChartType(cht)) <> "butterfly" Then Exit Sub
    Bfly_RefreshGap cht
End Sub
```

- [ ] **Passo 4: Ligações** — `modChartKinds.Kinds_Insert`: `Case "butterfly": Set Kinds_Insert = Bfly_Insert(ws, data)`. `modChartStyle.Style_ApplyNumberFormat`: o ramo da família `butterfly` passa a `Bfly_ApplyNumberFormat cht: Exit Sub`. `modAnnotations.Ann_GetCategoryData`, antes do `Case Else`:

```vb
        Case "butterfly"
            ' total (A + B) e a posicao do rotulo do total (Gap + B), lidos do bloco: Deco_ScaleClipped compara o topo com o maximo fixo
            Dim bb As Range
            Set bb = Block_Body(Registry_Block(cht))
            For i = 1 To n
                If i <= bb.Rows.Count - 1 Then
                    values(i) = Util_ToDouble(bb.Cells(i + 1, BF_COL_TOTAL).Value)
                    tops(i) = Util_ToDouble(bb.Cells(i + 1, BF_COL_TOTALPOS).Value)
                End If
            Next
```
(mova o `Dim bb As Range` para a lista de declarações do procedimento, antes de qualquer instrução executável). `modCoreApi`: em `ApplyToggle`, `Case "totals"`: `Case "butterfly": Bfly_SetTotals c, value`; `Case "labels"`: `Case "butterfly": Bfly_SetLabels c, value`; `PALETTE_FAMILIES = ",stacked,combo,waterfall,area,pie,mekko,gantt,butterfly,"` e, no despacho da paleta, `Case "butterfly": Bfly_ApplyColors c, value`; em `Core_SafeRefresh`, depois da linha do `Pie_RefreshAfterDataChange`: `Bfly_RefreshAfterDataChange cht` com a mesma linha de log (`"Core_SafeRefresh/Bfly"`).

- [ ] **Passo 5: GREEN e suíte completa** — filtro verde; `pwsh C:\temp\tqq-f3\build\run-tests.ps1 -TimeoutSec 900` verde. Exporte o gráfico da fixture para PNG numa pasta própria (`cht.Export`) e olhe: nomes no centro, barras espelhadas, totais à direita, legenda com duas entradas; registre no relatório.

- [ ] **Passo 6: Commit** — `git add src/core/modules/modChartButterfly.bas src/core/modules/modChartKinds.bas src/core/modules/modChartStyle.bas src/core/modules/modAnnotations.bas src/core/modules/modCoreApi.bas tests/vba/core/TestButterfly.bas tests/vba/core/TestCoreApi.bas` e `git commit -m "feat: grafico borboleta (barras espelhadas, nomes no centro, totais, eixo simetrico vivo)"` (+ trailer).

**Verificação no Excel real:** tabela Faixa/Mulheres/Homens; Mais tipos > Borboleta (a Tarefa 8 cria o botão; até lá, pela janela Imediata do VBE: `Application.Run "'ThinQQ.core.xlam'!Core_Insert", "butterfly"` com a tabela selecionada); mudar um valor mantém a simetria; redimensionar o gráfico mantém o nome no vão.

---
### Tarefa 2: Dispersão (`scatter`) e bolhas (`bubble`) — modelo: most-capable

**Files:**
- Create: `src/core/modules/modChartScatter.bas`
- Modify: `src/core/modules/modChartKinds.bas` (`Kinds_Insert`: `Case "scatter"` → `Scatter_Insert(ws, data, kind)`), `src/core/modules/modCoreApi.bas` (`ApplyToggle`: chaves `trend`/`quad` e família `scatter` em `labels`; `PALETTE_FAMILIES` + `Case "scatter": Scatter_ApplyColors`; `ReapplyOptions`: `trend`, `quad` na lista; `Core_SafeRefresh` chama `Scatter_RefreshAfterDataChange`)
- Test: `tests/vba/core/TestScatter.bas` (novo), `tests/vba/core/TestCoreApi.bas` (`FixtureFor` scatter/bubble, lista de tipos e de toggles do all-kinds, `ChoicesSnapshot` com `trend`/`quad`; **remover** `Test_Core_Insert_NewKindsNotReady_NotifiesWithoutChart`)

**Interfaces:**
- Consumes: `CalcSheet_AllocateBlock`, `Block_Body`, `Block_Ref`, `Style_ApplyBase`, `Style_Palette`, `Style_PaletteRGB`, `Style_ContrastFont`, `Registry_*`, `Util_ToDouble`, `Kinds_*`, `ChartData_ExternalAddress`, `Errors_Log`.
- Produces (`modChartScatter`): `SC_COL_NAME=1, SC_COL_X=2, SC_COL_Y=3, SC_COL_SIZE=4, SC_COL_GROUP=5`; `Scatter_GroupCount(blk) As Long` = `(Columns.Count - 5) \ 3`; `Scatter_XgCol(blk, k)`, `Scatter_YgCol(blk, k)`, `Scatter_SgCol(blk, k)`; `Scatter_BuildBlock(wb, id, data, kind) As Range`; `Scatter_Insert(ws, data, kind) As ChartObject`; `Scatter_SetLabels(cht, on)`; `Scatter_ApplyColors(cht, paletteName)`; `Scatter_SetTrend(cht, on)`; `Scatter_SetQuad(cht, on)`; `Scatter_ApplyQuadrants(cht)`; `Scatter_HideTrendEntries(cht)`; `Scatter_RefreshAfterDataChange(cht)`; `Scatter_DataSeries(cht, k) As Series` (k-ésima série não auxiliar); `Scatter_SeriesCount(cht) As Long`.
- Mensagens exatas: 6012 `"Dispersão: a tabela precisa de nome, X e Y (e, nas bolhas, tamanho)."`; 6013 `"No máximo 6 grupos (coluna de grupo)."`. Opções gravadas no Insert: `palette=padrao, n, legend=right, labels=1, trend=0, quad=0`. Meta: `E = AvgX`, `F = AvgY`.
- Tabela: `Nome | X | Y [| Grupo]` (scatter) e `Nome | X | Y | Tamanho [| Grupo]` (bubble); grupo = valor da célula como texto (`CStr(Value2)`, o mesmo que o `=src&""` do bloco: datas e números formatados batem com a fórmula), vazio vira o grupo "(sem grupo)".

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestScatter.bas`:

```vb
Attribute VB_Name = "TestScatter"
Option Explicit

Private Function Fixture(wb As Workbook, Optional withGroup As Boolean = True, Optional bubble As Boolean = False) As clsChartData
    Dim ws As Worksheet, lastCol As String
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Produto", "Preço", "Margem")
    ws.Range("A2:C2").Value = Array("P1", 10, 0.2): ws.Range("A3:C3").Value = Array("P2", 20, 0.35)
    ws.Range("A4:C4").Value = Array("P3", 30, 0.25): ws.Range("A5:C5").Value = Array("P4", 40, 0.5)
    lastCol = "C"
    If bubble Then
        ws.Range("D1").Value = "Volume": ws.Range("D2:D5").Value = Application.Transpose(Array(50, 120, 80, 20)): lastCol = "D"
    End If
    If withGroup Then
        ws.Range(IIf(bubble, "E1", "D1")).Value = "Grupo"
        ws.Range(IIf(bubble, "E2:E5", "D2:D5")).Value = Application.Transpose(Array("A", "B", "A", "B")): lastCol = IIf(bubble, "E", "D")
    End If
    Set Fixture = ChartData_FromRange(ws.Range("A1:" & lastCol & "5"))
End Function

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function TqqNameCount(wb As Workbook) As Long
    Dim nm As Name
    For Each nm In wb.Names
        If Left$(nm.Name, Len(ID_PREFIX)) = ID_PREFIX Then TqqNameCount = TqqNameCount + 1
    Next
End Function

Private Function XOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlCategory, xlPrimary)
    XOf = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (v - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
End Function

Private Function YOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, xlPrimary)
    YOf = cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight * (ax.MaximumScale - v) / (ax.MaximumScale - ax.MinimumScale)
End Function

Public Sub Test_Scatter_Insert_BlockSeriesAxesLabels()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range, body As Range, sa As Series, sb As Series, v As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Scatter_Insert(d.Source.Worksheet, d, "scatter").Chart
    AssertEqual "scatter", Registry_ChartType(cht)
    AssertEqual xlXYScatter, cht.ChartType
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 11, blk.Columns.Count, "5 + 3 colunas por grupo (2 grupos)": AssertEqual 2, Scatter_GroupCount(blk)
    AssertEqual "Xg1", body.Cells(1, Scatter_XgCol(blk, 1)).Value: AssertEqual "Sg2", body.Cells(1, Scatter_SgCol(blk, 2)).Value
    AssertEqual 10, body.Cells(2, SC_COL_X).Value: AssertEqual 0.2, body.Cells(2, SC_COL_Y).Value: AssertEqual "A", body.Cells(2, SC_COL_GROUP).Value
    AssertEqual 10, body.Cells(2, Scatter_XgCol(blk, 1)).Value, "P1 e do grupo A"
    AssertTrue IsError(body.Cells(2, Scatter_XgCol(blk, 2)).Value), "P1 nao esta no grupo B"
    AssertEqual 20, body.Cells(3, Scatter_XgCol(blk, 2)).Value
    AssertAlmost 25, CDbl(blk.Cells(1, 5).Value), 0.000001, "AvgX": AssertAlmost 0.325, CDbl(blk.Cells(1, 6).Value), 0.000001, "AvgY"
    AssertEqual 2, cht.SeriesCollection.Count
    Set sa = cht.SeriesCollection(1): Set sb = cht.SeriesCollection(2)
    AssertEqual "A", sa.Name: AssertEqual "B", sb.Name
    AssertTrue InStr(sa.Formula, CALC_SHEET) > 0, "series apontam para o bloco"
    AssertEqual xlNotPlotted, cht.DisplayBlanksAs
    v = sa.Values
    AssertTrue IsError(v(2)) Or IsEmpty(v(2)), "ponto do outro grupo nao plotado"
    AssertEqual "P1", sa.Points(1).DataLabel.Text: AssertEqual "P2", sb.Points(2).DataLabel.Text
    AssertEqual xlLabelPositionRight, sa.Points(1).DataLabel.Position
    AssertTrue cht.Axes(xlCategory).HasTitle And cht.Axes(xlValue).HasTitle
    AssertEqual "Preço", cht.Axes(xlCategory).AxisTitle.Text: AssertEqual "Margem", cht.Axes(xlValue).AxisTitle.Text
    AssertEqual xlTickLabelPositionLow, cht.Axes(xlCategory).TickLabelPosition
    AssertEqual xlTickLabelPositionLow, cht.Axes(xlValue).TickLabelPosition
    AssertFalse cht.Axes(xlValue).HasMajorGridlines
    AssertTrue cht.HasLegend: AssertEqual 2, cht.Legend.LegendEntries.Count
    AssertEqual "0", Registry_GetOption(cht, "trend"): AssertEqual "0", Registry_GetOption(cht, "quad"): AssertEqual "4", Registry_GetOption(cht, "n")
    d.Source.Worksheet.Range("B1").Value = "Preço (R$)"
    AssertEqual "Preço (R$)", cht.Axes(xlCategory).AxisTitle.Text, "titulo do eixo ligado ao cabecalho"
    CloseTemp wb
End Sub

Public Sub Test_Scatter_Geometry_PointsAndLabels()
    Dim wb As Workbook, d As clsChartData, cht As Chart, p As Point, lbl As DataLabel
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Scatter_Insert(d.Source.Worksheet, d, "scatter").Chart
    ForceLayout cht
    Set p = cht.SeriesCollection(1).Points(1)                   ' P1 (10; 0,2)
    AssertAlmost XOf(cht, 10), p.Left + p.Width / 2, 2, "centro do marcador em x(10)"
    AssertAlmost YOf(cht, 0.2), p.Top + p.Height / 2, 2, "centro do marcador em y(0,2)"
    Set lbl = p.DataLabel
    AssertTrue lbl.Left >= p.Left + p.Width / 2, "nome a direita do ponto"
    AssertTrue Abs((lbl.Top + lbl.Height / 2) - (p.Top + p.Height / 2)) < 8, "nome na altura do ponto"
    Set p = cht.SeriesCollection(2).Points(4)                   ' P4 (40; 0,5)
    AssertAlmost XOf(cht, 40), p.Left + p.Width / 2, 2
    AssertAlmost YOf(cht, 0.5), p.Top + p.Height / 2, 2
    CloseTemp wb
End Sub

' Review Focus 2: um ponto troca de grupo e outro ganha texto em X.
Public Sub Test_Scatter_GroupChange_MovesPointAndLabel()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, sa As Series, sb As Series, va As Variant, vb As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    Set sa = cht.SeriesCollection(1): Set sb = cht.SeriesCollection(2)
    ws.Range("D2").Value = "B"                                  ' P1: A -> B
    ws.Range("B4").Value = "x"                                  ' P3: X invalido
    cht.Refresh
    va = sa.Values: vb = sb.Values
    AssertTrue IsError(va(1)) Or IsEmpty(va(1)), "P1 saiu da serie A"
    AssertAlmost 0.2, CDbl(vb(1)), 0.000001, "P1 entrou na serie B"
    AssertEqual "P1", sb.Points(1).DataLabel.Text, "o nome acompanha o ponto na serie B"
    AssertTrue IsError(va(3)) Or IsEmpty(va(3)), "P3 com X invalido nao e plotado"
    AssertTrue IsError(Block_Body(Registry_Block(cht)).Cells(4, SC_COL_X).Value), "X invalido vira NA() no bloco"
    ForceLayout cht
    AssertTrue sb.Points(1).DataLabel.Left > 0, "rotulo de P1 na serie B tem posicao"
    CloseTemp wb
End Sub

Public Sub Test_Scatter_Trend_And_Quadrants()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, blk As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    Set blk = Registry_Block(cht)
    Scatter_SetTrend cht, True
    AssertEqual 1, cht.SeriesCollection(1).Trendlines.Count: AssertEqual 1, cht.SeriesCollection(2).Trendlines.Count
    AssertEqual xlLinear, cht.SeriesCollection(1).Trendlines(1).Type
    AssertEqual 2, cht.Legend.LegendEntries.Count, "entradas das tendencias fora da legenda"
    AssertEqual "1", Registry_GetOption(cht, "trend")
    Scatter_SetTrend cht, False
    AssertEqual 0, cht.SeriesCollection(1).Trendlines.Count: AssertEqual "0", Registry_GetOption(cht, "trend")
    Scatter_SetQuad cht, True
    AssertEqual xlAxisCrossesCustom, cht.Axes(xlCategory).Crosses: AssertEqual xlAxisCrossesCustom, cht.Axes(xlValue).Crosses
    AssertAlmost CDbl(blk.Cells(1, 5).Value), CDbl(cht.Axes(xlCategory).CrossesAt), 0.000001, "linha vertical na media de X"
    AssertAlmost CDbl(blk.Cells(1, 6).Value), CDbl(cht.Axes(xlValue).CrossesAt), 0.000001, "linha horizontal na media de Y"
    AssertEqual xlTickLabelPositionLow, cht.Axes(xlValue).TickLabelPosition, "rotulos ficam na borda"
    ws.Range("B5").Value = 80                                   ' media de X passa de 25 para 35
    Scatter_RefreshAfterDataChange cht
    AssertAlmost 35, CDbl(cht.Axes(xlCategory).CrossesAt), 0.000001, "quadrante acompanha os dados no refresh"
    Scatter_SetQuad cht, False
    AssertEqual xlAxisCrossesAutomatic, cht.Axes(xlCategory).Crosses: AssertEqual "0", Registry_GetOption(cht, "quad")
    CloseTemp wb
End Sub

Public Sub Test_Bubble_Insert_SizesNegativeAndLabels()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, blk As Range, body As Range, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, True, True): Set ws = d.Source.Worksheet
    ws.Range("D5").Value = -20                                  ' tamanho negativo: fora
    Set cht = Scatter_Insert(ws, d, "bubble").Chart
    AssertEqual "bubble", Registry_ChartType(cht): AssertEqual xlBubble, cht.ChartType
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 50, body.Cells(2, SC_COL_SIZE).Value
    AssertTrue IsError(body.Cells(5, SC_COL_SIZE).Value), "tamanho negativo vira NA()"
    AssertTrue IsError(body.Cells(5, Scatter_XgCol(blk, 2)).Value), "ponto sem tamanho nao e plotado"
    Set s = cht.SeriesCollection(1)
    AssertTrue InStr(s.BubbleSizes, CALC_SHEET) > 0, "tamanhos no bloco"
    AssertEqual 60, cht.ChartGroups(1).BubbleScale
    AssertFalse cht.ChartGroups(1).ShowNegativeBubbles
    AssertEqual xlLabelPositionCenter, s.Points(1).DataLabel.Position
    AssertEqual "P1", s.Points(1).DataLabel.Text
    AssertEqual "Tam", body.Cells(1, SC_COL_SIZE).Value, "cabecalho fixo do bloco (nao o da tabela)"
    CloseTemp wb
End Sub

Public Sub Test_Scatter_Refusals_Rollback_NoGroup()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, en As Long, nm As Name, i As Long, k As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, False): Set ws = d.Source.Worksheet
    ws.Range("A1:B5").Select
    Core_ClearLastNotice
    Core_Insert "scatter"
    AssertEqual "Dispersão: a tabela precisa de nome, X e Y (e, nas bolhas, tamanho).", Core_LastNotice()
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    ws.Range("A1:C5").Select
    Core_ClearLastNotice
    Core_Insert "bubble"
    AssertEqual "Dispersão: a tabela precisa de nome, X e Y (e, nas bolhas, tamanho).", Core_LastNotice()
    AssertEqual 0, ws.ChartObjects.Count
    Set cht = Scatter_Insert(ws, d, "scatter").Chart            ' sem coluna de grupo: uma serie com o nome do cabecalho de Y
    AssertEqual 1, cht.SeriesCollection.Count: AssertEqual "Margem", cht.SeriesCollection(1).Name
    AssertEqual 8, Registry_Block(cht).Columns.Count
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    ws.Range("D1").Value = "Grupo"
    For i = 2 To 5: ws.Range("D" & i).Value = "G" & i: Next
    ws.Range("A6:D8").Value = Array("P5", 1, 1, "G6"): ws.Range("A7:D7").Value = Array("P6", 2, 2, "G7"): ws.Range("A8:D8").Value = Array("P7", 3, 3, "G8")
    Set d = ChartData_FromRange(ws.Range("A1:D8"))
    On Error Resume Next
    Scatter_Insert ws, d, "scatter"
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 6013, en, "7 grupos"
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    Set d = ChartData_FromRange(ws.Range("A1:D5"))
    For Each k In Array(9999, 9998)
        If k = 9999 Then Util_FailAfterBlockForTest = True Else Util_FailAfterChartForTest = True
        On Error Resume Next
        Scatter_Insert ws, d, "scatter"
        en = Err.Number
        On Error GoTo 0
        Util_FailAfterBlockForTest = False: Util_FailAfterChartForTest = False
        AssertEqual vbObjectError + k, en
        AssertEqual 0, ws.ChartObjects.Count
        For Each nm In wb.Names
            AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
        Next
    Next
    Teardown wb
End Sub

Public Sub Test_Core_Insert_Scatter_ThroughApi()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D5").Select
    Core_Insert "scatter"
    Set cht = Core_ActiveChart()
    AssertEqual "scatter", Registry_ChartType(cht)
    AssertTrue Core_Supports("trend") And Core_Supports("quad") And Core_Supports("labels") And Core_Supports("legend")
    AssertFalse Core_Supports("numfmt") Or Core_Supports("gap") Or Core_Supports("totals") Or Core_Supports("samescale")
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(1).MarkerBackgroundColor
    AssertEqual RGB(127, 127, 127), cht.SeriesCollection(2).MarkerBackgroundColor
    Core_ClearLastNotice
    Core_SetOption "numfmt", "dec1"
    AssertEqual "'Formato' não se aplica a este tipo de gráfico.", Core_LastNotice()
    Core_SetToggle "trend", True
    AssertTrue Core_GetToggle("trend"): AssertEqual 1, cht.SeriesCollection(1).Trendlines.Count
    Core_SetOption "legend", "none"
    AssertFalse cht.HasLegend
    Core_SetOption "legend", "right"
    AssertEqual 2, cht.Legend.LegendEntries.Count, "religar a legenda nao traz as tendencias de volta"
    Core_SetToggle "labels", False
    AssertFalse cht.SeriesCollection(1).HasDataLabels
    Core_SetToggle "quad", True
    AssertEqual xlAxisCrossesCustom, cht.Axes(xlValue).Crosses
    Teardown wb
End Sub
```

Em `tests/vba/core/TestCoreApi.bas`: `FixtureFor` ganha, antes do `Case Else`:

```vb
        Case "scatter"
            ws.Range("A1:D1").Value = Array("Produto", "Preço", "Margem", "Grupo")
            ws.Range("A2:D2").Value = Array("P1", 10, 0.2, "A"): ws.Range("A3:D3").Value = Array("P2", 20, 0.35, "B")
            ws.Range("A4:D4").Value = Array("P3", 30, 0.25, "A")
            If kind = "bubble" Then
                ws.Columns("D").Insert
                ws.Range("D1").Value = "Volume": ws.Range("D2:D4").Value = Application.Transpose(Array(50, 120, 80))
            End If
            Set FixtureFor = ws.Range("A1").CurrentRegion
```
Na lista `kinds` do all-kinds acrescente `"scatter", "bubble"`; na lista `toggles`, `"trend", "quad"`; em `ChoicesSnapshot`, `"trend", "quad"`. Remova `Test_Core_Insert_NewKindsNotReady_NotifiesWithoutChart`. (`OneRowFixtureFor` continua funcionando: `ws.Rows("3:" & rng.Rows.Count).Delete`.)

- [ ] **Passo 2: RED** — `-Filter "TestScatter|AllKinds"` → saída 3.

- [ ] **Passo 3: Implementar `modChartScatter.bas`**:

```vb
Attribute VB_Name = "modChartScatter"
Option Explicit
' Dispersao (kind scatter) e bolhas (kind bubble), familia scatter: dois eixos de valor (xlXYScatter / xlBubble), UMA SERIE POR GRUPO
' (coluna de grupo opcional, a ultima da tabela), rotulo de cada ponto ligado ao nome (DataLabel.Formula), titulos dos eixos ligados aos
' cabecalhos (AxisTitle.Formula), linha de tendencia nativa (trend) e quadrantes nas medias (quad: Axes(xlCategory).CrossesAt = media de X
' desenha a linha vertical; Axes(xlValue).CrossesAt = media de Y, a horizontal; medido). Bolhas nao combinam com outros tipos de serie
' (o Excel converte o grafico inteiro): por isso nao ha auxiliares.
' Tabela: Nome | X | Y [| Grupo] (dispersao) e Nome | X | Y | Tamanho [| Grupo] (bolhas).
' Bloco (corpo): linha 1 = cabecalhos; n linhas: 1 Nome | 2 X | 3 Y | 4 Tam | 5 Grupo | 6..5+G Xg_k | 6+G..5+2G Yg_k | 6+2G..5+3G Sg_k,
' com X/Y = IF(ISNUMBER(src),src,NA()), Tam = IF(ISNUMBER(src),IF(src>0,src,NA()),NA()) (dispersao: 1), Grupo = =src&"" (o valor como texto), e
' Xg_k = IF(AND(Grupo="<nome k>",ISNUMBER(X),ISNUMBER(Y),ISNUMBER(Tam)),X,NA()) (idem Yg, Sg). DisplayBlanksAs = xlNotPlotted: o ponto NA
' nao e desenhado e o seu rotulo tambem nao (medido). Meta: E = AvgX, F = AvgY (AGGREGATE ignora os NA()).
' Um grupo que so aparece DEPOIS de inserir nao tem serie: "Alterar dados" recria (documentado no tutorial).
Public Const SC_COL_NAME As Long = 1
Public Const SC_COL_X As Long = 2
Public Const SC_COL_Y As Long = 3
Public Const SC_COL_SIZE As Long = 4
Public Const SC_COL_GROUP As Long = 5
Private Const SC_FIXED As Long = 5
Private Const SC_META_AVGX As Long = 5
Private Const SC_META_AVGY As Long = 6
Private Const SC_MAX_GROUPS As Long = 6
Private Const SC_ERR_COLS As String = "Dispersão: a tabela precisa de nome, X e Y (e, nas bolhas, tamanho)."
Private Const SC_ERR_GROUPS As String = "No máximo 6 grupos (coluna de grupo)."
Private Const SC_NO_GROUP As String = "(sem grupo)"
Private Const SC_AXIS_COLOR As Long = 10921638   ' RGB(166, 166, 166)
Private Const SC_TEXT_COLOR As Long = 5855577    ' RGB(89, 89, 89)

Public Function Scatter_GroupCount(blk As Range) As Long
    Scatter_GroupCount = (blk.Columns.Count - SC_FIXED) \ 3
End Function

Public Function Scatter_XgCol(blk As Range, k As Long) As Long
    Scatter_XgCol = SC_FIXED + k
End Function

Public Function Scatter_YgCol(blk As Range, k As Long) As Long
    Scatter_YgCol = SC_FIXED + Scatter_GroupCount(blk) + k
End Function

Public Function Scatter_SgCol(blk As Range, k As Long) As Long
    Scatter_SgCol = SC_FIXED + 2 * Scatter_GroupCount(blk) + k
End Function

' Numero minimo de colunas de valores: 2 (X, Y) na dispersao, 3 (X, Y, tamanho) nas bolhas. Erro 6012 (aviso de Core_Insert).
Private Sub Scatter_CheckData(data As clsChartData, kind As String)
    If data.SeriesCount < IIf(kind = "bubble", 3, 2) Then Err.Raise vbObjectError + 6012, "Scatter_Insert", SC_ERR_COLS
End Sub

' Coluna da tabela (1 = categoria) que traz o grupo; 0 se a tabela nao tem coluna de grupo.
Private Function GroupColumn(data As clsChartData, kind As String) As Long
    Dim need As Long
    need = IIf(kind = "bubble", 4, 3)                ' colunas de valores obrigatorias + 1
    If data.SeriesCount >= need Then GroupColumn = need + 1
End Function

' Nomes distintos dos grupos na ordem em que aparecem (VALOR da celula como texto, igual ao =src&"" do bloco: o texto exibido de uma data
' ou de um numero formatado nao bateria com a formula; vazio = "(sem grupo)"); sem coluna de grupo, um unico grupo com
' o nome do cabecalho de Y. Mais de SC_MAX_GROUPS => 6013.
Private Function GroupNames(data As clsChartData, kind As String) As Variant
    Dim gc As Long, i As Long, nm As String, names() As String, g As Long, k As Long, found As Boolean, v As Variant
    gc = GroupColumn(data, kind)
    ReDim names(1 To SC_MAX_GROUPS)
    If gc = 0 Then
        names(1) = data.SeriesName(2): g = 1
    Else
        For i = 1 To data.CategoryCount
            v = data.Source.Cells(i + 1, gc).Value2
            If IsError(v) Then nm = "" Else nm = CStr(v)
            If Len(nm) = 0 Then nm = SC_NO_GROUP
            found = False
            For k = 1 To g
                If names(k) = nm Then found = True: Exit For
            Next
            If Not found Then
                g = g + 1
                If g > SC_MAX_GROUPS Then Err.Raise vbObjectError + 6013, "Scatter_Insert", SC_ERR_GROUPS
                names(g) = nm
            End If
        Next
        If g = 0 Then names(1) = SC_NO_GROUP: g = 1
    End If
    ReDim Preserve names(1 To g)
    GroupNames = names
End Function

Public Function Scatter_BuildBlock(wb As Workbook, id As String, data As clsChartData, kind As String) As Range
    Dim blk As Range, body As Range, n As Long, i As Long, r As Long, names As Variant, g As Long, k As Long, gc As Long
    Dim sx As String, sy As String, ss As String, sg As String, x As String, y As String, t As String, grp As String, lit As String, cond As String
    Scatter_CheckData data, kind
    names = GroupNames(data, kind)
    g = UBound(names)
    gc = GroupColumn(data, kind)
    n = data.CategoryCount
    Set blk = CalcSheet_AllocateBlock(wb, id, kind, Mid$(ChartData_ExternalAddress(data.Source), 2), n + 1, SC_FIXED + 3 * g)
    Set body = Block_Body(blk)
    body.Cells(1, SC_COL_NAME).Value = "Nome": body.Cells(1, SC_COL_X).Value = "X": body.Cells(1, SC_COL_Y).Value = "Y"
    body.Cells(1, SC_COL_SIZE).Value = "Tam": body.Cells(1, SC_COL_GROUP).Value = "Grupo"
    For k = 1 To g
        body.Cells(1, Scatter_XgCol(blk, k)).Value = "Xg" & k
        body.Cells(1, Scatter_YgCol(blk, k)).Value = "Yg" & k
        body.Cells(1, Scatter_SgCol(blk, k)).Value = "Sg" & k
    Next
    For i = 1 To n
        r = i + 1
        sx = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 2)), 2)
        sy = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 3)), 2)
        body.Cells(r, SC_COL_NAME).Formula = ChartData_ExternalAddress(data.Source.Cells(r, 1))
        body.Cells(r, SC_COL_X).Formula = "=IF(ISNUMBER(" & sx & ")," & sx & ",NA())"
        body.Cells(r, SC_COL_Y).Formula = "=IF(ISNUMBER(" & sy & ")," & sy & ",NA())"
        If kind = "bubble" Then
            ss = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 4)), 2)
            body.Cells(r, SC_COL_SIZE).Formula = "=IF(ISNUMBER(" & ss & "),IF(" & ss & ">0," & ss & ",NA()),NA())"      ' IF aninhado: AND propagaria o erro
        Else
            body.Cells(r, SC_COL_SIZE).Formula = "=1"
        End If
        If gc > 0 Then
            sg = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, gc)), 2)
            body.Cells(r, SC_COL_GROUP).Formula = "=" & sg & "&"""""
        Else
            body.Cells(r, SC_COL_GROUP).Formula = "=""" & Replace(CStr(names(1)), """", """""") & """"
        End If
        x = body.Cells(r, SC_COL_X).Address(False, False): y = body.Cells(r, SC_COL_Y).Address(False, False)
        t = body.Cells(r, SC_COL_SIZE).Address(False, False): grp = body.Cells(r, SC_COL_GROUP).Address(False, False)
        For k = 1 To g
            lit = Replace(CStr(names(k)), """", """""")
            If CStr(names(k)) = SC_NO_GROUP And gc > 0 Then lit = ""
            cond = "AND(" & grp & "=""" & lit & """,ISNUMBER(" & x & "),ISNUMBER(" & y & "),ISNUMBER(" & t & "))"
            body.Cells(r, Scatter_XgCol(blk, k)).Formula = "=IF(" & cond & "," & x & ",NA())"
            body.Cells(r, Scatter_YgCol(blk, k)).Formula = "=IF(" & cond & "," & y & ",NA())"
            body.Cells(r, Scatter_SgCol(blk, k)).Formula = "=IF(" & cond & "," & t & ",NA())"
        Next
    Next
    blk.Cells(1, SC_META_AVGX).Formula = "=IFERROR(AGGREGATE(1,6," & body.Cells(2, SC_COL_X).Resize(n, 1).Address(True, True) & "),0)"
    blk.Cells(1, SC_META_AVGY).Formula = "=IFERROR(AGGREGATE(1,6," & body.Cells(2, SC_COL_Y).Resize(n, 1).Address(True, True) & "),0)"
    Set Scatter_BuildBlock = blk
End Function

Public Function Scatter_Insert(ws As Worksheet, data As clsChartData, kind As String) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, anchor As Range, blk As Range, body As Range, n As Long, g As Long, k As Long, s As Series
    Dim names As Variant, isBubble As Boolean, en As Long, es As String, ed As String, sref As String
    If Kinds_Family(kind) <> "scatter" Then Err.Raise 5, "Scatter_Insert", "Tipo desconhecido: " & kind
    Scatter_CheckData data, kind                     ' recusas (6012/6013) antes de alocar: nada a desfazer
    names = GroupNames(data, kind)
    isBubble = (kind = "bubble")
    On Error GoTo Fail
    id = Registry_NewId()
    Set blk = Scatter_BuildBlock(ws.Parent, id, data, kind)
    Util_TestHookAfterBlock
    Set body = Block_Body(blk)
    n = data.CategoryCount: g = Scatter_GroupCount(blk)
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    For k = 1 To g
        Set s = cht.SeriesCollection.NewSeries
        s.Name = CStr(names(k))
        s.Values = Block_Ref(body.Cells(2, Scatter_YgCol(blk, k)).Resize(n, 1))
    Next
    cht.ChartType = IIf(isBubble, xlBubble, xlXYScatter)
    For k = 1 To g                                   ' depois do tipo: uma serie XY precisa existir como XY para aceitar X numerico
        Set s = cht.SeriesCollection(k)
        s.XValues = Block_Ref(body.Cells(2, Scatter_XgCol(blk, k)).Resize(n, 1))
        If isBubble Then
            ' a documentacao pede R1C1 ao GRAVAR BubbleSizes e a sonda nao registrou a forma aceita: tenta A1 e, se recusada, R1C1
            ' (registre no relatorio qual forma o Excel aceitou)
            sref = Block_Ref(body.Cells(2, Scatter_SgCol(blk, k)).Resize(n, 1))
            On Error Resume Next
            s.BubbleSizes = sref
            en = Err.Number
            Err.Clear
            On Error GoTo Fail
            If en <> 0 Then s.BubbleSizes = Application.ConvertFormula(sref, xlA1, xlR1C1)
        End If
    Next
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlNotPlotted               ' ponto NA() (outro grupo, dado invalido) sem marcador e sem rotulo
    Style_ApplyBase cht
    Util_TestHookAfterChart
    Scatter_StyleAxes cht, data
    If isBubble Then
        cht.ChartGroups(1).BubbleScale = 60
        cht.ChartGroups(1).SizeRepresents = xlSizeIsArea
        cht.ChartGroups(1).ShowNegativeBubbles = False
    End If
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight
    cht.Legend.Font.Size = 9
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "n", CStr(n)
    Registry_SetOption cht, "legend", "right"
    Registry_SetOption cht, "trend", "0"
    Registry_SetOption cht, "quad", "0"
    Scatter_ApplyColors cht, "padrao"
    Scatter_SetLabels cht, True
    Set Scatter_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' Eixos visiveis (os numeros nao estao em rotulos): linha cinza, rotulos de 9 pt na BORDA (Low: com os quadrantes a linha cruza no meio
' e os rotulos ficariam sobre os pontos), sem grade, titulos ligados aos cabecalhos de X e Y.
Private Sub Scatter_StyleAxes(cht As Chart, data As clsChartData)
    Dim ax As Axis, k As Long
    cht.HasAxis(xlCategory, xlPrimary) = True
    cht.HasAxis(xlValue, xlPrimary) = True
    For k = 1 To 2
        Set ax = cht.Axes(IIf(k = 1, xlCategory, xlValue), xlPrimary)
        With ax
            .HasMajorGridlines = False
            .HasMinorGridlines = False
            .MajorTickMark = xlTickMarkOutside
            .MinorTickMark = xlTickMarkNone
            .TickLabelPosition = xlTickLabelPositionLow
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = SC_AXIS_COLOR
            .Format.Line.DashStyle = msoLineSolid
            .TickLabels.Font.Size = 9
            .TickLabels.Font.Color = SC_TEXT_COLOR
            .HasTitle = True
            .AxisTitle.Formula = ChartData_ExternalAddress(data.SeriesNameCell(k))
            .AxisTitle.Font.Size = 9
            .AxisTitle.Font.Bold = False
            .AxisTitle.Font.Color = SC_TEXT_COLOR
        End With
    Next
End Sub

Public Function Scatter_SeriesCount(cht As Chart) As Long
    Dim s As Series
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then Scatter_SeriesCount = Scatter_SeriesCount + 1
    Next
End Function

Public Function Scatter_DataSeries(cht As Chart, k As Long) As Series
    Dim s As Series, i As Long
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            i = i + 1
            If i = k Then Set Scatter_DataSeries = s: Exit Function
        End If
    Next
End Function

' Rotulos: o nome de cada ponto, ligado a coluna Nome do bloco (ponto a ponto; nunca Labels_LinkPoints, que gravaria formato numerico nas
' celulas de texto). Dispersao: a direita do marcador, cinza; bolhas: no centro, fonte por contraste com a cor da serie.
Public Sub Scatter_SetLabels(cht As Chart, turnOn As Boolean)
    Dim body As Range, n As Long, k As Long, i As Long, s As Series, isBubble As Boolean, c As Long
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    isBubble = (Registry_ChartType(cht) = "bubble")
    For k = 1 To Scatter_SeriesCount(cht)
        Set s = Scatter_DataSeries(cht, k)
        If Not turnOn Then
            s.HasDataLabels = False
        Else
            c = IIf(isBubble, Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB)), RGB(64, 64, 64))
            For i = 1 To n
                If i <= s.Points.Count Then
                    ' isolado por ponto (como Labels_LinkPoints): so a dispersao foi medida com rotulo em ponto NA(); um ponto NA() de bolha
                    ' que recuse o rotulo nao pode desfazer o Insert inteiro
                    On Error Resume Next
                    With s.Points(i)
                        .HasDataLabel = True
                        .DataLabel.Formula = Block_Ref(body.Cells(i + 1, SC_COL_NAME))
                        .DataLabel.Position = IIf(isBubble, xlLabelPositionCenter, xlLabelPositionRight)
                        .DataLabel.Font.Size = 9
                        .DataLabel.Font.Bold = False
                        .DataLabel.Font.Color = c
                    End With
                    If Err.Number <> 0 Then Errors_Log "Scatter_SetLabels" & vbTab & s.Name & vbTab & i & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
                    On Error GoTo 0
                End If
            Next
        End If
    Next
    Registry_SetOption cht, "labels", IIf(turnOn, "1", "0")
End Sub

' Cores por serie (grupo): marcadores XY so aceitam a cor por MarkerBackground/ForegroundColor (medido na Fase 2); as bolhas tambem pelo
' preenchimento; a tendencia (se existe) e os rotulos das bolhas acompanham.
Public Sub Scatter_ApplyColors(cht As Chart, paletteName As String)
    Dim pal As Variant, k As Long, s As Series, c As Long, isBubble As Boolean, i As Long
    pal = Style_Palette(paletteName)
    isBubble = (Registry_ChartType(cht) = "bubble")
    For k = 1 To Scatter_SeriesCount(cht)
        Set s = Scatter_DataSeries(cht, k)
        c = Style_PaletteRGB(cht, pal((k - 1) Mod 6))
        s.Format.Fill.Visible = msoTrue
        s.Format.Fill.ForeColor.RGB = c
        s.MarkerBackgroundColor = c
        s.MarkerForegroundColor = c
        If isBubble Then
            s.Format.Line.Visible = msoTrue
            s.Format.Line.ForeColor.RGB = RGB(255, 255, 255)
            s.Format.Line.Weight = 1
            If s.HasDataLabels Then s.DataLabels.Font.Color = Style_ContrastFont(c)
        Else
            s.MarkerStyle = xlMarkerStyleCircle
            s.MarkerSize = 7
            s.Format.Line.Visible = msoFalse         ' sem linha ligando os pontos
        End If
        For i = 1 To s.Trendlines.Count
            s.Trendlines(i).Format.Line.ForeColor.RGB = c
        Next
    Next
End Sub

' Linha de tendencia linear (regressao nativa do Excel) por serie, tracejada, na cor da serie; as entradas de legenda das tendencias sao
' apagadas (Scatter_HideTrendEntries). Desligar apaga as tendencias.
Public Sub Scatter_SetTrend(cht As Chart, turnOn As Boolean)
    Dim k As Long, s As Series, t As Trendline, pal As Variant
    pal = Style_Palette(Registry_GetOption(cht, "palette", "padrao"))
    For k = 1 To Scatter_SeriesCount(cht)
        Set s = Scatter_DataSeries(cht, k)
        Do While s.Trendlines.Count > 0
            s.Trendlines(1).Delete
        Loop
        If turnOn Then
            Set t = s.Trendlines.Add(xlLinear)
            t.Format.Line.ForeColor.RGB = Style_PaletteRGB(cht, pal((k - 1) Mod 6))
            t.Format.Line.Weight = 1
            t.Format.Line.DashStyle = msoLineDash
        End If
    Next
    Registry_SetOption cht, "trend", IIf(turnOn, "1", "0")
    Scatter_HideTrendEntries cht
End Sub

' As entradas de legenda das tendencias vem DEPOIS das entradas das series (medido): apaga as que excedem o numero de series. Idempotente;
' religar a legenda (Deco_SetLegend) as traz de volta, e o refresh as apaga de novo. Nunca propaga erro.
Public Sub Scatter_HideTrendEntries(cht As Chart)
    Dim ns As Long
    On Error Resume Next
    If Not cht.HasLegend Then Exit Sub
    ns = Scatter_SeriesCount(cht)
    Do While cht.Legend.LegendEntries.Count > ns
        cht.Legend.LegendEntries(cht.Legend.LegendEntries.Count).Delete
        If Err.Number <> 0 Then Exit Do
    Loop
End Sub

' Quadrantes: os eixos cruzam nas medias (Axes(xlCategory).CrossesAt = media de X: linha vertical; Axes(xlValue).CrossesAt = media de Y:
' horizontal), tracejados; os rotulos continuam na borda (TickLabelPosition Low). Desligar volta ao cruzamento automatico e a linha solida.
Public Sub Scatter_SetQuad(cht As Chart, turnOn As Boolean)
    Registry_SetOption cht, "quad", IIf(turnOn, "1", "0")
    If turnOn Then
        Scatter_ApplyQuadrants cht
    Else
        cht.Axes(xlCategory, xlPrimary).Crosses = xlAxisCrossesAutomatic
        cht.Axes(xlValue, xlPrimary).Crosses = xlAxisCrossesAutomatic
        cht.Axes(xlCategory, xlPrimary).Format.Line.DashStyle = msoLineSolid
        cht.Axes(xlValue, xlPrimary).Format.Line.DashStyle = msoLineSolid
    End If
End Sub

' Reaplica o cruzamento nas medias (meta E/F). So escreve quando mudou (chamada a cada refresh). Sem o ThinQQ as medias ficam no ultimo
' valor aplicado (como o intervalo do eixo do Gantt).
Public Sub Scatter_ApplyQuadrants(cht As Chart)
    Dim blk As Range, ax As Axis, v As Variant, k As Long
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    For k = 1 To 2
        v = blk.Cells(1, IIf(k = 1, SC_META_AVGX, SC_META_AVGY)).Value2
        If IsError(v) Or IsEmpty(v) Then Exit Sub
        Set ax = cht.Axes(IIf(k = 1, xlCategory, xlValue), xlPrimary)
        If ax.Crosses <> xlAxisCrossesCustom Or Abs(CDbl(ax.CrossesAt) - CDbl(v)) > 0.000000001 * (1 + Abs(CDbl(v))) Then ax.CrossesAt = CDbl(v)
        ax.TickLabelPosition = xlTickLabelPositionLow
        ax.Format.Line.DashStyle = msoLineDash
    Next
End Sub

' O que na dispersao depende dos DADOS e muda depois de inserir (Core_SafeRefresh): as medias dos quadrantes e as entradas de legenda das
' tendencias. Nunca propaga erro.
Public Sub Scatter_RefreshAfterDataChange(cht As Chart)
    On Error Resume Next
    If Kinds_Family(Registry_ChartType(cht)) <> "scatter" Then Exit Sub
    If Registry_GetOption(cht, "quad", "0") = "1" Then Scatter_ApplyQuadrants cht
    If Err.Number <> 0 Then Errors_Log "Scatter_RefreshAfterDataChange/quad" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Scatter_HideTrendEntries cht
End Sub
```

- [ ] **Passo 4: Ligações** — `Kinds_Insert`: `Case "scatter": Set Kinds_Insert = Scatter_Insert(ws, data, kind)`. `modCoreApi`: em `ApplyToggle`, a lista de chaves passa a `Case "totals", "labels", "connectors", "ind100", "vline", "secondary", "pieout", "trend", "quad"`; `Case "labels"` ganha `Case "scatter": Scatter_SetLabels c, value`; novos `Case "trend": Scatter_SetTrend c, value` e `Case "quad": Scatter_SetQuad c, value` (a matriz já recusou os tipos sem o recurso); `PALETTE_FAMILIES` ganha `scatter,` e o despacho `Case "scatter": Scatter_ApplyColors c, value`; em `ReapplyOptions`, `Array("totals", "labels", "connectors", "secondary", "ind100", "vline", "pieout", "trend", "quad")`; em `Core_SafeRefresh`, depois do `Bfly_RefreshAfterDataChange`: `Scatter_RefreshAfterDataChange cht` (log `"Core_SafeRefresh/Scatter"`).

- [ ] **Passo 5: GREEN e suíte completa.** Exporte um PNG da dispersão com tendência e quadrantes e um das bolhas; registre o que viu (quadrantes nas médias, nomes à direita/centro), qual forma de `BubbleSizes` o Excel aceitou (A1 ou R1C1) e se algum rótulo de ponto foi recusado (log `Scatter_SetLabels`).

- [ ] **Passo 6: Commit** — `git add src/core/modules/modChartScatter.bas src/core/modules/modChartKinds.bas src/core/modules/modCoreApi.bas tests/vba/core/TestScatter.bas tests/vba/core/TestCoreApi.bas` e `git commit -m "feat: dispersao e bolhas (series por grupo, nomes por ponto, titulos de eixo, tendencia, quadrantes)"` (+ trailer).

**Verificação no Excel real:** tabela Produto/Preço/Margem/Grupo; `Core_Insert "scatter"` pela janela Imediata até a Tarefa 8; mudar o grupo de um ponto move o ponto de série; Linha de tendência e Quadrantes (médias) pelos toggles (Tarefa 8) ou por `Core_SetToggle "quad", True`.

---
### Tarefa 3: Rótulos com soma consistente (`roundsum`) e "Arredondar intervalo" — modelo: most-capable

**Files:**
- Create: `src/core/modules/modRounding.bas`
- Modify: `src/core/modules/modChartStacked.bas` (`Stacked_SetLabels` liga a `Rnd_j` com `roundsum=1`), `src/core/modules/modChartArea.bas` (`Area_SetLabels` idem), `src/core/modules/modLabels.bas` (`Labels_ApplyCellFormat` chama `Round_ApplyCellFormat` quando o bloco tem `RndD`), `src/core/modules/modCoreApi.bas` (`ApplyToggle` chave `roundsum` → `Round_SetLabels`; `ReapplyOptions` com `roundsum`; `Core_RoundRange`)
- Test: `tests/vba/core/TestRounding.bas` (novo), `tests/vba/core/TestCoreApi.bas` (`toggles` e `ChoicesSnapshot` do all-kinds ganham `"roundsum"`)

**Interfaces:**
- Consumes: `Block_FindCol`, `Block_Extend`, `Block_SourceRange`, `Block_Body`, `Block_Ref`, `Registry_*`, `Labels_LinkPoints`, `Labels_SegmentFormat`, `Style_LocalizeCode`, `Style_ContrastFont`, `Style_DataSeriesColor`, `Style_IsHelperSeries`, `Area_ValCol`, `Area_SeriesCount`, `Kinds_*`, `ChartData_ExternalAddress`.
- Produces (`modRounding`): `Round_Decimals(numfmtKey) As Long` (`int` 0, `dec1` 1, `dec2` 2, `pct0` 2, `pct1` 3, `mil` -3, `mi` -5); `Round_Formula(v, rowRange, partialRange, target, decimals) As String` (pura; texto da fórmula do maior resto para uma célula; `decimals` é um endereço ou literal); `Round_SeriesCount(cht) As Long` (séries não auxiliares); `Round_DCol(blk)`, `Round_VCol(blk, j)`, `Round_RCol(blk, j)` (0 se não existem); `Round_EnsureColumns(cht) As Range` (estende o bloco com `RndD | V_1..V_S | Rnd_1..Rnd_S` na primeira vez e devolve o bloco); `Round_ApplyCellFormat(cht)` (grava `RndD` = casas do `numfmt` e o formato dos segmentos nas células `Rnd_j`); `Round_SetLabels(cht, on)` (7002 `"'Rótulos com soma consistente' não se aplica a este tipo de gráfico."` fora da matriz); `Round_ValidateRange(rng) As String` (`""` ou `"Selecione uma única coluna ou linha só com números (pelo menos 2)."`); `Round_WriteFormulas(src, dest, decimals)`.
- Produces (`modCoreApi`): `Core_RoundRange()` (seleção = faixa; valida; em build `Test` sai antes dos `InputBox`; pede casas decimais (tipo 1, padrão 0) e a célula de destino (tipo 8); grava por `Round_WriteFormulas`; avisa `"Fórmulas de arredondamento com soma consistente gravadas em <endereço>."`).
- Opção persistida: `roundsum` (`0|1`). Colunas acrescentadas à direita do bloco (cabeçalhos `RndD`, `V1..VS`, `Rnd1..RndS`); `V_j`: `=IF(ISNUMBER(src),src,0)` (empilhadas), `=Part_j` (100%), `=Val_j` (área); alvo `SUM(V_1..V_S)` nos 100% (a soma das participações: 1, ou 0 numa categoria vazia, que assim não ganha um "1%" fantasma) e a célula `Total` nos demais; empate de restos desfeito pela ordem das colunas: a **primeira** coluna empatada recebe a unidade (ou a perde, quando a soma passa do alvo).

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestRounding.bas`:

```vb
Attribute VB_Name = "TestRounding"
Option Explicit

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

' Tabela 3 series x 2 categorias a partir de A1; devolve a faixa.
Private Function Fixture(ws As Worksheet, v As Variant) As Range
    ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "S1", "S2", "S3")
    ws.Range("A2:D2").Value = Array(2024, v(0), v(1), v(2))
    ws.Range("A3:D3").Value = Array(2025, v(3), v(4), v(5))
    Set Fixture = ws.Range("A1:D3")
End Function

Private Function Pct(n As Long, Optional dec As String = "") As String
    Pct = CStr(n) & dec & "%"
End Function

Public Sub Test_Round_Decimals_And_Formula()
    Dim wb As Workbook, ws As Worksheet, col As Variant, i As Long, f As String
    AssertEqual 0, Round_Decimals("int"): AssertEqual 1, Round_Decimals("dec1"): AssertEqual 2, Round_Decimals("dec2")
    AssertEqual 2, Round_Decimals("pct0"): AssertEqual 3, Round_Decimals("pct1"): AssertEqual -3, Round_Decimals("mil"): AssertEqual -5, Round_Decimals("mi")
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array(0.335, 0.335, 0.33): ws.Range("E1").Formula = "=SUM(A1:C1)": ws.Range("F1").Value = 2
    col = Array("A", "B", "C")
    For i = 0 To 2
        f = Round_Formula(col(i) & "1", "$A$1:$C$1", "$A$1:" & col(i) & "1", "$E$1", "$F$1")
        AssertEqual "=", Left$(f, 1): AssertTrue InStr(f, "SUMPRODUCT") > 0 And InStr(f, "COUNTIF") = 0, "so ROUND/SUMPRODUCT/POWER/IF"
        ws.Range(col(i) & "2").Formula = f
    Next
    AssertAlmost 0.33, ws.Range("A2").Value, 0.0000001: AssertAlmost 0.34, ws.Range("B2").Value, 0.0000001: AssertAlmost 0.33, ws.Range("C2").Value, 0.0000001
    ws.Range("A1:C1").Value = Array(1.5, 1.5, 1.5): ws.Range("F1").Value = 0
    AssertEqual 1, ws.Range("A2").Value: AssertEqual 2, ws.Range("B2").Value: AssertEqual 2, ws.Range("C2").Value
    ws.Range("A1:C1").Value = Array(33.3333, 33.3333, 33.3334)
    AssertEqual 33, ws.Range("A2").Value: AssertEqual 33, ws.Range("B2").Value: AssertEqual 34, ws.Range("C2").Value
    ws.Range("A1:C1").Value = Array(1234, 2345, 3456): ws.Range("F1").Value = -3
    AssertEqual 1000, ws.Range("A2").Value: AssertEqual 2000, ws.Range("B2").Value: AssertEqual 4000, ws.Range("C2").Value
    ws.Range("A1:C1").Value = Array(-1.5, 2.5, 2.5): ws.Range("F1").Value = 0
    AssertEqual -2, ws.Range("A2").Value: AssertEqual 3, ws.Range("B2").Value: AssertEqual 3, ws.Range("C2").Value
    CloseTemp wb
End Sub

' Review Focus 3: tres participacoes de um terco somam 100%.
Public Sub Test_Round_Stacked100_SharesSumTo100()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, blk As Range, body As Range, j As Long, s As Series, dec As String, sum As Double
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set rng = Fixture(ws, Array(1, 1, 1, 2, 3, 5))
    rng.Select
    Core_Insert "stacked100"
    Set cht = Core_ActiveChart()
    AssertEqual Pct(33), cht.SeriesCollection(1).Points(1).DataLabel.Text, "antes: 33%"
    Core_SetToggle "roundsum", True
    AssertTrue Core_GetToggle("roundsum")
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertTrue Round_DCol(blk) > 0 And Round_RCol(blk, 3) > 0, "colunas RndD e Rnd3"
    AssertEqual 2, body.Cells(2, Round_DCol(blk)).Value, "pct0: 2 casas da fracao"
    AssertEqual "Pos", body.Cells(1, Stacked_PosCol(blk)).Value, "Pos continua achada pelo cabecalho"
    For j = 1 To 3
        Set s = cht.SeriesCollection(j)
        AssertTrue InStr(s.Points(1).DataLabel.Formula, CALC_SHEET) > 0, "rotulo " & j & " ligado ao bloco"
    Next
    AssertEqual Pct(34), cht.SeriesCollection(1).Points(1).DataLabel.Text, "empate: a primeira coluna recebe a unidade"
    AssertEqual Pct(33), cht.SeriesCollection(2).Points(1).DataLabel.Text
    AssertEqual Pct(33), cht.SeriesCollection(3).Points(1).DataLabel.Text
    sum = 0
    For j = 1 To 3: sum = sum + CDbl(body.Cells(2, Round_RCol(blk, j)).Value): Next
    AssertAlmost 1, sum, 0.0000001, "as participacoes arredondadas somam 100%"
    Core_SetOption "numfmt", "pct1"
    dec = Application.International(xlDecimalSeparator)
    AssertEqual 3, body.Cells(2, Round_DCol(blk)).Value, "pct1: 3 casas"
    AssertEqual Pct(33, dec & "4"), cht.SeriesCollection(1).Points(1).DataLabel.Text, "pct1: a primeira coluna recebe a unidade"
    AssertEqual Pct(33, dec & "3"), cht.SeriesCollection(3).Points(1).DataLabel.Text
    AssertEqual Pct(20, dec & "0"), cht.SeriesCollection(1).Points(2).DataLabel.Text, "linha 2: 2/10 exato"
    ws.Range("B3:D3").ClearContents
    cht.Refresh
    AssertEqual "", cht.SeriesCollection(1).Points(2).DataLabel.Text, "categoria vazia: sem 1% fantasma"
    ws.Range("B3:D3").Value = Array(2, 3, 5)
    Core_SetToggle "roundsum", False
    AssertFalse Core_GetToggle("roundsum")
    AssertEqual Pct(33, dec & "3"), cht.SeriesCollection(1).Points(1).DataLabel.Text, "desligado: participacao crua de novo (33,3, nao 33,4)"
    Teardown wb
End Sub

Public Sub Test_Round_Stacked_Values_SumToRoundedTotal()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, blk As Range, body As Range, cols As Long, j As Long, s As Series, src As Range
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set rng = Fixture(ws, Array(1.5, 1.5, 1.5, 10, 20, 30))
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    cols = Registry_Block(cht).Columns.Count
    AssertEqual "5", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "total 4,5 mostra 5"
    AssertEqual "2", cht.SeriesCollection(1).Points(1).DataLabel.Text, "antes: cada 1,5 mostra 2 (2+2+2 = 6)"
    Core_SetToggle "roundsum", True
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual cols + 7, blk.Columns.Count, "RndD + 3 V + 3 Rnd"
    AssertEqual "V1", body.Cells(1, Round_VCol(blk, 1)).Value: AssertEqual "Rnd3", body.Cells(1, Round_RCol(blk, 3)).Value
    AssertEqual 1.5, body.Cells(2, Round_VCol(blk, 1)).Value, "V = valor saneado da origem"
    AssertEqual "1", cht.SeriesCollection(1).Points(1).DataLabel.Text
    AssertEqual "2", cht.SeriesCollection(2).Points(1).DataLabel.Text
    AssertEqual "2", cht.SeriesCollection(3).Points(1).DataLabel.Text, "1 + 2 + 2 = 5 = total arredondado"
    AssertEqual "5", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    AssertEqual "10", cht.SeriesCollection(1).Points(2).DataLabel.Text, "linha exata inalterada"
    ws.Range("B2").Value = "x"
    AssertEqual 0, body.Cells(2, Round_VCol(blk, 1)).Value, "texto vale 0 (sem #VALOR!)"
    AssertFalse IsError(body.Cells(2, Round_RCol(blk, 1)).Value)
    ws.Range("B2").Value = 1.5
    Set src = Block_SourceRange(cht)
    Core_SetOption "numfmt", "dec1"
    AssertEqual 1, body.Cells(2, Round_DCol(blk)).Value
    AssertEqual "1" & Application.International(xlDecimalSeparator) & "5", cht.SeriesCollection(1).Points(1).DataLabel.Text, "com 1 casa nada precisa de ajuste"
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(255, 255, 255), cht.SeriesCollection(1).DataLabels.Font.Color, "fonte por contraste continua"
    Teardown wb
End Sub

Public Sub Test_Round_Area_Toggle_And_Rebuild()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, blk As Range, lbl As Series, before As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set rng = Fixture(ws, Array(10.4, 10.4, 10.4, 1, 1, 1))
    rng.Select
    Core_Insert "area"
    Set cht = Core_ActiveChart()
    AssertEqual "31", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    Core_SetToggle "roundsum", True
    Set blk = Registry_Block(cht)
    AssertEqual 3, Area_SeriesCount(blk), "contagem por cabecalho sobrevive a extensao"
    Set lbl = Chart_FindSeries(cht, "tqq_lbl_3")
    AssertTrue InStr(lbl.Points(1).DataLabel.Formula, Block_Body(blk).Cells(2, Round_RCol(blk, 3)).Address(True, True)) > 0, "rotulo de segmento ligado a Rnd3"
    AssertEqual "11", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text, "empate: a primeira coluna recebe a unidade"
    AssertEqual "10", lbl.Points(1).DataLabel.Text, "11 + 10 + 10 = 31"
    AssertTrue Core_RebuildFromRange(cht, rng), "rebuild"
    Set cht = Core_ActiveChart()
    AssertTrue Core_GetToggle("roundsum"), "opcao mantida"
    Set blk = Registry_Block(cht)
    AssertTrue Round_RCol(blk, 3) > 0, "colunas recriadas no bloco novo"
    AssertEqual "11", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text, "religado apos o rebuild"
    Core_SetToggle "roundsum", False
    AssertTrue InStr(Chart_FindSeries(cht, "tqq_lbl_3").Points(1).DataLabel.Formula, Block_Body(blk).Cells(2, Area_ValCol(blk, 3)).Address(True, True)) > 0, "desligado: Val3"
    before = blk.Columns.Count
    Core_SetToggle "roundsum", True
    AssertEqual before, Registry_Block(cht).Columns.Count, "religar nao estende de novo"
    Teardown wb
End Sub

Public Sub Test_Round_Refusals_And_LabelsOff()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, en As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set rng = Fixture(ws, Array(1, 2, 3, 4, 5, 6))
    rng.Select
    Core_Insert "clustered"
    Set cht = Core_ActiveChart()
    Core_ClearLastNotice
    Core_SetToggle "roundsum", True
    AssertEqual "'Rótulos com soma consistente' não se aplica a este tipo de gráfico.", Core_LastNotice()
    AssertFalse Core_GetToggle("roundsum")
    AssertEqual 0, Round_DCol(Registry_Block(cht)), "nada gravado no bloco"
    On Error Resume Next
    Round_SetLabels cht, True
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "labels", False
    Core_SetToggle "roundsum", True
    AssertFalse cht.SeriesCollection(1).HasDataLabels, "soma consistente com rotulos desligados nao liga rotulos"
    Core_SetToggle "labels", True
    AssertTrue InStr(cht.SeriesCollection(1).Points(1).DataLabel.Formula, CALC_SHEET) > 0, "religar os rotulos liga a Rnd"
    Teardown wb
End Sub

Public Sub Test_Round_WriteFormulas_Range_And_Validate()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:A3").Value = Application.Transpose(Array(0.335, 0.335, 0.33))
    AssertEqual "", Round_ValidateRange(ws.Range("A1:A3"))
    AssertEqual "Selecione uma única coluna ou linha só com números (pelo menos 2).", Round_ValidateRange(ws.Range("A1"))
    AssertEqual "Selecione uma única coluna ou linha só com números (pelo menos 2).", Round_ValidateRange(ws.Range("A1:B3"))
    ws.Range("B2").Value = "x"
    AssertEqual "Selecione uma única coluna ou linha só com números (pelo menos 2).", Round_ValidateRange(ws.Range("B1:B3"))
    Round_WriteFormulas ws.Range("A1:A3"), ws.Range("C1"), 2
    AssertAlmost 0.33, ws.Range("C1").Value, 0.0000001: AssertAlmost 0.34, ws.Range("C2").Value, 0.0000001: AssertAlmost 0.33, ws.Range("C3").Value, 0.0000001
    AssertTrue InStr(ws.Range("C1").Formula, "$A$1:$A$3") > 0 And InStr(ws.Range("C1").Formula, CALC_SHEET) = 0, "so a origem"
    ws.Range("E1:G1").Value = Array(1234, 2345, 3456)
    Round_WriteFormulas ws.Range("E1:G1"), ws.Range("E3"), -3
    AssertEqual 1000, ws.Range("E3").Value: AssertEqual 2000, ws.Range("F3").Value: AssertEqual 4000, ws.Range("G3").Value
    ws.Range("A1").Value = 0.5
    AssertAlmost 0.5, ws.Range("C1").Value, 0.0000001, "vivo: a origem mudou (0,5 + 0,335 + 0,33 = 1,165 -> 1,17: 0,5 + 0,34 + 0,33)"
    AssertAlmost 1.17, ws.Range("C1").Value + ws.Range("C2").Value + ws.Range("C3").Value, 0.0000001
    CloseTemp wb
End Sub

Public Sub Test_Core_RoundRange_TestBuild_ValidatesWithoutInputBox()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1").Select
    Core_ClearLastNotice
    Core_RoundRange
    AssertEqual "Selecione uma única coluna ou linha só com números (pelo menos 2).", Core_LastNotice()
    ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))
    ws.Range("A1:A3").Select
    Core_ClearLastNotice
    Core_RoundRange                                             ' build Test: sai antes do InputBox, sem aviso
    AssertEqual "", Core_LastNotice()
    CloseTemp wb
End Sub
```

Em `TestCoreApi.bas`: `toggles` do all-kinds ganha `"roundsum"`; `ChoicesSnapshot` ganha `"roundsum"`.

- [ ] **Passo 2: RED** — `-Filter "TestRounding|AllKinds"` → saída 3.

- [ ] **Passo 3: Implementar `modRounding.bas`**:

```vb
Attribute VB_Name = "modRounding"
Option Explicit
' Rotulos com soma consistente (opcao roundsum): arredondamento por MAIOR RESTO em formulas nativas do bloco (sem UDF: uma funcao de
' suplemento vira vinculo externo ao caminho do .xlam quando a pasta e aberta sem o ThinQQ, e o caminho muda a cada versao do core).
' Para cada categoria (linha), os S valores da linha sao arredondados a D casas e a diferenca para o alvo arredondado (nos 100%, a soma
' das participacoes: 1, ou 0 numa linha vazia; a coluna Total nos demais) e distribuida, uma unidade por vez, aos valores de maior resto (ou retirada dos de menor resto), com empate
' desfeito pela ordem das colunas (a PRIMEIRA coluna empatada recebe a unidade, ou a perde). So ROUND, SUMPRODUCT, POWER e IF (sem COUNTIF, sem entrada matricial): verificado no Excel 2019.
' As colunas RndD | V_1..V_S | Rnd_1..Rnd_S sao acrescentadas A DIREITA do bloco na primeira vez (Block_Extend); V_j = valor saneado
' (empilhadas: IF(ISNUMBER(src),src,0); 100%: =Part_j; area: =Val_j). Os rotulos de segmento passam a apontar para Rnd_j; o rotulo do
' total continua ligado ao Total, que o formato mostra como ROUND(Total, D) = soma dos Rnd_j.
Private Const RND_HDR_D As String = "RndD"
Private Const RND_ERR_RANGE As String = "Selecione uma única coluna ou linha só com números (pelo menos 2)."

Public Function Round_Decimals(numfmtKey As String) As Long
    Select Case numfmtKey
        Case "dec1": Round_Decimals = 1
        Case "dec2": Round_Decimals = 2
        Case "pct0": Round_Decimals = 2
        Case "pct1": Round_Decimals = 3
        Case "mil": Round_Decimals = -3
        Case "mi": Round_Decimals = -5
        Case Else: Round_Decimals = 0
    End Select
End Function

' Formula (texto em ingles, com "=") do maior resto para UMA celula: v = endereco do valor, rowRange = os S valores da linha,
' partialRange = do primeiro valor ate v (empate pela ordem), target = alvo (endereco ou expressao), decimals = endereco ou literal.
Public Function Round_Formula(v As String, rowRange As String, partialRange As String, target As String, decimals As String) As String
    Dim u As String, k As String, remExpr As String, rd As String, ra As String
    u = "POWER(10,-" & decimals & ")"
    k = "ROUND((ROUND(" & target & "," & decimals & ")-SUMPRODUCT(ROUND(" & rowRange & "," & decimals & ")))/" & u & ",0)"
    remExpr = "(" & v & "-ROUND(" & v & "," & decimals & "))/" & u
    rd = "(SUMPRODUCT(--((" & rowRange & "-ROUND(" & rowRange & "," & decimals & "))/" & u & ">" & remExpr & "))+SUMPRODUCT(--((" & partialRange & "-ROUND(" & partialRange & "," & decimals & "))/" & u & "=" & remExpr & ")))"
    ra = "(SUMPRODUCT(--((" & rowRange & "-ROUND(" & rowRange & "," & decimals & "))/" & u & "<" & remExpr & "))+SUMPRODUCT(--((" & partialRange & "-ROUND(" & partialRange & "," & decimals & "))/" & u & "=" & remExpr & ")))"
    Round_Formula = "=ROUND(" & v & "," & decimals & ")+" & u & "*(IF(" & k & ">0,--(" & rd & "<=" & k & "),0)-IF(" & k & "<0,--(" & ra & "<=-" & k & "),0))"
End Function

Public Function Round_SeriesCount(cht As Chart) As Long
    Dim s As Series
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then Round_SeriesCount = Round_SeriesCount + 1
    Next
End Function

Public Function Round_DCol(blk As Range) As Long
    Round_DCol = Block_FindCol(blk, RND_HDR_D)
End Function

Public Function Round_VCol(blk As Range, j As Long) As Long
    Round_VCol = Block_FindCol(blk, "V" & j)
End Function

Public Function Round_RCol(blk As Range, j As Long) As Long
    Round_RCol = Block_FindCol(blk, "Rnd" & j)
End Function

' Cria as colunas (uma vez) e grava as formulas. Devolve o bloco (ja estendido).
Public Function Round_EnsureColumns(cht As Chart) As Range
    Dim blk As Range, body As Range, wb As Workbook, id As String, kind As String, fam As String, is100 As Boolean, S As Long, n As Long
    Dim c0 As Long, i As Long, j As Long, r As Long, src As Range, vAddr As String, rowRng As String, partRng As String, target As String, dAddr As String, srcAddr As String
    Set blk = Registry_Block(cht)
    If Round_DCol(blk) > 0 Then Set Round_EnsureColumns = blk: Exit Function
    Set wb = Registry_Workbook(cht): id = Registry_IdFromChart(cht)
    kind = Registry_ChartType(cht): fam = Kinds_Family(kind): is100 = Kinds_Is100(kind)
    S = Round_SeriesCount(cht)
    If S < 1 Then Err.Raise vbObjectError + 7002, "Round_EnsureColumns", "O gráfico não tem séries de dados."
    If fam = "stacked" And Not is100 Then
        Set src = Block_SourceRange(cht)
        If src Is Nothing Then Err.Raise vbObjectError + 7002, "Round_EnsureColumns", "Não foi possível ler a tabela de origem do gráfico. Use 'Alterar dados' para reconstruí-lo."
    End If
    c0 = blk.Columns.Count
    Set blk = Block_Extend(wb, id, 2 * S + 1)
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    body.Cells(1, c0 + 1).Value = RND_HDR_D
    For j = 1 To S
        body.Cells(1, c0 + 1 + j).Value = "V" & j
        body.Cells(1, c0 + 1 + S + j).Value = "Rnd" & j
    Next
    For i = 1 To n
        r = i + 1
        body.Cells(r, c0 + 1).Value = Round_Decimals(Registry_GetOption(cht, "numfmt", "int"))
        For j = 1 To S
            If fam = "area" Then
                body.Cells(r, c0 + 1 + j).Formula = "=" & body.Cells(r, Area_ValCol(blk, j)).Address(False, False)
            ElseIf is100 Then
                body.Cells(r, c0 + 1 + j).Formula = "=" & body.Cells(r, 3 + j).Address(False, False)
            Else
                srcAddr = Mid$(ChartData_ExternalAddress(src.Cells(r, 1 + j)), 2)
                body.Cells(r, c0 + 1 + j).Formula = "=IF(ISNUMBER(" & srcAddr & ")," & srcAddr & ",0)"
            End If
        Next
        rowRng = body.Cells(r, c0 + 2).Resize(1, S).Address(False, False)
        dAddr = body.Cells(r, c0 + 1).Address(False, False)
        If is100 Then target = "SUM(" & rowRng & ")" Else target = body.Cells(r, 2).Address(False, False)     ' 100%: 1, ou 0 numa linha vazia (sem "1%" fantasma)
        For j = 1 To S
            vAddr = body.Cells(r, c0 + 1 + j).Address(False, False)
            partRng = body.Cells(r, c0 + 2).Resize(1, j).Address(False, False)
            body.Cells(r, c0 + 1 + S + j).Formula = Round_Formula(vAddr, rowRng, partRng, target, dAddr)
        Next
    Next
    Round_ApplyCellFormat cht
    Set Round_EnsureColumns = blk
End Function

' RndD = casas do formato atual; celulas Rnd_j com o formato dos segmentos (o rotulo ligado mostra o texto da celula). Sem as colunas nao faz nada.
Public Sub Round_ApplyCellFormat(cht As Chart)
    Dim blk As Range, body As Range, n As Long, dcol As Long, j As Long, rc As Long
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    dcol = Round_DCol(blk)
    If dcol = 0 Then Exit Sub
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    body.Cells(2, dcol).Resize(n, 1).Value = Round_Decimals(Registry_GetOption(cht, "numfmt", "int"))
    For j = 1 To Round_SeriesCount(cht)
        rc = Round_RCol(blk, j)
        If rc > 0 Then body.Cells(2, rc).Resize(n, 1).NumberFormatLocal = Style_LocalizeCode(Labels_SegmentFormat(cht))
    Next
End Sub

' Liga/desliga a opcao e religa os rotulos de segmento pela familia (Stacked_SetLabels/Area_SetLabels leem a opcao). Recusa (7002) fora da
' matriz. Desligar nao apaga as colunas (ficam calculadas; religar nao estende de novo).
Public Sub Round_SetLabels(cht As Chart, turnOn As Boolean)
    Dim kind As String, fam As String, s As Series
    kind = Registry_ChartType(cht): fam = Kinds_Family(kind)
    If Not Kinds_Supports(kind, "roundsum") Then Err.Raise vbObjectError + 7002, "Round_SetLabels", "'" & Kinds_FeatureLabel("roundsum") & "' não se aplica a este tipo de gráfico."
    If turnOn Then Round_EnsureColumns cht
    Registry_SetOption cht, "roundsum", IIf(turnOn, "1", "0")
    If Registry_GetOption(cht, "labels", "1") <> "1" Then Exit Sub
    If fam = "area" Then
        Area_SetLabels cht, True
    Else
        ' rotulos nativos nao se desfazem de uma ligacao por ponto ao receber propriedades no nivel da serie: recomeca do zero
        For Each s In cht.SeriesCollection
            If Not Style_IsHelperSeries(s) Then s.HasDataLabels = False
        Next
        Stacked_SetLabels cht, True
    End If
End Sub

' "" se a faixa serve ao Arredondar intervalo (uma area, uma coluna OU uma linha, >= 2 celulas, todas numericas); senao a mensagem.
Public Function Round_ValidateRange(rng As Range) As String
    Dim c As Range
    Round_ValidateRange = RND_ERR_RANGE
    If rng Is Nothing Then Exit Function
    If rng.Areas.Count <> 1 Then Exit Function
    If rng.Rows.Count > 1 And rng.Columns.Count > 1 Then Exit Function
    If rng.Cells.Count < 2 Then Exit Function
    For Each c In rng.Cells
        If IsError(c.Value) Or IsEmpty(c.Value) Or VarType(c.Value) = vbString Or VarType(c.Value) = vbBoolean Then Exit Function
        If Not IsNumeric(c.Value) Then Exit Function
    Next
    Round_ValidateRange = ""
End Function

' Grava em dest (celula inicial; a saida tem a forma da origem) uma formula por celula que referencia SO a faixa de origem: portatil.
Public Sub Round_WriteFormulas(src As Range, dest As Range, decimals As Long)
    Dim n As Long, i As Long, vertical As Boolean, rowRng As String, partRng As String, out As Range, fmt As String
    n = src.Cells.Count
    vertical = (src.Columns.Count = 1)
    rowRng = Mid$(ChartData_ExternalAddress(src), 2)
    For i = 1 To n
        If vertical Then
            partRng = Mid$(ChartData_ExternalAddress(src.Cells(1, 1).Resize(i, 1)), 2)
            Set out = dest.Cells(i, 1)
        Else
            partRng = Mid$(ChartData_ExternalAddress(src.Cells(1, 1).Resize(1, i)), 2)
            Set out = dest.Cells(1, i)
        End If
        out.Formula = Round_Formula(Mid$(ChartData_ExternalAddress(src.Cells(i)), 2), rowRng, partRng, "SUM(" & rowRng & ")", "(" & decimals & ")")
    Next
    If decimals > 0 Then fmt = "#,##0." & String$(decimals, "0") Else fmt = "#,##0"
    If vertical Then dest.Cells(1, 1).Resize(n, 1).NumberFormatLocal = Style_LocalizeCode(fmt) Else dest.Cells(1, 1).Resize(1, n).NumberFormatLocal = Style_LocalizeCode(fmt)
End Sub
```

- [ ] **Passo 4: Famílias e API** — `modChartStacked.Stacked_SetLabels`: declare `blk As Range, rc As Long, rs As Boolean`; substitua `If is100 Then Set body = Block_Body(Registry_Block(cht)): n = body.Rows.Count - 1` por `Set blk = Registry_Block(cht): Set body = Block_Body(blk): n = body.Rows.Count - 1: rs = (Registry_GetOption(cht, "roundsum", "0") = "1")`; e o trecho dos 100% por:

```vb
                ' soma consistente: o rotulo liga a Rnd_j (valor ou participacao arredondados por maior resto); 100%: a participacao Part_j
                rc = 0
                If rs Then rc = Round_RCol(blk, j)
                If rc > 0 Then
                    Labels_LinkPoints s, body.Cells(2, rc).Resize(n, 1), fmt, xlLabelPositionCenter, Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
                ElseIf is100 Then
                    Labels_LinkPoints s, body.Cells(2, 3 + j).Resize(n, 1), fmt, xlLabelPositionCenter, Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
                End If
```
`modChartArea.Area_SetLabels`: declare `rc As Long`; dentro do laço, antes do `Labels_LinkPoints`: `rc = 0: If Registry_GetOption(cht, "roundsum", "0") = "1" Then rc = Round_RCol(blk, j)` e o `Labels_LinkPoints` passa a usar `body.Cells(2, IIf(rc > 0, rc, Area_ValCol(blk, j))).Resize(n, 1)`. `modLabels.Labels_ApplyCellFormat`: antes do `If Registry_GetOption(cht, "vline"...` acrescente `If Round_DCol(blk) > 0 Then Round_ApplyCellFormat cht`. `modCoreApi`: `ApplyToggle` lista de chaves ganha `"roundsum"` e `Case "roundsum": Round_SetLabels c, value`; `ReapplyOptions` array ganha `"roundsum"` (depois de `"quad"`); novo:

```vb
' Arredondar intervalo: grava, a partir da celula escolhida, formulas de maior resto que referenciam so a faixa selecionada (uma coluna ou
' uma linha de numeros). Portatil: nada depende do ThinQQ. Em build Test sai antes dos InputBox (modais travariam o Excel invisivel).
Public Sub Core_RoundRange()
    On Error GoTo Fail
    Dim rng As Range, dest As Range, dec As Variant, msg As String, n As Long
    If TypeName(Selection) <> "Range" Then
        Core_Notify "Selecione uma única coluna ou linha só com números (pelo menos 2).", vbExclamation
        Exit Sub
    End If
    Set rng = Selection
    msg = Round_ValidateRange(rng)
    If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Sub
    If BUILD_CONFIG = "Test" Then Exit Sub
    dec = Application.InputBox("Casas decimais do resultado (0 = inteiros; 1 ou 2 = decimais; -3 = milhares):", APP_NAME, 0, Type:=1)
    If VarType(dec) = vbBoolean Then Exit Sub
    On Error Resume Next
    Set dest = Application.InputBox("Célula onde o resultado começa (cada célula recebe uma fórmula que referencia só os valores de origem):", APP_NAME, Type:=8)
    On Error GoTo Fail
    If dest Is Nothing Then Exit Sub
    Round_WriteFormulas rng, dest.Cells(1, 1), CLng(dec)
    n = rng.Cells.Count
    Core_Notify "Fórmulas de arredondamento com soma consistente gravadas em " & IIf(rng.Columns.Count = 1, dest.Cells(1, 1).Resize(n, 1).Address(False, False), dest.Cells(1, 1).Resize(1, n).Address(False, False)) & ".", vbInformation
    Exit Sub
Fail:
    Errors_Handle "Core_RoundRange"
End Sub
```

- [ ] **Passo 5: GREEN e suíte completa** (a extensão só ocorre ao ligar o recurso: os testes de contagem de colunas da Fase 2 continuam verdes).

- [ ] **Passo 6: Commit** — `git add src/core/modules/modRounding.bas src/core/modules/modChartStacked.bas src/core/modules/modChartArea.bas src/core/modules/modLabels.bas src/core/modules/modCoreApi.bas tests/vba/core/TestRounding.bas tests/vba/core/TestCoreApi.bas` e `git commit -m "feat: rotulos com soma consistente (maior resto em formulas do bloco) e Arredondar intervalo"` (+ trailer).

**Verificação no Excel real:** Colunas 100% com 1/1/1: 33% 33% 33%; `Core_SetToggle "roundsum", True` (botão na Tarefa 8): 34% 33% 33% (empate: a primeira coluna recebe a unidade); Formato 12,3%: 33,4/33,3/33,3; uma linha vazia sem rótulos; abrir a pasta sem o ThinQQ mantém os rótulos.

---
### Tarefa 4: Corte de eixo (`break`) — modelo: most-capable

**Files:**
- Create: `src/core/modules/modAxisBreak.bas`
- Modify: `src/core/modules/modChartStacked.bas` (`Stacked_SetTotals` e `Stacked_EnsureRoomAfterChange` usam a coluna `LblPos` quando existe), `src/core/modules/modAnnotations.bas` (`Ann_Refresh` trata `kind=break`; `Ann_LayoutBreak`; `Ann_RemoveAll` poupa as marcas; `Ann_GetCategoryData` limita o topo ao corte), `src/core/modules/modDecorations.bas` (`Deco_SameScale` ignora gráfico com `vbreak`), `src/core/modules/modChartUtil.bas` (`Chart_EnsureRoomAfterChange` chama `Break_EnsureRoom`), `src/core/modules/modCoreApi.bas` (`Core_SetAxisBreak`, `Core_SetAxisBreakAt`, `Core_RemoveAxisBreak`, `Core_SafeRefresh` chama `Break_RefreshAfterDataChange`, `ReapplyOptions` chama `Break_TryReapply`)
- Test: `tests/vba/core/TestAxisBreak.bas` (novo), `tests/vba/core/TestCoreApi.bas` (`ChoicesSnapshot` ganha `"vbreak"`)

**Interfaces:**
- Consumes: `Block_FindCol`, `Block_Extend`, `Block_Body`, `Block_Ref`, `Registry_*`, `Chart_FindSeries`, `Geom_*`, `Deco_ScaleText`, `Deco_ShrinkPlotWidth`, `Labels_MaxTotalWidth`, `Stacked_SetLabels`, `Style_IsHelperSeries`, `Util_ToDouble`, `Kinds_*`.
- Produces (`modAxisBreak`): `Break_Limit(cht) As Double` (0 = sem corte); `Break_EnsureColumn(cht) As Long` (coluna `LblPos = MIN(TotalPos, $BreakAt)`, criada uma vez à direita do bloco); `Break_Set(cht, limit)` (`limit <= 0` remove; 7002 fora da matriz; 7003 `"Corte de eixo: informe um valor maior que zero e menor que o maior valor do gráfico."`); `Break_Remove(cht)`; `Break_ApplyAxis(cht)`; `Break_EnsureRoom(cht)`; `Break_SyncLabels(cht)`; `Break_SyncMarks(cht)`; `Break_CategoryTops(cht) As Double()` (topo positivo por categoria: soma nas empilhadas, maior série nas agrupadas); `Break_Suggest(cht) As Double` (1,2 × segundo maior topo, ou metade do maior); `Break_RefreshAfterDataChange(cht)`; `Break_TryReapply(cht)` (reaplica `vbreak` após "Alterar dados"; 7003 limpa a opção e registra).
- Produces (`modAnnotations`): `Ann_LayoutBreak(cht, shp)` (privada; formas `tqq_ann_break_<i>` (cinza), `_w` (branca, 3 pt, atrás) e `_b` (cinza, 4 pt abaixo/à direita), `AlternativeText = "kind=break;cat=<i>;part=a|w|b"`).
- Produces (`modCoreApi`): `Core_SetAxisBreak()` (InputBox com sugestão; em `Test` sai antes), `Core_SetAxisBreakAt(limit As Double)`, `Core_RemoveAxisBreak()`.
- Opção persistida: `vbreak` (texto neutro de `Deco_ScaleText`; vazio = sem corte). Limite do corte (`BreakAt`): na linha de meta da própria coluna `LblPos` (dentro do bloco: a limpeza a apaga junto e `Block_Extend` a vê como parte do bloco, então ligar a soma consistente depois do corte funciona; uma célula fixa como a `J` ficaria fora do bloco de 5 ou 6 colunas). Mesma escala e corte são excludentes: `Break_Set` limpa `vmin/vmax`; `Deco_SameScale` conta o gráfico com `vbreak` como ignorado.

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestAxisBreak.bas`:

```vb
Attribute VB_Name = "TestAxisBreak"
Option Explicit

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Cat", "S1", "S2")
    ws.Range("A2:C2").Value = Array("a", 5, 5): ws.Range("A3:C3").Value = Array("b", 10, 10): ws.Range("A4:C4").Value = Array("c", 200, 30)
    Set Fixture = ws.Range("A1:C4")
End Function

Private Function YOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, xlPrimary)
    YOf = cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight * (ax.MaximumScale - v) / (ax.MaximumScale - ax.MinimumScale)
End Function

Private Function HasShape(cht As Chart, nm As String) As Boolean
    Dim shp As Shape
    For Each shp In cht.Shapes
        If shp.Name = nm Then HasShape = True: Exit Function
    Next
End Function

Public Sub Test_Break_Set_ClipsAxis_TotalsAtLimit_LivePos()
    Dim wb As Workbook, rng As Range, cht As Chart, blk As Range, body As Range, pc As Long, tot As Series, ax As Axis
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Set ax = cht.Axes(xlValue, xlPrimary)
    AssertTrue ax.MaximumScale >= 230, "antes: escala automatica cobre 230"
    Break_Set cht, 50
    AssertEqual "50", Registry_GetOption(cht, "vbreak")
    AssertFalse ax.MaximumScaleIsAuto: AssertEqual 50, ax.MaximumScale
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    pc = Block_FindCol(blk, "LblPos")
    AssertTrue pc > 0, "coluna LblPos"
    AssertEqual 50, CDbl(blk.Cells(1, pc).Value), "meta da coluna LblPos = BreakAt (dentro do bloco)"
    AssertEqual 10, body.Cells(2, pc).Value: AssertEqual 50, body.Cells(4, pc).Value, "LblPos = MIN(TotalPos, BreakAt)"
    AssertEqual "Pos", body.Cells(1, Stacked_PosCol(blk)).Value
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertTrue InStr(tot.Formula, body.Cells(2, pc).Resize(3, 1).Address(True, True)) > 0, "tqq_total aponta para LblPos"
    AssertEqual "230", tot.Points(3).DataLabel.Text, "o rotulo mostra o total real"
    ForceLayout cht
    AssertAlmost cht.PlotArea.InsideTop, tot.Points(3).Top, 2, "ponto do total da coluna cortada no topo do plot"
    AssertTrue tot.Points(3).DataLabel.Top >= 0 And tot.Points(3).DataLabel.Top + tot.Points(3).DataLabel.Height <= cht.PlotArea.InsideTop + 2, "rotulo acima do plot, dentro do grafico"
    AssertTrue cht.PlotArea.InsideTop >= 23.5, "folga de cima para o rotulo"
    Core_SafeRefresh cht
    AssertEqual 50, ax.MaximumScale, "refresh mantem o corte"
    Teardown wb
End Sub

' Review Focus 4: segmento atravessado, segmento acima do corte, marca de corte, e a volta quando os dados descem.
Public Sub Test_Break_SegmentCrossingLimit_LabelsAndMark()
    Dim wb As Workbook, rng As Range, cht As Chart, ws As Worksheet, s1 As Series, s2 As Series, m As Shape, p As Point
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetAxisBreakAt 50
    Set s1 = cht.SeriesCollection(1): Set s2 = cht.SeriesCollection(2)
    AssertEqual xlLabelPositionInsideBase, s1.Points(3).DataLabel.Position, "segmento atravessado: rotulo na base"
    AssertFalse s2.Points(3).HasDataLabel, "segmento inteiro acima do corte: sem rotulo"
    AssertEqual xlLabelPositionCenter, s1.Points(1).DataLabel.Position, "segmento abaixo do corte: centro"
    AssertTrue HasShape(cht, "tqq_ann_break_3") And HasShape(cht, "tqq_ann_break_3_w") And HasShape(cht, "tqq_ann_break_3_b"), "marca na coluna 3"
    AssertFalse HasShape(cht, "tqq_ann_break_1") Or HasShape(cht, "tqq_ann_break_2"), "sem marca nas colunas inteiras"
    ForceLayout cht
    Set m = cht.Shapes("tqq_ann_break_3"): Set p = s1.Points(3)
    AssertAlmost YOf(cht, 50) + 8, m.Top + m.Height / 2, 2.5, "marca 8 pt abaixo do topo do plot"
    AssertTrue m.Left <= p.Left + 0.5 And m.Left + m.Width >= p.Left + p.Width - 0.5, "marca atravessa a coluna"
    AssertEqual "kind=break;cat=3;part=a", m.AlternativeText
    ws.Range("B4").Value = 20: ws.Range("C4").Value = 20
    Core_SafeRefresh cht
    AssertFalse HasShape(cht, "tqq_ann_break_3"), "total 40 < 50: marca removida"
    AssertTrue s2.Points(3).HasDataLabel, "rotulo de volta"
    AssertEqual xlLabelPositionCenter, s1.Points(3).DataLabel.Position
    AssertEqual "40", Chart_FindSeries(cht, "tqq_total").Points(3).DataLabel.Text
    AssertEqual 50, cht.Axes(xlValue).MaximumScale, "o corte continua gravado"
    Teardown wb
End Sub

Public Sub Test_Break_Remove_RestoresAuto()
    Dim wb As Workbook, rng As Range, cht As Chart, blk As Range, body As Range, pc As Long
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetAxisBreakAt 50
    Core_RemoveAxisBreak
    AssertEqual "", Registry_GetOption(cht, "vbreak", "")
    AssertTrue cht.Axes(xlValue).MaximumScaleIsAuto
    AssertFalse HasShape(cht, "tqq_ann_break_3")
    Set blk = Registry_Block(cht): Set body = Block_Body(blk): pc = Block_FindCol(blk, "LblPos")
    AssertEqual "", CStr(blk.Cells(1, pc).Value), "limite apagado da meta da coluna LblPos"
    AssertEqual 230, body.Cells(4, pc).Value, "MIN com a meta vazia = TotalPos"
    AssertEqual xlLabelPositionCenter, cht.SeriesCollection(1).Points(3).DataLabel.Position
    AssertTrue cht.SeriesCollection(2).Points(3).HasDataLabel
    Teardown wb
End Sub

Public Sub Test_Break_Clustered_And_Bars()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, s As Series, m As Shape, tot As Series, blk As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "clustered"
    Set cht = Core_ActiveChart()
    Break_Set cht, 50
    Set s = cht.SeriesCollection(1)
    AssertEqual xlLabelPositionInsideBase, s.Points(3).DataLabel.Position, "agrupadas: ponto cortado com rotulo na base"
    AssertEqual xlLabelPositionOutsideEnd, s.Points(1).DataLabel.Position
    AssertTrue HasShape(cht, "tqq_ann_break_3")
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    rng.Select
    Core_Insert "barstacked"
    Set cht = Core_ActiveChart()
    Break_Set cht, 50
    AssertEqual 50, cht.Axes(xlValue, xlPrimary).MaximumScale
    Set blk = Registry_Block(cht)
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertTrue InStr(tot.Formula, Block_Body(blk).Cells(2, Block_FindCol(blk, "LblPos")).Resize(3, 1).Address(True, True)) > 0, "barras: X do total = LblPos"
    ForceLayout cht
    Set m = cht.Shapes("tqq_ann_break_3")
    AssertTrue m.Height > m.Width, "barras: marca atravessa a barra na vertical"
    AssertTrue cht.ChartArea.Width - (cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth) >= 20, "folga a direita para o rotulo do total"
    Teardown wb
End Sub

Public Sub Test_Break_Refusals_SameScale_Rebuild_RemoveAll()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, c2 As Chart, coll As New Collection, skipped As Long, en As Long
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked100"
    Core_ClearLastNotice
    Core_SetAxisBreakAt 0.5
    AssertEqual "'Corte de eixo' não se aplica a este tipo de gráfico.", Core_LastNotice()
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_ClearLastNotice
    Core_SetAxisBreakAt 500
    AssertEqual "Corte de eixo: informe um valor maior que zero e menor que o maior valor do gráfico.", Core_LastNotice()
    AssertEqual "", Registry_GetOption(cht, "vbreak", "")
    On Error Resume Next
    Break_Set cht, -1
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertEqual 0, en, "limite <= 0 e remover (sem erro)"
    ws.Range("A10:C12").Value = Array("x", 1, 1): ws.Range("A11:C11").Value = Array("y", 2, 2): ws.Range("A12:C12").Value = Array("z", 3, 3)
    ws.Range("A10:C12").Select
    Core_Insert "stacked"
    Set c2 = Core_ActiveChart()
    coll.Add cht: coll.Add c2
    AssertEqual 2, Deco_SameScale(coll, skipped), "mesma escala antes do corte"
    AssertTrue Len(Registry_GetOption(cht, "vmax", "")) > 0
    Break_Set cht, 50
    AssertEqual "", Registry_GetOption(cht, "vmax", ""), "o corte limpa a escala fixa"
    Set coll = New Collection: coll.Add cht: coll.Add c2
    AssertEqual 0, Deco_SameScale(coll, skipped): AssertEqual 1, skipped, "grafico com corte fica fora da mesma escala"
    AssertEqual 50, cht.Axes(xlValue).MaximumScale, "o corte nao foi desfeito"
    Ann_AddCagr cht, 1, 3
    Ann_RemoveAll cht
    AssertTrue HasShape(cht, "tqq_ann_break_3"), "Remover setas poupa a marca de corte"
    AssertFalse HasShape(cht, "tqq_ann_cagr_1")
    AssertTrue Core_RebuildFromRange(cht, rng), "rebuild"
    Set cht = Core_ActiveChart()
    AssertEqual "50", Registry_GetOption(cht, "vbreak"), "corte mantido no rebuild"
    AssertEqual 50, cht.Axes(xlValue).MaximumScale
    AssertTrue HasShape(cht, "tqq_ann_break_3")
    AssertTrue Block_FindCol(Registry_Block(cht), "LblPos") > 0
    Teardown wb
End Sub

Public Sub Test_Break_ThenRoundSum_CoexistInBlock()
    Dim wb As Workbook, rng As Range, cht As Chart, blk As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetAxisBreakAt 50
    Core_ClearLastNotice
    Core_SetToggle "roundsum", True
    AssertTrue Core_GetToggle("roundsum"), "soma consistente depois do corte (" & Core_LastNotice() & ")"
    Set blk = Registry_Block(cht)
    AssertTrue Round_DCol(blk) > 0 And Block_FindCol(blk, "LblPos") > 0
    AssertEqual 50, CDbl(blk.Cells(1, Block_FindCol(blk, "LblPos")).Value), "limite dentro do bloco"
    AssertEqual 50, cht.Axes(xlValue).MaximumScale
    Teardown wb
End Sub

Public Sub Test_Core_SetAxisBreak_TestBuild_NoInputBox()
    Dim wb As Workbook, rng As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked100"
    Core_ClearLastNotice
    Core_SetAxisBreak                                           ' a recusa pela matriz vem ANTES da saida do build Test
    AssertEqual "'Corte de eixo' não se aplica a este tipo de gráfico.", Core_LastNotice()
    rng.Worksheet.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    rng.Select
    Core_Insert "stacked"
    Core_ClearLastNotice
    Core_SetAxisBreak                                           ' build Test: sai antes do InputBox
    AssertEqual "", Core_LastNotice()
    AssertEqual "", Registry_GetOption(Core_ActiveChart(), "vbreak", ""), "nada aplicado sem o InputBox"
    AssertTrue Break_Suggest(Core_ActiveChart()) > 20 And Break_Suggest(Core_ActiveChart()) < 230, "sugestao entre o 2o e o maior topo"
    rng.Worksheet.Range("A1").Select
    Core_SetAxisBreak
    AssertEqual "", Core_LastNotice(), "sem grafico: nada"
    Teardown wb
End Sub
```

Em `TestCoreApi.bas`, `ChoicesSnapshot` ganha `"vbreak"`.

- [ ] **Passo 2: RED** — `-Filter "TestAxisBreak"` → saída 3.

- [ ] **Passo 3: Implementar `modAxisBreak.bas`**:

```vb
Attribute VB_Name = "modAxisBreak"
Option Explicit
' Corte de eixo (opcao vbreak; design da Fase 3, secao 4.6): o MAXIMO FIXO do eixo de valores corta as colunas/barras na borda da area de
' plotagem (medido: o Excel desenha ate a borda); o rotulo do total passa a usar a coluna LblPos = MIN(TotalPos, BreakAt) do bloco (vivo),
' os rotulos dos segmentos acima do corte sao ocultados e o do segmento atravessado vai para a base (Break_SyncLabels), e cada coluna
' cortada recebe uma marca de corte (3 linhas: branca por baixo e duas cinzas; formas tqq_ann_break_<i>, posicionadas por Ann_Refresh).
' Sem o ThinQQ: eixo fixo e LblPos sao vivos; rotulos de segmento e marcas ficam como na ultima sincronizacao. Mesma escala e corte sao
' excludentes (Break_Set limpa vmin/vmax; Deco_SameScale ignora graficos com corte).
Private Const BK_HDR As String = "LblPos"
Private Const BK_PREFIX As String = "tqq_ann_break_"
Private Const BK_TOP_ROOM As Double = 24        ' folga (pt) acima da area interna para o rotulo do total da coluna cortada
Private Const BK_RIGHT_MARGIN As Double = 10    ' barras: folga = largura do maior rotulo de total + esta margem
Private Const BK_ERR_LIMIT As String = "Corte de eixo: informe um valor maior que zero e menor que o maior valor do gráfico."
Private mBusy As Boolean

Public Function Break_Limit(cht As Chart) As Double
    Break_Limit = Val(Registry_GetOption(cht, "vbreak", ""))
    If Break_Limit < 0 Then Break_Limit = 0
End Function

' Celula do limite (BreakAt): a linha de meta da coluna LblPos (dentro do bloco: a limpeza a apaga junto e Block_Extend a ve como parte
' do bloco; vazio = sem corte, e MIN ignora a celula vazia). Nothing se a coluna ainda nao existe.
Private Function LimitCell(cht As Chart) As Range
    Dim blk As Range, c As Long
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    c = Block_FindCol(blk, BK_HDR)
    If c > 0 Then Set LimitCell = blk.Cells(1, c)
End Function

' Coluna LblPos (criada uma vez, a direita do bloco): MIN(TotalPos, $BreakAt), com BreakAt na linha de meta da propria coluna (LimitCell).
' Devolve a coluna no corpo.
Public Function Break_EnsureColumn(cht As Chart) As Long
    Dim blk As Range, body As Range, c As Long, n As Long, i As Long, lim As String
    Set blk = Registry_Block(cht)
    c = Block_FindCol(blk, BK_HDR)
    If c > 0 Then Break_EnsureColumn = c: Exit Function
    Set blk = Block_Extend(Registry_Workbook(cht), Registry_IdFromChart(cht), 1)
    Set body = Block_Body(blk)
    c = blk.Columns.Count
    n = body.Rows.Count - 1
    lim = blk.Cells(1, c).Address(True, True)           ' meta da propria coluna: dentro do bloco
    body.Cells(1, c).Value = BK_HDR
    For i = 1 To n
        body.Cells(i + 1, c).Formula = "=MIN(" & body.Cells(i + 1, 3).Address(False, False) & "," & lim & ")"
    Next
    Break_EnsureColumn = c
End Function

' Topo positivo de cada categoria: soma dos valores > 0 nas empilhadas, maior valor entre as series nas agrupadas (series auxiliares fora).
Public Function Break_CategoryTops(cht As Chart) As Double()
    Dim tops() As Double, n As Long, i As Long, s As Series, v As Variant, x As Double, flat As Boolean, kind As String
    kind = Registry_ChartType(cht)
    flat = (kind = "clustered" Or kind = "barclustered")
    n = Geom_ChartCategoryCount(cht)
    If n < 1 Then ReDim tops(0 To 0): Break_CategoryTops = tops: Exit Function
    ReDim tops(1 To n)
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            v = s.Values
            For i = 1 To n
                x = 0
                If IsArray(v) Then If LBound(v) + i - 1 <= UBound(v) Then x = Util_ToDouble(v(LBound(v) + i - 1))
                If x > 0 Then
                    If flat Then
                        If x > tops(i) Then tops(i) = x
                    Else
                        tops(i) = tops(i) + x
                    End If
                End If
            Next
        End If
    Next
    Break_CategoryTops = tops
End Function

' Sugestao de corte: 1,2 x o segundo maior topo (se menor que o maior), senao metade do maior.
Public Function Break_Suggest(cht As Chart) As Double
    Dim tops() As Double, i As Long, mx As Double, second As Double
    tops = Break_CategoryTops(cht)
    For i = LBound(tops) To UBound(tops)
        If tops(i) > mx Then
            second = mx: mx = tops(i)
        ElseIf tops(i) > second Then
            second = tops(i)
        End If
    Next
    If second > 0 And second * 1.2 < mx Then Break_Suggest = second * 1.2 Else Break_Suggest = mx / 2
End Function

Public Sub Break_Set(cht As Chart, limit As Double)
    Dim kind As String, tops() As Double, i As Long, mx As Double
    kind = Registry_ChartType(cht)
    If Not Kinds_Supports(kind, "break") Then Err.Raise vbObjectError + 7002, "Break_Set", "'" & Kinds_FeatureLabel("break") & "' não se aplica a este tipo de gráfico."
    If limit <= 0 Then Break_Remove cht: Exit Sub
    tops = Break_CategoryTops(cht)
    For i = LBound(tops) To UBound(tops)
        If tops(i) > mx Then mx = tops(i)
    Next
    If limit >= mx Then Err.Raise vbObjectError + 7003, "Break_Set", BK_ERR_LIMIT
    ' mesma escala e corte sao excludentes: a escala fixa (vmin/vmax) sai e o minimo volta ao automatico
    Registry_SetOption cht, "vmin", ""
    Registry_SetOption cht, "vmax", ""
    cht.Axes(xlValue, xlPrimary).MinimumScaleIsAuto = True
    Registry_SetOption cht, "vbreak", Deco_ScaleText(limit)
    Break_EnsureColumn cht
    LimitCell(cht).Value = limit
    Break_RepointTotals cht
    Break_ApplyAxis cht
    Break_EnsureRoom cht
    Break_SyncLabels cht
    Break_SyncMarks cht
    Ann_Refresh cht
End Sub

' O rotulo do total passa a seguir LblPos (vertical: Values; barras: XValues da auxiliar XY).
Private Sub Break_RepointTotals(cht As Chart)
    Dim s As Series, blk As Range, body As Range, c As Long, n As Long
    Set s = Chart_FindSeries(cht, "tqq_total")
    If s Is Nothing Then Exit Sub
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    c = Block_FindCol(blk, BK_HDR)
    If c = 0 Then Exit Sub
    n = body.Rows.Count - 1
    If Kinds_IsHorizontal(Registry_ChartType(cht)) Then
        s.XValues = Block_Ref(body.Cells(2, c).Resize(n, 1))
    Else
        s.Values = Block_Ref(body.Cells(2, c).Resize(n, 1))
    End If
End Sub

Public Sub Break_Remove(cht As Chart)
    Dim i As Long, lc As Range
    Registry_SetOption cht, "vbreak", ""
    Set lc = LimitCell(cht)
    If Not lc Is Nothing Then lc.ClearContents
    cht.Axes(xlValue, xlPrimary).MaximumScaleIsAuto = True
    For i = cht.Shapes.Count To 1 Step -1
        If Left$(cht.Shapes(i).Name, Len(BK_PREFIX)) = BK_PREFIX Then cht.Shapes(i).Delete
    Next
    Break_SyncLabels cht                             ' sem limite: todos os rotulos de volta as posicoes normais
    On Error Resume Next
    cht.PlotArea.Position = xlChartElementPositionAutomatic     ' a folga de cima (so reduz) volta; as demais folgas sao refeitas
    Chart_EnsureRoomAfterChange cht
    Ann_Refresh cht
    If Err.Number <> 0 Then Errors_Log "Break_Remove" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Reaplica o maximo fixo se o eixo voltou ao automatico (ex.: Escala automatica) ou difere do limite. Sem corte nao faz nada.
Public Sub Break_ApplyAxis(cht As Chart)
    Dim lim As Double, ax As Axis
    lim = Break_Limit(cht)
    If lim <= 0 Then Exit Sub
    Set ax = cht.Axes(xlValue, xlPrimary)
    If ax.MaximumScaleIsAuto Or Abs(ax.MaximumScale - lim) > 0.000000001 * (1 + lim) Then ax.MaximumScale = lim
End Sub

' Folga para o rotulo do total da coluna cortada: 24 pt acima da area interna (colunas) ou a largura do maior rotulo de total + margem a
' direita (barras). So aumenta a folga. Chamada por Break_Set e por Chart_EnsureRoomAfterChange. Nunca propaga erro.
Public Sub Break_EnsureRoom(cht As Chart)
    Dim k As Long, delta As Double, need As Double
    On Error Resume Next
    If Break_Limit(cht) <= 0 Then Exit Sub
    If Geom_ChartIsHorizontal(cht) Then
        need = Labels_MaxTotalWidth(cht) + BK_RIGHT_MARGIN
        For k = 1 To 3
            delta = need - (Geom_ChartRightEdge(cht) - (cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth))
            If delta <= 0.5 Then Exit For
            If Not Deco_ShrinkPlotWidth(cht, delta) Then Exit For
        Next
    Else
        For k = 1 To 3
            delta = BK_TOP_ROOM - cht.PlotArea.InsideTop
            If delta <= 0.01 Then Exit For
            cht.PlotArea.Height = cht.PlotArea.Height - delta
            cht.PlotArea.Top = cht.PlotArea.Top + delta
        Next
    End If
    If Err.Number <> 0 Then Errors_Log "Break_EnsureRoom" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Rotulos nativos dos segmentos (opcao labels ligada): segmento com a base >= limite fica sem rotulo; segmento atravessado pelo limite
' vai para a base (InsideBase); os demais voltam a posicao normal (Center nas empilhadas, OutsideEnd nas agrupadas). Sem limite, todos
' voltam ao normal. Um rotulo que precisa voltar e recriado por Stacked_SetLabels (um ponto re-ligado nao herda o conteudo da serie).
' So escreve o que mudou. Nunca propaga erro.
Public Sub Break_SyncLabels(cht As Chart)
    Dim lim As Double, kind As String, flat As Boolean, n As Long, i As Long, s As Series, v As Variant, x As Double, base() As Double
    Dim want As Long, hide As Boolean, needReset As Boolean, pass As Long, normalPos As Long
    On Error GoTo Fail
    If Registry_GetOption(cht, "labels", "1") <> "1" Then Exit Sub
    kind = Registry_ChartType(cht)
    flat = (kind = "clustered" Or kind = "barclustered")
    normalPos = IIf(flat, xlLabelPositionOutsideEnd, xlLabelPositionCenter)
    lim = Break_Limit(cht)
    n = Geom_ChartCategoryCount(cht)
    If n < 1 Then Exit Sub
    For pass = 1 To 2
        ReDim base(1 To n)
        For Each s In cht.SeriesCollection
            If Not Style_IsHelperSeries(s) Then
                v = s.Values
                For i = 1 To n
                    x = 0
                    If IsArray(v) Then If LBound(v) + i - 1 <= UBound(v) Then x = Util_ToDouble(v(LBound(v) + i - 1))
                    hide = False: want = normalPos
                    If lim > 0 And x > 0 Then
                        If base(i) >= lim Then
                            hide = True
                        ElseIf base(i) + x > lim Then
                            want = xlLabelPositionInsideBase
                        End If
                    End If
                    If x > 0 And Not flat Then base(i) = base(i) + x
                    If pass = 1 Then
                        If Not hide And Not s.Points(i).HasDataLabel Then needReset = True
                    Else
                        If hide Then
                            If s.Points(i).HasDataLabel Then s.Points(i).HasDataLabel = False
                        ElseIf s.Points(i).HasDataLabel Then
                            If s.Points(i).DataLabel.Position <> want Then s.Points(i).DataLabel.Position = want
                        End If
                    End If
                Next
            End If
        Next
        If pass = 1 And needReset Then Stacked_SetLabels cht, True
    Next
    Exit Sub
Fail:
    Errors_Log "Break_SyncLabels" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Uma marca por categoria com topo acima do limite: tqq_ann_break_<i>_w (branca, 3 pt, criada primeiro: fica atras), tqq_ann_break_<i>
' (cinza) e tqq_ann_break_<i>_b (cinza, deslocada). Cria as que faltam e apaga as que sobram; Ann_Refresh posiciona (Ann_LayoutBreak).
Public Sub Break_SyncMarks(cht As Chart)
    Dim lim As Double, tops() As Double, i As Long, nm As String, want As Boolean, shp As Shape, sfx As Variant, k As Long
    On Error GoTo Fail
    lim = Break_Limit(cht)
    tops = Break_CategoryTops(cht)
    For i = 1 To UBound(tops)
        nm = BK_PREFIX & i
        want = (lim > 0 And tops(i) > lim)
        If want And Not HasMark(cht, nm) Then
            For Each sfx In Array("_w", "", "_b")
                Set shp = cht.Shapes.AddLine(0, 0, 10, 0)
                shp.Name = nm & sfx
                shp.AlternativeText = "kind=break;cat=" & i & ";part=" & IIf(sfx = "_w", "w", IIf(sfx = "_b", "b", "a"))
                If sfx = "_w" Then
                    shp.Line.ForeColor.RGB = RGB(255, 255, 255): shp.Line.Weight = 3
                Else
                    shp.Line.ForeColor.RGB = RGB(89, 89, 89): shp.Line.Weight = 1
                End If
            Next
        ElseIf Not want And HasMark(cht, nm) Then
            For k = cht.Shapes.Count To 1 Step -1
                If cht.Shapes(k).Name = nm Or cht.Shapes(k).Name = nm & "_w" Or cht.Shapes(k).Name = nm & "_b" Then cht.Shapes(k).Delete
            Next
        End If
    Next
    Exit Sub
Fail:
    Errors_Log "Break_SyncMarks" & vbTab & Err.Number & vbTab & Err.Description
End Sub

Private Function HasMark(cht As Chart, nm As String) As Boolean
    Dim shp As Shape
    For Each shp In cht.Shapes
        If shp.Name = nm Then HasMark = True: Exit Function
    Next
End Function

' O que no corte depende dos DADOS e muda depois de inserir (Core_SafeRefresh): o maximo fixo (se o eixo voltou ao automatico), os rotulos
' dos segmentos e as marcas. Nunca propaga erro; guarda de reentrada (as marcas/rotulos disparam Calculate).
Public Sub Break_RefreshAfterDataChange(cht As Chart)
    If mBusy Then Exit Sub
    mBusy = True
    On Error Resume Next
    If Kinds_Supports(Registry_ChartType(cht), "break") Then
        If Break_Limit(cht) > 0 Then
            Break_ApplyAxis cht
            Break_SyncLabels cht
            Break_SyncMarks cht
        End If
    End If
    If Err.Number <> 0 Then Errors_Log "Break_RefreshAfterDataChange" & vbTab & Err.Number & vbTab & Err.Description
    mBusy = False
End Sub

' Depois de "Alterar dados" o grafico novo tem a opcao copiada mas nenhuma coluna/marca: reaplica; se o limite ja nao vale para os dados
' novos (7003), limpa a opcao e registra. Nunca propaga erro.
Public Sub Break_TryReapply(cht As Chart)
    Dim lim As Double
    On Error GoTo Fail
    lim = Break_Limit(cht)
    If lim <= 0 Then Exit Sub
    Break_Set cht, lim
    Exit Sub
Fail:
    Errors_Log "Break_TryReapply" & vbTab & Err.Number & vbTab & Err.Description
    On Error Resume Next
    Registry_SetOption cht, "vbreak", ""
End Sub
```

- [ ] **Passo 4: Ligações** — `modChartStacked.Stacked_SetTotals`: declare `pc As Long`; depois de `n = body.Rows.Count - 1` acrescente `pc = Block_FindCol(blk, "LblPos"): If pc = 0 Then pc = 3` e use `body.Cells(2, pc).Resize(n, 1)` como **valueCells** nas duas chamadas de `Chart_AddHelperSeries` (o rótulo continua ligado à coluna 2 e `Stacked_EnsureTotalsRoom` recebe `body.Cells(2, pc)`); em `Stacked_EnsureRoomAfterChange` idem (`pc`). `modAnnotations`: em `Ann_RemoveAll` acrescente `And Left$(nm, Len("tqq_ann_break_")) <> "tqq_ann_break_"` à condição; em `Ann_Refresh` acrescente, antes do `ElseIf kind = "ind100"`:

```vb
            ElseIf kind = "break" Then
                valid = Kinds_Supports(chartKind, "break") And Len(Registry_GetOption(cht, "vbreak", "")) > 0
                If valid Then
                    On Error Resume Next
                    Ann_LayoutBreak cht, shp
                    If Err.Number <> 0 Then
                        valid = False
                        mLastError = "Ann_Refresh " & nm & " " & Err.Number & " " & Err.Description
                        Errors_Log mLastError
                        Err.Clear
                    End If
                    On Error GoTo Done
                End If
                If shp.Visible <> IIf(valid, msoTrue, msoFalse) Then shp.Visible = IIf(valid, msoTrue, msoFalse)
```
e a função privada:

```vb
' Marca de corte: duas linhas diagonais atravessando a coluna (ou barra) cortada, 8 pt para dentro da borda do corte (a branca de 3 pt por
' baixo apaga o preenchimento; a parte "b" fica 4 pt deslocada). Colunas: de (cx - meia - 3, y + 3) a (cx + meia + 3, y - 3) com
' y = valor(limite) + 8; barras: de (x - 3, cy - meia - 3) a (x + 3, cy + meia + 3) com x = valor(limite) - 8.
Private Sub Ann_LayoutBreak(cht As Chart, shp As Shape)
    Dim cat As Long, part As String, lim As Double, half As Double, c As Double, v As Double, off As Double
    cat = CLng(Val(Ann_ParseParam(shp.AlternativeText, "cat")))
    part = Ann_ParseParam(shp.AlternativeText, "part")
    lim = Val(Registry_GetOption(cht, "vbreak", ""))
    If cat < 1 Or lim <= 0 Then Exit Sub
    half = Geom_ChartBarHalfWidth(cht)
    c = Geom_ChartCatCoord(cht, cat)
    v = Geom_ChartValCoord(cht, lim)
    off = IIf(part = "b", 4, 0)
    If Geom_ChartIsHorizontal(cht) Then
        Ann_MoveLine shp, v - 8 - 3 - off, c - half - 3, v - 8 + 3 - off, c + half + 3
    Else
        Ann_MoveLine shp, c - half - 3, v + 8 + 3 + off, c + half + 3, v + 8 - 3 + off
    End If
End Sub
```
Em `Ann_GetCategoryData` declare `lim As Double` e acrescente, depois do `End Select` (antes do `End Sub`), o limite do topo pelo corte: sem ele as setas CAGR e diferença de um gráfico cortado ancoram no topo real (ex.: 230) e ficam acima da área de plotagem; `values` (os números dos rótulos das setas) continua com os valores reais:

```vb
    lim = Break_Limit(cht)                          ' corte de eixo: o topo plotado para na borda do corte (as setas ancoram ali)
    If lim > 0 Then
        For i = 1 To n
            If tops(i) > lim Then tops(i) = lim
        Next
    End If
```
`modDecorations.FreeScaleKind`: `FreeScaleKind = IIf(Kinds_Supports(Registry_ChartType(cht), "samescale") And Len(Registry_GetOption(cht, "vbreak", "")) = 0, 1, 0)`. `modChartUtil.Chart_EnsureRoomAfterChange`: depois do `Gantt_EnsureRoom`, `Break_EnsureRoom cht` (log `"Chart_EnsureRoomAfterChange/corte"`). `modCoreApi`: `Core_SafeRefresh` chama `Break_RefreshAfterDataChange cht` logo depois de `Deco_ApplyFixedScale` (log `"Core_SafeRefresh/Break"`); `ReapplyOptions`: depois de `Deco_ApplyFixedScale nc`, `Break_TryReapply nc`; e os comandos:

```vb
' Corte de eixo: pede o valor (sugestao: Break_Suggest) e aplica. Em build Test sai antes do InputBox.
Public Sub Core_SetAxisBreak()
    On Error GoTo Fail
    Dim c As Chart, v As Variant
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    If Not Kinds_SupportsChart(c, "break") Then
        Core_Notify "'" & Kinds_FeatureLabel("break") & "' não se aplica a este tipo de gráfico.", vbInformation
        Exit Sub
    End If
    If BUILD_CONFIG = "Test" Then Exit Sub
    v = Application.InputBox("Corte de eixo: valor onde o eixo é cortado (as colunas acima dele ficam cortadas e o total continua no rótulo):", APP_NAME, Break_Suggest(c), Type:=1)
    If VarType(v) = vbBoolean Then Exit Sub
    Core_SetAxisBreakAt CDbl(v)
    Exit Sub
Fail:
    Errors_Handle "Core_SetAxisBreak"
End Sub

Public Sub Core_SetAxisBreakAt(limit As Double)
    On Error GoTo Fail
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    Break_Set c, limit
    Core_SafeRefresh c
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7002 Or Err.Number = vbObjectError + 7003 Then
        Core_Notify Err.Description, vbInformation       ' recusa (tipo sem corte, limite invalido): aviso, nada gravado
    Else
        Errors_Handle "Core_SetAxisBreakAt"
    End If
End Sub

Public Sub Core_RemoveAxisBreak()
    On Error GoTo Fail
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    If Len(Registry_GetOption(c, "vbreak", "")) = 0 Then Exit Sub
    Break_Set c, 0
    Core_SafeRefresh c
    Exit Sub
Fail:
    Errors_Handle "Core_RemoveAxisBreak"
End Sub
```

- [ ] **Passo 5: GREEN e suíte completa.** Exporte PNG de colunas e barras cortadas (coluna cortada, total acima, marca, rótulo na base).

- [ ] **Passo 6: Commit** — `git add src/core/modules/modAxisBreak.bas src/core/modules/modChartStacked.bas src/core/modules/modAnnotations.bas src/core/modules/modDecorations.bas src/core/modules/modChartUtil.bas src/core/modules/modCoreApi.bas tests/vba/core/TestAxisBreak.bas tests/vba/core/TestCoreApi.bas` e `git commit -m "feat: corte de eixo (maximo fixo, total vivo em LblPos, marca de corte e rotulos sincronizados)"` (+ trailer).

**Verificação no Excel real:** Colunas empilhadas com um valor 10× maior; `Core_SetAxisBreakAt` (ou Mais opções > Corte de eixo... após a Tarefa 8): coluna cortada, total real acima, marca; Remover corte de eixo devolve a escala.

---
### Tarefa 5: Copiar para o PowerPoint (imagem e gráfico vinculado) — modelo: standard

**Files:**
- Create: `src/core/modules/modPowerPoint.bas`
- Modify: `src/core/modules/modCoreApi.bas` (`Core_CopyToPowerPoint`)
- Test: `tests/vba/core/TestPowerPoint.bas` (novo)

**Interfaces:**
- Consumes: `Core_ActiveChart`, `Registry_Workbook`, `Core_Notify`, `Errors_Log`, `Errors_HandleErr`, `BUILD_CONFIG`.
- Produces (`modPowerPoint`, automação tardia, sem referência, sem `Declare`): `Ppt_SetAppForTest(app As Object)` (só em `Test`); `Ppt_IsRunning() As Boolean` (`GetObject` funciona); `Ppt_GetApp(create As Boolean) As Object` (`GetObject` → `CreateObject` se `create`; `Nothing` se não há/não cria; em build `Test`, só a instância injetada por `Ppt_SetAppForTest`, e sem ela `Nothing`: um teste nunca anexa ao PowerPoint do usuário); `Ppt_TargetSlide(app) As Object` (slide ativo da janela ativa; sem janela: último slide da primeira apresentação ou um novo; sem apresentação: `Presentations.Add(msoTrue)` + slide em branco); `Ppt_PasteImage(cht, slide) As Object` (`CopyPicture xlScreen, xlPicture` + `Shapes.Paste`, ajustado a 80% do slide e centralizado); `Ppt_PasteChart(co As ChartObject, slide) As Object` (`co.Copy` + `Shapes.Paste`: gráfico do PowerPoint vinculado à pasta salva, com as formas das anotações); `Ppt_FitSize(w, h, slideW, slideH, ByRef outW, ByRef outH)` (pura); `Ppt_LinkedPrecheck(wb) As String` (`""` ou `PPT_ERR_UNSAVED`); `Ppt_ErrorText(errNum, errDesc) As String`; constantes públicas `PPT_ERR_NOT_INSTALLED = "O PowerPoint não está instalado neste computador (ou não pôde ser iniciado)."`, `PPT_ERR_UNSAVED = "Salve a pasta de trabalho antes de copiar o gráfico vinculado: o vínculo aponta para o arquivo salvo."`, `PPT_NOTE_DIRTY = "A pasta tem alterações não salvas: o PowerPoint lê o arquivo salvo nas atualizações do vínculo."`.
- Produces (`modCoreApi`): `Core_CopyToPowerPoint(mode As String)` (`"image"` | `"chart"`): sem gráfico ThinQQ ativo não faz nada; modo desconhecido vai para o log; vinculado com pasta não salva → aviso `PPT_ERR_UNSAVED` **antes** de tocar no PowerPoint; PowerPoint ausente (429) → `PPT_ERR_NOT_INSTALLED`; sucesso → `"Gráfico copiado para o slide N de '<apresentação>' como imagem."` / `"... como gráfico vinculado."` (+ `PPT_NOTE_DIRTY` quando `Not wb.Saved`); o aviso sai **antes** de o PowerPoint ficar visível e ativo (a caixa do Release não fica escondida atrás dele); no build `Test` o PowerPoint não é mostrado nem ativado (não rouba o foco da área de trabalho no meio da suíte).
- Regra de teste: o teste de integração só cria o PowerPoint quando `Ppt_IsRunning() = False`, não há nenhum processo `POWERPNT` (consulta WMI: a ROT de `GetObject` não vê um PowerPoint aberto só para incorporação) e `CreateObject` funciona; usa `Presentations.Add(msoFalse)`, injeta a instância por `Ppt_SetAppForTest` e a encerra (`Quit`) ao final, só se ninguém mais anexou a ela (`Presentations.Count = 0`). A colagem usa a área de transferência do Windows (o conteúdo anterior do usuário é substituído): limitação registrada no relatório e no README (Tarefa 11).

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestPowerPoint.bas`:

```vb
Attribute VB_Name = "TestPowerPoint"
Option Explicit

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2024, 10, 5): ws.Range("A3:C3").Value = Array(2025, 20, 6)
    Set Fixture = ws.Range("A1:C3")
End Function

Public Sub Test_Ppt_FitSize()
    Dim w As Double, h As Double
    Ppt_FitSize 460, 250, 960, 540, w, h
    AssertAlmost 768, w, 0.001: AssertAlmost 250 * 768 / 460, h, 0.001, "80% da largura, proporcao mantida"
    Ppt_FitSize 460, 280, 960, 540, w, h
    AssertAlmost 432, h, 0.001, "o grafico padrao (460 x 280) e limitado pela altura": AssertAlmost 460 * 432 / 280, w, 0.001
    Ppt_FitSize 200, 600, 960, 540, w, h
    AssertAlmost 432, h, 0.001, "limitado por 80% da altura": AssertAlmost 200 * 432 / 600, w, 0.001
    Ppt_FitSize 0, 0, 960, 540, w, h
    AssertEqual 0, w, "tamanho invalido fica como esta"
End Sub

' Review Focus 5 (parte 1): pasta nao salva e vinculo.
Public Sub Test_Ppt_LinkedPrecheck_UnsavedWorkbook()
    Dim wb As Workbook, p As String
    Set wb = NewTempWorkbook()
    AssertEqual PPT_ERR_UNSAVED, Ppt_LinkedPrecheck(wb)
    p = TempDir() & "\ppt.xlsx"
    wb.SaveAs p, xlOpenXMLWorkbook
    AssertEqual "", Ppt_LinkedPrecheck(wb), "pasta salva"
    CloseTemp wb
    AssertEqual PPT_ERR_NOT_INSTALLED, Ppt_ErrorText(429, "x")
    AssertTrue InStr(Ppt_ErrorText(70, "Permissão negada"), "Permissão negada") > 0
End Sub

Public Sub Test_Core_CopyToPowerPoint_Unsaved_Linked_NoticeBeforePowerPoint()
    Dim wb As Workbook, rng As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Core_ClearLastNotice
    Core_CopyToPowerPoint "chart"
    AssertEqual PPT_ERR_UNSAVED, Core_LastNotice(), "vinculado com pasta nao salva: so o aviso"
    Core_ClearLastNotice
    Core_CopyToPowerPoint "xyz"
    AssertEqual "", Core_LastNotice(), "modo desconhecido: nada (so log)"
    rng.Worksheet.Range("A1").Select
    Core_ClearLastNotice
    Core_CopyToPowerPoint "image"
    AssertEqual "", Core_LastNotice(), "sem grafico: nada"
    Teardown wb
End Sub

' Review Focus 5 (parte 2): colagem real. Roda SO quando nenhum PowerPoint esta aberto (aplicativo de instancia unica: CreateObject anexaria
' ao do usuario) e o PowerPoint existe; senao sai verde registrando o motivo. Apresentacao sem janela; a instancia e encerrada ao final.
Public Sub Test_Ppt_RealPaste_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, en As Long, ed As String, shp As Object
    If Ppt_IsRunning() Then Errors_Log "TestPowerPoint: pulado (ha um PowerPoint aberto)": Exit Sub
    ' a ROT (GetObject) nao ve um PowerPoint aberto so para incorporacao: confere tambem os processos (regra global: nenhum POWERPNT)
    If GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'POWERPNT.EXE'").Count > 0 Then Errors_Log "TestPowerPoint: pulado (processo POWERPNT existente)": Exit Sub
    On Error Resume Next
    Set app = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If app Is Nothing Then Errors_Log "TestPowerPoint: pulado (PowerPoint nao instalado)": Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0)
    Set slide = pres.Slides.Add(1, 12)
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\ppt-src.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ann_AddCagr c, 1, 2
    Ppt_SetAppForTest app
    Core_ClearLastNotice
    Core_CopyToPowerPoint "image"
    AssertEqual 1, slide.Shapes.Count, "imagem colada"
    AssertEqual 13, slide.Shapes(1).Type, "msoPicture"
    AssertAlmost 432, CDbl(slide.Shapes(1).Height), 2, "limitado pela altura (432 pt = 80% de 540): largura ~710 (" & slide.Shapes(1).Width & ")"
    AssertTrue InStr(Core_LastNotice(), "slide 1") > 0 And InStr(Core_LastNotice(), "como imagem") > 0, Core_LastNotice()
    Core_ClearLastNotice
    Core_CopyToPowerPoint "chart"
    AssertEqual 2, slide.Shapes.Count, "grafico colado"
    Set shp = slide.Shapes(2)
    AssertEqual -1, shp.HasChart, "gráfico do PowerPoint"
    AssertEqual -1, CLng(shp.Chart.ChartData.IsLinked), "vinculado a pasta salva"
    AssertTrue shp.Chart.Shapes.Count >= 2, "formas da seta CAGR preservadas: " & shp.Chart.Shapes.Count
    AssertTrue InStr(Core_LastNotice(), "gráfico vinculado") > 0, Core_LastNotice()
    rng.Worksheet.Range("B2").Value = 11
    Core_ClearLastNotice
    Core_CopyToPowerPoint "chart"
    AssertTrue InStr(Core_LastNotice(), PPT_NOTE_DIRTY) > 0, "aviso de alteracoes nao salvas"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Ppt_SetAppForTest Nothing
    If Not pres Is Nothing Then pres.Close
    If Not app Is Nothing Then If app.Presentations.Count = 0 Then app.Quit      ' alguem anexou a esta instancia e abriu algo: nao encerra
    Set app = Nothing
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPowerPoint", ed
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "TestPowerPoint"` → saída 3.

- [ ] **Passo 3: Implementar `modPowerPoint.bas`**:

```vb
Attribute VB_Name = "modPowerPoint"
Option Explicit
' Copiar o grafico para o PowerPoint por automacao tardia (CreateObject/GetObject; sem referencia, sem Declare; constantes numericas).
' Imagem: Chart.CopyPicture (xlScreen, xlPicture) + Shapes.Paste (funciona com o Excel invisivel: medido). Grafico vinculado: ChartObject.Copy
' + Shapes.Paste cria um grafico do PowerPoint com ChartData.IsLinked = True e as formas das anotacoes preservadas (medido); exige a pasta
' SALVA (com pasta nova o vinculo aponta para um arquivo inexistente). O PowerPoint e aplicativo de instancia unica: GetObject anexa ao
' aberto; CreateObject abre um novo. Application.Visible = msoFalse e recusado pelo PowerPoint (os testes usam Presentations.Add(msoFalse)).
Private mTestApp As Object
Private Const PPT_LAYOUT_BLANK As Long = 12     ' ppLayoutBlank
Private Const PPT_FIT As Double = 0.8           ' fracao do slide que o grafico ocupa (largura ou altura, a que limitar)
Public Const PPT_ERR_NOT_INSTALLED As String = "O PowerPoint não está instalado neste computador (ou não pôde ser iniciado)."
Public Const PPT_ERR_UNSAVED As String = "Salve a pasta de trabalho antes de copiar o gráfico vinculado: o vínculo aponta para o arquivo salvo."
Public Const PPT_NOTE_DIRTY As String = "A pasta tem alterações não salvas: o PowerPoint lê o arquivo salvo nas atualizações do vínculo."

' Gancho de teste: instancia injetada (inerte fora do build Test).
Public Sub Ppt_SetAppForTest(app As Object)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    Set mTestApp = app
End Sub

Public Function Ppt_IsRunning() As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = GetObject(, "PowerPoint.Application")
    Ppt_IsRunning = (Err.Number = 0) And (Not o Is Nothing)
    Err.Clear
End Function

Public Function Ppt_GetApp(create As Boolean) As Object
    Dim o As Object
    If BUILD_CONFIG = "Test" Then Set Ppt_GetApp = mTestApp: Exit Function     ' testes: so a instancia injetada (sem ela, Nothing): nunca o PowerPoint do usuario
    On Error Resume Next
    Set o = GetObject(, "PowerPoint.Application")
    If o Is Nothing And create Then Set o = CreateObject("PowerPoint.Application")
    Err.Clear
    On Error GoTo 0
    Set Ppt_GetApp = o
End Function

' Slide de destino (ver cabecalho). Propaga erros do PowerPoint.
Public Function Ppt_TargetSlide(app As Object) As Object
    Dim pres As Object, slide As Object
    On Error Resume Next
    If app.Windows.Count > 0 Then Set slide = app.ActiveWindow.View.Slide
    Err.Clear
    On Error GoTo 0
    If Not slide Is Nothing Then Set Ppt_TargetSlide = slide: Exit Function
    If app.Presentations.Count > 0 Then
        Set pres = app.Presentations(1)
        If pres.Slides.Count = 0 Then Set slide = pres.Slides.Add(1, PPT_LAYOUT_BLANK) Else Set slide = pres.Slides(pres.Slides.Count)
    Else
        Set pres = app.Presentations.Add(-1)        ' msoTrue: com janela, para o usuario ver
        Set slide = pres.Slides.Add(1, PPT_LAYOUT_BLANK)
    End If
    Set Ppt_TargetSlide = slide
End Function

' Tamanho que cabe em PPT_FIT do slide mantendo a proporcao (pura). Tamanho invalido fica como esta.
Public Sub Ppt_FitSize(w As Double, h As Double, slideW As Double, slideH As Double, ByRef outW As Double, ByRef outH As Double)
    Dim f As Double
    outW = w: outH = h
    If w <= 0 Or h <= 0 Or slideW <= 0 Or slideH <= 0 Then Exit Sub
    f = PPT_FIT * slideW / w
    If h * f > PPT_FIT * slideH Then f = PPT_FIT * slideH / h
    outW = w * f: outH = h * f
End Sub

Public Function Ppt_LinkedPrecheck(wb As Workbook) As String
    If Len(wb.Path) = 0 Then Ppt_LinkedPrecheck = PPT_ERR_UNSAVED
End Function

Public Function Ppt_ErrorText(errNum As Long, errDesc As String) As String
    If errNum = 429 Then Ppt_ErrorText = PPT_ERR_NOT_INSTALLED Else Ppt_ErrorText = "Não foi possível colar no PowerPoint: " & errDesc & " (" & errNum & ")"
End Function

Private Sub Ppt_PlaceShape(shp As Object, slide As Object)
    Dim sw As Double, sh As Double, w As Double, h As Double
    sw = slide.Parent.PageSetup.SlideWidth
    sh = slide.Parent.PageSetup.SlideHeight
    Ppt_FitSize CDbl(shp.Width), CDbl(shp.Height), sw, sh, w, h
    shp.LockAspectRatio = -1
    shp.Width = w
    shp.Height = h
    shp.Left = (sw - shp.Width) / 2
    shp.Top = (sh - shp.Height) / 2
End Sub

Public Function Ppt_PasteImage(cht As Chart, slide As Object) As Object
    Dim sr As Object, shp As Object
    cht.CopyPicture xlScreen, xlPicture
    Set sr = slide.Shapes.Paste
    Set shp = sr.Item(1)
    Ppt_PlaceShape shp, slide
    Set Ppt_PasteImage = shp
End Function

Public Function Ppt_PasteChart(co As ChartObject, slide As Object) As Object
    Dim sr As Object, shp As Object
    co.Copy
    Set sr = slide.Shapes.Paste
    Set shp = sr.Item(1)
    Ppt_PlaceShape shp, slide
    Set Ppt_PasteChart = shp
End Function
```

E em `modCoreApi`:

```vb
' Copiar para o PowerPoint: "image" (figura) ou "chart" (grafico do PowerPoint vinculado a pasta salva). Sem grafico ativo nao faz nada;
' a pasta nao salva e recusada ANTES de abrir o PowerPoint; PowerPoint ausente (429) vira aviso.
Public Sub Core_CopyToPowerPoint(mode As String)
    On Error GoTo Fail
    Dim c As Chart, wb As Workbook, app As Object, slide As Object, shp As Object, msg As String, linked As Boolean, en As Long, ed As String
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    If mode <> "image" And mode <> "chart" Then
        Errors_Log "Core_CopyToPowerPoint: modo desconhecido (ignorado)" & vbTab & mode
        Exit Sub
    End If
    linked = (mode = "chart")
    Set wb = Registry_Workbook(c)
    If linked Then
        msg = Ppt_LinkedPrecheck(wb)
        If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Sub
    End If
    Set app = Ppt_GetApp(True)
    If app Is Nothing Then Core_Notify PPT_ERR_NOT_INSTALLED, vbExclamation: Exit Sub
    Set slide = Ppt_TargetSlide(app)
    If linked Then Set shp = Ppt_PasteChart(c.Parent, slide) Else Set shp = Ppt_PasteImage(c, slide)
    msg = "Gráfico copiado para o slide " & slide.SlideIndex & " de '" & slide.Parent.Name & "' como " & IIf(linked, "gráfico vinculado", "imagem") & "."
    If linked And Not wb.Saved Then msg = msg & " " & PPT_NOTE_DIRTY
    Core_Notify msg, vbInformation                   ' ANTES de ativar o PowerPoint: a caixa (Release) nao fica escondida atras dele
    If BUILD_CONFIG <> "Test" Then                   ' nos testes o PowerPoint do teste nao aparece nem rouba o foco
        On Error Resume Next
        app.Visible = -1
        app.Activate
        Err.Clear
        On Error GoTo Fail
    End If
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    If en = 429 Then
        Core_Notify PPT_ERR_NOT_INSTALLED, vbExclamation
    Else
        Errors_HandleErr "Core_CopyToPowerPoint(" & mode & ")", en, Ppt_ErrorText(en, ed)
    End If
End Sub
```

- [ ] **Passo 4: GREEN e suíte completa.** Registre no relatório se o teste real rodou ou foi pulado (e por quê), que a colagem usa a área de transferência do Windows (o teste real a sobrescreve: limitação conhecida), e confirme com `Get-Process powerpnt` que nenhum processo criado pelo teste sobrou (se sobrar, encerre **só** o processo que o teste criou: compare a lista antes e depois).

- [ ] **Passo 5: Commit** — `git add src/core/modules/modPowerPoint.bas src/core/modules/modCoreApi.bas tests/vba/core/TestPowerPoint.bas` e `git commit -m "feat: copiar grafico para o PowerPoint como imagem ou grafico vinculado (automacao tardia)"` (+ trailer).

**Verificação no Excel real:** com o PowerPoint fechado, `Core_CopyToPowerPoint "image"` abre o PowerPoint com a imagem; com uma apresentação aberta, cola no slide ativo; "chart" numa pasta não salva avisa; numa pasta salva cola o gráfico vinculado (Dados > Editar Vínculos no PowerPoint mostra a pasta).

---

### Tarefa 6: Ícones próprios na faixa (build) — modelo: standard

**Files:**
- Create: `build/lib/Ribbon-Icons.ps1`
- Modify: `build/lib/Build-Common.ps1` (`Add-RibbonToPackage -Icons`), `build/build.ps1` (`-RibbonIcons Png|Mso`), `tests/ps/Test-Ribbon.ps1`, `tests/vba/loader/TestRibbon.bas`
- Test: `tests/ps/Test-Ribbon.ps1`, `tests/vba/loader/TestRibbon.bas`

**Interfaces:**
- Produces (`Ribbon-Icons.ps1`, dot-sourced por `Build-Common.ps1`): `Get-TqqIconSpecs` (hashtable ordenada nome → `@{ Rects = @(@{x;y;w;h;c}, ...); Mso = '<imageMso de reserva>' }`, numa grade de 16 × 16; 13 nomes: `tqq_columns`, `tqq_bars`, `tqq_waterfall`, `tqq_lines`, `tqq_more`, `tqq_butterfly`, `tqq_scatter`, `tqq_bubble`, `tqq_options`, `tqq_arrows`, `tqq_ppt`, `tqq_round`, `tqq_break`); `Write-TqqPng -Path -Width -Height -Rgba` (PNG RGBA 8 bits, zlib "stored", CRC-32/Adler-32 próprios: bytes determinísticos); `New-TqqIconRgba -Size -Rects`; `Export-TqqIcons -OutDir` (grava `<nome>.png` (32 × 32) e `<nome>16.png` (16 × 16) para cada spec; devolve os caminhos); `Convert-TqqRibbonToMso -Xml` (texto do XML com cada `image="<nome>"`/`image="<nome>16"` trocado por `imageMso="<Mso>"`).
- Produces (`Build-Common.ps1`): `Add-RibbonToPackage -PackagePath -RibbonXmlPath [-Icons Png|Mso]` — `Png` (padrão): grava o XML como está, `customUI/images/<nome>.png` para cada `image=` referenciado (erro se o nome não está na tabela), `customUI/_rels/customUI14.xml.rels` (uma `Relationship Id="<nome>" Type=".../officeDocument/2006/relationships/image" Target="images/<nome>.png"` por nome) e `<Default Extension="png" ContentType="image/png"/>` em `[Content_Types].xml`; `Mso`: grava o XML convertido e nenhuma imagem. `build.ps1 -RibbonIcons Mso` (padrão `Png`).
- Produces (testes): `Test-Ribbon.ps1` verifica, para cada `image=` do XML do **pacote**, o PNG (assinatura, IHDR 32 ou 16 px conforme o sufixo `16`, abre em `System.Drawing`), a relação e o content type; sem `image=` (pacote `Mso`), que não há `customUI/images/`; o XML do pacote igual ao fonte (modo `Png`, ou fonte ainda sem `image=`) ou igual ao fonte convertido por `Convert-TqqRibbonToMso` (modo `Mso`). `TestRibbon.bas`: `Test_Ribbon_Images_NamesAndNoDoubleIcon` (todo `@image` casa `^tqq_[a-z0-9_]+$` e nenhum controle tem `image` e `imageMso` juntos; aceita zero imagens: a Tarefa 8 as introduz e aperta para `>= 13`) e `Test_Ribbon_IconFallbackMso_AreValid` (cada `Mso = '...'` da tabela de `build\lib\Ribbon-Icons.ps1` é um `imageMso` válido por `GetImageMso`, 13 reservas: o teste existente lê só o XML fonte).

- [ ] **Passo 1: Testes (RED)** — em `tests/ps/Test-Ribbon.ps1`, substitua o conteúdo por:

```powershell
# Verifica o pacote ThinQQ.test.xlam: customUI14.xml (identico ao fonte no modo Png), a relacao ui/extensibility, e os icones proprios:
# para cada image="x" do XML do pacote, customUI/images/x.png (PNG valido, 32 px ou 16 px conforme o sufixo "16"), a Relationship Id="x"
# em customUI/_rels/customUI14.xml.rels e o content type png. No modo Mso (sem image=) nao pode haver imagens.
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing
. (Join-Path $RepoRoot 'build\lib\Ribbon-Icons.ps1')
$pkg = Join-Path $RepoRoot 'dist\ThinQQ.test.xlam'
$src = Join-Path $RepoRoot 'src\loader\ribbon\customUI14.xml'
$script:fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }
function Read-Entry($zip, [string]$name) { $e = $zip.GetEntry($name); if (-not $e) { return $null }; $ms = New-Object System.IO.MemoryStream; $s = $e.Open(); $s.CopyTo($ms); $s.Dispose(); return ,$ms.ToArray() }
if (-not (Test-Path -LiteralPath $pkg)) { Write-Host "not ok - pacote ausente: $pkg"; exit 1 }
$fs = New-Object System.IO.FileStream($pkg, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
try {
    $xmlBytes = Read-Entry $zip 'customUI/customUI14.xml'
    Check ($null -ne $xmlBytes) 'pacote contem customUI/customUI14.xml'
    $xml = [Text.Encoding]::UTF8.GetString($xmlBytes)
    $images = @([regex]::Matches($xml, 'image="([^"]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    $srcXml = [IO.File]::ReadAllText($src)
    if ($images.Count -gt 0 -or $srcXml -notmatch 'image="') {
        # modo Png (ou fonte ainda sem image=, Tarefas 6 e 7): o XML do pacote e o fonte, byte a byte
        Check (([Convert]::ToBase64String($xmlBytes)) -eq ([Convert]::ToBase64String([IO.File]::ReadAllBytes($src)))) 'customUI14.xml do pacote identico ao fonte'
    } else {
        # modo Mso: o XML do pacote e exatamente o fonte convertido (cada image= trocado pela reserva da tabela)
        Check ($xml -eq (Convert-TqqRibbonToMso -Xml $srcXml)) 'modo Mso: customUI14.xml do pacote = fonte convertido para imageMso'
    }
    if ($images.Count -eq 0) { Check ($null -eq ($zip.Entries | Where-Object { $_.FullName -like 'customUI/images/*' })) 'pacote sem image=: sem imagens' }
    $r = Read-Entry $zip '_rels/.rels'
    Check ($null -ne $r) 'pacote contem _rels/.rels'
    if ($r) {
        $txt = [Text.Encoding]::UTF8.GetString($r)
        Check ($txt -match 'http://schemas\.microsoft\.com/office/2007/relationships/ui/extensibility') 'relationship ui/extensibility presente'
        Check ($txt -match 'customUI/customUI14\.xml') 'relationship aponta para customUI/customUI14.xml'
    }
    $specs = Get-TqqIconSpecs
    $relsBytes = Read-Entry $zip 'customUI/_rels/customUI14.xml.rels'
    $ct = [Text.Encoding]::UTF8.GetString((Read-Entry $zip '[Content_Types].xml'))
    if ($images.Count -gt 0) {
        Check ($null -ne $relsBytes) 'pacote contem customUI/_rels/customUI14.xml.rels'
        Check ($ct -match '<Default Extension="png" ContentType="image/png"\s*/>') 'content type png declarado'
        $rels = if ($relsBytes) { [Text.Encoding]::UTF8.GetString($relsBytes) } else { '' }
        foreach ($img in $images) {
            $base = $img -replace '16$', ''
            Check ($specs.Contains($base)) "icone '$img' existe na tabela de icones ($base)"
            $png = Read-Entry $zip "customUI/images/$img.png"
            Check ($null -ne $png) "pacote contem customUI/images/$img.png"
            if ($png) {
                $sig = [Convert]::ToBase64String($png[0..7])
                Check ($sig -eq [Convert]::ToBase64String([byte[]]@(0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A))) "$img.png: assinatura PNG"
                $w = ($png[16] -shl 24) -bor ($png[17] -shl 16) -bor ($png[18] -shl 8) -bor $png[19]
                $expected = if ($img -like '*16') { 16 } else { 32 }
                Check ($w -eq $expected) "$img.png: largura $expected ($w)"
                try { $ms = New-Object IO.MemoryStream(,$png); $bmp = [System.Drawing.Image]::FromStream($ms); Check ($bmp.Width -eq $expected -and $bmp.Height -eq $expected) "$img.png: System.Drawing abre $($bmp.Width)x$($bmp.Height)"; $bmp.Dispose() } catch { Check $false "$img.png: System.Drawing nao abre: $($_.Exception.Message)" }
            }
            Check ($rels -match ('Id="' + [regex]::Escape($img) + '"[^>]*Target="images/' + [regex]::Escape($img) + '\.png"')) "relacao Id=$img -> images/$img.png"
        }
        Check ($rels -match 'http://schemas\.openxmlformats\.org/officeDocument/2006/relationships/image') 'relacoes do tipo image'
    }
    # determinismo do gravador: duas gravacoes, mesmos bytes
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("tqq_icons_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force $tmp | Out-Null
    try {
        $a = Export-TqqIcons -OutDir $tmp
        $h1 = ($a | ForEach-Object { (Get-FileHash $_).Hash }) -join ','
        Remove-Item (Join-Path $tmp '*') -Force
        $b = Export-TqqIcons -OutDir $tmp
        $h2 = ($b | ForEach-Object { (Get-FileHash $_).Hash }) -join ','
        Check ($h1 -eq $h2 -and $a.Count -eq 2 * $specs.Count) "gravador de PNG deterministico ($($a.Count) arquivos para $($specs.Count) icones)"
        Check ((Convert-TqqRibbonToMso -Xml $srcXml) -notmatch 'image="') 'conversao para imageMso nao deixa image='
    } finally { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue }
} finally { $zip.Dispose(); $fs.Dispose() }
if ($script:fail -gt 0) { Write-Host "FALHAS ribbon-pacote: $script:fail"; exit 1 }
```

Em `tests/vba/loader/TestRibbon.bas`:

```vb
' Icones proprios (Fase 3): nomes image= no padrao tqq_<nome>[16] e nunca junto de imageMso no mesmo controle (Tarefa 8 exige >= 13).
Public Sub Test_Ribbon_Images_NamesAndNoDoubleIcon()
    Dim d As Object, n As Object, re As Object, nm As String, c As Long
    Set d = LoadRibbon()
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^tqq_[a-z0-9_]+$"
    For Each n In d.SelectNodes("//*[@image]")
        nm = n.Attributes.getNamedItem("image").Text
        AssertTrue re.Test(nm), "nome de imagem invalido: " & nm
        AssertTrue n.Attributes.getNamedItem("imageMso") Is Nothing, "image e imageMso no mesmo controle: " & n.Attributes.getNamedItem("id").Text
        c = c + 1
    Next
    AssertTrue c >= 0
End Sub

' As reservas imageMso da tabela de icones (build\lib\Ribbon-Icons.ps1, usadas no modo -RibbonIcons Mso) existem neste Office. O teste
' Test_Ribbon_ImageMso_AreValid le o XML fonte e nunca ve as reservas.
Public Sub Test_Ribbon_IconFallbackMso_AreValid()
    Dim txt As String, re As Object, m As Object, pic As Object, checked As Long
    txt = Transport_ReadTextFile(RepoPath() & "\build\lib\Ribbon-Icons.ps1")
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "Mso = '([^']+)'"
    For Each m In re.Execute(txt)
        Set pic = Nothing
        On Error Resume Next
        Set pic = Application.CommandBars.GetImageMso(CStr(m.SubMatches(0)), 16, 16)
        On Error GoTo 0
        AssertTrue Not pic Is Nothing, "imageMso de reserva invalido: " & m.SubMatches(0)
        checked = checked + 1
    Next
    AssertEqual 13, checked, "uma reserva por icone da tabela"
End Sub
```

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-f3\build\run-tests.ps1 -Target loader -TimeoutSec 400` → `Test-Ribbon.ps1` falha ao dot-sourcear `Ribbon-Icons.ps1` (inexistente).

- [ ] **Passo 3: Implementar `build/lib/Ribbon-Icons.ps1`** (toda a aritmética em `[uint64]`: literais hexadecimais do PowerShell que cabem em Int32 viram negativos):

```powershell
# Icones proprios da faixa: gravador de PNG puro (RGBA 8 bits; zlib com blocos "stored"; CRC-32 e Adler-32 proprios: bytes identicos em
# qualquer runtime, sem System.Drawing) e a tabela de icones (retangulos numa grade de 16 x 16, com o imageMso de reserva de cada um).
Set-StrictMode -Version Latest

$script:TqqNavy = @(31, 56, 100); $script:TqqBlue = @(68, 114, 196); $script:TqqLight = @(157, 195, 230); $script:TqqGray = @(127, 127, 127)
# nunca "R": o alias embutido r (Invoke-History) tem precedencia sobre funcoes e a chamada falharia
function New-TqqRect([double]$x, [double]$y, [double]$w, [double]$h, $c) { return @{ x = $x; y = $y; w = $w; h = $h; c = $c } }

function Get-TqqIconSpecs {
    $n = $script:TqqNavy; $b = $script:TqqBlue; $l = $script:TqqLight; $g = $script:TqqGray
    return [ordered]@{
        tqq_columns   = @{ Mso = 'ChartTypeColumnInsertGallery'; Rects = @((New-TqqRect 1 8 4 7 $n), (New-TqqRect 6 3 4 12 $b), (New-TqqRect 11 6 4 9 $l)) }
        tqq_bars      = @{ Mso = 'ChartTypeBarInsertGallery';    Rects = @((New-TqqRect 1 1 13 4 $n), (New-TqqRect 1 6 8 4 $b), (New-TqqRect 1 11 5 4 $l)) }
        tqq_waterfall = @{ Mso = 'ChartTypeOtherInsertGallery';  Rects = @((New-TqqRect 1 7 3 8 $n), (New-TqqRect 5 3 3 4 $b), (New-TqqRect 9 3 3 3 $l), (New-TqqRect 13 3 2 12 $n)) }
        tqq_lines     = @{ Mso = 'ChartTypeLineInsertGallery';   Rects = @((New-TqqRect 1 12 3 2 $b), (New-TqqRect 4 9 3 2 $b), (New-TqqRect 7 10 3 2 $b), (New-TqqRect 10 5 3 2 $b), (New-TqqRect 13 2 2 2 $b), (New-TqqRect 1 14 15 1 $g)) }
        tqq_more      = @{ Mso = 'ChartTypePieInsertGallery';    Rects = @((New-TqqRect 1 1 6 6 $n), (New-TqqRect 9 1 6 6 $b), (New-TqqRect 1 9 6 6 $l), (New-TqqRect 9 9 6 6 $g)) }
        tqq_butterfly = @{ Mso = 'ChartTypeBarInsertGallery';    Rects = @((New-TqqRect 1 2 5 3 $n), (New-TqqRect 3 6 3 3 $n), (New-TqqRect 0 10 6 3 $n), (New-TqqRect 10 2 5 3 $b), (New-TqqRect 10 6 4 3 $b), (New-TqqRect 10 10 3 3 $b)) }
        tqq_scatter   = @{ Mso = 'ChartTypeXYScatterInsertGallery'; Rects = @((New-TqqRect 2 10 3 3 $b), (New-TqqRect 6 7 3 3 $b), (New-TqqRect 9 11 3 3 $n), (New-TqqRect 11 3 3 3 $b), (New-TqqRect 5 2 3 3 $n), (New-TqqRect 0 0 1 16 $g), (New-TqqRect 0 15 16 1 $g)) }
        tqq_bubble    = @{ Mso = 'ChartTypeXYScatterInsertGallery'; Rects = @((New-TqqRect 1 8 7 7 $b), (New-TqqRect 9 2 6 6 $n), (New-TqqRect 7 1 3 3 $l), (New-TqqRect 0 0 1 16 $g), (New-TqqRect 0 15 16 1 $g)) }
        tqq_options   = @{ Mso = 'PropertySheet';                Rects = @((New-TqqRect 1 3 14 1 $g), (New-TqqRect 1 8 14 1 $g), (New-TqqRect 1 13 14 1 $g), (New-TqqRect 9 2 3 3 $n), (New-TqqRect 4 7 3 3 $b), (New-TqqRect 11 12 3 3 $n)) }
        tqq_arrows    = @{ Mso = 'ArrowsMore';                   Rects = @((New-TqqRect 1 7 10 2 $n), (New-TqqRect 9 4 2 8 $n), (New-TqqRect 11 5 2 6 $n), (New-TqqRect 13 6 2 4 $n)) }
        tqq_ppt       = @{ Mso = 'SlideNew';                     Rects = @((New-TqqRect 1 2 14 11 $n), (New-TqqRect 3 4 10 7 $l), (New-TqqRect 4 7 2 3 $b), (New-TqqRect 7 5 2 5 $b), (New-TqqRect 10 6 2 4 $b), (New-TqqRect 6 13 4 2 $g)) }
        tqq_round     = @{ Mso = 'DecimalsDecrease';             Rects = @((New-TqqRect 1 2 14 4 $l), (New-TqqRect 1 7 14 4 $b), (New-TqqRect 1 12 14 3 $n), (New-TqqRect 14 1 2 14 $g)) }
        tqq_break     = @{ Mso = 'ChartAxes';                    Rects = @((New-TqqRect 5 9 6 6 $b), (New-TqqRect 5 1 6 5 $b), (New-TqqRect 3 6 10 1 $g), (New-TqqRect 3 8 10 1 $g)) }
    }
}

# tabela do CRC-32 montada num bloco de script: $poly, $i, $c e $k nao vazam para os scripts que fazem dot-source do Build-Common
$script:TqqCrcTable = & {
    $t = New-Object uint64[] 256
    $poly = [uint64]3988292384
    for ($i = 0; $i -lt 256; $i++) { $c = [uint64]$i; for ($k = 0; $k -lt 8; $k++) { if (($c -band 1) -eq 1) { $c = $poly -bxor ($c -shr 1) } else { $c = $c -shr 1 } }; $t[$i] = $c }
    , $t
}
function Get-TqqCrc32([byte[]]$Bytes) {
    $c = [uint64]4294967295
    foreach ($b in $Bytes) { $c = $script:TqqCrcTable[[int](($c -bxor [uint64]$b) -band 255)] -bxor ($c -shr 8) }
    return [uint32]($c -bxor [uint64]4294967295)
}
function Get-TqqBE32([uint32]$v) { return [byte[]]@([byte](($v -shr 24) -band 255), [byte](($v -shr 16) -band 255), [byte](($v -shr 8) -band 255), [byte]($v -band 255)) }
function New-TqqPngChunk([string]$Type, [byte[]]$Data) {
    $t = [Text.Encoding]::ASCII.GetBytes($Type)
    return [byte[]]((Get-TqqBE32 ([uint32]$Data.Length)) + $t + $Data + (Get-TqqBE32 (Get-TqqCrc32 ([byte[]]($t + $Data)))))
}
function Get-TqqZlibStored([byte[]]$Raw) {
    $out = New-Object System.Collections.Generic.List[byte]
    $out.Add([byte]0x78); $out.Add([byte]0x01)
    $pos = 0
    do {
        $len = [Math]::Min(65535, $Raw.Length - $pos); $final = if ($pos + $len -ge $Raw.Length) { 1 } else { 0 }; $nlen = 65535 - $len
        $out.Add([byte]$final); $out.Add([byte]($len -band 255)); $out.Add([byte](($len -shr 8) -band 255)); $out.Add([byte]($nlen -band 255)); $out.Add([byte](($nlen -shr 8) -band 255))
        for ($i = 0; $i -lt $len; $i++) { $out.Add($Raw[$pos + $i]) }
        $pos += $len
    } while ($pos -lt $Raw.Length)
    [uint64]$a = 1; [uint64]$b = 0
    foreach ($x in $Raw) { $a = ($a + [uint64]$x) % 65521; $b = ($b + $a) % 65521 }
    foreach ($y in (Get-TqqBE32 ([uint32](($b -shl 16) -bor $a)))) { $out.Add([byte]$y) }
    return $out.ToArray()
}
function Write-TqqPng([string]$Path, [int]$Width, [int]$Height, [byte[]]$Rgba) {
    $raw = New-Object byte[] (($Width * 4 + 1) * $Height)
    for ($y = 0; $y -lt $Height; $y++) { $raw[$y * ($Width * 4 + 1)] = 0; [Array]::Copy($Rgba, $y * $Width * 4, $raw, $y * ($Width * 4 + 1) + 1, $Width * 4) }
    $ihdr = [byte[]]((Get-TqqBE32 ([uint32]$Width)) + (Get-TqqBE32 ([uint32]$Height)) + [byte[]]@(8, 6, 0, 0, 0))
    $png = [byte[]]([byte[]]@(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A) + (New-TqqPngChunk 'IHDR' $ihdr) + (New-TqqPngChunk 'IDAT' (Get-TqqZlibStored $raw)) + (New-TqqPngChunk 'IEND' ([byte[]]@())))
    [IO.File]::WriteAllBytes($Path, $png)
}
function New-TqqIconRgba([int]$Size, $Rects) {
    $px = New-Object byte[] ($Size * $Size * 4)
    foreach ($r in $Rects) {
        $x0 = [int][Math]::Floor($r.x * $Size / 16); $y0 = [int][Math]::Floor($r.y * $Size / 16)
        $x1 = [Math]::Min($Size, [int][Math]::Ceiling(($r.x + $r.w) * $Size / 16)); $y1 = [Math]::Min($Size, [int][Math]::Ceiling(($r.y + $r.h) * $Size / 16))
        for ($y = $y0; $y -lt $y1; $y++) { for ($x = $x0; $x -lt $x1; $x++) { $o = ($y * $Size + $x) * 4; $px[$o] = [byte]$r.c[0]; $px[$o + 1] = [byte]$r.c[1]; $px[$o + 2] = [byte]$r.c[2]; $px[$o + 3] = 255 } }
    }
    return $px
}
# Grava <nome>.png (32) e <nome>16.png (16) de cada icone em OutDir; devolve os caminhos (ordem fixa).
function Export-TqqIcons([string]$OutDir) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($e in (Get-TqqIconSpecs).GetEnumerator()) {
        $p32 = Join-Path $OutDir "$($e.Key).png"; Write-TqqPng $p32 32 32 (New-TqqIconRgba 32 $e.Value.Rects); $out.Add($p32)
        $p16 = Join-Path $OutDir "$($e.Key)16.png"; Write-TqqPng $p16 16 16 (New-TqqIconRgba 16 $e.Value.Rects); $out.Add($p16)
    }
    return $out.ToArray()
}
# Fallback: troca image="tqq_x" / image="tqq_x16" por imageMso="<reserva>" (so no texto que vai para o pacote; o fonte nao muda).
function Convert-TqqRibbonToMso([string]$Xml) {
    $specs = Get-TqqIconSpecs
    return [regex]::Replace($Xml, 'image="(tqq_[a-z0-9_]+?)(16)?"', { param($m) $k = $m.Groups[1].Value; if ($specs.Contains($k)) { 'imageMso="' + $specs[$k].Mso + '"' } else { throw "icone sem reserva: $k" } })
}
```

- [ ] **Passo 4: `Add-RibbonToPackage` e `build.ps1`** — em `Build-Common.ps1`, logo após o `$script:DistDir`, acrescente `. "$PSScriptRoot\Ribbon-Icons.ps1"` e substitua a função por:

```powershell
# Insere customUI/customUI14.xml no pacote (zip), a relacao em _rels/.rels e, no modo Png, os icones proprios: customUI/images/<nome>.png
# para cada image= do XML, customUI/_rels/customUI14.xml.rels (Id = nome) e o content type png. Modo Mso: XML com imageMso de reserva, sem imagens.
function Add-RibbonToPackage([string]$PackagePath, [string]$RibbonXmlPath, [ValidateSet('Png','Mso')][string]$Icons = 'Png') {
    [void]([xml](Get-Content $RibbonXmlPath -Raw))   # falha cedo se o XML estiver malformado
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $xmlText = [IO.File]::ReadAllText($RibbonXmlPath, (New-Object Text.UTF8Encoding($false)))
    $images = @([regex]::Matches($xmlText, 'image="([^"]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    if ($Icons -eq 'Mso') { $xmlText = Convert-TqqRibbonToMso -Xml $xmlText; $images = @() }
    $specs = Get-TqqIconSpecs
    foreach ($img in $images) { if (-not $specs.Contains(($img -replace '16$', ''))) { throw "customUI14.xml usa image=`"$img`" sem icone na tabela (Ribbon-Icons.ps1)" } }
    $iconDir = $null
    if ($images.Count -gt 0) { $iconDir = Join-Path ([IO.Path]::GetTempPath()) ("thinqq_icons_" + [guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Force $iconDir | Out-Null; [void](Export-TqqIcons -OutDir $iconDir) }
    $zip = [System.IO.Compression.ZipFile]::Open($PackagePath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        function Set-ZipText($zip, [string]$name, [string]$text) {
            $old = $zip.GetEntry($name); if ($old) { $old.Delete() }
            $e = $zip.CreateEntry($name); $w = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false))); $w.Write($text); $w.Dispose()
        }
        if ($Icons -eq 'Mso') {
            Set-ZipText $zip 'customUI/customUI14.xml' $xmlText
        } else {
            $existing = $zip.GetEntry('customUI/customUI14.xml'); if ($existing) { $existing.Delete() }
            $entry = $zip.CreateEntry('customUI/customUI14.xml'); $bytes = [IO.File]::ReadAllBytes($RibbonXmlPath); $s = $entry.Open(); $s.Write($bytes, 0, $bytes.Length); $s.Dispose()
        }
        $relsEntry = $zip.GetEntry('_rels/.rels')
        $reader = New-Object System.IO.StreamReader($relsEntry.Open()); $rels = $reader.ReadToEnd(); $reader.Dispose()
        if ($rels -notmatch 'customUI/customUI14\.xml') {
            $rel = '<Relationship Id="rIdThinQQ" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
            Set-ZipText $zip '_rels/.rels' ($rels -replace '</Relationships>', "$rel</Relationships>")
        }
        foreach ($old in @($zip.Entries | Where-Object { $_.FullName -like 'customUI/images/*' -or $_.FullName -eq 'customUI/_rels/customUI14.xml.rels' })) { $old.Delete() }
        if ($images.Count -gt 0) {
            foreach ($img in $images) {
                $e = $zip.CreateEntry("customUI/images/$img.png"); $b = [IO.File]::ReadAllBytes((Join-Path $iconDir "$img.png")); $s = $e.Open(); $s.Write($b, 0, $b.Length); $s.Dispose()
            }
            $imgRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                (($images | ForEach-Object { '<Relationship Id="' + $_ + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/' + $_ + '.png"/>' }) -join '') + '</Relationships>'
            Set-ZipText $zip 'customUI/_rels/customUI14.xml.rels' $imgRels
            $ctEntry = $zip.GetEntry('[Content_Types].xml')
            $reader = New-Object System.IO.StreamReader($ctEntry.Open()); $ct = $reader.ReadToEnd(); $reader.Dispose()
            if ($ct -notmatch 'Extension="png"') { Set-ZipText $zip '[Content_Types].xml' ($ct -replace '<Default Extension="rels"', '<Default Extension="png" ContentType="image/png"/><Default Extension="rels"') }
        }
    } finally { $zip.Dispose(); if ($iconDir) { Remove-Item $iconDir -Recurse -Force -ErrorAction SilentlyContinue } }
}
```
(Nota: `[Content_Types].xml` de um `.xlam` salvo pelo Excel sempre tem `<Default Extension="rels" .../>`; o teste do pacote confere o resultado.) Em `build.ps1`: parâmetro `[ValidateSet('Png','Mso')] [string]$RibbonIcons = 'Png'` e `Add-RibbonToPackage -PackagePath $outPath -RibbonXmlPath (...) -Icons $RibbonIcons`.

- [ ] **Passo 5: GREEN** — `pwsh C:\temp\tqq-f3\build\run-tests.ps1 -Target loader -TimeoutSec 400` verde (o XML atual não tem `image=`: o pacote continua idêntico ao fonte e sem imagens; o gravador é testado em pasta temporária). Depois `pwsh C:\temp\tqq-f3\build\build.ps1 -Target loader -Configuration Test -RibbonIcons Mso` e `pwsh C:\temp\tqq-f3\tests\ps\Test-Ribbon.ps1 -RepoRoot C:\temp\tqq-f3` → verde também; refaça o build `Png` ao final. Suíte completa verde.

- [ ] **Passo 6: Commit** — `git add build/lib/Ribbon-Icons.ps1 build/lib/Build-Common.ps1 build/build.ps1 tests/ps/Test-Ribbon.ps1 tests/vba/loader/TestRibbon.bas` e `git commit -m "build: icones proprios da faixa gerados no build (PNG puro, imagens e relacoes no pacote) com reserva imageMso"` (+ trailer).

**Verificação no Excel real:** nenhuma mudança visível até a Tarefa 8 (o XML ainda usa `imageMso`).

---
### Tarefa 7: Certificado, assinatura e política corporativa (scripts e verificação) — modelo: standard

**Files:**
- Create: `tools/New-ThinQQCert.ps1`, `tools/Trust-ThinQQPublisher.ps1`, `tools/Sign-ThinQQ.ps1`, `tests/ps/Test-Signing.ps1`
- Modify: `build/release.ps1` (`-SignPfx`, `-SignPassword`, `-SignTimestampUrl`), `build/run-tests.ps1` (chama `Test-Signing.ps1` num processo próprio, depois de `Test-Privacy.ps1`)
- Test: `tests/ps/Test-Signing.ps1` (via `run-tests.ps1`)

**Interfaces:**
- Produces (`tools/New-ThinQQCert.ps1`): parâmetros `-Subject` (padrão `CN=ThinQQ (autoassinado)`), `-OutDir` (padrão `dist\cert`), `-Years` (3), `-PfxPassword` (gerada e impressa se vazia), `-TrustRoot` (importa também em `Cert:\CurrentUser\Root`, avisando que o Windows pede confirmação), `-NoTrustPublisher`. Cria o certificado (`New-SelfSignedCertificate -Type CodeSigningCert`, RSA 2048, SHA-256, exportável) em `Cert:\CurrentUser\My` e remove a cópia que o cmdlet grava em `Cert:\CurrentUser\CA` (Autoridades de Certificação Intermediárias, inútil aqui), exporta `ThinQQ-publisher.cer` e `ThinQQ-signing.pfx`, importa o `.cer` em `Cert:\CurrentUser\TrustedPublisher` (sem diálogo: medido) e devolve `Thumbprint, Subject, Cer, Pfx, NotAfter`.
- Produces (`tools/Trust-ThinQQPublisher.ps1`): `-CerPath` (obrigatório), `-TrustRoot`, `-Scope CurrentUser|LocalMachine` (`LocalMachine` exige administrador). Importa em `TrustedPublisher` (e `Root`), idempotente; devolve o thumbprint.
- Produces (`tools/Sign-ThinQQ.ps1`): `-Files` (padrão `dist\ThinQQ.xlam`, `dist\ThinQQ.core.xlam`), `-PfxPath`/`-PfxPassword` ou `-Thumbprint`, `-TimestampUrl`, `-WhatIf`. Procura o `signtool.exe` (PATH; `Windows Kits\10\bin\*\x86` de preferência, depois `x64`) e a SIP do Office (`HKLM:\SOFTWARE\Microsoft\Cryptography\OID\EncodingType 0\CryptSIPDllGetSignedDataMsg\*` nas visões de 64 bits e `WOW6432Node`, com `Dll` contendo `msosip`); imprime o diagnóstico; sem os dois pré-requisitos sai com **código 2** e as instruções (SDK + pacote "Office Subject Interface Packages for Digitally Signing VBA Projects" + `regsvr32` como administrador); com eles, `signtool sign /fd SHA256 /f <pfx> /p <senha> [/tr <url> /td SHA256] <arquivo>` (ou `/sha1 <thumbprint>`) e `signtool verify /pa <arquivo>`; `-WhatIf` só diagnostica (código 0 com tudo presente, 2 sem).
- Produces (`build/release.ps1`): `-SignPfx <caminho>`, `-SignPassword`, `-SignTimestampUrl`: depois do build Release e da guarda de módulos de teste e **antes** de calcular os SHA-256 do manifesto, chama `tools/Sign-ThinQQ.ps1`; código ≠ 0 => o release falha com a mensagem do script (nada é publicado com assinatura incompleta).
- Produces (`tests/ps/Test-Signing.ps1`): **só pwsh** (o `run-tests` o chama uma vez, num processo próprio do mesmo `pwsh`; termina com `exit 0` ou `exit 1`): cria um certificado com assunto único `CN=ThinQQ Teste <guid>` numa pasta temporária, confere `.cer`/`.pfx` (abre o PFX com a senha e vê a chave privada), confere a presença em `TrustedPublisher`, roda `Trust-ThinQQPublisher.ps1` (idempotente), roda `Sign-ThinQQ.ps1 -WhatIf` e aceita código 0 (ferramentas presentes: então assina uma **cópia** de `dist\ThinQQ.xlam` em pasta temporária e exige código 0) ou 2 (ausentes: exige a palavra `signtool` ou `SIP` na saída); confere que `dist\ThinQQ.xlam` não muda; nunca toca `Root` nem grava em `HKLM`; no `finally`, mesmo em falha (inclusive se `New-ThinQQCert` falhar depois de criar o certificado), remove pelo assunto o certificado das três lojas em que ele pode estar (`My` com `-DeleteKey`, `CA` e `TrustedPublisher`), confere que nada com o assunto sobrou em `Cert:\CurrentUser` e apaga a pasta.

- [ ] **Passo 1: Teste (RED)** — `tests/ps/Test-Signing.ps1`:

```powershell
# Certificado autoassinado por usuario (sem administrador): criar, exportar, confiar como editor, diagnosticar/assinar e REMOVER. Nunca toca
# Cert:\...\Root nem grava em HKLM. Assunto unico por execucao; a limpeza do finally remove pelo ASSUNTO, nas tres lojas em que o certificado
# pode estar (My; CA, onde New-SelfSignedCertificate grava uma copia em "Autoridades de Certificacao Intermediarias"; TrustedPublisher),
# inclusive quando a criacao falha no meio, e confere que nada sobrou. O run-tests o chama num processo proprio: termina sempre com exit 0 ou 1.
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$tmp = Join-Path $env:TEMP ("tqq_sign_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $tmp | Out-Null
function Assert($cond, $msg) { if (-not $cond) { throw "FALHA: $msg" } }
function Get-DistHash([string]$p) { if (Test-Path -LiteralPath $p) { return (Get-FileHash -LiteralPath $p).Hash } else { return '' } }
$subject = 'CN=ThinQQ Teste ' + [guid]::NewGuid().ToString('N').Substring(0, 8)
$exitCode = 1
$dist = Join-Path $repo 'dist\ThinQQ.xlam'
$distHash = Get-DistHash $dist
try {
    $r = & (Join-Path $repo 'tools\New-ThinQQCert.ps1') -Subject $subject -OutDir $tmp -Years 1 -PfxPassword 'teste'
    $thumb = $r.Thumbprint
    Assert ($thumb -match '^[0-9A-F]{40}$') "thumbprint: $thumb"
    Assert ((Get-Item "Cert:\CurrentUser\My\$thumb").EnhancedKeyUsageList.ObjectId -contains '1.3.6.1.5.5.7.3.3') 'EKU de assinatura de codigo'
    Assert (-not (Test-Path "Cert:\CurrentUser\CA\$thumb")) 'nenhuma copia em Autoridades Intermediarias (New-ThinQQCert a remove)'
    Assert (Test-Path $r.Cer) '.cer exportado'; Assert (Test-Path $r.Pfx) '.pfx exportado'
    # EphemeralKeySet: abrir o PFX nao grava a chave privada num conteiner do perfil; Dispose ao final
    $pfx = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($r.Pfx, 'teste', [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet)
    try { Assert ($pfx.HasPrivateKey -and $pfx.Thumbprint -eq $thumb) 'pfx com a chave privada' } finally { $pfx.Dispose() }
    Assert (Test-Path "Cert:\CurrentUser\TrustedPublisher\$thumb") 'editor confiavel do usuario'
    $t2 = & (Join-Path $repo 'tools\Trust-ThinQQPublisher.ps1') -CerPath $r.Cer
    Assert ($t2 -eq $thumb) 'Trust-ThinQQPublisher idempotente'
    Assert (-not (Test-Path "Cert:\CurrentUser\Root\$thumb")) 'Root nao foi tocada'
    # diagnostico / assinatura: Sign-ThinQQ so escreve com Write-Host (fluxo de informacao): *>&1 o captura; 2>&1 nao captura nada
    $copy = Join-Path $tmp 'ThinQQ.xlam'
    if (Test-Path $dist) { Copy-Item $dist $copy } else { Set-Content $copy 'x' }
    $global:LASTEXITCODE = 0
    $out = & (Join-Path $repo 'tools\Sign-ThinQQ.ps1') -Files $copy -PfxPath $r.Pfx -PfxPassword 'teste' -WhatIf *>&1 | Out-String
    $code = $global:LASTEXITCODE
    Assert ($code -eq 0 -or $code -eq 2) "Sign-ThinQQ -WhatIf: codigo $code"
    Assert ($out -match 'signtool' -or $out -match 'SIP') "diagnostico cita signtool/SIP: $out"
    if ($code -eq 0 -and (Test-Path $dist)) {
        $global:LASTEXITCODE = 0
        & (Join-Path $repo 'tools\Sign-ThinQQ.ps1') -Files $copy -PfxPath $r.Pfx -PfxPassword 'teste' *>&1 | Out-Null
        Assert ($global:LASTEXITCODE -eq 0) 'assinatura real da copia (ferramentas presentes)'
        Write-Host 'Test-Signing: assinatura real executada numa copia'
    } else { Write-Host "Test-Signing: ferramentas de assinatura ausentes (codigo $code): so o diagnostico foi verificado" }
    Assert ((Get-DistHash $dist) -eq $distHash) 'dist\ThinQQ.xlam nao mudou (a assinatura real e feita numa copia)'
    Write-Host "Test-Signing: OK (PowerShell $($PSVersionTable.PSVersion))" -ForegroundColor Green
    $exitCode = 0
} catch {
    Write-Host "not ok - $($_.Exception.Message)" -ForegroundColor Red
} finally {
    foreach ($store in 'My', 'CA', 'TrustedPublisher') {
        Get-ChildItem "Cert:\CurrentUser\$store" -ErrorAction SilentlyContinue | Where-Object Subject -eq $subject | ForEach-Object {
            if ($store -eq 'My') { Remove-Item $_.PSPath -DeleteKey -ErrorAction SilentlyContinue } else { Remove-Item $_.PSPath -ErrorAction SilentlyContinue } }
    }
    $left = @(Get-ChildItem Cert:\CurrentUser -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.PSObject.Properties['Subject'] -and $_.Subject -eq $subject })
    if ($left.Count -gt 0) { Write-Host "not ok - certificado de teste sobrou em: $($left.PSParentPath -join ', ')" -ForegroundColor Red; $exitCode = 1 }
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
exit $exitCode
```

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-f3\tests\ps\Test-Signing.ps1` → falha (scripts inexistentes).

- [ ] **Passo 3: Implementar os scripts** — `tools/New-ThinQQCert.ps1`:

```powershell
<#
.SYNOPSIS  Cria um certificado autoassinado de assinatura de codigo para o ThinQQ (por usuario, sem administrador), exporta .cer e .pfx
           e instala o .cer como Editor Confiavel do usuario. Com -TrustRoot tambem o instala em Autoridades Raiz Confiaveis do usuario
           (o Windows pede confirmacao numa janela). Sem a raiz, o Office mostra o editor como "nao verificado".
.EXAMPLE   pwsh tools/New-ThinQQCert.ps1 -Subject "CN=Minha Empresa - ThinQQ" -TrustRoot
#>
param(
    [string]$Subject = 'CN=ThinQQ (autoassinado)',
    [string]$OutDir = (Join-Path $PSScriptRoot '..\dist\cert'),
    [int]$Years = 3,
    [string]$PfxPassword = '',
    [switch]$TrustRoot,
    [switch]$NoTrustPublisher
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutDir | Out-Null
$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $Subject -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
    -KeyExportPolicy Exportable -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears($Years) -FriendlyName 'ThinQQ assinatura de codigo'
# o cmdlet tambem grava uma copia em "Autoridades de Certificacao Intermediarias" (Cert:\CurrentUser\CA); ela nao serve para nada aqui
Get-ChildItem Cert:\CurrentUser\CA | Where-Object Thumbprint -eq $cert.Thumbprint | Remove-Item
$cer = Join-Path $OutDir 'ThinQQ-publisher.cer'
$pfx = Join-Path $OutDir 'ThinQQ-signing.pfx'
Export-Certificate -Cert $cert -FilePath $cer -Type CERT | Out-Null
if (-not $PfxPassword) {
    $PfxPassword = -join ((48..57 + 65..90 + 97..122) | Get-Random -Count 16 | ForEach-Object { [char]$_ })
    Write-Host "Senha do PFX gerada (guarde-a; nao e gravada em lugar nenhum): $PfxPassword"
}
Export-PfxCertificate -Cert $cert -FilePath $pfx -Password (ConvertTo-SecureString $PfxPassword -AsPlainText -Force) | Out-Null
if (-not $NoTrustPublisher) { Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null }
if ($TrustRoot) {
    Write-Host 'Instalando o certificado em Autoridades Raiz Confiaveis do usuario: o Windows vai pedir confirmacao.' -ForegroundColor Yellow
    Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
}
Write-Host "Certificado $($cert.Thumbprint) ($Subject) criado; .cer e .pfx em $OutDir"
[pscustomobject]@{ Thumbprint = $cert.Thumbprint; Subject = $cert.Subject; Cer = $cer; Pfx = $pfx; NotAfter = $cert.NotAfter }
```

`tools/Trust-ThinQQPublisher.ps1`:

```powershell
<#
.SYNOPSIS  Instala um certificado de editor (.cer) como Editor Confiavel (e, com -TrustRoot, como raiz confiavel) do usuario atual ou da maquina.
.EXAMPLE   pwsh tools/Trust-ThinQQPublisher.ps1 -CerPath .\ThinQQ-publisher.cer -TrustRoot
.NOTES     -Scope LocalMachine exige administrador. Em empresas, prefira distribuir o .cer por GPO (Computer Configuration > Policies >
           Windows Settings > Security Settings > Public Key Policies > Trusted Publishers).
#>
param(
    [Parameter(Mandatory)][string]$CerPath,
    [switch]$TrustRoot,
    [ValidateSet('CurrentUser', 'LocalMachine')][string]$Scope = 'CurrentUser'
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $CerPath)) { throw "Certificado nao encontrado: $CerPath" }
$c = Import-Certificate -FilePath $CerPath -CertStoreLocation "Cert:\$Scope\TrustedPublisher"
if ($TrustRoot) {
    Write-Host 'Instalando em Autoridades Raiz Confiaveis: o Windows pode pedir confirmacao.' -ForegroundColor Yellow
    Import-Certificate -FilePath $CerPath -CertStoreLocation "Cert:\$Scope\Root" | Out-Null
}
Write-Host "Editor confiavel ($Scope): $($c.Subject) [$($c.Thumbprint)]"
$c.Thumbprint
```

`tools/Sign-ThinQQ.ps1`:

```powershell
<#
.SYNOPSIS  Assina os projetos VBA dos .xlam por linha de comando (SignTool do Windows SDK + Office Subject Interface Packages) ou so diagnostica (-WhatIf).
.EXAMPLE   pwsh tools/Sign-ThinQQ.ps1 -PfxPath .\dist\cert\ThinQQ-signing.pfx -PfxPassword ... -TimestampUrl http://timestamp.digicert.com
.NOTES     Pre-requisitos (uma vez, com administrador): Windows SDK (signtool.exe) e o pacote "Office Subject Interface Packages for Digitally
           Signing VBA Projects" (Microsoft), com as DLLs registradas por regsvr32 (leia o readme.txt do pacote). Sem eles este script sai com
           codigo 2 e nada e assinado. O VBE (Ferramentas > Assinatura Digital) continua sendo o caminho manual.
#>
param(
    [string[]]$Files = @((Join-Path $PSScriptRoot '..\dist\ThinQQ.xlam'), (Join-Path $PSScriptRoot '..\dist\ThinQQ.core.xlam')),
    [string]$PfxPath = '',
    [string]$PfxPassword = '',
    [string]$Thumbprint = '',
    [string]$TimestampUrl = '',
    [switch]$WhatIf
)
$ErrorActionPreference = 'Stop'
function Find-SignTool {
    $c = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    foreach ($arch in 'x86', 'x64') {
        $cands = Get-ChildItem -Path (Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin') -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue | Where-Object { $_.FullName -like "*\$arch\signtool.exe" } | Sort-Object FullName -Descending
        if ($cands) { return $cands[0].FullName }
    }
    return $null
}
# A SIP do Office fica na visao de 64 bits ou, registrada pelo regsvr32 de 32 bits (a do SignTool x86, que este script prefere), na
# WOW6432Node: procura nas duas. PSObject.Properties evita o erro de modo estrito numa subchave sem o valor Dll.
function Test-OfficeSip {
    foreach ($root in 'HKLM:\SOFTWARE\Microsoft\Cryptography\OID\EncodingType 0\CryptSIPDllGetSignedDataMsg', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Cryptography\OID\EncodingType 0\CryptSIPDllGetSignedDataMsg') {
        if (-not (Test-Path $root)) { continue }
        foreach ($k in (Get-ChildItem $root -ErrorAction SilentlyContinue)) {
            $p = Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue
            if ($p -and $p.PSObject.Properties['Dll'] -and $p.Dll -match 'msosip') { return $true }
        }
    }
    return $false
}
$signtool = Find-SignTool
$sip = Test-OfficeSip
Write-Host ("signtool.exe: " + $(if ($signtool) { $signtool } else { 'NAO ENCONTRADO (instale o Windows SDK)' }))
Write-Host ("SIP do Office (msosip/msosipx registradas): " + $(if ($sip) { 'presente' } else { 'AUSENTE (instale o pacote Office SIPs e registre as DLLs com regsvr32, como administrador)' }))
if (-not $signtool -or -not $sip) {
    Write-Host 'Pre-requisitos ausentes: nada foi assinado. Caminho manual: VBE > Ferramentas > Assinatura Digital (gera assinatura V3).' -ForegroundColor Yellow
    exit 2
}
if ($WhatIf) { Write-Host 'WhatIf: pre-requisitos presentes; a assinatura rodaria agora.'; exit 0 }
if (-not $PfxPath -and -not $Thumbprint) { throw 'Informe -PfxPath (e -PfxPassword) ou -Thumbprint.' }
foreach ($f in $Files) {
    if (-not (Test-Path -LiteralPath $f)) { throw "Arquivo nao encontrado: $f" }
    $stArgs = @('sign', '/fd', 'SHA256')                 # nunca $args: e a variavel automatica dos argumentos do script
    if ($PfxPath) { $stArgs += @('/f', $PfxPath); if ($PfxPassword) { $stArgs += @('/p', $PfxPassword) } } else { $stArgs += @('/sha1', $Thumbprint) }
    if ($TimestampUrl) { $stArgs += @('/tr', $TimestampUrl, '/td', 'SHA256') }
    $stArgs += $f
    & $signtool @stArgs
    if ($LASTEXITCODE -ne 0) { throw "signtool sign falhou ($LASTEXITCODE) em $f" }
    & $signtool verify /pa $f
    if ($LASTEXITCODE -ne 0) { throw "signtool verify falhou ($LASTEXITCODE) em $f" }
    Write-Host "Assinado e verificado: $f"
}
exit 0
```

- [ ] **Passo 4: `release.ps1` e `run-tests.ps1`** — em `release.ps1`, parâmetros `[string]$SignPfx = ''`, `[string]$SignPassword = ''`, `[string]$SignTimestampUrl = ''`; depois do laço que confere os artefatos (`Get-TestMarkersInXlam`) e antes de `$notes = ''`:

```powershell
    if ($SignPfx) {
        $global:LASTEXITCODE = 0
        & "$PSScriptRoot\..\tools\Sign-ThinQQ.ps1" -Files @((Join-Path $DistDir 'ThinQQ.xlam'), (Join-Path $DistDir 'ThinQQ.core.xlam')) -PfxPath $SignPfx -PfxPassword $SignPassword -TimestampUrl $SignTimestampUrl
        if ($global:LASTEXITCODE -ne 0) { throw "Assinatura dos .xlam falhou (codigo $global:LASTEXITCODE): o release nao foi gerado. Veja tools/Sign-ThinQQ.ps1." }
    }
```
(os SHA-256 do manifesto são calculados depois, sobre os arquivos assinados). Em `run-tests.ps1`, depois do bloco `=== privacidade ===`:

```powershell
if (-not $timedOut) {
    Write-Host '=== assinatura (certificado por usuario: cria, confia, diagnostica e remove) ==='
    # processo proprio (como os testes do instalador): o exit 2 do Sign-ThinQQ e uma excecao do teste nao vazam para o run-tests
    $global:LASTEXITCODE = 0
    & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $RepoRoot 'tests\ps\Test-Signing.ps1')
    if ($LASTEXITCODE -ne 0) { $totalFailed += 1 }
}
```

- [ ] **Passo 5: GREEN e suíte completa.** Registre no relatório o diagnóstico impresso por `Sign-ThinQQ.ps1 -WhatIf` nesta máquina (esperado: signtool e SIP ausentes, código 2) e confirme com `Get-ChildItem Cert:\CurrentUser\My, Cert:\CurrentUser\CA, Cert:\CurrentUser\TrustedPublisher, Cert:\CurrentUser\Root | Where-Object Subject -like '*ThinQQ Teste*'` que nada sobrou (vazio; um certificado com outro assunto ThinQQ não é do teste: registre no relatório e não o remova).

- [ ] **Passo 6: Commit** — `git add tools/New-ThinQQCert.ps1 tools/Trust-ThinQQPublisher.ps1 tools/Sign-ThinQQ.ps1 tests/ps/Test-Signing.ps1 build/release.ps1 build/run-tests.ps1` e `git commit -m "build: scripts de certificado autoassinado, editor confiavel e assinatura (signtool + Office SIPs), gancho -SignPfx no release"` (+ trailer).

**Verificação no Excel real (opcional, usuário):** `pwsh tools/New-ThinQQCert.ps1 -TrustRoot` (confirmar o diálogo), assinar `dist\ThinQQ.core.xlam` pelo VBE (Ferramentas > Assinatura Digital) e abrir: sem aviso de macro.

---

### Tarefa 8: Faixa de opções 1.2.0 (menus, ícones próprios, controles novos) e rótulos no tutorial — modelo: standard

**Files:**
- Modify: `src/loader/ribbon/customUI14.xml`, `src/loader/modules/modRibbon.bas`, `tests/vba/loader/TestRibbon.bas`, `docs/tutorial/tutorial.html` (maquete, `#tipos`, tabela de controles, classes `k` → `ui` da matriz)
- Test: `tests/vba/loader/TestRibbon.bas`, `tests/ps/Test-Ribbon.ps1` (já aperta as imagens), `tests/ps/Test-Docs.ps1` (rótulos)

**Interfaces:**
- Consumes (core 1.2.0): `Core_Insert` dos 20 tipos, `Core_SetToggle` (`roundsum`, `trend`, `quad`), `Core_SetAxisBreak`, `Core_RemoveAxisBreak`, `Core_RoundRange`, `Core_CopyToPowerPoint(mode)`, `Core_Supports` (`break`, `roundsum`, `trend`, `quad`). Ícones da Tarefa 6 (`tqq_*` e `tqq_*16`).
- Produces (XML; ids e rótulos exatos):

```
grpInsert: mnuColumns image="tqq_columns" | mnuBars image="tqq_bars" | mnuWaterfall image="tqq_waterfall" | mnuLines image="tqq_lines" | mnuOthers image="tqq_more"
  (os cinco trocam imageMso por image; os 17 botões existentes ficam iguais)
  mnuOthers ganha, depois de btnDoughnut:
    btnButterfly "Borboleta" tag=butterfly image="tqq_butterfly16" supertip="Duas séries espelhadas (esquerda e direita) com os nomes das categorias no centro. Tabela: categorias e exatamente 2 colunas de valores."
    btnScatter "Dispersão" tag=scatter image="tqq_scatter16" supertip="Pontos em dois eixos de valor, com o nome ao lado de cada ponto. Tabela: nome, X, Y e, opcionalmente, uma coluna de grupo (uma série por grupo)."
    btnBubble "Bolhas" tag=bubble image="tqq_bubble16" supertip="Como a dispersão, com o tamanho da bolha numa terceira coluna. Tabela: nome, X, Y, tamanho e, opcionalmente, grupo."
grpChart: (toggles e dropdowns iguais)
  mnuMore image="tqq_options16" (em vez de imageMso):
    tglInd100, tglVline, tglSecondary, tglPieOut (iguais)
    tglRoundSum "Rótulos com soma consistente" tag=roundsum image="tqq_round16" onAction=Rx_OnToggle getPressed=Rx_GetPressed getEnabled=Rx_GetEnabled supertip="Arredonda os rótulos de segmento para que somem exatamente o total mostrado (34% + 33% + 33% = 100%), pelo método do maior resto, em fórmulas da planilha."
    tglTrend "Linha de tendência" tag=trend image="tqq_lines16" onAction=Rx_OnToggle getPressed=Rx_GetPressed getEnabled=Rx_GetEnabled supertip="Regressão linear por série (dispersão e bolhas)."
    tglQuad "Quadrantes (médias)" tag=quad image="tqq_scatter16" onAction=Rx_OnToggle getPressed=Rx_GetPressed getEnabled=Rx_GetEnabled supertip="Os eixos passam a cruzar nas médias de X e de Y, dividindo o gráfico em quatro quadrantes (dispersão e bolhas)."
    menuSeparator msepBreak
    btnBreak "Corte de eixo..." tag=break image="tqq_break16" onAction=Rx_OnAxisBreak getEnabled=Rx_GetEnabled supertip="Corta o eixo de valores num valor que você informa: as colunas maiores ficam cortadas, com uma marca, e o total continua no rótulo."
    btnBreakOff "Remover corte de eixo" tag=break onAction=Rx_OnAxisBreakOff getEnabled=Rx_GetEnabled
    menuSeparator msepRound
    btnRoundRange "Arredondar intervalo..." image="tqq_round16" onAction=Rx_OnRoundRange (sem getEnabled: age sobre células) supertip="Grava, a partir de uma célula que você escolhe, fórmulas que arredondam a coluna ou linha selecionada de modo que a soma dos arredondados seja o total arredondado (maior resto). Só fórmulas nativas."
    menuSeparator msepScale, btnSameScale, btnResetScale (iguais)
  mnuArrows image="tqq_arrows16" (em vez de imageMso; filhos iguais)
  menu mnuPpt "PowerPoint" image="tqq_ppt16" getEnabled=Rx_GetEnabled (sem tag) screentip="Copiar para o PowerPoint":
    btnPptImage "Copiar como imagem" onAction=Rx_OnPptImage getEnabled=Rx_GetEnabled supertip="Cola uma imagem do gráfico no slide ativo (abre o PowerPoint se preciso)."
    btnPptChart "Copiar como gráfico vinculado" onAction=Rx_OnPptChart getEnabled=Rx_GetEnabled supertip="Cola um gráfico do PowerPoint vinculado a esta pasta de trabalho salva (Dados > Editar Vínculos no PowerPoint)."
  separator sep3, btnRefresh, btnChangeSource (iguais)
grpAbout: inalterado
```
  Contagens: 8 menus, 20 botões de inserção, 10 toggles, 4 dropdowns, 3 botões de anotação, 16 atributos `image=` (14 nomes distintos, pois `tqq_round16` e `tqq_scatter16` aparecem duas vezes; 13 ícones da tabela `Get-TqqIconSpecs`, com `tqq_lines` em 32 px no menu e `tqq_lines16` no alternador). Nenhum filho de menu com `size`; nenhum controle com `image` e `imageMso`.
- Produces (`modRibbon`): `Rx_OnAxisBreak`, `Rx_OnAxisBreakOff`, `Rx_OnRoundRange`, `Rx_OnPptImage`, `Rx_OnPptChart` (todos `On Error Resume Next`, `CoreOrWarn`, `Loader_RunCore`; os dois de corte invalidam a faixa com `Loader_InvalidateRibbon`; os de PowerPoint e o de arredondar não mudam o estado da faixa).
- Produces (tutorial, parcial): maquete com "Borboleta, Dispersão, Bolhas" no menu Mais tipos, "PowerPoint ▾" no grupo do gráfico; `#tipos` com 3 linhas novas (menu › item, o que faz, tabela) e a tabela "Controles" com `Rótulos com soma consistente`, `Linha de tendência`, `Quadrantes (médias)`, `Corte de eixo...`, `Remover corte de eixo`, `Arredondar intervalo...`, `PowerPoint`, `Copiar como imagem`, `Copiar como gráfico vinculado` (todos em `<span class="ui">`); na matriz (Tarefa 0), troque `class="k"` por `class="ui"` nos três rótulos de linha (Borboleta, Dispersão, Bolhas) e nos cabeçalhos `roundsum`, `trend` e `quad`; o cabeçalho `data-f="break"` ("Corte de eixo", o rótulo do recurso e da mensagem 7002) **continua** `<span class="k">`, porque não é o texto de nenhum controle (a faixa tem "Corte de eixo..." e "Remover corte de eixo") e o `Test-Docs` confere todo `ui` contra a faixa. A Tarefa 11 escreve as seções completas.

- [ ] **Passo 1: Testes (RED)** — em `tests/vba/loader/TestRibbon.bas`:
  - `Test_Ribbon_Xml_HasAllCallbacks`: acrescente `"Rx_OnAxisBreak", "Rx_OnAxisBreakOff", "Rx_OnRoundRange", "Rx_OnPptImage", "Rx_OnPptChart"` à lista e `AssertTrue InStr(xml, "tag=""butterfly""") > 0`.
  - `Test_Ribbon_Tags_MatchCoreKeys`: `AssertEqual 20, ...Rx_OnInsert...`.
  - `Test_Ribbon_Menus_Structure`: `AssertEqual 8, d.SelectNodes("//u:menu").Length, "5 menus de insercao + Mais opcoes + Setas + PowerPoint"`; a lista de kinds válidos ganha `butterfly,scatter,bubble`; `AssertEqual 20` nos botões de inserção; `AssertEqual 7, d.SelectNodes("//u:menu[@id='mnuMore']/u:toggleButton").Length`; e:

```vb
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnBreak'][@tag='break'][@onAction='Rx_OnAxisBreak'][@getEnabled='Rx_GetEnabled']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnBreakOff'][@tag='break'][@onAction='Rx_OnAxisBreakOff']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnRoundRange'][@onAction='Rx_OnRoundRange']") Is Nothing
    AssertTrue d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnRoundRange']/@getEnabled") Is Nothing, "Arredondar intervalo age sobre celulas: sempre ativo"
    AssertEqual "Rx_GetEnabled", d.SelectSingleNode("//u:menu[@id='mnuPpt']").Attributes.getNamedItem("getEnabled").Text
    AssertEqual 2, d.SelectNodes("//u:menu[@id='mnuPpt']/u:button").Length
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:button[@id='btnPptImage'][@onAction='Rx_OnPptImage']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:button[@id='btnPptChart'][@onAction='Rx_OnPptChart']") Is Nothing
```
  - `Test_Ribbon_Toggles_And_Dropdowns_Phase2`: `AssertEqual 10` toggles; lista `",totals,labels,connectors,ind100,vline,secondary,pieout,roundsum,trend,quad,"`.
  - `Test_Ribbon_SelectedChartGroup_EnabledExceptRefresh`: no laço dos botões de `mnuMore`, `btnBreak`/`btnBreakOff` **têm** `getEnabled` e `btnSameScale`/`btnResetScale`/`btnRoundRange` **não**; o laço de `mnuArrows` fica; acrescente o de `mnuPpt` (todos com `getEnabled`).
  - `Test_Ribbon_Images_NamesAndNoDoubleIcon`: `AssertTrue c >= 13, "icones proprios: " & c`.
  - `Test_Ribbon_Enabled_FollowsCapabilityOfSelectedChart`: depois da pizza, acrescente `For Each t In Array("roundsum", "trend", "quad", "break")`: `AssertFalse Rx_IsEnabledFor(CStr(t)), "pizza: '" & t & "' deveria estar desabilitado"`; e após a cascata de duas séries: `InsertAndSelect ws, "A1:B5", "clustered"` (a tabela Categoria/Valor serve) com `AssertTrue Rx_IsEnabledFor("break")` e `AssertFalse Rx_IsEnabledFor("roundsum")`.
  - `ExpectedSig` já cobre `button/onAction` e `menu/getEnabled`.

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-f3\build\run-tests.ps1 -Target loader -TimeoutSec 400` → `not ok` nos testes alterados (e `Test-Docs` ainda verde, pois o XML não mudou).

- [ ] **Passo 3: Implementar** — XML conforme as Interfaces (mantenha os supertips existentes; os novos acima); `modRibbon.bas`:

```vb
Public Sub Rx_OnAxisBreak(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_SetAxisBreak"
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_OnAxisBreakOff(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_RemoveAxisBreak"
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_OnRoundRange(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_RoundRange"
End Sub

Public Sub Rx_OnPptImage(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_CopyToPowerPoint", "image"
End Sub

Public Sub Rx_OnPptChart(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_CopyToPowerPoint", "chart"
End Sub
```
Tutorial: a maquete (`.rb-btn` do menu Mais tipos e um novo `<div class="rb-btn sm"><span class="ui">PowerPoint</span><span class="car">▾</span></div>` ao lado de Setas), três linhas em `#tipos` (`Mais tipos › Borboleta`: "Duas séries espelhadas com os nomes no centro; totais no fim da barra direita." / "Categorias e exatamente 2 colunas de valores."; `Mais tipos › Dispersão`: "Pontos em dois eixos de valor, uma série por grupo, nome ao lado de cada ponto." / "Nome, X, Y e, opcionalmente, Grupo."; `Mais tipos › Bolhas`: "Como a dispersão, com o tamanho da bolha." / "Nome, X, Y, Tamanho e, opcionalmente, Grupo.") e, na tabela de controles, uma linha por controle novo (rótulo exato em `<span class="ui">` + uma frase). Na matriz (Tarefa 0), troque `class="k"` por `class="ui"` nos três rótulos de linha (Borboleta, Dispersão, Bolhas) e nos cabeçalhos `roundsum`, `trend` e `quad`; o cabeçalho `data-f="break"` ("Corte de eixo", o rótulo do recurso e da mensagem 7002) **continua** `<span class="k">`, porque não é o texto de nenhum controle (a faixa tem "Corte de eixo..." e "Remover corte de eixo") e o `Test-Docs` confere todo `ui` contra a faixa. Depois `pwsh C:\temp\tqq-f3\build\build.ps1 -Target loader -Configuration Release` e abra `dist\ThinQQ.xlam` **numa instância própria** (`New-ExcelApp` de `Build-Common.ps1`, com `Visible = $true` e mantendo `EnableEvents = $false`, o padrão de `New-ExcelApp`: com eventos, o `Workbook_Open` do Release roda `Loader_Startup`, que usa o `%LOCALAPPDATA%\ThinQQ` real) para ver a faixa com os ícones (captura de tela opcional); feche com `Close-ExcelApp`. Nunca instale.

- [ ] **Passo 4: GREEN e suíte completa** (`core` + `loader` + PowerShell; `Test-Ribbon.ps1` agora confere as 14 imagens distintas no pacote).

- [ ] **Passo 5: Commit** — `git add src/loader/ribbon/customUI14.xml src/loader/modules/modRibbon.bas tests/vba/loader/TestRibbon.bas docs/tutorial/tutorial.html` e `git commit -m "feat: faixa 1.2.0 com icones proprios, Borboleta/Dispersao/Bolhas, soma consistente, tendencia, quadrantes, corte de eixo, arredondar intervalo e PowerPoint"` (+ trailer).

**Verificação no Excel real (usuário):** aba ThinQQ com ícones próprios nos cinco menus; Mais tipos com Borboleta, Dispersão e Bolhas; com uma dispersão selecionada, Linha de tendência e Quadrantes ativos e Formato cinza; com colunas empilhadas, Rótulos com soma consistente e Corte de eixo... ativos; PowerPoint ▾ ativo só com gráfico selecionado; nenhum ícone em branco.

---
### Tarefa 9: Exemplos (planilha "Exemplos 3") — modelo: standard

**Files:**
- Modify: `src/core/modules/modSamples.bas`; `docs/tutorial/tutorial.html` e `README.md` **só** na contagem de exemplos (`Test-Docs` exige que os dois citem o maior número de exemplos de `Samples_Create`: `18 exemplos` → `26 exemplos`, `duas planilhas` → `três planilhas`, `"Exemplos" e "Exemplos 2"` → `"Exemplos", "Exemplos 2" e "Exemplos 3"`; no item 25 das verificações manuais do README, `percorrer "Exemplos 2"` → `percorrer "Exemplos 2" e "Exemplos 3"`; o conteúdo novo fica para a Tarefa 11)
- Test: `tests/vba/core/TestSamples.bas`, `tests/ps/Test-Docs.ps1` (existente: `tutorial e README citam $nEx exemplos`)

**Interfaces:**
- Consumes: `Kinds_Insert`, `Scatter_SetTrend`, `Scatter_SetQuad`, `Round_SetLabels`, `Break_Set`, `Core_SafeRefresh`.
- Produces: `Samples_CreateSheet3(wb) As Worksheet` (privada), chamada por `Samples_Create` depois de `Samples_CreateSheet2`; planilha "Exemplos 3" (depois de "Exemplos 2") com 8 seções (19 a 26), títulos na coluna A nas linhas `1, 22, 43, 64, 85, 106, 127, 148`, tabela em `A(t+1)`, gráfico 2 colunas à direita, larguras `A = 18`, `B:F = 14`:

| # | Linha | Título | Tabela | Gráfico |
|---|---|---|---|---|
| 19 | 1 | `19) Borboleta: duas séries espelhadas` | Faixa/Mulheres/Homens: 0-14 120 130; 15-29 200 210; 30-59 310 290; 60+ 90 60 | `butterfly` |
| 20 | 22 | `20) Dispersão por grupo, com tendência e quadrantes` | Produto/Preço/Margem/Grupo: P1 10 0,20 A; P2 20 0,35 B; P3 30 0,25 A; P4 40 0,50 B; P5 25 0,30 A; P6 35 0,45 B (Margem em `0%`) | `scatter` + `Scatter_SetTrend True` + `Scatter_SetQuad True` |
| 21 | 43 | `21) Bolhas: tamanho = volume` | Produto/Preço/Margem/Volume/Grupo: P1 10 0,20 50 A; P2 20 0,35 120 B; P3 30 0,25 80 A; P4 40 0,50 20 B | `bubble` |
| 22 | 64 | `22) Colunas 100% com rótulos de soma consistente (34+33+33)` | Ano/S1/S2/S3: 2024 1 1 1; 2025 2 3 5 | `stacked100` + `Round_SetLabels True` |
| 23 | 85 | `23) Colunas empilhadas com corte de eixo em 50` | Cat/S1/S2: a 5 5; b 10 10; c 200 30 | `stacked` + `Break_Set 50` |
| 24 | 106 | `24) Barras empilhadas com corte de eixo` | a mesma tabela do 23 | `barstacked` + `Break_Set 50` |
| 25 | 127 | `25) Área empilhada com soma consistente` | Ano/Sul/Norte/Leste: 2024 10,4 10,4 10,4; 2025 1 1 1; 2026 2,5 2,5 2,5 | `area` + `Round_SetLabels True` |
| 26 | 148 | `26) Dispersão sem grupo, com quadrantes` | Produto/Preço/Margem: as 6 linhas do 20 sem a coluna Grupo | `scatter` + `Scatter_SetQuad True` |

  `Settle` (= `Core_SafeRefresh`) após cada exemplo, como na planilha 2. Ao final, `Samples_Create` continua ativando "Exemplos" com `A1`.

- [ ] **Passo 1: Testes (RED)** — em `tests/vba/core/TestSamples.bas`: renomeie `Test_Samples_BothSheets` para `Test_Samples_AllSheets`, troque no comentário do topo do módulo `Samples_Create monta 19 graficos` por `Samples_Create monta 27 graficos`, acrescente `CheckSheet3KindsAndFeatures wb`, `CheckSheet3Layout wb` e `CheckSheet3Labels wb` depois de `CheckSheet2Labels wb`, e as funções:

```vb
Private Function TitleRows3() As Variant
    Dim r(0 To 7) As Long, k As Long
    For k = 0 To 7: r(k) = 1 + 21 * k: Next
    TitleRows3 = r
End Function

Private Function ExpectedKinds3() As Variant
    ExpectedKinds3 = Array("butterfly", "scatter", "bubble", "stacked100", "stacked", "barstacked", "area", "scatter")
End Function

Private Sub CheckSheet3KindsAndFeatures(wb As Workbook)
    Dim ws As Worksheet, co As ChartObject, ek As Variant, i As Long, c As Chart, blk As Range
    Set ws = wb.Worksheets("Exemplos 3")
    AssertEqual wb.Worksheets("Exemplos 2").Index + 1, ws.Index, "a terceira planilha vem logo depois da segunda"
    AssertEqual 8, ws.ChartObjects.Count, "8 secoes, um grafico cada"
    ek = ExpectedKinds3()
    For Each co In ws.ChartObjects
        AssertTrue Registry_IsThinQQChart(co.Chart), co.Name
        i = i + 1
        AssertEqual CStr(ek(i - 1)), Registry_ChartType(co.Chart), "tipo do grafico " & i & " de Exemplos 3"
    Next
    Set c = ws.ChartObjects(2).Chart
    AssertEqual "1", Registry_GetOption(c, "trend"), "20: tendencia": AssertEqual "1", Registry_GetOption(c, "quad"), "20: quadrantes"
    AssertEqual 2, c.SeriesCollection.Count, "20: dois grupos": AssertEqual 1, c.SeriesCollection(1).Trendlines.Count
    AssertEqual xlAxisCrossesCustom, c.Axes(xlCategory).Crosses
    AssertEqual xlBubble, ws.ChartObjects(3).Chart.ChartType, "21: bolhas"
    Set c = ws.ChartObjects(4).Chart: Set blk = Registry_Block(c)
    AssertEqual "1", Registry_GetOption(c, "roundsum"), "22: soma consistente": AssertTrue Round_RCol(blk, 3) > 0
    Set c = ws.ChartObjects(5).Chart
    AssertEqual "50", Registry_GetOption(c, "vbreak"), "23: corte em 50": AssertEqual 50, c.Axes(xlValue).MaximumScale
    AssertTrue HasShape(c, "tqq_ann_break_3"), "23: marca de corte"
    Set c = ws.ChartObjects(6).Chart
    AssertEqual "50", Registry_GetOption(c, "vbreak"), "24: corte nas barras": AssertTrue HasShape(c, "tqq_ann_break_3")
    AssertEqual "1", Registry_GetOption(ws.ChartObjects(7).Chart, "roundsum"), "25: area com soma consistente"
    Set c = ws.ChartObjects(8).Chart
    AssertEqual 1, c.SeriesCollection.Count, "26: sem grupo = uma serie": AssertEqual "1", Registry_GetOption(c, "quad")
End Sub

Private Sub CheckSheet3Layout(wb As Workbook)
    Dim ws As Worksheet, i As Long, j As Long, tr As Variant, a As ChartObject, b As ChartObject
    Set ws = wb.Worksheets("Exemplos 3")
    tr = TitleRows3()
    For i = 1 To ws.ChartObjects.Count
        Set a = ws.ChartObjects(i)
        AssertTrue a.Top >= ws.Cells(tr(i - 1), 1).Top, "grafico " & i & " acima do titulo da secao " & (i + 18)
        If i < 8 Then AssertTrue a.Top + a.Height <= ws.Cells(tr(i), 1).Top, "grafico " & i & " invade a secao seguinte"
        For j = i + 1 To ws.ChartObjects.Count
            Set b = ws.ChartObjects(j)
            AssertFalse a.Left < b.Left + b.Width And b.Left < a.Left + a.Width And a.Top < b.Top + b.Height And b.Top < a.Top + a.Height, "graficos " & i & " e " & j & " sobrepostos"
        Next
    Next
    For i = 0 To 7
        AssertTrue Len(CStr(ws.Cells(tr(i), 1).Value)) > 0, "titulo " & (i + 19)
    Next
    AssertEqual "19) ", Left$(CStr(ws.Cells(1, 1).Value), 4): AssertEqual "26) ", Left$(CStr(ws.Cells(148, 1).Value), 4)
End Sub

' Textos realmente exibidos (expectativas das celulas e da localidade).
Private Sub CheckSheet3Labels(wb As Workbook)
    Dim ws As Worksheet, cht As Chart
    Set ws = wb.Worksheets("Exemplos 3")
    Set cht = ws.ChartObjects(1).Chart
    AssertEqual Format$(ws.Range("B2").Value + ws.Range("C2").Value, "#,##0"), Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "19: total da 1a faixa"
    AssertEqual Format$(ws.Range("B2").Value, "#,##0"), Bfly_SideSeries(cht, 1).Points(1).DataLabel.Text, "19: lado esquerdo sem sinal"
    AssertEqual "P1", ws.ChartObjects(2).Chart.SeriesCollection(1).Points(1).DataLabel.Text, "20: nome do ponto"
    Set cht = ws.ChartObjects(4).Chart
    AssertEqual Format$(0.34, "0%"), cht.SeriesCollection(1).Points(1).DataLabel.Text, "22: 34 + 33 + 33 (empate: a primeira coluna recebe o resto)"
    AssertEqual Format$(0.33, "0%"), cht.SeriesCollection(3).Points(1).DataLabel.Text
    AssertEqual Format$(230, "#,##0"), Chart_FindSeries(ws.ChartObjects(5).Chart, "tqq_total").Points(3).DataLabel.Text, "23: total real acima do corte"
End Sub
```
(`HasShape` já existe no módulo.) `CheckNoErrorsAnywhere` continua valendo para a planilha nova.

- [ ] **Passo 2: RED** — `-Filter "TestSamples"` → `not ok` (planilha "Exemplos 3" inexistente).

- [ ] **Passo 3: Implementar** — em `modSamples.bas`, depois de `Samples_CreateSheet2`:

```vb
' Terceira planilha ("Exemplos 3"): os tipos e recursos da Fase 3, 8 secoes (19 a 26) a cada 21 linhas, pelos mesmos caminhos do usuario.
Private Function Samples_CreateSheet3(wb As Workbook) As Worksheet
    Dim ws As Worksheet, d As clsChartData, co As ChartObject
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets("Exemplos 2"))
    ws.Name = "Exemplos 3"
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B:F").ColumnWidth = 14

    WriteTitle ws, 1, "19) Borboleta: duas séries espelhadas"
    WriteTable ws.Range("A2"), Array(Array("Faixa", "Mulheres", "Homens"), Array("0-14", 120, 130), Array("15-29", 200, 210), Array("30-59", 310, 290), Array("60+", 90, 60))
    Set d = ChartData_FromRange(ws.Range("A2:C6"))
    Set co = Kinds_Insert(ws, d, "butterfly")
    Settle co

    WriteTitle ws, 22, "20) Dispersão por grupo, com tendência e quadrantes"
    WriteTable ws.Range("A23"), Array(Array("Produto", "Preço", "Margem", "Grupo"), Array("P1", 10, 0.2, "A"), Array("P2", 20, 0.35, "B"), _
                                      Array("P3", 30, 0.25, "A"), Array("P4", 40, 0.5, "B"), Array("P5", 25, 0.3, "A"), Array("P6", 35, 0.45, "B"))
    ws.Range("C24:C29").NumberFormat = "0%"
    Set d = ChartData_FromRange(ws.Range("A23:D29"))
    Set co = Kinds_Insert(ws, d, "scatter")
    Scatter_SetTrend co.Chart, True
    Scatter_SetQuad co.Chart, True
    Settle co

    WriteTitle ws, 43, "21) Bolhas: tamanho = volume"
    WriteTable ws.Range("A44"), Array(Array("Produto", "Preço", "Margem", "Volume", "Grupo"), Array("P1", 10, 0.2, 50, "A"), Array("P2", 20, 0.35, 120, "B"), _
                                      Array("P3", 30, 0.25, 80, "A"), Array("P4", 40, 0.5, 20, "B"))
    ws.Range("C45:C48").NumberFormat = "0%"
    Set d = ChartData_FromRange(ws.Range("A44:E48"))
    Set co = Kinds_Insert(ws, d, "bubble")
    Settle co

    WriteTitle ws, 64, "22) Colunas 100% com rótulos de soma consistente (34+33+33)"
    WriteTable ws.Range("A65"), Array(Array("Ano", "S1", "S2", "S3"), Array(2024, 1, 1, 1), Array(2025, 2, 3, 5))
    Set d = ChartData_FromRange(ws.Range("A65:D67"))
    Set co = Kinds_Insert(ws, d, "stacked100")
    Round_SetLabels co.Chart, True
    Settle co

    WriteTitle ws, 85, "23) Colunas empilhadas com corte de eixo em 50"
    WriteTable ws.Range("A86"), Array(Array("Cat", "S1", "S2"), Array("a", 5, 5), Array("b", 10, 10), Array("c", 200, 30))
    Set d = ChartData_FromRange(ws.Range("A86:C89"))
    Set co = Kinds_Insert(ws, d, "stacked")
    Break_Set co.Chart, 50
    Settle co

    WriteTitle ws, 106, "24) Barras empilhadas com corte de eixo"
    WriteTable ws.Range("A107"), Array(Array("Cat", "S1", "S2"), Array("a", 5, 5), Array("b", 10, 10), Array("c", 200, 30))
    Set d = ChartData_FromRange(ws.Range("A107:C110"))
    Set co = Kinds_Insert(ws, d, "barstacked")
    Break_Set co.Chart, 50
    Settle co

    WriteTitle ws, 127, "25) Área empilhada com soma consistente"
    WriteTable ws.Range("A128"), Array(Array("Ano", "Sul", "Norte", "Leste"), Array(2024, 10.4, 10.4, 10.4), Array(2025, 1, 1, 1), Array(2026, 2.5, 2.5, 2.5))
    Set d = ChartData_FromRange(ws.Range("A128:D131"))
    Set co = Kinds_Insert(ws, d, "area")
    Round_SetLabels co.Chart, True
    Settle co

    WriteTitle ws, 148, "26) Dispersão sem grupo, com quadrantes"
    WriteTable ws.Range("A149"), Array(Array("Produto", "Preço", "Margem"), Array("P1", 10, 0.2), Array("P2", 20, 0.35), Array("P3", 30, 0.25), _
                                       Array("P4", 40, 0.5), Array("P5", 25, 0.3), Array("P6", 35, 0.45))
    ws.Range("C150:C155").NumberFormat = "0%"
    Set d = ChartData_FromRange(ws.Range("A149:C155"))
    Set co = Kinds_Insert(ws, d, "scatter")
    Scatter_SetQuad co.Chart, True
    Settle co
    Set Samples_CreateSheet3 = ws
End Function
```
Em `Samples_Create`, depois de `Samples_CreateSheet2 wb`: `Samples_CreateSheet3 wb`. Atualize o comentário do cabeçalho do módulo (três planilhas).

- [ ] **Passo 4: GREEN e suíte completa.** Antes de rodar a suíte, troque `18 exemplos` por `26 exemplos` (e `duas planilhas` por `três planilhas`; `"Exemplos" e "Exemplos 2"` por `"Exemplos", "Exemplos 2" e "Exemplos 3"`; no item 25 do README, `percorrer "Exemplos 2"` por `percorrer "Exemplos 2" e "Exemplos 3"`) em `docs/tutorial/tutorial.html` e em `README.md` (só essas menções; `Test-Docs` compara o número citado com o maior exemplo de `Samples_Create`). Exporte os 8 gráficos para PNG (planilha ativada e rolada até cada gráfico antes do `Export`) e olhe cada um; registre.

- [ ] **Passo 5: Commit** — `git add src/core/modules/modSamples.bas tests/vba/core/TestSamples.bas docs/tutorial/tutorial.html README.md` e `git commit -m "feat: planilha Exemplos 3 com os tipos e recursos da Fase 3"` (+ trailer).

**Verificação no Excel real:** ThinQQ > Exemplos abre a pasta com três planilhas; percorrer "Exemplos 3".

---

### Tarefa 10: Pasta de demonstração (dados fictícios de materiais de construção, um exemplo por funcionalidade, índice com hiperlinks) — modelo: most-capable

**Files:**
- Create: `tools/DemoCoverage.ps1`, `tools/New-DemoWorkbook.ps1`, `tests/ps/Test-DemoWorkbook.ps1`, `exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx` (gerado e commitado)
- Modify: `build/run-tests.ps1` (chama `Test-DemoWorkbook.ps1`), `.gitignore` (nada: `exemplos/` é versionado)

**Interfaces:**
- Consumes (core 1.2.0 de **Teste**, via `Application.Run` no Excel do script): `Core_Insert`, `Core_SetToggle`, `Core_GetToggle`, `Core_SetOption`, `Core_AddAnnotation(kind, from, to)`, `Core_SameScale`, `Core_SetAxisBreakAt`, `Round_WriteFormulas(src, dest, decimals)`, `Core_ClearLastNotice`/`Core_LastNotice` (o gerador falha se um comando deixar um aviso). O core de Teste é usado porque em Release `Core_Notify` abre `MsgBox` e travaria a automação; o `.xlsx` resultante não leva código.
- Produces (`tools/DemoCoverage.ps1`, dot-sourced pelo gerador e pelo teste): `Get-TqqDemoFeatures` (lista fixa de recursos visíveis: `totals, labels, legend-none, lineend, numfmt, gap, palette, cagr, diff, bracket, vline, ind100, samescale, secondary, pieout, gantt-weeks, gantt-months, roundsum, roundrange, break, trend, quad, connectors`) e `Get-TqqDemoCoverage` (26 itens `@{ Num; Sheet; Title; Kind; Features; Menu; Desc }`, abaixo). Regra verificada pelo teste: todo tipo de `modChartKinds.bas` aparece em `Kind` de ≥ 1 item e todo recurso de `Get-TqqDemoFeatures` aparece em `Features` de ≥ 1 item; o índice tem exatamente um hiperlink por item.
- Produces (`tools/New-DemoWorkbook.ps1 -CoreXlam <dist\ThinQQ.core.test.xlam da cópia> [-OutFile <xlsx>]`): Excel novo invisível (`New-ExcelApp`), `EnableEvents = $false`, abre o core de Teste, cria a pasta com `Indice` (primeira), `Dados` (8 lojas × 8 departamentos × 24 meses: Loja, Cidade, Região, Departamento, Mês, Vendas, Custo, Margem, Unidades, Ticket médio, Canal B2C/B2B — gerador congruencial com semente fixa 20261011), `Lojas` (loja, cidade, região, meta anual, inauguração), `Obras` (cronograma de reformas/expansões) e uma planilha `E01`..`E26` por item (A1 = hiperlink "Voltar ao índice", A3 = legenda "<título> — Menu: <menu>", tabela em `A5`, gráfico 2 colunas à direita), salva como `.xlsx` (51) com `RemovePersonalInformation = $true` (sem autor em `docProps`; nunca altere `Application.UserName`, que grava no registro do Office do usuário). Nomes fictícios ("Casa & Obra Materiais", lojas "Loja Vale Verde", ...), nenhuma pessoa, empresa real ou dado pessoal.
- Produces (`tests/ps/Test-DemoWorkbook.ps1`): abre o `.xlsx` num Excel próprio (somente leitura); primeira planilha = `Indice`; cada hiperlink do índice tem `SubAddress` para uma planilha existente, que contém ≥ 1 `ChartObject`; nº de hiperlinks = nº de itens da cobertura; cobertura × matriz (tipos **e recursos** lidos de `modChartKinds.bas`; um recurso conta coberto por um item com o mesmo nome ou com o prefixo `<recurso>-`, como `legend-none`) e × a lista fixa de `Get-TqqDemoFeatures`; toda planilha `E##` tem em `A1` um hiperlink para `Indice`; nenhuma célula com valor de erro (fora da `_ThinQQ`); `_ThinQQ` com `Visible = xlSheetVeryHidden` (2); tamanho < 3 MB; `docProps/core.xml` sem autor (`dc:creator`/`cp:lastModifiedBy` vazios) e nenhuma parte XML com id de máquina ou caminho de perfil (o `.xlsx` é binário: o `Test-Privacy` não o lê). Corpo em `try/catch`, `exit 0|1` no fim; o `run-tests` o chama num processo próprio.

Itens da cobertura (`Num` → planilha `E<Num>`; `Kind`; recursos; tabela de origem gerada a partir de `Dados`):
| Num | Kind | Features | Tabela | Comandos depois de `Core_Insert` |
|---|---|---|---|---|
| 1 | stacked | totals, labels | Departamento × 2024/2025 (vendas, R$ mil; `SUMIFS` sobre `Dados`) | — |
| 2 | clustered | gap | Loja × 2024/2025 | `Core_SetOption gap 120` (as agrupadas já nascem com 80) |
| 3 | stacked100 | ind100 | Região × B2C/B2B | — |
| 4 | line | lineend | Mês (12 de 2025) × Vendas/Custo da loja 1 | `Core_SetOption legend lineend` |
| 5 | waterfall | connectors | Etapa/Valor (Receita 2025, -Custo, -Despesas, EBITDA e, -Impostos, Lucro e) | — |
| 6 | barstacked | diff | Região × 2024/2025 | `Core_AddAnnotation diff 1 n` |
| 7 | barclustered | numfmt | Loja × 2024/2025 (R$) | `Core_SetOption numfmt mil` |
| 8 | barstacked100 | — | Loja × B2C/B2B | — |
| 9 | barwaterfall | — | mesma do 5 | — |
| 10 | area | vline | Mês (12) × 3 departamentos | `Core_SetToggle vline True` |
| 11 | area100 | — | mesma do 10 | — |
| 12 | combo | secondary | Mês (12) × Vendas/Custo/Margem % | — (o secundário liga sozinho; o gerador afirma com `Core_GetToggle secondary`) |
| 13 | mekko | — | Região × 3 grupos de departamento | — |
| 14 | mekkounits | — | Região × Largura (nº de lojas) × 3 grupos | — |
| 15 | gantt | bracket, gantt-weeks | 4 obras de 2026 (≤ 14 semanas) | `Core_AddAnnotation bracket 1 2` |
| 16 | gantt | gantt-months | 5 obras ao longo de 18 meses | — |
| 17 | pie | pieout | Departamento × vendas 2025 | `Core_SetToggle pieout True` |
| 18 | doughnut | palette | mesma do 17 | `Core_SetOption palette cinza` |
| 19 | butterfly | — | Loja × B2C/B2B | — |
| 20 | scatter | trend, quad | Loja × Ticket médio × Margem % × Região | `Core_SetToggle trend True`, `quad True` |
| 21 | bubble | — | Departamento × Ticket × Margem % × Vendas × Grupo | — |
| 22 | stacked | cagr | Ano (2024, 2025) × 3 departamentos (`SUMIFS`) | `Core_AddAnnotation cagr 1 2` |
| 23 | stacked100 | roundsum, roundrange | Região × 3 departamentos (participações) + coluna de `Round_WriteFormulas` ao lado da tabela | `Core_SetToggle roundsum True`; `Round_WriteFormulas` |
| 24 | stacked | break | Departamento 2025 (cimento ≫ demais) | `Core_SetAxisBreakAt <2º maior × 1,2>` |
| 25 | clustered (2 gráficos) | samescale | Loja 1 × 2024/2025 e Loja 2 × 2024/2025 | `Core_SameScale` |
| 26 | line | legend-none | Mês (12) × Vendas da loja 1 | `Core_SetOption legend none` |

- [ ] **Passo 1: Cobertura e teste (RED)** — `tools/DemoCoverage.ps1` (UTF-8 **com BOM**, como `installer/install.ps1`: tem acentos e também é lido pelo Windows PowerShell 5.1):

```powershell
# Lista de cobertura da pasta de demonstracao: a MESMA fonte para o gerador (tools/New-DemoWorkbook.ps1) e o teste (tests/ps/Test-DemoWorkbook.ps1).
function Get-TqqDemoFeatures {
    return @('totals', 'labels', 'legend-none', 'lineend', 'numfmt', 'gap', 'palette', 'cagr', 'diff', 'bracket', 'vline', 'ind100', 'samescale',
             'secondary', 'pieout', 'gantt-weeks', 'gantt-months', 'roundsum', 'roundrange', 'break', 'trend', 'quad', 'connectors')
}
function Get-TqqDemoCoverage {
    $items = @(
        @{ Num = 1;  Kind = 'stacked';       Features = @('totals', 'labels'); Menu = 'Colunas > Colunas empilhadas';        Title = 'Vendas por departamento (colunas empilhadas, totais e rótulos)' }
        @{ Num = 2;  Kind = 'clustered';     Features = @('gap');              Menu = 'Colunas > Colunas agrupadas';         Title = 'Vendas por loja (agrupadas, Espaço muito largo)' }
        @{ Num = 3;  Kind = 'stacked100';    Features = @('ind100');           Menu = 'Colunas > Colunas 100%';              Title = 'Mix B2C/B2B por região (100% e indicador)' }
        @{ Num = 4;  Kind = 'line';          Features = @('lineend');          Menu = 'Linha e área > Linha';                Title = 'Vendas e custo mensais (nome no fim da linha)' }
        @{ Num = 5;  Kind = 'waterfall';     Features = @('connectors');       Menu = 'Cascata > Cascata (colunas)';         Title = 'Da receita ao lucro (cascata com conectores)' }
        @{ Num = 6;  Kind = 'barstacked';    Features = @('diff');             Menu = 'Barras > Barras empilhadas';          Title = 'Vendas por região (barras e seta de diferença)' }
        @{ Num = 7;  Kind = 'barclustered';  Features = @('numfmt');           Menu = 'Barras > Barras agrupadas';           Title = 'Vendas por loja em milhares (Formato)' }
        @{ Num = 8;  Kind = 'barstacked100'; Features = @();                   Menu = 'Barras > Barras 100%';                Title = 'Mix B2C/B2B por loja (barras 100%)' }
        @{ Num = 9;  Kind = 'barwaterfall';  Features = @();                   Menu = 'Cascata > Cascata (barras)';          Title = 'Da receita ao lucro (cascata horizontal)' }
        @{ Num = 10; Kind = 'area';          Features = @('vline');            Menu = 'Linha e área > Área empilhada';       Title = 'Vendas mensais de três departamentos (área e linha de valor)' }
        @{ Num = 11; Kind = 'area100';       Features = @();                   Menu = 'Linha e área > Área 100%';            Title = 'Participação mensal (área 100%)' }
        @{ Num = 12; Kind = 'combo';         Features = @('secondary');        Menu = 'Linha e área > Combinação (colunas + linha)'; Title = 'Vendas, custo e margem % (combinação com eixo secundário)' }
        @{ Num = 13; Kind = 'mekko';         Features = @();                   Menu = 'Mais tipos > Mekko (%)';              Title = 'Regiões x grupos de produto (Mekko %)' }
        @{ Num = 14; Kind = 'mekkounits';    Features = @();                   Menu = 'Mais tipos > Mekko (unidades)';       Title = 'Regiões x grupos, largura = nº de lojas (Mekko unidades)' }
        @{ Num = 15; Kind = 'gantt';         Features = @('bracket', 'gantt-weeks'); Menu = 'Mais tipos > Gantt';            Title = 'Reformas do trimestre (Gantt em semanas, colchete)' }
        @{ Num = 16; Kind = 'gantt';         Features = @('gantt-months');     Menu = 'Mais tipos > Gantt';                  Title = 'Plano de expansão (Gantt em meses)' }
        @{ Num = 17; Kind = 'pie';           Features = @('pieout');           Menu = 'Mais tipos > Pizza';                  Title = 'Vendas 2025 por departamento (pizza, rótulos fora)' }
        @{ Num = 18; Kind = 'doughnut';      Features = @('palette');          Menu = 'Mais tipos > Rosca';                  Title = 'Vendas 2025 por departamento (rosca, tons de cinza)' }
        @{ Num = 19; Kind = 'butterfly';     Features = @();                   Menu = 'Mais tipos > Borboleta';              Title = 'B2C x B2B por loja (borboleta)' }
        @{ Num = 20; Kind = 'scatter';       Features = @('trend', 'quad');    Menu = 'Mais tipos > Dispersão';              Title = 'Ticket médio x margem por loja (dispersão, tendência, quadrantes)' }
        @{ Num = 21; Kind = 'bubble';        Features = @();                   Menu = 'Mais tipos > Bolhas';                 Title = 'Ticket x margem x vendas por departamento (bolhas)' }
        @{ Num = 22; Kind = 'stacked';       Features = @('cagr');             Menu = 'Colunas > Colunas empilhadas';        Title = 'Crescimento 2024-2025 (seta CAGR)' }
        @{ Num = 23; Kind = 'stacked100';    Features = @('roundsum', 'roundrange'); Menu = 'Mais opções > Rótulos com soma consistente'; Title = 'Participações que somam 100% (soma consistente no gráfico e na tabela)' }
        @{ Num = 24; Kind = 'stacked';       Features = @('break');            Menu = 'Mais opções > Corte de eixo...';      Title = 'Cimento domina: corte de eixo' }
        @{ Num = 25; Kind = 'clustered';     Features = @('samescale');        Menu = 'Mais opções > Mesma escala';          Title = 'Duas lojas lado a lado (mesma escala)' }
        @{ Num = 26; Kind = 'line';          Features = @('legend-none');      Menu = 'Legenda > Nenhuma';                   Title = 'Vendas mensais da loja 1 (sem legenda)' }
    )
    foreach ($i in $items) { $i.Sheet = ('E{0:00}' -f $i.Num); $i.Desc = $i.Title }
    return $items
}
```

`tests/ps/Test-DemoWorkbook.ps1`:

```powershell
# Pasta de demonstracao: indice com hiperlinks internos, um grafico por exemplo, cobertura = tipos e recursos da MATRIZ (lidos de
# modChartKinds.bas) + a lista fixa de recursos, sem erros, _ThinQQ muito oculta, < 3 MB, docProps sem autor. Excel proprio (New-ExcelApp),
# somente leitura, nada e gravado. O run-tests o chama num processo proprio; termina sempre com exit 0 ou exit 1.
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$script:fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }
try {
    . (Join-Path $RepoRoot 'build\lib\Build-Common.ps1')
    . (Join-Path $RepoRoot 'tools\DemoCoverage.ps1')
    $path = Join-Path $RepoRoot 'exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx'
    Check (Test-Path $path) "pasta de demonstracao existe: $path"
    if (-not (Test-Path $path)) { exit 1 }
    Check ((Get-Item $path).Length -lt 3MB) "tamanho < 3 MB ($((Get-Item $path).Length) bytes)"
    $cov = Get-TqqDemoCoverage
    $feats = Get-TqqDemoFeatures
    $kindsSrc = [IO.File]::ReadAllText((Join-Path $RepoRoot 'src\core\modules\modChartKinds.bas'))
    $allKinds = @(([regex]::Match($kindsSrc, 'Private Const KINDS As String = "(?<k>[^"]+)"').Groups['k'].Value).Trim(',').Split(','))
    foreach ($k in $allKinds) { Check (@($cov | Where-Object { $_.Kind -eq $k }).Count -ge 1) "cobertura tem o tipo '$k'" }
    # recursos da MATRIZ (a mesma leitura do Test-Docs), nao so a lista fixa de DemoCoverage.ps1 (que viria do mesmo arquivo da cobertura):
    # um recurso conta coberto por um item com o mesmo nome ou com o prefixo "<recurso>-" (legend -> legend-none)
    $matrixFeatures = @([regex]::Matches($kindsSrc, 'Case "(?<f>[a-z0-9]+)":\s*list =') | ForEach-Object { $_.Groups['f'].Value })
    Check ($matrixFeatures.Count -ge 20) "recursos da matriz lidos de modChartKinds.bas ($($matrixFeatures.Count))"
    foreach ($f in $matrixFeatures) { Check (@($cov | Where-Object { $_.Features -contains $f -or @($_.Features | Where-Object { $_ -like "$f-*" }).Count -gt 0 }).Count -ge 1) "cobertura tem o recurso da matriz '$f'" }
    foreach ($f in $feats) { Check (@($cov | Where-Object { $_.Features -contains $f }).Count -ge 1) "cobertura tem o recurso '$f'" }
    Check (@($cov | ForEach-Object { $_.Sheet } | Sort-Object -Unique).Count -eq $cov.Count) 'planilhas da cobertura sem repeticao'
    # privacidade do binario (o Test-Privacy pula arquivo com byte NUL): docProps sem autor e nenhuma parte XML com id ou caminho de perfil
    # (as duas expressoes do Test-Privacy, montadas em pedacos para este arquivo nao casar com elas)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $privPatterns = @(('(?<!\d)14' + '3272(?!\d)'), ('(?i)[a-z]:[\\/]+' + 'Us' + 'ers[\\/]+(?![<%$\{]|\.\.\.)[^\\/\s''"`<>]+'))
    $privHits = @()
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        foreach ($e in @($zip.Entries | Where-Object { $_.FullName -like '*.xml' -and ($_.FullName -like 'docProps/*' -or $_.FullName -like 'xl/*') })) {
            $sr = New-Object IO.StreamReader($e.Open()); $txt = $sr.ReadToEnd(); $sr.Dispose()
            foreach ($p in $privPatterns) { if ($txt -match $p) { $privHits += $e.FullName } }
            if ($e.FullName -eq 'docProps/core.xml') {
                Check ($txt -notmatch '<dc:creator>[^<]+</dc:creator>') 'docProps/core.xml sem autor (dc:creator vazio ou ausente)'
                Check ($txt -notmatch '<cp:lastModifiedBy>[^<]+</cp:lastModifiedBy>') 'docProps/core.xml sem cp:lastModifiedBy'
            }
        }
    } finally { $zip.Dispose() }
    Check ($privHits.Count -eq 0) "partes XML sem id de maquina nem caminho de perfil ($($privHits -join ', '))"
    $xl = New-ExcelApp
    $wd = Start-ExcelWatchdog $xl 240
    try {
        $wb = $xl.Workbooks.Open($path, 0, $true)
        Check ($wb.Worksheets.Item(1).Name -eq 'Indice') 'primeira planilha e Indice'
        $idx = $wb.Worksheets.Item('Indice')
        $links = @(); foreach ($h in $idx.Hyperlinks) { $links += $h }
        Check ($links.Count -eq $cov.Count) "indice com um hiperlink por exemplo ($($links.Count) de $($cov.Count))"
        $sheetNames = @(); foreach ($s in $wb.Sheets) { $sheetNames += $s.Name }
        foreach ($h in $links) {
            $target = ($h.SubAddress -replace "^'?(.*?)'?!.*$", '$1')
            Check ($sheetNames -contains $target) "destino existe: $($h.SubAddress)"
            if ($sheetNames -contains $target) {
                $ws = $wb.Worksheets.Item($target)
                Check ($ws.ChartObjects().Count -ge 1) "'$target' tem grafico"
                $back = $ws.Range('A1').Hyperlinks
                Check ($back.Count -eq 1 -and $back.Item(1).SubAddress -match 'Indice') "'$target'!A1 volta ao indice"
            }
        }
        foreach ($it in $cov) { Check ($sheetNames -contains $it.Sheet) "planilha $($it.Sheet) do item $($it.Num)" }
        $calc = $null; try { $calc = $wb.Worksheets.Item('_ThinQQ') } catch { }
        Check ($null -ne $calc -and $calc.Visible -eq 2) '_ThinQQ muito oculta'
        $errs = 0
        foreach ($ws in $wb.Worksheets) {
            if ($ws.Name -eq '_ThinQQ') { continue }
            $ur = $ws.UsedRange
            $vals = $ur.Value2
            if ($null -eq $vals) { continue }
            if ($vals -is [array]) { foreach ($v in $vals) { if ($v -is [int] -and $v -lt -2146826000 -and $v -gt -2146827000) { $errs++ } } }
            elseif ($vals -is [int] -and $vals -lt -2146826000) { $errs++ }
        }
        Check ($errs -eq 0) "nenhuma celula com valor de erro ($errs)"
        $wb.Close($false)
    } catch { Check $false "excecao: $($_.Exception.Message)" }
    finally { Stop-ExcelWatchdog $wd; Close-ExcelApp $xl }
} catch {
    Check $false "excecao: $($_.Exception.Message)"
}
if ($script:fail -gt 0) { Write-Host "FALHAS demonstracao: $script:fail"; exit 1 }
Write-Host 'Test-DemoWorkbook: OK'
exit 0
```
(Valores de erro do Excel chegam ao PowerShell como `Int32` entre -2146826281 e -2146826246.)

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-f3\tests\ps\Test-DemoWorkbook.ps1 -RepoRoot C:\temp\tqq-f3` → `not ok - pasta de demonstracao existe`.

- [ ] **Passo 3: Gerador `tools/New-DemoWorkbook.ps1`** (UTF-8 **com BOM**, pelo mesmo motivo):

```powershell
<#
.SYNOPSIS  Gera exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx: dados FICTICIOS deterministicos de uma rede de materiais de construcao e um
           exemplo por funcionalidade do ThinQQ, pela API real do core (Application.Run num Excel novo e invisivel).
.EXAMPLE   robocopy <repo> C:\temp\tqq-f3 /MIR /XD .git .superpowers dist ; pwsh C:\temp\tqq-f3\build\build.ps1 -Target core -Configuration Test
           pwsh tools\New-DemoWorkbook.ps1 -CoreXlam C:\temp\tqq-f3\dist\ThinQQ.core.test.xlam
.NOTES     Core de TESTE de proposito: em Release, Core_Notify abre MsgBox e travaria a automacao. O .xlsx nao leva codigo.
#>
param(
    [Parameter(Mandatory)][string]$CoreXlam,
    [string]$OutFile = (Join-Path $PSScriptRoot '..\exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx'),
    [int]$TimeoutSec = 600
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\build\lib\Build-Common.ps1')
. (Join-Path $PSScriptRoot 'DemoCoverage.ps1')
if (-not (Test-Path $CoreXlam)) { throw "core de teste nao encontrado: $CoreXlam" }
$OutFile = [IO.Path]::GetFullPath($OutFile)
New-Item -ItemType Directory -Force (Split-Path $OutFile -Parent) | Out-Null

# ---- dados ficticios deterministicos (gerador congruencial; nada de pessoas, empresas reais ou dados pessoais)
$script:seed = 20261011
function Next-Rand { $script:seed = ([int64]$script:seed * 1103515245 + 12345) % 2147483648; return [double]$script:seed / 2147483648 }
$company = 'Casa & Obra Materiais (rede fictícia)'
$stores = @(
    @('Loja Vale Verde', 'Vale Verde', 'Sul'), @('Loja Porto Claro', 'Porto Claro', 'Sudeste'), @('Loja Serra Azul', 'Serra Azul', 'Sudeste'),
    @('Loja Rio Dourado', 'Rio Dourado', 'Norte'), @('Loja Campo Novo', 'Campo Novo', 'Centro-Oeste'), @('Loja Lagoa Alta', 'Lagoa Alta', 'Nordeste'),
    @('Loja Monte Bonito', 'Monte Bonito', 'Sul'), @('Loja Boa Vista', 'Boa Vista do Sul', 'Nordeste'))
$depts = @('Cimento e argamassa', 'Tintas', 'Elétrica', 'Hidráulica', 'Pisos e revestimentos', 'Ferramentas', 'Madeiras', 'Ferragens')
$deptGroup = @{ 'Cimento e argamassa' = 'Construção'; 'Madeiras' = 'Construção'; 'Tintas' = 'Acabamento'; 'Pisos e revestimentos' = 'Acabamento'; 'Elétrica' = 'Instalações'; 'Hidráulica' = 'Instalações'; 'Ferramentas' = 'Instalações'; 'Ferragens' = 'Acabamento' }
$months = @(); for ($m = 0; $m -lt 24; $m++) { $months += (Get-Date -Year 2024 -Month 1 -Day 1).AddMonths($m).Date }
$b2b = @{}; foreach ($s in $stores) { $b2b[$s[0]] = 0.2 + 0.3 * (Next-Rand) }
$rows = New-Object System.Collections.Generic.List[object]
foreach ($s in $stores) {
    foreach ($d in $depts) {
        $base = (60 + 140 * (Next-Rand)) * 1000
        if ($d -eq 'Cimento e argamassa') { $base = $base * 5 }
        for ($m = 0; $m -lt 24; $m++) {
            $season = 1 + 0.15 * [Math]::Sin(2 * [Math]::PI * ($m % 12) / 12)
            $sales = [Math]::Round($base * $season * (1 + 0.012 * $m) * (0.9 + 0.2 * (Next-Rand)), 0)
            $cost = [Math]::Round($sales * (0.6 + 0.12 * (Next-Rand)), 0)
            $units = [Math]::Max(1, [Math]::Round($sales / (90 + 110 * (Next-Rand)), 0))
            $ticket = [Math]::Round($sales / [Math]::Max(1, [Math]::Round($units / (1.5 + (Next-Rand)), 0)), 2)
            $rows.Add(@($s[0], $s[1], $s[2], $d, $months[$m], $sales, $cost, ($sales - $cost), $units, $ticket, [Math]::Round($sales * $b2b[$s[0]], 0), [Math]::Round($sales * (1 - $b2b[$s[0]]), 0)))
        }
    }
}
# agregados (R$ mil) para as tabelas de exemplo
function Sum-Where([scriptblock]$pred, [int]$col) { $t = 0.0; foreach ($r in $rows) { if (& $pred $r) { $t += [double]$r[$col] } }; return [Math]::Round($t / 1000, 0) }
$deptYear = @{}; foreach ($d in $depts) { $deptYear[$d] = @((Sum-Where { param($r) $r[3] -eq $d -and $r[4].Year -eq 2024 } 5), (Sum-Where { param($r) $r[3] -eq $d -and $r[4].Year -eq 2025 } 5)) }
$storeYear = @{}; foreach ($s in $stores) { $storeYear[$s[0]] = @((Sum-Where { param($r) $r[0] -eq $s[0] -and $r[4].Year -eq 2024 } 5), (Sum-Where { param($r) $r[0] -eq $s[0] -and $r[4].Year -eq 2025 } 5)) }
$regions = @('Sul', 'Sudeste', 'Norte', 'Nordeste', 'Centro-Oeste')
$regionMix = @{}; foreach ($g in $regions) { $regionMix[$g] = @((Sum-Where { param($r) $r[2] -eq $g } 11), (Sum-Where { param($r) $r[2] -eq $g } 10)) }
$storeMix = @{}; foreach ($s in $stores) { $storeMix[$s[0]] = @((Sum-Where { param($r) $r[0] -eq $s[0] } 11), (Sum-Where { param($r) $r[0] -eq $s[0] } 10)) }
$s1 = $stores[0][0]
$store1Month = @(); for ($m = 12; $m -lt 24; $m++) { $mm = $months[$m]; $store1Month += ,@($mm, (Sum-Where { param($r) $r[0] -eq $s1 -and $r[4] -eq $mm } 5), (Sum-Where { param($r) $r[0] -eq $s1 -and $r[4] -eq $mm } 6)) }
$receita = Sum-Where { param($r) $r[4].Year -eq 2025 } 5
$custo = Sum-Where { param($r) $r[4].Year -eq 2025 } 6
$despesas = [Math]::Round($receita * 0.12, 0); $ebitda = $receita - $custo - $despesas; $impostos = [Math]::Round($ebitda * 0.09, 0)
$regionGroups = @{}; foreach ($g in $regions) { $regionGroups[$g] = @('Construção', 'Acabamento', 'Instalações') | ForEach-Object { $gg = $_; Sum-Where { param($r) $r[2] -eq $g -and $deptGroup[$r[3]] -eq $gg -and $r[4].Year -eq 2025 } 5 } }
$regionStores = @{}; foreach ($g in $regions) { $regionStores[$g] = @($stores | Where-Object { $_[2] -eq $g }).Count }
$storeStats = @{}; foreach ($s in $stores) { $v = Sum-Where { param($r) $r[0] -eq $s[0] -and $r[4].Year -eq 2025 } 5; $c = Sum-Where { param($r) $r[0] -eq $s[0] -and $r[4].Year -eq 2025 } 6; $tk = 0.0; $n = 0; foreach ($r in $rows) { if ($r[0] -eq $s[0] -and $r[4].Year -eq 2025) { $tk += [double]$r[9]; $n++ } }; $storeStats[$s[0]] = @([Math]::Round($tk / $n, 2), [Math]::Round(($v - $c) / $v, 3), $v) }
$deptStats = @{}; foreach ($d in $depts) { $v = Sum-Where { param($r) $r[3] -eq $d -and $r[4].Year -eq 2025 } 5; $c = Sum-Where { param($r) $r[3] -eq $d -and $r[4].Year -eq 2025 } 6; $tk = 0.0; $n = 0; foreach ($r in $rows) { if ($r[3] -eq $d -and $r[4].Year -eq 2025) { $tk += [double]$r[9]; $n++ } }; $deptStats[$d] = @([Math]::Round($tk / $n, 2), [Math]::Round(($v - $c) / $v, 3), $v) }
$obrasQ = @(@('Reforma Loja Vale Verde', [datetime]'2026-03-02', [datetime]'2026-03-27'), @('Troca de telhado Porto Claro', [datetime]'2026-03-16', [datetime]'2026-04-24'), @('Novo depósito Serra Azul', [datetime]'2026-04-06', [datetime]'2026-05-15'), @('Inauguração do depósito', [datetime]'2026-05-18', $null))
$obrasY = @(@('Projeto da loja Campo Alegre', [datetime]'2026-01-05', [datetime]'2026-03-27'), @('Obra civil', [datetime]'2026-04-06', [datetime]'2026-12-18'), @('Instalações e acabamento', [datetime]'2027-01-04', [datetime]'2027-04-30'), @('Montagem e estoque', [datetime]'2027-05-03', [datetime]'2027-06-11'), @('Inauguração', [datetime]'2027-06-15', $null))

# ---- Excel
$xl = New-ExcelApp
$wd = Start-ExcelWatchdog $xl $TimeoutSec
try {
    $xl.EnableEvents = $false
    $core = $xl.Workbooks.Open($CoreXlam)
    $run = "'$($core.Name)'!"
    function Core([string]$proc) { $a = @($args); switch ($a.Count) { 0 { return $xl.Run($run + $proc) } 1 { return $xl.Run($run + $proc, $a[0]) } 2 { return $xl.Run($run + $proc, $a[0], $a[1]) } default { return $xl.Run($run + $proc, $a[0], $a[1], $a[2]) } } }
    function Assert-NoNotice([string]$ctx) { $n = Core 'Core_LastNotice'; if ($n) { throw "aviso do core em ${ctx}: $n" } }
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $idx = $wb.Worksheets.Item(1); $idx.Name = 'Indice'
    # Dados
    $wsD = $wb.Worksheets.Add([Type]::Missing, $idx); $wsD.Name = 'Dados'
    $wsD.Range('A1:L1').Value2 = [object[]]('Loja', 'Cidade', 'Região', 'Departamento', 'Mês', 'Vendas', 'Custo', 'Margem', 'Unidades', 'Ticket médio', 'Vendas B2B', 'Vendas B2C')
    $arr = New-Object 'object[,]' $rows.Count, 12
    for ($i = 0; $i -lt $rows.Count; $i++) { for ($j = 0; $j -lt 12; $j++) { $arr[$i, $j] = $rows[$i][$j] } }
    $wsD.Range('A2').Resize($rows.Count, 12).Value2 = $arr
    $wsD.Range('E2').Resize($rows.Count, 1).NumberFormat = 'mmm/yyyy'
    $wsD.Range('F2').Resize($rows.Count, 7).NumberFormat = '#,##0'
    $wsD.Range('J2').Resize($rows.Count, 1).NumberFormat = '#,##0.00'
    $wsD.Range('A1:L1').Font.Bold = $true; $wsD.Columns('A:L').AutoFit() | Out-Null
    # Lojas e Obras
    $wsL = $wb.Worksheets.Add([Type]::Missing, $wsD); $wsL.Name = 'Lojas'
    $wsL.Range('A1:E1').Value2 = [object[]]('Loja', 'Cidade', 'Região', 'Meta 2025 (R$ mil)', 'Inauguração')
    for ($i = 0; $i -lt $stores.Count; $i++) { $s = $stores[$i]; $wsL.Range("A$($i + 2):E$($i + 2)").Value2 = [object[]]($s[0], $s[1], $s[2], [Math]::Round($storeYear[$s[0]][1] * 1.08, 0), (Get-Date -Year (2012 + $i) -Month (1 + 3 * ($i % 4)) -Day 10).Date) }
    $wsL.Range('E2:E9').NumberFormat = 'dd/mm/yyyy'; $wsL.Range('A1:E1').Font.Bold = $true; $wsL.Columns('A:E').AutoFit() | Out-Null
    $wsO = $wb.Worksheets.Add([Type]::Missing, $wsL); $wsO.Name = 'Obras'
    $wsO.Range('A1:C1').Value2 = [object[]]('Obra', 'Início', 'Fim')
    $r = 2; foreach ($o in ($obrasQ + $obrasY)) { $wsO.Cells.Item($r, 1).Value2 = $o[0]; $wsO.Cells.Item($r, 2).Value2 = $o[1]; if ($o[2]) { $wsO.Cells.Item($r, 3).Value2 = $o[2] }; $r++ }
    $wsO.Range('B2:C20').NumberFormat = 'dd/mm/yyyy'; $wsO.Range('A1:C1').Font.Bold = $true; $wsO.Columns('A:C').AutoFit() | Out-Null

    # ---- um exemplo por item
    function Write-Table($ws, [string]$topLeft, [object[]]$header, [object[][]]$body, [string]$numFmt = '#,##0') {
        $ws.Range($topLeft).Resize(1, $header.Count).Value2 = $header
        $ws.Range($topLeft).Resize(1, $header.Count).Font.Bold = $true
        for ($i = 0; $i -lt $body.Count; $i++) { $ws.Range($topLeft).Offset($i + 1, 0).Resize(1, $body[$i].Count).Value2 = [object[]]$body[$i] }
        $rng = $ws.Range($topLeft).Resize($body.Count + 1, $header.Count)
        if ($numFmt) { $rng.Offset(1, 1).Resize($body.Count, $header.Count - 1).NumberFormat = $numFmt }
        return ,$rng                                   # a virgula: sem ela o PowerShell desenrola o Range celula por celula
    }
    function New-ExampleSheet($after, $item) {
        $ws = $wb.Worksheets.Add([Type]::Missing, $after)
        $ws.Name = $item.Sheet
        $ws.Hyperlinks.Add($ws.Range('A1'), '', "'Indice'!A1", 'Voltar ao índice', 'Voltar ao índice') | Out-Null
        $ws.Range('A3').Value2 = "$($item.Num)) $($item.Title) — Menu ThinQQ: $($item.Menu)"
        $ws.Range('A3').Font.Bold = $true; $ws.Range('A3').Font.Size = 12
        $ws.Columns('A').ColumnWidth = 24; $ws.Columns('B:G').ColumnWidth = 14
        return $ws
    }
    function Insert-Chart($ws, $rng, [string]$kind) {
        $ws.Activate(); $rng.Select() | Out-Null
        Core 'Core_ClearLastNotice' | Out-Null
        Core 'Core_Insert' $kind | Out-Null
        Assert-NoNotice "Core_Insert $kind em $($ws.Name)"
        if ($ws.ChartObjects().Count -lt 1) { throw "sem grafico em $($ws.Name)" }
    }
    $prev = $wsO
    $deptBody = @($depts | ForEach-Object { ,@($_, $deptYear[$_][0], $deptYear[$_][1]) })
    $storeBody = @($stores | ForEach-Object { ,@($_[0], $storeYear[$_[0]][0], $storeYear[$_[0]][1]) })
    $regionMixBody = @($regions | ForEach-Object { ,@($_, $regionMix[$_][0], $regionMix[$_][1]) })
    $storeMixBody = @($stores | ForEach-Object { ,@($_[0], $storeMix[$_[0]][0], $storeMix[$_[0]][1]) })
    $monthBody = @($store1Month | ForEach-Object { ,@($_[0], $_[1], $_[2]) })
    $wfBody = @(@('Receita 2025', $receita), @('Custo das mercadorias', -$custo), @('Despesas', -$despesas), @('EBITDA', 'e'), @('Impostos', -$impostos), @('Lucro', 'e'))
    $threeDepts = @('Cimento e argamassa', 'Tintas', 'Elétrica')
    $monthDeptBody = @(); for ($m = 12; $m -lt 24; $m++) { $mm = $months[$m]; $monthDeptBody += ,@($mm, (Sum-Where { param($r) $r[3] -eq $threeDepts[0] -and $r[4] -eq $mm } 5), (Sum-Where { param($r) $r[3] -eq $threeDepts[1] -and $r[4] -eq $mm } 5), (Sum-Where { param($r) $r[3] -eq $threeDepts[2] -and $r[4] -eq $mm } 5)) }
    $comboBody = @($store1Month | ForEach-Object { ,@($_[0], $_[1], $_[2], [Math]::Round(($_[1] - $_[2]) / $_[1], 3)) })
    $mekkoBody = @($regions | ForEach-Object { ,@($_, $regionGroups[$_][0], $regionGroups[$_][1], $regionGroups[$_][2]) })
    $mekkoUnitsBody = @($regions | ForEach-Object { ,@($_, $regionStores[$_], $regionGroups[$_][0], $regionGroups[$_][1], $regionGroups[$_][2]) })
    $pieBody = @($depts | ForEach-Object { ,@($_, $deptYear[$_][1]) })
    $scatterBody = @($stores | ForEach-Object { ,@($_[0], $storeStats[$_[0]][0], $storeStats[$_[0]][1], $_[2]) })
    $bubbleBody = @($depts | ForEach-Object { ,@($_, $deptStats[$_][0], $deptStats[$_][1], $deptStats[$_][2], $deptGroup[$_]) })
    $shareBody = @($regions | ForEach-Object { $t = $regionGroups[$_][0] + $regionGroups[$_][1] + $regionGroups[$_][2]; ,@($_, $regionGroups[$_][0], $regionGroups[$_][1], $regionGroups[$_][2]) })
    $breakBody = @($depts | ForEach-Object { ,@($_, $deptYear[$_][1]) })
    foreach ($item in (Get-TqqDemoCoverage)) {
        $ws = New-ExampleSheet $prev $item
        $prev = $ws
        switch ($item.Num) {
            1  { $rng = Write-Table $ws 'A5' @('Departamento', '2024', '2025') $deptBody
                 for ($i = 0; $i -lt $depts.Count; $i++) { foreach ($c in 0, 1) { $ws.Cells.Item(6 + $i, 2 + $c).Formula = "=ROUND(SUMIFS(Dados!`$F:`$F,Dados!`$D:`$D,`$A$(6 + $i),Dados!`$E:`$E,`">=`"&DATE($(2024 + $c),1,1),Dados!`$E:`$E,`"<`"&DATE($(2025 + $c),1,1))/1000,0)" } }
                 Insert-Chart $ws $rng 'stacked' }
            2  { $rng = Write-Table $ws 'A5' @('Loja', '2024', '2025') $storeBody; Insert-Chart $ws $rng 'clustered'; Core 'Core_SetOption' 'gap' '120' | Out-Null }
            3  { $rng = Write-Table $ws 'A5' @('Região', 'B2C', 'B2B') $regionMixBody; Insert-Chart $ws $rng 'stacked100' }
            4  { $rng = Write-Table $ws 'A5' @('Mês', 'Vendas', 'Custo') $monthBody; $ws.Range('A6:A17').NumberFormat = 'dd/mm/yyyy'; $ws.Range('A6:A17').NumberFormat = 'mmm/yy'; Insert-Chart $ws $rng 'line'; Core 'Core_SetOption' 'legend' 'lineend' | Out-Null }
            5  { $rng = Write-Table $ws 'A5' @('Etapa', 'R$ mil') $wfBody; Insert-Chart $ws $rng 'waterfall' }
            6  { $rng = Write-Table $ws 'A5' @('Região', '2024', '2025') @($regions | ForEach-Object { $g = $_; ,@($g, (Sum-Where { param($r) $r[2] -eq $g -and $r[4].Year -eq 2024 } 5), (Sum-Where { param($r) $r[2] -eq $g -and $r[4].Year -eq 2025 } 5)) }); Insert-Chart $ws $rng 'barstacked'; Core 'Core_AddAnnotation' 'diff' 1 $regions.Count | Out-Null }
            7  { $rng = Write-Table $ws 'A5' @('Loja', '2024', '2025') @($stores | ForEach-Object { ,@($_[0], $storeYear[$_[0]][0] * 1000, $storeYear[$_[0]][1] * 1000) }); Insert-Chart $ws $rng 'barclustered'; Core 'Core_SetOption' 'numfmt' 'mil' | Out-Null }
            8  { $rng = Write-Table $ws 'A5' @('Loja', 'B2C', 'B2B') $storeMixBody; Insert-Chart $ws $rng 'barstacked100' }
            9  { $rng = Write-Table $ws 'A5' @('Etapa', 'R$ mil') $wfBody; Insert-Chart $ws $rng 'barwaterfall' }
            10 { $rng = Write-Table $ws 'A5' @('Mês', $threeDepts[0], $threeDepts[1], $threeDepts[2]) $monthDeptBody; $ws.Range('A6:A17').NumberFormat = 'dd/mm/yyyy'; $ws.Range('A6:A17').NumberFormat = 'mmm/yy'; Insert-Chart $ws $rng 'area'; Core 'Core_SetToggle' 'vline' $true | Out-Null }
            11 { $rng = Write-Table $ws 'A5' @('Mês', $threeDepts[0], $threeDepts[1], $threeDepts[2]) $monthDeptBody; $ws.Range('A6:A17').NumberFormat = 'dd/mm/yyyy'; $ws.Range('A6:A17').NumberFormat = 'mmm/yy'; Insert-Chart $ws $rng 'area100' }
            12 { $rng = Write-Table $ws 'A5' @('Mês', 'Vendas', 'Custo', 'Margem %') $comboBody; $ws.Range('A6:A17').NumberFormat = 'dd/mm/yyyy'; $ws.Range('A6:A17').NumberFormat = 'mmm/yy'; $ws.Range('D6:D17').NumberFormat = '0.0%'; Insert-Chart $ws $rng 'combo'; if (-not (Core 'Core_GetToggle' 'secondary')) { throw 'E12: eixo secundario nao ligou' } }
            13 { $rng = Write-Table $ws 'A5' @('Região', 'Construção', 'Acabamento', 'Instalações') $mekkoBody; Insert-Chart $ws $rng 'mekko' }
            14 { $rng = Write-Table $ws 'A5' @('Região', 'Lojas', 'Construção', 'Acabamento', 'Instalações') $mekkoUnitsBody; Insert-Chart $ws $rng 'mekkounits' }
            15 { $rng = Write-Table $ws 'A5' @('Obra', 'Início', 'Fim') @($obrasQ | ForEach-Object { ,@($_[0], $_[1], $_[2]) }) ''; $ws.Range('B6:C9').NumberFormat = 'dd/mm/yyyy'; Insert-Chart $ws $rng 'gantt'; Core 'Core_AddAnnotation' 'bracket' 1 2 | Out-Null }
            16 { $rng = Write-Table $ws 'A5' @('Obra', 'Início', 'Fim') @($obrasY | ForEach-Object { ,@($_[0], $_[1], $_[2]) }) ''; $ws.Range('B6:C10').NumberFormat = 'dd/mm/yyyy'; Insert-Chart $ws $rng 'gantt' }
            17 { $rng = Write-Table $ws 'A5' @('Departamento', 'Vendas 2025') $pieBody; Insert-Chart $ws $rng 'pie'; Core 'Core_SetToggle' 'pieout' $true | Out-Null }
            18 { $rng = Write-Table $ws 'A5' @('Departamento', 'Vendas 2025') $pieBody; Insert-Chart $ws $rng 'doughnut'; Core 'Core_SetOption' 'palette' 'cinza' | Out-Null }
            19 { $rng = Write-Table $ws 'A5' @('Loja', 'B2C', 'B2B') $storeMixBody; Insert-Chart $ws $rng 'butterfly' }
            20 { $rng = Write-Table $ws 'A5' @('Loja', 'Ticket médio', 'Margem', 'Região') $scatterBody '#,##0.00'; $ws.Range('C6:C13').NumberFormat = '0.0%'; Insert-Chart $ws $rng 'scatter'; Core 'Core_SetToggle' 'trend' $true | Out-Null; Core 'Core_SetToggle' 'quad' $true | Out-Null }
            21 { $rng = Write-Table $ws 'A5' @('Departamento', 'Ticket médio', 'Margem', 'Vendas 2025', 'Grupo') $bubbleBody '#,##0.00'; $ws.Range('C6:C13').NumberFormat = '0.0%'; $ws.Range('D6:D13').NumberFormat = '#,##0'; Insert-Chart $ws $rng 'bubble' }
            22 { $rng = Write-Table $ws 'A5' @('Ano', $threeDepts[0], $threeDepts[1], $threeDepts[2]) @(@(2024, 0, 0, 0), @(2025, 0, 0, 0))
                 for ($y = 0; $y -lt 2; $y++) { for ($c = 0; $c -lt 3; $c++) { $ws.Cells.Item(6 + $y, 2 + $c).Formula = "=ROUND(SUMIFS(Dados!`$F:`$F,Dados!`$D:`$D,`"$($threeDepts[$c])`",Dados!`$E:`$E,`">=`"&DATE($(2024 + $y),1,1),Dados!`$E:`$E,`"<`"&DATE($(2025 + $y),1,1))/1000,0)" } }
                 Insert-Chart $ws $rng 'stacked'; Core 'Core_AddAnnotation' 'cagr' 1 2 | Out-Null }
            23 { $rng = Write-Table $ws 'A5' @('Região', 'Construção', 'Acabamento', 'Instalações') $shareBody; Insert-Chart $ws $rng 'stacked100'; Core 'Core_SetToggle' 'roundsum' $true | Out-Null
                 $ws.Range('A14').Value2 = 'Arredondar intervalo: participações da região Sul (2 casas) que somam exatamente 100%'; $ws.Range('A14').Font.Italic = $true
                 $ws.Range('A15:B15').Value2 = [object[]]('Grupo', 'Participação'); $ws.Range('A15:B15').Font.Bold = $true
                 $t = $shareBody[0][1] + $shareBody[0][2] + $shareBody[0][3]
                 $ws.Range('A16:B16').Value2 = [object[]]('Construção', $shareBody[0][1] / $t)      # uma linha por vez: um array de arrays nao e matriz 2-D
                 $ws.Range('A17:B17').Value2 = [object[]]('Acabamento', $shareBody[0][2] / $t)
                 $ws.Range('A18:B18').Value2 = [object[]]('Instalações', $shareBody[0][3] / $t)
                 $ws.Range('B16:B18').NumberFormat = '0.000%'; $ws.Range('C15').Value2 = 'Soma consistente'; $ws.Range('C15').Font.Bold = $true
                 Core 'Round_WriteFormulas' $ws.Range('B16:B18') $ws.Range('C16') 2 | Out-Null; $ws.Range('C16:C18').NumberFormat = '0%' }
            24 { $rng = Write-Table $ws 'A5' @('Departamento', 'Vendas 2025') $breakBody; Insert-Chart $ws $rng 'stacked'
                 $sorted = @($breakBody | ForEach-Object { [double]$_[1] } | Sort-Object -Descending); Core 'Core_SetAxisBreakAt' ([double]([Math]::Round($sorted[1] * 1.2, 0))) | Out-Null; Assert-NoNotice 'corte de eixo' }
            25 { $rng = Write-Table $ws 'A5' @('Ano', 'Vendas', 'Custo') @(@(2024, $storeYear[$stores[0][0]][0], (Sum-Where { param($r) $r[0] -eq $stores[0][0] -and $r[4].Year -eq 2024 } 6)), @(2025, $storeYear[$stores[0][0]][1], (Sum-Where { param($r) $r[0] -eq $stores[0][0] -and $r[4].Year -eq 2025 } 6)))
                 $ws.Range('A4').Value2 = $stores[0][0]; Insert-Chart $ws $rng 'clustered'; $co1 = $ws.ChartObjects().Item(1); $co1.Width = 300; $co1.Height = 200
                 $rng2 = Write-Table $ws 'A22' @('Ano', 'Vendas', 'Custo') @(@(2024, $storeYear[$stores[3][0]][0], (Sum-Where { param($r) $r[0] -eq $stores[3][0] -and $r[4].Year -eq 2024 } 6)), @(2025, $storeYear[$stores[3][0]][1], (Sum-Where { param($r) $r[0] -eq $stores[3][0] -and $r[4].Year -eq 2025 } 6)))
                 $ws.Range('A21').Value2 = $stores[3][0]; Insert-Chart $ws $rng2 'clustered'; $co2 = $ws.ChartObjects().Item(2); $co2.Width = 300; $co2.Height = 200; $co2.Left = $co1.Left + 320; $co2.Top = $co1.Top
                 Core 'Core_ClearLastNotice' | Out-Null; Core 'Core_SameScale' | Out-Null; if ((Core 'Core_LastNotice') -notmatch 'aplicada a 2') { throw "mesma escala: $(Core 'Core_LastNotice')" } }
            26 { $rng = Write-Table $ws 'A5' @('Mês', 'Vendas') @($store1Month | ForEach-Object { ,@($_[0], $_[1]) }); $ws.Range('A6:A17').NumberFormat = 'dd/mm/yyyy'; $ws.Range('A6:A17').NumberFormat = 'mmm/yy'; Insert-Chart $ws $rng 'line'; Core 'Core_SetOption' 'legend' 'none' | Out-Null }
        }
        Core 'Core_ClearLastNotice' | Out-Null
    }
    # ---- indice
    $idx.Activate()
    $idx.Range('A1').Value2 = "ThinQQ — pasta de demonstração: $company"; $idx.Range('A1').Font.Bold = $true; $idx.Range('A1').Font.Size = 14
    $idx.Range('A2').Value2 = 'Dados fictícios, gerados por tools/New-DemoWorkbook.ps1. Clique na funcionalidade para ir ao exemplo; cada exemplo tem "Voltar ao índice" em A1. Os gráficos são gráficos nativos do Excel: abra esta pasta em qualquer Excel, com ou sem o ThinQQ.'
    $idx.Range('A4:D4').Value2 = [object[]]('Nº', 'Funcionalidade', 'Tipo de gráfico', 'Descrição'); $idx.Range('A4:D4').Font.Bold = $true
    $r = 5
    foreach ($item in (Get-TqqDemoCoverage)) {
        $idx.Cells.Item($r, 1).Value2 = $item.Num
        $idx.Hyperlinks.Add($idx.Cells.Item($r, 2), '', "'$($item.Sheet)'!A1", $item.Title, $item.Title) | Out-Null
        $idx.Cells.Item($r, 3).Value2 = $item.Kind
        $idx.Cells.Item($r, 4).Value2 = "$($item.Desc). Menu: $($item.Menu)."
        $r++
    }
    $idx.Columns('A').ColumnWidth = 5; $idx.Columns('B').ColumnWidth = 70; $idx.Columns('C').ColumnWidth = 16; $idx.Columns('D').ColumnWidth = 90
    $idx.Range('A5').Select() | Out-Null
    if (Test-Path $OutFile) { Remove-Item $OutFile -Force }
    $wb.RemovePersonalInformation = $true             # docProps sem autor (nunca altere $xl.UserName: grava no registro do Office do usuario)
    $wb.SaveAs($OutFile, 51)
    $wb.Close($false)
    $core.Close($false)
    Write-Host "Gerado: $OutFile ($((Get-Item $OutFile).Length) bytes)"
} finally { Stop-ExcelWatchdog $wd; Close-ExcelApp $xl }
```
(Notas para o implementador: `Write-Table` recebe `$body` como `object[][]`; dentro de `@(... | ForEach-Object { ,@(...) })` e em `+= ,@(...)` a vírgula unária preserva cada linha como array, mas num literal de várias linhas **não** se põe vírgula antes da primeira (`@(@(a), @(b))`; `@(,@(a), @(b))` embrulha só a primeira linha num array a mais e a célula receberia um array); funções que devolvem um `Range` usam `return ,$rng` (o PowerShell desenrola coleções COM: sem a vírgula, `$rng.Select()` selecionaria célula por célula e o `Core_Insert` leria a região atual da última célula, que no E25 inclui o nome da loja em `A4`/`A21`); `Range.Value2` de várias linhas exige matriz 2-D: grave linha a linha; datas em `mmm/yy` passam antes por `dd/mm/yyyy` (de `General` direto, o rótulo ligado mostra o mês em inglês: armadilha medida na Fase 2); os dois `.ps1` de `tools/` desta tarefa vão em UTF-8 **com BOM** (têm acentos; sem BOM o Windows PowerShell 5.1 os lê como ANSI); `Core` encaminha até 3 argumentos, inclusive `Range`; `Core_SameScale` sem seleção múltipla age em todos os gráficos ThinQQ da planilha ativa; a tabela do exemplo 1 e 22 usa `SUMIFS` sobre `Dados` para o gráfico ser vivo; os demais são valores.)

- [ ] **Passo 4: Gerar e testar** — na cópia: `pwsh C:\temp\tqq-f3\build\build.ps1 -Target core -Configuration Test`; depois, **a partir do repositório real** (o gerador grava em `exemplos\`): `pwsh <repo>\tools\New-DemoWorkbook.ps1 -CoreXlam C:\temp\tqq-f3\dist\ThinQQ.core.test.xlam` → `Gerado: ...xlsx (N bytes)`; abra o `.xlsx` **numa instância própria** (ou exporte alguns gráficos para PNG) e olhe o índice, os hiperlinks, 4 ou 5 exemplos; `robocopy` de novo e `pwsh C:\temp\tqq-f3\tests\ps\Test-DemoWorkbook.ps1 -RepoRoot C:\temp\tqq-f3` → `Test-DemoWorkbook: OK`. Em `build/run-tests.ps1`, depois do bloco da assinatura:

```powershell
if (-not $timedOut) {
    Write-Host '=== demonstracao (exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx) ==='
    $global:LASTEXITCODE = 0
    & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $RepoRoot 'tests\ps\Test-DemoWorkbook.ps1') -RepoRoot $RepoRoot
    if ($LASTEXITCODE -ne 0) { $totalFailed += 1 }
}
```

- [ ] **Passo 5: GREEN e suíte completa** (a suíte passa a abrir mais um Excel próprio por ~20 s; total da suíte ~9–10 min: o `-TimeoutSec 900` basta, registre o tempo no relatório).

- [ ] **Passo 6: Commit** — `git add tools/DemoCoverage.ps1 tools/New-DemoWorkbook.ps1 tests/ps/Test-DemoWorkbook.ps1 exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx build/run-tests.ps1` e `git commit -m "feat: pasta de demonstracao (materiais de construcao) gerada por script, com indice por funcionalidade e teste"` (+ trailer).

**Verificação no Excel real:** abrir `exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx`; clicar em alguns itens do índice e voltar; com o ThinQQ instalado, clicar num gráfico e ver a faixa reconhecê-lo; sem o ThinQQ, os gráficos continuam vivos (mudar um número em `Dados` muda o exemplo 1).

---
### Tarefa 11: Tutorial, README (instalação completa e TI), CHANGELOG, versões 1.2.0 e ensaio de release — modelo: standard

**Files:**
- Modify: `docs/tutorial/tutorial.html`, `README.md`, `CHANGELOG.md`, `version.json`, `release/notes.txt`, `installer/LEIA-ME.txt`, `tests/ps/Test-Docs.ps1`
- Test: `tests/ps/Test-Docs.ps1` (via `run-tests.ps1`), ensaio `build/release.ps1` na cópia

**Interfaces:**
- Consumes: rótulos exatos do `customUI14.xml` (Tarefa 8); mensagens exatas das tarefas 1–5 (6011, 6012, 6013, 7003, `PPT_ERR_*`, `PPT_NOTE_DIRTY`, "Fórmulas de arredondamento..."); scripts e resultado da Tarefa 7; `exemplos/*.xlsx` e `Test-DemoWorkbook` (Tarefa 10); 26 exemplos (Tarefa 9).
- Produces: documentação coerente com o código, versões `1.2.0` (loader e core), `Test-Docs` ampliado, release de ensaio verde.

- [ ] **Passo 1: Ampliar `Test-Docs.ps1` (RED)** — acrescente (dentro do `if ($parsed)` as que usam `$xml`):

```powershell
foreach ($id in 'borboleta', 'dispersao', 'arredondamento', 'corte', 'powerpoint', 'instalacao-ti') {
    Check ($null -ne $xml.SelectSingleNode("//*[local-name()='section'][@id='$id']")) "tutorial tem a seção #$id"
    Check ($null -ne $xml.SelectSingleNode("//nav//a[@href='#$id']")) "o índice leva à seção #$id"
}
Check ($readme -match '(?m)^## Instalação') 'README tem a seção "Instalação"'
foreach ($t in 'Instalar-ThinQQ.cmd', 'Desbloquear', 'Central de Confiabilidade', 'Local Confiável', 'antivírus', 'Desinstala', 'Instalação manual', 'Releases') { Check ($readme -match [regex]::Escape($t)) "README/Instalação cita '$t'" }
Check ($readme -match '(?m)^### Ambientes corporativos \(TI\)') 'README tem "Ambientes corporativos (TI)"'
foreach ($t in 'Editores Confiáveis', 'tools/New-ThinQQCert.ps1', 'tools/Sign-ThinQQ.ps1', 'Assinatura Digital', 'Subject Interface Packages') { Check ($readme -match [regex]::Escape($t)) "README/TI cita '$t'" }
Check ($readme -match 'ThinQQ-Demonstracao-Materiais-Construcao\.xlsx') 'README cita a pasta de demonstração'
Check ($cl -match '(?m)^## \[1\.2\.0\] - \d{4}-\d{2}-\d{2}') 'CHANGELOG tem a seção 1.2.0'
Check (((Read-Utf8 (Join-Path $RepoRoot 'version.json')) | ConvertFrom-Json).loader -eq '1.2.0') 'loader 1.2.0'
```
O limite de 160 KB do tutorial já foi trocado na Tarefa 0 (confira; não duplique a verificação). Na lista `$messages` acrescente: `'A borboleta precisa de exatamente 2 colunas de valores (lado esquerdo e lado direito).'`, `'Dispersão: a tabela precisa de nome, X e Y (e, nas bolhas, tamanho).'`, `'No máximo 6 grupos (coluna de grupo).'`, `'Corte de eixo: informe um valor maior que zero e menor que o maior valor do gráfico.'`, `'O PowerPoint não está instalado neste computador (ou não pôde ser iniciado).'`, `'Salve a pasta de trabalho antes de copiar o gráfico vinculado: o vínculo aponta para o arquivo salvo.'`, `'A pasta tem alterações não salvas: o PowerPoint lê o arquivo salvo nas atualizações do vínculo.'`, `'Selecione uma única coluna ou linha só com números (pelo menos 2).'`, `'Fórmulas de arredondamento com soma consistente gravadas em'`. A verificação `Samples_Create tem os exemplos numerados até $nEx` passa a exigir `-ge 26` e `tutorial e README citam 26 exemplos`. Rode `pwsh C:\temp\tqq-f3\tests\ps\Test-Docs.ps1 -RepoRoot C:\temp\tqq-f3` → `not ok` nas seções/mensagens/versões/instalação.

- [ ] **Passo 2: Tutorial** — conteúdo obrigatório (cada rótulo em `<span class="ui">` com o texto exato do XML; ids únicos; sem recursos externos; o único `<script>` inline existente; um `h1`):
  - "Neste tutorial": índice com as seções novas.
  - `#tipos`: as três linhas novas (Tarefa 8) com link para `#borboleta`/`#dispersao`.
  - `#borboleta`: tabela de exemplo (Faixa/Mulheres/Homens), o que o gráfico mostra (lado esquerdo em espelho, nomes no centro, totais à direita, mesma escala dos dois lados), limites (exatamente 2 colunas; negativos/texto valem 0; nomes muito longos reduzem as barras), mensagem 6011, controles que valem (`Totais`, `Rótulos`, `Formato`, `Paleta`, `Espaço`, `Legenda`, `Mesma escala`) e os que ficam cinza.
  - `#dispersao`: tabelas de exemplo (dispersão com `Grupo`; bolhas com `Volume`), uma série por grupo (máximo 6; mensagem 6013; **um grupo novo criado depois precisa de `Alterar dados`**), nomes por ponto, títulos dos eixos ligados aos cabeçalhos, `Linha de tendência`, `Quadrantes (médias)` (as linhas dos eixos cruzam nas médias; sem o ThinQQ ficam na última posição), tamanho negativo/texto fica fora, mensagem 6012, `Formato`/`Espaço`/`Totais` cinza, rótulos sobrepostos em nuvens densas.
  - `#arredondamento`: o problema (33 + 33 + 33 = 99), `Rótulos com soma consistente` (onde vale: colunas/barras empilhadas e 100%, áreas; como funciona: maior resto em fórmulas da planilha; o total mostrado = soma dos rótulos; `Formato` muda as casas), `Arredondar intervalo...` (passo a passo com as duas perguntas e a frase de confirmação; fórmulas só com a faixa de origem; a mensagem de validação), por que não há função `TQQ.ROUND` (vínculo externo ao `.xlam` sem o ThinQQ).
  - `#corte`: `Corte de eixo...` (sugestão, o que acontece: colunas cortadas, marca, total real, rótulo do segmento atravessado na base, segmentos acima sem rótulo), `Remover corte de eixo`, limites (não é compressão de escala; `Mesma escala` e corte não convivem; `Escala automática` não remove o corte; sem o ThinQQ a marca e os rótulos de segmento ficam como estavam), mensagem 7003, tipos (colunas/barras empilhadas e agrupadas).
  - `#powerpoint`: `PowerPoint` › `Copiar como imagem` e `Copiar como gráfico vinculado` (o que cada um cola; slide ativo ou apresentação nova; a pasta precisa estar salva para o vínculo; as mensagens `PPT_ERR_NOT_INSTALLED`, `PPT_ERR_UNSAVED` e `PPT_NOTE_DIRTY` citadas literalmente; no PowerPoint, Dados > Editar Vínculos e "Quebrar vínculo" para incorporar; renomear/mover o arquivo quebra o vínculo).
  - `#instalacao-ti`: o mesmo roteiro da subseção "Ambientes corporativos (TI)" do README (certificado interno, assinatura V3 pelo VBE, `signtool` + Office SIPs, `tools/New-ThinQQCert.ps1`, `tools/Trust-ThinQQPublisher.ps1`, `tools/Sign-ThinQQ.ps1`, Editores Confiáveis por GPO, local confiável, política "Only trust VBA macros that use V3 signatures", reassinar a cada release).
  - `#matriz`: já com 20 × 20 (Tarefas 0 e 8); "Opções do gráfico selecionado": linhas dos controles novos; "Limitações": acrescentar os limites acima e os do README (área de transferência no PowerPoint; borboleta com `Mesma escala`); "Como funciona": vão central da borboleta (`GapK`), séries por grupo com `NA()`, colunas acrescentadas à direita do bloco (soma consistente, corte), `LblPos`; "Exemplos": 26 exemplos em três planilhas ("Exemplos", "Exemplos 2" e "Exemplos 3", já trocado na Tarefa 9) e a frase das planilhas de Exemplos com "17 dos 20 tipos" (eram 14 dos 17); as outras menções a "17 tipos" (legenda da maquete da faixa e "reúnem 17 tipos" na introdução) passam a 20; rodapé `ThinQQ 1.2.0.`
  - Renderize em Edge headless (como nas fases anteriores) em 1200 px e 500 px e corrija quebras.

- [ ] **Passo 3: README** — seções novas/alteradas:
  - `## Instalação` (para quem não é técnico; substitui a "Instalação (usuários)" atual): **Requisitos** (Excel do Office 16.0 para Windows: 2016, 2019, 2021 ou Microsoft 365; sem Mac/web; sem administrador); **Passo a passo**: 1) baixar `ThinQQ-x.y.z.zip` na página de Releases; 2) descompactar a pasta inteira (botão direito > Extrair tudo); 3) fechar o Excel; 4) dois cliques em `Instalar-ThinQQ.cmd` (se o Windows perguntar, "Executar"); 5) abrir o Excel e procurar a aba ThinQQ ao lado de Inserir; **Como as atualizações chegam** (aviso ao abrir o Excel, no máximo 1×/dia; "Atualizar agora"; ThinQQ > Verificar atualizações); **Desinstalar** (comando do `uninstall.ps1`; o que ele remove); **Instalação manual** (os 5 passos do LEIA-ME); **Problemas mais comuns** (5 itens, cada um com "o que você vê / o que fazer"): macros bloqueadas ou arquivo bloqueado pelo Windows (Propriedades > **Desbloquear**), Central de Confiabilidade (macros desabilitadas / aviso de segurança), Local Confiável (`%LOCALAPPDATA%\ThinQQ` com subpastas), antivírus/proxy (download bloqueado, quarentena do `.xlam`), a aba não aparece (fechar todas as janelas do Excel, reinstalar, ThinQQ > Sobre).
  - `### Ambientes corporativos (TI)`: o resultado da Fase 3 (§4.5 da spec): o que a TI pode fazer (certificado interno de assinatura de código; assinar os `.xlam` pelo VBE (Ferramentas > Assinatura Digital, assinatura V3) ou por `signtool` + Office Subject Interface Packages num servidor de build com `tools/Sign-ThinQQ.ps1`; distribuir o certificado em **Editores Confiáveis** por GPO (`tools/Trust-ThinQQPublisher.ps1` para um usuário); ou local confiável corporativo; reassinar a cada release e publicar os SHA-256 dos arquivos assinados) e o que o repositório entrega (`tools/New-ThinQQCert.ps1` para um certificado autoassinado de teste; `release.ps1 -SignPfx`); limites (sem admin/SDK a assinatura automática não roda: código 2).
  - "Recursos" → bloco **Fase 3 (1.2.0)**; "Galeria" com Borboleta/Dispersão/Bolhas; "Desenvolvimento" com `tools/` e os testes novos; "Antes de distribuir" com `v1.2.0`; "Verificações manuais no Excel real" com os itens das tarefas 1–10 (numerados a partir de 27; inclua o item da pasta de demonstração e o dos ícones próprios; confira o item 25, trocado na Tarefa 9 para três planilhas e "Exemplos 3"); "Limitações conhecidas" atualizadas (inclua: `Copiar para o PowerPoint` usa a área de transferência do Windows e substitui o que estava nela; `Mesma escala` entre uma borboleta e gráficos de outro tipo tira a borboleta do centro (use só entre borboletas); gráfico com corte de eixo fica fora da `Mesma escala`); "Estrutura" com `tools/` e `exemplos/`; links para `design.md`/plano da Fase 3 (caminhos finais em `docs/superpowers/...`, a cargo do controlador).
  - `CHANGELOG.md`: `## [1.2.0] - <data>` com **Adicionado** (borboleta; dispersão e bolhas; rótulos com soma consistente e Arredondar intervalo; corte de eixo; PowerPoint; ícones próprios; scripts de certificado/assinatura; pasta de demonstração; Exemplos 3; seção de instalação), **Alterado** (faixa; matriz 20 × 20; colunas Pos/VLine por cabeçalho; `release.ps1 -SignPfx`; `run-tests` com Test-Signing e Test-DemoWorkbook), **Corrigido** (o que surgir nas tarefas). `version.json` → `{ "loader": "1.2.0", "core": "1.2.0" }`. `release/notes.txt` → 6 a 8 linhas com `•`. `LEIA-ME.txt` → `versions\1.2.0\` e um parágrafo "Problemas comuns" alinhado ao README.

- [ ] **Passo 4: GREEN** — `pwsh C:\temp\tqq-f3\build\run-tests.ps1 -TimeoutSec 1200` → `Todos os testes passaram` (Test-Docs inclusive).

- [ ] **Passo 5: Ensaio de release (na cópia, nunca no repositório real)**

```powershell
robocopy "<repo>" C:\temp\tqq-f3 /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS
pwsh C:\temp\tqq-f3\build\release.ps1 -Version 1.2.0 -LoaderVersion 1.2.0 -BaseUrl C:\temp\tqq-f3-rel -PublishTo C:\temp\tqq-f3-rel
# esperado: "Release pronto: ...\ThinQQ-1.2.0.zip"; manifest.json com coreVersion/loaderVersion 1.2.0 e dois SHA-256
Expand-Archive C:\temp\tqq-f3-rel\ThinQQ-1.2.0.zip C:\temp\tqq-f3-rel-x
pwsh C:\temp\tqq-f3-rel-x\install.ps1 -Root C:\temp\tqq-f3-rel-root\ThinQQ -RegistryRoot HKCU:\Software\ThinQQ_Test_f3 -SkipExcelCheck
pwsh C:\temp\tqq-f3-rel-x\uninstall.ps1 -Root C:\temp\tqq-f3-rel-root\ThinQQ -RegistryRoot HKCU:\Software\ThinQQ_Test_f3 -SkipExcelCheck
Remove-Item HKCU:\Software\ThinQQ_Test_f3 -Recurse -ErrorAction SilentlyContinue
# ensaio da assinatura (diagnostico): esperado codigo 2 nesta maquina (sem signtool/SIP), release abortado com a mensagem do script
pwsh C:\temp\tqq-f3\tools\New-ThinQQCert.ps1 -Subject "CN=ThinQQ ensaio" -OutDir C:\temp\tqq-f3-rel\cert -PfxPassword x -NoTrustPublisher
pwsh C:\temp\tqq-f3\build\release.ps1 -Version 1.2.0 -LoaderVersion 1.2.0 -BaseUrl C:\temp\tqq-f3-rel -SignPfx C:\temp\tqq-f3-rel\cert\ThinQQ-signing.pfx -SignPassword x
# remover o certificado do ensaio (My e, por garantia, CA; o ensaio usou -NoTrustPublisher): comando, nao comentario
foreach ($s in 'My', 'CA') { Get-ChildItem "Cert:\CurrentUser\$s" | Where-Object Subject -eq 'CN=ThinQQ ensaio' | ForEach-Object { if ($s -eq 'My') { Remove-Item $_.PSPath -DeleteKey } else { Remove-Item $_.PSPath } } }
Remove-Item C:\temp\tqq-f3-rel, C:\temp\tqq-f3-rel-x, C:\temp\tqq-f3-rel-root -Recurse -Force
```
  Nada do ensaio é commitado.

- [ ] **Passo 6: Commit** — `git add docs/tutorial/tutorial.html README.md CHANGELOG.md version.json release/notes.txt installer/LEIA-ME.txt tests/ps/Test-Docs.ps1` e `git commit -m "docs: tutorial, README (instalacao e TI) e CHANGELOG da Fase 3; versao 1.2.0 do core e do loader"` (+ trailer).

**Verificação no Excel real:** ThinQQ > Tutorial abre o tutorial novo; ThinQQ > Sobre mostra Loader 1.2.0 / Core 1.2.0 após instalação pelo ZIP.

---

### Tarefa 12: Integração final e verificação — modelo: standard

**Files:**
- Nenhum arquivo de produção novo. Correções pequenas achadas aqui entram em commits `fix:` próprios, cada um com teste.

**Interfaces:** consome tudo.

- [ ] **Passo 1: Cópia limpa e suíte completa com troca de loader** — `robocopy "<repo>" C:\temp\tqq-f3 /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS` e `pwsh C:\temp\tqq-f3\build\run-tests.ps1 -LoaderSwap -TimeoutSec 1500`. Esperado: `Todos os testes passaram`, `loader-swap: ok`; anote os totais, o tempo e os 15 testes mais lentos; não remova testes.
- [ ] **Passo 2: Build Release (Png e Mso) e guarda** — `pwsh C:\temp\tqq-f3\build\build.ps1 -Target all -Configuration Release` → `OK dist\ThinQQ.xlam`, `OK dist\ThinQQ.core.xlam`; `pwsh C:\temp\tqq-f3\tests\ps\Test-ReleaseGuard.ps1`; depois `-Target loader -Configuration Release -RibbonIcons Mso` e `Test-Ribbon.ps1` contra o pacote de **teste** reconstruído com `-RibbonIcons Mso` (`build.ps1 -Target loader -Configuration Test -RibbonIcons Mso`) → verde; refaça o build de teste `Png`. Se `src/core` mudou depois do commit da Tarefa 10 (correções das Tarefas 11 e 12), regenere a demonstração (Tarefa 10, Passo 4), rode o `Test-DemoWorkbook` e commite o `.xlsx` novo.
- [ ] **Passo 3: Compatibilidade entre versões (sem automação contra o core velho)** — loader 1.2.0 com core 1.1.0 é coberto por `TestLoaderMain.Test_RunCore_MissingProc_ReturnsEmpty_NoRaise` e pelo fallback de `Rx_GetEnabled`; os botões novos chamam `Core_SetAxisBreak`/`Core_RoundRange`/`Core_CopyToPowerPoint` que, ausentes, devolvem `Empty` (só uma linha de log). Confirme verdes e registre. Não abra um core 1.1.0 Release por automação (MsgBox trava).
- [ ] **Passo 4: Auditoria de aderência** — `grep` nos fontes: `Declare ` (nenhum, inclusive em `modPowerPoint`), `MsgBox` fora de `Core_Notify`/`RxMsg`/`Loader_Msg`/formulários, `InputBox` só em `Core_RoundRange`, `Core_SetAxisBreak` e `Rx_OnSettings` (guardados por `BUILD_CONFIG = "Test"`) e em `Core_ChangeSource` (sem guarda desde a Fase 2; nenhum teste o chama; acrescentar a guarda é opcional, como `fix:` com teste); `Application.Run` nos testes (nenhum); `NumberFormat =` só com `dd/mm`, `mmm/yy`, `mmm/yyyy`, `dd/mm/yyyy`, `General`, `0%`, `0.0%`, `0.000%`, `#,##0`, `#,##0.00` (os três últimos só no gerador e nos exemplos: células de tabela); `CreateObject("PowerPoint.Application")` só em `modPowerPoint` e em `TestPowerPoint` (com as guardas de `Ppt_IsRunning` e do processo `POWERPNT`; em build `Test`, `Ppt_GetApp` só usa a instância injetada); `Cert:\` só em `tools/` e `Test-Signing.ps1`, nunca `Cert:\LocalMachine` fora de `Trust-ThinQQPublisher -Scope LocalMachine`; `Test_Core_AllPublicApi_HaveErrorContainment` continua verde com `Core_RoundRange`, `Core_SetAxisBreak`, `Core_SetAxisBreakAt`, `Core_RemoveAxisBreak`, `Core_CopyToPowerPoint`; `Test-Privacy.ps1` verde (nenhum caminho de usuário, PID ou e-mail nos arquivos de texto novos, inclusive `tools/`); o `.xlsx` gerado é conferido pelo `Test-DemoWorkbook` (docProps sem autor, nenhuma parte XML com id ou caminho de perfil; o `Test-Privacy` não lê binários; nomes fictícios).
- [ ] **Passo 5: Estado do repositório** — `git status` limpo; `git log --oneline <base>..HEAD` lista os commits da Fase 3; nenhum push; `dist/` e `tests/results/` ignorados; `exemplos/*.xlsx` versionado; `release/manifest.json` inexistente; nenhum certificado `ThinQQ Teste`/`ThinQQ ensaio` nas lojas do usuário (`Get-ChildItem Cert:\CurrentUser\My, Cert:\CurrentUser\CA, Cert:\CurrentUser\TrustedPublisher, Cert:\CurrentUser\Root | Where-Object { $_.Subject -like '*ThinQQ Teste*' -or $_.Subject -like '*ThinQQ ensaio*' }` vazio; um certificado com outro assunto ThinQQ não foi criado pelos testes: registre no relatório e não o remova); nenhum processo `POWERPNT` criado pelos testes; os processos Excel do usuário anotados no início continuam vivos.
- [ ] **Passo 6: Relatório final** — consolidar: contagens da suíte, decisões tomadas nas sondas das tarefas (vão da borboleta, quadrantes, corte, PowerPoint pulado ou real, assinatura diagnosticada), lista "minor (deferred)" de todos os relatórios, pastas temporárias removidas (`C:\temp\tqq-f3*`), e a lista numerada das verificações manuais no Excel real (a do README).

**Verificação no Excel real (usuário, após instalar pelo ZIP):** percorrer a lista "Verificações manuais no Excel real" do README na ordem; qualquer desvio volta como `fix:` com teste.

---

## Autorrevisão (feita ao escrever o plano)

- **Cobertura da spec:** borboleta ✓ (1); dispersão e bolhas com grupos, nomes, títulos, tendência e quadrantes ✓ (2); arredondamento por fórmula nativa + Arredondar intervalo, UDF recusada e justificada ✓ (3); PowerPoint imagem e vinculado com as mensagens ✓ (5); certificado/editor confiável/assinatura (scripts, teste, gancho do release, roteiro manual) ✓ (7, 11); corte de eixo como variante viável ✓ (4); ícones próprios com fallback e teste do pacote ✓ (6, 8); itens baratos da §10: nenhum (justificado na spec §11); faixa ✓ (8), matriz ✓ (0), exemplos ✓ (9), demonstração ✓ (10), docs/versão/release/instalação ✓ (11), verificação final ✓ (12). Pré-flight na spec §12.
- **Placeholders:** nenhum "TBD/TODO"; todo passo traz o código ou a fórmula; toda mensagem de usuário aparece com o texto exato onde é testada.
- **Consistência de nomes:** `Block_FindCol/Block_Extend/Block_SourceRange` (T0) usados em T3/T4; `Style_PaletteRGB` (T0) em T2; `Bfly_*` (T1) em T9/T10 via `Kinds_Insert`; `Scatter_SetTrend/SetQuad` (T2) em T9; `Round_SetLabels/Round_WriteFormulas/Round_RCol` (T3) em T9/T10; `Break_Set/Core_SetAxisBreakAt` (T4) em T9/T10; `Core_CopyToPowerPoint` (T5) em T8; `Get-TqqIconSpecs/Export-TqqIcons/Convert-TqqRibbonToMso` (T6) em `Build-Common` e `Test-Ribbon`; nomes `tqq_*`/`tqq_*16` (T6) no XML (T8); `Get-TqqDemoCoverage/Get-TqqDemoFeatures` (T10) no gerador e no teste; erros `6011` (borboleta), `6012/6013` (dispersão), `6014/6015` (extensão de bloco), `7003` (corte), `7002` (opções), `7001` (anotações).
- **Review Focus → testes:** 1 → `Test_Bfly_LongName_GapFitsAndAxisSymmetric` (T1); 2 → `Test_Scatter_GroupChange_MovesPointAndLabel` (T2); 3 → `Test_Round_Stacked100_SharesSumTo100` (T3); 4 → `Test_Break_SegmentCrossingLimit_LabelsAndMark` (T4); 5 → `Test_Ppt_LinkedPrecheck_UnsavedWorkbook`, `Test_Ppt_RealPaste_WhenAvailable` (T5).
