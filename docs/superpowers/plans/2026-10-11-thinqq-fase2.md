# ThinQQ — Plano de Implementação (Fase 2)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Entregar o ThinQQ 1.1.0 (core e loader): barras horizontais (empilhadas, agrupadas, 100% e cascata), cascata com várias séries e "de cima para baixo", área e combinação, Mekko (% e unidades), Gantt (com colchetes), pizza e rosca, linhas de valor, indicador 100%, mesma escala, rótulos no fim da linha, faixa de opções reorganizada em menus — e, antes de tudo, a correção residual da limpeza de blocos (Tarefa 0).

**Architecture:** Mantém a da Fase 1: `ThinQQ.xlam` (loader: faixa, boot, atualizador) chama `Core_*` em `ThinQQ.core.xlam` via `Application.Run`; cada gráfico é um gráfico nativo do Excel ligado por fórmulas a um bloco da planilha muito oculta `_ThinQQ`, com séries auxiliares `tqq_*` para rótulos/totais e formas `tqq_ann_*` só para o que não é dado (setas, indicador 100%, colchetes). Novidades estruturais: matriz de capacidades por tipo (`modChartKinds`), séries auxiliares orientáveis (linha em gráficos verticais, XY no grupo secundário em horizontais), geometria/anotações por orientação, módulos novos por família (área/combinação, pizza, Mekko, Gantt, decorações), faixa com `menu`s habilitados por capacidade.

**Tech Stack:** VBA 7 (Excel 2016+ Windows, 32/64 bits, sem `Declare`), customUI14, PowerShell 7 (build/testes/release) e 5.1 (instalador), Node.js (fixtures HTTP), Git.

**Spec:** `docs/superpowers/specs/2026-10-11-thinqq-fase2-design.md` (adendo; argumenta cada decisão) sobre `docs/superpowers/specs/2026-10-10-thinqq-design.md` (Fase 1, autoridade da arquitetura). Ledger e relatórios da Fase 1: `.superpowers/sdd/2026-10-10-thinqq-fase1/` (progress.md, task-N-report.md, final-fix-report.md).

**Estado de execução:** a Tarefa 0 está **concluída** (commits `b19615e` e `6b036c3` da `feat/fase2`; não reexecutar). As regras de ambiente para os implementadores ficam em `.superpowers/sdd/2026-10-11-thinqq-fase2/implementer-rules.md` (fora do repositório; o controlador as entrega junto com cada tarefa). O plano foi corrigido por uma varredura de coerência (pre-flight) antes da Tarefa 1.

## Global Constraints

- **Ambiente:** Excel 2019 32 bits pt-BR no Windows 11; testes em Excel invisível próprio. O VBA compila sob demanda: o build roda o lint (`build/lib/Vba-Lint.ps1`) e uma compilação completa (`Compile VBAProject`); erro de compilação => saída 3; travamento morto pelo watchdog => saída 2. Fontes UTF-8 com **somente caracteres cp1252** nos arquivos VBA (o lint falha com `→`, `≥`, aspas tipográficas fora do cp1252, emojis; `–` e `…` existem no cp1252 mas evite-os em código), sem ` _` perdido (continuação seguida de texto ou de linha vazia), `Option Explicit`, declarações antes dos procedimentos, `Attribute VB_Name` na primeira linha de cada `.bas`.
- **Testes:** `Public Sub Test_x()` sem parâmetros em `tests/vba/{core,loader}/Test*.bas`, executados por um despachante gerado (nunca `Application.Run`), via `pwsh build/run-tests.ps1` (suíte completa ~2–3 min; parâmetros `-Target core|loader|all`, `-NoBuild`, `-TimeoutSec`, `-Filter "Modulo.Teste|outro"` (pula as suítes PowerShell), `-LoaderSwap`). Erro de compilação em teste novo => saída 3 com `módulo:linha`. Asserts: `AssertTrue/False/Equal/Almost`, `NewTempWorkbook`, `CloseTemp`, `TempDir`, `RepoPath`.
- **Instâncias do usuário:** os processos Excel **do usuário** (anote os PIDs de `Get-Process excel` antes de começar) nunca são tocados nem encerrados; um deles mantém `dist\ThinQQ.core.test.xlam` travado no repositório real. Por isso **builds e testes rodam numa CÓPIA**: `robocopy "<repo>" C:\temp\tqq-f2 /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS` (repetir antes de cada rodada; edite sempre o repositório real; commite a partir do repositório real; `cd` no comando só com caminho absoluto). Nunca `build.ps1 -Install`; nunca tocar em `%LOCALAPPDATA%\ThinQQ` nem no registro real do usuário além do que os testes existentes do atualizador já salvam/restauram. A partir da Tarefa 1, `Errors_Log` no build `Test` grava em `%TEMP%\ThinQQ-test-logs\thinqq.log` (antes, os testes de falha acrescentavam linhas ao log real do usuário). Pouco espaço em disco: apague as suas pastas temporárias (`C:\temp\tqq-f2`, `%TEMP%\thinqq_test_*` são limpas pelo runner) e **nunca apague nada fora das suas próprias pastas temporárias**.
- **Código:** procedimentos `Core_*` nunca propagam erro (`Errors_Handle`/`Errors_HandleErr`); `MsgBox`/formulários só registram no build `Test` (`Core_Notify`, `RxMsg`, `Loader_Msg`; formulários não abrem quando `BUILD_CONFIG = "Test"`); fórmulas via `.Formula`/`.FormulaR1C1` em inglês; formatos de **rótulo** de gráfico via `Style_AssignNumberFormat`/`NumberFormatLocal` (códigos interpretados na localidade: nunca `DataLabel.NumberFormat` direto) e formatos de **célula** do bloco via `NumberFormatLocal` com `Style_LocalizeCode` para números e `Range.NumberFormat` em inglês para datas (`dd/mm`, `mmm/yy`); testes de texto renderizado via `DataLabel.Text` com expectativas derivadas de `Application.International(xlDecimalSeparator/xlThousandsSeparator)` ou `Format$`; textos de interface em pt-BR; identificadores em inglês com os prefixos existentes (`Core_`, `Stacked_`, `Waterfall_`, `Geom_`, `Ann_`, `Labels_`, `Style_`, `Chart_`, `Util_`, `Registry_`, `CalcSheet_`, `Samples_`, `Rx_`, `Loader_`) e os novos (`Kinds_`, `Area_`, `Combo_`, `Pie_`, `Mekko_`, `Gantt_`, `Deco_`).
- **Todo tipo novo de gráfico entrega, obrigatoriamente:** `Insert` com rollback (bloco e `ChartObject` apagados em falha; `Util_TestHookAfterBlock` após alocar o bloco), opções persistidas por `Registry_SetOption` (incl. `n`), rótulos ligados a células (ou nativos, na pizza), matriz de anotações com recusa explícita (erro `vbObjectError + 7001` com mensagem por tipo, que `Core_AddAnnotation` mostra sem registrar falha), testes de geometria **independentes das fórmulas de produção** (regra de três sobre `PlotArea.Inside*` e `Axis.MinimumScale/MaximumScale` reais), um exemplo em `Samples_Create`, seção no tutorial, entrada na faixa validada por `Test-Ribbon`/`TestRibbon` e cobertura em `Test-Docs`.
- **Commits:** pequenos, em português, prefixos `feat:`/`fix:`/`test:`/`docs:`/`build:`; trailer `Co-Authored-By:` **indicado no prompt de cada implementador** (os implementadores usam `Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>`); "(+ trailer)" nos comandos de commit deste plano significa essa linha. Branch `feat/fase2` (já existe); **não** fazer push, **não** criar nem trocar de branch, não tocar em outras branches.
- **Versões:** `version.json` passa a `{"loader":"1.1.0","core":"1.1.0"}` só na Tarefa 14; `APP_VERSION`/`BUILD_CONFIG`/`BUILD_SUFFIX` vêm de `modBuildInfo` gerado pelo build. O build importa automaticamente todo `.bas` de `src/<alvo>/modules`, `.cls` de `classes`, `.json`+`.vba` de `forms` e todo `Test*.bas` de `tests/vba/<alvo>` (nenhum registro manual de módulo é necessário).
- **Processo:** cada tarefa é executada por um implementador novo que só vê o texto da tarefa; tudo o que ele precisa (interfaces, fórmulas, fixtures, mensagens exatas, comandos) está na própria tarefa. Cada tarefa termina com a suíte completa verde na cópia e um commit por tarefa (ou mais, se útil).

## Review Focus

1. **Auxiliares XY em gráficos horizontais desalinham depois de mudar os dados ou o tamanho do gráfico** (o eixo X secundário não acompanha o primário): o usuário espera o rótulo/total no fim da barra. Teste em Tarefa 3 (`Test_Bars_TotalsHelper_StaysAlignedAfterDataChangeAndResize`).
2. **Cascata com várias séries em que uma coluna tem sinais misturados e cruza o zero** (ex.: início 130, -200 e -20): o usuário espera cada segmento com a extensão certa e o conector no total líquido. Teste em Tarefa 4 (`Test_WaterfallMulti_Block_MixedSignsCrossingZero`).
3. **Mekko com uma linha de total zero (ou largura zero/texto)**: o usuário espera o gráfico inserido sem `#DIV/0!` e sem coluna fantasma. Teste em Tarefa 9 (`Test_Mekko_ZeroTotalRow_NoErrorNoWidth`).
4. **Gantt com texto onde deveria haver data, só marcos, ou uma única atividade; e mudança de dados que troca a unidade da escala (semanas → meses)**: o usuário espera recusa clara no primeiro caso e, no segundo, a escala atualizada ao clicar em Atualizar. Testes em Tarefa 10 (`Test_Gantt_InvalidDates_RefusedWithMessage`, `Test_Gantt_ScaleUnitSwitches_OnRefresh`, `Test_Gantt_OnlyMilestones_And_SingleActivity`).
5. **Pizza com valores negativos, zero, texto e mais de 6 fatias, com rótulos fora e nomes longos**: o usuário espera fatias só dos positivos, cores ciclando a paleta e rótulos dentro da área do gráfico. Teste em Tarefa 8 (`Test_Pie_NegativeZeroText_AndSevenSlices`).

---

## Estrutura de arquivos (novos e alterados)

```
src/core/modules/
├── modChartKinds.bas        NOVO  Kinds_IsValid/IsHorizontal/Is100/Family/Supports/ExpectedCategoryRefs/Insert (despacho)
├── modChartUtil.bas         ALT   Chart_AddHelperSeries, Chart_SetupSecondaryAxes, Chart_SyncSecondaryX, Chart_HideHelperLegendEntries (ganchos Util_*ForTest da limpeza: Tarefa 0, feito)
├── modCalcSheet.bas         ALT   Block_Ref (limpeza por rótulos ligados e falha fechada: Tarefa 0, feito)
├── modErrors.bas            ALT   Errors_LogDir: log do build Test em %TEMP%\ThinQQ-test-logs (Tarefa 1)
├── modLabels.bas            ALT   Labels_LinkPoints(horizontal), Labels_RefreshPositions por orientação e isolado por ponto
├── modChartStyle.bas        ALT   Style_ApplyBase tolera pizza e inverte categorias em horizontais; Style_ApplyPointColors; Style_ApplyNumberFormat por família
├── modGeometry.bas          ALT   Geom_CategoryCenterY, Geom_ValueToX, Geom_ChartIsHorizontal, Geom_ChartCatCoord, Geom_ChartValCoord, Geom_ChartValueToX, Geom_ChartCategoryCenterY
├── modAnnotations.bas       ALT   layouts por orientação; kinds ind100 (Ann_AddInd100/Ann_RemoveInd100) e bracket; recusas por tipo; Ann_GetCategoryData para area/combo/barras
├── modChartStacked.bas      ALT   kinds barstacked/barclustered/barstacked100; colunas Pos/VLine; totais por auxiliar orientável
├── modChartWaterfall.bas    ALT   várias séries (cols 21+), build-down (col 20), barwaterfall
├── modChartArea.bas         NOVO  area, area100
├── modChartCombo.bas        NOVO  combo (+ eixo secundário)
├── modChartPie.bas          NOVO  pie, doughnut
├── modChartMekko.bas        NOVO  mekko, mekkounits
├── modChartGantt.bas        NOVO  gantt (+ escala, marcos)
├── modDecorations.bas       NOVO  linha de valor, indicador 100%, legenda, mesma escala
├── modCoreApi.bas           ALT   Core_Supports, Core_SameScale, Core_ResetScale, despacho por família, colchete
└── modSamples.bas           ALT   planilha "Exemplos 2"
src/loader/ribbon/customUI14.xml  ALT  menus; src/loader/modules/modRibbon.bas ALT (Rx_GetEnabled por tag, Rx_OnSameScale, Rx_OnResetScale)
tests/vba/core/  TestKinds.bas, TestBars.bas, TestWaterfall2.bas, TestArea.bas, TestPie.bas, TestMekko.bas, TestGantt.bas, TestDecorations.bas (NOVOS);
                 TestCalcSheet, TestCoreApi, TestGeometry, TestAnnotations, TestWaterfall, TestStacked, TestSamples (ALT)
tests/vba/loader/TestRibbon.bas  ALT;   tests/ps/Test-Docs.ps1  ALT (menus)
docs/tutorial/tutorial.html, README.md, CHANGELOG.md, version.json, release/notes.txt  ALT (Tarefa 14)
```

## Como trabalhar (todas as tarefas)

```powershell
# 1) editar no repositório real; 2) espelhar para a cópia; 3) rodar lá
robocopy "<repo>" C:\temp\tqq-f2 /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS
pwsh C:\temp\tqq-f2\build\run-tests.ps1 -Target core -Filter "TestBars" -TimeoutSec 240     # iteração (RED/GREEN)
pwsh C:\temp\tqq-f2\build\run-tests.ps1 -TimeoutSec 600                                     # suíte completa antes de commitar
# saída 3 = erro de compilação (módulo:linha na tela); 2 = travou (watchdog); 1 = asserts; 0 = ok
```
Fixtures: cada módulo de teste tem a sua `Private Function Fixture(wb)`; use `Teardown`/`Watchers_Stop` quando chamar `Core_*` (estado global dos observadores). Nunca deixe `Util_*ForTest` ligado em caminho de erro (padrão `On Error GoTo Cleanup` que reseta e repropaga, como em `TestSamples`).

---

### Tarefa 0: Limpeza de blocos conta rótulos ligados e falha fechada (carry-over da Fase 1) — CONCLUÍDA

**Estado:** concluída e revisada nos commits `b19615e` e `6b036c3` da `feat/fase2` (relatório: `.superpowers/sdd/2026-10-11-thinqq-fase2/task-0-report.md`). **Não reexecutar:** os passos de teste, implementação e commit foram retirados deste plano; o texto abaixo só registra o que já existe e o comportamento do Excel observado, para as tarefas seguintes.

**O que existe no código:**
- `src/core/modules/modCalcSheet.bas`: `CalcSheet_Cleanup(wb) As Long` considera referência ao bloco qualquer `Series.Formula` **e** qualquer `DataLabel.Formula` (rótulo ligado) de qualquer gráfico da pasta (embutidos, dentro de grupos de formas via `GroupItems`, planilhas de gráfico em `wb.Charts`), pelas privadas `AllChartReferences`/`ChartReferences`/`LabelFormula`. **Falha fechada:** qualquer erro de leitura (forma, item de grupo, planilha de gráfico, gráfico, `Series.Formula`, `Point.HasDataLabel`, `DataLabel.Formula`, ou uma citação `_ThinQQ!...` que não pode ser interpretada) interrompe a passada e a limpeza não remove nada: `Errors_Log "CalcSheet_Cleanup: leitura de gráfico falhou; nenhum bloco removido [detalhe]"` e devolve 0. `Public Function CalcSheet_CleanupReadFailed() As Boolean` devolve o resultado da última limpeza (zerado ao entrar).
- Varredura de rótulos limitada: em cada série só os primeiros `MaxBlockRows(wb)` pontos são lidos (altura do maior bloco `tqq_*` da `_ThinQQ`); a leitura para no primeiro rótulo que cita `_ThinQQ`; série cuja `Series.Formula` já cita `_ThinQQ` não tem rótulos lidos; orçamento de 20000 pontos por passada (estourou = falha fechada).
- Exceção estreita: gráfico nativo chartex (cascata/funil/treemap do Excel 2016+, `ChartType` 117..123 ou 140) cujo `Series.Formula` lança o erro 445 é ignorado (não é falha; o ThinQQ não cria chartex).
- Ganchos de teste em `src/core/modules/modChartUtil.bas` (efeito só com `BUILD_CONFIG = "Test"`, como `Util_FailAfterBlockForTest`): `Util_FailChartReadForTest As Boolean`; `Util_FailChartReadSiteForTest As Long` com as constantes `READFAIL_SHAPE = 1`, `READFAIL_GROUP = 2`, `READFAIL_CHARTSHEET = 3`, `READFAIL_CHART = 4`, `READFAIL_SERIES = 5`, `READFAIL_HASLABEL = 6`, `READFAIL_FORMULA = 7`; `Util_LabelReadsForTest` (contador de pontos lidos); `Util_LabelBudgetForTest` (substitui o limite de 20000).
- Testes já existentes (nenhuma tarefa posterior pode reutilizar estes nomes): em `TestCoreApi.bas` `Test_Core_Cleanup_Stacked100NoTotals_CopyKeepsBlock`, `Test_Core_Cleanup_ChartInsideGroup_KeepsBlock`, `Test_Core_Insert_WhileCleanupAborts_StillInsertsAndDoesNotOverlapOrphan`; em `TestCalcSheet.bas` `Test_CalcSheet_Cleanup_ReadFailure_DeletesNothing`, `Test_CalcSheet_Cleanup_ReadFailureAtEverySite_DeletesNothing`, `Test_CalcSheet_Cleanup_StopsScanningAfterFirstFailure`, `Test_CalcSheet_Cleanup_PointBudgetExceeded_FailsClosed`, `Test_CalcSheet_Cleanup_LargeChart_IsBoundedAndNotAFailure`, `Test_CalcSheet_Cleanup_SeriesReferencingCalcSheet_ReadsNoLabels`, `Test_CalcSheet_Cleanup_LabelsOnlyReference_KeepsBlock_StopsAtFirstLinkedLabel`, `Test_CalcSheet_Cleanup_PlainLabels_DoNotBlockCleanup`, `Test_CalcSheet_Cleanup_UnparseableCalcSheetMention_FailsClosed`, `Test_CalcSheet_Cleanup_NativeChartEx_IsNotAReadFailure`, `Test_CalcSheet_LabelFormula_ObservedBehavior`. Suíte ao fim da tarefa: core 173, loader 95, PowerShell verde.

**Comportamento observado no Excel 2019 32-bit pt-BR (fixado em `Test_CalcSheet_LabelFormula_ObservedBehavior`):**
- Ponto **sem** rótulo: `Point.HasDataLabel = False`, sem erro; ler `Point.DataLabel` lança erro (-2147467259; em rótulo apagado/parcial, -2147024809). Sempre teste `HasDataLabel` antes de tocar em `DataLabel`.
- Rótulo automático: `DataLabel.Formula` não lança e devolve o **texto exibido** (ex.: `10`, igual a `.Text`); rótulo editado à mão (`.Text = "texto livre"`) devolve o texto; rótulo ligado a célula devolve `=Dados!$B$2` (sempre em A1, mesmo com `Application.ReferenceStyle = xlR1C1`).
- `Series.HasDataLabels = False` quando só alguns pontos têm rótulo (não serve para pular séries).
- Chartex: `HasChart = True`, `ChartType = 119`, `Points.Count` normal, mas `Series.Formula` e `DataLabel.Formula` lançam 445.
- Custo medido: 3000 leituras de `HasDataLabel` + `.Formula` com rótulos levam cerca de 0,17 s.

**Regras que valem para as tarefas seguintes:**
- Todos os rótulos ligados de uma série apontam para o bloco do próprio gráfico (a varredura assume isso): nunca ligue rótulos de uma série a um bloco que nenhuma série do gráfico cite.
- Séries auxiliares com pontos `NA()` (Tarefas 4, 9 e 10) entram na varredura: os testes dessas tarefas confirmam que uma limpeza com o gráfico novo e um bloco órfão remove só o órfão e não marca falha.
- A pendência "o `Errors_Log` do build Test escreve no log real do usuário" é resolvida na Tarefa 1.

**Verificação manual pendente (usuário, Excel real; entra na lista do README na Tarefa 14):** inserir Colunas 100%, desligar Totais, copiar/colar o gráfico, apagar o original, clicar Atualizar: a cópia mantém as participações; inserir outro gráfico: a cópia continua certa (o mesmo roteiro já é coberto, em Excel invisível, por `Test_Core_Cleanup_Stacked100NoTotals_CopyKeepsBlock`).

---

### Tarefa 1: Matriz de tipos, capacidades e séries auxiliares orientáveis (infraestrutura)

**Files:**
- Create: `src/core/modules/modChartKinds.bas`
- Modify: `src/core/modules/modChartUtil.bas`, `src/core/modules/modCalcSheet.bas` (`Block_Ref`), `src/core/modules/modLabels.bas`, `src/core/modules/modChartStyle.bas`, `src/core/modules/modChartStacked.bas`, `src/core/modules/modChartWaterfall.bas` (opção `n`), `src/core/modules/modCoreApi.bas`, `src/core/modules/modErrors.bas` (log do build Test fora da pasta do usuário)
- Test: `tests/vba/core/TestKinds.bas` (novo), `tests/vba/core/TestCoreApi.bas`, `tests/vba/core/TestStacked.bas`

**Interfaces:**
- Produces (`modChartKinds`):
  - `Kinds_IsValid(kind As String) As Boolean` — os 17 tipos: `stacked, clustered, stacked100, line, waterfall, barstacked, barclustered, barstacked100, barwaterfall, area, area100, combo, mekko, mekkounits, gantt, pie, doughnut`.
  - `Kinds_IsHorizontal(kind) As Boolean` — `barstacked, barclustered, barstacked100, barwaterfall, gantt`.
  - `Kinds_Is100(kind) As Boolean` — `stacked100, barstacked100, area100, mekko`.
  - `Kinds_Family(kind) As String` — `stacked` (stacked, clustered, stacked100, line, barstacked, barclustered, barstacked100), `waterfall` (waterfall, barwaterfall), `area` (area, area100), `combo`, `mekko` (mekko, mekkounits), `gantt`, `pie` (pie, doughnut); `""` se inválido.
  - `Kinds_Supports(kind, feature As String) As Boolean` — matriz do design §7 (features: `totals, labels, connectors, numfmt, palette, gap, legend, ind100, vline, secondary, pieout, cagr, diff, bracket, samescale`); `False` para tipo ou feature desconhecidos. `pieout` vale **só para `pie`**: a rosca não aceita `DataLabels.Position` `InsideEnd`/`OutsideEnd` (o Excel lança 1004; ver Tarefa 8).
  - `Kinds_ExpectedCategoryRefs(kind, n As Long) As Long` — `5*n` para a família `mekko`, senão `n`.
  - `Kinds_Insert(ws As Worksheet, data As clsChartData, kind As String) As ChartObject` — despacho por família; nesta tarefa só `stacked` → `Stacked_Insert(ws, data, kind)` e `waterfall` → `Waterfall_Insert(ws, data)`; demais famílias: `Err.Raise vbObjectError + 7200, "Kinds_Insert", "Este tipo de gráfico ainda não está disponível nesta versão: " & kind`. Tarefas seguintes acrescentam os `Case`.
- Produces (`modChartUtil`):
  - `Chart_AddHelperSeries(cht As Chart, name As String, valueCells As Range, Optional posCells As Range = Nothing) As Series` — `posCells Is Nothing`: série `xlLine` invisível (linha e marcador ocultos) com `Values = Block_Ref(valueCells)` (grupo primário); com `posCells`: série `xlXYScatter` no grupo **secundário**, `XValues = Block_Ref(valueCells)`, `Values = Block_Ref(posCells)`, sem marcador nem linha, e chama `Chart_SetupSecondaryAxes cht, posCells.Rows.Count`. Não mexe na legenda (quem chama decide com `Chart_HideHelperLegendEntries`).
  - `Chart_SetupSecondaryAxes(cht As Chart, n As Long)` — `cht.HasAxis(xlValue, xlSecondary) = True`; eixo Y secundário: `MinimumScale = 0`, `MaximumScale = n`, `ReversePlotOrder = False`, sem marcas, sem rótulos, linha invisível, sem grade; se `cht.HasAxis(xlCategory, xlSecondary)` for True, chama `Chart_SyncSecondaryX`.
  - `Chart_SyncSecondaryX(cht As Chart)` — se existe eixo X secundário (`HasAxis(xlCategory, xlSecondary)`), copia `MinimumScale/MaximumScale` do eixo de valores primário para ele e o oculta; senão não faz nada. Nunca propaga erro.
  - `Chart_HideHelperLegendEntries(cht As Chart)` — remove da legenda toda entrada de série `tqq_*`; idempotente e correta com várias auxiliares (o teste de hoje em `Stacked_HideTotalLegendEntry`, `LegendEntries.Count < SeriesCollection.Count`, para depois da primeira remoção). Algoritmo: sem legenda ou sem séries `tqq_*`, nada; `h` = nº de séries `tqq_*`; se `LegendEntries.Count = SeriesCollection.Count - h`, já está feito (sai); senão, se `LegendEntries.Count <> SeriesCollection.Count` (estado parcial), lê `Legend.Position` e `Legend.Font.Size`, religa a legenda (`HasLegend = False: HasLegend = True`) e restaura os dois; então calcula o índice de legenda de cada série `tqq_*` pela ordem dos grupos (`cht.ChartGroups`, inclusive os do eixo secundário) e das séries dentro de cada grupo, e apaga essas entradas **do maior índice para o menor** (apagar renumera as seguintes). `Stacked_HideTotalLegendEntry` passa a delegar a esta função (`Test_Stacked_Legend_ExtraSeries_KeepsDataEntries`, da Fase 1, continua verde). Cuidado: em pilhas com legenda vertical o Excel pode exibir as entradas em ordem invertida, e grupos secundários mudam a ordem; o teste abaixo confere o mapeamento pela cor da chave da legenda (se `LegendKey.Format.Fill.ForeColor.RGB` não refletir a cor da série no Excel 2019, use `LegendKey.Interior.Color` e registre no relatório); as Tarefas 3 e 4 conferem por `Chart.Export`.
- Produces (`modCalcSheet`): `Block_Ref(cells As Range) As String` = `"='_ThinQQ'!" & cells.Address(True, True)`.
- Produces (`modLabels`): `Labels_LinkPoints(s, valueCells, numFmt, Optional fixedPos As Long = 0, Optional fontColor As Long = -1, Optional horizontal As Boolean = False)` — com `horizontal`, a posição por sinal é `xlLabelPositionRight` (≥ 0) / `xlLabelPositionLeft` (< 0); `Labels_RefreshPositions` usa `Geom_ChartIsHorizontal`-equivalente (`Kinds_IsHorizontal(Registry_ChartType(cht))`) para escolher Right/Left ou Above/Below e isola cada ponto com `On Error Resume Next` (um ponto não plotado — `NA()` — não interrompe os demais). `Labels_RefreshPositions` e `Labels_ApplyCellFormat` decidem por **família**, não pelo nome do tipo: `Kinds_Family(kind) = "waterfall"` → série `tqq_lbl` e coluna `WF_COL_LABELVAL` (hoje ambas comparam `Registry_ChartType(cht) = "waterfall"`, o que deixaria a cascata horizontal da Tarefa 5 de fora); colunas de participação (4..) quando `Kinds_Is100(kind)` na família `stacked`.
- Produces (`modChartStyle`): `Style_ApplyBase` pula o bloco de eixos quando `cht.ChartType` é pizza/rosca (`xlPie`, `xlDoughnut`, `xlPieExploded`, `xlDoughnutExploded`) e, quando o gráfico é de barras (`xlBarStacked`, `xlBarClustered`, `xlBarStacked100`), define `Axes(xlCategory).ReversePlotOrder = True` (primeira categoria no topo) mantendo `TickLabelPosition = xlTickLabelPositionLow`.
- Produces (`modCoreApi`): `Core_Supports(feature As String) As Boolean` — `False` sem gráfico ThinQQ ativo; `"changesource"` → True se há gráfico; demais → `Kinds_Supports(Registry_ChartType(c), feature)`; nunca propaga erro. `Core_Insert` aceita todo `Kinds_IsValid`, despacha por `Kinds_Insert` e trata `vbObjectError + 7200` com `Core_Notify Err.Description, vbInformation` (sem `Errors_Handle`). `Core_ChartSourceMismatch` compara com `Kinds_ExpectedCategoryRefs(kind, CLng(Registry_GetOption(cht, "n", "0")))` quando a opção `n` existe; senão com `blk.Rows.Count - 2` como hoje. `Core_RebuildFromRange` também passa a construir por `Kinds_Insert(ws, data, kind)` (hoje: `If kind = "waterfall" Then Waterfall_Insert ... Else Stacked_Insert ...`), para que cada tipo das tarefas seguintes seja reconstruído pelo próprio `Insert`.
- Produces: `Stacked_Insert` e `Waterfall_Insert` passam a gravar `Registry_SetOption cht, "n", CStr(data.CategoryCount)`. `Labels_SegmentFormat/TotalFormat` e `Style_ApplyNumberFormat` usam `Kinds_Is100` em vez de comparar com `"stacked100"`.
- Produces (`modErrors`): `Public Function Errors_LogDir() As String` = `Environ$("TEMP") & "\ThinQQ-test-logs"` quando `BUILD_CONFIG = "Test"`, senão `Environ$("LOCALAPPDATA") & "\ThinQQ\logs"` (como hoje); `Errors_Log` grava em `Errors_LogDir() & "\thinqq.log"` (o corpo continua inteiro sob `On Error Resume Next`). Motivo: os testes de falha (rollback, limpeza fechada) escreviam no log real do usuário. A pasta de teste é pequena e reaproveitada (não precisa ser apagada).

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestKinds.bas`:

```vb
Attribute VB_Name = "TestKinds"
Option Explicit

Public Sub Test_Kinds_Valid_And_Family()
    AssertTrue Kinds_IsValid("stacked"): AssertTrue Kinds_IsValid("barwaterfall"): AssertTrue Kinds_IsValid("doughnut")
    AssertFalse Kinds_IsValid("pizza"): AssertFalse Kinds_IsValid("")
    AssertEqual "stacked", Kinds_Family("barclustered")
    AssertEqual "waterfall", Kinds_Family("barwaterfall")
    AssertEqual "area", Kinds_Family("area100")
    AssertEqual "combo", Kinds_Family("combo")
    AssertEqual "mekko", Kinds_Family("mekkounits")
    AssertEqual "gantt", Kinds_Family("gantt")
    AssertEqual "pie", Kinds_Family("doughnut")
    AssertEqual "", Kinds_Family("x")
End Sub

Public Sub Test_Kinds_Horizontal_100_And_Refs()
    AssertTrue Kinds_IsHorizontal("barstacked"): AssertTrue Kinds_IsHorizontal("gantt")
    AssertFalse Kinds_IsHorizontal("stacked"): AssertFalse Kinds_IsHorizontal("mekko")
    AssertTrue Kinds_Is100("stacked100"): AssertTrue Kinds_Is100("mekko"): AssertTrue Kinds_Is100("area100")
    AssertFalse Kinds_Is100("mekkounits"): AssertFalse Kinds_Is100("barstacked")
    AssertEqual 20, Kinds_ExpectedCategoryRefs("mekko", 4)
    AssertEqual 4, Kinds_ExpectedCategoryRefs("gantt", 4)
End Sub

Public Sub Test_Kinds_Supports_Matrix()
    AssertTrue Kinds_Supports("stacked", "totals"): AssertFalse Kinds_Supports("clustered", "totals"): AssertTrue Kinds_Supports("mekko", "totals")
    AssertTrue Kinds_Supports("gantt", "labels"): AssertFalse Kinds_Supports("gantt", "numfmt")
    AssertTrue Kinds_Supports("waterfall", "connectors"): AssertFalse Kinds_Supports("line", "connectors")
    AssertTrue Kinds_Supports("gantt", "gap"): AssertFalse Kinds_Supports("pie", "gap"): AssertFalse Kinds_Supports("mekko", "gap")
    AssertTrue Kinds_Supports("stacked100", "ind100"): AssertTrue Kinds_Supports("mekko", "ind100"): AssertFalse Kinds_Supports("mekkounits", "ind100")
    AssertTrue Kinds_Supports("mekkounits", "vline"): AssertFalse Kinds_Supports("stacked100", "vline"): AssertTrue Kinds_Supports("barclustered", "vline")
    AssertTrue Kinds_Supports("combo", "secondary"): AssertFalse Kinds_Supports("stacked", "secondary")
    AssertTrue Kinds_Supports("pie", "pieout"): AssertFalse Kinds_Supports("doughnut", "pieout"): AssertFalse Kinds_Supports("area", "pieout")
    AssertTrue Kinds_Supports("barwaterfall", "cagr"): AssertFalse Kinds_Supports("pie", "cagr"): AssertFalse Kinds_Supports("mekko", "cagr")
    AssertTrue Kinds_Supports("stacked", "diff"): AssertFalse Kinds_Supports("stacked100", "diff"): AssertFalse Kinds_Supports("area100", "diff")
    AssertTrue Kinds_Supports("gantt", "bracket"): AssertFalse Kinds_Supports("stacked", "bracket")
    AssertTrue Kinds_Supports("line", "samescale"): AssertFalse Kinds_Supports("pie", "samescale"): AssertFalse Kinds_Supports("mekko", "samescale")
    AssertTrue Kinds_Supports("stacked", "legend"): AssertFalse Kinds_Supports("gantt", "legend")
    AssertTrue Kinds_Supports("pie", "palette"): AssertTrue Kinds_Supports("waterfall", "labels")
    AssertFalse Kinds_Supports("stacked", "feature_inexistente")
    AssertFalse Kinds_Supports("tipo_inexistente", "labels")
End Sub

Public Sub Test_Chart_AddHelperSeries_Vertical_IsInvisibleLine()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, body As Range, s As Series
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Ano", "Valor")
    ws.Range("A2:B2").Value = Array(2022, 10): ws.Range("A3:B3").Value = Array(2023, 20)
    Set d = ChartData_FromRange(ws.Range("A1:B3"))
    Set cht = Stacked_Insert(ws, d, "stacked").Chart
    Set body = Block_Body(Registry_Block(cht))
    Set s = Chart_AddHelperSeries(cht, "tqq_x", body.Cells(2, 2).Resize(2, 1))
    AssertEqual "tqq_x", s.Name
    AssertEqual xlLine, s.ChartType
    AssertEqual msoFalse, s.Format.Line.Visible
    AssertEqual xlMarkerStyleNone, s.MarkerStyle
    AssertEqual xlPrimary, s.AxisGroup
    AssertTrue InStr(s.Formula, CALC_SHEET) > 0
    AssertEqual "='" & CALC_SHEET & "'!" & body.Cells(2, 2).Resize(2, 1).Address(True, True), Block_Ref(body.Cells(2, 2).Resize(2, 1)), "Block_Ref = referencia absoluta na _ThinQQ"
    CloseTemp wb
End Sub

Public Sub Test_Chart_HideHelperLegendEntries_RemovesAllTqq()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, body As Range, pal As Variant, keys As String, k As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 10, 5): ws.Range("A3:C3").Value = Array(2023, 20, 6)
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Set cht = Stacked_Insert(ws, d, "stacked").Chart          ' ja tem tqq_total fora da legenda
    Set body = Block_Body(Registry_Block(cht))
    Chart_AddHelperSeries cht, "tqq_a", body.Cells(2, 2).Resize(2, 1)
    Chart_AddHelperSeries cht, "tqq_b", body.Cells(2, 3).Resize(2, 1)
    AssertTrue cht.Legend.LegendEntries.Count >= 3, "as auxiliares novas aparecem na legenda (estado parcial)"
    Chart_HideHelperLegendEntries cht
    AssertEqual 2, cht.Legend.LegendEntries.Count, "estado parcial resolvido: so Sul e Norte"
    AssertEqual xlLegendPositionRight, cht.Legend.Position, "religar a legenda preserva a posicao"
    AssertEqual 9, cht.Legend.Font.Size, "religar a legenda preserva a fonte"
    cht.HasLegend = False: cht.HasLegend = True               ' religa: todas as entradas voltam
    AssertEqual 5, cht.Legend.LegendEntries.Count
    Chart_HideHelperLegendEntries cht
    AssertEqual 2, cht.Legend.LegendEntries.Count, "so Sul e Norte"
    Chart_HideHelperLegendEntries cht
    AssertEqual 2, cht.Legend.LegendEntries.Count, "idempotente"
    ' mapeamento: as entradas que ficaram sao as das series de dados (cor da chave = cor da serie)
    pal = Style_Palette("padrao")
    For k = 1 To 2: keys = keys & ";" & cht.Legend.LegendEntries(k).LegendKey.Format.Fill.ForeColor.RGB: Next
    keys = keys & ";"
    AssertTrue InStr(keys, ";" & CLng(pal(0)) & ";") > 0 And InStr(keys, ";" & CLng(pal(1)) & ";") > 0, "entradas restantes = Sul e Norte (cores " & keys & ")"
    CloseTemp wb
End Sub
```

Em `tests/vba/core/TestCoreApi.bas`:

```vb
Public Sub Test_Core_Supports_FollowsActiveChart()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    AssertTrue Core_Supports("totals"): AssertTrue Core_Supports("changesource")
    AssertFalse Core_Supports("secondary"): AssertFalse Core_Supports("pieout")
    AssertEqual "4", Core_GetOption("n"), "numero de categorias persistido"
    ws.Range("A1").Select
    AssertFalse Core_Supports("totals"): AssertFalse Core_Supports("changesource")
    Teardown wb
End Sub

' Temporario: a Tarefa 10 o remove ao implementar o Gantt (o ultimo tipo a ficar pronto). Tipo valido ainda sem
' implementacao avisa e nao deixa bloco.
Public Sub Test_Core_Insert_KindNotReady_NotifiesWithoutChart()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_ClearLastNotice
    Core_Insert "gantt"
    AssertEqual 0, ws.ChartObjects.Count
    AssertEqual 0, TqqNameCount(wb)
    AssertTrue InStr(Core_LastNotice(), "ainda não está disponível") > 0, Core_LastNotice()
    Teardown wb
End Sub

' Build Test: o log vai para %TEMP%\ThinQQ-test-logs, nunca para o log real do usuario.
Public Sub Test_Errors_Log_TestBuildWritesToTempFolder()
    Dim marker As String, p As String, f As Integer, txt As String
    AssertEqual Environ$("TEMP") & "\ThinQQ-test-logs", Errors_LogDir()
    Randomize
    marker = "tqq-log-probe-" & Format$(Now, "hhnnss") & "-" & CStr(Int(Rnd * 100000))
    Errors_Log marker
    p = Errors_LogDir() & "\thinqq.log"
    AssertTrue Len(Dir(p)) > 0, "log de teste criado em " & p
    f = FreeFile
    Open p For Binary Access Read As #f
    txt = Space$(LOF(f))
    Get #f, , txt
    Close #f
    AssertTrue InStr(txt, marker) > 0, "a linha foi para o log de teste"
End Sub
```

Em `tests/vba/core/TestStacked.bas` (comportamento novo de `Style_ApplyBase` em barras, testado já aqui com um gráfico cru):

```vb
Public Sub Test_Style_ApplyBase_Bars_ReverseCategories_AndPieSafe()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, s As Series
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B3").Value = Application.Transpose(Array(Array("a", "b", "c"), Array(1, 2, 3)))
    Set co = ws.ChartObjects.Add(10, 10, 300, 200)
    Set s = co.Chart.SeriesCollection.NewSeries: s.Values = "=" & ws.Name & "!$B$1:$B$3"
    co.Chart.ChartType = xlBarStacked
    Style_ApplyBase co.Chart
    AssertTrue co.Chart.Axes(xlCategory).ReversePlotOrder, "barras: primeira categoria no topo"
    AssertEqual xlTickLabelPositionLow, co.Chart.Axes(xlCategory).TickLabelPosition
    co.Chart.ChartType = xlPie
    Style_ApplyBase co.Chart                                   ' nao pode levantar erro (sem eixos)
    AssertFalse co.Chart.HasTitle
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-f2\build\run-tests.ps1 -Target core -Filter "TestKinds|Core_Supports|KindNotReady|ApplyBase_Bars|Errors_Log" -TimeoutSec 240` → saída 3 (símbolos indefinidos).

- [ ] **Passo 3: Implementar** — esboço de `modChartKinds.bas` (a matriz é dado, não lógica: use listas separadas por vírgula e `InStr`, como `WF_MARKERS`):

```vb
Attribute VB_Name = "modChartKinds"
Option Explicit
' Fonte unica dos tipos de grafico e do que cada um suporta (faixa, Core_SetToggle/SetOption, Rebuild).
Private Const KINDS As String = ",stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,gantt,pie,doughnut,"

Public Function Kinds_IsValid(kind As String) As Boolean
    If Len(kind) = 0 Then Exit Function
    Kinds_IsValid = InStr(KINDS, "," & kind & ",") > 0
End Function

Private Function InList(list As String, kind As String) As Boolean
    InList = InStr("," & list & ",", "," & kind & ",") > 0
End Function

Public Function Kinds_Supports(kind As String, feature As String) As Boolean
    Dim list As String
    If Not Kinds_IsValid(kind) Then Exit Function
    Select Case feature
        Case "totals":     list = "stacked,stacked100,barstacked,barstacked100,area,area100,combo,mekko,mekkounits"
        Case "labels":     list = Mid$(KINDS, 2, Len(KINDS) - 2)
        Case "connectors": list = "stacked,stacked100,barstacked,barstacked100,waterfall,barwaterfall,combo"
        Case "numfmt":     list = Replace(Mid$(KINDS, 2, Len(KINDS) - 2), ",gantt", "")
        ' ... palette, gap, legend, ind100, vline, secondary, pieout, cagr, diff, bracket, samescale conforme o design §7
        Case Else: Exit Function
    End Select
    Kinds_Supports = InList(list, kind)
End Function
```

`Kinds_Insert`: `Select Case Kinds_Family(kind)` com `Case "stacked": Set Kinds_Insert = Stacked_Insert(ws, data, kind)`, `Case "waterfall": Set Kinds_Insert = Waterfall_Insert(ws, data)`, `Case Else: Err.Raise vbObjectError + 7200, ...`. Em `Core_Insert`, trocar o `Select Case kind` por `If Not Kinds_IsValid(kind) Then Core_Notify ...: Exit Sub`, chamar `Kinds_Insert` e, no `Fail:`, `If Err.Number = vbObjectError + 7200 Then Core_Notify Err.Description, vbInformation Else Errors_Handle ...` (restaurando `ScreenUpdating` antes, como hoje). `Chart_AddHelperSeries` segue o padrão de `Stacked_SetTotals` (vertical) e, para `posCells`, `s.ChartType = xlXYScatter: s.AxisGroup = xlSecondary: s.XValues = Block_Ref(valueCells): s.Values = Block_Ref(posCells): s.MarkerStyle = xlMarkerStyleNone: s.Format.Line.Visible = msoFalse` — definir `ChartType` **antes** de `XValues` (uma série XY precisa existir como XY para aceitar `XValues` numéricos).

- [ ] **Passo 4: GREEN e suíte completa** — filtro acima verde; depois `pwsh C:\temp\tqq-f2\build\run-tests.ps1 -TimeoutSec 600` verde (os testes da Fase 1 de 100%/totais continuam passando com `Kinds_Is100`).

- [ ] **Passo 5: Commit**

```bash
git add src/core/modules/modChartKinds.bas src/core/modules/modChartUtil.bas src/core/modules/modCalcSheet.bas src/core/modules/modLabels.bas src/core/modules/modChartStyle.bas src/core/modules/modChartStacked.bas src/core/modules/modChartWaterfall.bas src/core/modules/modCoreApi.bas src/core/modules/modErrors.bas tests/vba/core/TestKinds.bas tests/vba/core/TestCoreApi.bas tests/vba/core/TestStacked.bas
git commit -m "feat: matriz de tipos e capacidades (Kinds_*), Core_Supports, series auxiliares orientaveis e despacho de insercao por familia" -m "<linha Co-Authored-By indicada no seu prompt>"
```

**Verificação no Excel real:** nenhuma mudança visível ainda (os tipos antigos continuam iguais); conferir só que Colunas empilhadas/Cascata inserem e que `Atualizar` não avisa "tabela mudou de tamanho" (opção `n` coerente).

---

### Tarefa 2: Geometria e anotações por orientação (colunas e barras)

**Files:**
- Modify: `src/core/modules/modGeometry.bas`, `src/core/modules/modAnnotations.bas`
- Test: `tests/vba/core/TestGeometry.bas`, `tests/vba/core/TestAnnotations.bas`

**Interfaces:**
- Consumes: `Kinds_IsHorizontal`, `Kinds_Is100`, `Kinds_Supports`, `Kinds_Family` (Tarefa 1); `Registry_ChartType`, `Style_IsHelperSeries`, `Waterfall_EndValues`.
- Produces (`modGeometry`, puras): `Geom_CategoryCenterY(insideTop, insideHeight, nCats, idx) As Double` = `insideTop + (idx - 0.5) * insideHeight / nCats` (categoria 1 **no topo**, que é como `Style_ApplyBase` inverte as barras); `Geom_ValueToX(insideLeft, insideWidth, axisMin, axisMax, v) As Double` = `insideLeft + insideWidth * (v - axisMin) / (axisMax - axisMin)` (escala degenerada → `insideLeft`).
- Produces (`modGeometry`, de gráfico): `Geom_ChartIsHorizontal(cht) As Boolean` (= `Kinds_IsHorizontal(Registry_ChartType(cht))`); `Geom_ChartCategoryCenterY(cht, idx)`; `Geom_ChartValueToX(cht, v)` (eixo `Axes(xlValue, xlPrimary)`); `Geom_ChartCatCoord(cht, idx) As Double` (x nas colunas, y nas barras); `Geom_ChartValCoord(cht, v) As Double` (y nas colunas, x nas barras); `Geom_ChartBarHalfWidth` usa `InsideHeight` em horizontais. `Geom_ChartCategoryCenterX`/`Geom_ChartValueToY` continuam como estão.
- Produces (`modAnnotations`): `Ann_GetCategoryData` cobre `barstacked/barstacked100` (como empilhadas; topo 1 em 100%), `barclustered` (como `clustered`), `barwaterfall` (como `waterfall`), `area` (soma dos positivos no topo), `area100` (topo 1), `combo` (soma das séries de coluna: todas as não-auxiliares **menos a última**). `Ann_AddCagr`/`Ann_AddDifference` recusam por `Kinds_Supports(kind, "cagr"/"diff")` com `vbObjectError + 7001` e mensagens exatas: 100% → `"A seta de diferença não está disponível para gráficos 100%."` (texto atual, mantido); outros tipos → `"A seta de diferença não está disponível para este tipo de gráfico."` e `"A seta CAGR não está disponível para este tipo de gráfico."`. `Ann_Refresh` valida `cagr/diff` por `Kinds_Supports` (oculta a forma quando o tipo não suporta). Layouts por orientação (abaixo). `Ann_EnsureRoom(cht)` (renomeia `Ann_EnsureHeadroom`, privada): vertical como hoje (52 pt acima quando totais ligados; chamada só por `Ann_AddCagr`, como hoje); horizontal garante `ChartArea.Width - (InsideLeft + InsideWidth) >= 60` reduzindo `PlotArea.Width` e, quando o gráfico tem alguma forma `tqq_ann_diff_*`, também `ChartArea.Height - (InsideTop + InsideHeight) >= 34` reduzindo `PlotArea.Height` (espaço para a seta de diferença abaixo da última barra e para o rótulo dela: 10 até a seta + 6 + rótulo de ~16 + 2); as duas só aumentam a folga. Em horizontais, `Ann_AddCagr` **e** `Ann_AddDifference` chamam `Ann_EnsureRoom` depois de criar as formas e antes de `Ann_Refresh`.

Layouts horizontais (coordenadas: `c = Geom_ChartCatCoord`, `v = Geom_ChartValCoord`, `half = Geom_ChartBarHalfWidth`, `extra = 18` se `totals=1`):
- **CAGR**: de `(v(tops(from)) + 10 + extra, c(from))` a `(v(tops(to)) + 10 + extra, c(to))` (seta ao longo das categorias, ponta em `to`); x limitado a `ChartArea.Width - 24`; rótulo à direita: `lbl.Left = max(x1, x2) + 4`, `lbl.Top = (c(from) + c(to)) / 2 - lbl.Height / 2`; `Ann_ClampShape`.
- **Diferença**: `yArrow = c(to) + half + 10` (abaixo da barra final; limitado a `ChartArea.Height - 2`); linha de nível **vertical** de `(v(tops(from)), c(from) + half)` a `(v(tops(from)), yArrow + 4)`; seta dupla **horizontal** de `(v(tops(from)), yArrow)` a `(v(tops(to)), yArrow)`; rótulo centralizado abaixo: `lbl.Left = (x1 + x2) / 2 - lbl.Width / 2`, `lbl.Top = yArrow + 6`; texto igual ao vertical (`+45 (+300%)`).
- **Conector (cascata)**: vertical em `x = v(tops(i))`, de `c(i) + half` a `c(j) - half`.
- Vertical: idêntico ao atual, reescrito sobre `Geom_ChartCatCoord/ValCoord` (os testes da Fase 1 em `TestAnnotations` são a rede de segurança).

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestGeometry.bas`:

```vb
Public Sub Test_Geom_CategoryCenterY_And_ValueToX()
    AssertAlmost 125, Geom_CategoryCenterY(100, 200, 4, 1)     ' slot de 50: categoria 1 no topo, centro em 100+25
    AssertAlmost 275, Geom_CategoryCenterY(100, 200, 4, 4)
    AssertAlmost 50, Geom_ValueToX(50, 300, 0, 100, 0)
    AssertAlmost 350, Geom_ValueToX(50, 300, 0, 100, 100)
    AssertAlmost 200, Geom_ValueToX(50, 300, -100, 100, 0)
    AssertAlmost 50, Geom_ValueToX(50, 300, 5, 5, 5)           ' escala degenerada
End Sub
```

`tests/vba/core/TestAnnotations.bas` (fixture "crua" de barras: bloco com tipo `barstacked` e gráfico nomeado, sem depender da Tarefa 3):

```vb
' Grafico de barras cru registrado como barstacked: bloco com categorias e totais, uma serie, estilo base.
Private Function RawBarFixture(wb As Workbook) As Chart
    Dim ws As Worksheet, blk As Range, body As Range, co As ChartObject, s As Series, i As Long
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Cat", "Valor")
    ws.Range("A2:B2").Value = Array("a", 10): ws.Range("A3:B3").Value = Array("b", 25): ws.Range("A4:B4").Value = Array("c", 40)
    Set blk = CalcSheet_AllocateBlock(wb, "geo001", "barstacked", "'Dados'!$A$1:$B$4", 4, 4)
    Set body = Block_Body(blk)
    body.Cells(1, 1).Value = "Categoria": body.Cells(1, 2).Value = "Total"
    For i = 1 To 3
        body.Cells(i + 1, 1).Formula = "=Dados!$A$" & (i + 1)
        body.Cells(i + 1, 2).Formula = "=Dados!$B$" & (i + 1)
    Next
    Set co = ws.ChartObjects.Add(150, 10, 460, 280)
    co.Name = ID_PREFIX & "geo001"
    Set s = co.Chart.SeriesCollection.NewSeries
    s.Name = "=Dados!$B$1": s.Values = "=Dados!$B$2:$B$4": s.XValues = "=Dados!$A$2:$A$4"
    co.Chart.ChartType = xlBarStacked
    Style_ApplyBase co.Chart
    Registry_SetOption co.Chart, "numfmt", "int"
    Set RawBarFixture = co.Chart
End Function

Public Sub Test_Geom_ChartCoords_HorizontalChart()
    Dim wb As Workbook, cht As Chart, ax As Axis, il As Double, it As Double, iw As Double, ih As Double
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    AssertTrue Geom_ChartIsHorizontal(cht)
    il = cht.PlotArea.InsideLeft: it = cht.PlotArea.InsideTop: iw = cht.PlotArea.InsideWidth: ih = cht.PlotArea.InsideHeight
    Set ax = cht.Axes(xlValue)
    AssertAlmost it + ih / 6, Geom_ChartCatCoord(cht, 1), 0.01, "categoria 1 no topo"
    AssertAlmost it + ih * 5 / 6, Geom_ChartCatCoord(cht, 3), 0.01
    AssertAlmost il + iw * (25 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale), Geom_ChartValCoord(cht, 25), 0.01
    AssertAlmost (ih / 3) / (1 + cht.ChartGroups(1).GapWidth / 100) / 2, Geom_ChartBarHalfWidth(cht), 0.01, "meia largura pela altura interna"
    CloseTemp wb
End Sub

Public Sub Test_Ann_Cagr_Horizontal_RunsAlongCategoriesRightOfBars()
    Dim wb As Workbook, cht As Chart, ln As Shape, lbl As Shape, ax As Axis, xEnd As Double
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    Set ln = Ann_AddCagr(cht, 1, 3)
    Set ax = cht.Axes(xlValue)
    xEnd = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (40 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
    AssertAlmost Geom_ChartCatCoord(cht, 1), ln.Top, 1, "comeca na faixa da categoria 1"
    AssertAlmost Geom_ChartCatCoord(cht, 3) - Geom_ChartCatCoord(cht, 1), ln.Height, 1
    AssertAlmost xEnd + 10, ln.Left + ln.Width, 1.5, "fim da seta 10 pt a direita da barra mais longa"
    AssertTrue ln.VerticalFlip = msoFalse And ln.HorizontalFlip = msoFalse, "seta desce para a direita (ponta em 'c')"
    Set lbl = cht.Shapes("tqq_ann_cagr_1_lbl")
    AssertEqual "CAGR +100,0%", lbl.TextFrame2.TextRange.Text      ' (40/10)^(1/2)-1
    AssertTrue lbl.Left >= ln.Left + ln.Width + 2, "rotulo a direita da seta"
    AssertTrue lbl.Left + lbl.Width <= cht.ChartArea.Width + 0.5, "dentro do grafico"
    CloseTemp wb
End Sub

Public Sub Test_Ann_Diff_Horizontal_BelowLastBar()
    Dim wb As Workbook, cht As Chart, ln As Shape, lvl As Shape, lbl As Shape
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    Set ln = Ann_AddDifference(cht, 1, 3)
    Set lvl = cht.Shapes("tqq_ann_diff_1_lvl"): Set lbl = cht.Shapes("tqq_ann_diff_1_lbl")
    AssertAlmost 0, ln.Height, 0.01, "seta de diferenca horizontal"
    AssertTrue ln.Top >= Geom_ChartCatCoord(cht, 3) + Geom_ChartBarHalfWidth(cht) + 8, "abaixo da ultima barra"
    AssertAlmost Geom_ChartValCoord(cht, 10), ln.Left, 1
    AssertAlmost Geom_ChartValCoord(cht, 40) - Geom_ChartValCoord(cht, 10), ln.Width, 1
    AssertAlmost 0, lvl.Width, 0.01, "linha de nivel vertical"
    AssertAlmost Geom_ChartValCoord(cht, 10), lvl.Left, 1
    AssertEqual "+30 (+300%)", lbl.TextFrame2.TextRange.Text
    AssertTrue lbl.Top >= ln.Top + 4, "rotulo abaixo da seta"
    AssertTrue cht.ChartArea.Height - (cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight) >= 33.5, "folga abaixo da area de plotagem (Ann_EnsureRoom)"
    AssertTrue lbl.Top + lbl.Height <= cht.ChartArea.Height + 0.5, "rotulo da diferenca dentro do grafico"
    CloseTemp wb
End Sub

Public Sub Test_Ann_Refusals_ByKind()
    Dim wb As Workbook, cht As Chart, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    Registry_Block(cht).Cells(1, 2).Value = "pie"            ' finge outro tipo: a recusa vem da matriz
    On Error Resume Next
    Ann_AddCagr cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    AssertEqual vbObjectError + 7001, en
    AssertEqual "A seta CAGR não está disponível para este tipo de gráfico.", ed
    Ann_AddDifference cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7001, en
    AssertEqual "A seta de diferença não está disponível para este tipo de gráfico.", ed
    Registry_Block(cht).Cells(1, 2).Value = "barstacked100"
    On Error Resume Next
    Ann_AddDifference cht, 1, 3
    ed = Err.Description: Err.Clear
    On Error GoTo 0
    AssertEqual "A seta de diferença não está disponível para gráficos 100%.", ed
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-f2\build\run-tests.ps1 -Target core -Filter "TestGeometry|Horizontal|Refusals" -TimeoutSec 240` → saída 3 (`Geom_ChartCatCoord` etc. indefinidos).

- [ ] **Passo 3: Implementar** — em `modGeometry` as funções novas; em `modAnnotations`, reescrever `Ann_LayoutCagr/Diff/Conn` com um ramo `If Geom_ChartIsHorizontal(cht) Then ... Else ... End If` por função (não tente um layout "genérico" com troca de eixos: os dois ramos explícitos são mais fáceis de revisar). `Ann_GetCategoryData`:

```vb
    Select Case kind
        Case "waterfall", "barwaterfall": ' End do bloco (como hoje)
        Case "clustered", "barclustered", "line": ' serie 1
        Case "combo": ' soma de todas as nao-auxiliares menos a ultima (a linha)
        Case Else: ' soma (stacked, stacked100, barstacked, barstacked100, area, area100); topos = 1 se Kinds_Is100(kind)
    End Select
```

- [ ] **Passo 4: GREEN** — filtro verde; depois a suíte completa: **todos** os testes antigos de `TestAnnotations`/`TestSamples` (CAGR acima dos totais, conectores da cascata, 100%) continuam verdes.

- [ ] **Passo 5: Commit** — `git commit -m "feat: geometria e setas por orientacao (colunas e barras), recusas de anotacao pela matriz de tipos"` (+ trailer).

**Verificação no Excel real:** nada visível ainda; conferir no gráfico de exemplo 1 que CAGR e diferença continuam no mesmo lugar (regressão visual).

---

### Tarefa 3: Barras horizontais (empilhadas, agrupadas, 100%)

**Files:**
- Modify: `src/core/modules/modChartStacked.bas`, `src/core/modules/modLabels.bas` (`Labels_ApplyCellFormat` em 100%), `src/core/modules/modCoreApi.bas` (`Core_SafeRefresh` chama `Chart_SyncSecondaryX` em horizontais)
- Test: `tests/vba/core/TestBars.bas` (novo), `tests/vba/core/TestCoreApi.bas`

**Interfaces:**
- Consumes: `Chart_AddHelperSeries`, `Chart_SetupSecondaryAxes`, `Chart_SyncSecondaryX`, `Chart_HideHelperLegendEntries`, `Labels_LinkPoints(..., horizontal)`, `Kinds_*`, `Geom_*`, `Ann_*` (Tarefas 1–2).
- Produces: `Stacked_BuildBlock(wb, id, data, kind)` com colunas `1 Categoria | 2 Total | 3 TotalPos | [4..3+S Part_j só em 100%] | P Pos | P+1 VLine` para **todos** os tipos da família (`Pos = n - i + 0.5` como fórmula `=2.5`; `VLine = AVERAGE(coluna Total)` com referências absolutas dentro de `_ThinQQ`; `=NA()` em 100%); `Stacked_PosCol(blk) As Long` (= `blk.Columns.Count - 1`), `Stacked_VLineCol(blk) As Long` (= `blk.Columns.Count`); `Stacked_Insert(ws, data, kind)` aceita `barstacked` (`xlBarStacked`, gap 50), `barclustered` (`xlBarClustered`, gap 80), `barstacked100` (`xlBarStacked100`, gap 50); `Stacked_SetTotals` cria `tqq_total` por `Chart_AddHelperSeries` (em horizontais: XY secundária com `X = TotalPos`, `Y = Pos`) e liga os rótulos com `horizontal:=Kinds_IsHorizontal(kind)`; `Stacked_SetLabels` mantém as posições (Center em empilhadas/100%, OutsideEnd em agrupadas); `Stacked_SetConnectors` aceita `xlBarStacked/xlBarStacked100`; `Labels_ApplyCellFormat` em 100% formata só `Part_j` (`Columns.Count - 5` colunas a partir da 4).
- Produces: `Core_SafeRefresh` chama `Chart_SyncSecondaryX cht` quando `Kinds_IsHorizontal(Registry_ChartType(cht))`.
- Comparações da Fase 1 que esta tarefa troca (sem elas os testes abaixo falham): `Stacked_SetTotals` recusa por `Not Kinds_Supports(kind, "totals")` (hoje `<> "stacked" And <> "stacked100"`) e `Stacked_Insert` liga os totais com `Kinds_Supports(kind, "totals")` (hoje `kind = "stacked" Or kind = "stacked100"`); `Stacked_SetLabels` usa `is100 = Kinds_Is100(kind)` (participações também em `barstacked100`) e põe `xlLabelPositionOutsideEnd` também em séries `xlBarClustered` (hoje só `xlColumnClustered`); `Stacked_BuildBlock` usa `Kinds_Is100(kind)` para as colunas `Part_j` e `TotalPos = Total` também em `barclustered` (como `clustered`/`line`); `Stacked_Insert`: gap 80 em `clustered`/`barclustered`, 50 nos demais (exceto `line`).

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestBars.bas`:

```vb
Attribute VB_Name = "TestBars"
Option Explicit

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Regiao", "Loja", "Online")
    ws.Range("A2:C2").Value = Array("Sul", 10, 5)
    ws.Range("A3:C3").Value = Array("Sudeste", 25, 10)
    ws.Range("A4:C4").Value = Array("Norte", 40, 20)
    Set Fixture = ChartData_FromRange(ws.Range("A1:C4"))
End Function

Private Function IsScatter(ct As Long) As Boolean
    IsScatter = (ct = xlXYScatter Or ct = xlXYScatterLines Or ct = xlXYScatterLinesNoMarkers Or ct = xlXYScatterSmooth Or ct = xlXYScatterSmoothNoMarkers)
End Function

Public Sub Test_Bars_Insert_KindsAndAxes()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tot As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    AssertEqual "barstacked", Registry_ChartType(cht)
    AssertEqual xlBarStacked, cht.SeriesCollection(1).ChartType
    AssertTrue cht.Axes(xlCategory).ReversePlotOrder, "primeira categoria no topo"
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertTrue Not tot Is Nothing
    AssertTrue IsScatter(tot.ChartType), "totais em XY (barras nao combinam com linha)"
    AssertEqual xlSecondary, tot.AxisGroup
    AssertEqual 0, cht.Axes(xlValue, xlSecondary).MinimumScale
    AssertEqual 3, cht.Axes(xlValue, xlSecondary).MaximumScale
    AssertEqual xlTickLabelPositionNone, cht.Axes(xlValue, xlSecondary).TickLabelPosition
    AssertEqual "50", Registry_GetOption(cht, "gap"): AssertEqual "3", Registry_GetOption(cht, "n")
    AssertEqual "1", Registry_GetOption(cht, "totals"): AssertEqual "1", Registry_GetOption(cht, "labels")
    AssertEqual 2, cht.Legend.LegendEntries.Count, "tqq_total fora da legenda"
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barclustered").Chart
    AssertEqual xlBarClustered, cht.SeriesCollection(1).ChartType
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual xlLabelPositionOutsideEnd, cht.SeriesCollection(1).DataLabels.Position
    AssertEqual "80", Registry_GetOption(cht, "gap")
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked100").Chart
    AssertEqual xlBarStacked100, cht.SeriesCollection(1).ChartType
    AssertEqual "67%", cht.SeriesCollection(1).Points(1).DataLabel.Text, "segmento 100% mostra a participacao (10/15)"
    AssertEqual "15", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    Stacked_SetConnectors cht, True
    AssertTrue cht.ChartGroups(1).HasSeriesLines
    CloseTemp wb
End Sub

Public Sub Test_Bars_Block_PosAndVLineColumns()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range, blk As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 5, blk.Columns.Count
    AssertEqual 2.5, body.Cells(2, Stacked_PosCol(blk)).Value: AssertEqual 0.5, body.Cells(4, Stacked_PosCol(blk)).Value
    AssertAlmost 36.666667, body.Cells(2, Stacked_VLineCol(blk)).Value, 0.0001, "media dos totais 15, 35, 60"
    AssertAlmost 36.666667, body.Cells(4, Stacked_VLineCol(blk)).Value, 0.0001
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked100").Chart
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 7, blk.Columns.Count, "3 + 2 participacoes + Pos + VLine"
    AssertTrue IsError(body.Cells(2, Stacked_VLineCol(blk)).Value), "100%: sem linha de valor"
    CloseTemp wb
End Sub

' Sonda e regra de ouro das barras: o rotulo do total i fica no fim da barra i (eixo de valores primario) e na faixa i.
Public Sub Test_Bars_TotalsHelper_AlignedWithBarsAndSlots()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tot As Series, ax As Axis, i As Long
    Dim il As Double, it As Double, iw As Double, ih As Double, xEnd As Double, yc As Double, lbl As DataLabel, totals As Variant
    totals = Array(15, 35, 60)
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Set tot = Chart_FindSeries(cht, "tqq_total")
    il = cht.PlotArea.InsideLeft: it = cht.PlotArea.InsideTop: iw = cht.PlotArea.InsideWidth: ih = cht.PlotArea.InsideHeight
    AssertTrue iw > 50 And ih > 50, "area interna real"
    Set ax = cht.Axes(xlValue, xlPrimary)
    For i = 1 To 3
        xEnd = il + iw * (totals(i - 1) - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
        yc = it + (i - 0.5) * ih / 3
        Set lbl = tot.Points(i).DataLabel
        AssertTrue lbl.Left >= xEnd - 2 And lbl.Left <= xEnd + 14, "rotulo " & i & " a direita do fim da barra: left=" & lbl.Left & " fim=" & xEnd
        AssertTrue Abs((lbl.Top + lbl.Height / 2) - yc) <= ih / 6, "rotulo " & i & " na faixa da categoria: centro=" & (lbl.Top + lbl.Height / 2) & " esperado=" & yc
    Next
    CloseTemp wb
End Sub

' Review Focus 1: depois de mudar dados (nova escala) e redimensionar, os totais continuam no fim das barras.
Public Sub Test_Bars_TotalsHelper_StaysAlignedAfterDataChangeAndResize()
    Dim wb As Workbook, d As clsChartData, co As ChartObject, cht As Chart, tot As Series, ax As Axis
    Dim il As Double, iw As Double, xEnd As Double, lbl As DataLabel
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set co = Stacked_Insert(d.Source.Worksheet, d, "barstacked")
    Set cht = co.Chart
    d.Source.Worksheet.Range("C4").Value = 200           ' total Norte = 240: eixo rescala
    co.Width = 620
    Core_SafeRefresh cht
    Set tot = Chart_FindSeries(cht, "tqq_total")
    Set ax = cht.Axes(xlValue, xlPrimary)
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth
    xEnd = il + iw * (240 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
    Set lbl = tot.Points(3).DataLabel
    AssertEqual "240", lbl.Text
    AssertTrue lbl.Left >= xEnd - 2 And lbl.Left <= xEnd + 14, "total 3 segue o fim da barra apos mudanca: left=" & lbl.Left & " fim=" & xEnd
    xEnd = il + iw * (15 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
    Set lbl = tot.Points(1).DataLabel
    AssertTrue lbl.Left >= xEnd - 2 And lbl.Left <= xEnd + 14, "total 1 segue a nova escala"
    CloseTemp wb
End Sub

Public Sub Test_Bars_TotalLabels_RightLeftBySign()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tot As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("B2:C2").Value = Array(-10, -5)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertEqual xlLabelPositionLeft, tot.Points(1).DataLabel.Position, "total negativo a esquerda"
    AssertEqual xlLabelPositionRight, tot.Points(2).DataLabel.Position
    d.Source.Worksheet.Range("B2:C2").Value = Array(10, 5)
    Labels_RefreshPositions cht
    AssertEqual xlLabelPositionRight, tot.Points(1).DataLabel.Position
    CloseTemp wb
End Sub

Public Sub Test_Bars_Annotations_UseTotals()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Ann_AddCagr cht, 1, 3
    Ann_AddDifference cht, 1, 3
    AssertEqual "CAGR +100,0%", cht.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text     ' (60/15)^(1/2)-1
    AssertEqual "+45 (+300%)", cht.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text
    AssertTrue cht.Shapes("tqq_ann_cagr_1").Left + cht.Shapes("tqq_ann_cagr_1").Width <= cht.ChartArea.Width, "CAGR dentro do grafico (folga a direita)"
    CloseTemp wb
End Sub

Public Sub Test_Bars_Insert_RollsBackOnFailure()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, n As Long, nm As Name
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Stacked_Insert ws, d, "barclustered"
    n = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, n
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    CloseTemp wb
End Sub
```

Em `TestCoreApi.bas`:

```vb
Public Sub Test_Core_Insert_Bars_ThroughApi()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "barstacked"
    Set cht = Core_ActiveChart()
    AssertEqual "barstacked", Registry_ChartType(cht)
    AssertTrue Core_GetToggle("totals")
    Core_SetToggle "totals", False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    Core_SetToggle "totals", True
    AssertTrue Not Chart_FindSeries(cht, "tqq_total") Is Nothing
    Core_SetOption "numfmt", "dec1"
    AssertEqual "150" & Application.International(xlDecimalSeparator) & "0", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    Core_SetOption "gap", "120"
    AssertEqual 120, cht.ChartGroups(1).GapWidth
    Teardown wb
End Sub
```

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-f2\build\run-tests.ps1 -Target core -Filter "TestBars|Insert_Bars" -TimeoutSec 240` → saída 3 (`Stacked_PosCol` indefinido) e, após declarar, `not ok` nos inserts de barras.

- [ ] **Passo 3: Implementar** — em `Stacked_BuildBlock`, calcule `cols = 3 + IIf(Kinds_Is100(kind), S, 0) + 2` e escreva `Pos`/`VLine` nas últimas duas colunas; em `Stacked_Insert`, `Select Case kind` com os três tipos novos e `Registry_SetOption cht, "n", ...`; `Stacked_SetTotals`:

```vb
    If s Is Nothing Then
        If Kinds_IsHorizontal(kind) Then
            Set s = Chart_AddHelperSeries(cht, "tqq_total", body.Cells(2, 3).Resize(n, 1), body.Cells(2, Stacked_PosCol(blk)).Resize(n, 1))
        Else
            Set s = Chart_AddHelperSeries(cht, "tqq_total", body.Cells(2, 3).Resize(n, 1))
        End If
    End If
    Chart_HideHelperLegendEntries cht
    Labels_LinkPoints s, body.Cells(2, 2).Resize(n, 1), fmt, , , Kinds_IsHorizontal(kind)
```

**Decisão da sonda (Passo 3b):** se `Test_Bars_TotalsHelper_AlignedWithBarsAndSlots` falhar no eixo X (rótulos fora do fim das barras) com o eixo X secundário não exibido, ative em `Chart_SetupSecondaryAxes` `cht.HasAxis(xlCategory, xlSecondary) = True` e deixe `Chart_SyncSecondaryX` copiar `Min/Max` do primário (oculto); `Core_SafeRefresh` já o chama em horizontais. Registre no relatório qual caminho foi necessário, com os números medidos.

- [ ] **Passo 4: GREEN e suíte completa**; `-LoaderSwap` não é necessário aqui. Exporte `cht.Export TempDir() & "\bars.png"` de um `barstacked` e de um `barstacked100` e confira na imagem que a legenda mostra só `Loja` e `Online` (mapeamento de `Chart_HideHelperLegendEntries` com a auxiliar XY no grupo secundário) e que os totais ficam no fim das barras; registre no relatório o que viu.

- [ ] **Passo 5: Commit** — `git commit -m "feat: barras horizontais empilhadas, agrupadas e 100% com totais em serie XY, rotulos por sinal e setas"` (+ trailer).

**Verificação no Excel real:** inserir os três tipos de barras a partir de uma tabela 3x3; totais no fim das barras; mudar um valor e ver os totais acompanharem; redimensionar; Seta CAGR e Seta diferença; Formato dec1; 100% com participações e conectores.

---

### Tarefa 4: Cascata com várias séries e cascata "de cima para baixo" (`e` na primeira linha)

**Files:**
- Modify: `src/core/modules/modChartWaterfall.bas`, `src/core/modules/modLabels.bas` (`Labels_LinkPoints` isola cada ponto: um ponto `NA()` não interrompe a ligação dos demais), `src/core/modules/modCoreApi.bas` (`Core_Insert` trata a faixa de validação 6007..6099 como aviso)
- Test: `tests/vba/core/TestWaterfall2.bas` (novo), `tests/vba/core/TestWaterfall.bas` (um teste atualizado)

**Interfaces:**
- Consumes: `Chart_AddHelperSeries`, `Chart_HideHelperLegendEntries`, `Block_Ref`, `Kinds_*`, `Style_Palette`, `Style_ContrastFont`, `Labels_LinkPoints`.
- Produces (constantes): `WF_COL_POS = 19`, `WF_COL_BD = 20`, `WF_MS_BASE = 21`, `WF_MS_WIDTH = 10`, `WF_MAX_SERIES = 6`; `Waterfall_SeriesCol(j As Long, part As String) As Long` com `part` em `D, CP, CN, A, B, M, UP, DN, UPN, DNN` (= `WF_MS_BASE + WF_MS_WIDTH*(j-1) + índice da parte`, `D` = 0 … `DNN` = 9); `Waterfall_SeriesCount(cht) As Long` (= `max(1, (blk.Columns.Count - 20) \ 10)`).
- Produces: `Waterfall_BuildBlock(wb, id, data)` com 20 colunas (S = 1) ou `20 + 10S` (S > 1), fórmulas do design §4.3; **build-down** nas colunas 3/6/20; `Waterfall_Insert(ws, data)` monta, para S > 1, as séries na ordem `tqq_base, tqq_dn_S..tqq_dn_1, <nome da série 1>..<nome da série S> (UP_j), tqq_upn_S..tqq_upn_1, tqq_dnn_1..tqq_dnn_S, tqq_total`, liga `HasLegend = True` com `Chart_HideHelperLegendEntries` e `Registry_SetOption "legend","right"`; com S = 1 o gráfico é idêntico ao da Fase 1 (`HasLegend = False`). Recusa S > 6 com `Err.Raise vbObjectError + 6007, "Waterfall_Insert", "A cascata aceita no máximo 6 séries."` **antes** de alocar o bloco. `Waterfall_SetLabels(cht, on)`: S = 1 como hoje; S > 1 cria `tqq_lbl_j` (auxiliar em `M_j`, rótulos ligados a `D_j`, posição Center, cor de fonte por contraste com a cor da série) e `tqq_lbl` (em `LabelPos`, ligado a `LabelVal`, só linhas de total). `Waterfall_ApplyColors(cht, palette)`: S = 1 como hoje; S > 1: toda série da série `j` (UP/DN/UPN/DNN) recebe `pal((j-1) Mod 6)` e `tqq_total` recebe `pal(3)` (cinza). `Waterfall_EndValues` inalterada (col 6).
- Mensagem via `Core_Insert`: acrescente ao `Fail:` de `Core_Insert` o ramo `If Err.Number >= vbObjectError + 6007 And Err.Number <= vbObjectError + 6099 Then Core_Notify Err.Description, vbExclamation` (faixa reservada a **erros de validação dos dados** dos tipos novos: 6007 cascata, 6008 combinação, 6009 Mekko, 6010 Gantt; os 6001..6006 continuam sendo falhas estruturais tratadas por `Errors_Handle`). Texto: `"A cascata aceita no máximo 6 séries."`.

Fórmulas por série `j` (R1C1 relativo com referências absolutas de coluna; `Start` = col 5, `IsTotal` = col 3; `CP_0/CN_0` = 0 — use `IF(j=1, 0, célula)` ao montar a string em VBA, não `MAX` sobre células vazias):

```
D_j   = IF(RC3, 0, IF(ISNUMBER(src_j), src_j, 0))
CP_j  = CP_{j-1} + MAX(0, D_j)               CN_j = CN_{j-1} + MAX(0, -D_j)
A_j   = IF(D_j >= 0, RC5 + CP_{j-1}, RC5 - CN_j)
B_j   = IF(D_j >= 0, RC5 + CP_j, RC5 - CN_{j-1})
M_j   = (A_j + B_j) / 2
UP_j  = IF(D_j >= 0, MAX(0,B_j) - MAX(0,A_j), 0)      DN_j  = IF(D_j < 0, MAX(0,B_j) - MAX(0,A_j), 0)
UPN_j = IF(D_j >= 0, MIN(0,A_j) - MIN(0,B_j), 0)      DNN_j = IF(D_j < 0, MIN(0,A_j) - MIN(0,B_j), 0)
col 4 Delta = SUM(D_1..D_S)      col 7 Lo = IF(RC3, MIN(0,RC6), RC5 - CN_S)      col 8 Hi = IF(RC3, MAX(0,RC6), RC5 + CP_S)
col 3 IsTotal (S>1) = SUMPRODUCT(--ISNUMBER(MATCH(LOWER(TRIM(src_1:src_S&"")),{"e","t","=","total","subtotal"},0)))>0
      (sem IFERROR: dentro de SUMPRODUCT ele não é avaliado como matriz no Excel 2019 sem Ctrl+Shift+Enter; um erro na
       origem já vira FALSE em ISNUMBER(MATCH(...)))
col 17 LabelVal (S>1) = IF(RC3, RC16, NA())      col 18 LabelPos (S>1) = IF(RC3, IF(RC16>=0, RC8, RC7), NA())
col 19 Pos = n - i + 0.5 (constante por linha)
col 20 BD = AND(<marcador na linha 1>, COUNT(<células de valor da linha n>)>0, NOT(<marcador na linha n>))   (mesma fórmula em todas as linhas)
col 3 IsTotal, última linha: = OR(<marcador>, RC20)
col 6 End, linha 1: = IF(RC20, AGGREGATE(9,6,<células de valor da linha n>) - SUM(R[1]C4:R[n-2]C4), RC5 + RC4)   (com n = 2 a SUM cobre só linhas de total, que valem 0)
col 6 End, última linha: = IF(RC20, AGGREGATE(9,6,<células de valor da linha n>), RC5 + RC4)
```
Com S = 1 as colunas 10..15 continuam alimentando o gráfico; com S > 1 elas ficam calculadas mas não são usadas pelas séries (`UP/DN/UPN/DNN` são).

`<marcador na linha k>` é uma expressão sobre as **células de origem** da linha k (referências absolutas), **nunca** sobre a célula `IsTotal` (col 3): `IsTotal` da última linha depende de `BD` (`OR(<marcador>, RC20)`), então usar `IsTotal` dentro de `BD` cria referência circular (também com n = 1, em que a linha 1 é a última). S = 1: `IF(ISERROR(src),FALSE,IF(ISNUMBER(src),FALSE,OR(LOWER(TRIM(src&""))={<lista>})))` (a expressão da col 3 da Fase 1, com a origem da linha k); S > 1: `SUMPRODUCT(--ISNUMBER(MATCH(LOWER(TRIM(src_1:src_S&"")),{<lista>},0)))>0` sobre a linha k. A `<lista>` (constante de matriz) é montada em VBA a partir de `WF_MARKERS`, como a Fase 1 já faz na col 3 (`Test_Waterfall_MarkerList_FormulaMatchesFunction` continua valendo).

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestWaterfall2.bas`:

```vb
Attribute VB_Name = "TestWaterfall2"
Option Explicit
' Fase 2: cascata com varias series, build-down ('e' na primeira linha) e cascata horizontal.

Private Function MultiFixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:C1").Value = Array("Etapa", "Loja", "Online")
    ws.Range("A2:C2").Value = Array("Inicio", 100, 50)
    ws.Range("A3:C3").Value = Array("Ajuste", -30, 10)
    ws.Range("A4:C4").Value = Array("Subtotal", "e", "e")
    ws.Range("A5:C5").Value = Array("Queda", -200, -20)
    ws.Range("A6:B6").Value = Array("Fim", "e")
    Set MultiFixture = ChartData_FromRange(ws.Range("A1:C6"))
End Function

' Review Focus 2: coluna com sinais misturados que cruza o zero (Queda: inicio 130, -200 e -20).
Public Sub Test_WaterfallMulti_Block_MixedSignsCrossingZero()
    Dim wb As Workbook, body As Range, blk As Range
    Set wb = NewTempWorkbook()
    Set blk = Waterfall_BuildBlock(wb, "wfm001", MultiFixture(wb)): Set body = Block_Body(blk)
    AssertEqual 40, blk.Columns.Count, "20 + 10 por serie"
    ' End: 150, 130, 130, -90, -90
    AssertEqual 150, body.Cells(2, WF_COL_END).Value: AssertEqual 130, body.Cells(3, WF_COL_END).Value
    AssertEqual 130, body.Cells(4, WF_COL_END).Value: AssertEqual -90, body.Cells(5, WF_COL_END).Value
    AssertEqual -90, body.Cells(6, WF_COL_END).Value
    ' Inicio: duas series positivas empilhadas a partir de 0
    AssertEqual 100, body.Cells(2, Waterfall_SeriesCol(1, "UP")).Value: AssertEqual 50, body.Cells(2, Waterfall_SeriesCol(2, "UP")).Value
    AssertEqual 50, body.Cells(2, Waterfall_SeriesCol(1, "M")).Value: AssertEqual 125, body.Cells(2, Waterfall_SeriesCol(2, "M")).Value
    ' Ajuste: Loja desce 30 a partir de 150 (120..150), Online sobe 10 (150..160); Base = 120
    AssertEqual 120, body.Cells(3, Waterfall_SeriesCol(1, "A")).Value: AssertEqual 150, body.Cells(3, Waterfall_SeriesCol(1, "B")).Value
    AssertEqual 30, body.Cells(3, Waterfall_SeriesCol(1, "DN")).Value: AssertEqual 0, body.Cells(3, Waterfall_SeriesCol(1, "UP")).Value
    AssertEqual 150, body.Cells(3, Waterfall_SeriesCol(2, "A")).Value: AssertEqual 160, body.Cells(3, Waterfall_SeriesCol(2, "B")).Value
    AssertEqual 10, body.Cells(3, Waterfall_SeriesCol(2, "UP")).Value
    AssertEqual 120, body.Cells(3, WF_COL_BASE).Value: AssertEqual 120, body.Cells(3, 7).Value: AssertEqual 160, body.Cells(3, 8).Value
    ' Subtotal: total 130, sem segmentos
    AssertEqual True, body.Cells(4, 3).Value: AssertEqual 130, body.Cells(4, WF_COL_TOTAL).Value
    AssertEqual 0, body.Cells(4, Waterfall_SeriesCol(1, "DN")).Value: AssertEqual 130, body.Cells(4, WF_COL_LABELVAL).Value
    ' Queda: Loja -200 cruza o zero (-70..130): DN 130 acima, DNN -70 abaixo; Online -20 (-90..-70): DNN -20
    AssertEqual -70, body.Cells(5, Waterfall_SeriesCol(1, "A")).Value: AssertEqual 130, body.Cells(5, Waterfall_SeriesCol(1, "B")).Value
    AssertEqual 130, body.Cells(5, Waterfall_SeriesCol(1, "DN")).Value: AssertEqual -70, body.Cells(5, Waterfall_SeriesCol(1, "DNN")).Value
    AssertEqual -90, body.Cells(5, Waterfall_SeriesCol(2, "A")).Value: AssertEqual -70, body.Cells(5, Waterfall_SeriesCol(2, "B")).Value
    AssertEqual 0, body.Cells(5, Waterfall_SeriesCol(2, "DN")).Value: AssertEqual -20, body.Cells(5, Waterfall_SeriesCol(2, "DNN")).Value
    AssertEqual 0, body.Cells(5, WF_COL_BASE).Value: AssertEqual -90, body.Cells(5, 7).Value: AssertEqual 130, body.Cells(5, 8).Value
    AssertEqual -220, body.Cells(5, 4).Value, "delta liquido"
    AssertTrue IsError(body.Cells(5, WF_COL_LABELVAL).Value), "linha sem total: rotulo de total NA"
    ' Fim: total -90, rotulo em Lo
    AssertEqual True, body.Cells(6, 3).Value: AssertEqual -90, body.Cells(6, WF_COL_TOTAL).Value: AssertEqual -90, body.Cells(6, WF_COL_LABELPOS).Value
    AssertEqual 4.5, body.Cells(2, WF_COL_POS).Value: AssertEqual False, body.Cells(2, WF_COL_BD).Value
    CloseTemp wb
End Sub

Public Sub Test_WaterfallMulti_Insert_SeriesOrderColorsLegend()
    Dim wb As Workbook, d As clsChartData, cht As Chart, pal As Variant, names As String, s As Series
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "waterfall", Registry_ChartType(cht)
    AssertEqual 2, Waterfall_SeriesCount(cht)
    AssertEqual 13, cht.SeriesCollection.Count, "base + 4 por serie + total + 2 rotulos de serie + rotulo de total"
    For Each s In cht.SeriesCollection: names = names & s.Name & ";": Next
    AssertEqual "tqq_base;tqq_dn_2;tqq_dn_1;Loja;Online;tqq_upn_2;tqq_upn_1;tqq_dnn_1;tqq_dnn_2;tqq_total;tqq_lbl_1;tqq_lbl_2;tqq_lbl;", names
    pal = Style_Palette("padrao")
    AssertEqual CLng(pal(0)), Chart_FindSeries(cht, "Loja").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(0)), Chart_FindSeries(cht, "tqq_dnn_1").Format.Fill.ForeColor.RGB, "partes da serie 1 com a mesma cor"
    AssertEqual CLng(pal(1)), Chart_FindSeries(cht, "tqq_dn_2").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(3)), Chart_FindSeries(cht, "tqq_total").Format.Fill.ForeColor.RGB, "totais em cinza"
    AssertEqual msoFalse, Chart_FindSeries(cht, "tqq_base").Format.Fill.Visible
    AssertTrue cht.HasLegend
    AssertEqual 2, cht.Legend.LegendEntries.Count, "legenda so com Loja e Online"
    AssertEqual "right", Registry_GetOption(cht, "legend")
    AssertEqual "5", Registry_GetOption(cht, "n")
    CloseTemp wb
End Sub

Public Sub Test_WaterfallMulti_Labels_SegmentsAndTotals()
    Dim wb As Workbook, d As clsChartData, cht As Chart, l1 As Series, l2 As Series, lt As Series
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    Set l1 = Chart_FindSeries(cht, "tqq_lbl_1"): Set l2 = Chart_FindSeries(cht, "tqq_lbl_2"): Set lt = Chart_FindSeries(cht, "tqq_lbl")
    AssertEqual "100", l1.Points(1).DataLabel.Text
    AssertEqual "-30", l1.Points(2).DataLabel.Text
    AssertEqual "10", l2.Points(2).DataLabel.Text
    AssertEqual xlLabelPositionCenter, l1.Points(2).DataLabel.Position
    AssertEqual "-200", l1.Points(4).DataLabel.Text: AssertEqual "-20", l2.Points(4).DataLabel.Text
    AssertEqual "130", lt.Points(3).DataLabel.Text: AssertEqual xlLabelPositionAbove, lt.Points(3).DataLabel.Position
    AssertEqual "-90", lt.Points(5).DataLabel.Text: AssertEqual xlLabelPositionBelow, lt.Points(5).DataLabel.Position
    Waterfall_SetLabels cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl_1") Is Nothing And Chart_FindSeries(cht, "tqq_lbl") Is Nothing
    AssertEqual 10, cht.SeriesCollection.Count
    Waterfall_SetLabels cht, True
    AssertEqual 13, cht.SeriesCollection.Count
    ' auxiliares com pontos NA() entram na varredura da limpeza (Tarefa 0): o grafico novo nao pode bloquear nem falhar
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    AssertEqual 1, CalcSheet_Cleanup(wb), "limpeza com a cascata multi remove so o orfao"
    AssertFalse CalcSheet_CleanupReadFailed(), "pontos NA dos rotulos nao sao falha de leitura"
    CloseTemp wb
End Sub

Public Sub Test_WaterfallMulti_Connectors_FollowNetEnd()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ax As Axis, it As Double, ih As Double
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    it = cht.PlotArea.InsideTop: ih = cht.PlotArea.InsideHeight: Set ax = cht.Axes(xlValue)
    AssertAlmost it + ih * (ax.MaximumScale - 150) / (ax.MaximumScale - ax.MinimumScale), cht.Shapes("tqq_ann_conn_1").Top, 1
    AssertAlmost it + ih * (ax.MaximumScale - 130) / (ax.MaximumScale - ax.MinimumScale), cht.Shapes("tqq_ann_conn_2").Top, 1, "conector no liquido (-30+10)"
    AssertAlmost it + ih * (ax.MaximumScale + 90) / (ax.MaximumScale - ax.MinimumScale), cht.Shapes("tqq_ann_conn_4").Top, 1
    CloseTemp wb
End Sub

' 'e' na primeira linha + numero na ultima: o primeiro total e calculado para tras (think-cell), a ultima linha vira total.
Public Sub Test_Waterfall_BuildDown_FirstRowE_LastRowNumber()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, body As Range, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Receita", "e")
    ws.Range("A3:B3").Value = Array("Custos", -700)
    ws.Range("A4:B4").Value = Array("Despesas", -250)
    ws.Range("A5:B5").Value = Array("Lucro", 250)
    Set d = ChartData_FromRange(ws.Range("A1:B5"))
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wfbd01", d))
    AssertEqual True, body.Cells(2, WF_COL_BD).Value
    AssertEqual True, body.Cells(2, 3).Value: AssertEqual 1200, body.Cells(2, WF_COL_END).Value, "Receita = 250 - (-700 - 250)"
    AssertEqual 1200, body.Cells(2, WF_COL_TOTAL).Value: AssertEqual 0, body.Cells(2, WF_COL_BASE).Value
    AssertEqual 500, body.Cells(3, WF_COL_END).Value: AssertEqual 250, body.Cells(4, WF_COL_END).Value
    AssertEqual True, body.Cells(5, 3).Value, "ultima linha vira total": AssertEqual 250, body.Cells(5, WF_COL_TOTAL).Value
    AssertEqual 250, body.Cells(5, WF_COL_LABELVAL).Value
    ws.Range("B5").Value = 300
    AssertEqual 1250, body.Cells(2, WF_COL_END).Value, "recalcula ao mudar o total final"
    Set cht = Waterfall_Insert(ws, d).Chart
    AssertEqual "1" & Application.International(xlThousandsSeparator) & "250", Chart_FindSeries(cht, "tqq_lbl").Points(1).DataLabel.Text, "formato int com separador de milhar"
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_TooManySeries_Refused()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, en As Long, ed As String, nm As Name
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:H1").Value = Array("Etapa", "S1", "S2", "S3", "S4", "S5", "S6", "S7")
    ws.Range("A2:H2").Value = Array("x", 1, 1, 1, 1, 1, 1, 1)
    ws.Range("A3:H3").Value = Array("y", 1, 1, 1, 1, 1, 1, 1)
    Set d = ChartData_FromRange(ws.Range("A1:H3"))
    On Error Resume Next
    Waterfall_Insert ws, d
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 6007, en
    AssertEqual "A cascata aceita no máximo 6 séries.", ed
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nenhum bloco alocado"
    Next
    ws.Range("A1:H3").Select
    Core_ClearLastNotice
    Core_Insert "waterfall"
    AssertEqual "A cascata aceita no máximo 6 séries.", Core_LastNotice()
    Watchers_Stop
    CloseTemp wb
End Sub

Public Sub Test_WaterfallMulti_Insert_RollsBackOnFailure()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, n As Long, nm As Name
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb): Set ws = d.Source.Worksheet
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Waterfall_Insert ws, d
    n = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, n
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    CloseTemp wb
End Sub
```

Em `tests/vba/core/TestWaterfall.bas`, substitua `Test_Waterfall_Block_TotalFirstRow` (comportamento novo: `e` na primeira linha com número na última = build-down):

```vb
Public Sub Test_Waterfall_Block_TotalFirstRow_IsBuildDown()
    Dim wb As Workbook, body As Range
    Set wb = NewTempWorkbook()
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0013", Simple(wb, Array("e", 5))))
    AssertEqual True, body.Cells(2, WF_COL_BD).Value
    AssertEqual 5, body.Cells(2, WF_COL_END).Value, "primeiro total calculado para tras"
    AssertEqual 5, body.Cells(2, WF_COL_TOTAL).Value
    AssertEqual True, body.Cells(3, 3).Value: AssertEqual 5, body.Cells(3, WF_COL_END).Value
    ' 'e' na primeira linha SEM numero na ultima: nada desenhado (como na Fase 1)
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0016", Simple(wb, Array("e", "e"))))
    AssertEqual False, body.Cells(2, WF_COL_BD).Value: AssertEqual 0, body.Cells(2, WF_COL_END).Value
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-f2\build\run-tests.ps1 -Target core -Filter "TestWaterfall2|TotalFirstRow" -TimeoutSec 240` → saída 3 e depois `not ok`.

- [ ] **Passo 3: Implementar** — `Waterfall_BuildBlock`: alocar `20 + IIf(S > 1, 10*S, 0)` colunas; escrever as colunas 1..18 como hoje para S = 1 (apenas trocando `IsTotal` da última linha, `End` da 1.ª/última linha e acrescentando 19/20); para S > 1 escrever `IsTotal` (SUMPRODUCT), `Delta`, `Lo`, `Hi`, `LabelVal`, `LabelPos` nas versões multi e as 10 colunas por série. Em `Waterfall_Insert`, após validar `S <= 6`, montar as séries conforme a ordem acima com `Block_Ref`; `Waterfall_ApplyColors` lê `j` do sufixo `_j` dos nomes `tqq_dn_j/tqq_upn_j/tqq_dnn_j` e, para as não-auxiliares, pela ordem. `Labels_LinkPoints`: envolva o corpo do `For` por ponto em `On Error Resume Next ... If Err.Number <> 0 Then Err.Clear` (pontos `NA()` de `tqq_lbl` em multi).

- [ ] **Passo 4: GREEN e suíte completa** (os testes da Fase 1 em `TestWaterfall` e `TestSamples` continuam verdes; `Test_Samples_Content` exige `B111 = "e"` etc.: nada muda nos exemplos ainda). Exporte `cht.Export TempDir() & "\wfmulti.png"` da cascata de 2 séries e confira na imagem que a legenda mostra `Loja` e `Online` (e não `tqq_dn_*`, que têm a mesma cor: só a imagem distingue o mapeamento de `Chart_HideHelperLegendEntries`); registre no relatório. Se ler `HasDataLabel`/`DataLabel.Formula` de um ponto `NA()` lançar erro na limpeza, trate-o em `modCalcSheet` sem relaxar a falha fechada (um ponto `NA()` sem rótulo devolve `HasDataLabel = False`) e registre.

- [ ] **Passo 5: Commit** — `git commit -m "feat: cascata com varias series (segmentos por serie, legenda) e cascata de cima para baixo com 'e' na primeira linha"` (+ trailer).

**Verificação no Excel real:** tabela Etapa | Loja | Online com `e` nas linhas de total: segmentos coloridos por série, legenda com 2 entradas, rótulos por segmento, conectores no líquido; mudar -200 para -100 e ver tudo acompanhar; tabela com `e` na primeira linha e número na última: primeira barra calculada.

---

### Tarefa 5: Cascata horizontal (`barwaterfall`)

**Files:**
- Modify: `src/core/modules/modChartWaterfall.bas`, `src/core/modules/modChartKinds.bas` (`Kinds_Insert` passa `kind` à cascata), `src/core/modules/modCoreApi.bas` (`ApplyToggle "labels"/"connectors"` por família, não por `= "waterfall"`), `src/core/modules/modAnnotations.bas` (`Ann_Refresh` reconstrói os conectores por família)
- Test: `tests/vba/core/TestWaterfall2.bas`, `tests/vba/core/TestCoreApi.bas`

**Interfaces:**
- Consumes: Tarefas 1–4; `Ann_SetWaterfallConnectors` (layout horizontal da Tarefa 2), `Chart_SyncSecondaryX`.
- Produces: `Waterfall_Insert(ws, data, Optional kind As String = "waterfall")` — `barwaterfall` usa `xlBarStacked`, `Style_ApplyBase` (categorias invertidas), `Registry` tipo `barwaterfall` (`Waterfall_BuildBlock(wb, id, data, Optional kind As String = "waterfall")` grava `kind` na meta do bloco, de onde `Registry_ChartType` lê; hoje grava `"waterfall"` fixo); `Waterfall_SetLabels` em horizontal cria os auxiliares por `Chart_AddHelperSeries(cht, nome, valueCells, posCells = coluna 19)` (XY secundária) e liga com `horizontal:=True` (Right/Left); `Kinds_Insert` chama `Waterfall_Insert(ws, data, kind)` para a família `waterfall`. Em `modCoreApi`, `ApplyToggle`/`ApplyOption` decidem por `Kinds_Family(kind)` (`"waterfall"` → `Waterfall_SetLabels`, `Ann_SetWaterfallConnectors`, `Waterfall_ApplyColors`). `Ann_Refresh` reconstrói os conectores quando `Kinds_Family(chartKind) = "waterfall"` (hoje compara `chartKind = "waterfall"`: a cascata horizontal não refaria os conectores ao mudar o número de categorias).

- [ ] **Passo 1: Testes (RED)** — acrescentar a `tests/vba/core/TestWaterfall2.bas`:

```vb
Private Function SimpleFixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Inicio", 100)
    ws.Range("A3:B3").Value = Array("A", -150)
    ws.Range("A4:B4").Value = Array("B", 30)
    ws.Range("A5:B5").Value = Array("Subtotal", "e")
    ws.Range("A6:B6").Value = Array("C", 70)
    ws.Range("A7:B7").Value = Array("Fim", "e")
    Set SimpleFixture = ChartData_FromRange(ws.Range("A1:B7"))
End Function

Public Sub Test_BarWaterfall_Insert_SeriesConnectorsAndLabels()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lbl As Series, ax As Axis, il As Double, iw As Double, c As Shape
    Set wb = NewTempWorkbook(): Set d = SimpleFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d, "barwaterfall").Chart
    AssertEqual "barwaterfall", Registry_ChartType(cht)
    AssertEqual xlBarStacked, cht.SeriesCollection(1).ChartType
    AssertTrue cht.Axes(xlCategory).ReversePlotOrder
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    AssertEqual xlSecondary, lbl.AxisGroup
    AssertEqual "100", lbl.Points(1).DataLabel.Text: AssertEqual xlLabelPositionRight, lbl.Points(1).DataLabel.Position
    AssertEqual "-150", lbl.Points(2).DataLabel.Text: AssertEqual xlLabelPositionLeft, lbl.Points(2).DataLabel.Position
    AssertEqual "-20", lbl.Points(4).DataLabel.Text, "subtotal"
    ' conectores verticais em x = End_i (regra de tres independente): End = 100, -50, -20, -20, 50
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth: Set ax = cht.Axes(xlValue, xlPrimary)
    Set c = cht.Shapes("tqq_ann_conn_1")
    AssertAlmost 0, c.Width, 0.01, "conector vertical"
    AssertAlmost il + iw * (100 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale), c.Left, 1
    AssertAlmost il + iw * (-50 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale), cht.Shapes("tqq_ann_conn_2").Left, 1
    AssertTrue c.Top > cht.PlotArea.InsideTop And c.Top + c.Height < cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight / 3, "entre as barras 1 e 2"
    AssertEqual "1", Registry_GetOption(cht, "connectors")
    CloseTemp wb
End Sub

Public Sub Test_BarWaterfall_LabelHelper_AlignedWithBarEnds()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lbl As Series, ax As Axis, il As Double, it As Double, iw As Double, ih As Double
    Dim x As Double, dl As DataLabel
    Set wb = NewTempWorkbook(): Set d = SimpleFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d, "barwaterfall").Chart
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    il = cht.PlotArea.InsideLeft: it = cht.PlotArea.InsideTop: iw = cht.PlotArea.InsideWidth: ih = cht.PlotArea.InsideHeight
    Set ax = cht.Axes(xlValue, xlPrimary)
    x = il + iw * (100 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)      ' Hi da barra 1
    Set dl = lbl.Points(1).DataLabel
    AssertTrue dl.Left >= x - 2 And dl.Left <= x + 14, "rotulo 1 no fim da barra: " & dl.Left & " vs " & x
    AssertTrue Abs((dl.Top + dl.Height / 2) - (it + ih / 12)) <= ih / 12, "rotulo 1 na primeira faixa (6 categorias)"
    x = il + iw * (-50 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)      ' Lo da barra 2 (rotulo a esquerda)
    Set dl = lbl.Points(2).DataLabel
    AssertTrue dl.Left + dl.Width <= x + 2 And dl.Left + dl.Width >= x - 14, "rotulo 2 a esquerda do Lo: " & (dl.Left + dl.Width) & " vs " & x
    CloseTemp wb
End Sub

Public Sub Test_BarWaterfall_Multi_And_Core()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C6").Select
    Core_Insert "barwaterfall"
    Set cht = Core_ActiveChart()
    AssertEqual "barwaterfall", Registry_ChartType(cht)
    AssertEqual 13, cht.SeriesCollection.Count
    AssertEqual "-30", Chart_FindSeries(cht, "tqq_lbl_1").Points(2).DataLabel.Text
    AssertTrue Core_GetToggle("connectors")
    Core_SetToggle "connectors", False
    AssertEqual 0, cht.Shapes.Count
    Core_SetToggle "labels", False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl_1") Is Nothing
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), Chart_FindSeries(cht, "Loja").Format.Fill.ForeColor.RGB
    Core_AddAnnotation "diff", 1, 5
    AssertEqual "-240 (-160%)", cht.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text
    Watchers_Stop
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "BarWaterfall"` → saída 3 / `not ok`.

- [ ] **Passo 3: Implementar** — `Waterfall_Insert` ganha o parâmetro `kind`; `horiz = Kinds_IsHorizontal(kind)`; `cht.ChartType = IIf(horiz, xlBarStacked, xlColumnStacked)`; `Waterfall_SetLabels` (nos dois ramos, S = 1 e S > 1) usa a coluna 19 como `posCells` quando horizontal; `Core_SafeRefresh` já sincroniza o eixo X secundário em horizontais (Tarefa 3). Em `modCoreApi.ApplyToggle`: `Select Case Kinds_Family(kind)`.

- [ ] **Passo 4: GREEN e suíte completa.**

- [ ] **Passo 5: Commit** — `git commit -m "feat: cascata horizontal com rotulos em serie XY e conectores verticais"` (+ trailer).

**Verificação no Excel real:** Cascata (barras) com a tabela de 7 linhas; rótulos à direita/esquerda; conectores verticais; redimensionar e mudar valores.

---

### Tarefa 6: Área (empilhada e 100%), combinação (colunas + linha, eixo secundário) e opção de legenda

**Files:**
- Create: `src/core/modules/modChartArea.bas`, `src/core/modules/modChartCombo.bas`, `src/core/modules/modDecorations.bas` (só `Deco_SetLegend` nesta tarefa)
- Modify: `src/core/modules/modChartStacked.bas` (`Stacked_BuildBlock` com `kind = "combo"`; `Stacked_SetTotals/SetLabels/SetConnectors` decidem por `Kinds_Supports`; `Stacked_SetLabels` reaplica "fim da linha"), `src/core/modules/modChartStyle.bas` (`Style_HideValueAxis(ax)`, borda branca em áreas), `src/core/modules/modLabels.bas` (`Labels_ApplyCellFormat` por família), `src/core/modules/modChartKinds.bas` (`Kinds_Insert`: `area` → `Area_Insert`, `combo` → `Combo_Insert`), `src/core/modules/modCoreApi.bas` (`ValidateOption "legend"`, `ApplyOption "legend"`, `ApplyToggle "secondary"` e `"totals"/"labels"` da família `area`, erro 7002 tratado como aviso)
- Test: `tests/vba/core/TestArea.bas` (novo), `tests/vba/core/TestCoreApi.bas`

**Interfaces:**
- Consumes: Tarefas 1–3 (`Chart_AddHelperSeries`, `Block_Ref`, `Kinds_*`, `Labels_LinkPoints`, `Stacked_BuildBlock`, `Stacked_SetTotals/SetLabels/SetConnectors`, `Chart_HideHelperLegendEntries`).
- Produces (`modChartArea`): `Area_BuildBlock(wb, id, data, kind) As Range` — colunas `1 Categoria | 2 Total (=AGGREGATE(9,6,linha)) | 3 TotalPos (=Total; area100: =1) | 4..3+S Val_j | 4+S..3+2S Mid_j | 4+2S Pos | 5+2S VLine` com `Val_j = IF(ISNUMBER(src),src,0)` (area100: `IF(AND(ISNUMBER(src),Total<>0),src/Total,0)`), `Mid_j = SUM(Val_1..Val_j) - Val_j/2`, `VLine = AVERAGE(Total)` (`NA()` em area100); `Area_SeriesCount(blk) = (Columns.Count - 5) \ 2`; `Area_ValCol(blk, j) = 3 + j`; `Area_MidCol(blk, j) = 3 + S + j`; `Area_PosCol(blk) = Columns.Count - 1`; `Area_VLineCol(blk) = Columns.Count`; `Area_Insert(ws, data, kind) As ChartObject` (`xlAreaStacked`/`xlAreaStacked100`, séries da origem, `Style_ApplyBase`, cores da paleta com **borda branca 1 pt**, legenda à direita, opções `numfmt=int, palette=padrao, n, legend=right`, rótulos e totais ligados); `Area_SetLabels(cht, on)` (um auxiliar `tqq_lbl_j` por série em `Mid_j`, rótulo ligado a `Val_j`, posição Center, cor por contraste); `Area_SetTotals(cht, on)` (`tqq_total` em `TotalPos`, rótulo ligado a `Total`, negrito, acima/abaixo pelo sinal).
- Produces (`modChartCombo`): `Combo_Insert(ws, data) As ChartObject` — exige `data.SeriesCount >= 2`, senão `Err.Raise vbObjectError + 6008, "Combo_Insert", "A combinação precisa de pelo menos 2 colunas de valores: as primeiras viram colunas empilhadas e a última vira linha."`; bloco `Stacked_BuildBlock(wb, id, data, "combo")` (Total/TotalPos só das séries 1..S-1); séries 1..S-1 `xlColumnStacked`, série S `xlLineMarkers`; `Combo_DefaultSecondary(data) As Boolean` (= ambas as magnitudes > 0 e `lineMax > 5*colMax Or lineMax*5 < colMax`, com `lineMax` = maior |valor| da última coluna e `colMax` = maior |soma das demais| por linha, lidos por `Util_ToDouble`); `Combo_SetSecondary(cht, on)` (`AxisGroup` da linha; eixo de valores secundário oculto por `Style_HideValueAxis`; opção `secondary`); `Combo_LineSeries(cht) As Series` (última série não-auxiliar). Totais/rótulos/conectores/cores via `Stacked_*` (kind `combo` está na matriz).
- Produces (`modDecorations`): `Deco_SetLegend(cht, mode As String)` com `mode` em `right|none|lineend`: `right` liga a legenda à direita (9 pt) e esconde auxiliares; `none` desliga; `lineend` só para `line` e `combo` (senão `Err.Raise vbObjectError + 7002, "Deco_SetLegend", "Rótulos no fim da linha só existem em gráficos de linha e combinação."`): desliga a legenda e, em cada série de linha não-auxiliar, o último ponto ganha `HasDataLabel = True`, `ShowSeriesName = True`, `Separator = " "` (mantendo `ShowValue` conforme o toggle de rótulos); `right/none` desfazem o `ShowSeriesName`. Persiste `legend`. `Deco_ApplyLineEnd(cht)` (reaplica quando `legend=lineend`; chamada no fim de `Stacked_SetLabels`).
- Produces (`modChartStyle`): `Style_HideValueAxis(ax As Axis)` (grade, marcas, rótulos e linha ocultos; extraído de `Style_ApplyBase`); `Style_ApplySeriesColors` em séries de área: preenchimento da paleta + `Format.Line.Visible = msoTrue`, `ForeColor = RGB(255,255,255)`, `Weight = 1`.
- Produces (`modChartStacked`): `Stacked_Insert` passa a gravar `Registry_SetOption cht, "legend", "right"` (a legenda já nasce à direita, mas a opção nunca era gravada e `Core_GetOption("legend")` devolveria `""`); `Combo_Insert` também grava `legend=right` (e `Area_Insert`, acima).
- Produces (`modCoreApi`): `ValidateOption("legend", v)` aceita `right|none|lineend`; `ApplyOption "legend"` → `Deco_SetLegend` **sem** passar pelo `Registry_SetOption` genérico do início de `ApplyOption` (hoje `ApplyOption` grava a opção antes de despachar): `Deco_SetLegend` valida, aplica e só então grava, de modo que a recusa 7002 não deixa `legend=lineend` gravado; as demais chaves continuam gravando antes de despachar (o despacho de `numfmt` lê a opção do registro, ex.: `Labels_SegmentFormat`: inverter a ordem para todas quebraria o 100%); `ApplyToggle "secondary"` → `Combo_SetSecondary` (só `combo`); `ApplyToggle "totals"`/`"labels"` na família `area` → `Area_SetTotals`/`Area_SetLabels` (a família `combo` usa os `Stacked_*`); em `Core_SetToggle/Core_SetOption`, erro `vbObjectError + 7002` vira `Core_Notify Err.Description, vbInformation` sem persistir a opção (mesmo tratamento do 7001 em `Core_AddAnnotation`); `Core_Insert` trata `vbObjectError + 6008` como aviso (`Core_Notify Err.Description`).
- Produces (`modLabels`): `Labels_ApplyCellFormat` por família: `stacked` (col 2 + `Part_j`), `waterfall` (col 17, `LabelVal`, e `D_j` quando S > 1), `area` (col 2 + `Val_j`), `combo` (como `stacked`).

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestArea.bas`:

```vb
Attribute VB_Name = "TestArea"
Option Explicit

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
    ws.Range("A2:D2").Value = Array(2022, 10, 20, 5)
    ws.Range("A3:D3").Value = Array(2023, 12, 22, 6)
    ws.Range("A4:D4").Value = Array(2024, 15, 25, 8)
    ws.Range("A5:D5").Value = Array(2025, 18, 30, 9)
    Set Fixture = ChartData_FromRange(ws.Range("A1:D5"))
End Function

Private Function ComboFixture(wb As Workbook, Optional lineBig As Boolean = False) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "Vendas", "Custo", "Margem")
    ws.Range("A2:D2").Value = Array(2022, 100, 60, IIf(lineBig, 100, 0.42))
    ws.Range("A3:D3").Value = Array(2023, 120, 70, IIf(lineBig, 120, 0.45))
    ws.Range("A4:D4").Value = Array(2024, 150, 80, IIf(lineBig, 150, 0.47))
    Set ComboFixture = ChartData_FromRange(ws.Range("A1:D4"))
End Function

Public Sub Test_Area_Block_ValMidTotals()
    Dim wb As Workbook, blk As Range, body As Range
    Set wb = NewTempWorkbook()
    Set blk = Area_BuildBlock(wb, "ar0001", Fixture(wb), "area"): Set body = Block_Body(blk)
    AssertEqual 11, blk.Columns.Count: AssertEqual 3, Area_SeriesCount(blk)
    AssertEqual 35, body.Cells(2, 2).Value: AssertEqual 35, body.Cells(2, 3).Value
    AssertEqual 10, body.Cells(2, Area_ValCol(blk, 1)).Value: AssertEqual 5, body.Cells(2, Area_ValCol(blk, 3)).Value
    AssertEqual 5, body.Cells(2, Area_MidCol(blk, 1)).Value: AssertEqual 20, body.Cells(2, Area_MidCol(blk, 2)).Value
    AssertAlmost 32.5, body.Cells(2, Area_MidCol(blk, 3)).Value
    AssertEqual 3.5, body.Cells(2, Area_PosCol(blk)).Value
    AssertEqual 45, body.Cells(2, Area_VLineCol(blk)).Value, "media de 35, 40, 48, 57"
    Set blk = Area_BuildBlock(wb, "ar0002", Fixture(wb), "area100"): Set body = Block_Body(blk)
    AssertEqual 1, body.Cells(2, 3).Value
    AssertAlmost 10 / 35, body.Cells(2, Area_ValCol(blk, 1)).Value
    AssertAlmost 5 / 35, body.Cells(2, Area_MidCol(blk, 1)).Value
    AssertTrue IsError(body.Cells(2, Area_VLineCol(blk)).Value)
    CloseTemp wb
End Sub

Public Sub Test_Area_Insert_LabelsTotalsLegendBorder()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area").Chart
    AssertEqual "area", Registry_ChartType(cht)
    AssertEqual xlAreaStacked, cht.SeriesCollection(1).ChartType
    AssertEqual RGB(255, 255, 255), cht.SeriesCollection(1).Format.Line.ForeColor.RGB, "borda branca separa as faixas"
    AssertEqual "10", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text
    AssertEqual "5", Chart_FindSeries(cht, "tqq_lbl_3").Points(1).DataLabel.Text
    AssertEqual xlLabelPositionCenter, Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Position
    AssertEqual "35", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    AssertEqual xlLabelPositionAbove, Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Position
    AssertEqual 3, cht.Legend.LegendEntries.Count
    AssertEqual "right", Registry_GetOption(cht, "legend"): AssertEqual "4", Registry_GetOption(cht, "n")
    Area_SetLabels cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl_1") Is Nothing: AssertEqual "0", Registry_GetOption(cht, "labels")
    Area_SetTotals cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual 3, cht.SeriesCollection.Count
    CloseTemp wb
End Sub

Public Sub Test_Area100_LabelsShowShare()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area100").Chart
    AssertEqual xlAreaStacked100, cht.SeriesCollection(1).ChartType
    AssertEqual "29%", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text, "10/35"
    AssertEqual "35", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "total absoluto"
    CloseTemp wb
End Sub

Public Sub Test_Area_Annotations_And_IndependentGeometry()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ln As Shape, expectStart As Double, il As Double, iw As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area").Chart
    Set ln = Ann_AddCagr(cht, 1, 4)
    AssertEqual "CAGR +17,7%", cht.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text      ' (57/35)^(1/3)-1 = 17,65%
    Ann_AddDifference cht, 1, 4
    AssertEqual "+22 (+63%)", cht.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth
    If cht.Axes(xlCategory).AxisBetweenCategories Then expectStart = il + iw / 8 Else expectStart = il
    AssertAlmost expectStart, ln.Left, 1, "inicio da seta na categoria 1 conforme a posicao do eixo"
    CloseTemp wb
End Sub

Public Sub Test_Combo_Insert_LineSecondaryTotals()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ln As Series
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb)
    Set cht = Combo_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "combo", Registry_ChartType(cht)
    AssertEqual xlColumnStacked, cht.SeriesCollection(1).ChartType
    Set ln = Combo_LineSeries(cht)
    AssertEqual "Margem", ln.Name
    AssertEqual xlLineMarkers, ln.ChartType
    AssertEqual xlSecondary, ln.AxisGroup, "magnitudes diferem por mais de 5x: eixo secundario por padrao"
    AssertEqual "1", Registry_GetOption(cht, "secondary")
    AssertEqual xlTickLabelPositionNone, cht.Axes(xlValue, xlSecondary).TickLabelPosition
    AssertEqual "160", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "total so das colunas"
    AssertEqual 3, cht.Legend.LegendEntries.Count
    Combo_SetSecondary cht, False
    AssertEqual xlPrimary, Combo_LineSeries(cht).AxisGroup
    AssertEqual "0", Registry_GetOption(cht, "secondary")
    Ann_AddCagr cht, 1, 3
    AssertEqual "CAGR +19,9%", cht.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text      ' (230/160)^(1/2)-1
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb, True)
    Set cht = Combo_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "0", Registry_GetOption(cht, "secondary"), "mesma ordem de grandeza: eixo unico"
    AssertEqual xlPrimary, Combo_LineSeries(cht).AxisGroup
    CloseTemp wb
End Sub

Public Sub Test_Combo_RequiresTwoValueColumns()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, en As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Ano", "Valor"): ws.Range("A2:B2").Value = Array(2022, 1): ws.Range("A3:B3").Value = Array(2023, 2)
    Set d = ChartData_FromRange(ws.Range("A1:B3"))
    On Error Resume Next
    Combo_Insert ws, d
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 6008, en
    AssertEqual 0, ws.ChartObjects.Count
    ws.Range("A1:B3").Select
    Core_ClearLastNotice
    Core_Insert "combo"
    AssertTrue InStr(Core_LastNotice(), "pelo menos 2 colunas") > 0, Core_LastNotice()
    Watchers_Stop
    CloseTemp wb
End Sub

Public Sub Test_Legend_Modes_LineEnd_Only_LineAndCombo()
    Dim wb As Workbook, d As clsChartData, cht As Chart, p As DataLabel, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "line").Chart
    Deco_SetLegend cht, "lineend"
    AssertFalse cht.HasLegend
    Set p = cht.SeriesCollection(1).Points(4).DataLabel
    AssertTrue p.ShowSeriesName, "nome da serie no ultimo ponto"
    AssertTrue InStr(p.Text, "Sul") > 0, p.Text
    AssertEqual "lineend", Registry_GetOption(cht, "legend")
    Stacked_SetLabels cht, False
    AssertTrue cht.SeriesCollection(1).Points(4).DataLabel.ShowSeriesName, "fim da linha sobrevive ao toggle de rotulos"
    Deco_SetLegend cht, "none"
    AssertFalse cht.HasLegend
    AssertFalse cht.SeriesCollection(1).Points(4).HasDataLabel, "sem rotulos e sem fim da linha"
    Deco_SetLegend cht, "right"
    AssertTrue cht.HasLegend: AssertEqual xlLegendPositionRight, cht.Legend.Position
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    On Error Resume Next
    Deco_SetLegend cht, "lineend"
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    AssertEqual "Rótulos no fim da linha só existem em gráficos de linha e combinação.", ed
    AssertEqual "right", Registry_GetOption(cht, "legend", "right")
    CloseTemp wb
End Sub
```

Em `TestCoreApi.bas`:

```vb
Public Sub Test_Core_Legend_And_Secondary_Options()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    Core_ClearLastNotice
    Core_SetOption "legend", "lineend"
    AssertEqual "Rótulos no fim da linha só existem em gráficos de linha e combinação.", Core_LastNotice()
    AssertEqual "right", Core_GetOption("legend"), "nao persiste"
    Core_SetOption "legend", "none"
    AssertFalse Core_ActiveChart().HasLegend
    Core_SetToggle "secondary", True
    AssertFalse Core_GetToggle("secondary"), "secondary so em combinacao"
    Teardown wb
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "TestArea|Legend_And_Secondary"` → saída 3.

- [ ] **Passo 3: Implementar** — `Area_Insert` segue o esqueleto de `Stacked_Insert` (rollback idêntico: `co.Delete` + `CalcSheet_DeleteBlock`). `Combo_Insert`: montar como `Stacked_Insert` com `kind = "combo"` e, antes de `Style_ApplyBase`, `cht.SeriesCollection(S).ChartType = xlLineMarkers`; depois `Combo_SetSecondary cht, Combo_DefaultSecondary(data)`. `Stacked_SetTotals` já recusa por `Not Kinds_Supports(kind, "totals")` desde a Tarefa 3 (o que inclui `combo`). `Deco_SetLegend "lineend"`: percorra `cht.SeriesCollection`, pule `Style_IsHelperSeries`, considere apenas séries com `ChartType` `xlLine`/`xlLineMarkers`.

- [ ] **Passo 4: GREEN e suíte completa.**

- [ ] **Passo 5: Commit** — `git commit -m "feat: area empilhada e 100%, combinacao colunas + linha com eixo secundario e legenda (direita, nenhuma, fim da linha)"` (+ trailer).

**Verificação no Excel real:** inserir Área empilhada e Área 100% (rótulos no meio das faixas, totais acima, bordas brancas); Combinação com uma coluna de percentual (linha no eixo secundário, totais só das colunas); Legenda > Nome no fim da linha num gráfico de linha; Legenda > Nenhuma.

---

### Tarefa 7: Decorações: linha de valor, indicador 100% e mesma escala

**Files:**
- Modify: `src/core/modules/modDecorations.bas`, `src/core/modules/modAnnotations.bas` (kind `ind100`: `Ann_AddInd100`/`Ann_RemoveInd100`), `src/core/modules/modChartStacked.bas` e `modChartArea.bas` (`ind100` ligado por padrão nos 100%), `src/core/modules/modCoreApi.bas` (`Core_SameScale`, `Core_ResetScale`, `Core_SelectedCharts`, toggles `ind100`/`vline`, `Core_SafeRefresh` reaplica escala fixa)
- Test: `tests/vba/core/TestDecorations.bas` (novo); ajustar `TestCoreApi.Test_Core_AddAnnotation_Diff_On100Chart_NoThrow` (contar só formas `tqq_ann_diff*`) e `TestSamples.Test_Samples_Content` (o 100% passa a ter o indicador: `Ann_ShapeCount` deve ignorar `tqq_ann_ind100*`)

**Interfaces:**
- Consumes: `Chart_AddHelperSeries`, `Stacked_VLineCol/PosCol`, `Area_VLineCol/PosCol`, `Kinds_Supports`, `Geom_ChartValCoord`, `Ann_Refresh`, `Registry_*`.
- Produces (`modDecorations`):
  - `Deco_SetValueLine(cht, on)` — recusa com `vbObjectError + 7002, "A linha de valor não está disponível para este tipo de gráfico."` quando `Not Kinds_Supports(kind, "vline")`; liga: série `tqq_vline` por `Chart_AddHelperSeries` (vertical: linha sobre a coluna `VLine`; horizontal: XY com `X = VLine`, `Y = Pos`, do centro da 1.ª faixa ao da última), linha **visível** cinza `RGB(127,127,127)`, 1 pt, `msoLineDash`, sem marcadores, fora da legenda; rótulo ligado só no **último** ponto (`m = vcells.Rows.Count`: n nas famílias `stacked`/`combo`/`area`, 5n no Mekko; `s.Points(m).DataLabel.Formula = Block_Ref(vcells.Cells(m, 1))`, posição Right em vertical e Above em horizontal, 9 pt cinza); a coluna `VLine` do bloco recebe `NumberFormatLocal = Style_LocalizeCode("""Média """ & Style_NumberFormat(numfmt))`; opção `vline=1`. Desliga: apaga a série, `vline=0`. `Deco_VLineCells(cht) As Range` / `Deco_PosCells(cht) As Range` por família (`stacked`/`combo` → `Stacked_*Col`; `area` → `Area_*Col`; `mekko` acrescentado na Tarefa 9).
  - `Deco_SetInd100(cht, on)` — recusa `7002, "O indicador 100% só existe em gráficos 100% e Mekko (%)."`; liga: `Ann_AddInd100 cht` (as duas formas; ver `modAnnotations` abaixo), `Deco_EnsureIndicatorRoom(cht)` (vertical: `PlotArea.InsideLeft >= 46`, horizontal: `InsideTop >= 20`; só aumenta a folga), opção `ind100=1`, `Ann_Refresh`; desliga: `Ann_RemoveInd100 cht`, `ind100=0`.
  - `Deco_SameScale(charts As Collection, ByRef skipped As Long) As Long` — para cada gráfico com `Kinds_Supports(kind, "samescale")`: `Axes(xlValue).MinimumScaleIsAuto = True`, `MaximumScaleIsAuto = True`, lê min/max; aplica `min(mins)`/`max(maxs)` a todos os aceitos e persiste `vmin`/`vmax` com `Deco_ScaleText(v) = Trim$(Str$(v))` (ponto decimal; `Val` lê de volta); devolve quantos receberam a escala; `skipped` = quantos foram ignorados por tipo.
  - `Deco_ResetScale(cht)` — `vmin`/`vmax` = `""` e eixo automático.
  - `Deco_ApplyFixedScale(cht)` — se `vmin` não vazio, define `MinimumScale/MaximumScale` (só quando diferentes); nunca propaga erro.
- Produces (`modAnnotations`): `Public Function Ann_AddInd100(cht As Chart) As Shape` (as formas ficam em `modAnnotations` porque `Ann_NewName`/`Ann_AddLabel` são privadas do módulo; `modDecorations` não as enxerga): linha `tqq_ann_ind100` (1 pt, `RGB(89,89,89)`, `AlternativeText = "kind=ind100"`) e rótulo `tqq_ann_ind100_lbl` criado por `Ann_AddLabel cht, "tqq_ann_ind100"` (logo `AlternativeText = "kind=label;owner=tqq_ann_ind100"`), ajustado para texto "100%", 9 pt, sem negrito, sem preenchimento e sem borda; se as formas já existem, não cria de novo; `Public Sub Ann_RemoveInd100(cht As Chart)` apaga as duas. `Ann_Refresh` posiciona `kind=ind100` (válido quando `Kinds_Supports(kind, "ind100")`, senão oculto): vertical: `y = Geom_ChartValCoord(cht, 1)`, tique de `(InsideLeft - 8, y)` a `(InsideLeft - 1, y)`, rótulo `Left = InsideLeft - 10 - lbl.Width`, `Top = y - lbl.Height/2`; horizontal: `x = Geom_ChartValCoord(cht, 1)`, tique de `(x, InsideTop - 8)` a `(x, InsideTop - 1)`, rótulo `Left = x - lbl.Width/2`, `Top = InsideTop - 10 - lbl.Height`. `Ann_RemoveAll` **não** remove `ind100` (é um toggle).
- Produces (`modCoreApi`): `Core_SelectedCharts() As Collection` (gráfico ativo ThinQQ; senão todos os `Selection.ShapeRange` com `HasChart` e `Registry_IsThinQQChart`; nunca propaga erro); `Core_SameScale()` — com < 2 selecionados usa todos os gráficos ThinQQ da planilha ativa; com < 2 no total notifica `"Mesma escala: selecione dois ou mais gráficos do ThinQQ (ou nenhum, para usar todos os gráficos da planilha)."`; senão aplica e notifica `"Mesma escala aplicada a N gráfico(s)."` + (se `skipped > 0`) `" K ignorado(s): tipo sem escala livre (100%, Mekko (%), pizza, Gantt)."`; `Core_ResetScale()` — gráfico ativo ou seleção; notifica `"Escala automática restaurada em N gráfico(s)."`; `ApplyToggle "ind100"` → `Deco_SetInd100`, `"vline"` → `Deco_SetValueLine`; `Core_SafeRefresh` chama `Deco_ApplyFixedScale cht` antes de `Labels_RefreshPositions`. `Stacked_Insert`/`Area_Insert` chamam `Deco_SetInd100 cht, True` quando `Kinds_Is100(kind)`.

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestDecorations.bas`:

```vb
Attribute VB_Name = "TestDecorations"
Option Explicit

Private Function Fixture(wb As Workbook, Optional mult As Double = 1) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
    ws.Range("A2:D2").Value = Array(2022, 10 * mult, 20 * mult, 5 * mult)
    ws.Range("A3:D3").Value = Array(2023, 12 * mult, 22 * mult, 6 * mult)
    ws.Range("A4:D4").Value = Array(2024, 15 * mult, 25 * mult, 8 * mult)
    ws.Range("A5:D5").Value = Array(2025, 18 * mult, 30 * mult, 9 * mult)
    Set Fixture = ChartData_FromRange(ws.Range("A1:D5"))
End Function

Private Function CountPrefix(cht As Chart, p As String) As Long
    Dim i As Long
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, Len(p)) = p Then CountPrefix = CountPrefix + 1
    Next
End Function

Public Sub Test_VLine_Stacked_MeanLineWithLinkedLabel()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, v As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Deco_SetValueLine cht, True
    Set s = Chart_FindSeries(cht, "tqq_vline")
    AssertTrue Not s Is Nothing
    v = s.Values
    AssertAlmost 45, v(1): AssertAlmost 45, v(4)
    AssertEqual msoTrue, s.Format.Line.Visible
    AssertEqual msoLineDash, s.Format.Line.DashStyle
    AssertEqual "Média 45", s.Points(4).DataLabel.Text
    AssertFalse s.Points(1).HasDataLabel, "so o ultimo ponto tem rotulo"
    AssertEqual "1", Registry_GetOption(cht, "vline")
    AssertEqual 3, cht.Legend.LegendEntries.Count, "fora da legenda"
    d.Source.Worksheet.Range("B5").Value = 118          ' totais 35, 40, 48, 157 -> media 70
    v = s.Values
    AssertAlmost 70, v(1), 0.001, "a media e formula do bloco"
    AssertEqual "Média 70", s.Points(4).DataLabel.Text
    Deco_SetValueLine cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_vline") Is Nothing: AssertEqual "0", Registry_GetOption(cht, "vline")
    CloseTemp wb
End Sub

Public Sub Test_VLine_Bars_IsScatterAcrossCategories()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, x As Variant, y As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barclustered").Chart
    Deco_SetValueLine cht, True
    Set s = Chart_FindSeries(cht, "tqq_vline")
    AssertEqual xlSecondary, s.AxisGroup
    x = s.XValues: y = s.Values
    AssertAlmost 45, x(1), 0.001: AssertAlmost 45, x(4), 0.001
    AssertAlmost 3.5, y(1): AssertAlmost 0.5, y(4)
    AssertEqual "Média 45", s.Points(4).DataLabel.Text
    AssertEqual msoTrue, s.Format.Line.Visible
    CloseTemp wb
End Sub

Public Sub Test_VLine_Refused_On100AndWaterfall()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long, ed As String, ws As Worksheet
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Stacked_Insert(ws, d, "stacked100").Chart
    On Error Resume Next
    Deco_SetValueLine cht, True
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    AssertEqual "A linha de valor não está disponível para este tipo de gráfico.", ed
    AssertTrue Chart_FindSeries(cht, "tqq_vline") Is Nothing
    Set cht = Waterfall_Insert(ws, d).Chart
    On Error Resume Next
    Deco_SetValueLine cht, True
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en, "cascata: sem linha de valor"
    AssertEqual "A linha de valor não está disponível para este tipo de gráfico.", ed
    AssertTrue Chart_FindSeries(cht, "tqq_vline") Is Nothing
    ws.Range("A1:D5").Select
    Core_Insert "stacked100"
    Core_ClearLastNotice
    Core_SetToggle "vline", True
    ' via API o aviso fala da linha de valor (a Tarefa 11 troca o texto pela mensagem generica da matriz,
    ' "'Linha de valor (media)' nao se aplica..."; a chamada direta acima mantem a mensagem especifica)
    AssertTrue InStr(Core_LastNotice(), "inha de valor") > 0, Core_LastNotice()
    AssertFalse Core_GetToggle("vline")
    Watchers_Stop
    CloseTemp wb
End Sub

Public Sub Test_Ind100_DefaultOn_Geometry_And_Toggle()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tick As Shape, lbl As Shape, yTop As Double, ax As Axis
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked100").Chart
    AssertEqual "1", Registry_GetOption(cht, "ind100"), "ligado por padrao nos 100%"
    Set tick = cht.Shapes("tqq_ann_ind100"): Set lbl = cht.Shapes("tqq_ann_ind100_lbl")
    AssertEqual "100%", lbl.TextFrame2.TextRange.Text
    Set ax = cht.Axes(xlValue)
    yTop = cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight * (ax.MaximumScale - 1) / (ax.MaximumScale - ax.MinimumScale)
    AssertAlmost yTop, tick.Top, 1, "tique no nivel 100%"
    AssertTrue tick.Left + tick.Width <= cht.PlotArea.InsideLeft, "tique a esquerda da area de plotagem"
    AssertTrue lbl.Left + lbl.Width <= tick.Left + 0.5, "rotulo a esquerda do tique"
    AssertTrue lbl.Left >= 0, "dentro do grafico (folga criada)"
    Deco_SetInd100 cht, False
    AssertEqual 0, CountPrefix(cht, "tqq_ann_ind100"): AssertEqual "0", Registry_GetOption(cht, "ind100")
    Deco_SetInd100 cht, True
    AssertEqual 2, CountPrefix(cht, "tqq_ann_ind100")
    Ann_RemoveAll cht
    AssertEqual 2, CountPrefix(cht, "tqq_ann_ind100"), "Remover setas nao remove o indicador"
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked100").Chart
    Set tick = cht.Shapes("tqq_ann_ind100")
    AssertAlmost cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth, tick.Left, 1, "barras: tique no fim (100%) no topo"
    AssertTrue tick.Top + tick.Height <= cht.PlotArea.InsideTop + 0.5
    CloseTemp wb
End Sub

Public Sub Test_Ind100_Refused_OnStacked()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    AssertEqual "", Registry_GetOption(cht, "ind100", "")
    On Error Resume Next
    Deco_SetInd100 cht, True
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    AssertEqual "O indicador 100% só existe em gráficos 100% e Mekko (%).", ed
    AssertEqual 0, CountPrefix(cht, "tqq_ann_ind100")
    CloseTemp wb
End Sub

Public Sub Test_SameScale_Union_Persist_Refresh_Reset()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, c1 As Chart, c2 As Chart, coll As New Collection, skipped As Long, n As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set c1 = Stacked_Insert(ws, d, "stacked").Chart
    ws.Range("F1:I1").Value = Array("Ano", "A", "B", "C")
    ws.Range("F2:I2").Value = Array(2022, 100, 200, 50): ws.Range("F3:I3").Value = Array(2023, 120, 220, 60)
    Set c2 = Stacked_Insert(ws, ChartData_FromRange(ws.Range("F1:I3")), "clustered").Chart
    coll.Add c1: coll.Add c2
    n = Deco_SameScale(coll, skipped)
    AssertEqual 2, n: AssertEqual 0, skipped
    AssertEqual c2.Axes(xlValue).MaximumScale, c1.Axes(xlValue).MaximumScale
    AssertEqual c2.Axes(xlValue).MinimumScale, c1.Axes(xlValue).MinimumScale
    AssertTrue c1.Axes(xlValue).MaximumScale >= 220, "uniao cobre o maior grafico"
    AssertTrue Len(Registry_GetOption(c1, "vmax")) > 0 And Len(Registry_GetOption(c1, "vmin")) > 0
    AssertAlmost c1.Axes(xlValue).MaximumScale, Val(Registry_GetOption(c1, "vmax")), 0.000001, "texto neutro com ponto"
    c1.Axes(xlValue).MaximumScaleIsAuto = True                   ' alguem mexeu: o refresh reaplica
    Core_SafeRefresh c1
    AssertEqual c2.Axes(xlValue).MaximumScale, c1.Axes(xlValue).MaximumScale, "escala fixa reaplicada no refresh"
    Deco_ResetScale c1
    AssertTrue c1.Axes(xlValue).MaximumScaleIsAuto
    AssertEqual "", Registry_GetOption(c1, "vmax", "")
    CloseTemp wb
End Sub

Public Sub Test_Core_SameScale_SheetAndSelection_Notices()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, c3 As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D5").Select: Core_Insert "stacked"
    ws.Range("A1:D5").Select: Core_Insert "clustered"
    ws.Range("A1:D5").Select: Core_Insert "stacked100"
    Set c3 = Core_ActiveChart()
    ws.Range("A1").Select
    Core_ClearLastNotice
    Core_SameScale
    AssertEqual "Mesma escala aplicada a 2 gráfico(s). 1 ignorado(s): tipo sem escala livre (100%, Mekko (%), pizza, Gantt).", Core_LastNotice()
    AssertEqual "", Registry_GetOption(c3, "vmax", ""), "100% ignorado"
    Deco_ResetScale ws.ChartObjects(1).Chart: Deco_ResetScale ws.ChartObjects(2).Chart
    ws.Shapes.Range(Array(ws.ChartObjects(1).Name, ws.ChartObjects(2).Name)).Select
    Core_ClearLastNotice
    Core_SameScale
    AssertEqual "Mesma escala aplicada a 2 gráfico(s).", Core_LastNotice()
    AssertEqual ws.ChartObjects(1).Chart.Axes(xlValue).MaximumScale, ws.ChartObjects(2).Chart.Axes(xlValue).MaximumScale
    ws.ChartObjects(1).Select
    Core_ClearLastNotice
    Core_ResetScale
    AssertEqual "Escala automática restaurada em 1 gráfico(s).", Core_LastNotice()
    AssertTrue ws.ChartObjects(1).Chart.Axes(xlValue).MaximumScaleIsAuto
    Teardown wb
End Sub

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "TestDecorations"` → saída 3; depois de declarar as funções, `not ok`.

- [ ] **Passo 3: Implementar** — `Deco_SetValueLine` (vertical): `Set s = Chart_AddHelperSeries(cht, "tqq_vline", vcells)`; depois `s.Format.Line.Visible = msoTrue` etc.; o rótulo do último ponto, com `m = vcells.Rows.Count` (não o número de categorias: no Mekko da Tarefa 9 são 5n pontos): `s.Points(m).HasDataLabel = True: s.Points(m).DataLabel.Formula = Block_Ref(vcells.Cells(m, 1))`. Em `Core_SelectedCharts`, a seleção múltipla de gráficos chega como `TypeName(Selection) = "DrawingObjects"`; use `Selection.ShapeRange` sob `On Error Resume Next`. Em `Deco_SameScale`, leia min/max **depois** de ligar `IsAuto` (o Excel recalcula a escala automática na leitura).

- [ ] **Passo 4: GREEN e suíte completa** — incluindo os dois testes antigos ajustados (100% agora tem 2 formas `tqq_ann_ind100*`).

- [ ] **Passo 5: Commit** — `git commit -m "feat: linha de valor na media, indicador 100% (ligado por padrao nos 100%) e mesma escala entre graficos"` (+ trailer).

**Verificação no Excel real:** Colunas 100% nasce com "100%" à esquerda; Mais opções > Linha de valor num empilhado (linha tracejada com "Média N" à direita; mudar um valor move a linha); dois gráficos com escalas diferentes: selecionar os dois (Ctrl+clique) > Mesma escala; Escala automática; sem seleção: mensagem com a contagem.

---

### Tarefa 8: Pizza e rosca

**Files:**
- Create: `src/core/modules/modChartPie.bas`
- Modify: `src/core/modules/modChartKinds.bas` (`Kinds_Insert`: `pie` → `Pie_Insert`), `src/core/modules/modChartStyle.bas` (`Style_ApplyNumberFormat` família `pie`; `Style_ApplyPointColors`), `src/core/modules/modLabels.bas` (`Labels_ApplyCellFormat`/`Labels_RefreshPositions` ignoram a família `pie`), `src/core/modules/modDecorations.bas` (`Deco_SetLegend` na família `pie`: precedência com `pieout`), `src/core/modules/modCoreApi.bas` (`ApplyToggle "pieout"` e `"labels"` da família `pie`, `ApplyOption "palette"` por família)
- Test: `tests/vba/core/TestPie.bas` (novo)

**Interfaces:**
- Consumes: `CalcSheet_AllocateBlock`, `Block_Ref`, `Style_ApplyBase` (tolerante a pizza, Tarefa 1), `Style_Palette`, `Style_SetColor`, `Style_ContrastFont`, `Style_LocalizeCode`, `Style_NumberFormat`, `Kinds_*`, `Registry_*`.
- Produces: `Pie_BuildBlock(wb, id, data, kind) As Range` — `1 Categoria | 2 Valor (=IF(AND(ISNUMBER(src),src>0),src,0)) | 3 Share (=IF(SUM(Valor)>0, Valor/SUM(Valor), 0))`; `Pie_Insert(ws, data, kind) As ChartObject` — uma série (nome = cabeçalho da 1.ª coluna de valores, `Values = Block_Ref(coluna Valor)`, `XValues = categorias da origem`), `xlPie` ou `xlDoughnut` (`DoughnutHoleSize = 50`), `FirstSliceAngle = 0`, `Style_ApplyBase`, `Style_ApplyPointColors cht, "padrao"` (`VaryByCategories = True`, ponto k com `pal((k-1) Mod 6)`, borda branca 1 pt), legenda à direita, rótulos **dentro** (pizza: `xlLabelPositionInsideEnd`; rosca: **nunca define `Position`** — o Excel recusa `InsideEnd`/`OutsideEnd` na rosca com erro 1004 — e os rótulos ficam na posição padrão, sobre o anel), opções `numfmt=int, palette=padrao, legend=right, labels=1, n` e, só na pizza, `pieout=0`; `Pie_SetLabels(cht, on)`; `Pie_SetOutside(cht, on)` — só `pie`: na rosca levanta `Err.Raise vbObjectError + 7002, "Pie_SetOutside", "'Rótulos fora da pizza' não se aplica a este tipo de gráfico."` sem gravar nada (o mesmo texto que a Tarefa 11 gera pela matriz); `pieout=1`: rótulos `xlLabelPositionOutsideEnd` com `ShowCategoryName = True`, `Separator = vbLf`, `ShowLeaderLines = True`, fonte cinza-escuro, legenda oculta; `pieout=0`: `xlLabelPositionInsideEnd`, sem nome da categoria, cor da fonte por contraste com a fatia, legenda conforme a opção `legend`; `Pie_ApplyNumberFormat(cht, key As String)` — recebe a **chave** (`int`, `pct0`, ...), grava `numfmt=key` e reaplica o conteúdo dos rótulos: `pct0/pct1`: `ShowPercentage = True`, `ShowValue = False`, formato localizado `0%`/`0.0%`; demais: `ShowValue = True`, `ShowPercentage = False`, formato localizado de `Style_NumberFormat(key)`; `Style_ApplyPointColors(cht, paletteName)` em `modChartStyle`. `Style_ApplyNumberFormat(cht, fmt As String)` mantém a assinatura de hoje (recebe o **código** de formato de `Style_NumberFormat`) e, na família `pie`, chama `Pie_ApplyNumberFormat cht, Registry_GetOption(cht, "numfmt", "int")` e sai (`ApplyOption` grava `numfmt` antes de despachar, então a chave lida já é a nova); `ApplyOption "palette"` despacha por família (`pie` → `Style_ApplyPointColors`); `ApplyToggle "labels"` na família `pie` → `Pie_SetLabels` (não o `Stacked_SetLabels` de reserva, que poria `Position = Center` e apagaria nome e participação).
- Precedência da legenda na família `pie` (`Deco_SetLegend` e `Pie_SetOutside` mexem em `HasLegend`): `HasLegend = (legend <> "none") And (pieout <> "1")` — rótulos fora vencem (os nomes já estão nos rótulos). `Pie_SetOutside` não grava `legend`; `Deco_SetLegend` (Tarefa 6) ganha o ramo `pie`: grava `legend` e aplica a mesma regra (com `pieout=1`, `right` fica gravado, mas a legenda continua oculta até desligar os rótulos fora); `lineend` continua recusado (7002). Rótulos nativos: `Labels_ApplyCellFormat` e `Labels_RefreshPositions` saem cedo na família `pie`.
- Anotações: recusadas pela matriz (7001); `vline`/`ind100`/`samescale`: recusados (7002 / ignorado).

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestPie.bas`:

```vb
Attribute VB_Name = "TestPie"
Option Explicit

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Produto", "Vendas")
    ws.Range("A2:B2").Value = Array("Eletro", 50)
    ws.Range("A3:B3").Value = Array("Móveis", 30)
    ws.Range("A4:B4").Value = Array("Tintas", 20)
    Set Fixture = ChartData_FromRange(ws.Range("A1:B4"))
End Function

Private Function Dist(cht As Chart, dl As DataLabel) As Double
    Dim cx As Double, cy As Double
    cx = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth / 2
    cy = cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight / 2
    Dist = Sqr((dl.Left + dl.Width / 2 - cx) ^ 2 + (dl.Top + dl.Height / 2 - cy) ^ 2)
End Function

Private Function Radius(cht As Chart) As Double
    Radius = IIf(cht.PlotArea.InsideWidth < cht.PlotArea.InsideHeight, cht.PlotArea.InsideWidth, cht.PlotArea.InsideHeight) / 2
End Function

Public Sub Test_Pie_Insert_SeriesColorsLabelsLegend()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, pal As Variant, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "pie").Chart
    AssertEqual "pie", Registry_ChartType(cht)
    AssertEqual xlPie, cht.ChartType
    AssertEqual 1, cht.SeriesCollection.Count
    Set s = cht.SeriesCollection(1)
    AssertEqual "Vendas", s.Name
    AssertEqual 0, cht.ChartGroups(1).FirstSliceAngle
    pal = Style_Palette("padrao")
    AssertEqual CLng(pal(0)), s.Points(1).Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(2)), s.Points(3).Format.Fill.ForeColor.RGB
    AssertTrue s.HasDataLabels
    AssertEqual xlLabelPositionInsideEnd, s.DataLabels.Position
    AssertTrue s.DataLabels.ShowValue: AssertFalse s.DataLabels.ShowPercentage: AssertFalse s.DataLabels.ShowCategoryName
    AssertEqual "50", s.Points(1).DataLabel.Text
    AssertTrue cht.HasLegend: AssertEqual 3, cht.Legend.LegendEntries.Count
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual 50, body.Cells(2, 2).Value: AssertAlmost 0.5, body.Cells(2, 3).Value
    AssertEqual "0", Registry_GetOption(cht, "pieout"): AssertEqual "3", Registry_GetOption(cht, "n")
    AssertEqual "right", Registry_GetOption(cht, "legend")
    Pie_SetLabels cht, False
    AssertFalse s.HasDataLabels: AssertEqual "0", Registry_GetOption(cht, "labels")
    CloseTemp wb
End Sub

Public Sub Test_Pie_NumFmtPct_ShowsPercentage()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "pie").Chart
    Set s = cht.SeriesCollection(1)
    Pie_ApplyNumberFormat cht, "pct0"
    AssertTrue s.DataLabels.ShowPercentage: AssertFalse s.DataLabels.ShowValue
    AssertEqual "50%", s.Points(1).DataLabel.Text: AssertEqual "20%", s.Points(3).DataLabel.Text
    Pie_ApplyNumberFormat cht, "dec1"
    AssertEqual "50" & Application.International(xlDecimalSeparator) & "0", s.Points(1).DataLabel.Text
    CloseTemp wb
End Sub

Public Sub Test_Pie_Outside_CategoryNamesNoLegend_AndGeometry()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "pie").Chart
    Set s = cht.SeriesCollection(1)
    AssertTrue Dist(cht, s.Points(1).DataLabel) < Radius(cht), "rotulo interno dentro do circulo"
    Pie_SetOutside cht, True
    AssertEqual "1", Registry_GetOption(cht, "pieout")
    AssertEqual xlLabelPositionOutsideEnd, s.DataLabels.Position
    AssertTrue s.DataLabels.ShowCategoryName
    AssertTrue InStr(s.Points(1).DataLabel.Text, "Eletro") > 0 And InStr(s.Points(1).DataLabel.Text, "50") > 0, s.Points(1).DataLabel.Text
    AssertFalse cht.HasLegend
    AssertTrue Dist(cht, s.Points(1).DataLabel) > Radius(cht), "rotulo externo fora do circulo"
    Deco_SetLegend cht, "right"
    AssertFalse cht.HasLegend, "rotulos fora vencem a opcao de legenda"
    AssertEqual "right", Registry_GetOption(cht, "legend")
    Pie_SetOutside cht, False
    AssertTrue cht.HasLegend: AssertFalse s.DataLabels.ShowCategoryName
    AssertEqual xlLabelPositionInsideEnd, s.DataLabels.Position
    Deco_SetLegend cht, "none"
    Pie_SetOutside cht, True: Pie_SetOutside cht, False
    AssertFalse cht.HasLegend, "desligar rotulos fora respeita legend=none"
    CloseTemp wb
End Sub

Public Sub Test_Doughnut_HoleAndKind()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "doughnut").Chart
    AssertEqual "doughnut", Registry_ChartType(cht)
    AssertEqual xlDoughnut, cht.ChartType
    AssertEqual 50, cht.ChartGroups(1).DoughnutHoleSize
    AssertEqual "30", cht.SeriesCollection(1).Points(2).DataLabel.Text
    CloseTemp wb
End Sub

' Sonda (Excel 2019): a rosca recusa DataLabels.Position InsideEnd/OutsideEnd (erro 1004). Por isso a rosca nunca
' define Position e nao tem 'pieout'. Se esta sonda falhar, o Excel mudou: revise a regra (nao "conserte" o teste).
Public Sub Test_Doughnut_RejectsLabelPosition_Probe_And_NoPieOut()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "doughnut").Chart
    AssertTrue cht.SeriesCollection(1).HasDataLabels, "rosca com rotulos na posicao padrao"
    On Error Resume Next
    cht.SeriesCollection(1).DataLabels.Position = xlLabelPositionOutsideEnd
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertTrue en <> 0, "sonda: o Excel aceitou OutsideEnd na rosca"
    On Error Resume Next
    Pie_SetOutside cht, True
    en = Err.Number: ed = Err.Description: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    AssertEqual "'Rótulos fora da pizza' não se aplica a este tipo de gráfico.", ed
    AssertEqual "", Registry_GetOption(cht, "pieout", ""), "rosca nao grava pieout"
    AssertTrue cht.HasLegend
    CloseTemp wb
End Sub

' Review Focus 5: negativos, zero e texto viram 0 (sem tocar na tabela), 7 fatias ciclam a paleta, rotulos externos cabem.
Public Sub Test_Pie_NegativeZeroText_AndSevenSlices()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, s As Series, body As Range, i As Long, pal As Variant, dl As DataLabel
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Produto", "Vendas")
    ws.Range("A2:B2").Value = Array("Eletroportáteis e utilidades", 50)
    ws.Range("A3:B3").Value = Array("Devoluções", -10)
    ws.Range("A4:B4").Value = Array("Sem venda", 0)
    ws.Range("A5:B5").Value = Array("Indisponível", "n/d")
    ws.Range("A6:B6").Value = Array("Móveis", 30)
    ws.Range("A7:B7").Value = Array("Tintas", 20)
    ws.Range("A8:B8").Value = Array("Ferramentas", 10)
    Set d = ChartData_FromRange(ws.Range("A1:B8"))
    Set cht = Pie_Insert(ws, d, "pie").Chart
    Set s = cht.SeriesCollection(1): Set body = Block_Body(Registry_Block(cht))
    AssertEqual 0, body.Cells(3, 2).Value: AssertEqual 0, body.Cells(5, 2).Value
    AssertEqual -10, ws.Range("B3").Value, "tabela do usuario intacta"
    AssertAlmost 50 / 110, body.Cells(2, 3).Value
    AssertEqual 7, s.Points.Count
    pal = Style_Palette("padrao")
    AssertEqual CLng(pal(0)), s.Points(7).Format.Fill.ForeColor.RGB, "setima fatia volta ao inicio da paleta"
    AssertEqual "", s.Points(2).DataLabel.Text, "fatia zero sem rotulo"
    Pie_SetOutside cht, True
    For i = 1 To 7
        Set dl = s.Points(i).DataLabel
        AssertTrue dl.Left >= 0 And dl.Top >= 0 And dl.Left + dl.Width <= cht.ChartArea.Width + 0.5, "rotulo " & i & " dentro do grafico"
    Next
    CloseTemp wb
End Sub

Public Sub Test_Pie_Refusals_And_Rollback()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, en As Long, ed As String, coll As New Collection, skipped As Long, nm As Name
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Pie_Insert(ws, d, "pie").Chart
    On Error Resume Next
    Ann_AddCagr cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    AssertEqual vbObjectError + 7001, en: AssertEqual "A seta CAGR não está disponível para este tipo de gráfico.", ed
    Deco_SetValueLine cht, True
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    coll.Add cht
    AssertEqual 0, Deco_SameScale(coll, skipped): AssertEqual 1, skipped
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Pie_Insert ws, d, "doughnut"
    en = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, en
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    CloseTemp wb
End Sub

Public Sub Test_Core_Insert_Pie_TogglesAndOptions()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:B4").Select
    Core_Insert "pie"
    Set cht = Core_ActiveChart()
    AssertEqual "pie", Registry_ChartType(cht)
    AssertTrue Core_Supports("pieout"): AssertFalse Core_Supports("totals")
    Core_SetToggle "pieout", True
    AssertTrue Core_GetToggle("pieout"): AssertFalse cht.HasLegend
    Core_SetOption "numfmt", "pct1"
    AssertEqual "50" & Application.International(xlDecimalSeparator) & "0%", Mid$(cht.SeriesCollection(1).Points(1).DataLabel.Text, InStr(cht.SeriesCollection(1).Points(1).DataLabel.Text, "50"))
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB
    Core_SetToggle "labels", False
    AssertFalse cht.SeriesCollection(1).HasDataLabels
    Core_SetToggle "labels", True
    AssertTrue cht.SeriesCollection(1).DataLabels.ShowPercentage, "religar por Pie_SetLabels mantem pct1 (participacao)"
    AssertTrue cht.SeriesCollection(1).DataLabels.ShowCategoryName, "religar mantem os rotulos fora com o nome"
    AssertEqual xlLabelPositionOutsideEnd, cht.SeriesCollection(1).DataLabels.Position
    Watchers_Stop
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "TestPie"` → saída 3.

- [ ] **Passo 3: Implementar** — `Pie_Insert` no esqueleto de `Stacked_Insert` (rollback). Em `Pie_ApplyLabelContent(cht)` (privada, chamada por `Pie_SetLabels`, `Pie_SetOutside` e `Pie_ApplyNumberFormat`): leia `numfmt` e `pieout` do registro e aplique tudo de uma vez (posição, conteúdo, separador, linhas de chamada, fonte); cor da fonte por ponto só quando dentro; na rosca, nunca escreva `Position`. `Style_ApplyPointColors`: `cht.ChartGroups(1).VaryByCategories = True` e, para `k = 1 To s.Points.Count`, `Style_SetColor s.Points(k).Format.Fill.ForeColor, pal((k - 1) Mod 6)`; `s.Format.Line.Visible = msoTrue: ForeColor = RGB(255,255,255): Weight = 1`.

- [ ] **Passo 4: GREEN e suíte completa.**

- [ ] **Passo 5: Commit** — `git commit -m "feat: pizza e rosca com rotulos dentro/fora, valor ou participacao e cores por fatia"` (+ trailer).

**Verificação no Excel real:** Pizza e Rosca de uma tabela de 3 linhas; Mais opções > Rótulos fora da pizza (nomes + linhas de chamada, legenda some; na rosca o comando avisa que não se aplica); Formato 12% mostra participações; Paleta; Seta CAGR avisa que não se aplica.

---

### Tarefa 9: Mekko (eixo percentual e unidades)

**Files:**
- Create: `src/core/modules/modChartMekko.bas`
- Modify: `src/core/modules/modChartKinds.bas` (`Kinds_Insert`: `mekko` → `Mekko_Insert`), `src/core/modules/modChartStyle.bas` (`Style_ApplyNumberFormat` família `mekko`), `src/core/modules/modLabels.bas` (`Labels_ApplyCellFormat` família `mekko`; `Labels_RefreshPositions` ignora `mekko`), `src/core/modules/modDecorations.bas` (`Deco_VLineCells`: `mekko` → `Mekko_VLineCells`), `src/core/modules/modCoreApi.bas` (família `mekko`: `ApplyToggle "labels"/"totals"` → `Mekko_SetLabels`/`Mekko_SetTotals`, `ApplyOption "palette"` → `Mekko_ApplyColors`)
- Test: `tests/vba/core/TestMekko.bas` (novo)

**Interfaces:**
- Consumes: Tarefas 1, 6, 7 (`Chart_AddHelperSeries`, `Deco_SetInd100`, `Deco_SetValueLine`, `Style_ApplySeriesColors` com borda branca em áreas).
- Produces (constantes/colunas): `MK_ROWS_PER_CAT = 5`; `MK_COL_X = 1`, `MK_COL_CAT = 2`, `MK_COL_W = 3`, `MK_COL_TOTAL = 4`; `Mekko_SeriesCount(blk) = (blk.Columns.Count - 8) \ 3`; `Mekko_YCol(j) = 4 + j`; `Mekko_LblCol(blk, j) = 4 + S + j`; `Mekko_LblValCol(blk, j) = 4 + 2*S + j`; `Mekko_TotPosCol(blk) = 5 + 3*S`; `Mekko_TotLblCol(blk) = 6 + 3*S`; `Mekko_CatPosCol(blk) = 7 + 3*S`; `Mekko_VLineCol(blk) = 8 + 3*S`; `Mekko_CenterRow(i) = 5*(i-1) + 3` (linha do corpo, contando o cabeçalho como linha 1: a linha `C` da categoria `i` é `body.Cells(1 + Mekko_CenterRow(i), ...)`).
- Produces: `Mekko_BuildBlock(wb, id, data, kind) As Range` (corpo `5n` linhas + cabeçalho; fórmulas do design §4.5 — `percent`: séries = colunas 2..S+1 da tabela, `W = AGGREGATE(9,6,valores da linha)`; `units`: coluna 2 = largura (`W = IF(AND(ISNUMBER(w),w>0),w,0)`), séries = colunas 3.., exige `data.SeriesCount >= 2` senão `Err.Raise vbObjectError + 6009, "Mekko_Insert", "Mekko (unidades): a segunda coluna da tabela é a largura; são necessárias pelo menos 3 colunas."`); `WT = SUM(W de todas as linhas C)`; `X`: `L0/L = 1+ROUND(1000*IF(WT=0,0,SUM(W_C1:W_C(i-1))/WT),0)` (i = 1 → 1), `C = 1+ROUND(1000*IF(WT=0,0,(SUM(W_C1:W_C(i-1))+SUM(W_C1:W_Ci))/2/WT),0)`, `R/R0 = 1+ROUND(1000*IF(WT=0,0,SUM(W_C1:W_Ci)/WT),0)`; `Y_j` = 0 em `L0/R0`; em `L/C/R`, com `v = IF(ISNUMBER(src),src,0)`: participação `IF(Total=0, 0, v/Total)` (`percent`: linha de total zero não gera `#DIV/0!`) ou o valor `v` (`units`); `Lbl_j` (só `C`) `= SUM(Y_1..Y_j) - Y_j/2`, `LblVal_j` (só `C`) = mesmo que `Y_j`; `TotPos` (`C`) = 1 (`percent`) ou `Total` (`units`); `TotLbl` (`C`) = `IF(WT=0,0,W/WT)` (`percent`) ou `Total` (`units`); `CatPos` (`C`) = 0; `Categoria` (`C`) = célula da categoria; `VLine` (todas as `5n` linhas): `units` → `= IF(WT=0,0,SUMPRODUCT(W_C1:W_Cn, Total_C1:Total_Cn)/WT)`; `percent` → `=NA()` (sem linha de valor); fora das linhas indicadas, `NA()` (nas colunas `Lbl/LblVal/TotPos/TotLbl/CatPos`) ou vazio (`Categoria`, `W`, `Total`).
- Produces: `Mekko_Insert(ws, data, kind) As ChartObject` — séries `j` com `Values = Block_Ref(Y_j, 5n linhas)` e `XValues = Block_Ref(X)`, `xlAreaStacked`; `Style_ApplyBase` e **depois** `Axes(xlCategory)`: `CategoryType = xlTimeScale`, `BaseUnitIsAuto = False`, `BaseUnit = xlDays`, `AxisBetweenCategories = False` (X = 1 na borda esquerda da área de plotagem, sem meia fatia), `MinimumScale = 1`, `MaximumScale = 1001`, `MajorUnit = 1000`, sem rótulos/marcas, linha cinza; eixo de valores: `percent` → `MinimumScale = 0`, `MaximumScale = 1` (oculto); `units` → automático (oculto); cores por `Mekko_ApplyColors(cht, paletteName)` = `Style_ApplySeriesColors` (área: borda branca 1 pt, Tarefa 6) seguido de `Format.Line.Weight = 1.5` nas séries de área (separadores mais visíveis); `ApplyOption "palette"` na família `mekko` chama `Mekko_ApplyColors` (senão a troca de paleta voltaria os separadores para 1 pt); auxiliares: `tqq_lbl_j` (`Lbl_j`, rótulo ligado a `LblVal_j`, Center, contraste), `tqq_total` (`TotPos`, rótulo `TotLbl`, Above, negrito), `tqq_cat` (`CatPos`, rótulo = célula `Categoria`, Below, 9 pt cinza); `Mekko_LinkLabels(s, valueCells, labelCells, fmt, pos, fontColor)` liga só os pontos cujo `labelCells(r)` não é erro, sob `On Error Resume Next` por ponto; legenda à direita sem auxiliares; opções `numfmt` (`pct0` em `percent`, `int` em `units`), `palette=padrao`, `n`, `legend=right`, `totals=1`, `labels=1`; `Deco_SetInd100 cht, True` em `percent`. Por fim garante `PlotArea.InsideTop >= 24` (só aumenta: `delta = 24 - InsideTop`, `PlotArea.Height = PlotArea.Height - delta`, `PlotArea.Top = PlotArea.Top + delta`, como `Ann_EnsureRoom`), para os rótulos de `tqq_total`, acima do nível 100%, caberem no gráfico. `Mekko_SetLabels(cht, on)`, `Mekko_SetTotals(cht, on)`, `Mekko_ApplyColors(cht, paletteName)`, `Mekko_VLineCells(cht) As Range` (`5n` linhas), `Mekko_PosCells` não existe (vertical). `Style_ApplyNumberFormat` família `mekko`: segmentos com `Labels_SegmentFormat` (percent → `pct0/pct1`), totais: `percent` → `pct0`, `units` → `numfmt`; `Labels_ApplyCellFormat` família `mekko`: `LblVal_j` e `TotLbl`.
- Anotações: recusadas (7001). `vline` só em `mekkounits`; `ind100` só em `mekko`; `gap`/`connectors`: não (matriz).

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestMekko.bas`:

```vb
Attribute VB_Name = "TestMekko"
Option Explicit

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Regiao", "Eletro", "Moveis")
    ws.Range("A2:C2").Value = Array("Sul", 40, 10)
    ws.Range("A3:C3").Value = Array("Sudeste", 20, 10)
    ws.Range("A4:C4").Value = Array("Norte", 10, 10)
    Set Fixture = ChartData_FromRange(ws.Range("A1:C4"))
End Function

Private Function UnitsFixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Regiao", "Largura", "Eletro", "Moveis")
    ws.Range("A2:D2").Value = Array("Sul", 50, 40, 10)
    ws.Range("A3:D3").Value = Array("Sudeste", 30, 20, 10)
    ws.Range("A4:D4").Value = Array("Norte", 20, 10, 10)
    Set UnitsFixture = ChartData_FromRange(ws.Range("A1:D4"))
End Function

Private Function C(i As Long) As Long     ' linha do corpo (1-based, cabecalho = 1) do centro da categoria i
    C = 1 + Mekko_CenterRow(i)
End Function

Public Sub Test_Mekko_Block_XSharesAndHelpers()
    Dim wb As Workbook, blk As Range, body As Range
    Set wb = NewTempWorkbook()
    Set blk = Mekko_BuildBlock(wb, "mk0001", Fixture(wb), "mekko"): Set body = Block_Body(blk)
    AssertEqual 14, blk.Columns.Count: AssertEqual 16, body.Rows.Count, "corpo: cabecalho + 5 linhas por categoria (o bloco tem ainda a linha de meta)"
    AssertEqual 2, Mekko_SeriesCount(blk)
    ' X: larguras 50%, 30%, 20% -> bordas 1, 501, 801, 1001; centros 251, 651, 901
    AssertEqual 1, body.Cells(2, MK_COL_X).Value: AssertEqual 1, body.Cells(3, MK_COL_X).Value
    AssertEqual 251, body.Cells(C(1), MK_COL_X).Value
    AssertEqual 501, body.Cells(5, MK_COL_X).Value: AssertEqual 501, body.Cells(6, MK_COL_X).Value
    AssertEqual 501, body.Cells(7, MK_COL_X).Value: AssertEqual 651, body.Cells(C(2), MK_COL_X).Value
    AssertEqual 901, body.Cells(C(3), MK_COL_X).Value: AssertEqual 1001, body.Cells(16, MK_COL_X).Value
    AssertEqual "Sul", body.Cells(C(1), MK_COL_CAT).Value: AssertEqual "", CStr(body.Cells(2, MK_COL_CAT).Value)
    AssertEqual 50, body.Cells(C(1), MK_COL_W).Value: AssertEqual 50, body.Cells(C(1), MK_COL_TOTAL).Value
    ' Y: participacoes nas linhas L/C/R, zero em L0/R0
    AssertAlmost 0.8, body.Cells(C(1), Mekko_YCol(1)).Value: AssertAlmost 0.2, body.Cells(C(1), Mekko_YCol(2)).Value
    AssertEqual 0, body.Cells(2, Mekko_YCol(1)).Value: AssertAlmost 0.8, body.Cells(3, Mekko_YCol(1)).Value
    AssertEqual 0, body.Cells(6, Mekko_YCol(1)).Value, "R0 da categoria 1"
    AssertAlmost 2 / 3, body.Cells(C(2), Mekko_YCol(1)).Value, 0.000001
    ' auxiliares so nas linhas C
    AssertAlmost 0.4, body.Cells(C(1), Mekko_LblCol(blk, 1)).Value: AssertAlmost 0.9, body.Cells(C(1), Mekko_LblCol(blk, 2)).Value
    AssertTrue IsError(body.Cells(2, Mekko_LblCol(blk, 1)).Value)
    AssertAlmost 0.8, body.Cells(C(1), Mekko_LblValCol(blk, 1)).Value
    AssertEqual 1, body.Cells(C(1), Mekko_TotPosCol(blk)).Value: AssertAlmost 0.5, body.Cells(C(1), Mekko_TotLblCol(blk)).Value, 0.000001, "largura 50%"
    AssertEqual 0, body.Cells(C(1), Mekko_CatPosCol(blk)).Value: AssertTrue IsError(body.Cells(5, Mekko_CatPosCol(blk)).Value), "linha R: NA"
    AssertTrue IsError(body.Cells(2, Mekko_VLineCol(blk)).Value), "percent: sem linha de valor"
    CloseTemp wb
End Sub

Public Sub Test_Mekko_Insert_AreaOnDateAxis_LabelsTotalsCategories()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Mekko_Insert(d.Source.Worksheet, d, "mekko").Chart
    AssertEqual "mekko", Registry_ChartType(cht)
    AssertEqual xlAreaStacked, cht.SeriesCollection(1).ChartType
    AssertEqual xlTimeScale, cht.Axes(xlCategory).CategoryType
    AssertEqual 1, cht.Axes(xlCategory).MinimumScale: AssertEqual 1001, cht.Axes(xlCategory).MaximumScale
    AssertEqual 0, cht.Axes(xlValue).MinimumScale: AssertEqual 1, cht.Axes(xlValue).MaximumScale
    AssertEqual 15, cht.SeriesCollection(1).Points.Count
    AssertEqual "80%", Chart_FindSeries(cht, "tqq_lbl_1").Points(C(1) - 1).DataLabel.Text
    AssertEqual "20%", Chart_FindSeries(cht, "tqq_lbl_2").Points(C(1) - 1).DataLabel.Text
    AssertEqual "50%", Chart_FindSeries(cht, "tqq_total").Points(C(1) - 1).DataLabel.Text, "largura da coluna"
    AssertEqual "Sul", Chart_FindSeries(cht, "tqq_cat").Points(C(1) - 1).DataLabel.Text
    AssertEqual xlLabelPositionBelow, Chart_FindSeries(cht, "tqq_cat").Points(C(1) - 1).DataLabel.Position
    AssertEqual 2, cht.Legend.LegendEntries.Count
    AssertEqual "1", Registry_GetOption(cht, "ind100"): AssertEqual "3", Registry_GetOption(cht, "n")
    AssertEqual "pct0", Registry_GetOption(cht, "numfmt")
    AssertEqual RGB(255, 255, 255), cht.SeriesCollection(1).Format.Line.ForeColor.RGB, "separadores brancos"
    AssertEqual 1.5, cht.SeriesCollection(1).Format.Line.Weight, "separadores de 1,5 pt"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    AssertEqual 1, CalcSheet_Cleanup(wb), "limpeza com o grafico novo (auxiliares com pontos NA) remove so o orfao"
    AssertFalse CalcSheet_CleanupReadFailed(), "pontos NA dos auxiliares nao sao falha de leitura"
    Mekko_SetLabels cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl_1") Is Nothing
    Mekko_SetTotals cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertTrue Not Chart_FindSeries(cht, "tqq_cat") Is Nothing, "categorias sempre"
    CloseTemp wb
End Sub

' Geometria independente: o centro de cada coluna e a soma das larguras anteriores + metade da propria.
Public Sub Test_Mekko_Geometry_Independent()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lbl As Series, il As Double, iw As Double, dl As DataLabel, centers As Variant, i As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Mekko_Insert(d.Source.Worksheet, d, "mekko").Chart
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth
    centers = Array(0.25, 0.65, 0.9)
    Set lbl = Chart_FindSeries(cht, "tqq_lbl_1")
    For i = 1 To 3
        Set dl = lbl.Points(C(i) - 1).DataLabel
        AssertAlmost il + iw * centers(i - 1), dl.Left + dl.Width / 2, 8, "rotulo da categoria " & i & " no centro da coluna"
    Next
    Set dl = Chart_FindSeries(cht, "tqq_total").Points(C(1) - 1).DataLabel
    AssertTrue dl.Top + dl.Height <= cht.PlotArea.InsideTop + 2, "total acima do nivel 100%"
    AssertTrue dl.Top >= 0, "rotulo de total dentro do grafico"
    AssertTrue cht.PlotArea.InsideTop >= 23.5, "folga acima da area de plotagem para os totais"
    AssertTrue Not cht.Shapes("tqq_ann_ind100") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_MekkoUnits_Block_VLine_AndNumFmt()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range, body As Range, s As Series, v As Variant
    Set wb = NewTempWorkbook(): Set d = UnitsFixture(wb)
    Set cht = Mekko_Insert(d.Source.Worksheet, d, "mekkounits").Chart
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual "mekkounits", Registry_ChartType(cht)
    AssertEqual 2, Mekko_SeriesCount(blk)
    AssertEqual 50, body.Cells(C(1), MK_COL_W).Value: AssertEqual 251, body.Cells(C(1), MK_COL_X).Value
    AssertEqual 40, body.Cells(C(1), Mekko_YCol(1)).Value, "unidades: valor absoluto"
    AssertEqual 50, body.Cells(C(1), Mekko_TotPosCol(blk)).Value: AssertEqual 50, body.Cells(C(1), Mekko_TotLblCol(blk)).Value
    AssertEqual 38, body.Cells(2, Mekko_VLineCol(blk)).Value, "media ponderada pela largura (2500+900+400)/100"
    AssertEqual "40", Chart_FindSeries(cht, "tqq_lbl_1").Points(C(1) - 1).DataLabel.Text
    AssertEqual "50", Chart_FindSeries(cht, "tqq_total").Points(C(1) - 1).DataLabel.Text
    AssertEqual "int", Registry_GetOption(cht, "numfmt"): AssertEqual "", Registry_GetOption(cht, "ind100", "")
    AssertTrue cht.Axes(xlValue).MaximumScaleIsAuto
    Deco_SetValueLine cht, True
    Set s = Chart_FindSeries(cht, "tqq_vline")
    v = s.Values
    AssertAlmost 38, v(1): AssertAlmost 38, v(15)
    AssertEqual "Média 38", s.Points(15).DataLabel.Text
    CloseTemp wb
End Sub

' Review Focus 3: linha com total zero vira coluna de largura nula, sem #DIV/0! e sem coluna fantasma.
Public Sub Test_Mekko_ZeroTotalRow_NoErrorNoWidth()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range, r As Long, cc As Long, blk As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("B3:C3").Value = Array(0, 0)              ' Sudeste some: larguras 50/70 e 20/70
    Set cht = Mekko_Insert(d.Source.Worksheet, d, "mekko").Chart
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 715, body.Cells(5, MK_COL_X).Value, "borda direita de Sul = 1+ROUND(1000*50/70)"
    AssertEqual 715, body.Cells(C(2), MK_COL_X).Value, "categoria vazia tem largura zero"
    AssertEqual 715, body.Cells(10, MK_COL_X).Value: AssertEqual 1001, body.Cells(16, MK_COL_X).Value
    For r = 2 To 16
        For cc = 1 To 4 + Mekko_SeriesCount(blk)
            AssertFalse IsError(body.Cells(r, cc).Value), "erro em " & body.Cells(r, cc).Address
        Next
    Next
    AssertEqual 0, body.Cells(C(2), Mekko_YCol(1)).Value
    AssertEqual 15, cht.SeriesCollection(1).Points.Count
    CloseTemp wb
End Sub

Public Sub Test_Mekko_Refusals_Rollback_And_UnitsNeeds3Cols()
    Dim wb As Workbook, d As clsChartData, d2 As clsChartData, ws As Worksheet, cht As Chart, en As Long, ed As String, nm As Name
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Mekko_Insert(ws, d, "mekko").Chart
    On Error Resume Next
    Ann_AddDifference cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    AssertEqual vbObjectError + 7001, en: AssertEqual "A seta de diferença não está disponível para este tipo de gráfico.", ed
    Deco_SetValueLine cht, True
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en, "linha de valor so em unidades"
    Set d2 = ChartData_FromRange(ws.Range("A1:B4"))          ' 2 colunas: largura sem nenhuma serie
    On Error Resume Next
    Mekko_Insert ws, d2, "mekkounits"
    en = Err.Number: ed = Err.Description: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 6009, en
    AssertEqual "Mekko (unidades): a segunda coluna da tabela é a largura; são necessárias pelo menos 3 colunas.", ed
    AssertEqual 1, ws.ChartObjects.Count, "nada inserido na recusa"
    ws.ChartObjects(1).Delete: CalcSheet_Cleanup wb
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Mekko_Insert ws, d, "mekko"
    en = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, en
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    CloseTemp wb
End Sub

Public Sub Test_Core_Insert_Mekko_TogglesAndOptions()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C4").Select
    Core_Insert "mekko"
    Set cht = Core_ActiveChart()
    AssertEqual "mekko", Registry_ChartType(cht)
    AssertTrue Core_Supports("ind100"): AssertFalse Core_Supports("gap"): AssertFalse Core_Supports("cagr")
    Core_SetOption "numfmt", "pct1"
    AssertEqual "80" & Application.International(xlDecimalSeparator) & "0%", Chart_FindSeries(cht, "tqq_lbl_1").Points(C(1) - 1).DataLabel.Text
    Core_SetToggle "ind100", False
    AssertFalse Core_GetToggle("ind100")
    Core_SetToggle "totals", False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(1).Format.Fill.ForeColor.RGB
    AssertEqual 1.5, cht.SeriesCollection(1).Format.Line.Weight, "trocar a paleta mantem os separadores do Mekko"
    Core_ClearLastNotice                                          ' o contador conta desde aqui, nao desde testes anteriores
    Core_Refresh
    AssertEqual 0, Core_NoticeCount(), "sem aviso de tabela mudada (n = 3 categorias, 15 pontos)"
    Watchers_Stop
    CloseTemp wb
End Sub
```

Observação: a fixture de 3 colunas em modo `mekkounits` tem 1 série (coluna 2 = largura, coluna 3 = série) e **é válida**; o 6009 só ocorre com 2 colunas (largura sem série), como no teste.

- [ ] **Passo 2: RED** — `-Filter "TestMekko"` → saída 3.

- [ ] **Passo 3: Implementar** — ordem em `Mekko_Insert`: bloco → `ChartObjects.Add` → séries (`XValues` depois de `ChartType = xlAreaStacked`) → `Style_ApplyBase` → eixo de categorias em `xlTimeScale` com escala fixa (se o Excel recusar `MinimumScale = 1`, use `2..1002` e desloque o `+1` das fórmulas de `X` para `+2`; registre no relatório) → cores (`Mekko_ApplyColors`) → auxiliares → opções → `Deco_SetInd100` (percent) → folga `InsideTop >= 24`. Em `Mekko_LinkLabels`, defina `s.HasDataLabels = True` uma vez e depois, por ponto, `If Not IsError(labelCells.Cells(r, 1).Value) Then` ligue (`DataLabel.Formula = Block_Ref(célula)`); pontos `NA()` não são plotados e não mostram rótulo.

- [ ] **Passo 4: GREEN e suíte completa.**

- [ ] **Passo 5: Commit** — `git commit -m "feat: Mekko (eixo percentual e unidades) como area empilhada em eixo de datas, com rotulos, totais e categorias ligados"` (+ trailer).

**Verificação no Excel real:** Mekko (%) de uma tabela 3x3: colunas com larguras proporcionais aos totais, separadores brancos, participações nos segmentos, larguras no topo, nomes embaixo, "100%" à esquerda; Mekko (unidades) com coluna Largura; Linha de valor na variante unidades; mudar um valor redistribui as larguras ao vivo.

---

### Tarefa 10: Gantt (atividades, barras, marcos, escala de datas e colchetes)

**Files:**
- Create: `src/core/modules/modChartGantt.bas`
- Modify: `src/core/modules/modChartKinds.bas` (`Kinds_Insert`: `gantt` → `Gantt_Insert`), `src/core/modules/modAnnotations.bas` (kind `bracket`; `Ann_SetShapeVisible` também cobre `_t1/_t2`), `src/core/modules/modCoreApi.bas` (`Core_AddAnnotation "bracket"`, com `de = até` válido; `Core_SafeRefresh` chama `Gantt_RefreshScale`; família `gantt`: `ApplyToggle "labels"` → `Gantt_SetLabels`, `ApplyOption "palette"` → `Gantt_ApplyColors`), `src/core/modules/modLabels.bas` e `modChartStyle.bas` (família `gantt` sem formato numérico de rótulos)
- Test: `tests/vba/core/TestGantt.bas` (novo), `tests/vba/core/TestCoreApi.bas` (remover `Test_Core_Insert_KindNotReady_NotifiesWithoutChart`, da Tarefa 1: com o Gantt pronto não resta tipo válido sem implementação)

**Interfaces:**
- Consumes: `Chart_AddHelperSeries`, `Chart_SetupSecondaryAxes`, `Chart_SyncSecondaryX`, `Geom_ChartValueToX`, `Geom_ChartCatCoord`, `Ann_NewName/Ann_AddLabel/Ann_MoveLine/Ann_ClampShape` (privadas de `modAnnotations`, onde `Ann_AddBracket` também fica), `Kinds_*`, `Style_Palette`, `Util_ToDouble`, `Block_Ref`.
- Produces (constantes): `GT_COL_ACT = 1, GT_COL_START = 2, GT_COL_END = 3, GT_COL_ISMS = 4, GT_COL_BARSTART = 5, GT_COL_DUR = 6, GT_COL_MSX = 7, GT_COL_POS = 8, GT_COL_LBLX = 9, GT_COL_LBLTXT = 10, GT_COL_UNIT = 12, GT_COL_AXMIN = 13, GT_COL_AXMAX = 14, GT_COL_TICK = 15, GT_COL_TICKY = 16, GT_COL_TICKLBL = 17, GT_COL_DMIN = 18, GT_COL_DMAX = 19, GT_SCALE_SLOTS = 40`; bloco com 19 colunas (a 11 fica vazia, separando atividades da escala) e `MAX(n, 40)` linhas de corpo além do cabeçalho.
- Produces: `Gantt_BuildBlock(wb, id, data) As Range` — valida `data.SeriesCount >= 2` (senão `Err.Raise vbObjectError + 6010, "Gantt_Insert", "Gantt: a tabela precisa de 3 colunas: atividade, início e fim (datas). Deixe o fim vazio para um marco."`) e pelo menos uma célula numérica na coluna de início (senão `6010, "Gantt: a coluna de início não tem datas válidas (use células de data, não texto)."`), tudo **antes** de alocar; fórmulas do design §4.6, com `LblX = IF(IsMs, Inicio, Fim + 1)` (a barra cobre o dia do fim inteiro e termina em `Fim + 1`: com `LblX = Fim` o rótulo invadiria a barra em 1 dia); `LblTxt` = texto da 4.ª coluna (`=src4&""`) quando `data.SeriesCount >= 3`, senão a data exibida `=IF(IsMs, Inicio, Fim)` (o último dia da atividade, não `LblX`) com `NumberFormat = "dd/mm"`; `DMin`/`DMax` (colunas 18/19, linha 2 do corpo) = `AGGREGATE(5,6,Inicio)` e `AGGREGATE(4,6,Inicio,Fim)`, e as fórmulas da escala referenciam essas duas células; `TickLbl` com `NumberFormat` definido por `Gantt_RefreshScale`.
- Produces: `Gantt_Insert(ws, data) As ChartObject` — `xlBarStacked` com `tqq_start` (`BarStart`, preenchimento invisível) e `tqq_dur` (`Dur`, `pal(1)`), `XValues` = células de atividade da origem, gap 50, `Style_ApplyBase` (categorias invertidas; eixo de valores oculto), auxiliares XY secundárias: `tqq_ms` (`X = MsX`, `Y = Pos`, marcador `xlMarkerStyleDiamond` tamanho 9 `pal(0)`), `tqq_lbl` (`X = LblX`, `Y = Pos`, rótulos ligados a `LblTxt`, Right, 9 pt cinza), `tqq_scale` (`X = Tick` linhas 2..41, `Y = TickY`, rótulos ligados a `TickLbl`, Above, 8 pt cinza; `ErrorBar Direction:=xlY, Include:=xlErrorBarIncludeMinusValues, Type:=xlErrorBarTypeFixedValue, Amount:=n`, barras de erro cinza-claro tracejadas sem extremidade (`EndStyle = xlNoCap`)); `HasLegend = False`; opções `palette=padrao, gap=50, labels=1, n`; `Gantt_RefreshScale` ao final. Os rótulos de `tqq_lbl` e de `tqq_scale` são ligados ponto a ponto com `p.HasDataLabel = True: p.DataLabel.Formula = Block_Ref(célula)` (mais fonte e posição), **nunca** por `Labels_LinkPoints`, que sobrescreve o formato `dd/mm`/`mmm/yy` das células com o `NumberFormatLocal` de um formato numérico. Depois dos auxiliares, garante `PlotArea.InsideTop >= 24` (só aumenta, como `Ann_EnsureRoom`) para a escala acima da área de plotagem caber no gráfico (os colchetes pedem mais, abaixo).
- Produces: `Gantt_RefreshScale(cht)` — lê `Unit/AxisMin/AxisMax` (linha 2 do corpo), aplica ao eixo de valores primário (`MinimumScale/MaximumScale`), `Chart_SyncSecondaryX`, `Axes(xlValue, xlSecondary).MaximumScale = n`, e `NumberFormat` da coluna `TickLbl`: `w` → `"dd/mm"`, `m` → `"mmm/yy"`, `q`/`y` → `"General"`; nunca propaga erro (registra). `Gantt_SetLabels(cht, on)`; `Gantt_ApplyColors(cht, palette)` (`tqq_dur = pal(1)`, `tqq_ms = pal(0)`); `Gantt_Dates(cht, ByRef starts() As Double, ByRef ends() As Double)` (1..n, lidos das colunas `GT_COL_START`/`GT_COL_END` do bloco; `NA()` ou vazio = 0).
- Produces (`modAnnotations`): `Ann_AddBracket(cht, fromIdx, toIdx) As Shape` — recusa `7001, "O colchete só está disponível em gráficos Gantt."` quando `Not Kinds_Supports(kind, "bracket")`; cria linha `nm` (`kind=bracket;from=i;to=j`), tiques `nm_t1`/`nm_t2` (`kind=tick;owner=nm`, 5 pt para baixo nas pontas) e rótulo `nm_lbl`; garante `PlotArea.InsideTop >= 40 + 14 * (nº de colchetes)`; `Ann_Refresh` posiciona (válido com `1 <= i <= j <= n`: **`i = j` é válido** no colchete, ao contrário de CAGR/diferença/conectores, que continuam exigindo `i < j`): para cada atividade k em `i..j`, último dia `le_k = IF(fim_k > 0, fim_k, início_k)` e fim plotado `pe_k = IF(fim_k > 0, fim_k + 1, início_k)` (a barra vai até `Fim + 1`; o marco fica no próprio dia); `d1 = menor início`, `d2 = maior le_k` (zeros ignorados), `x1 = Geom_ChartValueToX(cht, d1)`, `x2 = Geom_ChartValueToX(cht, maior pe_k)`, `y = InsideTop - 20 - 14 * (k - 1)` (`k` = ordinal do colchete pelo sufixo numérico do nome), rótulo `Format$(d1, "dd/mm") & " - " & Format$(d2, "dd/mm")` centralizado acima da linha; `Ann_RemoveAll` remove colchetes e tiques; `Core_AddAnnotation` aceita `kind = "bracket"`: exige `n >= 1` (não 2) e só pede categorias quando `fromIdx < 1 Or toIdx > n Or fromIdx > toIdx` (com `de = até` não reabre o formulário); para `cagr`/`diff` a regra de hoje (`n >= 2` e `fromIdx < toIdx`) continua.
- `Core_SafeRefresh`: família `gantt` → `Gantt_RefreshScale cht` (antes de `Ann_Refresh`). `Labels_RefreshPositions`, `Labels_ApplyCellFormat`, `Style_ApplyNumberFormat`: família `gantt` sai cedo.
- `modCoreApi`, família `gantt` (os testes desta tarefa usam os dois pela API; a Tarefa 11 só reorganiza os despachos): `ApplyToggle "labels"` → `Gantt_SetLabels`; `ApplyOption "palette"` → `Gantt_ApplyColors`. Remova `Test_Core_Insert_KindNotReady_NotifiesWithoutChart` de `TestCoreApi.bas` (ele insere `gantt` esperando o aviso 7200); o ramo 7200 de `Kinds_Insert` fica como proteção, mas nenhum tipo válido chega mais a ele.

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestGantt.bas`:

```vb
Attribute VB_Name = "TestGantt"
Option Explicit
' Datas: 2026-03-02 e 2026-04-20 sao segundas-feiras (WEEKDAY(d,3) = 0).

Private Function Fixture(wb As Workbook, Optional withOwner As Boolean = False) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Plano"
    ws.Range("A1:C1").Value = Array("Atividade", "Início", "Fim")
    ws.Range("A2").Value = "Planejamento": ws.Range("B2").Value = DateSerial(2026, 3, 2): ws.Range("C2").Value = DateSerial(2026, 3, 13)
    ws.Range("A3").Value = "Execução": ws.Range("B3").Value = DateSerial(2026, 3, 9): ws.Range("C3").Value = DateSerial(2026, 4, 17)
    ws.Range("A4").Value = "Entrega": ws.Range("B4").Value = DateSerial(2026, 4, 20)
    ws.Range("B2:C4").NumberFormat = "dd/mm/yyyy"
    If withOwner Then
        ws.Range("D1").Value = "Responsável": ws.Range("D2:D4").Value = Application.Transpose(Array("Ana", "Bruno", "Carla"))
        Set Fixture = ChartData_FromRange(ws.Range("A1:D4"))
    Else
        Set Fixture = ChartData_FromRange(ws.Range("A1:C4"))
    End If
End Function

Private Function ValidTicks(body As Range) As Long
    Dim r As Long
    For r = 2 To GT_SCALE_SLOTS + 1
        If Not IsError(body.Cells(r, GT_COL_TICK).Value) Then ValidTicks = ValidTicks + 1
    Next
End Function

Public Sub Test_Gantt_Block_FormulasAndWeekScale()
    Dim wb As Workbook, blk As Range, body As Range
    Set wb = NewTempWorkbook()
    Set blk = Gantt_BuildBlock(wb, "gt0001", Fixture(wb)): Set body = Block_Body(blk)
    AssertEqual 19, blk.Columns.Count: AssertEqual 41, body.Rows.Count, "corpo: cabecalho + max(n, 40) (o bloco tem ainda a linha de meta)"
    AssertEqual CDbl(DateSerial(2026, 3, 2)), body.Cells(2, GT_COL_DMIN).Value: AssertEqual CDbl(DateSerial(2026, 4, 20)), body.Cells(2, GT_COL_DMAX).Value
    AssertEqual False, body.Cells(2, GT_COL_ISMS).Value: AssertEqual True, body.Cells(4, GT_COL_ISMS).Value
    AssertEqual CDbl(DateSerial(2026, 3, 2)), body.Cells(2, GT_COL_BARSTART).Value
    AssertEqual 12, body.Cells(2, GT_COL_DUR).Value, "fim inclusivo: 2 a 13 de marco"
    AssertEqual 40, body.Cells(3, GT_COL_DUR).Value
    AssertTrue IsError(body.Cells(4, GT_COL_DUR).Value), "marco nao tem barra"
    AssertTrue IsError(body.Cells(2, GT_COL_MSX).Value): AssertEqual CDbl(DateSerial(2026, 4, 20)), body.Cells(4, GT_COL_MSX).Value
    AssertEqual 2.5, body.Cells(2, GT_COL_POS).Value: AssertEqual 0.5, body.Cells(4, GT_COL_POS).Value
    AssertEqual CDbl(DateSerial(2026, 3, 14)), body.Cells(2, GT_COL_LBLX).Value, "rotulo depois do fim da barra (Fim + 1)"
    AssertEqual CDbl(DateSerial(2026, 4, 20)), body.Cells(4, GT_COL_LBLX).Value, "marco: no proprio dia"
    AssertEqual "w", body.Cells(2, GT_COL_UNIT).Value, "49 dias: semanas"
    AssertEqual CDbl(DateSerial(2026, 3, 2)), body.Cells(2, GT_COL_AXMIN).Value
    AssertEqual CDbl(DateSerial(2026, 4, 27)), body.Cells(2, GT_COL_AXMAX).Value, "segunda seguinte ao fim"
    AssertEqual 8, ValidTicks(body)
    AssertEqual CDbl(DateSerial(2026, 3, 9)), body.Cells(3, GT_COL_TICK).Value
    AssertEqual 3, body.Cells(2, GT_COL_TICKY).Value: AssertTrue IsError(body.Cells(10, GT_COL_TICKY).Value)
    CloseTemp wb
End Sub

Public Sub Test_Gantt_Insert_BarsMilestonesScale()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ms As Series, sc As Series, lbl As Series, v As Variant, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "gantt", Registry_ChartType(cht)
    AssertEqual xlBarStacked, Chart_FindSeries(cht, "tqq_start").ChartType
    AssertEqual msoFalse, Chart_FindSeries(cht, "tqq_start").Format.Fill.Visible
    AssertTrue cht.Axes(xlCategory).ReversePlotOrder
    AssertEqual CDbl(DateSerial(2026, 3, 2)), cht.Axes(xlValue, xlPrimary).MinimumScale
    AssertEqual CDbl(DateSerial(2026, 4, 27)), cht.Axes(xlValue, xlPrimary).MaximumScale
    Set ms = Chart_FindSeries(cht, "tqq_ms")
    AssertEqual xlSecondary, ms.AxisGroup: AssertEqual xlMarkerStyleDiamond, ms.MarkerStyle
    v = ms.XValues
    AssertEqual CDbl(DateSerial(2026, 4, 20)), v(3)
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    AssertEqual Format$(DateSerial(2026, 3, 13), "dd/mm"), lbl.Points(1).DataLabel.Text
    AssertEqual Format$(DateSerial(2026, 4, 20), "dd/mm"), lbl.Points(3).DataLabel.Text, "marco rotulado com a propria data"
    AssertEqual xlLabelPositionRight, lbl.Points(1).DataLabel.Position
    Set sc = Chart_FindSeries(cht, "tqq_scale")
    AssertEqual 40, sc.Points.Count
    AssertEqual Format$(DateSerial(2026, 3, 2), "dd/mm"), sc.Points(1).DataLabel.Text
    AssertEqual xlLabelPositionAbove, sc.Points(1).DataLabel.Position
    AssertTrue sc.HasErrorBars, "separadores verticais por barras de erro"
    AssertFalse cht.HasLegend
    AssertEqual "50", Registry_GetOption(cht, "gap"): AssertEqual "3", Registry_GetOption(cht, "n"): AssertEqual "1", Registry_GetOption(cht, "labels")
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual "dd/mm", body.Cells(2, GT_COL_TICKLBL).NumberFormat
    AssertEqual "dd/mm", body.Cells(2, GT_COL_LBLTXT).NumberFormat, "ligar os rotulos nao troca o formato de data das celulas"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    AssertEqual 1, CalcSheet_Cleanup(wb), "limpeza com o grafico novo (auxiliares com pontos NA) remove so o orfao"
    AssertFalse CalcSheet_CleanupReadFailed(), "pontos NA dos auxiliares nao sao falha de leitura"
    Gantt_SetLabels cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl") Is Nothing: AssertEqual "0", Registry_GetOption(cht, "labels")
    CloseTemp wb
End Sub

Public Sub Test_Gantt_Geometry_Independent()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lbl As Series, ax As Axis, il As Double, it As Double, iw As Double, ih As Double, x As Double, dl As DataLabel
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    il = cht.PlotArea.InsideLeft: it = cht.PlotArea.InsideTop: iw = cht.PlotArea.InsideWidth: ih = cht.PlotArea.InsideHeight
    Set ax = cht.Axes(xlValue, xlPrimary)
    x = il + iw * (CDbl(DateSerial(2026, 4, 18)) - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)      ' fim visual da barra = Fim + 1
    Set dl = lbl.Points(2).DataLabel
    AssertTrue dl.Left >= x - 2 And dl.Left <= x + 14, "rotulo da Execucao no fim da barra: " & dl.Left & " vs " & x
    AssertTrue Abs((dl.Top + dl.Height / 2) - (it + 1.5 * ih / 3)) <= ih / 6, "na faixa 2"
    x = il + iw * (CDbl(DateSerial(2026, 4, 20)) - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
    Set dl = lbl.Points(3).DataLabel
    AssertTrue dl.Left >= x - 2 And dl.Left <= x + 16, "rotulo do marco a direita do losango"
    AssertTrue Abs((dl.Top + dl.Height / 2) - (it + 2.5 * ih / 3)) <= ih / 6, "na faixa 3"
    Set dl = Chart_FindSeries(cht, "tqq_scale").Points(1).DataLabel
    AssertTrue dl.Top + dl.Height <= it + 2, "rotulo da escala acima da area de plotagem"
    AssertTrue Abs((dl.Left + dl.Width / 2) - il) <= 12, "primeira marca na borda esquerda (AxisMin)"
    AssertTrue dl.Top >= 0, "rotulo da escala dentro do grafico"
    AssertTrue cht.PlotArea.InsideTop >= 23.5, "folga acima da area de plotagem para a escala"
    CloseTemp wb
End Sub

' Review Focus 4: a unidade da escala muda com os dados e o refresh reaplica eixo e formatos.
Public Sub Test_Gantt_ScaleUnitSwitches_OnRefresh()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    Set body = Block_Body(Registry_Block(cht))
    d.Source.Worksheet.Range("C3").Value = DateSerial(2027, 1, 15)       ' 319 dias: meses
    AssertEqual "m", body.Cells(2, GT_COL_UNIT).Value
    Gantt_RefreshScale cht
    AssertEqual CDbl(DateSerial(2026, 3, 1)), cht.Axes(xlValue, xlPrimary).MinimumScale
    AssertEqual CDbl(DateSerial(2027, 2, 1)), cht.Axes(xlValue, xlPrimary).MaximumScale
    AssertEqual 11, ValidTicks(body), "mar/26 a jan/27"
    AssertEqual "mmm/yy", body.Cells(2, GT_COL_TICKLBL).NumberFormat
    d.Source.Worksheet.Range("C3").Value = DateSerial(2029, 6, 30)       ' > 1100 dias: trimestres
    Core_SafeRefresh cht
    AssertEqual "q", body.Cells(2, GT_COL_UNIT).Value
    AssertEqual CDbl(DateSerial(2026, 1, 1)), cht.Axes(xlValue, xlPrimary).MinimumScale
    AssertEqual CDbl(DateSerial(2029, 7, 1)), cht.Axes(xlValue, xlPrimary).MaximumScale
    AssertEqual 14, ValidTicks(body)
    AssertEqual "T1/2026", body.Cells(2, GT_COL_TICKLBL).Value
    AssertEqual "T1/2026", Chart_FindSeries(cht, "tqq_scale").Points(1).DataLabel.Text
    CloseTemp wb
End Sub

Public Sub Test_Gantt_InvalidDates_RefusedWithMessage()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, en As Long, ed As String, nm As Name
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("Atividade", "Início", "Fim")
    ws.Range("A2:C2").Value = Array("Planejamento", "2 de março", "13 de março")
    ws.Range("A3:C3").Value = Array("Execução", "depois", "")
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    On Error Resume Next
    Gantt_Insert ws, d
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 6010, en
    AssertEqual "Gantt: a coluna de início não tem datas válidas (use células de data, não texto).", ed
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nada alocado"
    Next
    ws.Range("A1:B3").Select
    Core_ClearLastNotice
    Core_Insert "gantt"
    AssertTrue InStr(Core_LastNotice(), "3 colunas") > 0, Core_LastNotice()
    Watchers_Stop
    CloseTemp wb
End Sub

Public Sub Test_Gantt_OnlyMilestones_And_SingleActivity()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("Atividade", "Início", "Fim")
    ws.Range("A2").Value = "Kick-off": ws.Range("B2").Value = DateSerial(2026, 5, 6)
    Set d = ChartData_FromRange(ws.Range("A1:C2"))
    Set cht = Gantt_Insert(ws, d).Chart
    AssertEqual "1", Registry_GetOption(cht, "n")
    AssertEqual 1, Chart_FindSeries(cht, "tqq_ms").Points.Count
    AssertEqual 1, cht.Axes(xlValue, xlSecondary).MaximumScale
    AssertEqual CDbl(DateSerial(2026, 5, 4)), cht.Axes(xlValue, xlPrimary).MinimumScale, "segunda anterior"
    AssertEqual CDbl(DateSerial(2026, 5, 11)), cht.Axes(xlValue, xlPrimary).MaximumScale
    AssertEqual Format$(DateSerial(2026, 5, 6), "dd/mm"), Chart_FindSeries(cht, "tqq_lbl").Points(1).DataLabel.Text
    CloseTemp wb
End Sub

Public Sub Test_Gantt_FourthColumn_AsLabel()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, True)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "Ana", Chart_FindSeries(cht, "tqq_lbl").Points(1).DataLabel.Text
    AssertEqual "Carla", Chart_FindSeries(cht, "tqq_lbl").Points(3).DataLabel.Text
    CloseTemp wb
End Sub

Public Sub Test_Gantt_Bracket_Annotation()
    Dim wb As Workbook, d As clsChartData, cht As Chart, br As Shape, lbl As Shape, ax As Axis, il As Double, iw As Double, en As Long, ed As String, other As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    Set br = Ann_AddBracket(cht, 1, 2)
    AssertEqual "tqq_ann_bracket_1", br.Name
    AssertEqual "kind=bracket;from=1;to=2", br.AlternativeText
    AssertTrue Not cht.Shapes("tqq_ann_bracket_1_t1") Is Nothing And Not cht.Shapes("tqq_ann_bracket_1_t2") Is Nothing
    Set lbl = cht.Shapes("tqq_ann_bracket_1_lbl")
    AssertEqual Format$(DateSerial(2026, 3, 2), "dd/mm") & " - " & Format$(DateSerial(2026, 4, 17), "dd/mm"), lbl.TextFrame2.TextRange.Text
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth: Set ax = cht.Axes(xlValue, xlPrimary)
    AssertAlmost il + iw * (CDbl(DateSerial(2026, 3, 2)) - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale), br.Left, 1
    AssertAlmost iw * (CDbl(DateSerial(2026, 4, 18)) - CDbl(DateSerial(2026, 3, 2))) / (ax.MaximumScale - ax.MinimumScale), br.Width, 1, "ate o fim visual da barra (Fim + 1)"
    AssertTrue br.Top < cht.PlotArea.InsideTop - 10, "colchete acima da area de plotagem"
    AssertTrue lbl.Top + lbl.Height <= br.Top + 0.5, "rotulo acima do colchete"
    AssertTrue lbl.Top >= 0
    Ann_AddBracket cht, 3, 3
    AssertTrue cht.Shapes("tqq_ann_bracket_2").Top < br.Top, "segundo colchete empilha acima"
    Ann_RemoveAll cht
    AssertEqual 0, cht.Shapes.Count
    d.Source.Worksheet.Range("F1:G1").Value = Array("Ano", "V"): d.Source.Worksheet.Range("F2:G2").Value = Array(2022, 1): d.Source.Worksheet.Range("F3:G3").Value = Array(2023, 2)
    Set other = Stacked_Insert(d.Source.Worksheet, ChartData_FromRange(d.Source.Worksheet.Range("F1:G3")), "stacked").Chart
    On Error Resume Next
    Ann_AddBracket other, 1, 2
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7001, en: AssertEqual "O colchete só está disponível em gráficos Gantt.", ed
    CloseTemp wb
End Sub

Public Sub Test_Core_Insert_Gantt_And_Rollback()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, n As Long, nm As Name
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C4").Select
    Core_Insert "gantt"
    Set cht = Core_ActiveChart()
    AssertEqual "gantt", Registry_ChartType(cht)
    AssertTrue Core_Supports("bracket"): AssertFalse Core_Supports("numfmt"): AssertTrue Core_Supports("gap")
    Core_SetOption "gap", "80"
    AssertEqual 80, cht.ChartGroups(1).GapWidth
    Core_AddAnnotation "bracket", 1, 3
    AssertTrue Not cht.Shapes("tqq_ann_bracket_1") Is Nothing
    Core_AddAnnotation "bracket", 2, 2
    AssertTrue Not cht.Shapes("tqq_ann_bracket_2") Is Nothing, "colchete sobre uma unica atividade (de = ate) sem reabrir o formulario"
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(127, 127, 127), Chart_FindSeries(cht, "tqq_dur").Format.Fill.ForeColor.RGB, "Gantt_ApplyColors pela API: tqq_dur = pal(1)"
    Core_SetToggle "labels", False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl") Is Nothing
    Core_ClearLastNotice
    Core_Refresh
    AssertEqual 0, Core_NoticeCount(), "sem aviso de tabela mudada"
    ws.ChartObjects(1).Delete: CalcSheet_Cleanup wb
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Gantt_Insert ws, d
    n = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, n
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    Watchers_Stop
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "TestGantt"` → saída 3.

- [ ] **Passo 3: Implementar** — fórmulas da escala (linha 2 do corpo; `U = célula Unit`, `I = coluna Inicio das linhas 2..n+1`, `F = Fim`):

```
Unit    = IF(dmax-dmin<=98,"w",IF(dmax-dmin<=1100,"m",IF(dmax-dmin<=3700,"q","y")))      com dmin = célula DMin (col 18) e dmax = célula DMax (col 19), não inline
AxisMin = IF(U="w",dmin-WEEKDAY(dmin,3),IF(U="m",DATE(YEAR(dmin),MONTH(dmin),1),IF(U="q",DATE(YEAR(dmin),3*INT((MONTH(dmin)-1)/3)+1,1),DATE(YEAR(dmin),1,1))))
AxisMax = IF(U="w",dmax-WEEKDAY(dmax,3)+7,IF(U="m",EDATE(DATE(YEAR(dmax),MONTH(dmax),1),1),IF(U="q",EDATE(DATE(YEAR(dmax),3*INT((MONTH(dmax)-1)/3)+1,1),3),DATE(YEAR(dmax)+1,1,1))))
Tick_k  = IF(U="w",AxisMin+7*k,EDATE(AxisMin,k*IF(U="m",1,IF(U="q",3,12))))  -> IF(tick<AxisMax,tick,NA())   (k = linha - 2)
TickY   = IF(ISNUMBER(Tick),n,NA())
TickLbl = IF(U="q","T"&(INT((MONTH(Tick)-1)/3)+1)&"/"&YEAR(Tick),IF(U="y",YEAR(Tick),Tick))
```
`dmin` e `dmax` ficam nas colunas `GT_COL_DMIN = 18` e `GT_COL_DMAX = 19` (linha 2 do corpo; o bloco tem **19** colunas, como nos testes): `DMin = AGGREGATE(5,6,I)`, `DMax = AGGREGATE(4,6,I,F)`; as fórmulas acima referenciam essas duas células. Barras de erro: `sc.ErrorBar Direction:=xlY, Include:=xlErrorBarIncludeMinusValues, Type:=xlErrorBarTypeFixedValue, Amount:=n` depois de definir `Values/XValues`; `sc.ErrorBars.EndStyle = xlNoCap`; `sc.ErrorBars.Format.Line.ForeColor.RGB = RGB(217,217,217)`, `DashStyle = msoLineDash`. Se as barras de erro descerem para o lado errado (visível no PNG `cht.Export`), inverta para `xlErrorBarIncludePlusValues` e registre.

- [ ] **Passo 4: GREEN e suíte completa** — exporte um PNG (`cht.Export TempDir() & "\gantt.png"`) e olhe: escala no topo, separadores descendo, losango no marco, rótulos à direita.

- [ ] **Passo 5: Commit** — `git commit -m "feat: Gantt com barras, marcos, escala de datas por formula (semana/mes/trimestre/ano) e colchetes"` (+ trailer).

**Verificação no Excel real:** tabela com 3 atividades (uma sem fim = marco); escala semanal no topo com separadores; mudar o fim para o ano seguinte e clicar Atualizar: escala mensal; Setas > Colchete (Gantt); Espaço muda o espaçamento das linhas; 4.ª coluna "Responsável" aparece à direita das barras.

---

### Tarefa 11: Integração da API para todos os tipos (toggles, opções, reconstrução, anotações)

**Files:**
- Modify: `src/core/modules/modCoreApi.bas`, `src/core/modules/modChartKinds.bas` (`Kinds_FeatureLabel`), `src/core/modules/modLabels.bas`, `src/core/modules/modChartStyle.bas` (varredura final das famílias)
- Test: `tests/vba/core/TestCoreApi.bas` (adicionar os testes abaixo; o teste temporário `Test_Core_Insert_KindNotReady_NotifiesWithoutChart` já saiu na Tarefa 10)

**Interfaces:**
- Produces: `Kinds_FeatureLabel(feature) As String` (`totals` → `Totais`, `labels` → `Rótulos`, `connectors` → `Conectores`, `ind100` → `Indicador 100%`, `vline` → `Linha de valor (média)`, `secondary` → `Eixo secundário (linha)`, `pieout` → `Rótulos fora da pizza`, `gap` → `Espaço`, `legend` → `Legenda`, `numfmt` → `Formato`, `palette` → `Paleta`) — os rótulos exatos dos controles da faixa (Tarefa 12), para o aviso citar o botão que o usuário clicou.
- `ApplyToggle(c, key, value)`: chaves fora da lista de toggles (`totals, labels, connectors, ind100, vline, secondary, pieout`) saem sem fazer nada (silenciosas, como hoje); para uma chave da lista, se `Not Kinds_Supports(kind, key)` → `Err.Raise vbObjectError + 7002, "Core", "'" & Kinds_FeatureLabel(key) & "' não se aplica a este tipo de gráfico."`; senão despacha por família (`totals`: stacked/combo → `Stacked_SetTotals`, area → `Area_SetTotals`, mekko → `Mekko_SetTotals`; `labels`: stacked/combo → `Stacked_SetLabels`, waterfall → `Waterfall_SetLabels`, area → `Area_SetLabels`, mekko → `Mekko_SetLabels`, gantt → `Gantt_SetLabels`, pie → `Pie_SetLabels`; `connectors`: waterfall → `Ann_SetWaterfallConnectors`, demais → `Stacked_SetConnectors`; `ind100` → `Deco_SetInd100`; `vline` → `Deco_SetValueLine`; `secondary` → `Combo_SetSecondary`; `pieout` → `Pie_SetOutside`). `Core_SetToggle` trata 7002 como aviso (já feito na Tarefa 6); chaves desconhecidas continuam silenciosas (`Test_Core_InvalidKeysAndValues_DoNotRaise`).
- `ApplyOption(c, key, value)`: valida o valor (como hoje) e, **antes de gravar**, recusa `gap` e `legend` não suportados com `7002` e `"Espaço entre colunas não se aplica a este tipo de gráfico."` / `"Legenda não se aplica a este tipo de gráfico."` (o ramo especial de `gap` em linha de `Core_SetOption` é substituído por este; `Test_Core_SetOption_Gap_OnLineChart_NotPersisted` continua válido); `numfmt` → `Style_ApplyNumberFormat` (que despacha por família: `pie` → `Pie_ApplyNumberFormat`, `gantt` → nada) ; `palette` por família (`waterfall` → `Waterfall_ApplyColors`, `gantt` → `Gantt_ApplyColors`, `pie` → `Style_ApplyPointColors`, `mekko` → `Mekko_ApplyColors`, demais → `Style_ApplySeriesColors`); `legend` → `Deco_SetLegend`, que só grava ao aceitar (Tarefa 6; precedência com `pieout` na pizza: Tarefa 8).
- `Core_AddAnnotation(kind, from, to)`: aceita `cagr|diff|bracket`; **antes de qualquer outra verificação** (inclusive a de `n < 2`, que hoje vem primeiro e recusaria uma pizza ou um Gantt de uma linha com a mensagem errada) verifica `Kinds_Supports(chartKind, kind)` e notifica a mesma mensagem de `Ann_Add*` (sem abrir formulário); depois aplica as regras de índices da Tarefa 10 (`bracket`: `n >= 1` e `de <= até`; `cagr`/`diff`: `n >= 2` e `de < até`); `bracket` → `Ann_AddBracket` (despacho da Tarefa 10).
- `Core_RebuildFromRange(c, rng)`: constrói com `Kinds_Insert(ws, data, kind)` (desde a Tarefa 1); copia as opções do bloco antigo; reaplica, **só quando `Kinds_Supports`**, `numfmt`, `palette`, `gap`, `legend`, e os toggles `totals, labels, connectors, ind100, vline, secondary, pieout` (valor do registro; `labels` com padrão `1`, os demais `0`; `ind100` com padrão `1` em tipos 100%); por fim `Deco_ApplyFixedScale`. Mensagens e rollback como hoje.
- `Core_Refresh` (sem gráfico ativo) continua varrendo `ws.ChartObjects` com prefixo `tqq_`.

- [ ] **Passo 1: Testes (RED)** — em `TestCoreApi.bas`:

```vb
' Tabela de dados adequada a cada tipo, escrita a partir de A1 de ws; devolve a faixa.
Private Function FixtureFor(ws As Worksheet, kind As String) As Range
    Select Case Kinds_Family(kind)
        Case "waterfall"
            ws.Range("A1:C1").Value = Array("Etapa", "Loja", "Online")
            ws.Range("A2:C2").Value = Array("Inicio", 100, 50): ws.Range("A3:C3").Value = Array("Ajuste", -30, 10)
            ws.Range("A4:C4").Value = Array("Fim", "e", "e")
            Set FixtureFor = ws.Range("A1:C4")
        Case "mekko"
            ws.Range("A1:D1").Value = Array("Regiao", "Largura", "Eletro", "Moveis")
            ws.Range("A2:D2").Value = Array("Sul", 50, 40, 10): ws.Range("A3:D3").Value = Array("Norte", 30, 20, 10)
            If kind = "mekko" Then ws.Columns("B").Delete          ' percentual: sem coluna de largura
            Set FixtureFor = ws.Range("A1").CurrentRegion
        Case "gantt"
            ws.Range("A1:C1").Value = Array("Atividade", "Início", "Fim")
            ws.Range("A2").Value = "A": ws.Range("B2").Value = DateSerial(2026, 3, 2): ws.Range("C2").Value = DateSerial(2026, 3, 13)
            ws.Range("A3").Value = "B": ws.Range("B3").Value = DateSerial(2026, 3, 16)
            Set FixtureFor = ws.Range("A1:C3")
        Case "pie"
            ws.Range("A1:B1").Value = Array("Produto", "Vendas")
            ws.Range("A2:B2").Value = Array("Eletro", 50): ws.Range("A3:B3").Value = Array("Moveis", 30)
            Set FixtureFor = ws.Range("A1:B3")
        Case Else
            ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
            ws.Range("A2:D2").Value = Array(2022, 10, 20, 5): ws.Range("A3:D3").Value = Array(2023, 12, 22, 6)
            ws.Range("A4:D4").Value = Array(2024, 15, 25, 8)
            Set FixtureFor = ws.Range("A1:D4")
    End Select
End Function

' Todo tipo: insere pela API, exercita cada toggle/opcao conforme a matriz (suportado = funciona; nao suportado = aviso,
' nada persiste), reconstroi mantendo o tipo e as opcoes, e aceita/recusa setas pela matriz.
Public Sub Test_Core_AllKinds_InsertTogglesOptionsRebuild()
    Dim kinds As Variant, toggles As Variant, k As Long, t As Long, wb As Workbook, ws As Worksheet, rng As Range, c As Chart, kind As String, key As String
    kinds = Array("stacked", "clustered", "stacked100", "line", "waterfall", "barstacked", "barclustered", "barstacked100", "barwaterfall", _
                  "area", "area100", "combo", "mekko", "mekkounits", "gantt", "pie", "doughnut")
    toggles = Array("totals", "labels", "connectors", "ind100", "vline", "secondary", "pieout")
    For k = LBound(kinds) To UBound(kinds)
        kind = kinds(k)
        Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
        Set rng = FixtureFor(ws, kind)
        rng.Select
        Core_ClearLastNotice
        Core_Insert kind
        Set c = Core_ActiveChart()
        AssertTrue Not c Is Nothing, kind & ": nao inseriu (" & Core_LastNotice() & ")"
        AssertEqual kind, Registry_ChartType(c)
        For t = LBound(toggles) To UBound(toggles)
            key = toggles(t)
            Core_ClearLastNotice
            If Kinds_Supports(kind, key) Then
                Core_SetToggle key, True
                AssertTrue Core_GetToggle(key), kind & "/" & key & " ligado"
                Core_SetToggle key, False
                AssertFalse Core_GetToggle(key), kind & "/" & key & " desligado"
                Core_SetToggle key, True
            Else
                Core_SetToggle key, True
                AssertEqual "'" & Kinds_FeatureLabel(key) & "' não se aplica a este tipo de gráfico.", Core_LastNotice(), kind & "/" & key
                AssertFalse Core_GetToggle(key), kind & "/" & key & " nao persiste"
            End If
        Next
        If Kinds_Supports(kind, "numfmt") Then Core_SetOption "numfmt", "dec1": AssertEqual "dec1", Core_GetOption("numfmt"), kind
        Core_SetOption "palette", "cinza": AssertEqual "cinza", Core_GetOption("palette"), kind
        Core_ClearLastNotice
        Core_SetOption "gap", "80"
        If Kinds_Supports(kind, "gap") Then
            AssertEqual "80", Core_GetOption("gap"), kind & "/gap"
        Else
            AssertEqual "Espaço entre colunas não se aplica a este tipo de gráfico.", Core_LastNotice(), kind & "/gap"
        End If
        If Kinds_Supports(kind, "legend") Then Core_SetOption "legend", "none": AssertFalse c.HasLegend, kind & "/legend"
        Core_ClearLastNotice
        AssertTrue Core_RebuildFromRange(c, rng), kind & ": rebuild (" & Core_LastNotice() & ")"
        Set c = Core_ActiveChart()
        AssertEqual kind, Registry_ChartType(c), kind & ": tipo apos rebuild"
        If Kinds_Supports(kind, "numfmt") Then AssertEqual "dec1", Core_GetOption("numfmt"), kind & ": numfmt apos rebuild"
        AssertEqual "cinza", Core_GetOption("palette"), kind & ": paleta apos rebuild"
        If Kinds_Supports(kind, "vline") Then AssertTrue Not Chart_FindSeries(c, "tqq_vline") Is Nothing, kind & ": linha de valor apos rebuild"
        AssertEqual 1, TqqNameCount(wb), kind & ": um bloco"
        Core_ClearLastNotice
        Core_AddAnnotation "cagr", 1, 2
        If Kinds_Supports(kind, "cagr") Then
            AssertTrue Not c.Shapes("tqq_ann_cagr_1") Is Nothing, kind & ": CAGR"
        Else
            AssertEqual "A seta CAGR não está disponível para este tipo de gráfico.", Core_LastNotice(), kind & ": recusa CAGR"
        End If
        Teardown wb
    Next
End Sub

Public Sub Test_Core_AddAnnotation_Unsupported_NotifiesBeforeForm()
    Dim wb As Workbook, ws As Worksheet, rng As Range
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set rng = FixtureFor(ws, "stacked")
    rng.Select
    Core_Insert "stacked"
    Core_ClearLastNotice
    Core_AddAnnotation "bracket"                       ' sem indices: nao pode abrir formulario nem falhar
    AssertEqual "O colchete só está disponível em gráficos Gantt.", Core_LastNotice()
    AssertEqual 1, Core_NoticeCount(), "um unico aviso (conta desde Core_ClearLastNotice)"
    Teardown wb
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "AllKinds|NotifiesBeforeForm"` → `not ok` (mensagens genéricas ainda não existem; `bracket` não despachado).

- [ ] **Passo 3: Implementar** — concentre os despachos por família em três `Select Case` (`ApplyToggle`, `ApplyOption`, `Core_AddAnnotation`); mantenha `Core_RebuildFromRange` com a mesma estrutura de rollback (`FailNew`); a lista de toggles/opções a reaplicar fica numa única função privada `ReapplyOptions(nc)` usada só ali.

- [ ] **Passo 4: GREEN e suíte completa** — meça o tempo total; se passar de 5 minutos, reduza as fixtures do teste de todos os tipos (não remova tipos).

- [ ] **Passo 5: Commit** — `git commit -m "feat: API do core unificada para todos os tipos (toggles e opcoes pela matriz, reconstrucao generica, colchete)"` (+ trailer).

**Verificação no Excel real:** para 3 tipos diferentes (barras, Mekko, Gantt): Alterar dados mantém tipo e opções; botões não aplicáveis avisam em vez de falhar.

---

### Tarefa 12: Faixa de opções em menus, habilitação por capacidade e tutorial com os rótulos novos (loader)

**Files:**
- Modify: `src/loader/ribbon/customUI14.xml`, `src/loader/modules/modRibbon.bas`, `tests/vba/loader/TestRibbon.bas`, `tests/ps/Test-Docs.ps1`, `docs/tutorial/tutorial.html` (maquete da faixa e tabela "Tipos de gráfico"; as seções completas vêm na Tarefa 14)
- Test: `tests/vba/loader/TestRibbon.bas`, `tests/ps/Test-Ribbon.ps1` (inalterado), `tests/ps/Test-Docs.ps1`

**Interfaces:**
- Consumes (core 1.1.0): `Core_Supports(feature)`, `Core_SameScale()`, `Core_ResetScale()`, `Core_Insert(kind)` com os 17 tipos, `Core_AddAnnotation("bracket")`, opção `legend`. Com core 1.0.0, `Loader_RunCore` devolve `Empty` para procedimentos inexistentes: a faixa cai nos comportamentos antigos.
- Produces (XML; ids e rótulos exatos — o tutorial e os testes dependem deles):

```
grpInsert "Inserir gráfico"
  menu mnuColumns   "Colunas"        size=large imageMso=ChartTypeColumnInsertGallery
      btnStacked "Colunas empilhadas" tag=stacked | btnClustered "Colunas agrupadas" tag=clustered | btnStacked100 "Colunas 100%" tag=stacked100
  menu mnuBars      "Barras"         size=large imageMso=ChartTypeBarInsertGallery
      btnBarStacked "Barras empilhadas" tag=barstacked | btnBarClustered "Barras agrupadas" tag=barclustered | btnBarStacked100 "Barras 100%" tag=barstacked100
  menu mnuWaterfall "Cascata"        size=large imageMso=ChartTypeOtherInsertGallery
      btnWaterfall "Cascata (colunas)" tag=waterfall | btnBarWaterfall "Cascata (barras)" tag=barwaterfall
  menu mnuLines     "Linha e área"   size=large imageMso=ChartTypeLineInsertGallery
      btnLine "Linha" tag=line | btnArea "Área empilhada" tag=area | btnArea100 "Área 100%" tag=area100 | btnCombo "Combinação (colunas + linha)" tag=combo
  menu mnuOthers    "Mais tipos"     size=large imageMso=ChartTypePieInsertGallery
      btnMekko "Mekko (%)" tag=mekko | btnMekkoUnits "Mekko (unidades)" tag=mekkounits | btnGantt "Gantt" tag=gantt | btnPie "Pizza" tag=pie | btnDoughnut "Rosca" tag=doughnut
  (todos os botões: onAction=Rx_OnInsert; supertips curtos com o formato da tabela esperada)
  (nenhum filho de menu leva size: só o próprio menu no grupo tem size=large; um atributo inválido faz o Excel
   descartar a aba inteira, em silêncio)
grpChart "Gráfico selecionado"
  tglTotals, tglLabels, tglConnectors (como hoje; getEnabled=Rx_GetEnabled, getPressed=Rx_GetPressed)
  separator
  ddFormat, ddPalette, ddGap (como hoje) | ddLegend "Legenda" tag=legend: item right "Direita", none "Nenhuma", lineend "Nome no fim da linha"
  separator
  menu mnuMore "Mais opções" imageMso=PropertySheet   (sem getEnabled: contém botões sempre ativos)
      tglInd100 "Indicador 100%" tag=ind100 | tglVline "Linha de valor (média)" tag=vline | tglSecondary "Eixo secundário (linha)" tag=secondary | tglPieOut "Rótulos fora da pizza" tag=pieout
      menuSeparator | btnSameScale "Mesma escala" onAction=Rx_OnSameScale | btnResetScale "Escala automática" onAction=Rx_OnResetScale
  menu mnuArrows "Setas" imageMso=ArrowsMore getEnabled=Rx_GetEnabled (sem tag: há gráfico)
      btnCagr "Seta CAGR" tag=cagr | btnDiff "Seta diferença" tag=diff | btnBracket "Colchete (Gantt)" tag=bracket | menuSeparator | btnRemoveAnn "Remover setas"
  separator | btnRefresh "Atualizar" (sempre) | btnChangeSource "Alterar dados" getEnabled=Rx_GetEnabled
grpAbout: inalterado
```
  `imageMso` dos itens novos: use nomes já validados na Fase 1 quando fizer sentido (`PercentStyle` para Indicador 100%, `ShapeStraightConnector` para Linha de valor, `ChartAxes` para Eixo secundário e Mesma escala, `Undo` para Escala automática, `TextBoxInsert` para Rótulos fora, `ChartTypePieInsertGallery` para Mais tipos); `Test_Ribbon_ImageMso_AreValid` decide: troque qualquer nome reprovado por `PropertySheet`.
- Produces (`modRibbon`): `Rx_GetEnabled` → com `control.Tag` não vazio, `SupportsCached(tag)`; sem tag, `HasChartCached()`; `SupportsCached` guarda num `Scripting.Dictionary` por ciclo (limpo em `Loader_InvalidateRibbon` e `Rx_OnLoad`) o resultado de `Loader_RunCore("Core_Supports", tag)` e cai em `HasChartCached()` quando vem `Empty` (core antigo); `Rx_OnSameScale(control)` e `Rx_OnResetScale(control)` (`On Error Resume Next`; `CoreOrWarn`; `Loader_RunCore`; `Loader_InvalidateRibbon`); `Rx_GetSelectedId` com padrão `"right"` para `legend`.
- Produces (`Test-Docs.ps1`): a verificação "todo controle da faixa aparece no tutorial" passa a funcionar de fato e inclui `menu`. **Hoje ela não verifica nada:** o XPath `//*[self::button or self::toggleButton or ...]` não casa nenhum nó, porque o `customUI` tem namespace padrão (0 nós; com `local-name()` são 39). Troque por `//*[local-name()='button' or local-name()='toggleButton' or local-name()='dropDown' or local-name()='menu' or local-name()='group' or local-name()='tab']`. Como a verificação nunca rodou, rótulos da Fase 1 podem faltar no tutorial (a aba e o grupo `ThinQQ`, os grupos `Inserir gráfico` e `Gráfico selecionado`, os botões do grupo `ThinQQ`): acrescente os que faltarem em `<span class="ui">`. Limiares `known >= 60` e `ui >= 45`. O limite de tamanho do tutorial sobe **já nesta tarefa** de `-le 61440` para `-le 122880` (120 KB): a maquete e a tabela de tipos podem passar de 60 KB antes da Tarefa 14.
- Produces (`tutorial.html`, parcial): maquete da faixa com os cinco menus e os dois menus do grupo do gráfico (cada rótulo em `<span class="ui">`), e uma seção `<section id="tipos">` "Tipos de gráfico" com uma tabela: um item por botão de inserção (rótulo exato + 1 frase + formato da tabela), mais uma linha por controle novo (`Legenda`, `Mais opções` e seus itens, `Setas` e `Colchete (Gantt)`, `Mesma escala`, `Escala automática`). Conteúdo real, não esqueleto; a Tarefa 14 aprofunda.

- [ ] **Passo 1: Testes (RED)** — em `tests/vba/loader/TestRibbon.bas`:

```vb
Public Sub Test_Ribbon_Menus_Structure()
    Dim d As Object, m As Object, kids As Object, n As Object, t As String
    Set d = LoadRibbon()
    AssertEqual 7, d.SelectNodes("//u:menu").Length, "5 menus de insercao + Mais opcoes + Setas"
    AssertEqual 5, d.SelectNodes("//u:group[@id='grpInsert']/u:menu[@size='large']").Length
    AssertEqual 0, d.SelectNodes("//u:menu/*[@size]").Length, "filho de menu com 'size' (invalido: o Excel descarta a aba inteira sem aviso)"
    For Each m In d.SelectNodes("//u:group[@id='grpInsert']/u:menu")
        Set kids = m.SelectNodes("u:button")
        AssertTrue kids.Length >= 2, "menu " & m.Attributes.getNamedItem("id").Text & " com poucos itens"
        For Each n In kids
            AssertEqual "Rx_OnInsert", n.Attributes.getNamedItem("onAction").Text
            t = n.Attributes.getNamedItem("tag").Text
            AssertTrue InStr(",stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,gantt,pie,doughnut,", "," & t & ",") > 0, "kind invalido: " & t
        Next
    Next
    AssertEqual 17, d.SelectNodes("//u:button[@onAction='Rx_OnInsert']").Length
    AssertTrue d.SelectSingleNode("//u:menu[@id='mnuMore']/@getEnabled") Is Nothing, "Mais opcoes contem botoes sempre ativos"
    AssertEqual 4, d.SelectNodes("//u:menu[@id='mnuMore']/u:toggleButton").Length
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnSameScale'][@onAction='Rx_OnSameScale']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnResetScale'][@onAction='Rx_OnResetScale']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuArrows']/u:button[@tag='bracket']") Is Nothing
    AssertEqual "Rx_GetEnabled", d.SelectSingleNode("//u:menu[@id='mnuArrows']").Attributes.getNamedItem("getEnabled").Text
End Sub

Public Sub Test_Ribbon_Toggles_And_Dropdowns_Phase2()
    Dim d As Object, n As Object, t As String
    Set d = LoadRibbon()
    AssertEqual 7, d.SelectNodes("//u:toggleButton").Length
    For Each n In d.SelectNodes("//u:toggleButton")
        t = n.Attributes.getNamedItem("tag").Text
        AssertTrue InStr(",totals,labels,connectors,ind100,vline,secondary,pieout,", "," & t & ",") > 0, "toggle invalido: " & t
        AssertEqual "Rx_GetEnabled", n.Attributes.getNamedItem("getEnabled").Text, "toggle sem getEnabled: " & t
        AssertEqual "Rx_GetPressed", n.Attributes.getNamedItem("getPressed").Text
    Next
    AssertEqual 4, d.SelectNodes("//u:dropDown").Length
    AssertEqual "right,none,lineend", ItemIds(d.SelectSingleNode("//u:dropDown[@tag='legend']"))
    For Each n In d.SelectNodes("//u:button[@onAction='Rx_OnAnnotation']")
        t = n.Attributes.getNamedItem("tag").Text
        AssertTrue InStr(",cagr,diff,bracket,", "," & t & ",") > 0, "annotation kind invalido: " & t
    Next
End Sub
```

Atualize os testes existentes: `Test_Ribbon_Xml_HasAllCallbacks` inclui `Rx_OnSameScale`, `Rx_OnResetScale` e `tag="gantt"`; `ExpectedSig` **não precisa mudar** (já devolve `iribboncontrol,byref *` para qualquer `getEnabled`, inclusive o de `menu`, e um `toggleButton` dentro de `menu` tem o mesmo `baseName`); `Test_Ribbon_Tags_MatchCoreKeys` passa a esperar 17 botões de inserção e 7 toggles (ou delegue a `Test_Ribbon_Menus_Structure`/`..._Phase2` e remova as contagens antigas); `Test_Ribbon_Dropdowns_ItemsMatchCoreOptions` espera 4 dropdowns; `Test_Ribbon_SelectedChartGroup_EnabledExceptRefresh`: sem `getEnabled` só `btnRefresh` e `mnuMore` (dentro de `mnuMore`, os toggles têm `getEnabled` e os dois botões de escala não).

`tests/ps/Test-Docs.ps1`: no laço "tutorial documenta o controle", XPath `//*[local-name()='button' or local-name()='toggleButton' or local-name()='dropDown' or local-name()='menu' or local-name()='group' or local-name()='tab']` (o atual, com `self::`, casa 0 nós por causa do namespace padrão); `Check ($known.Count -ge 60)`; `Check ($ui.Count -ge 45)`; `Check ($bytes -le 122880) "tutorial com até 120 KB ($bytes bytes)"`.

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-f2\build\run-tests.ps1 -Target loader -TimeoutSec 300` (sem filtro, para rodar também `Test-Ribbon.ps1` e `Test-Docs.ps1`): `not ok` nos testes novos; `Test-Docs` reprova os rótulos ausentes no tutorial.

- [ ] **Passo 3: Implementar** — XML e `modRibbon` conforme acima; na maquete do tutorial, troque os cinco botões por cinco "menus" com seta (`▾` em HTML é permitido: o tutorial não passa pelo lint cp1252) e acrescente a tabela "Tipos de gráfico". Rode também `pwsh C:\temp\tqq-f2\build\build.ps1 -Target loader -Configuration Release` e abra `dist\ThinQQ.xlam` **numa instância própria do Excel de automação** (`New-ExcelApp` de `Build-Common.ps1`, visível, com `$xl.Visible = $true`) só para ver a faixa renderizada; feche-a com `Close-ExcelApp`. Nunca instale.

- [ ] **Passo 4: GREEN e suíte completa** (`core` + `loader` + PowerShell).

- [ ] **Passo 5: Commit** — `git commit -m "feat: faixa de opcoes com menus por familia de grafico, controles habilitados por capacidade, Mesma escala e Escala automatica"` (+ trailer).

**Verificação no Excel real (usuário):** aba ThinQQ com os 5 menus; com uma pizza selecionada, Totais/Conectores/Espaço ficam cinza e Rótulos fora da pizza habilitado; com célula selecionada só Atualizar/Mesma escala/Escala automática ativos; nenhum ícone em branco.

---

### Tarefa 13: Exemplos (planilha "Exemplos 2")

**Files:**
- Modify: `src/core/modules/modSamples.bas`
- Test: `tests/vba/core/TestSamples.bas`

**Interfaces:**
- Consumes: todas as `*_Insert`/`Kinds_Insert`, `Deco_SetValueLine`, `Deco_SameScale`, `Pie_SetOutside`, `Ann_AddBracket`, `Ann_AddDifference`.
- Produces: `Samples_Create` mantém a planilha "Exemplos" (6 exemplos, intocados) e cria "Exemplos 2" (depois dela; `Samples_CreateSheet2(wb) As Worksheet`, privada) com 12 seções, títulos na coluna A nas linhas `1, 22, 43, ..., 232` (a cada 21 linhas), tabela em `A(t+1)`, gráfico nascendo 2 colunas à direita da tabela (como hoje), larguras de coluna `A = 18`, `B:E = 14`:

| # | Linha do título | Título (coluna A) | Tabela | Gráfico |
|---|---|---|---|---|
| 7 | 1 | `7) Barras empilhadas com totais e seta de diferença` | Regiao/Loja/Online (Sul 10 5; Sudeste 25 10; Norte 40 20) | `barstacked` + `Ann_AddDifference 1, 3` |
| 8 | 22 | `8) Cascata com várias séries (segmentos por série)` | Etapa/Loja/Online: Inicio 100 50; Ajuste -30 10; Subtotal e e; Queda -200 -20; Fim e | `waterfall` (2 séries) |
| 9 | 43 | `9) Cascata horizontal` | Etapa/Valor: Receita 1200; Custos -700; Despesas -250; EBITDA e; Impostos -80; Lucro e | `barwaterfall` |
| 10 | 64 | `10) Mekko (%): largura = total da linha` | Regiao/Eletro/Moveis: Sul 40 10; Sudeste 20 10; Norte 10 10 | `mekko` |
| 11 | 85 | `11) Mekko (unidades) com linha de valor` | Regiao/Largura/Eletro/Moveis: Sul 50 40 10; Sudeste 30 20 10; Norte 20 10 10 | `mekkounits` + `Deco_SetValueLine True` |
| 12 | 106 | `12) Gantt com marco e colchete` | Atividade/Início/Fim: Planejamento 02/03/2026–13/03/2026; Execução 09/03/2026–17/04/2026; Entrega 20/04/2026 (sem fim) | `gantt` + `Ann_AddBracket 1, 2` |
| 13 | 127 | `13) Rosca (rótulos sobre o anel)` | Produto/Vendas: Eletro 50; Móveis 30; Tintas 20 | `doughnut` |
| 14 | 148 | `14) Pizza com rótulos fora (nomes e valores)` | mesma tabela | `pie` + `Pie_SetOutside True` (a rosca não aceita rótulos fora: Tarefa 8) |
| 15 | 169 | `15) Área empilhada` | Ano/Sul/Sudeste/Nordeste (2022..2025 como no exemplo 1) | `area` |
| 16 | 190 | `16) Combinação: colunas + linha no eixo secundário` | Ano/Vendas/Custo/Margem: 2022 100 60 0,42; 2023 120 70 0,45; 2024 150 80 0,47 | `combo` |
| 17 | 211 | `17) Colunas empilhadas com linha de valor (média)` | tabela do exemplo 1 | `stacked` + `Deco_SetValueLine True` |
| 18 | 232 | `18) Mesma escala: duas regiões lado a lado` | duas tabelas (A233:C236 e A239:C242: Ano/Loja/Online com valores 10..60 e 100..600) | dois `clustered` redimensionados para 300x200, lado a lado (`co2.Left = co1.Left + 320`, mesmo `Top`), depois `Deco_SameScale` |

  Ao final, `Samples_Create` ativa "Exemplos" com `A1` selecionada (comportamento atual preservado). Falha em qualquer exemplo fecha a pasta e propaga (como hoje).

- [ ] **Passo 1: Testes (RED)** — em `tests/vba/core/TestSamples.bas`. Para a suíte não ficar lenta (cada `Samples_Create` passa a montar 19 gráficos), a pasta de exemplos é criada **uma vez** para todas as verificações de conteúdo: transforme `Test_Samples_Create`, `Test_Samples_Layout_NoOverlap` e `Test_Samples_Content`, da Fase 1, em `Private Sub CheckSheet1Create(wb As Workbook)`, `Private Sub CheckSheet1Layout(wb As Workbook)` e `Private Sub CheckSheet1Content(wb As Workbook)` (corpo inalterado, sem o `Samples_Create` e o `CloseTemp` de cada um) e acrescente:

```vb
Private Function TitleRows2() As Variant
    Dim r(0 To 11) As Long, k As Long
    For k = 0 To 11: r(k) = 1 + 21 * k: Next
    TitleRows2 = r
End Function

' Uma unica pasta de exemplos para todas as verificacoes de conteudo (Samples_Create monta 19 graficos).
Public Sub Test_Samples_BothSheets()
    Dim wb As Workbook, en As Long, es As String, ed As String
    Set wb = Samples_Create()
    On Error GoTo Cleanup
    AssertEqual "Exemplos", wb.ActiveSheet.Name, "a primeira planilha continua ativa"
    CheckSheet1Create wb
    CheckSheet1Layout wb
    CheckSheet1Content wb
    CheckSheet2KindsAndFeatures wb
    CheckSheet2Layout wb
    CloseTemp wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

Private Sub CheckSheet2KindsAndFeatures(wb As Workbook)
    Dim ws As Worksheet, co As ChartObject, kinds As String, c1 As Chart, c2 As Chart
    Set ws = wb.Worksheets("Exemplos 2")
    AssertEqual 13, ws.ChartObjects.Count, "12 secoes, a ultima com 2 graficos"
    For Each co In ws.ChartObjects
        AssertTrue Registry_IsThinQQChart(co.Chart), co.Name
        kinds = kinds & Registry_ChartType(co.Chart) & ";"
    Next
    AssertTrue InStr(kinds, "barstacked;") > 0 And InStr(kinds, "barwaterfall;") > 0 And InStr(kinds, "mekko;") > 0 And InStr(kinds, "mekkounits;") > 0, kinds
    AssertTrue InStr(kinds, "gantt;") > 0 And InStr(kinds, "pie;") > 0 And InStr(kinds, "doughnut;") > 0 And InStr(kinds, "area;") > 0 And InStr(kinds, "combo;") > 0, kinds
    AssertTrue Not ws.ChartObjects(1).Chart.Shapes("tqq_ann_diff_1") Is Nothing, "7: seta de diferenca"
    AssertEqual 2, Waterfall_SeriesCount(ws.ChartObjects(2).Chart), "8: cascata com 2 series"
    AssertTrue Not Chart_FindSeries(ws.ChartObjects(5).Chart, "tqq_vline") Is Nothing, "11: linha de valor"
    AssertTrue Not ws.ChartObjects(6).Chart.Shapes("tqq_ann_bracket_1") Is Nothing, "12: colchete"
    AssertEqual "doughnut", Registry_ChartType(ws.ChartObjects(7).Chart), "13: rosca"
    AssertEqual "pie", Registry_ChartType(ws.ChartObjects(8).Chart), "14: pizza"
    AssertEqual "1", Registry_GetOption(ws.ChartObjects(8).Chart, "pieout"), "14: rotulos fora"
    AssertEqual "1", Registry_GetOption(ws.ChartObjects(10).Chart, "secondary"), "16: linha no eixo secundario"
    AssertTrue Not Chart_FindSeries(ws.ChartObjects(11).Chart, "tqq_vline") Is Nothing, "17: linha de valor"
    Set c1 = ws.ChartObjects(12).Chart: Set c2 = ws.ChartObjects(13).Chart
    AssertEqual c1.Axes(xlValue).MaximumScale, c2.Axes(xlValue).MaximumScale, "18: mesma escala"
    AssertTrue Len(Registry_GetOption(c1, "vmax")) > 0
    AssertEqual ws.ChartObjects(12).Top, ws.ChartObjects(13).Top
End Sub

Private Sub CheckSheet2Layout(wb As Workbook)
    Dim ws As Worksheet, i As Long, j As Long, tr As Variant, a As ChartObject, b As ChartObject, sec As Long
    Set ws = wb.Worksheets("Exemplos 2")
    tr = TitleRows2()
    For i = 1 To ws.ChartObjects.Count
        Set a = ws.ChartObjects(i)
        sec = IIf(i <= 12, i, 12)
        AssertTrue a.Top >= ws.Cells(tr(sec - 1), 1).Top, "grafico " & i & " acima do titulo da secao " & sec
        If sec < 12 Then AssertTrue a.Top + a.Height <= ws.Cells(tr(sec), 1).Top, "grafico " & i & " invade a secao seguinte"
        For j = i + 1 To ws.ChartObjects.Count
            Set b = ws.ChartObjects(j)
            AssertFalse a.Left < b.Left + b.Width And b.Left < a.Left + a.Width And a.Top < b.Top + b.Height And b.Top < a.Top + a.Height, "graficos " & i & " e " & j & " sobrepostos"
        Next
    Next
    For i = 0 To 11
        AssertTrue Len(CStr(ws.Cells(tr(i), 1).Value)) > 0, "titulo " & (i + 7)
    Next
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "TestSamples"` → `not ok` em `Test_Samples_BothSheets` (planilha "Exemplos 2" inexistente).

- [ ] **Passo 3: Implementar** — reutilize `WriteTable`/`WriteTitle`; datas do Gantt com `DateSerial` e `NumberFormat = "dd/mm/yyyy"` nas células; percentuais da combinação como números (0,42) com `NumberFormat = "0%"`; para o exemplo 18 insira os dois gráficos, redimensione e reposicione **antes** de `Deco_SameScale` (a escala automática depende do tamanho); exporte `cht.Export` dos 13 gráficos para PNG numa pasta temporária e olhe cada um (o relatório da tarefa lista o que foi visto e corrigido).

- [ ] **Passo 4: GREEN e suíte completa** (as verificações da Fase 1, agora em `CheckSheet1*`, continuam verdes; a de conteúdo já ignora `tqq_ann_ind100*` desde a Tarefa 7; `Test_Samples_Failure_ClosesWorkbook_NoSideEffects` e `Test_Core_InsertSamples_NeverRaises` ficam como estão: a suíte monta a pasta completa só duas vezes). Registre no relatório o tempo dos testes de exemplos.

- [ ] **Passo 5: Commit** — `git commit -m "feat: planilha Exemplos 2 com os tipos e decoracoes da Fase 2"` (+ trailer).

**Verificação no Excel real:** ThinQQ > Exemplos abre a pasta com duas planilhas; percorrer "Exemplos 2" e conferir cada gráfico contra o título.

---

### Tarefa 14: Tutorial completo, README, CHANGELOG, versões 1.1.0 e ensaio de release

**Files:**
- Modify: `docs/tutorial/tutorial.html`, `README.md`, `CHANGELOG.md`, `version.json`, `release/notes.txt`, `installer/LEIA-ME.txt` (exemplo de pasta `versions\1.1.0`), `tests/ps/Test-Docs.ps1`
- Test: `tests/ps/Test-Docs.ps1` (via `run-tests.ps1`), ensaio `build/release.ps1` na cópia

**Interfaces:**
- Consumes: todos os rótulos do `customUI14.xml` (Tarefa 12), mensagens exatas das tarefas anteriores (recusas 7001/7002, validações 6007..6010, avisos de Mesma escala), comportamentos registrados nos relatórios (`.superpowers/sdd/.../task-N-report.md`): caminho escolhido na sonda das barras (Tarefa 3), direção das barras de erro (Tarefa 10), escala do eixo do Mekko (Tarefa 9).
- Produces: documentação coerente com o código, versões `1.1.0`, `Test-Docs` ampliado, dry run de release verde.

- [ ] **Passo 1: Ampliar `Test-Docs.ps1` (RED)** — acrescente:

```powershell
foreach ($id in 'tipos', 'barras', 'cascata', 'linha-area', 'mekko', 'gantt', 'pizza', 'decoracoes') {
    Check ($null -ne $xml.SelectSingleNode("//*[local-name()='section'][@id='$id']")) "tutorial tem a seção #$id"
}
foreach ($msg in 'A seta de diferença não está disponível para gráficos 100%', 'O colchete só está disponível em gráficos Gantt', 'A cascata aceita no máximo 6 séries') {
    Check ($html -match [regex]::Escape($msg)) "tutorial cita a mensagem '$msg'"
}
Check ($cl -match '(?m)^## \[1\.1\.0\] - \d{4}-\d{2}-\d{2}') 'CHANGELOG tem a seção 1.1.0'
Check (((Read-Utf8 (Join-Path $RepoRoot 'version.json')) | ConvertFrom-Json).loader -eq '1.1.0') 'loader 1.1.0'
```
(colocadas depois das definições de `$xml`, `$html` e `$cl`, dentro do bloco `if ($parsed)` as que usam `$xml`; o limite de tamanho do tutorial já é `-le 122880` (120 KB) desde a Tarefa 12; as demais verificações continuam). Rode `pwsh C:\temp\tqq-f2\tests\ps\Test-Docs.ps1 -RepoRoot C:\temp\tqq-f2` → `not ok` nas seções/mensagens/versões.

- [ ] **Passo 2: Tutorial** — conteúdo obrigatório (cada rótulo de controle em `<span class="ui">` com o texto **exato** do XML; `id`s únicos; sem recursos externos; um único `<script>` inline existente):
  - "Neste tutorial": índice com as seções novas.
  - `#tipos` (da Tarefa 12, ampliada): tabela de todos os botões de inserção com o formato da tabela de origem de cada um.
  - `#barras`: Barras empilhadas/agrupadas/100%, primeira categoria no topo, totais no fim das barras, Cascata (barras); nota de limitação conforme o relatório da Tarefa 3 (se o eixo X secundário precisou ser sincronizado, dizer que sem o ThinQQ os rótulos de total podem sair do lugar após grandes mudanças de dados; senão omitir).
  - `#cascata` (existente, ampliada): várias séries (tabela de exemplo Etapa/Loja/Online, `e` em qualquer célula da linha marca a linha como total, valores nessa linha são ignorados, máximo de 6 séries — citar a mensagem), build-up × build-down (`e` na primeira linha + número na última: o primeiro total é calculado para trás; exemplo Receita/Custos/Despesas/Lucro), rótulos por segmento e por total, horizontal.
  - `#linha-area`: Linha, Área empilhada, Área 100% (rótulos no meio da faixa, participações), Combinação (última coluna vira linha; `Eixo secundário (linha)` ligado automaticamente quando as grandezas diferem mais de 5×; o toggle alterna), `Legenda` (Direita, Nenhuma, Nome no fim da linha — só linha e combinação; citar a mensagem de recusa).
  - `#mekko`: as duas variantes com as tabelas de exemplo (10 e 11 dos Exemplos), larguras proporcionais, separadores, rótulos (participação; largura no topo; nomes embaixo), `Indicador 100%`, `Linha de valor (média)` ponderada só em unidades, sem setas (mensagem), eixo de valores livre em unidades.
  - `#gantt`: tabela (Atividade, Início, Fim; fim vazio = marco; fim é o último dia da atividade; 4.ª coluna opcional vira rótulo), escala automática (semanas até 14 semanas; meses até 3 anos; trimestres até 10 anos; anos), separadores, `Colchete (Gantt)`, `Espaço` = espaçamento das linhas, `Formato` não se aplica; **limitação**: o intervalo do eixo só é recalculado pelo ThinQQ (botão `Atualizar` ou evento de cálculo): sem o ThinQQ as barras/marcos/rótulos seguem os dados, mas o intervalo visível fica o último aplicado.
  - `#pizza`: Pizza e Rosca, `Rótulos fora da pizza` (só na pizza: a rosca mantém os rótulos sobre o anel e o botão fica cinza), `Formato` 12%/12,3% mostra participações, negativos/zero/texto não viram fatia, paleta por fatia, sem setas.
  - `#decoracoes`: `Linha de valor (média)` (média dos totais; da primeira à última categoria; vertical em barras), `Indicador 100%` (ligado por padrão nos 100% e no Mekko (%)), `Mesma escala` / `Escala automática` (regras de seleção, mensagens, escala fixa — clicar de novo após grandes mudanças, tipos ignorados), tabela de "o que se aplica a cada tipo" (matriz do design §7 em linguagem de usuário).
  - "Opções do gráfico selecionado": linhas novas (`Legenda`, `Mais opções` e seus quatro toggles, `Setas` com `Seta CAGR`, `Seta diferença`, `Colchete (Gantt)`, `Remover setas`); nota de que os controles cinza não se aplicam ao tipo selecionado.
  - FAQ "Copiei ou movi um gráfico": atualizar (Tarefa 0: a cópia mantém dados e rótulos ligados, inclusive 100% sem totais).
  - "Limitações": remover "Barras, Mekko, Gantt e pizza ficam para a Fase 2"; acrescentar: mesma escala é fixa; linha de valor só na média; Gantt sem fins de semana/responsável/âncoras; Mekko sem setas; rótulos no fim da linha só em linha/combinação; cascata até 6 séries.
  - "Como funciona": séries XY auxiliares nos gráficos horizontais; Mekko como área em eixo de datas; `_ThinQQ` com blocos maiores (5 linhas por coluna do Mekko; 40 linhas de escala no Gantt).
  Renderize em Edge headless (como na Fase 1) em 1200 px e 500 px e corrija quebras.
- [ ] **Passo 3: README, CHANGELOG, versões, notas** — README: "Recursos" (lista completa 1.1.0), "Antes de distribuir" com `v1.1.0`, "Verificações manuais no Excel real" com os itens das tarefas 0, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13 consolidados (numerados, um por linha), "Limitações conhecidas" atualizadas, links para o design e o plano da Fase 2; todo caminho `ThinQQ > Rótulo` com rótulo existente. CHANGELOG: `## [1.1.0] - <data do commit>` com **Adicionado** (barras, cascata horizontal e com várias séries, cascata de cima para baixo, área, combinação, Mekko, Gantt, pizza/rosca, linha de valor, indicador 100%, mesma escala, legenda no fim da linha, faixa em menus), **Alterado** (faixa; controles habilitados por tipo; `e` na primeira linha passa a calcular para trás), **Corrigido** (tudo que estava em "Não lançado" + Tarefa 0); mantenha `## [Não lançado]` vazio no topo. `version.json` → `{ "loader": "1.1.0", "core": "1.1.0" }`. `release/notes.txt` → 6 a 8 linhas com `•` (vai para o manifesto). `LEIA-ME.txt` → exemplo `versions\1.1.0\`.
- [ ] **Passo 4: GREEN** — `pwsh C:\temp\tqq-f2\build\run-tests.ps1 -TimeoutSec 900` → `Todos os testes passaram` (Test-Docs inclusive).
- [ ] **Passo 5: Ensaio de release (na cópia, nunca no repositório real)**

```powershell
robocopy "<repo>" C:\temp\tqq-f2 /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS
pwsh C:\temp\tqq-f2\build\release.ps1 -Version 1.1.0 -LoaderVersion 1.1.0 -BaseUrl C:\temp\tqq-rel -PublishTo C:\temp\tqq-rel
# esperado: "Release pronto: ...\ThinQQ-1.1.0.zip"; manifest.json com coreVersion/loaderVersion 1.1.0 e dois SHA-256
Expand-Archive C:\temp\tqq-rel\ThinQQ-1.1.0.zip C:\temp\tqq-rel-x
pwsh C:\temp\tqq-rel-x\install.ps1 -Root C:\temp\tqq-rel-root\ThinQQ -RegistryRoot HKCU:\Software\ThinQQ_Test_f2 -SkipExcelCheck
# esperado: "ThinQQ 1.1.0 instalado"; versions\1.1.0 com ThinQQ.core.xlam e tutorial.html
pwsh C:\temp\tqq-rel-x\uninstall.ps1 -Root C:\temp\tqq-rel-root\ThinQQ -RegistryRoot HKCU:\Software\ThinQQ_Test_f2 -SkipExcelCheck
# -SkipExcelCheck é obrigatório aqui: o Excel do usuário está aberto e, sem ele, o script lança 'Feche o Excel antes de desinstalar'
Remove-Item HKCU:\Software\ThinQQ_Test_f2 -Recurse -ErrorAction SilentlyContinue      # chave de teste propria
Remove-Item C:\temp\tqq-rel, C:\temp\tqq-rel-x, C:\temp\tqq-rel-root -Recurse -Force
```
  Nada do ensaio é commitado (`release/manifest.json` só existe após a publicação real, como o `release/README.md` explica).
- [ ] **Passo 6: Commit** — `git commit -m "docs: tutorial, README e CHANGELOG da Fase 2; versao 1.1.0 do core e do loader"` (+ trailer).

**Verificação no Excel real:** ThinQQ > Tutorial abre o tutorial novo; ThinQQ > Sobre mostra Loader 1.1.0 / Core 1.1.0 após instalação pelo ZIP (quando o usuário decidir instalar).

---

### Tarefa 15: Integração final e verificação

**Files:**
- Nenhum arquivo de produção novo. Possíveis correções pequenas achadas aqui entram em commits `fix:` próprios, cada um com teste.

**Interfaces:** consome tudo.

- [ ] **Passo 1: Cópia limpa e suíte completa com troca de loader**

```powershell
robocopy "<repo>" C:\temp\tqq-f2 /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS
pwsh C:\temp\tqq-f2\build\run-tests.ps1 -LoaderSwap -TimeoutSec 900
```
  Esperado: `Todos os testes passaram`, `loader-swap: ok`; anote no relatório os totais (`# passed N failed 0` do core e do loader), o tempo e a saída 0. Se o tempo total passar de 6 minutos, registre quais testes são os mais lentos (TAP com `Timer` por teste pode ser adicionado em `modTestRunner` como `fix:` de infraestrutura) — não remova testes.
- [ ] **Passo 2: Build Release e guarda** — `pwsh C:\temp\tqq-f2\build\build.ps1 -Target all -Configuration Release` → `OK dist\ThinQQ.xlam`, `OK dist\ThinQQ.core.xlam`; `pwsh C:\temp\tqq-f2\tests\ps\Test-ReleaseGuard.ps1` → sem módulos de teste no Release.
- [ ] **Passo 3: Compatibilidade entre versões (sem automação contra o core velho)** — loader 1.1.0 com core 1.0.0 é coberto por testes que já rodam no Passo 1: `TestLoaderMain.Test_RunCore_MissingProc_ReturnsEmpty_NoRaise` (`Loader_RunCore("Core_NaoExiste")` devolve `Empty` sem erro, com a raiz redirecionada) e o fallback de `Rx_GetEnabled` para `HasChartCached()` quando `Core_Supports` devolve `Empty` (Tarefa 12); confirme que estão verdes e registre. **Não** abra um core 1.0.0 Release por automação para chamar `Core_*` por `Loader_RunCore`: (a) qualquer `Core_*` que avise o usuário (ex.: `Core_Insert "pie"` → "Tipo de gráfico desconhecido") passa por `Core_Notify`, que em Release mostra `MsgBox` e trava a automação; (b) no loader Release, o caminho de erro de `Loader_RunCore` grava `Loader_Log` em `%LOCALAPPDATA%\ThinQQ\logs` (e cria as pastas), a pasta real do usuário, proibida por este plano (`Loader_SetRootOverride` é inerte em Release). Os callbacks `Rx_*` não podem ser chamados por automação (exigem `IRibbonControl`); a faixa com core antigo fica para o roteiro manual do usuário.
- [ ] **Passo 4: Auditoria de aderência às regras do plano** — `grep` nos fontes por: `Declare ` (nenhum), `NumberFormat =` fora de células de data (`Range.NumberFormat` só com `dd/mm`, `mmm/yy`, `General`, `0%`, `dd/mm/yyyy`), `MsgBox` fora de `Core_Notify`/`RxMsg`/`Loader_Msg`/formulários, `Application.Run` nos testes (nenhum), handlers que dependem de `On Error Resume Next` para limpar recursos num `Core_*` (a Fase 1 descobriu, em `Core_ApplyPendingLoader`, que dentro de um handler ativo ele não protege: um erro novo escapa ao chamador; por isso a restauração fica em `RestoreLoaderFiles`, procedimento separado). A regra **não** acusa os handlers que desfazem e repropagam de propósito — os `Fail:` dos `*_Insert` (`On Error Resume Next` → apagar gráfico/bloco → `On Error GoTo 0` → `Err.Raise en, es, ed`, como `Stacked_Insert`/`Waterfall_Insert` da Fase 1) e o `FailNew:` de `Core_RebuildFromRange` (termina em `Errors_HandleErr`): depois de `On Error GoTo 0` dentro de um handler só podem vir `Err.Raise`, `Errors_Handle`/`Errors_HandleErr`, `Core_Notify`, `Errors_Log`, atribuições e `Exit`; liste no relatório qualquer outro caso. Confirme que todo `Public Sub`/`Public Function Core_*` tem `On Error` antes da primeira linha executável (estenda `Test_Ribbon_AllCallbacks_HaveErrorContainment` com um equivalente no core, `Test_Core_AllPublicApi_HaveErrorContainment`, lendo `modCoreApi` via `VBProject.VBComponents("modCoreApi").CodeModule` — exceções permitidas: getters puros `Core_Ping`, `Core_Version`, `Core_LoaderPath`, `Core_LastNotice`, `Core_NoticeCount`, `Core_ClearLastNotice`, `Core_TestSetApplyBusy`, e `Core_ApplyPendingLoader`, cuja guarda de reentrada `If mApplyBusy Then Exit Sub` vem antes do `On Error` de propósito e cujo corpo roda em `ApplyPendingLoaderImpl`, com `On Error GoTo Fail`).
- [ ] **Passo 5: Estado do repositório** — `git status` limpo; `git log --oneline feat/fase1..feat/fase2` lista os commits da Fase 2 (um por tarefa, mais `fix:`); nenhum push; `dist/` e `tests/results/` continuam ignorados; `release/manifest.json` inexistente.
- [ ] **Passo 6: Relatório final** — consolidar: contagens da suíte, decisões tomadas nas sondas (barras, Gantt, Mekko), lista de "minor (deferred)" de todos os relatórios de tarefa, pastas temporárias removidas (`C:\temp\tqq-f2`, `tqq-rel*`), confirmação de que os Excel do usuário (os PIDs anotados no início) continuam vivos e intocados (`Get-Process excel | Select-Object Id, StartTime`), e a lista numerada das verificações manuais no Excel real que ficam para o usuário (a mesma do README).

**Verificação no Excel real (usuário, após instalar pelo ZIP):** percorrer a lista "Verificações manuais no Excel real" do README na ordem; qualquer desvio volta como `fix:` com teste.

---

## Autorrevisão (feita ao escrever o plano)

- **Cobertura da spec/adendo:** Tarefa 0 (carry-over C1 residual) ✓ (concluída: `b19615e`, `6b036c3`); barras ✓ (3); cascata horizontal ✓ (5); várias séries ✓ (4); build-down ✓ (4); Mekko % e unidades ✓ (9); Gantt com atividades, barras, marcos, escala e colchetes ✓ (10); pizza/rosca dentro/fora ✓ (8); área ✓ e combinação com eixo secundário ✓ (6); linhas de valor na média ✓, indicador 100% ✓, mesma escala ✓ (7); rótulos no fim da linha ✓ (6); faixa em menus ✓ (12); exemplos ✓ (13); tutorial/README/CHANGELOG/versão/release ✓ (14); integração ✓ (15). Ícones próprios na faixa: adiado (design §1/§10).
- **Placeholders:** nenhum "TBD/TODO"; todo passo de código traz o código ou a fórmula; cada mensagem de usuário aparece com o texto exato onde é testada.
- **Consistência de nomes:** `Chart_AddHelperSeries(cht, name, valueCells, posCells)` (T1) usada em T3/T4/T5/T7/T9/T10; `Stacked_PosCol/VLineCol` (T3) usadas em T7; `Area_*Col` (T6) em T7; `Mekko_*Col`/`Mekko_CenterRow` (T9) em T9/T13; `GT_COL_*` (T10); `Kinds_Supports/Family/Is100/IsHorizontal/FeatureLabel` (T1/T11); `Deco_SetValueLine/SetInd100/SetLegend/SameScale/ResetScale/ApplyFixedScale` (T6/T7); `Ann_AddInd100/Ann_RemoveInd100` (T7); `Mekko_ApplyColors` (T9); `Errors_LogDir` (T1); `Waterfall_Insert(ws, data, kind)` (T5) e `Waterfall_SeriesCol/SeriesCount` (T4); `Geom_ChartCatCoord/ValCoord/ValueToX/CategoryCenterY` (T2); erros `7001` (anotação), `7002` (decoração/opção), `7200` (tipo não pronto; o teste temporário com `gantt` sai na T10), `6007..6010` (validação de dados).
- **Varredura de coerência (pre-flight, antes da Tarefa 1):** aplicada sobre este plano. Principais correções: teste temporário de tipo não pronto com `gantt` (sai na T10); `pieout` só na pizza (rosca recusa `InsideEnd`/`OutsideEnd`; exemplo 14 = pizza); totais 1.250 com o separador da localidade; `body.Rows.Count` (o bloco tem a linha de meta) nos testes de Mekko e Gantt; `legend=right` gravado por `Stacked_Insert`/`Combo_Insert` e `ApplyOption "legend"` gravando só quando `Deco_SetLegend` aceita; `IsTotal` sem `IFERROR` e `BD` sobre as células de origem (sem circularidade); folgas para setas/totais/escala; dispatch de `labels`/`palette` por família na tarefa que o testa; colchete com `de = até`; `Test-Docs` com `local-name()`; `menu` sem filhos com `size`; `-SkipExcelCheck` no ensaio; compatibilidade com core 1.0.0 sem automação; log do build Test fora da pasta do usuário.
- **Review Focus → testes:** 1 → `Test_Bars_TotalsHelper_StaysAlignedAfterDataChangeAndResize` (T3); 2 → `Test_WaterfallMulti_Block_MixedSignsCrossingZero` (T4); 3 → `Test_Mekko_ZeroTotalRow_NoErrorNoWidth` (T9); 4 → `Test_Gantt_InvalidDates_RefusedWithMessage`, `Test_Gantt_ScaleUnitSwitches_OnRefresh`, `Test_Gantt_OnlyMilestones_And_SingleActivity` (T10); 5 → `Test_Pie_NegativeZeroText_AndSevenSlices` (T8).

