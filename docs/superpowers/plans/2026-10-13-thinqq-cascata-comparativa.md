# ThinQQ — Plano de Implementação: Cascata comparativa (1.3.0)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Entregar o tipo de gráfico **Cascata comparativa** (`waterfallpair`) do ThinQQ 1.3.0: duas pontes lado a lado em cada etapa, conectores por série, setas em cotovelo com a variação entre as duas barras de cada etapa numa elipse, totais com "valor (percentual do primeiro total)", nomes das séries dentro da primeira barra (ou legenda), tudo em séries nativas ligadas ao bloco da `_ThinQQ` (vivo sem o ThinQQ), com faixa, ícone, Exemplos, pasta de demonstração (E27, E28), galeria, tutorial, README e CHANGELOG.

**Architecture:** A das Fases 1 a 3: `ThinQQ.xlam` (loader: faixa) chama `Core_*` em `ThinQQ.core.xlam`; cada gráfico é nativo e ligado por fórmulas a um bloco da planilha muito oculta `_ThinQQ`. Novidade: família própria `waterfallpair` (módulo `modChartWaterfallPair.bas`, prefixo `Wfp_`) com colunas empilhadas em fatias expandidas (A, B, vão), eixo de categorias de 2 níveis e auxiliares **XY no grupo principal** (x = número da fatia) para conectores (barras de erro), setas (linhas com ponta e rótulo em elipse), sala das setas e nomes; o ThinQQ só recalibra no refresh a meia-largura das barras (`HW`), o valor por ponto das setas (`PtV`), os rótulos finos (`LblOn`) e "nome dentro da barra ou legenda" (`Fit`).

**Tech Stack:** VBA 7 (Excel 2016+ Windows, 32/64 bits, sem `Declare`, sem referências), customUI14, PowerShell 7 e 5.1 (build, testes, ferramentas), Git.

**Spec:** `docs/superpowers/specs/2026-10-13-thinqq-cascata-comparativa-design.md` (cópia, feita na Tarefa 0, do rascunho `.superpowers/sdd/2026-10-13-cascata-comparativa/design.md`: pesquisa, sondas medidas, decisões, matriz §6, pré-flight §11) sobre as specs das Fases 1, 2 e 3 em `docs/superpowers/specs/`. Ledger e relatórios da Fase 3 (armadilhas medidas): `.superpowers/sdd/2026-10-11-thinqq-fase3/`. Regras de ambiente: `.superpowers/sdd/2026-10-13-cascata-comparativa/implementer-rules.md`.

**Estado de execução:** nenhuma tarefa concluída. Branch: `feat/cascata-comparativa` (criada da correção 1.2.3 pelo controlador); o implementador **não** cria nem troca de branch. Se a 1.2.3 for publicada antes do merge, o **controlador** rebaseia esta branch na `main`: no CHANGELOG tira o `### Corrigido` da 1.2.3 do `[Não lançado]` (ele fica na `[1.2.3]`); `release/notes.txt` segue a variante da Tarefa 8, Passo 4, escolhida pela existência da tag `v1.2.3`. O implementador não decide pela publicação da 1.2.3.

## Global Constraints

- **Ambiente:** Excel 2019 32-bit pt-BR no Windows 11; testes em Excel invisível próprio. O VBA compila sob demanda: o build roda o lint (`build/lib/Vba-Lint.ps1`) e uma compilação completa; erro de compilação ⇒ saída 3 (módulo:linha na tela); travamento morto pelo watchdog ⇒ saída 2; asserts ⇒ 1; ok ⇒ 0. Fontes VBA em UTF-8 com **somente caracteres cp1252** (acentos do português sim; nada de `→`, `≥`, `−`, `…`, aspas tipográficas, emojis), `Option Explicit`, declarações antes dos procedimentos, `Attribute VB_Name` na primeira linha, nenhuma continuação ` _` seguida de texto, nada de `ElseIf` num `If` de uma linha.
- **Testes:** `Public Sub Test_x()` sem parâmetros em `tests/vba/{core,loader}/Test*.bas`, despachados por módulo gerado (nunca `Application.Run`), via `pwsh build/run-tests.ps1` (parâmetros `-Target core|loader|all`, `-NoBuild`, `-TimeoutSec`, `-Filter "Modulo.Teste|outro"` — pula as suítes PowerShell —, `-LoaderSwap`). Asserts: `AssertTrue/False/Equal/Almost`, `NewTempWorkbook`, `CloseTemp`, `TempDir`, `ForceLayout` (exporta o gráfico: só depois dele a geometria de rótulos e pontos é a desenhada), `RepoPath`. Cada módulo de teste novo define os próprios `Fixture`, `Teardown` (`Watchers_Stop` + `CloseTemp`) e `TqqNameCount`.
- **Instâncias do usuário:** os processos Excel/PowerPoint já abertos quando a tarefa começa (anote `Get-Process excel,powerpnt -ErrorAction SilentlyContinue | Select-Object Id, StartTime`) nunca são tocados nem encerrados; um deles pode travar `dist\ThinQQ.core.test.xlam` no repositório real. Por isso **builds e testes rodam numa CÓPIA**: `robocopy "<repo>" C:\temp\tqq-cc /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS` (repita antes de cada rodada; edite sempre o repositório real; commite do repositório real; `cd` só com caminho absoluto). Nunca `build.ps1 -Install`; nunca tocar em `%LOCALAPPDATA%\ThinQQ`, no registro nem nas lojas de certificados além do que os testes existentes já fazem. **Disco C: com pouco espaço:** apague `C:\temp\tqq-cc` e `C:\temp\tqq-cc-*` ao terminar; nunca apague nada fora delas.
- **Código:** `Core_*` nunca propaga erro (`Errors_Handle`/`Errors_HandleErr`); `MsgBox` só em `Core_Notify`/`Core_Confirm` (log-only no build Test), `RxMsg`, `Loader_Msg` e formulários; fórmulas em inglês via `.Formula`; formatos de **rótulo** via `Style_AssignNumberFormat`/`NumberFormatLocal` (nunca `DataLabel.NumberFormat`), formatos de **célula** do bloco por `NumberFormatLocal` com `Style_LocalizeCode` (números) e `Range.NumberFormat` em inglês só para datas, texto (`@`) e o formato copiado da origem; testes de texto renderizado via `DataLabel.Text` com expectativas derivadas de `Format$`/`Application.International`; textos de interface em pt-BR; identificadores em inglês com o prefixo novo `Wfp_`.
- **Todo tipo novo entrega:** `Insert` com rollback (`Util_TestHookAfterBlock` depois do bloco, `Util_TestHookAfterChart` depois do gráfico), opções por `Registry_SetOption` (inclusive `n`), rótulos ligados a células, recusas pela matriz (7001 anotações, 7002 opções), testes de geometria **independentes das fórmulas de produção** (regra de três sobre `PlotArea.Inside*` e `Axes(xlValue).MinimumScale/MaximumScale` reais), exemplo em `Samples_Create`, seção no tutorial, entrada na faixa validada por `Test-Ribbon`/`TestRibbon`, cobertura em `Test-Docs` e na demonstração.
- **Armadilhas medidas (Fases 2 e 3 e as sondas desta spec, §3; não redescubra):** rótulos e `Format$` seguem a localidade; um ponto `NA()` só some (com o rótulo) com `DisplayBlanksAs = xlNotPlotted` — este tipo usa `xlNotPlotted` e o texto ligado dos pontos `NA()` é `""`, nunca `#N/D`; a geometria de um `DataLabel`/`Point` só vale depois de `ForceLayout`; `Chart.Export` grava PNG de 0 byte para gráfico fora da janela do Excel invisível (exporte gráficos perto de A1 ou role até eles); `Series.Name` devolve o texto, não a referência (procure séries de dados pelo índice ou por `Wfp_SideSeries`); `Chart.ChartType` lê `xlCombination` (−4111) quando há auxiliares XY: confira `SeriesCollection(k).ChartType`; um texto gravado como **constante** "+7%" vira o número 0,07 (os textos do bloco são **fórmulas**); com `XValues` de 2 colunas e a interna vazia o Excel junta tudo num nível (a interna é `" "`); a linha do eixo de categorias de 2 níveis desenha separadores (fica invisível); pontos XY do grupo principal entram na escala automática do eixo de valores e usam x = número da categoria; gravar uma célula que o gráfico lê dispara o evento Calculate, que reentra no refresh (guarda `mBusy` e gravações só quando o valor muda); laços de geometria têm teto de passadas e tolerância; ganchos de teste (`Util_*ForTest`, `Util_FailThinSyncForTest`) são zerados em todo caminho, inclusive de erro (`On Error GoTo Cleanup` que reseta e repropaga); `IIf` avalia os dois ramos (não use com índices inválidos); PowerShell: nomes de variável não diferenciam maiúsculas (`$R` = `$r`), a vírgula tem precedência sobre a aritmética dentro de `@(...)` (use parênteses), uma função devolve no pipeline tudo o que não for atribuído (`[void]`/`| Out-Null`), e `Range.NumberFormat` gravado pelo PowerShell num Windows pt-BR é lido como código local (use o `Set-NumFmt` do gerador da demonstração).
- **Commits:** pequenos, em português, prefixos `feat:`/`fix:`/`test:`/`docs:`/`build:`, terminando com a linha `Co-Authored-By:` indicada no prompt do implementador ("(+ trailer)" nos comandos abaixo). **Não** fazer push, **não** criar nem trocar de branch.
- **Versões:** `version.json` não muda (o `publicar-release.bat 1.3.0` sobe loader e core). O build importa todo `.bas` de `src/<alvo>/modules` e todo `Test*.bas` de `tests/vba/<alvo>`: um módulo novo não precisa ser registrado.
- **Processo:** cada tarefa é executada por um implementador novo que só vê o texto da tarefa; tudo o que ele precisa está nela. Cada tarefa termina com a suíte completa verde na cópia (`pwsh C:\temp\tqq-cc\build\run-tests.ps1 -TimeoutSec 900`) e um commit (ou mais). Modelo sugerido: **standard** (código com testes), **most-capable** (família nova, geometria, setas, demonstração), **cheap** (edições mecânicas).

## Review Focus

1. **Etapa que cruza o zero e total negativo** (série A: 10, −25, +8, total −7; série B: 5, −12, +20, total 13): o usuário espera cada barra na extensão certa dos dois lados do zero, a linha do zero visível, o rótulo do total negativo **abaixo** da barra ("−7 (−70%)") e a seta saindo de cima das barras. Testes: `Test_Wfp_CrossesZero_NegativeTotal` (T1) e `Test_Wfp_Labels_NegativeTotalBelow` (T4).
2. **Percentual da seta igual ao da imagem do think-cell**: −4 → −7 é "+75%" (o custo cresceu), −6 → −2 é "−67%", 27 → 29 "+7%", 11 → 10 "−9%"; com o primeiro valor zero ou sinais opostos, "n/d"; com Variação em valor, "+3". Testes: `Test_Wfp_Arrows_TextsLikeTheImage`, `Test_Wfp_Arrows_NdAndAbs` (T3).
3. **Gráfico redimensionado ou em Mesma escala**: a seta continua 3 pt acima do degrau e 22 pt acima da barra de total com rótulo (acima do próprio rótulo), a elipse inteira dentro do gráfico. Testes: `Test_Wfp_Arrows_Geometry`, `Test_Wfp_Arrows_ResizeRecalibrates`, `Test_Wfp_Arrows_SameScale_FixedAxis` (T3), `Test_Wfp_Labels_StepsCenterTotalsAbove` (T4).
4. **Nome de série que não cabe na primeira barra** (barra baixa): o usuário espera a legenda à direita em vez de um nome transbordando; voltando a caber, os nomes voltam. Teste: `Test_Wfp_NamesInsideBars_And_LegendModes` (T1).
5. **Sem o ThinQQ**: mudar um número da tabela move barras, conectores, setas, o texto da elipse e o "(41%)" do total. Testes: `Test_Wfp_DataChange_LiveWithoutThinQQ` (T1), `Test_Wfp_Connectors_GeometryAndLive` (T2), `Test_Wfp_Arrows_LiveWithoutThinQQ` (T3).

---

## Estrutura de arquivos (novos e alterados)

```
docs/superpowers/specs/2026-10-13-thinqq-cascata-comparativa-design.md   NOVO (T0: cópia do design)
docs/superpowers/plans/2026-10-13-thinqq-cascata-comparativa.md          NOVO (T0: cópia deste plano)
src/core/modules/
├── modChartWaterfallPair.bas   NOVO  Wfp_* (T1 construção; T2 conectores; T3 setas; T4 rótulos, totais, formato)
├── modChartKinds.bas           ALT   tipo, família, matriz 21 x 23, rótulos dos recursos, Kinds_ExpectedCategoryRefs (T0); Kinds_Insert (T1)
├── modLabels.bas               ALT   saída cedo da família (T0); Labels_ApplyCellFormat -> Wfp_ApplyCellFormat (T4)
├── modChartStyle.bas           ALT   saída cedo da família (T0); Style_ApplyNumberFormat -> Wfp_ApplyNumberFormat (T4)
├── modAnnotations.bas          ALT   Ann_GetCategoryData (T1)
├── modDecorations.bas          ALT   Deco_SetLegend: modo inbar e família (T1)
├── modChartUtil.bas            ALT   Chart_EnsureRoomAfterChange -> Wfp_RefreshAfterDataChange (T3)
├── modCoreApi.bas              ALT   despachos, legenda inbar, refresh, mismatch (T1); conectores (T2); setas e ReapplyOptions (T3); totais e rótulos (T4)
└── modSamples.bas              ALT   exemplo 27 em "Exemplos 3" (T6)
src/loader/ribbon/customUI14.xml  ALT (T5);  build/lib/Ribbon-Icons.ps1 ALT (ícone tqq_waterfallpair, T5)
tests/vba/core/TestWaterfallPair.bas  NOVO (T1..T4);  TestKinds.bas, TestCoreApi.bas (T0, T1, T6), TestSamples.bas (T6)  ALT
tests/vba/loader/TestRibbon.bas  ALT (T5);  tests/ps/Test-Docs.ps1 ALT (T0, T5, T8);  tests/ps/Test-DemoWorkbook.ps1 ALT (T0, T7)
tools/DemoCoverage.ps1, tools/New-DemoWorkbook.ps1  ALT (T7);  exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx  REGERADO (T7)
tools/Export-DemoChartImages.ps1, tools/Build-TutorialGallery.ps1  ALT só nos comentários (T7)
docs/tutorial/tutorial.html  ALT (T0 matriz; T5 seção e rótulos; T6/T7 contagens e galeria; T8 texto completo)
README.md, CHANGELOG.md, release/notes.txt  ALT (T6/T7 contagens; T8)
```

## Como trabalhar (todas as tarefas)

```powershell
# 1) editar no repositório real; 2) espelhar para a cópia; 3) rodar lá
robocopy "<repo>" C:\temp\tqq-cc /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS
pwsh C:\temp\tqq-cc\build\run-tests.ps1 -Target core -Filter "TestWaterfallPair" -TimeoutSec 400     # iteração (RED/GREEN)
pwsh C:\temp\tqq-cc\build\run-tests.ps1 -TimeoutSec 900                                             # suíte completa antes de commitar
Remove-Item C:\temp\tqq-cc -Recurse -Force                                                          # ao terminar a tarefa
```
`<repo>` é o caminho do repositório real (o diretório de trabalho do implementador). Fixtures: cada módulo de teste tem a sua `Private Function Fixture(...)`; use `Teardown` quando chamar `Core_*` (observadores).

---

### Tarefa 0: Matriz 1.3.0, família nova nos pontos comuns, spec promovida (infraestrutura) — modelo: standard

**Files:**
- Create: `docs/superpowers/specs/2026-10-13-thinqq-cascata-comparativa-design.md` (cópia byte a byte de `.superpowers/sdd/2026-10-13-cascata-comparativa/design.md`; se o controlador já a criou, não sobrescreva)
- Create: `docs/superpowers/plans/2026-10-13-thinqq-cascata-comparativa.md` (cópia byte a byte de `.superpowers/sdd/2026-10-13-cascata-comparativa/plan.md`, já com o pré-flight aplicado; se o controlador já a criou, não sobrescreva)
- Modify: `src/core/modules/modChartKinds.bas`, `src/core/modules/modLabels.bas`, `src/core/modules/modChartStyle.bas`, `docs/tutorial/tutorial.html` (só `#matriz`), `tests/ps/Test-Docs.ps1` (contagens), `tests/ps/Test-DemoWorkbook.ps1` (lista temporária de pendências)
- Test: `tests/vba/core/TestKinds.bas`, `tests/vba/core/TestCoreApi.bas` (teste temporário)

**Interfaces:**
- Produces (`modChartKinds`): `KINDS` com 21 tipos (`...,bubble,waterfallpair,`); `Kinds_Family("waterfallpair") = "waterfallpair"`; `Kinds_Supports` com a matriz do spec §6 (23 recursos: os 20 de antes, `waterfallpair` acrescentado em `totals`, `connectors`, `numfmt`, `gap`, `samescale`, e os novos `pairdiff`, `pairabs`, `inbar` = só `waterfallpair`; `labels`, `palette`, `legend` já o incluem por `allKinds`); `Kinds_FeatureLabel`: `pairdiff` "Setas de variação", `pairabs` "Variação em valor", `inbar` "Nome dentro das barras"; `Kinds_ExpectedCategoryRefs("waterfallpair", n) = 2 * (3n - 1)`; `Kinds_Insert`: a família nova cai no `Case Else` (7200 "ainda não está disponível") até a Tarefa 1.
- Produces (`modLabels`): `Labels_RefreshPositions` sai cedo na família `waterfallpair` (permanente: o refresh da família cuida); `Labels_ApplyCellFormat` sai cedo nela (a Tarefa 4 troca por `Wfp_ApplyCellFormat cht: Exit Sub`).
- Produces (`modChartStyle`): `Style_ApplyNumberFormat` sai cedo na família (a Tarefa 4 troca por `Wfp_ApplyNumberFormat cht: Exit Sub`).
- Produces (tutorial `#matriz`): linha `data-kind="waterfallpair"` nas três tabelas e a 4.ª tabela "Cascata comparativa" (`data-f` `pairdiff`, `pairabs`, `inbar`), rótulos novos em `<span class="k">` (a Tarefa 5 os troca para `ui`).
- Produces (`Test-DemoWorkbook.ps1`): `$pendingDemo` (tipo e recursos ainda sem exemplo na demonstração; a Tarefa 7 remove).

- [ ] **Passo 1: Promover o spec e o plano** — `Copy-Item "<repo>\.superpowers\sdd\2026-10-13-cascata-comparativa\design.md" "<repo>\docs\superpowers\specs\2026-10-13-thinqq-cascata-comparativa-design.md"` (só se o destino não existir). Confira que o arquivo tem `## 6. Matriz de capacidades 1.3.0` e `## 7. Faixa de opções`. Copie também o plano: `Copy-Item "<repo>\.superpowers\sdd\2026-10-13-cascata-comparativa\plan.md" "<repo>\docs\superpowers\plans\2026-10-13-thinqq-cascata-comparativa.md"` (só se o destino não existir; o `Test-Privacy` varre `docs/`: o plano só cita `C:\temp\...` e `<repo>`).

- [ ] **Passo 2: Testes (RED)** — em `tests/vba/core/TestKinds.bas`, troque a função `SpecSection6` inteira por estas duas (a antiga passa a ser um atalho; o teste da Fase 3 continua igual):

```vb
' Texto de uma secao de um spec de docs\superpowers\specs (bytes; so ASCII importa: nomes de recursos e tipos), de startHdr ate endHdr.
Private Function SpecSectionOf(fileName As String, startHdr As String, endHdr As String) As String
    Dim f As Integer, s As String, p As String, a As Long, b As Long
    p = RepoPath() & "\docs\superpowers\specs\" & fileName
    f = FreeFile
    Open p For Binary Access Read As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    a = InStr(s, startHdr)
    If a = 0 Then Err.Raise vbObjectError + 1, "TestKinds", "secao '" & startHdr & "' nao encontrada em " & p
    b = InStr(a + 1, s, endHdr)
    If b = 0 Then Err.Raise vbObjectError + 1, "TestKinds", "secao '" & endHdr & "' nao encontrada em " & p
    SpecSectionOf = Mid$(s, a, b - a)
End Function

' Texto da secao 6 do spec da Fase 3 (a tabela "| feature | tipos |").
Private Function SpecSection6() As String
    SpecSection6 = SpecSectionOf("2026-10-11-thinqq-fase3-design.md", "## 6. ", "## 7. ")
End Function
```

E acrescente ao fim do módulo:

```vb
Public Sub Test_Kinds_PairWaterfall_TypeFamilyAndLabels()
    AssertTrue Kinds_IsValid("waterfallpair")
    AssertEqual "waterfallpair", Kinds_Family("waterfallpair")
    AssertEqual "waterfall", Kinds_Family("waterfall"), "a cascata atual continua na familia dela"
    AssertFalse Kinds_IsHorizontal("waterfallpair"): AssertFalse Kinds_Is100("waterfallpair")
    AssertEqual 28, Kinds_ExpectedCategoryRefs("waterfallpair", 5), "categorias de 2 colunas x (3n - 1) fatias"
    AssertEqual 4, Kinds_ExpectedCategoryRefs("waterfallpair", 1)
    AssertEqual 5, Kinds_ExpectedCategoryRefs("waterfall", 5), "os demais tipos: uma referencia por categoria"
    AssertEqual "Setas de variação", Kinds_FeatureLabel("pairdiff")
    AssertEqual "Variação em valor", Kinds_FeatureLabel("pairabs")
    AssertEqual "Nome dentro das barras", Kinds_FeatureLabel("inbar")
    AssertEqual "'Nome dentro das barras' não se aplica a este tipo de gráfico.", Kinds_NotApplicableText("inbar")
    AssertTrue Kinds_Supports("waterfallpair", "legend"), "a legenda da comparativa nao depende do numero de series"
    AssertFalse Kinds_Supports("waterfall", "pairdiff"): AssertFalse Kinds_Supports("butterfly", "inbar"): AssertFalse Kinds_Supports("stacked", "pairabs")
End Sub

' Cascata comparativa (1.3.0): a matriz inteira (21 tipos x 23 recursos) conferida contra o TEXTO da tabela da secao 6 do spec da cascata
' comparativa (nao contra a propria matriz). O teste da Fase 3 (20 x 20, spec da Fase 3) continua valendo para os tipos e recursos antigos.
Public Sub Test_Kinds_PairWaterfall_Matrix_ExhaustiveFromSpec()
    Dim sec As String, lines() As String, i As Long, cells() As String, feat As String, expect As String, k As Variant, allKinds As Variant
    Dim nFeat As Long, nPairs As Long, want As Boolean
    allKinds = Array("stacked", "clustered", "stacked100", "line", "waterfall", "barstacked", "barclustered", "barstacked100", "barwaterfall", "area", "area100", "combo", "mekko", "mekkounits", "gantt", "pie", "doughnut", "butterfly", "scatter", "bubble", "waterfallpair")
    AssertEqual 21, UBound(allKinds) + 1
    sec = SpecSectionOf("2026-10-13-thinqq-cascata-comparativa-design.md", "## 6. ", "## 7. ")
    lines = Split(Replace(sec, vbCr, ""), vbLf)
    For i = LBound(lines) To UBound(lines)
        If Left$(lines(i), 2) = "| " Then
            cells = Split(lines(i), "|")
            feat = Trim$(Replace(cells(1), "**", ""))
            If feat <> "feature" And Left$(feat, 3) <> "---" Then
                nFeat = nFeat + 1
                expect = SpecKinds(cells(2), allKinds)
                For Each k In allKinds
                    want = (InStr(expect, "," & k & ",") > 0)
                    AssertEqual want, Kinds_Supports(CStr(k), feat, 2), feat & " x " & k & " (spec: " & Trim$(cells(2)) & ")"
                    nPairs = nPairs + 1
                Next
            End If
        End If
    Next
    AssertEqual 23, nFeat, "23 recursos na tabela do spec"
    AssertEqual 483, nPairs
End Sub
```

Em `tests/vba/core/TestCoreApi.bas` (temporário; a Tarefa 1 o remove):

```vb
' Temporario (Tarefa 0 da cascata comparativa): a Tarefa 1 o remove ao implementar o tipo. Tipo valido ainda sem implementacao avisa e nao
' deixa bloco.
Public Sub Test_Core_Insert_PairWaterfallNotReady_NotifiesWithoutChart()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_ClearLastNotice
    Core_Insert "waterfallpair"
    AssertEqual 0, ws.ChartObjects.Count
    AssertEqual 0, TqqNameCount(wb)
    AssertTrue InStr(Core_LastNotice(), "ainda não está disponível") > 0, Core_LastNotice()
    Teardown wb
End Sub
```

- [ ] **Passo 3: RED** — `robocopy ...` e `pwsh C:\temp\tqq-cc\build\run-tests.ps1 -Target core -Filter "PairWaterfall|Phase3_Matrix" -TimeoutSec 300` → falhas nos asserts novos (`Kinds_IsValid("waterfallpair")` falso, 0 recurso novo); `Test_Kinds_Phase3_Matrix_ExhaustiveFromSpec` continua verde.

- [ ] **Passo 4: Implementar `modChartKinds`** — troque a constante e as funções indicadas (o resto do módulo fica):

```vb
Private Const KINDS As String = ",stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,gantt,pie,doughnut,butterfly,scatter,bubble,waterfallpair,"
```

Em `Kinds_Family`, antes do `End If` final:

```vb
    ElseIf kind = "waterfallpair" Then
        Kinds_Family = "waterfallpair"
```

Em `Kinds_Supports`, troque estas linhas do `Select Case` e acrescente as três últimas antes do `Case Else` (cada uma numa linha só: o `Test-Docs` lê `Case "x": list = "..."` por expressão regular):

```vb
        Case "totals":     list = "stacked,stacked100,barstacked,barstacked100,area,area100,combo,mekko,mekkounits,butterfly,waterfallpair"
        Case "connectors": list = "stacked,stacked100,barstacked,barstacked100,waterfall,barwaterfall,combo,waterfallpair"
        Case "numfmt":     list = "stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,pie,doughnut,butterfly,waterfallpair"
        Case "gap":        list = "stacked,clustered,stacked100,barstacked,barclustered,barstacked100,waterfall,barwaterfall,combo,gantt,butterfly,waterfallpair"
        Case "samescale":  list = "stacked,clustered,line,waterfall,barstacked,barclustered,barwaterfall,area,combo,mekkounits,butterfly,waterfallpair"
        Case "pairdiff":   list = "waterfallpair"
        Case "pairabs":    list = "waterfallpair"
        Case "inbar":      list = "waterfallpair"
```

Em `Kinds_FeatureLabel`, antes do `End Select` (uma linha cada, no formato das existentes):

```vb
        Case "pairdiff":   Kinds_FeatureLabel = "Setas de variação"
        Case "pairabs":    Kinds_FeatureLabel = "Variação em valor"
        Case "inbar":      Kinds_FeatureLabel = "Nome dentro das barras"
```

Troque `Kinds_ExpectedCategoryRefs` inteira:

```vb
' Numero de referencias de categoria que a primeira serie do grafico deve ter (Core_ChartSourceMismatch): o Mekko usa 5 pontos por categoria;
' a cascata comparativa tem 3n - 1 fatias com categorias de 2 colunas (eixo de 2 niveis); os demais, uma por categoria.
Public Function Kinds_ExpectedCategoryRefs(kind As String, n As Long) As Long
    Select Case Kinds_Family(kind)
        Case "mekko": Kinds_ExpectedCategoryRefs = 5 * n
        Case "waterfallpair": Kinds_ExpectedCategoryRefs = 2 * (3 * n - 1)
        Case Else: Kinds_ExpectedCategoryRefs = n
    End Select
End Function
```

`Kinds_Insert` fica como está (a família nova cai no 7200 até a Tarefa 1).

- [ ] **Passo 5: Saídas cedo da família** — `modLabels.Labels_RefreshPositions`, logo depois da linha do `scatter`:

```vb
    If Kinds_Family(kind) = "waterfallpair" Then GoTo Done  ' Cascata comparativa: posicoes e rotulos finos pelo refresh da familia (Wfp_SyncLabels)
```

`modLabels.Labels_ApplyCellFormat`, logo depois da linha do `scatter`:

```vb
    If Kinds_Family(kind) = "waterfallpair" Then Exit Sub     ' Cascata comparativa: formato pela propria familia (Tarefa 4: Wfp_ApplyCellFormat)
```

`modChartStyle.Style_ApplyNumberFormat`, logo depois da linha do `scatter`:

```vb
    If Kinds_Family(Registry_ChartType(cht)) = "waterfallpair" Then Exit Sub   ' Cascata comparativa (Tarefa 4: Wfp_ApplyNumberFormat)
```

- [ ] **Passo 6: Tutorial (`#matriz`)** — em `docs/tutorial/tutorial.html`, nas três tabelas `class="mx"`, acrescente depois da linha `data-kind="bubble"` de cada uma (na ordem das colunas de cada tabela):

Tabela 1 (Totais, Rótulos, Conectores, Formato, Paleta, Espaço, Legenda, fim da linha):
```html
      <tr data-kind="waterfallpair"><th scope="row"><span class="k">Cascata comparativa</span></th><td class="y">Sim</td><td class="y">Sim</td><td class="y">Sim</td><td class="y">Sim</td><td class="y">Sim</td><td class="y">Sim</td><td class="y">Sim</td><td>—</td></tr>
```
Tabela 2 (Indicador 100%, Linha de valor, Eixo secundário, Rótulos fora, CAGR, Diferença, Colchete, Mesma escala):
```html
      <tr data-kind="waterfallpair"><th scope="row"><span class="k">Cascata comparativa</span></th><td>—</td><td>—</td><td>—</td><td>—</td><td>—</td><td>—</td><td>—</td><td class="y">Sim</td></tr>
```
Tabela 3 (Rótulos com soma consistente, Corte de eixo, Linha de tendência, Quadrantes):
```html
      <tr data-kind="waterfallpair"><th scope="row"><span class="k">Cascata comparativa</span></th><td>—</td><td>—</td><td>—</td><td>—</td></tr>
```
Depois do `</div>` que fecha a terceira tabela e antes do `<ul>` seguinte, a 4.ª tabela:
```html
  <h4>Cascata comparativa</h4>
  <div class="tw" tabindex="0" role="region" aria-label="Tabela do que cada tipo aceita: recursos da cascata comparativa (em telas estreitas, role para os lados)">
  <table class="mx">
    <thead><tr><th scope="col">Tipo</th><th scope="col" data-f="pairdiff"><span class="k">Setas de variação</span></th><th scope="col" data-f="pairabs"><span class="k">Variação em valor</span></th><th scope="col" data-f="inbar"><span class="k">Nome dentro das barras</span></th></tr></thead>
    <tbody>
      <tr data-kind="stacked"><th scope="row"><span class="ui">Colunas empilhadas</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="clustered"><th scope="row"><span class="ui">Colunas agrupadas</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="stacked100"><th scope="row"><span class="ui">Colunas 100%</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="barstacked"><th scope="row"><span class="ui">Barras empilhadas</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="barclustered"><th scope="row"><span class="ui">Barras agrupadas</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="barstacked100"><th scope="row"><span class="ui">Barras 100%</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="waterfall"><th scope="row"><span class="ui">Cascata (colunas)</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="barwaterfall"><th scope="row"><span class="ui">Cascata (barras)</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="line"><th scope="row"><span class="ui">Linha</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="area"><th scope="row"><span class="ui">Área empilhada</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="area100"><th scope="row"><span class="ui">Área 100%</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="combo"><th scope="row"><span class="ui">Combinação (colunas + linha)</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="mekko"><th scope="row"><span class="ui">Mekko (%)</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="mekkounits"><th scope="row"><span class="ui">Mekko (unidades)</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="gantt"><th scope="row"><span class="ui">Gantt</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="pie"><th scope="row"><span class="ui">Pizza</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="doughnut"><th scope="row"><span class="ui">Rosca</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="butterfly"><th scope="row"><span class="ui">Borboleta</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="scatter"><th scope="row"><span class="ui">Dispersão</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="bubble"><th scope="row"><span class="ui">Bolhas</span></th><td>—</td><td>—</td><td>—</td></tr>
      <tr data-kind="waterfallpair"><th scope="row"><span class="k">Cascata comparativa</span></th><td class="y">Sim</td><td class="y">Sim</td><td class="y">Sim</td></tr>
    </tbody>
  </table>
  </div>
```
(O traço das células é o caractere U+2014 "—", o mesmo das tabelas existentes.) Os rótulos novos ficam em `<span class="k">` (sem `ui`): o `Test-Docs` confere todo `ui` contra os rótulos da faixa, e "Cascata comparativa", "Setas de variação" e "Variação em valor" só entram na faixa na Tarefa 5; o cabeçalho é conferido contra `Kinds_FeatureLabel` (lido do código), qualquer que seja a classe.

- [ ] **Passo 7: Test-Docs e Test-DemoWorkbook** — `tests/ps/Test-Docs.ps1`: troque as duas linhas de contagem por

```powershell
    Check ($allKinds.Count -eq 21) "21 tipos de gráfico em modChartKinds.bas ($($allKinds.Count))"
    Check ($matrix.Count -eq 23) "23 recursos na matriz de capacidades de modChartKinds.bas ($($matrix.Count))"
```

`tests/ps/Test-DemoWorkbook.ps1`: logo depois de `$feats = @(Get-TqqDemoFeatures)` acrescente

```powershell
    # TEMPORARIO (Tarefa 0 da cascata comparativa): tipo e recursos que ainda nao tem exemplo na demonstracao; a Tarefa 7 gera E27/E28 e remove
    # esta lista (e o "if" das duas linhas seguintes)
    $pendingDemo = @('waterfallpair', 'pairdiff', 'pairabs', 'inbar')
```

e troque as duas linhas `foreach ($k in $allKinds) {...}` e `foreach ($f in $matrixFeatures) {...}` por

```powershell
    foreach ($k in $allKinds) { if ($pendingDemo -contains $k) { Write-Host "ok - cobertura do tipo '$k' # SKIP pendente (Tarefa 7)"; continue }; Check (@($cov | Where-Object { $_.Kind -eq $k }).Count -ge 1) "cobertura tem o tipo '$k'" }
```
```powershell
    foreach ($f in $matrixFeatures) { if ($pendingDemo -contains $f) { Write-Host "ok - cobertura do recurso '$f' # SKIP pendente (Tarefa 7)"; continue }; Check (@($cov | Where-Object { $_.Features -contains $f -or @($_.Features | Where-Object { $_ -like "$f-*" }).Count -gt 0 }).Count -ge 1) "cobertura tem o recurso da matriz '$f'" }
```

- [ ] **Passo 8: GREEN e suíte completa** — filtro do passo 3 verde (os dois testes novos, o temporário e o da Fase 3); depois `pwsh C:\temp\tqq-cc\build\run-tests.ps1 -TimeoutSec 900` verde (inclusive `Test-Docs` com 21/23 e `Test-DemoWorkbook` com as 4 pendências em SKIP). Saída esperada no fim: "Todos os testes passaram", código 0.

- [ ] **Passo 9: Commit** — `git add docs/superpowers/specs/2026-10-13-thinqq-cascata-comparativa-design.md docs/superpowers/plans/2026-10-13-thinqq-cascata-comparativa.md src/core/modules/modChartKinds.bas src/core/modules/modLabels.bas src/core/modules/modChartStyle.bas docs/tutorial/tutorial.html tests/ps/Test-Docs.ps1 tests/ps/Test-DemoWorkbook.ps1 tests/vba/core/TestKinds.bas tests/vba/core/TestCoreApi.bas` e `git commit -m "feat: matriz 1.3.0 (cascata comparativa: tipo, familia, setas de variacao, variacao em valor, nome dentro das barras), spec e plano"` (+ trailer).

**Verificação no Excel real:** nenhuma mudança visível; Cascata (colunas) e Borboleta continuam iguais.

---
### Tarefa 1: Construção da cascata comparativa (bloco, barras, eixo de 2 níveis, cores, linha do zero, nomes e legenda) — modelo: most-capable

**Files:**
- Create: `src/core/modules/modChartWaterfallPair.bas`
- Modify: `src/core/modules/modChartKinds.bas` (`Kinds_Insert`), `src/core/modules/modCoreApi.bas` (paleta, espaço, legenda `inbar`, refresh, mismatch), `src/core/modules/modDecorations.bas` (`Deco_SetLegend`), `src/core/modules/modAnnotations.bas` (`Ann_GetCategoryData`)
- Test: `tests/vba/core/TestWaterfallPair.bas` (novo), `tests/vba/core/TestCoreApi.bas` (remove o teste temporário da Tarefa 0)

**Interfaces:**
- Consumes: `CalcSheet_AllocateBlock`, `Block_Body`, `Block_Ref`, `ChartData_ExternalAddress`, `Chart_FreeTop`, `Chart_FindSeries`, `Chart_HideHelperLegendEntries`, `Style_ApplyBase`, `Style_Palette`, `Style_SetColor`, `Style_IsHelperSeries`, `Style_ContrastFont`, `Style_DataSeriesColor`, `Labels_EstimateWidth`, `Registry_*`, `Util_*`, `Kinds_*`, `Errors_Log`, `WF_MARKERS` (`modChartWaterfall`).
- Produces (`modChartWaterfallPair`): as constantes `WFP_COL_*`, `WFP_META_*`, `WFP_ROWS_PER_PAIR` (6), `WFP_DEFAULT_GAP` (30) do design §4.2; `Wfp_PartIndex(part) As Long`, `Wfp_SideCol(j, part) As Long` (lado 1 = colunas 2..13, lado 2 = 14..25; partes `Src, Tot, D, Start, End, Lo, Hi, Base, PP, PN, Val, Big`), `Wfp_SlotCount(n)` (3n − 1), `Wfp_SlotOf(i, j)` (3(i − 1) + j), `Wfp_HalfWidth(gapPct)` (0,5 / (1 + gap/100)), `Wfp_StepCount(cht)`, `Wfp_SideSeries(cht, j) As Series`, `Wfp_BuildBlock(wb, id, data) As Range`, `Wfp_Insert(ws, data) As ChartObject` (6016/6017 antes de alocar), `Wfp_ApplyColors(cht, paletteName)`, `Wfp_SetLegend(cht, mode)` (`right|none|inbar`), `Wfp_SyncNames(cht)`, `Wfp_SyncGap(cht)`, `Wfp_RefreshAfterDataChange(cht)` (nunca propaga erro; guarda `mBusy`), `Wfp_SourceRowsSpanned(cht) As Long`; privadas `AddBar`, `AddXY(cht, name, xCells, yCells, withLine) As Series`, `SetMeta`, `ShowLegend`, `Wfp_RefreshFontColors`, `PointHasLabel`.
- Mensagens exatas: 6016 "A cascata comparativa precisa de exatamente 2 colunas de valores (uma por série); a tabela tem só 1."; 6017 "A cascata comparativa compara exatamente 2 séries; a tabela tem " & N & " colunas de valores. Selecione a coluna das etapas e só as duas colunas a comparar."
- Opções gravadas pelo Insert nesta tarefa: `numfmt=int, palette=padrao, gap=30, n, totals=1, legend=inbar` (as Tarefas 2 a 4 acrescentam `connectors`, `pairdiff`, `pairabs`, `labels`).

- [ ] **Passo 1: Testes (RED)** — crie `tests/vba/core/TestWaterfallPair.bas`:

```vb
Attribute VB_Name = "TestWaterfallPair"
Option Explicit
' Cascata comparativa (waterfallpair). Geometria por regra de tres sobre a area interna e a escala REAL do eixo (XOfX, YOf), independente das
' formulas do bloco; textos com Format$ (localidade).

' A imagem do think-cell: Segmento 1 | Segmento 2, Bruto 27 | 29, tres componentes que descem, Liquido = total.
Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Etapa", "Segmento 1", "Segmento 2")
    ws.Range("A2:C2").Value = Array("Bruto", 27, 29)
    ws.Range("A3:C3").Value = Array("Componente 1", -4, -7)
    ws.Range("A4:C4").Value = Array("Componente 2", -6, -2)
    ws.Range("A5:C5").Value = Array("Componente 3", -6, -10)
    ws.Range("A6:C6").Value = Array("Líquido", "e", "e")
    Set Fixture = ChartData_FromRange(ws.Range("A1:C6"))
End Function

' Cruza o zero: a serie A desce abaixo de zero e termina negativa; a B cruza para cima.
Private Function FixtureZero(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Etapa", "Loja A", "Loja B")
    ws.Range("A2:C2").Value = Array(2024, 10, 5)
    ws.Range("A3:C3").Value = Array("Queda", -25, -12)
    ws.Range("A4:C4").Value = Array("Alta", 8, 20)
    ws.Range("A5:C5").Value = Array("Fim", "t", "total")
    Set FixtureZero = ChartData_FromRange(ws.Range("A1:C5"))
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

Private Function SlotN(cht As Chart) As Long
    SlotN = cht.SeriesCollection(1).Points.Count
End Function

' x (pt) da fatia (ou do X de uma auxiliar XY do grupo principal: o mesmo numero), eixo de categorias "entre" as marcas.
Private Function XOfX(cht As Chart, x As Double) As Double
    XOfX = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (x - 0.5) / SlotN(cht)
End Function

Private Function YOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, xlPrimary)
    YOf = cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight * (ax.MaximumScale - v) / (ax.MaximumScale - ax.MinimumScale)
End Function

Private Function BodyOf(cht As Chart) As Range
    Set BodyOf = Block_Body(Registry_Block(cht))
End Function

Private Function SideVal(cht As Chart, i As Long, j As Long, part As String) As Double
    SideVal = Util_ToDouble(BodyOf(cht).Cells(i + 1, Wfp_SideCol(j, part)).Value)
End Function

' "27,23,17,11,11": o fim da ponte do lado j em cada etapa.
Private Function EndsText(cht As Chart, j As Long) As String
    Dim i As Long
    For i = 1 To Wfp_StepCount(cht)
        EndsText = EndsText & IIf(i > 1, ",", "") & CStr(SideVal(cht, i, j, "End"))
    Next
End Function

Public Sub Test_Wfp_Insert_SeriesBlockOptions()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "waterfallpair", Registry_ChartType(cht)
    AssertEqual xlColumnStacked, cht.SeriesCollection(2).ChartType, "colunas empilhadas (o grafico le xlCombination por causa das auxiliares XY)"
    AssertEqual "tqq_base", cht.SeriesCollection(1).Name: AssertEqual "Segmento 1", cht.SeriesCollection(2).Name
    AssertEqual "tqq_aneg", cht.SeriesCollection(3).Name: AssertEqual "Segmento 2", cht.SeriesCollection(4).Name
    AssertEqual "tqq_bneg", cht.SeriesCollection(5).Name
    AssertEqual msoFalse, cht.SeriesCollection(1).Format.Fill.Visible, "base invisivel"
    AssertEqual 14, SlotN(cht), "3n - 1 fatias (A, B e o vao; a ultima etapa sem vao)"
    AssertEqual 5, Wfp_StepCount(cht)
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 58, blk.Columns.Count: AssertEqual 31, body.Rows.Count, "cabecalho + 6n linhas"
    AssertEqual "27,23,17,11,11", EndsText(cht, 1): AssertEqual "29,22,20,10,10", EndsText(cht, 2)
    AssertTrue CBool(body.Cells(2, Wfp_SideCol(1, "Big")).Value), "a 1a barra e grande (inicio)"
    AssertFalse CBool(body.Cells(3, Wfp_SideCol(1, "Big")).Value): AssertTrue CBool(body.Cells(6, Wfp_SideCol(2, "Big")).Value), "o total e grande"
    AssertEqual "Bruto", body.Cells(2, WFP_COL_OUTER).Text, "fatia A leva o nome da etapa"
    AssertTrue IsEmpty(body.Cells(3, WFP_COL_OUTER).Value), "fatia B sem nome (o grupo do eixo de 2 niveis cobre A e B)"
    AssertEqual " ", CStr(body.Cells(4, WFP_COL_OUTER).Value), "vao"
    AssertEqual " ", CStr(body.Cells(2, WFP_COL_INNER).Value), "nivel interno"
    AssertEqual 27, body.Cells(2, WFP_COL_APOS).Value: AssertEqual 0, body.Cells(2, WFP_COL_BPOS).Value
    AssertEqual 29, body.Cells(3, WFP_COL_BPOS).Value: AssertEqual 0, body.Cells(3, WFP_COL_APOS).Value
    AssertEqual 23, body.Cells(5, WFP_COL_SBASE).Value, "Componente 1 de A: base 23": AssertEqual 4, body.Cells(5, WFP_COL_APOS).Value
    AssertEqual 0, Application.WorksheetFunction.Sum(body.Cells(4, WFP_COL_SBASE).Resize(1, 5)), "vao vazio"
    AssertEqual "int", Registry_GetOption(cht, "numfmt"): AssertEqual "padrao", Registry_GetOption(cht, "palette")
    AssertEqual "30", Registry_GetOption(cht, "gap"): AssertEqual "5", Registry_GetOption(cht, "n")
    AssertEqual "1", Registry_GetOption(cht, "totals"): AssertEqual "inbar", Registry_GetOption(cht, "legend")
    AssertEqual 30, cht.ChartGroups(1).GapWidth
    AssertAlmost 0.5 / 1.3, Util_ToDouble(blk.Cells(1, WFP_META_HW).Value), 0.000001, "meia-largura em fatias"
    AssertTrue cht.Axes(xlCategory).TickLabels.MultiLevel, "eixo de categorias de 2 niveis"
    AssertEqual msoFalse, cht.Axes(xlCategory).Format.Line.Visible, "sem a linha do eixo (separadores)"
    AssertEqual xlNotPlotted, cht.DisplayBlanksAs
    AssertTrue Not Chart_FindSeries(cht, "tqq_zero") Is Nothing: AssertTrue Not Chart_FindSeries(cht, "tqq_names") Is Nothing
    AssertEqual xlPrimary, Chart_FindSeries(cht, "tqq_zero").AxisGroup, "auxiliar XY no grupo principal"
    AssertFalse Core_ChartSourceMismatch(cht)
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Geometry_SlotsAndBars()
    Dim wb As Workbook, d As clsChartData, cht As Chart, i As Long, j As Long, k As Long, p As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    ForceLayout cht
    For i = 1 To 5
        For j = 1 To 2
            k = Wfp_SlotOf(i, j)
            If j = 1 Then Set p = cht.SeriesCollection(2).Points(k) Else Set p = cht.SeriesCollection(4).Points(k)
            AssertAlmost XOfX(cht, CDbl(k)), p.Left + p.Width / 2, 1, "centro da barra " & i & "/" & j
            AssertAlmost cht.PlotArea.InsideWidth / SlotN(cht) / 1.3, p.Width, 1, "largura (espaco 30%) " & i & "/" & j
            AssertAlmost YOf(cht, SideVal(cht, i, j, "Hi")), p.Top, 1.5, "topo da barra " & i & "/" & j
            AssertAlmost YOf(cht, SideVal(cht, i, j, "Lo")), p.Top + p.Height, 1.5, "base da barra " & i & "/" & j
        Next
    Next
    CloseTemp wb
End Sub

Public Sub Test_Wfp_CrossesZero_NegativeTotal()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range, pp As Point, pn As Point, z As Series
    Set wb = NewTempWorkbook(): Set d = FixtureZero(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    Set body = BodyOf(cht)
    AssertEqual "10,-15,-7,-7", EndsText(cht, 1): AssertEqual "5,-7,13,13", EndsText(cht, 2)
    AssertEqual "2024", body.Cells(2, WFP_COL_OUTER).Text, "ano como categoria"
    ' A2 (fatia 4) cruza o zero: base 0, +10 e -15
    AssertEqual 0, body.Cells(5, WFP_COL_SBASE).Value: AssertEqual 10, body.Cells(5, WFP_COL_APOS).Value: AssertEqual -15, body.Cells(5, WFP_COL_ANEG).Value
    ' A3 (fatia 7) inteira abaixo de zero: base invisivel -7 e a parte visivel -8 (de -7 a -15)
    AssertEqual -7, body.Cells(8, WFP_COL_SBASE).Value: AssertEqual 0, body.Cells(8, WFP_COL_APOS).Value: AssertEqual -8, body.Cells(8, WFP_COL_ANEG).Value
    ' B3 (fatia 8) cruza para cima: -7 a 13
    AssertEqual 13, body.Cells(9, WFP_COL_BPOS).Value: AssertEqual -7, body.Cells(9, WFP_COL_BNEG).Value
    ForceLayout cht
    Set pp = cht.SeriesCollection(2).Points(4): Set pn = cht.SeriesCollection(3).Points(4)
    AssertAlmost YOf(cht, 10), pp.Top, 1.5, "A2 parte positiva: topo em 10": AssertAlmost YOf(cht, 0), pp.Top + pp.Height, 1.5
    AssertAlmost YOf(cht, 0), pn.Top, 1.5, "A2 parte negativa: de 0": AssertAlmost YOf(cht, -15), pn.Top + pn.Height, 1.5, "a -15"
    Set pn = cht.SeriesCollection(3).Points(7)
    AssertAlmost YOf(cht, -7), pn.Top, 1.5, "A3 de -7": AssertAlmost YOf(cht, -15), pn.Top + pn.Height, 1.5, "a -15"
    Set pn = cht.SeriesCollection(3).Points(10)
    AssertAlmost YOf(cht, 0), pn.Top, 1.5, "total negativo de 0": AssertAlmost YOf(cht, -7), pn.Top + pn.Height, 1.5, "a -7"
    Set z = Chart_FindSeries(cht, "tqq_zero")
    AssertAlmost YOf(cht, 0), z.Points(1).Top, 1, "linha do zero na altura do zero"
    AssertAlmost cht.PlotArea.InsideLeft, z.Points(1).Left, 1, "da borda esquerda"
    AssertAlmost cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth, z.Points(2).Left, 1, "a borda direita"
    CloseTemp wb
End Sub

' Marcadores de total por celula (cada serie tem a sua ponte), texto e vazio valem 0, marcador na 1a linha = total de nada (0).
Public Sub Test_Wfp_Markers_TextAndFirstRow()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Etapa", "S1", "S2")
    ws.Range("A2:C2").Value = Array("Inicio", 10, "abc")
    ws.Range("A3:C3").Value = Array("Sub", "t", "'=")          ' "'=": texto (um "=" sozinho seria lido como formula)
    ws.Range("A4").Value = "Mais": ws.Range("B4").Value = "x": ws.Range("C4").Value = 5
    ws.Range("A5:C5").Value = Array("Fim", " Total ", "SUBTOTAL")
    Set d = ChartData_FromRange(ws.Range("A1:C5"))
    Set cht = Wfp_Insert(ws, d).Chart
    AssertEqual "10,10,10,10", EndsText(cht, 1), "t e Total (com espacos) sao totais; texto vale 0"
    AssertEqual "0,0,5,5", EndsText(cht, 2), "texto na 1a linha = 0; '=' e SUBTOTAL sao totais"
    AssertTrue CBool(BodyOf(cht).Cells(3, Wfp_SideCol(1, "Tot")).Value): AssertFalse CBool(BodyOf(cht).Cells(4, Wfp_SideCol(1, "Tot")).Value)
    ws.Range("B2").Value = "e"
    AssertEqual 0, SideVal(cht, 1, 1, "End"), "marcador na 1a linha: total de nada"
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Refusals_And_Rollback()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, en As Long, ed As String, nm As Name, d2 As clsChartData, cs As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:B6").Select
    Core_ClearLastNotice
    Core_Insert "waterfallpair"
    AssertEqual "A cascata comparativa precisa de exatamente 2 colunas de valores (uma por série); a tabela tem só 1.", Core_LastNotice()
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    ws.Range("D1").Value = "Segmento 3": ws.Range("D2:D6").Value = 1
    ws.Range("A1:D6").Select
    Core_ClearLastNotice
    Core_Insert "waterfallpair"
    AssertEqual "A cascata comparativa compara exatamente 2 séries; a tabela tem 3 colunas de valores. Selecione a coluna das etapas e só as duas colunas a comparar.", Core_LastNotice()
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    ws.Range("D1:D6").ClearContents
    Set cht = Wfp_Insert(ws, d).Chart
    On Error Resume Next
    Ann_AddCagr cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    AssertEqual vbObjectError + 7001, en: AssertEqual "A seta CAGR não está disponível para este tipo de gráfico.", ed
    Deco_SetValueLine cht, True
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en, "linha de valor: fora da matriz"
    ' "Nome dentro das barras" em outro tipo: recusa 7002
    ws.Range("H1:J1").Value = Array("Ano", "A", "B"): ws.Range("H2:J2").Value = Array(2024, 1, 2): ws.Range("H3:J3").Value = Array(2025, 2, 3)
    Set d2 = ChartData_FromRange(ws.Range("H1:J3"))
    Set cs = Stacked_Insert(ws, d2, "stacked").Chart
    On Error Resume Next
    Deco_SetLegend cs, "inbar"
    en = Err.Number: ed = Err.Description: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en: AssertEqual "'Nome dentro das barras' não se aplica a este tipo de gráfico.", ed
    ' rollback
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
    CalcSheet_Cleanup wb
    AssertEqual 0, TqqNameCount(wb), "blocos orfaos limpos"
    On Error GoTo Cleanup
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Wfp_Insert ws, d
    en = Err.Number
    On Error GoTo Cleanup
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, en
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    Util_FailAfterChartForTest = True
    On Error Resume Next
    Wfp_Insert ws, d
    en = Err.Number
    On Error GoTo Cleanup
    Util_FailAfterChartForTest = False
    AssertEqual vbObjectError + 9998, en
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    Teardown wb
    Exit Sub
Cleanup:
    en = Err.Number: ed = Err.Description
    Util_FailAfterBlockForTest = False: Util_FailAfterChartForTest = False
    Teardown wb
    Err.Raise en, "TestWaterfallPair", ed
End Sub

' Review Focus 4: nomes verticais dentro da 1a barra quando cabem; Direita / Nenhuma / Nome dentro das barras; barra baixa -> legenda.
Public Sub Test_Wfp_NamesInsideBars_And_LegendModes()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, s As Series, lbl As DataLabel, p As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C6").Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    Set s = Chart_FindSeries(cht, "tqq_names")
    cht.Refresh
    AssertEqual "Segmento 1", s.Points(1).DataLabel.Text: AssertEqual "Segmento 2", s.Points(2).DataLabel.Text
    AssertEqual xlUpward, s.Points(1).DataLabel.Orientation
    AssertFalse cht.HasLegend, "os nomes cabem: sem legenda"
    AssertEqual 1, Util_ToDouble(Registry_Block(cht).Cells(1, WFP_META_FIT).Value)
    ForceLayout cht
    Set lbl = s.Points(1).DataLabel: Set p = cht.SeriesCollection(2).Points(1)
    AssertAlmost p.Left + p.Width / 2, lbl.Left + lbl.Width / 2, 1.5, "nome centrado na 1a barra"
    AssertTrue lbl.Top >= p.Top - 1 And lbl.Top + lbl.Height <= p.Top + p.Height + 1, "nome dentro da barra"
    AssertEqual RGB(255, 255, 255), lbl.Font.Color, "fonte branca na barra escura"
    AssertEqual RGB(64, 64, 64), s.Points(2).DataLabel.Font.Color, "fonte escura na barra clara"
    Core_SetOption "legend", "right"
    cht.Refresh
    AssertTrue cht.HasLegend: AssertEqual 2, cht.Legend.LegendEntries.Count, "legenda so com as duas series"
    AssertEqual "", s.Points(1).DataLabel.Text, "Direita: sem nomes nas barras"
    Core_SetOption "legend", "none"
    cht.Refresh
    AssertFalse cht.HasLegend: AssertEqual "", s.Points(1).DataLabel.Text
    Core_SetOption "legend", "inbar"
    cht.Refresh
    AssertFalse cht.HasLegend: AssertEqual "Segmento 1", s.Points(1).DataLabel.Text
    ' nao cabe: 1a barra de A baixa -> os dois nomes somem e a legenda aparece (troca automatica no refresh)
    ws.Range("B2").Value = 2
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual "", s.Points(1).DataLabel.Text: AssertEqual "", s.Points(2).DataLabel.Text
    AssertTrue cht.HasLegend, "legenda no lugar dos nomes": AssertEqual 2, cht.Legend.LegendEntries.Count
    AssertEqual "inbar", Registry_GetOption(cht, "legend"), "a escolha do usuario continua sendo 'dentro das barras'"
    ws.Range("B2").Value = 27
    Core_SafeRefresh cht
    cht.Refresh
    AssertFalse cht.HasLegend, "voltou a caber": AssertEqual "Segmento 1", s.Points(1).DataLabel.Text
    Teardown wb
End Sub

Public Sub Test_Wfp_CategoryLabels_TwoLevels_DateAndBlank()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A3").Value = DateSerial(2026, 1, 15): ws.Range("A3").NumberFormat = "dd/mm/yyyy"
    ws.Range("A4").ClearContents
    Set cht = Wfp_Insert(ws, d).Chart
    Set body = BodyOf(cht)
    AssertEqual ws.Range("A3").Text, body.Cells(5, WFP_COL_OUTER).Text, "data com o formato da origem (fatia A2)"
    AssertEqual " ", CStr(body.Cells(8, WFP_COL_OUTER).Value), "categoria vazia vira um espaco (fatia A3)"
    AssertTrue IsEmpty(body.Cells(6, WFP_COL_OUTER).Value), "fatia B2 sem nome"
    AssertEqual " ", CStr(body.Cells(7, WFP_COL_OUTER).Value), "vao 2"
    AssertEqual xlTickLabelPositionLow, cht.Axes(xlCategory).TickLabelPosition
    AssertEqual xlHorizontal, cht.Axes(xlCategory).TickLabels.Orientation
    CloseTemp wb
End Sub

' Review Focus 5 (parte das barras): sem o ThinQQ (nenhum refresh), mudar a origem move a ponte.
Public Sub Test_Wfp_DataChange_LiveWithoutThinQQ()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, v As Variant, p As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    ws.Range("C3").Value = -12
    cht.Refresh
    AssertEqual "29,17,15,5,5", EndsText(cht, 2)
    v = cht.SeriesCollection(4).Values: AssertEqual 12, v(5), "B2: parte visivel 12"
    v = cht.SeriesCollection(1).Values: AssertEqual 17, v(5), "B2: base 17"
    ForceLayout cht
    Set p = cht.SeriesCollection(4).Points(5)
    AssertAlmost YOf(cht, 29), p.Top, 1.5: AssertAlmost YOf(cht, 17), p.Top + p.Height, 1.5
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Gap_SyncsHalfWidth()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart(): Set blk = Registry_Block(cht)
    Core_SetOption "gap", "80"
    AssertEqual 80, cht.ChartGroups(1).GapWidth
    AssertAlmost 0.5 / 1.8, Util_ToDouble(blk.Cells(1, WFP_META_HW).Value), 0.000001, "HW acompanha o Espaco"
    cht.ChartGroups(1).GapWidth = 120                 ' mudado a mao no Excel
    Core_SafeRefresh cht
    AssertAlmost 0.5 / 2.2, Util_ToDouble(blk.Cells(1, WFP_META_HW).Value), 0.000001, "o refresh le o espaco real"
    Teardown wb
End Sub

Public Sub Test_Wfp_SourceSizeChange_Detected()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    AssertEqual 5, Wfp_SourceRowsSpanned(cht)
    AssertFalse Core_ChartSourceMismatch(cht)
    ws.Rows(4).Insert
    ws.Range("A4:C4").Value = Array("Extra", -1, -1)
    AssertEqual 6, Wfp_SourceRowsSpanned(cht)
    AssertTrue Core_ChartSourceMismatch(cht), "linha inserida no meio da tabela"
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Palette_And_Orphan()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    AssertEqual RGB(31, 56, 100), cht.SeriesCollection(2).Format.Fill.ForeColor.RGB, "serie 1: pal(0)"
    AssertEqual RGB(157, 195, 230), cht.SeriesCollection(4).Format.Fill.ForeColor.RGB, "serie 2: pal(2)"
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(2).Format.Fill.ForeColor.RGB: AssertEqual RGB(64, 64, 64), cht.SeriesCollection(3).Format.Fill.ForeColor.RGB
    AssertEqual RGB(166, 166, 166), cht.SeriesCollection(4).Format.Fill.ForeColor.RGB: AssertEqual RGB(166, 166, 166), cht.SeriesCollection(5).Format.Fill.ForeColor.RGB
    AssertEqual msoFalse, cht.SeriesCollection(1).Format.Fill.Visible, "a base continua invisivel"
    AssertEqual RGB(64, 64, 64), Chart_FindSeries(cht, "tqq_names").Points(2).DataLabel.Font.Color, "contraste com o cinza claro"
    ws.ChartObjects(1).Delete
    AssertEqual 1, CalcSheet_Cleanup(wb), "o bloco do grafico apagado e removido"
    Teardown wb
End Sub
```

Em `tests/vba/core/TestCoreApi.bas`, apague o teste temporário `Test_Core_Insert_PairWaterfallNotReady_NotifiesWithoutChart` (Tarefa 0).

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-cc\build\run-tests.ps1 -Target core -Filter "TestWaterfallPair" -TimeoutSec 400` → saída 3 (`Wfp_Insert` indefinido).

- [ ] **Passo 3: Implementar `src/core/modules/modChartWaterfallPair.bas`**:

```vb
Attribute VB_Name = "modChartWaterfallPair"
Option Explicit
' Cascata comparativa (kind waterfallpair, familia waterfallpair, vertical): duas pontes (series 1 e 2 da tabela) lado a lado em cada etapa,
' como o think-cell monta com lacunas de categoria (design da cascata comparativa, secoes 3 e 4). Colunas EMPILHADAS com 3n - 1 categorias
' (fatias): A_i = 3i - 2, B_i = 3i - 1 e o vao_i = 3i (a ultima etapa nao tem vao). Series na ordem tqq_base (invisivel), <serie 1> (parte
' >= 0 de A), tqq_aneg (parte <= 0 de A), <serie 2> (parte >= 0 de B), tqq_bneg: cada fatia so tem valores do seu lado, e base invisivel +
' parte positiva + parte negativa cruzam o zero como a cascata da Fase 1. Categorias de 2 niveis (Outer: nome da etapa na fatia A, vazio na B,
' " " no vao; Inner: " "): o Excel centra o nome sob o par A-B (medido). A linha do eixo de categorias fica invisivel (desenharia os separadores
' do eixo de 2 niveis); a base e a auxiliar XY tqq_zero. Conectores (tqq_conn_a/b), setas de variacao (tqq_pd), sala das setas (tqq_pdroom) e
' nomes das series (tqq_names) sao auxiliares XY no grupo PRINCIPAL (x = numero da fatia, medido) ligadas ao bloco: vivas sem o ThinQQ.
' Rotulos de degrau (tqq_lbl_a/b) e das barras grandes (tqq_lbltot) sao auxiliares de linha. O ThinQQ so recalibra, no refresh, a meia-largura
' das barras (HW), o valor por ponto das setas (PtV), os rotulos finos (LblOn) e "nome dentro da barra ou legenda" (Fit).
' Bloco (corpo, linha 1 = cabecalhos, 6n linhas): etapas (n linhas) 1 Cat | 2..13 lado A | 14..25 lado B (Src Tot D Start End Lo Hi Base PP PN
' Val Big) | 26 ArrowTxt | 27 TotTxtA | 28 TotTxtB; fatias (3n - 1 linhas) 29 Outer .. 43 Top; conectores (n - 1) 44..48; setas (6n - 1:
' 5 pontos por par + linha vazia) 49..51; sala (n) 52..53; nomes (2) 54..56; zero (2) 57..58. Meta (linha 1 do bloco): E PtV, F HW, G Div,
' H Dec, I Suf, J SufD, K Abs, L Lbl, M TotPct, N NameOn, O Fit.
Public Const WFP_COL_CAT As Long = 1
Public Const WFP_SIDE_BASE As Long = 2
Public Const WFP_SIDE_WIDTH As Long = 12
Public Const WFP_COL_ARROWTXT As Long = 26
Public Const WFP_COL_TOTTXT_A As Long = 27
Public Const WFP_COL_TOTTXT_B As Long = 28
Public Const WFP_COL_OUTER As Long = 29
Public Const WFP_COL_INNER As Long = 30
Public Const WFP_COL_SBASE As Long = 31
Public Const WFP_COL_APOS As Long = 32
Public Const WFP_COL_ANEG As Long = 33
Public Const WFP_COL_BPOS As Long = 34
Public Const WFP_COL_BNEG As Long = 35
Public Const WFP_COL_LBLAPOS As Long = 36
Public Const WFP_COL_LBLAVAL As Long = 37
Public Const WFP_COL_LBLBPOS As Long = 38
Public Const WFP_COL_LBLBVAL As Long = 39
Public Const WFP_COL_TOTPOS As Long = 40
Public Const WFP_COL_TOTTXT As Long = 41
Public Const WFP_COL_LBLON As Long = 42
Public Const WFP_COL_TOP As Long = 43
Public Const WFP_COL_CXA As Long = 44
Public Const WFP_COL_CYA As Long = 45
Public Const WFP_COL_CLEN As Long = 46
Public Const WFP_COL_CXB As Long = 47
Public Const WFP_COL_CYB As Long = 48
Public Const WFP_COL_PX As Long = 49
Public Const WFP_COL_PY As Long = 50
Public Const WFP_COL_PLBL As Long = 51
Public Const WFP_COL_RX As Long = 52
Public Const WFP_COL_RY As Long = 53
Public Const WFP_COL_NX As Long = 54
Public Const WFP_COL_NY As Long = 55
Public Const WFP_COL_NTXT As Long = 56
Public Const WFP_COL_ZX As Long = 57
Public Const WFP_COL_ZY As Long = 58
Private Const WFP_COLS As Long = 58
Public Const WFP_META_PTV As Long = 5
Public Const WFP_META_HW As Long = 6
Public Const WFP_META_DIV As Long = 7
Public Const WFP_META_DEC As Long = 8
Public Const WFP_META_SUF As Long = 9
Public Const WFP_META_SUFD As Long = 10
Public Const WFP_META_ABS As Long = 11
Public Const WFP_META_LBL As Long = 12
Public Const WFP_META_TOTPCT As Long = 13
Public Const WFP_META_NAMEON As Long = 14
Public Const WFP_META_FIT As Long = 15
Public Const WFP_ROWS_PER_PAIR As Long = 6       ' linhas do bloco por etapa (as setas tem 5 pontos e uma linha vazia por par)
Public Const WFP_DEFAULT_GAP As Long = 30        ' espaco entre as barras do par (%), o padrao do tipo
Public Const WFP_PT_GAP As Long = 3              ' pt entre o topo da barra e o inicio da seta
Public Const WFP_PT_LBLROOM As Long = 22        ' pt acima de uma barra grande com rotulo: 5,5 de afastamento + 14 do rotulo (medidos) + 2,5
Public Const WFP_PT_RISE As Long = 12            ' pt do inicio mais alto da seta ate a linha horizontal
Public Const WFP_PT_ROOM As Long = 12            ' pt da linha horizontal ate o ponto de sala (meia elipse de 19,9 pt medida + respiro)
Private Const WFP_NAME_MARGIN As Double = 8      ' pt alem da largura estimada do nome (texto vertical)
Private Const WFP_NAME_MIN_W As Double = 14      ' largura (pt) de um texto vertical de 9 pt (medido: 14,0)
Private Const WFP_ERR_FEW As String = "A cascata comparativa precisa de exatamente 2 colunas de valores (uma por série); a tabela tem só 1."
Private Const WFP_ERR_MANY1 As String = "A cascata comparativa compara exatamente 2 séries; a tabela tem "
Private Const WFP_ERR_MANY2 As String = " colunas de valores. Selecione a coluna das etapas e só as duas colunas a comparar."
Private mBusy As Boolean                         ' o refresh grava celulas do bloco: o evento Calculate do grafico reentra em Core_SafeRefresh

Private Sub Wfp_CheckData(data As clsChartData)
    If data.SeriesCount < 2 Then Err.Raise vbObjectError + 6016, "Wfp_Insert", WFP_ERR_FEW
    If data.SeriesCount > 2 Then Err.Raise vbObjectError + 6017, "Wfp_Insert", WFP_ERR_MANY1 & data.SeriesCount & WFP_ERR_MANY2
End Sub

' Indice da parte do lado (0..11) na secao das etapas; -1 para parte desconhecida.
Public Function Wfp_PartIndex(part As String) As Long
    Select Case part
        Case "Src": Wfp_PartIndex = 0
        Case "Tot": Wfp_PartIndex = 1
        Case "D": Wfp_PartIndex = 2
        Case "Start": Wfp_PartIndex = 3
        Case "End": Wfp_PartIndex = 4
        Case "Lo": Wfp_PartIndex = 5
        Case "Hi": Wfp_PartIndex = 6
        Case "Base": Wfp_PartIndex = 7
        Case "PP": Wfp_PartIndex = 8
        Case "PN": Wfp_PartIndex = 9
        Case "Val": Wfp_PartIndex = 10
        Case "Big": Wfp_PartIndex = 11
        Case Else: Wfp_PartIndex = -1
    End Select
End Function

' Coluna (no corpo do bloco) da parte do lado j (1 = serie 1, 2 = serie 2); 0 se o lado ou a parte nao existem.
Public Function Wfp_SideCol(j As Long, part As String) As Long
    Dim k As Long
    k = Wfp_PartIndex(part)
    If k >= 0 And (j = 1 Or j = 2) Then Wfp_SideCol = WFP_SIDE_BASE + WFP_SIDE_WIDTH * (j - 1) + k
End Function

' Numero de fatias (categorias do grafico) para n etapas: A e B por etapa e um vao entre etapas.
Public Function Wfp_SlotCount(n As Long) As Long
    If n > 0 Then Wfp_SlotCount = 3 * n - 1
End Function

' Fatia da barra do lado j (1 = A, 2 = B) na etapa i.
Public Function Wfp_SlotOf(i As Long, j As Long) As Long
    Wfp_SlotOf = 3 * (i - 1) + j
End Function

' Meia-largura da barra em unidades de fatia para o espaco (GapWidth, %): a barra ocupa 1 / (1 + gap/100) da fatia (medido). Pura.
Public Function Wfp_HalfWidth(gapPct As Double) As Double
    Wfp_HalfWidth = 0.5 / (1 + gapPct / 100)
End Function

' Numero de etapas do grafico, pela altura do bloco (meta + cabecalho + 6n). 0 sem bloco.
Public Function Wfp_StepCount(cht As Chart) As Long
    Dim blk As Range
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    Wfp_StepCount = (blk.Rows.Count - 2) \ WFP_ROWS_PER_PAIR
End Function

' j-esima serie NAO auxiliar: 1 = serie 1 (parte >= 0 de A), 2 = serie 2. Nothing se nao existe.
Public Function Wfp_SideSeries(cht As Chart, j As Long) As Series
    Dim s As Series, k As Long
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            k = k + 1
            If k = j Then Set Wfp_SideSeries = s: Exit Function
        End If
    Next
End Function

Private Function SR(body As Range, r As Long, j As Long, part As String) As String
    SR = body.Cells(r, Wfp_SideCol(j, part)).Address(False, False)
End Function

Private Function MetaRef(blk As Range, col As Long) As String
    MetaRef = blk.Cells(1, col).Address(True, True)
End Function

' Constante de matriz com os marcadores de total ({"e","t",...}), a mesma lista da cascata (WF_MARKERS).
Private Function MarkerList() As String
    MarkerList = "{""" & Replace(WF_MARKERS, ",", """,""") & """}"
End Function

' Calculo manual: recalcula o bloco (Range.Calculate segue as dependencias dentro da faixa). Em automatico nao faz nada.
Private Sub Wfp_CalcIfManual(blk As Range)
    If Application.Calculation <> xlCalculationAutomatic Then blk.Calculate
End Sub

' Grava uma celula de meta so quando o valor muda (cada gravacao recalcula e dispara o Calculate do grafico, que reentra no refresh).
Private Sub SetMeta(blk As Range, col As Long, v As Variant)
    Dim cur As Variant
    cur = blk.Cells(1, col).Value
    If VarType(v) = vbString Then
        If Not IsError(cur) Then
            If CStr(cur) = CStr(v) Then Exit Sub
        End If
    ElseIf Not IsError(cur) And Not IsEmpty(cur) And IsNumeric(cur) Then
        If Abs(CDbl(cur) - CDbl(v)) <= 0.000000001 * (1 + Abs(CDbl(v))) Then Exit Sub
    End If
    blk.Cells(1, col).Value = v
    Wfp_CalcIfManual blk
End Sub

Public Function Wfp_BuildBlock(wb As Workbook, id As String, data As clsChartData) As Range
    Dim blk As Range, body As Range, n As Long, ns As Long, i As Long, j As Long, k As Long, r As Long, st As Long
    Dim s As String, t As String, d As String, sta As String, en As String, lo As String, hi As String, bg As String, txt As String
    Dim a As String, b As String, ref1 As String, sA As String, sB As String, yH As String, lblOn As String, srcCat As String
    Dim ptv As String, hw As String, dv As String, dc As String, sf As String, sfd As String, ab As String, lb As String, tp As String, nOn As String, fit As String
    Dim parts As Variant, hdr As Variant, hiMax As Double, loMin As Double, v As Double
    Wfp_CheckData data
    n = data.CategoryCount
    ns = Wfp_SlotCount(n)
    Set blk = CalcSheet_AllocateBlock(wb, id, "waterfallpair", Mid$(ChartData_ExternalAddress(data.Source), 2), WFP_ROWS_PER_PAIR * n + 1, WFP_COLS)
    Set body = Block_Body(blk)
    body.Cells(1, WFP_COL_CAT).Value = "Cat"
    parts = Array("Src", "Tot", "D", "Start", "End", "Lo", "Hi", "Base", "PP", "PN", "Val", "Big")
    For j = 1 To 2
        For k = 0 To 11
            body.Cells(1, Wfp_SideCol(j, CStr(parts(k)))).Value = parts(k) & IIf(j = 1, "A", "B")
        Next
    Next
    hdr = Array("ArrowTxt", "TotTxtA", "TotTxtB", "Outer", "Inner", "SBase", "APos", "ANeg", "BPos", "BNeg", "LblAPos", "LblAVal", "LblBPos", "LblBVal", "TotPos", "TotTxt", "LblOn", "Top", "CxA", "CyA", "CLen", "CxB", "CyB", "PX", "PY", "PLbl", "RX", "RY", "NX", "NY", "NTxt", "ZX", "ZY")
    For k = 0 To UBound(hdr)
        body.Cells(1, WFP_COL_ARROWTXT + k).Value = hdr(k)
    Next
    ' meta: valores iniciais (as formulas abaixo as leem); os textos de sufixo em formato texto ("%" como numero viraria 0%)
    blk.Cells(1, WFP_META_SUF).NumberFormat = "@"
    blk.Cells(1, WFP_META_SUFD).NumberFormat = "@"
    blk.Cells(1, WFP_META_PTV).Value = 1
    blk.Cells(1, WFP_META_HW).Value = Wfp_HalfWidth(WFP_DEFAULT_GAP)
    blk.Cells(1, WFP_META_DIV).Value = 1
    blk.Cells(1, WFP_META_DEC).Value = 0
    blk.Cells(1, WFP_META_ABS).Value = 0
    blk.Cells(1, WFP_META_LBL).Value = 1
    blk.Cells(1, WFP_META_TOTPCT).Value = 1
    blk.Cells(1, WFP_META_NAMEON).Value = 1
    blk.Cells(1, WFP_META_FIT).Value = 1
    ptv = MetaRef(blk, WFP_META_PTV): hw = MetaRef(blk, WFP_META_HW): dv = MetaRef(blk, WFP_META_DIV): dc = MetaRef(blk, WFP_META_DEC)
    sf = MetaRef(blk, WFP_META_SUF): sfd = MetaRef(blk, WFP_META_SUFD): ab = MetaRef(blk, WFP_META_ABS): lb = MetaRef(blk, WFP_META_LBL)
    tp = MetaRef(blk, WFP_META_TOTPCT): nOn = MetaRef(blk, WFP_META_NAMEON): fit = MetaRef(blk, WFP_META_FIT)
    ' etapas: a ponte de cada serie, independente (marcadores por celula)
    For i = 1 To n
        r = i + 1
        srcCat = Mid$(ChartData_ExternalAddress(data.Source.Cells(i + 1, 1)), 2)
        body.Cells(r, WFP_COL_CAT).Formula = "=IF(ISBLANK(" & srcCat & "),"" ""," & srcCat & ")"
        body.Cells(r, WFP_COL_CAT).NumberFormat = data.Source.Cells(i + 1, 1).NumberFormat
        For j = 1 To 2
            s = SR(body, r, j, "Src"): t = SR(body, r, j, "Tot"): d = SR(body, r, j, "D"): sta = SR(body, r, j, "Start")
            en = SR(body, r, j, "End"): lo = SR(body, r, j, "Lo"): hi = SR(body, r, j, "Hi"): bg = SR(body, r, j, "Big")
            body.Cells(r, Wfp_SideCol(j, "Src")).Formula = ChartData_ExternalAddress(data.ValueCell(i, j))
            body.Cells(r, Wfp_SideCol(j, "Tot")).Formula = "=IF(ISERROR(" & s & "),FALSE,IF(ISNUMBER(" & s & "),FALSE,OR(LOWER(TRIM(" & s & "&""""))=" & MarkerList() & ")))"
            body.Cells(r, Wfp_SideCol(j, "D")).Formula = "=IF(" & t & ",0,IF(ISNUMBER(" & s & ")," & s & ",0))"
            If i = 1 Then
                body.Cells(r, Wfp_SideCol(j, "Start")).Formula = "=0"
            Else
                body.Cells(r, Wfp_SideCol(j, "Start")).Formula = "=" & SR(body, r - 1, j, "End")
            End If
            body.Cells(r, Wfp_SideCol(j, "End")).Formula = "=IF(" & t & "," & sta & "," & sta & "+" & d & ")"
            body.Cells(r, Wfp_SideCol(j, "Lo")).Formula = "=IF(" & t & ",MIN(0," & en & "),MIN(" & sta & "," & en & "))"
            body.Cells(r, Wfp_SideCol(j, "Hi")).Formula = "=IF(" & t & ",MAX(0," & en & "),MAX(" & sta & "," & en & "))"
            body.Cells(r, Wfp_SideCol(j, "Base")).Formula = "=MAX(0," & lo & ")+MIN(0," & hi & ")"
            body.Cells(r, Wfp_SideCol(j, "PP")).Formula = "=MAX(0," & hi & ")-MAX(0," & lo & ")"
            body.Cells(r, Wfp_SideCol(j, "PN")).Formula = "=MIN(0," & lo & ")-MIN(0," & hi & ")"
            If i = 1 Then
                body.Cells(r, Wfp_SideCol(j, "Big")).Formula = "=TRUE"
            Else
                body.Cells(r, Wfp_SideCol(j, "Big")).Formula = "=" & t
            End If
            body.Cells(r, Wfp_SideCol(j, "Val")).Formula = "=IF(" & bg & "," & en & "," & d & ")"
            ' texto das barras grandes: o valor (FIXED: separadores do Excel que calculou o bloco por ultimo) e, depois da 1a etapa, o percentual do 1o total da serie
            ref1 = body.Cells(2, Wfp_SideCol(j, "End")).Address(True, True)
            If i = 1 Then
                txt = "=IF(" & bg & ",FIXED(ROUND(" & en & "/" & dv & "," & dc & ")," & dc & ")&" & sf & ","""")"
            Else
                txt = "=IF(" & bg & ",FIXED(ROUND(" & en & "/" & dv & "," & dc & ")," & dc & ")&" & sf & "&IF(AND(" & tp & "=1," & ref1 & "<>0),"" (""&FIXED(ROUND(100*" & en & "/" & ref1 & ",0),0)&""%)"",""""),"""")"
            End If
            If j = 1 Then body.Cells(r, WFP_COL_TOTTXT_A).Formula = txt Else body.Cells(r, WFP_COL_TOTTXT_B).Formula = txt
        Next
        ' texto da seta: (B - A) / A em % inteiro com sinal; "n/d" com A = 0 ou sinais opostos; com Abs = 1, B - A no formato do grafico
        a = SR(body, r, 1, "Val"): b = SR(body, r, 2, "Val")
        body.Cells(r, WFP_COL_ARROWTXT).Formula = "=IF(" & ab & "=1,IF(ROUND((" & b & "-" & a & ")/" & dv & "," & dc & ")>0,""+"","""")&FIXED(ROUND((" & b & "-" & a & ")/" & dv & "," & dc & ")," & dc & ")&" & sfd & ",IF(OR(" & a & "=0," & a & "*" & b & "<0),""n/d"",IF(ROUND(100*(" & b & "-" & a & ")/" & a & ",0)>0,""+"","""")&FIXED(ROUND(100*(" & b & "-" & a & ")/" & a & ",0),0)&""%""))"
    Next
    ' fatias: o que o grafico plota
    For i = 1 To n
        st = i + 1
        For j = 1 To 2
            k = Wfp_SlotOf(i, j): r = k + 1
            bg = SR(body, st, j, "Big"): lo = SR(body, st, j, "Lo"): hi = SR(body, st, j, "Hi"): en = SR(body, st, j, "End"): d = SR(body, st, j, "D")
            lblOn = body.Cells(r, WFP_COL_LBLON).Address(False, False)
            If j = 1 Then
                body.Cells(r, WFP_COL_OUTER).Formula = "=" & body.Cells(st, WFP_COL_CAT).Address(False, False)
                body.Cells(r, WFP_COL_OUTER).NumberFormat = data.Source.Cells(i + 1, 1).NumberFormat
            End If
            body.Cells(r, WFP_COL_INNER).Value = " "
            body.Cells(r, WFP_COL_SBASE).Formula = "=" & SR(body, st, j, "Base")
            If j = 1 Then
                body.Cells(r, WFP_COL_APOS).Formula = "=" & SR(body, st, 1, "PP")
                body.Cells(r, WFP_COL_ANEG).Formula = "=" & SR(body, st, 1, "PN")
                body.Cells(r, WFP_COL_BPOS).Formula = "=0"
                body.Cells(r, WFP_COL_BNEG).Formula = "=0"
                body.Cells(r, WFP_COL_LBLAPOS).Formula = "=IF(" & bg & ",NA(),(" & lo & "+" & hi & ")/2)"
                body.Cells(r, WFP_COL_LBLAVAL).Formula = "=IF(OR(" & bg & "," & lblOn & "=0),""""," & d & ")"
                body.Cells(r, WFP_COL_LBLBPOS).Formula = "=NA()"
                body.Cells(r, WFP_COL_LBLBVAL).Formula = "="""""
                body.Cells(r, WFP_COL_TOTTXT).Formula = "=" & body.Cells(st, WFP_COL_TOTTXT_A).Address(False, False)
            Else
                body.Cells(r, WFP_COL_APOS).Formula = "=0"
                body.Cells(r, WFP_COL_ANEG).Formula = "=0"
                body.Cells(r, WFP_COL_BPOS).Formula = "=" & SR(body, st, 2, "PP")
                body.Cells(r, WFP_COL_BNEG).Formula = "=" & SR(body, st, 2, "PN")
                body.Cells(r, WFP_COL_LBLAPOS).Formula = "=NA()"
                body.Cells(r, WFP_COL_LBLAVAL).Formula = "="""""
                body.Cells(r, WFP_COL_LBLBPOS).Formula = "=IF(" & bg & ",NA(),(" & lo & "+" & hi & ")/2)"
                body.Cells(r, WFP_COL_LBLBVAL).Formula = "=IF(OR(" & bg & "," & lblOn & "=0),""""," & d & ")"
                body.Cells(r, WFP_COL_TOTTXT).Formula = "=" & body.Cells(st, WFP_COL_TOTTXT_B).Address(False, False)
            End If
            body.Cells(r, WFP_COL_TOTPOS).Formula = "=IF(" & bg & ",IF(" & en & ">=0," & hi & "," & lo & "),NA())"
            body.Cells(r, WFP_COL_LBLON).Value = 1
            body.Cells(r, WFP_COL_TOP).Formula = "=" & hi
        Next
        If i < n Then
            r = 3 * i + 1                                  ' vao da etapa i (fatia 3i)
            body.Cells(r, WFP_COL_OUTER).Value = " "
            body.Cells(r, WFP_COL_INNER).Value = " "
            body.Cells(r, WFP_COL_SBASE).Resize(1, 5).Formula = "=0"
            body.Cells(r, WFP_COL_LBLAPOS).Formula = "=NA()"
            body.Cells(r, WFP_COL_LBLAVAL).Formula = "="""""
            body.Cells(r, WFP_COL_LBLBPOS).Formula = "=NA()"
            body.Cells(r, WFP_COL_LBLBVAL).Formula = "="""""
            body.Cells(r, WFP_COL_TOTPOS).Formula = "=NA()"
            body.Cells(r, WFP_COL_TOTTXT).Formula = "="""""
            body.Cells(r, WFP_COL_LBLON).Value = 1
            body.Cells(r, WFP_COL_TOP).Formula = "=0"
        End If
    Next
    ' conectores: do lado direito da barra i ao lado esquerdo da barra i + 1 da MESMA serie, na altura do fim da barra i
    For i = 1 To n - 1
        r = i + 1
        body.Cells(r, WFP_COL_CXA).Formula = "=" & (3 * i - 2) & "+" & hw
        body.Cells(r, WFP_COL_CYA).Formula = "=" & SR(body, r, 1, "End")
        body.Cells(r, WFP_COL_CLEN).Formula = "=3-2*" & hw
        body.Cells(r, WFP_COL_CXB).Formula = "=" & (3 * i - 1) & "+" & hw
        body.Cells(r, WFP_COL_CYB).Formula = "=" & SR(body, r, 2, "End")
    Next
    ' setas: 5 pontos por par (sobe de A, horizontal, desce em B) e uma linha vazia; o afastamento em pontos vezes PtV (valor por ponto)
    For i = 1 To n
        st = i + 1
        sA = "(" & SR(body, st, 1, "Hi") & "+" & ptv & "*IF(AND(" & SR(body, st, 1, "Big") & "," & lb & "=1," & SR(body, st, 1, "End") & ">=0)," & WFP_PT_LBLROOM & "," & WFP_PT_GAP & "))"
        sB = "(" & SR(body, st, 2, "Hi") & "+" & ptv & "*IF(AND(" & SR(body, st, 2, "Big") & "," & lb & "=1," & SR(body, st, 2, "End") & ">=0)," & WFP_PT_LBLROOM & "," & WFP_PT_GAP & "))"
        yH = "(MAX(" & sA & "," & sB & ")+" & ptv & "*" & WFP_PT_RISE & ")"
        r = WFP_ROWS_PER_PAIR * (i - 1) + 2
        body.Cells(r, WFP_COL_PX).Formula = "=" & (3 * i - 2): body.Cells(r, WFP_COL_PY).Formula = "=" & sA
        body.Cells(r + 1, WFP_COL_PX).Formula = "=" & (3 * i - 2): body.Cells(r + 1, WFP_COL_PY).Formula = "=" & yH
        body.Cells(r + 2, WFP_COL_PX).Formula = "=" & (6 * i - 3) & "/2": body.Cells(r + 2, WFP_COL_PY).Formula = "=" & yH
        body.Cells(r + 2, WFP_COL_PLBL).Formula = "=" & body.Cells(st, WFP_COL_ARROWTXT).Address(False, False)
        body.Cells(r + 3, WFP_COL_PX).Formula = "=" & (3 * i - 1): body.Cells(r + 3, WFP_COL_PY).Formula = "=" & yH
        body.Cells(r + 4, WFP_COL_PX).Formula = "=" & (3 * i - 1): body.Cells(r + 4, WFP_COL_PY).Formula = "=" & sB
        body.Cells(st, WFP_COL_RX).Formula = "=" & (6 * i - 3) & "/2"
        body.Cells(st, WFP_COL_RY).Formula = "=" & yH & "+" & ptv & "*" & WFP_PT_ROOM
    Next
    ' nomes das series: meio da 1a barra de cada lado; texto so com NameOn = 1 e Fit = 1
    For j = 1 To 2
        r = j + 1
        body.Cells(r, WFP_COL_NX).Formula = "=" & j
        body.Cells(r, WFP_COL_NY).Formula = "=(" & SR(body, 2, j, "Lo") & "+" & SR(body, 2, j, "Hi") & ")/2"
        body.Cells(r, WFP_COL_NTXT).Formula = "=IF(AND(" & nOn & "=1," & fit & "=1)," & Mid$(ChartData_ExternalAddress(data.SeriesNameCell(j)), 2) & "&"""","""")"
    Next
    ' linha do zero: de 0,5 a 3n - 0,5 (bordas da area interna)
    body.Cells(2, WFP_COL_ZX).Formula = "=1/2": body.Cells(2, WFP_COL_ZY).Formula = "=0"
    body.Cells(3, WFP_COL_ZX).Formula = "=" & (2 * ns + 1) & "/2": body.Cells(3, WFP_COL_ZY).Formula = "=0"
    ' PtV inicial: a faixa dos dados em 180 pt (o refresh calibra pelo eixo real)
    Wfp_CalcIfManual blk
    hiMax = Application.WorksheetFunction.Max(body.Cells(2, WFP_COL_TOP).Resize(ns, 1))
    For i = 1 To n
        For j = 1 To 2
            v = Util_ToDouble(body.Cells(i + 1, Wfp_SideCol(j, "Lo")).Value)
            If v < loMin Then loMin = v
        Next
    Next
    v = (hiMax - loMin) / 180
    If v <= 0 Then v = 1
    blk.Cells(1, WFP_META_PTV).Value = v
    Wfp_CalcIfManual blk
    Set Wfp_BuildBlock = blk
End Function

Private Function AddBar(cht As Chart, name As String, valueCells As Range, cats As String) As Series
    Dim s As Series
    Set s = cht.SeriesCollection.NewSeries
    s.Name = name
    s.Values = Block_Ref(valueCells)
    s.XValues = cats
    Set AddBar = s
End Function

' Auxiliar XY no grupo PRINCIPAL (x = numero da fatia, medido), sem marcador; withLine: linha ligando os pontos (setas, zero). Em falha desfaz.
Private Function AddXY(cht As Chart, name As String, xCells As Range, yCells As Range, withLine As Boolean) As Series
    Dim s As Series, en As Long, es As String, ed As String
    On Error GoTo Fail
    Set s = cht.SeriesCollection.NewSeries
    s.Name = name
    If withLine Then s.ChartType = xlXYScatterLinesNoMarkers Else s.ChartType = xlXYScatter     ' antes de XValues: XY aceita X numerico
    s.XValues = Block_Ref(xCells)
    s.Values = Block_Ref(yCells)
    s.MarkerStyle = xlMarkerStyleNone
    If Not withLine Then s.Format.Line.Visible = msoFalse
    Set AddXY = s
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not s Is Nothing Then s.Delete
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

Private Sub Wfp_StyleCategoryAxis(cht As Chart)
    With cht.Axes(xlCategory)
        .TickLabelPosition = xlTickLabelPositionLow
        .TickLabels.MultiLevel = True
        .TickLabels.Orientation = xlHorizontal
        .Format.Line.Visible = msoFalse                ' a linha do eixo desenharia os separadores do eixo de 2 niveis (medido)
    End With
End Sub

Private Sub Wfp_AddZeroLine(cht As Chart)
    Dim body As Range, s As Series
    If Not Chart_FindSeries(cht, "tqq_zero") Is Nothing Then Exit Sub
    Set body = Block_Body(Registry_Block(cht))
    Set s = AddXY(cht, "tqq_zero", body.Cells(2, WFP_COL_ZX).Resize(2, 1), body.Cells(2, WFP_COL_ZY).Resize(2, 1), True)
    s.Format.Line.Visible = msoTrue
    s.Format.Line.ForeColor.RGB = RGB(166, 166, 166)
    s.Format.Line.Weight = 0.75
End Sub

' Nomes das series: auxiliar XY tqq_names (meio da 1a barra de cada lado) com o rotulo ligado a NTxt, vertical, 9 pt, fonte por contraste.
Private Sub Wfp_SetNames(cht As Chart)
    Dim body As Range, s As Series, k As Long
    Set body = Block_Body(Registry_Block(cht))
    Set s = Chart_FindSeries(cht, "tqq_names")
    If s Is Nothing Then Set s = AddXY(cht, "tqq_names", body.Cells(2, WFP_COL_NX).Resize(2, 1), body.Cells(2, WFP_COL_NY).Resize(2, 1), False)
    For k = 1 To 2
        With s.Points(k)
            .HasDataLabel = True
            .DataLabel.Formula = Block_Ref(body.Cells(k + 1, WFP_COL_NTXT))
            .DataLabel.Position = xlLabelPositionCenter
            .DataLabel.Orientation = xlUpward
            .DataLabel.Font.Size = 9
            .DataLabel.Font.Bold = False
            .DataLabel.Font.Color = Style_ContrastFont(Style_DataSeriesColor(cht, k))
        End With
    Next
End Sub

Public Function Wfp_Insert(ws As Worksheet, data As clsChartData) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, anchor As Range, blk As Range, body As Range, n As Long, ns As Long, cats As String
    Dim en As Long, es As String, ed As String
    Wfp_CheckData data                               ' recusa (6016/6017) antes de alocar: nada a desfazer
    On Error GoTo Fail
    id = Registry_NewId()
    Set blk = Wfp_BuildBlock(ws.Parent, id, data)
    Util_TestHookAfterBlock
    Set body = Block_Body(blk)
    n = data.CategoryCount
    ns = Wfp_SlotCount(n)
    cats = Block_Ref(body.Cells(2, WFP_COL_OUTER).Resize(ns, 2))     ' 2 colunas: eixo de categorias de 2 niveis
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    AddBar cht, "tqq_base", body.Cells(2, WFP_COL_SBASE).Resize(ns, 1), cats
    AddBar cht, ChartData_ExternalAddress(data.SeriesNameCell(1)), body.Cells(2, WFP_COL_APOS).Resize(ns, 1), cats
    AddBar cht, "tqq_aneg", body.Cells(2, WFP_COL_ANEG).Resize(ns, 1), cats
    AddBar cht, ChartData_ExternalAddress(data.SeriesNameCell(2)), body.Cells(2, WFP_COL_BPOS).Resize(ns, 1), cats
    AddBar cht, "tqq_bneg", body.Cells(2, WFP_COL_BNEG).Resize(ns, 1), cats
    cht.ChartType = xlColumnStacked
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlNotPlotted               ' as linhas vazias do bloco separam as setas; as colunas sao formulas (nunca vazias)
    Style_ApplyBase cht
    Util_TestHookAfterChart
    cht.ChartGroups(1).GapWidth = WFP_DEFAULT_GAP
    Wfp_StyleCategoryAxis cht
    cht.HasLegend = False
    Registry_SetOption cht, "numfmt", "int"
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "gap", CStr(WFP_DEFAULT_GAP)
    Registry_SetOption cht, "n", CStr(n)
    Registry_SetOption cht, "totals", "1"
    Registry_SetOption cht, "legend", "inbar"
    Wfp_ApplyColors cht, "padrao"
    Wfp_AddZeroLine cht
    Wfp_SetNames cht
    Wfp_SetLegend cht, "inbar"
    Wfp_RefreshAfterDataChange cht
    Set Wfp_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' Cores: serie 1 = pal(0) e serie 2 = pal(2) (o contraste escuro/claro da imagem do think-cell), as partes negativas com a cor da serie, a base
' invisivel; a fonte dos rotulos e dos nomes acompanha (contraste).
Public Sub Wfp_ApplyColors(cht As Chart, paletteName As String)
    Dim pal As Variant, s As Series, k As Long, c As Variant
    pal = Style_Palette(paletteName)
    For Each s In cht.SeriesCollection
        c = Empty
        Select Case s.Name
            Case "tqq_base"
                s.Format.Fill.Visible = msoFalse
                s.Format.Line.Visible = msoFalse
            Case "tqq_aneg"
                c = pal(0)
            Case "tqq_bneg"
                c = pal(2)
            Case Else
                If Not Style_IsHelperSeries(s) Then
                    k = k + 1
                    If k = 1 Then c = pal(0) Else c = pal(2)
                End If
        End Select
        If Not IsEmpty(c) Then
            s.Format.Fill.Visible = msoTrue
            Style_SetColor s.Format.Fill.ForeColor, c
            s.Format.Line.Visible = msoFalse
        End If
    Next
    Wfp_RefreshFontColors cht
End Sub

Private Function PointHasLabel(s As Series, i As Long) As Boolean
    On Error Resume Next
    PointHasLabel = s.Points(i).HasDataLabel
End Function

' Fonte por contraste com a cor atual de cada serie: rotulos de degrau (ponto a ponto: uma escrita no nivel da serie recriaria rotulos apagados)
' e o nome de cada serie.
Private Sub Wfp_RefreshFontColors(cht As Chart)
    Dim s As Series, k As Long, i As Long, fc As Long
    For k = 1 To 2
        fc = Style_ContrastFont(Style_DataSeriesColor(cht, k))
        If k = 1 Then Set s = Chart_FindSeries(cht, "tqq_lbl_a") Else Set s = Chart_FindSeries(cht, "tqq_lbl_b")
        If Not s Is Nothing Then
            For i = 1 To s.Points.Count
                If PointHasLabel(s, i) Then s.Points(i).DataLabel.Font.Color = fc
            Next
        End If
        Set s = Chart_FindSeries(cht, "tqq_names")
        If Not s Is Nothing Then
            If s.Points.Count >= k Then
                If PointHasLabel(s, k) Then s.Points(k).DataLabel.Font.Color = fc
            End If
        End If
    Next
End Sub

Private Sub ShowLegend(cht As Chart, turnOn As Boolean)
    If turnOn Then
        If Not cht.HasLegend Then cht.HasLegend = True
        cht.Legend.Position = xlLegendPositionRight
        cht.Legend.Font.Size = 9
        Chart_HideHelperLegendEntries cht
    ElseIf cht.HasLegend Then
        cht.HasLegend = False
    End If
End Sub

' Legenda do tipo: right (legenda a direita, sem nomes nas barras), none (nada), inbar (nomes verticais na 1a barra; se algum nao cabe, os dois
' somem e a legenda aparece: Wfp_SyncNames). Quem chama grava a opcao "legend" antes (Wfp_SyncNames a le).
Public Sub Wfp_SetLegend(cht As Chart, mode As String)
    Dim blk As Range
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    SetMeta blk, WFP_META_NAMEON, IIf(mode = "inbar", 1, 0)
    Select Case mode
        Case "right": ShowLegend cht, True
        Case "none": ShowLegend cht, False
        Case Else: Wfp_SyncNames cht
    End Select
End Sub

' "Nome dentro das barras": o nome cabe se a 1a barra da serie tem altura >= largura estimada do nome + 8 pt e largura >= 14 pt (texto
' vertical de 9 pt, medido). Se algum nao cabe, Fit = 0 (os dois nomes ficam vazios) e a legenda aparece; voltando a caber, o contrario (a
' legenda so muda quando o estado dela nao corresponde: legenda visivel = algum nome nao cabe). So com a opcao inbar; grava so o que muda.
Public Sub Wfp_SyncNames(cht As Chart)
    Dim blk As Range, body As Range, ax As Axis, ppu As Double, barW As Double, j As Long, h As Double, need As Double, fits As Boolean, s As Series, ns As Long
    If Registry_GetOption(cht, "legend", "inbar") <> "inbar" Then Exit Sub
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    Set body = Block_Body(blk)
    ns = Wfp_SlotCount(Wfp_StepCount(cht))
    Set ax = cht.Axes(xlValue, xlPrimary)
    If ns < 1 Or ax.MaximumScale <= ax.MinimumScale Then Exit Sub
    ppu = cht.PlotArea.InsideHeight / (ax.MaximumScale - ax.MinimumScale)
    barW = cht.PlotArea.InsideWidth / ns / (1 + cht.ChartGroups(1).GapWidth / 100)
    fits = (barW >= WFP_NAME_MIN_W)
    For j = 1 To 2
        Set s = Wfp_SideSeries(cht, j)
        h = (Util_ToDouble(body.Cells(2, Wfp_SideCol(j, "Hi")).Value) - Util_ToDouble(body.Cells(2, Wfp_SideCol(j, "Lo")).Value)) * ppu
        need = WFP_NAME_MARGIN
        If Not s Is Nothing Then need = need + Labels_EstimateWidth(s.Name)
        If h < need Then fits = False
    Next
    SetMeta blk, WFP_META_NAMEON, 1
    SetMeta blk, WFP_META_FIT, IIf(fits, 1, 0)
    If cht.HasLegend = fits Then ShowLegend cht, Not fits
End Sub

' Meia-largura das barras (meta HW) a partir do espaco REAL do grupo de colunas (mudado pela faixa ou a mao no Excel). Grava so se mudou.
Public Sub Wfp_SyncGap(cht As Chart)
    Dim blk As Range
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    SetMeta blk, WFP_META_HW, Wfp_HalfWidth(CDbl(cht.ChartGroups(1).GapWidth))
End Sub

' O que depende dos DADOS, do tamanho ou do espaco e muda depois de inserir (Core_SafeRefresh): meia-largura, nomes/legenda. Guarda de
' reentrada (as gravacoes disparam o Calculate do grafico). Nunca propaga erro.
Public Sub Wfp_RefreshAfterDataChange(cht As Chart)
    On Error GoTo Done
    If mBusy Then Exit Sub
    If Kinds_Family(Registry_ChartType(cht)) <> "waterfallpair" Then Exit Sub
    mBusy = True
    Wfp_SyncGap cht
    Wfp_SyncNames cht
Done:
    If Err.Number <> 0 Then Errors_Log "Wfp_RefreshAfterDataChange" & vbTab & Err.Number & vbTab & Err.Description
    mBusy = False
End Sub

' Linhas da tabela de origem entre a 1a e a ultima etapa (pelas referencias Src do lado 1 no bloco): linhas inseridas/apagadas no meio da
' tabela mudam esse numero (as series apontam para o bloco, como no Mekko). -1 se nao for possivel ler.
Public Function Wfp_SourceRowsSpanned(cht As Chart) As Long
    Dim body As Range, n As Long, r1 As Long, r2 As Long
    On Error GoTo Fail
    Wfp_SourceRowsSpanned = -1
    n = Wfp_StepCount(cht)
    If n < 1 Then Exit Function
    Set body = Block_Body(Registry_Block(cht))
    r1 = RefRow(CStr(body.Cells(2, Wfp_SideCol(1, "Src")).Formula))
    r2 = RefRow(CStr(body.Cells(n + 1, Wfp_SideCol(1, "Src")).Formula))
    If r1 < 1 Or r2 < r1 Then Exit Function
    Wfp_SourceRowsSpanned = r2 - r1 + 1
    Exit Function
Fail:
    Wfp_SourceRowsSpanned = -1
End Function

' Linha da referencia ='Planilha'!$B$7 (o numero depois do ultimo "$"); 0 sem referencia (#REF!).
Private Function RefRow(f As String) As Long
    Dim p As Long
    If InStr(f, "#REF!") > 0 Then Exit Function
    p = InStrRev(f, "$")
    If p = 0 Then Exit Function
    RefRow = CLng(Val(Mid$(f, p + 1)))
End Function
```

- [ ] **Passo 4: Ligações** —
  - `modChartKinds.Kinds_Insert`, antes do `Case Else`: `Case "waterfallpair": Set Kinds_Insert = Wfp_Insert(ws, data)`.
  - `modCoreApi`: `PALETTE_FAMILIES = ",stacked,combo,waterfall,area,pie,mekko,gantt,butterfly,scatter,waterfallpair,"`; no despacho da paleta de `ApplyOption`, antes do `Case "stacked", "combo"`: `Case "waterfallpair": Wfp_ApplyColors c, value              ' cascata comparativa: serie 1 pal(0), serie 2 pal(2), contraste dos rotulos`; em `ApplyOption`, troque `Case "gap": ApplyGapWidth c, CLng(Val(value))` por
    ```vb
            Case "gap"
                ApplyGapWidth c, CLng(Val(value))
                If fam = "waterfallpair" Then Wfp_SyncGap c         ' cascata comparativa: meia-largura das barras (conectores) acompanha o espaco
    ```
  - `modCoreApi.ValidateOption`, `Case "legend"`: `Case "right", "none", "lineend", "inbar"`.
  - `modCoreApi.Core_SafeRefresh`, depois da linha do `Bfly_RefreshAfterDataChange` (e do seu log): 
    ```vb
        Wfp_RefreshAfterDataChange cht                   ' cascata comparativa: meia-largura das barras, valor por ponto das setas, rotulos finos, nomes ou legenda
        If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Wfp" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    ```
  - `modCoreApi.Core_ChartSourceMismatch`, dentro do `If Not Core_ChartSourceMismatch And Len(nOpt) > 0 Then`, depois da linha da dispersão:
    ```vb
        If Kinds_Family(Registry_ChartType(cht)) = "waterfallpair" Then Core_ChartSourceMismatch = (Wfp_SourceRowsSpanned(cht) <> CLng(Val(nOpt)))
    ```
  - `modDecorations.Deco_SetLegend`: no `Select Case mode` inicial, `Case "right", "none", "lineend", "inbar"`; logo depois do bloco do `lineend`, acrescente
    ```vb
        If mode = "inbar" Then
            If Not Kinds_Supports(kind, "inbar") Then Err.Raise vbObjectError + 7002, "Deco_SetLegend", Kinds_NotApplicableText("inbar")
        End If
    ```
    e, logo depois do bloco `If Kinds_Family(kind) = "pie" Then ... End If`,
    ```vb
        If Kinds_Family(kind) = "waterfallpair" Then
            ' Cascata comparativa: Direita (legenda), Nenhuma, Nome dentro das barras (nomes verticais na 1a barra; a legenda se nao couberem).
            ' A opcao e gravada antes: Wfp_SyncNames a le.
            Registry_SetOption cht, "legend", mode
            Wfp_SetLegend cht, mode
            On Error Resume Next
            cht.PlotArea.Position = xlChartElementPositionAutomatic
            If Err.Number <> 0 Then Errors_Log "Deco_SetLegend/PlotArea.Position" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
            On Error GoTo 0
            Chart_EnsureRoomAfterChange cht
            Exit Sub
        End If
    ```
  - `modAnnotations.Ann_GetCategoryData`, antes do `Case Else`:
    ```vb
            Case "waterfallpair"
                ' cascata comparativa: o topo (Hi) de cada fatia, lido do bloco (Deco_ScaleClipped compara com o maximo fixo da Mesma escala)
                Set bb = Block_Body(Registry_Block(cht))
                For i = 1 To n
                    If i <= bb.Rows.Count - 1 Then
                        values(i) = Util_ToDouble(bb.Cells(i + 1, WFP_COL_TOP).Value)
                        tops(i) = values(i)
                    End If
                Next
    ```

- [ ] **Passo 5: GREEN e suíte completa** — filtro `TestWaterfallPair` verde (11 testes); `pwsh C:\temp\tqq-cc\build\run-tests.ps1 -TimeoutSec 900` verde. Exporte o gráfico da fixture (`cht.Export` para uma pasta sua em `C:\temp\tqq-cc-png`, gráfico perto de A1) e olhe: duas pontes lado a lado, nomes das etapas centrados sob cada par, linha do zero, nomes "Segmento 1"/"Segmento 2" na vertical dentro das barras de Bruto; registre no relatório e apague a pasta.

- [ ] **Passo 6: Commit** — `git add src/core/modules/modChartWaterfallPair.bas src/core/modules/modChartKinds.bas src/core/modules/modCoreApi.bas src/core/modules/modDecorations.bas src/core/modules/modAnnotations.bas tests/vba/core/TestWaterfallPair.bas tests/vba/core/TestCoreApi.bas` e `git commit -m "feat: cascata comparativa (duas pontes lado a lado, eixo de 2 niveis, nomes dentro das barras ou legenda)"` (+ trailer).

**Verificação no Excel real:** tabela Etapa/Segmento 1/Segmento 2 da imagem; pela janela Imediata do VBE (a faixa vem na Tarefa 5): `Application.Run "'ThinQQ.core.xlam'!Core_Insert", "waterfallpair"` com a tabela selecionada; mudar um valor move as barras; Legenda › Direita/Nenhuma (a faixa já tem esses itens) troca os nomes pela legenda.

---
### Tarefa 2: Conectores por série (auxiliares XY com barra de erro, vivos) — modelo: standard

**Files:**
- Modify: `src/core/modules/modChartWaterfallPair.bas` (`Wfp_SetConnectors`; uma linha em `Wfp_Insert`), `src/core/modules/modCoreApi.bas` (`ApplyToggle`, `Case "connectors"`)
- Test: `tests/vba/core/TestWaterfallPair.bas`

**Interfaces:**
- Consumes (Tarefa 1): `AddXY`, `Wfp_StepCount`, `Wfp_SlotOf`, `WFP_COL_CXA/CYA/CLEN/CXB/CYB`, `WFP_META_HW`, `Wfp_SyncGap` (já chamada no refresh e em `ApplyOption "gap"`).
- Produces: `Wfp_SetConnectors(cht, turnOn)` — apaga e recria `tqq_conn_a`/`tqq_conn_b` (n − 1 pontos cada, barra de erro X custom `CLen` passada em R1C1, sem ponta, cinza escuro RGB(89,89,89), tracejada `msoLineDash`, 0,75 pt), grava `connectors`; com uma etapa só não cria séries (a opção fica gravada). `Wfp_Insert` passa a chamar `Wfp_SetConnectors cht, True` logo depois de `Wfp_SetNames cht`.

- [ ] **Passo 1: Testes (RED)** — acrescente a `tests/vba/core/TestWaterfallPair.bas`:

```vb
' Review Focus 5 (conectores): do lado direito da barra i ao lado esquerdo da barra i + 1 da MESMA serie, na altura do fim da barra i; vivos.
Public Sub Test_Wfp_Connectors_GeometryAndLive()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, body As Range, ca As Series, cb As Series, cs As Series
    Dim i As Long, j As Long, sIdx As Long, x As Double, clen As Double, hw As Double, p As Point, q As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    Set ca = Chart_FindSeries(cht, "tqq_conn_a"): Set cb = Chart_FindSeries(cht, "tqq_conn_b")
    AssertTrue Not ca Is Nothing, "conector da serie 1": AssertTrue Not cb Is Nothing, "conector da serie 2"
    AssertEqual "1", Registry_GetOption(cht, "connectors")
    AssertEqual 4, ca.Points.Count, "n - 1 conectores": AssertTrue ca.HasErrorBars
    AssertEqual xlNoCap, ca.ErrorBars.EndStyle: AssertEqual msoLineDash, ca.ErrorBars.Format.Line.DashStyle
    AssertEqual xlPrimary, ca.AxisGroup
    Set body = BodyOf(cht)
    hw = Util_ToDouble(Registry_Block(cht).Cells(1, WFP_META_HW).Value)
    clen = Util_ToDouble(body.Cells(2, WFP_COL_CLEN).Value)
    AssertAlmost 3 - 2 * hw, clen, 0.000001
    ForceLayout cht
    For i = 1 To 4
        For j = 1 To 2
            If j = 1 Then
                Set cs = ca: sIdx = 2: x = Util_ToDouble(body.Cells(i + 1, WFP_COL_CXA).Value)
            Else
                Set cs = cb: sIdx = 4: x = Util_ToDouble(body.Cells(i + 1, WFP_COL_CXB).Value)
            End If
            AssertAlmost Wfp_SlotOf(i, j) + hw, x, 0.000001, "x do conector " & i & "/" & j
            Set p = cht.SeriesCollection(sIdx).Points(Wfp_SlotOf(i, j))
            Set q = cht.SeriesCollection(sIdx).Points(Wfp_SlotOf(i + 1, j))
            AssertAlmost p.Left + p.Width, XOfX(cht, x), 1, "sai da borda direita da barra " & i & "/" & j
            AssertAlmost q.Left, XOfX(cht, x + clen), 1, "chega a borda esquerda da proxima barra da mesma serie " & i & "/" & j
            AssertAlmost YOf(cht, SideVal(cht, i, j, "End")), cs.Points(i).Top, 1, "na altura do fim da barra " & i & "/" & j
        Next
    Next
    ws.Range("B3").Value = -9
    cht.Refresh
    AssertEqual 18, Util_ToDouble(body.Cells(3, WFP_COL_CYA).Value), "sem o ThinQQ o conector acompanha o dado"
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Connectors_ToggleGapLegendAndSingleStep()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, c1 As Chart, body As Range, x As Double, clen As Double, p As Point, q As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    Core_SetToggle "connectors", False
    AssertTrue Chart_FindSeries(cht, "tqq_conn_a") Is Nothing: AssertTrue Chart_FindSeries(cht, "tqq_conn_b") Is Nothing
    AssertEqual "0", Registry_GetOption(cht, "connectors"): AssertFalse Core_GetToggle("connectors")
    Core_SetToggle "connectors", True
    AssertEqual 4, Chart_FindSeries(cht, "tqq_conn_a").Points.Count
    Core_SetOption "gap", "80"
    ForceLayout cht
    Set body = BodyOf(cht)
    x = Util_ToDouble(body.Cells(2, WFP_COL_CXA).Value): clen = Util_ToDouble(body.Cells(2, WFP_COL_CLEN).Value)
    Set p = cht.SeriesCollection(2).Points(1): Set q = cht.SeriesCollection(2).Points(4)
    AssertAlmost p.Left + p.Width, XOfX(cht, x), 1, "espaco 80: o conector continua na borda da barra"
    AssertAlmost q.Left, XOfX(cht, x + clen), 1
    Core_SetOption "legend", "right"
    AssertEqual 2, cht.Legend.LegendEntries.Count, "os conectores ficam fora da legenda"
    ws.Range("E1:G1").Value = Array("Etapa", "S1", "S2"): ws.Range("E2:G2").Value = Array("Unico", 5, 7)
    Set c1 = Wfp_Insert(ws, ChartData_FromRange(ws.Range("E1:G2"))).Chart
    AssertEqual 2, SlotN(c1), "uma etapa: 2 fatias, sem vao"
    AssertTrue Chart_FindSeries(c1, "tqq_conn_a") Is Nothing, "uma etapa: sem conectores"
    AssertEqual "1", Registry_GetOption(c1, "connectors"), "a opcao fica gravada"
    Teardown wb
End Sub

' Copia do grafico (como colada: outro nome) com o original apagado: a limpeza nao remove o bloco (colunas, XY e rotulos citam o bloco).
Public Sub Test_Wfp_CopiedChart_KeepsBlockAlive()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, dup As Object
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Wfp_Insert ws, d
    Set dup = ws.ChartObjects(1).Duplicate
    ws.ChartObjects(2).Name = "Gráfico 2"
    ws.ChartObjects(1).Delete
    AssertEqual 0, CalcSheet_Cleanup(wb), "a copia cita o bloco: nada e removido"
    AssertFalse CalcSheet_CleanupReadFailed()
    AssertEqual 1, TqqNameCount(wb)
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-cc\build\run-tests.ps1 -Target core -Filter "Wfp_Connectors" -TimeoutSec 400` → falha (`tqq_conn_a` inexistente; o alternador ainda é recusado pelo despacho com 7002). O `Test_Wfp_CopiedChart_KeepsBlockAlive` não depende de código novo (regressão da limpeza com as séries do tipo): roda no GREEN do Passo 4.

- [ ] **Passo 3: Implementar** — acrescente a `modChartWaterfallPair.bas` (depois de `Wfp_SetNames`):

```vb
' Conectores: duas auxiliares XY sem marcador (tqq_conn_a, tqq_conn_b), um ponto por conector (n - 1) no lado direito da barra i, com barra de
' erro X custom (CLen do bloco) ate o lado esquerdo da barra i + 1 da MESMA serie, tracejada e sem ponta (medido; vivos sem o ThinQQ). Passam
' sobre a barra da outra serie (como no think-cell). Uma etapa so: sem conectores (a opcao fica gravada).
' Faixa da barra de erro em R1C1 (a forma que o Excel aceita para valores personalizados; em A1 pode recusar com 1004 ou virar constante).
Public Sub Wfp_SetConnectors(cht As Chart, turnOn As Boolean)
    Dim body As Range, n As Long, j As Long, s As Series, nm As String
    For j = 1 To 2
        nm = "tqq_conn_" & Mid$("ab", j, 1)
        Set s = Chart_FindSeries(cht, nm)
        If Not s Is Nothing Then s.Delete
    Next
    Registry_SetOption cht, "connectors", IIf(turnOn, "1", "0")
    If Not turnOn Then Exit Sub
    n = Wfp_StepCount(cht)
    If n < 2 Then Exit Sub
    Set body = Block_Body(Registry_Block(cht))
    For j = 1 To 2
        nm = "tqq_conn_" & Mid$("ab", j, 1)
        If j = 1 Then
            Set s = AddXY(cht, nm, body.Cells(2, WFP_COL_CXA).Resize(n - 1, 1), body.Cells(2, WFP_COL_CYA).Resize(n - 1, 1), False)
        Else
            Set s = AddXY(cht, nm, body.Cells(2, WFP_COL_CXB).Resize(n - 1, 1), body.Cells(2, WFP_COL_CYB).Resize(n - 1, 1), False)
        End If
        s.ErrorBar Direction:=xlX, Include:=xlErrorBarIncludePlusValues, Type:=xlErrorBarTypeCustom, Amount:="='" & CALC_SHEET & "'!" & body.Cells(2, WFP_COL_CLEN).Resize(n - 1, 1).Address(True, True, xlR1C1)
        With s.ErrorBars
            .EndStyle = xlNoCap
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(89, 89, 89)
            .Format.Line.DashStyle = msoLineDash
            .Format.Line.Weight = 0.75
        End With
    Next
    If cht.HasLegend Then Chart_HideHelperLegendEntries cht
End Sub
```

Em `Wfp_Insert`, logo depois da linha `Wfp_SetNames cht`, acrescente `Wfp_SetConnectors cht, True`.
Em `modCoreApi.ApplyToggle`, `Case "connectors"`, antes de `Case "stacked", "combo"`: `Case "waterfallpair": Wfp_SetConnectors c, value     ' cascata comparativa: auxiliares XY com barra de erro, por serie`.

- [ ] **Passo 4: GREEN e suíte completa** — `-Filter "TestWaterfallPair"` verde (14 testes); `pwsh C:\temp\tqq-cc\build\run-tests.ps1 -TimeoutSec 900` verde. Exporte o gráfico da fixture para PNG (pasta própria, apagada depois) e confira: linhas tracejadas finas do topo de cada barra à próxima barra da mesma série. Exporte também depois de `Core_SetOption "gap", "120"` e depois de mudar `B3` sem chamar nenhum `Wfp_*`: nos três PNG os tracejados terminam na borda da próxima barra da mesma série (se não, a barra de erro não está ligada ao bloco: registre e corrija).

- [ ] **Passo 5: Commit** — `git add src/core/modules/modChartWaterfallPair.bas src/core/modules/modCoreApi.bas tests/vba/core/TestWaterfallPair.bas` e `git commit -m "feat: conectores por serie na cascata comparativa (barras de erro, vivos sem o ThinQQ)"` (+ trailer).

**Verificação no Excel real:** Conectores liga e desliga; Espaço › Largo (80%) mantém os conectores encostados nas barras; mudar um valor move o conector sem clicar em Atualizar.

---

### Tarefa 3: Setas de variação (cotovelo com elipse, percentual ou valor, sala calibrada) — modelo: most-capable

**Files:**
- Modify: `src/core/modules/modChartWaterfallPair.bas` (constantes, `Wfp_SetArrows`, `Wfp_SetArrowMode`, `Wfp_RefreshRoom`; linhas em `Wfp_Insert` e `Wfp_RefreshAfterDataChange`), `src/core/modules/modCoreApi.bas` (`ApplyToggle`: `pairdiff`, `pairabs`; `ReapplyOptions`), `src/core/modules/modChartUtil.bas` (`Chart_EnsureRoomAfterChange`)
- Test: `tests/vba/core/TestWaterfallPair.bas`

**Interfaces:**
- Consumes (Tarefa 1): colunas `WFP_COL_PX/PY/PLBL/RX/RY`, meta `WFP_META_PTV`, `WFP_META_ABS`, `WFP_META_LBL`; `AddXY`, `SetMeta`, `Wfp_CalcIfManual`, `Wfp_StepCount`, `Wfp_RefreshAfterDataChange`. Nos testes (existentes): `Deco_SameScale`, `Deco_ResetScale`, `Stacked_Insert`.
- Produces: `Wfp_SetArrows(cht, turnOn)` (apaga e recria `tqq_pd` e `tqq_pdroom`; grava `pairdiff`; termina chamando `Wfp_RefreshAfterDataChange`), `Wfp_SetArrowMode(cht, absMode)` (meta `Abs`; grava `pairabs`), privada `Wfp_RefreshRoom(cht)` (chamada só dentro de `Wfp_RefreshAfterDataChange`, depois de `Wfp_SyncNames`: `PtV = (máx − mín do eixo) / InsideHeight`, grava se mudou mais de 2%, no máximo 3 passadas). `Chart_EnsureRoomAfterChange` termina chamando `Wfp_RefreshAfterDataChange` (só age na família). `ReapplyOptions` reaplica `pairdiff` (padrão "1") e `pairabs` (padrão "0").

- [ ] **Passo 1: Testes (RED)** — acrescente a `tests/vba/core/TestWaterfallPair.bas`:

```vb
' ROUND do Excel (metade para longe do zero), nao o Round do VBA (banqueiro).
Private Function XlRound(x As Double) As Double
    XlRound = Sgn(x) * Int(Abs(x) + 0.5)
End Function

' Texto esperado da seta: (B - A) / A em % inteiro com sinal; "n/d" com A = 0 ou sinais opostos.
Private Function PctText(a As Double, b As Double) As String
    Dim r As Double
    If a = 0 Or a * b < 0 Then PctText = "n/d": Exit Function
    r = XlRound(100 * (b - a) / a)
    PctText = IIf(r > 0, "+", "") & Format$(r, "0") & "%"
End Function

Private Function ArrowLabel(cht As Chart, i As Long) As String
    ArrowLabel = Chart_FindSeries(cht, "tqq_pd").Points(WFP_ROWS_PER_PAIR * (i - 1) + 3).DataLabel.Text
End Function

Private Function PtvOf(cht As Chart) As Double
    PtvOf = Util_ToDouble(Registry_Block(cht).Cells(1, WFP_META_PTV).Value)
End Function

' Review Focus 2: os percentuais da imagem do think-cell (custo de -4 para -7 = +75%).
Public Sub Test_Wfp_Arrows_TextsLikeTheImage()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, i As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    cht.Refresh
    AssertEqual "+7%", ArrowLabel(cht, 1): AssertEqual "+75%", ArrowLabel(cht, 2): AssertEqual "-67%", ArrowLabel(cht, 3)
    AssertEqual "+67%", ArrowLabel(cht, 4): AssertEqual "-9%", ArrowLabel(cht, 5)
    For i = 1 To 5
        AssertEqual PctText(SideVal(cht, i, 1, "Val"), SideVal(cht, i, 2, "Val")), ArrowLabel(cht, i), "seta " & i & " pelos valores das barras"
    Next
    AssertEqual "1", Registry_GetOption(cht, "pairdiff"): AssertEqual "0", Registry_GetOption(cht, "pairabs")
    Set s = Chart_FindSeries(cht, "tqq_pd")
    AssertEqual 29, s.Points.Count, "5 pontos por par e uma linha vazia entre pares"
    AssertEqual msoArrowheadTriangle, s.Format.Line.EndArrowheadStyle
    AssertEqual msoShapeOval, s.Points(3).DataLabel.Format.AutoShapeType, "rotulo em elipse"
    AssertEqual xlPrimary, s.AxisGroup
    AssertTrue Not Chart_FindSeries(cht, "tqq_pdroom") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Arrows_NdAndAbs()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Etapa", "S1", "S2")
    ws.Range("A2:C2").Value = Array("Inicio", 10, 10)
    ws.Range("A3:C3").Value = Array("Zero", 0, 3)
    ws.Range("A4:C4").Value = Array("Troca", -2, 4)
    ws.Range("A5:C5").Value = Array("Fim", "e", "e")
    Set d = ChartData_FromRange(ws.Range("A1:C5"))
    Set cht = Wfp_Insert(ws, d).Chart
    cht.Refresh
    AssertEqual "0%", ArrowLabel(cht, 1), "iguais"
    AssertEqual "n/d", ArrowLabel(cht, 2), "a serie 1 vale zero"
    AssertEqual "n/d", ArrowLabel(cht, 3), "sinais opostos"
    AssertEqual "+113%", ArrowLabel(cht, 4), "totais 8 e 17: (17 - 8) / 8 = 112,5% -> 113%"
    Wfp_SetArrowMode cht, True
    cht.Refresh
    AssertEqual "0", ArrowLabel(cht, 1): AssertEqual "+3", ArrowLabel(cht, 2): AssertEqual "+6", ArrowLabel(cht, 3): AssertEqual "+9", ArrowLabel(cht, 4)
    AssertEqual "1", Registry_GetOption(cht, "pairabs")
    CloseTemp wb
End Sub

' Review Focus 3: a seta sai 3 pt acima do degrau e 22 pt acima da barra grande com rotulo (regra de tres no eixo real), sobe 12 pt, a elipse
' fica centrada na linha horizontal, entre as duas barras, dentro do grafico.
Public Sub Test_Wfp_Arrows_Geometry()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, i As Long, r0 As Long, p1 As Point, p2 As Point, p3 As Point, p4 As Point, p5 As Point
    Dim gA As Double, gB As Double, lbl As DataLabel, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    ForceLayout cht                                  ' layout desenhado, calibracao nele, e o desenho final
    Core_SafeRefresh cht
    ForceLayout cht
    Set s = Chart_FindSeries(cht, "tqq_pd"): Set body = BodyOf(cht)
    For i = 1 To 5
        r0 = WFP_ROWS_PER_PAIR * (i - 1)
        Set p1 = s.Points(r0 + 1): Set p2 = s.Points(r0 + 2): Set p3 = s.Points(r0 + 3): Set p4 = s.Points(r0 + 4): Set p5 = s.Points(r0 + 5)
        If CBool(body.Cells(i + 1, Wfp_SideCol(1, "Big")).Value) Then gA = WFP_PT_LBLROOM Else gA = WFP_PT_GAP      ' (rotulos ligados e totais positivos na fixture)
        If CBool(body.Cells(i + 1, Wfp_SideCol(2, "Big")).Value) Then gB = WFP_PT_LBLROOM Else gB = WFP_PT_GAP
        AssertAlmost XOfX(cht, 3 * i - 2), p1.Left, 1, "par " & i & ": sai do centro da barra A"
        AssertAlmost XOfX(cht, 3 * i - 1), p5.Left, 1, "par " & i & ": chega ao centro da barra B"
        AssertAlmost YOf(cht, SideVal(cht, i, 1, "Hi")) - gA, p1.Top, 1.5, "par " & i & ": inicio acima da barra A"
        AssertAlmost YOf(cht, SideVal(cht, i, 2, "Hi")) - gB, p5.Top, 1.5, "par " & i & ": ponta acima da barra B"
        AssertAlmost p2.Top, p4.Top, 0.5, "par " & i & ": trecho horizontal"
        AssertAlmost IIf(p1.Top < p5.Top, p1.Top, p5.Top) - 12, p2.Top, 1.5, "par " & i & ": 12 pt acima do inicio mais alto"
        Set lbl = p3.DataLabel
        AssertAlmost XOfX(cht, 3 * i - 1.5), lbl.Left + lbl.Width / 2, 1.5, "par " & i & ": elipse entre as barras"
        AssertAlmost p3.Top, lbl.Top + lbl.Height / 2, 1.5, "par " & i & ": elipse centrada na linha"
        AssertTrue lbl.Top >= 0, "par " & i & ": elipse dentro do grafico"
    Next
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Arrows_RoomCalibrated_Stable_SmallChart()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ax As Axis, ptv As Double, r As Series, i As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    ForceLayout cht
    Core_SafeRefresh cht
    Set ax = cht.Axes(xlValue, xlPrimary)
    ptv = PtvOf(cht)
    AssertAlmost (ax.MaximumScale - ax.MinimumScale) / cht.PlotArea.InsideHeight, ptv, 0.02 * ptv, "PtV = faixa do eixo / altura interna"
    Core_SafeRefresh cht
    Core_SafeRefresh cht
    AssertAlmost ptv, PtvOf(cht), 0.02 * ptv, "estavel em dois refresh"
    ForceLayout cht
    Set r = Chart_FindSeries(cht, "tqq_pdroom")
    For i = 1 To 5
        AssertTrue r.Points(i).Top >= cht.PlotArea.InsideTop - 0.5, "a sala da elipse " & i & " cabe na escala"
    Next
    cht.Parent.Width = 220: cht.Parent.Height = 140
    Core_SafeRefresh cht
    AssertTrue PtvOf(cht) > 0, "grafico pequeno: a calibracao termina com valor positivo"
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Arrows_ResizeRecalibrates()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ax As Axis, ptv0 As Double, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    ptv0 = PtvOf(cht)
    cht.Parent.Height = 200
    ForceLayout cht
    Core_SafeRefresh cht
    ForceLayout cht
    Set ax = cht.Axes(xlValue, xlPrimary)
    AssertAlmost (ax.MaximumScale - ax.MinimumScale) / cht.PlotArea.InsideHeight, PtvOf(cht), 0.02 * PtvOf(cht), "recalibrado no tamanho novo"
    AssertTrue PtvOf(cht) > ptv0 * 1.1, "grafico mais baixo: mais valor por ponto (" & ptv0 & " -> " & PtvOf(cht) & ")"
    Set s = Chart_FindSeries(cht, "tqq_pd")
    AssertAlmost YOf(cht, SideVal(cht, 5, 1, "Hi")) - WFP_PT_LBLROOM, s.Points(WFP_ROWS_PER_PAIR * 4 + 1).Top, 1.5, "o afastamento continua 22 pt no total"
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Arrows_ToggleModeAndRefusal()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    Core_SetToggle "pairdiff", False
    AssertTrue Chart_FindSeries(cht, "tqq_pd") Is Nothing: AssertTrue Chart_FindSeries(cht, "tqq_pdroom") Is Nothing
    AssertEqual "0", Registry_GetOption(cht, "pairdiff"): AssertFalse Core_GetToggle("pairdiff")
    Core_SetToggle "pairdiff", True
    cht.Refresh
    AssertEqual "+7%", ArrowLabel(cht, 1)
    Core_SetToggle "pairabs", True
    cht.Refresh
    AssertEqual "+2", ArrowLabel(cht, 1): AssertTrue Core_GetToggle("pairabs")
    Core_SetToggle "pairabs", False
    cht.Refresh
    AssertEqual "+7%", ArrowLabel(cht, 1)
    Core_SetOption "legend", "right"
    AssertEqual 2, cht.Legend.LegendEntries.Count, "as auxiliares das setas ficam fora da legenda"
    ws.Range("H1:J1").Value = Array("Ano", "A", "B"): ws.Range("H2:J2").Value = Array(2024, 1, 2): ws.Range("H3:J3").Value = Array(2025, 2, 3)
    ws.Range("H1:J3").Select
    Core_Insert "stacked"
    Core_ClearLastNotice
    Core_SetToggle "pairdiff", True
    AssertEqual "'Setas de variação' não se aplica a este tipo de gráfico.", Core_LastNotice()
    Core_ClearLastNotice
    Core_SetToggle "pairabs", True
    AssertEqual "'Variação em valor' não se aplica a este tipo de gráfico.", Core_LastNotice()
    Teardown wb
End Sub

' Review Focus 5 (setas): sem o ThinQQ, mudar a origem muda o texto da elipse e a posicao da seta (formulas do bloco).
Public Sub Test_Wfp_Arrows_LiveWithoutThinQQ()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, ptv As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    ws.Range("C3").Value = -12
    cht.Refresh
    ptv = PtvOf(cht)
    AssertEqual "+200%", ArrowLabel(cht, 2)
    AssertEqual PctText(11, 5), ArrowLabel(cht, 5), "liquido B = 5"
    AssertEqual 17, SideVal(cht, 3, 2, "Hi"), "B3 agora vai de 17 a 15 (antes 22 a 20)": AssertAlmost SideVal(cht, 3, 2, "Hi") + 3 * ptv, Util_ToDouble(BodyOf(cht).Cells(2 * WFP_ROWS_PER_PAIR + 6, WFP_COL_PY).Value), 0.000001, "ponta do par 3 acompanha o topo novo de B"
    CloseTemp wb
End Sub

' Review Focus 3 (Mesma escala): com o eixo fixado por Deco_SameScale, PtV = (max - min fixos) / altura interna e a seta continua 3 pt acima do
' degrau; Escala automatica volta a calibrar pelo eixo automatico.
Public Sub Test_Wfp_Arrows_SameScale_FixedAxis()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, c1 As Chart, c2 As Chart, coll As Collection, skipped As Long, ax As Axis, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set c1 = Wfp_Insert(ws, d).Chart
    ForceLayout c1
    ws.Range("H1:I1").Value = Array("Ano", "Valor"): ws.Range("H2:I2").Value = Array(2024, 80): ws.Range("H3:I3").Value = Array(2025, 90)
    Set c2 = Stacked_Insert(ws, ChartData_FromRange(ws.Range("H1:I3")), "stacked").Chart
    Set coll = New Collection: coll.Add c1: coll.Add c2
    AssertEqual 2, Deco_SameScale(coll, skipped), "as duas entram (escala livre)"
    Core_SafeRefresh c1
    ForceLayout c1
    Set ax = c1.Axes(xlValue, xlPrimary)
    AssertFalse ax.MaximumScaleIsAuto, "eixo fixo"
    AssertAlmost (ax.MaximumScale - ax.MinimumScale) / c1.PlotArea.InsideHeight, PtvOf(c1), 0.02 * PtvOf(c1), "PtV pelo eixo fixo"
    Set s = Chart_FindSeries(c1, "tqq_pd")
    AssertAlmost YOf(c1, SideVal(c1, 2, 1, "Hi")) - WFP_PT_GAP, s.Points(WFP_ROWS_PER_PAIR + 1).Top, 1.5, "par 2: 3 pt acima do degrau com o eixo fixo"
    AssertTrue Deco_ResetScale(c1)
    Core_SafeRefresh c1
    Set ax = c1.Axes(xlValue, xlPrimary)
    AssertAlmost (ax.MaximumScale - ax.MinimumScale) / c1.PlotArea.InsideHeight, PtvOf(c1), 0.02 * PtvOf(c1), "Escala automatica: recalibrado"
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "Wfp_Arrows"` → saída 3 (`Wfp_SetArrowMode` indefinido).

- [ ] **Passo 3: Implementar** — em `modChartWaterfallPair.bas`, nas declarações, depois de `Private Const WFP_NAME_MIN_W ...`:

```vb
Private Const WFP_ROOM_PASSES As Long = 3        ' teto de passadas da calibracao do valor por ponto (PtV)
Private Const WFP_ROOM_TOL As Double = 0.02      ' so grava PtV quando muda mais de 2%
```

Acrescente os procedimentos (depois de `Wfp_SetConnectors`):

```vb
' Setas de variacao: tqq_pd (XY com linhas; 5 pontos por par e uma linha vazia entre pares; com DisplayBlanksAs = xlNotPlotted cada par vira um
' cotovelo com a sua propria ponta, medido), cinza de 1 pt, ponta triangular curta; rotulo so no ponto do meio de cada par (P3), ligado a PLbl,
' numa elipse branca de borda cinza (DataLabel.Format.AutoShapeType, medido). tqq_pdroom: um ponto invisivel 12 pt acima de cada elipse (a escala
' automatica deixa a sala; medido: o maximo foi de 35 a 40). Desligar apaga as duas series (as colunas do bloco ficam).
Public Sub Wfp_SetArrows(cht As Chart, turnOn As Boolean)
    Dim body As Range, n As Long, i As Long, s As Series, nm As Variant, rows As Long
    For Each nm In Array("tqq_pd", "tqq_pdroom")
        Set s = Chart_FindSeries(cht, CStr(nm))
        If Not s Is Nothing Then s.Delete
    Next
    Registry_SetOption cht, "pairdiff", IIf(turnOn, "1", "0")
    If Not turnOn Then Exit Sub
    n = Wfp_StepCount(cht)
    If n < 1 Then Exit Sub
    Set body = Block_Body(Registry_Block(cht))
    rows = WFP_ROWS_PER_PAIR * n - 1                 ' sem a linha vazia depois do ultimo par
    Set s = AddXY(cht, "tqq_pd", body.Cells(2, WFP_COL_PX).Resize(rows, 1), body.Cells(2, WFP_COL_PY).Resize(rows, 1), True)
    With s.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(89, 89, 89)
        .Weight = 1
        .EndArrowheadStyle = msoArrowheadTriangle
        .EndArrowheadLength = msoArrowheadShort
        .EndArrowheadWidth = msoArrowheadNarrow
    End With
    For i = 1 To n
        With s.Points(WFP_ROWS_PER_PAIR * (i - 1) + 3)
            .HasDataLabel = True
            .DataLabel.Formula = Block_Ref(body.Cells(WFP_ROWS_PER_PAIR * (i - 1) + 4, WFP_COL_PLBL))
            .DataLabel.Position = xlLabelPositionCenter
            .DataLabel.Font.Size = 9
            .DataLabel.Font.Bold = False
            .DataLabel.Font.Color = RGB(64, 64, 64)
            .DataLabel.Format.AutoShapeType = msoShapeOval
            .DataLabel.Format.Fill.Visible = msoTrue
            .DataLabel.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .DataLabel.Format.Line.Visible = msoTrue
            .DataLabel.Format.Line.ForeColor.RGB = RGB(166, 166, 166)
            .DataLabel.Format.Line.Weight = 0.75
        End With
    Next
    AddXY cht, "tqq_pdroom", body.Cells(2, WFP_COL_RX).Resize(n, 1), body.Cells(2, WFP_COL_RY).Resize(n, 1), False
    If cht.HasLegend Then Chart_HideHelperLegendEntries cht
    Wfp_RefreshAfterDataChange cht                   ' calibra o valor por ponto com as setas no grafico
End Sub

' Variacao em valor (B - A no formato do grafico, com sinal) em vez do percentual: meta Abs (o texto das setas e formula no bloco).
Public Sub Wfp_SetArrowMode(cht As Chart, absMode As Boolean)
    Dim blk As Range
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    SetMeta blk, WFP_META_ABS, IIf(absMode, 1, 0)
    Registry_SetOption cht, "pairabs", IIf(absMode, "1", "0")
End Sub

' Valor por ponto (meta PtV) = (maximo - minimo do eixo de valores) / altura interna: os afastamentos em pontos das setas viram valor nas
' formulas do bloco. Grava so quando muda mais de 2% (a escala automatica depende dos proprios pontos das setas; medido: converge numa gravacao);
' teto de 3 passadas; calculo manual recalcula o bloco. So dentro de Wfp_RefreshAfterDataChange (guarda de reentrada).
Private Sub Wfp_RefreshRoom(cht As Chart)
    Dim blk As Range, ax As Axis, pass As Long, h As Double, ptv As Double, cur As Double
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    Set ax = cht.Axes(xlValue, xlPrimary)
    For pass = 1 To WFP_ROOM_PASSES
        h = cht.PlotArea.InsideHeight
        If h <= 0 Or ax.MaximumScale <= ax.MinimumScale Then Exit For
        ptv = (ax.MaximumScale - ax.MinimumScale) / h
        cur = Util_ToDouble(blk.Cells(1, WFP_META_PTV).Value)
        If Abs(ptv - cur) <= WFP_ROOM_TOL * ptv Then Exit For
        blk.Cells(1, WFP_META_PTV).Value = ptv
        Wfp_CalcIfManual blk
    Next
    If pass > WFP_ROOM_PASSES Then Errors_Log "Wfp_RefreshRoom" & vbTab & "sem convergencia em " & WFP_ROOM_PASSES & " passadas" & vbTab & ptv   ' o teto agiu: fica no log
End Sub
```

Em `Wfp_RefreshAfterDataChange`, logo depois de `Wfp_SyncNames cht`: `Wfp_RefreshRoom cht                  ' depois dos nomes: a legenda que aparece ou some muda a altura interna`.
Em `Wfp_Insert`, logo depois de `Wfp_SetConnectors cht, True`: `Wfp_SetArrowMode cht, False` e, na linha seguinte, `Wfp_SetArrows cht, True`.

`modCoreApi.ApplyToggle`: na lista do primeiro `Select Case key` acrescente `"pairdiff", "pairabs"`; antes do `Case "roundsum"` do segundo `Select Case`:

```vb
        Case "pairdiff"
            If fam = "waterfallpair" Then Wfp_SetArrows c, value Else RefuseUndispatched key, c, fam      ' setas de variacao entre as barras do par
        Case "pairabs"
            If fam = "waterfallpair" Then Wfp_SetArrowMode c, value Else RefuseUndispatched key, c, fam   ' variacao em valor em vez de percentual
```

`modCoreApi.ReapplyOptions`: a lista do `For Each k In Array(...)` passa a `Array("totals", "labels", "connectors", "secondary", "ind100", "vline", "pieout", "trend", "quad", "roundsum", "pairdiff", "pairabs")` e o padrão a `dflt = IIf(k = "labels" Or k = "ind100" Or k = "pairdiff", "1", "0")`.

`modChartUtil.Chart_EnsureRoomAfterChange`, no fim (depois da linha do corte de eixo):

```vb
    Wfp_RefreshAfterDataChange cht                ' cascata comparativa: valor por ponto das setas e nomes/legenda no tamanho novo (nos demais tipos nao faz nada)
    If Err.Number <> 0 Then Errors_Log "Chart_EnsureRoomAfterChange/comparativa" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
```

- [ ] **Passo 4: GREEN e suíte completa** — `-Filter "TestWaterfallPair"` verde (22 testes); suíte completa verde. PNG da fixture: um cotovelo por etapa com a elipse ("+7%", "+75%", "-67%", "+67%", "-9%"), ponta sobre a barra B; registre no relatório.

- [ ] **Passo 5: Commit** — `git add src/core/modules/modChartWaterfallPair.bas src/core/modules/modCoreApi.bas src/core/modules/modChartUtil.bas tests/vba/core/TestWaterfallPair.bas` e `git commit -m "feat: setas de variacao da cascata comparativa (cotovelo com elipse, percentual ou valor, sala calibrada)"` (+ trailer).

**Verificação no Excel real:** os percentuais da imagem; redimensionar o gráfico (o evento Resize recalibra) mantém a seta acima dos rótulos; sem o ThinQQ (pasta aberta noutro Excel sem o suplemento) mudar um número muda a elipse.

---

### Tarefa 4: Rótulos, totais com percentual, formato e rótulos finos — modelo: standard

**Files:**
- Modify: `src/core/modules/modChartWaterfallPair.bas` (constante `WFP_LBL_MIN_H`; `Wfp_SetLabels`, `Wfp_SetTotalPct`, `Wfp_TotalPctOn`, `Wfp_FormatParts`, `Wfp_ApplyCellFormat`, `Wfp_ApplyNumberFormat`, `Wfp_SyncLabels`; linhas em `Wfp_Insert` e `Wfp_RefreshAfterDataChange`), `src/core/modules/modCoreApi.bas` (`ApplyToggle` totals/labels), `src/core/modules/modChartStyle.bas` e `src/core/modules/modLabels.bas` (troca das saídas cedo da Tarefa 0 por despachos)
- Test: `tests/vba/core/TestWaterfallPair.bas`

**Interfaces:**
- Consumes: `Chart_AddHelperSeries` (vertical: linha invisível), `Labels_LinkPoints`, `Labels_EstimateWidth`, `Style_NumberFormat`, `Style_LocalizeCode`, `Style_AssignNumberFormat`, `SetMeta`, `PointHasLabel`, `Wfp_RefreshAfterDataChange`.
- Produces: `Wfp_SetLabels(cht, turnOn)` (cria/religa `tqq_lbl_a`, `tqq_lbl_b`, `tqq_lbltot`; desligar apaga as três; grava `labels` e a meta `Lbl`), `Wfp_SetTotalPct(cht, turnOn)` (meta `TotPct`, opção `totals`), `Wfp_TotalPctOn(cht) As Boolean`, `Wfp_FormatParts(key, div, dec, suf, sufd)` (pura), `Wfp_ApplyCellFormat(cht)`, `Wfp_ApplyNumberFormat(cht)`, privada `Wfp_SyncLabels(cht)` (no refresh, por último). `Style_ApplyNumberFormat` → `Wfp_ApplyNumberFormat cht: Exit Sub`; `Labels_ApplyCellFormat` → `Wfp_ApplyCellFormat cht: Exit Sub`. `Wfp_Insert` liga os rótulos.

- [ ] **Passo 1: Testes (RED)** — acrescente a `tests/vba/core/TestWaterfallPair.bas`:

```vb
Private Function LblText(cht As Chart, seriesName As String, k As Long) As String
    LblText = Chart_FindSeries(cht, seriesName).Points(k).DataLabel.Text
End Function

Public Sub Test_Wfp_Labels_StepsCenterTotalsAbove()
    Dim wb As Workbook, d As clsChartData, cht As Chart, la As Series, lt As Series, lbl As DataLabel, p As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    cht.Refresh
    AssertEqual "1", Registry_GetOption(cht, "labels")
    AssertEqual Format$(-4, "#,##0"), LblText(cht, "tqq_lbl_a", 4), "degrau A2 assinado"
    AssertEqual Format$(-7, "#,##0"), LblText(cht, "tqq_lbl_b", 5), "degrau B2"
    AssertEqual Format$(-2, "#,##0"), LblText(cht, "tqq_lbl_b", 8), "degrau B3 de 2 visivel no tamanho padrao"
    AssertEqual "", LblText(cht, "tqq_lbl_a", 1), "barra grande: sem rotulo de degrau"
    AssertEqual "", LblText(cht, "tqq_lbl_a", 3), "vao"
    AssertEqual Format$(27, "#,##0"), LblText(cht, "tqq_lbltot", 1), "barra inicial: so o valor"
    AssertEqual Format$(11, "#,##0") & " (" & Format$(41, "0") & "%)", LblText(cht, "tqq_lbltot", 13), "total A: valor e % do 1o total (11 / 27)"
    AssertEqual Format$(10, "#,##0") & " (" & Format$(34, "0") & "%)", LblText(cht, "tqq_lbltot", 14), "total B (10 / 29)"
    Set la = Chart_FindSeries(cht, "tqq_lbl_a"): Set lt = Chart_FindSeries(cht, "tqq_lbltot")
    AssertEqual xlLabelPositionCenter, la.Points(4).DataLabel.Position
    AssertEqual xlLabelPositionAbove, lt.Points(13).DataLabel.Position
    AssertTrue lt.Points(13).DataLabel.Font.Bold
    AssertEqual RGB(255, 255, 255), la.Points(4).DataLabel.Font.Color, "contraste na barra escura"
    AssertEqual RGB(64, 64, 64), Chart_FindSeries(cht, "tqq_lbl_b").Points(5).DataLabel.Font.Color, "contraste na barra clara"
    ForceLayout cht
    Set lbl = la.Points(4).DataLabel
    AssertTrue lbl.Top + lbl.Height / 2 > YOf(cht, 27) And lbl.Top + lbl.Height / 2 < YOf(cht, 23), "rotulo do degrau dentro da barra (de 27 a 23)"
    Set lbl = lt.Points(13).DataLabel: Set p = cht.SeriesCollection(2).Points(13)
    AssertTrue Chart_FindSeries(cht, "tqq_pd").Points(WFP_ROWS_PER_PAIR * 4 + 1).Top <= lbl.Top - 1, "a seta do total comeca acima do rotulo do total"
    AssertTrue lbl.Top + lbl.Height <= p.Top + 1, "rotulo do total acima da barra"
    CloseTemp wb
End Sub

' Review Focus 1 (rotulos): total negativo com o rotulo ABAIXO da barra.
Public Sub Test_Wfp_Labels_NegativeTotalBelow()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lt As Series, lbl As DataLabel, p As Point
    Set wb = NewTempWorkbook(): Set d = FixtureZero(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    cht.Refresh
    Set lt = Chart_FindSeries(cht, "tqq_lbltot")
    AssertEqual Format$(-7, "#,##0") & " (" & Format$(-70, "0") & "%)", LblText(cht, "tqq_lbltot", 10), "total A -7 de um inicio de 10"
    AssertEqual xlLabelPositionBelow, lt.Points(10).DataLabel.Position, "total negativo: abaixo"
    AssertEqual xlLabelPositionAbove, lt.Points(11).DataLabel.Position, "total positivo: acima"
    ForceLayout cht
    Set lbl = lt.Points(10).DataLabel: Set p = cht.SeriesCollection(3).Points(10)
    AssertTrue lbl.Top >= p.Top + p.Height - 1, "abaixo da barra negativa"
    CloseTemp wb
End Sub

' Degrau fino demais: texto vazio (LblOn = 0) e volta quando cabe; rotulo apagado pelo usuario continua apagado.
Public Sub Test_Wfp_Labels_ThinTextHidden_UserDeletedStays()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, la As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    ws.Range("B4").Value = -0.2
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual 0, Util_ToDouble(BodyOf(cht).Cells(8, WFP_COL_LBLON).Value), "degrau A3 de 0,2: fino demais"
    AssertEqual "", LblText(cht, "tqq_lbl_a", 7)
    ws.Range("B4").Value = -6
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual Format$(-6, "#,##0"), LblText(cht, "tqq_lbl_a", 7), "voltou a caber"
    Set la = Chart_FindSeries(cht, "tqq_lbl_a")
    la.Points(4).HasDataLabel = False                ' o usuario apagou o rotulo do degrau A2
    Core_SafeRefresh cht
    AssertFalse la.Points(4).HasDataLabel, "o refresh nao recria rotulo apagado pelo usuario"
    Teardown wb
End Sub

Public Sub Test_Wfp_Totals_PctToggle_And_NumFmt()
    Dim wb As Workbook, d As clsChartData, cht As Chart, dec As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    AssertTrue Wfp_TotalPctOn(cht): AssertTrue Core_GetToggle("totals")
    Core_SetToggle "totals", False
    cht.Refresh
    AssertEqual Format$(11, "#,##0"), LblText(cht, "tqq_lbltot", 13), "sem o percentual"
    AssertFalse Wfp_TotalPctOn(cht): AssertEqual "0", Registry_GetOption(cht, "totals")
    Core_SetToggle "totals", True
    Core_SetOption "numfmt", "dec1"
    cht.Refresh
    dec = Application.International(xlDecimalSeparator)
    AssertEqual "-4" & dec & "0", LblText(cht, "tqq_lbl_a", 4), "degrau com 1 casa"
    AssertEqual "11" & dec & "0 (41%)", LblText(cht, "tqq_lbltot", 13), "total com 1 casa e o percentual inteiro"
    Core_SetToggle "pairabs", True
    cht.Refresh
    AssertEqual "+2" & dec & "0", Chart_FindSeries(cht, "tqq_pd").Points(3).DataLabel.Text, "variacao em valor no formato do grafico"
    Core_SetOption "numfmt", "pct0"
    AssertEqual " p.p.", CStr(Registry_Block(cht).Cells(1, WFP_META_SUFD).Value), "diferenca de percentuais em pontos percentuais"
    Teardown wb
End Sub

Public Sub Test_Wfp_Labels_Toggle_MovesArrowStart()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, k As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    k = WFP_ROWS_PER_PAIR * 4 + 1                    ' P1 do par 5 (Liquido A, barra grande)
    Core_SetToggle "labels", False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl_a") Is Nothing: AssertTrue Chart_FindSeries(cht, "tqq_lbltot") Is Nothing
    AssertEqual 0, Util_ToDouble(Registry_Block(cht).Cells(1, WFP_META_LBL).Value)
    ForceLayout cht
    Set s = Chart_FindSeries(cht, "tqq_pd")
    AssertAlmost YOf(cht, SideVal(cht, 5, 1, "Hi")) - 3, s.Points(k).Top, 1.5, "sem rotulos a seta sai 3 pt acima do total"
    Core_SetToggle "labels", True
    ForceLayout cht
    AssertAlmost YOf(cht, SideVal(cht, 5, 1, "Hi")) - WFP_PT_LBLROOM, s.Points(k).Top, 1.5, "com rotulos, 22 pt (acima do rotulo do total)"
    AssertTrue Not Chart_FindSeries(cht, "tqq_lbl_b") Is Nothing
    Core_SetOption "legend", "right"
    AssertEqual 2, cht.Legend.LegendEntries.Count, "rotulos fora da legenda"
    Teardown wb
End Sub

' Um degrau que vira subtotal e um total que vira degrau DEPOIS de inserir: os pontos que eram NA() na insercao aparecem ligados, com o texto,
' a posicao e a cor certos (o refresh nao religa rotulos).
Public Sub Test_Wfp_Labels_MarkerChange_PointsThatWereNA()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    ws.Range("C5").Value = "e"                       ' Componente 3 de B vira subtotal (fatia 11): 29 - 7 - 2 = 20
    ws.Range("B6").Value = -5                        ' Liquido de A vira degrau (fatia 13)
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual Format$(20, "#,##0") & " (" & Format$(69, "0") & "%)", LblText(cht, "tqq_lbltot", 11), "subtotal novo de B (20 / 29)"
    AssertEqual xlLabelPositionAbove, Chart_FindSeries(cht, "tqq_lbltot").Points(11).DataLabel.Position
    AssertEqual Format$(-5, "#,##0"), LblText(cht, "tqq_lbl_a", 13), "degrau novo de A"
    AssertEqual xlLabelPositionCenter, Chart_FindSeries(cht, "tqq_lbl_a").Points(13).DataLabel.Position
    AssertEqual RGB(255, 255, 255), Chart_FindSeries(cht, "tqq_lbl_a").Points(13).DataLabel.Font.Color
    CloseTemp wb
End Sub
```

Se o `Test_Wfp_Labels_MarkerChange_PointsThatWereNA` falhar no Excel real (rótulo recusado em ponto `NA()` na inserção), a correção prevista é tirar o `NA()` das três colunas de posição: `LblAPos`/`LblBPos` = `(Lo+Hi)/2` do próprio lado (nas fatias do outro lado e nos vãos, `=0`), `TotPos` = `IF(End>=0,Hi,Lo)` em toda fatia de barra (vão `=0`) — o texto ligado (`""`) já esconde o que não deve aparecer, todo ponto é plotado e ligado na inserção; `Wfp_SyncLabels` passa a usar `Big` (não `IsError(tp)`) para decidir Acima/Abaixo. Registre o desvio no relatório.

- [ ] **Passo 2: RED** — `-Filter "Wfp_Labels|Wfp_Totals"` → saída 3 (`Wfp_TotalPctOn` indefinido).

- [ ] **Passo 3: Implementar** — em `modChartWaterfallPair.bas`, nas declarações, depois de `Private Const WFP_ROOM_TOL ...`:

```vb
Private Const WFP_LBL_MIN_H As Double = 10       ' degrau mais baixo que isto (pt) fica sem texto (texto de 9 pt; a caixa de 14 pt tem margem)
```

Acrescente os procedimentos (depois de `Wfp_RefreshRoom`):

```vb
' Rotulos: tqq_lbl_a / tqq_lbl_b (auxiliares de linha em LblAPos/LblBPos, ligadas a LblAVal/LblBVal: o degrau assinado no centro da barra, fonte
' por contraste) e tqq_lbltot (auxiliar de linha em TotPos ligada a TotTxt: "valor (percentual do 1o total)" das barras grandes, negrito, acima;
' abaixo quando o total e negativo: Wfp_SyncLabels). Pontos NA() nao aparecem (xlNotPlotted) e o texto deles e "". Desligar apaga as tres e
' zera a meta Lbl (as setas descem para 3 pt das barras grandes).
Public Sub Wfp_SetLabels(cht As Chart, turnOn As Boolean)
    Dim blk As Range, body As Range, ns As Long, fmt As String, s As Series, nm As Variant
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    If Not turnOn Then
        For Each nm In Array("tqq_lbl_a", "tqq_lbl_b", "tqq_lbltot")
            Set s = Chart_FindSeries(cht, CStr(nm))
            If Not s Is Nothing Then s.Delete
        Next
        Registry_SetOption cht, "labels", "0"
        SetMeta blk, WFP_META_LBL, 0
        Exit Sub
    End If
    Set body = Block_Body(blk)
    ns = Wfp_SlotCount(Wfp_StepCount(cht))
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    Set s = Chart_FindSeries(cht, "tqq_lbl_a")
    If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_lbl_a", body.Cells(2, WFP_COL_LBLAPOS).Resize(ns, 1))
    Labels_LinkPoints s, body.Cells(2, WFP_COL_LBLAVAL).Resize(ns, 1), fmt, xlLabelPositionCenter, Style_ContrastFont(Style_DataSeriesColor(cht, 1))
    Set s = Chart_FindSeries(cht, "tqq_lbl_b")
    If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_lbl_b", body.Cells(2, WFP_COL_LBLBPOS).Resize(ns, 1))
    Labels_LinkPoints s, body.Cells(2, WFP_COL_LBLBVAL).Resize(ns, 1), fmt, xlLabelPositionCenter, Style_ContrastFont(Style_DataSeriesColor(cht, 2))
    Set s = Chart_FindSeries(cht, "tqq_lbltot")
    If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_lbltot", body.Cells(2, WFP_COL_TOTPOS).Resize(ns, 1))
    Labels_LinkPoints s, body.Cells(2, WFP_COL_TOTTXT).Resize(ns, 1), fmt, xlLabelPositionAbove
    s.DataLabels.Font.Bold = True                    ' logo depois de ligar (todos os pontos tem rotulo); o refresh nunca escreve no nivel da serie
    If cht.HasLegend Then Chart_HideHelperLegendEntries cht
    Registry_SetOption cht, "labels", "1"
    SetMeta blk, WFP_META_LBL, 1
    Wfp_ApplyCellFormat cht
    Wfp_RefreshAfterDataChange cht
End Sub

' Totais (alternador Totais neste tipo): o "(41%)" depois do valor das barras de total (meta TotPct; o texto e formula no bloco).
Public Sub Wfp_SetTotalPct(cht As Chart, turnOn As Boolean)
    Dim blk As Range
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    SetMeta blk, WFP_META_TOTPCT, IIf(turnOn, 1, 0)
    Registry_SetOption cht, "totals", IIf(turnOn, "1", "0")
End Sub

Public Function Wfp_TotalPctOn(cht As Chart) As Boolean
    Dim blk As Range
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    Wfp_TotalPctOn = (Util_ToDouble(blk.Cells(1, WFP_META_TOTPCT).Value) = 1)
End Function

' Partes do formato (chave numfmt) dos textos montados por FIXED no bloco: divisor, casas, sufixo dos valores e sufixo das diferencas (nos
' percentuais a diferenca de duas taxas e em pontos percentuais). Pura.
Public Sub Wfp_FormatParts(key As String, ByRef div As Double, ByRef dec As Long, ByRef suf As String, ByRef sufd As String)
    div = 1: dec = 0: suf = "": sufd = ""
    Select Case key
        Case "dec1": dec = 1
        Case "dec2": dec = 2
        Case "pct0": div = 0.01: suf = "%": sufd = " p.p."
        Case "pct1": div = 0.01: dec = 1: suf = "%": sufd = " p.p."
        Case "mil": div = 1000
        Case "mi": div = 1000000: dec = 1
    End Select
End Sub

' Formato nas celulas do bloco: os degraus (numero; NumberFormatLocal do codigo do grafico) e a meta dos textos FIXED (totais e setas).
Public Sub Wfp_ApplyCellFormat(cht As Chart)
    Dim blk As Range, body As Range, ns As Long, key As String, loc As String, div As Double, dec As Long, suf As String, sufd As String
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    ns = Wfp_SlotCount(Wfp_StepCount(cht))
    If ns < 1 Then Exit Sub
    Set body = Block_Body(blk)
    key = Registry_GetOption(cht, "numfmt", "int")
    loc = Style_LocalizeCode(Style_NumberFormat(key))
    body.Cells(2, WFP_COL_LBLAVAL).Resize(ns, 1).NumberFormatLocal = loc
    body.Cells(2, WFP_COL_LBLBVAL).Resize(ns, 1).NumberFormatLocal = loc
    Wfp_FormatParts key, div, dec, suf, sufd
    SetMeta blk, WFP_META_DIV, div
    SetMeta blk, WFP_META_DEC, dec
    SetMeta blk, WFP_META_SUF, suf
    SetMeta blk, WFP_META_SUFD, sufd
End Sub

' numfmt: celulas e meta (Wfp_ApplyCellFormat) e o formato de cada rotulo de degrau, ponto a ponto (ponto sem rotulo fica como esta).
Public Sub Wfp_ApplyNumberFormat(cht As Chart)
    Dim fmt As String, s As Series, nm As Variant, i As Long
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    Wfp_ApplyCellFormat cht
    For Each nm In Array("tqq_lbl_a", "tqq_lbl_b")
        Set s = Chart_FindSeries(cht, CStr(nm))
        If Not s Is Nothing Then
            For i = 1 To s.Points.Count
                If PointHasLabel(s, i) Then Style_AssignNumberFormat s.Points(i).DataLabel, fmt
            Next
        End If
    Next
End Sub

' Rotulos de degrau finos demais (altura < 10 pt ou texto mais largo que a fatia) ficam com o texto vazio (coluna LblOn = 0) e voltam quando
' cabem; o rotulo continua existindo, entao um rotulo apagado pelo usuario continua apagado (nada e recriado aqui). Barra grande com total
' negativo: rotulo abaixo (TotPos < Top). So escreve o que muda; cada ponto e isolado. Chamada por ultimo em Wfp_RefreshAfterDataChange.
Private Sub Wfp_SyncLabels(cht As Chart)
    Dim blk As Range, body As Range, n As Long, ns As Long, i As Long, j As Long, k As Long, ax As Axis, ppu As Double, slotPt As Double
    Dim sec As String, d As Double, want As Long, arr As Variant, changed As Boolean, s As Series, tp As Variant, tt As Variant, pos As Long
    If Registry_GetOption(cht, "labels", "1") <> "1" Then Exit Sub
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    n = Wfp_StepCount(cht)
    If n < 1 Then Exit Sub
    ns = Wfp_SlotCount(n)
    Set body = Block_Body(blk)
    Set ax = cht.Axes(xlValue, xlPrimary)
    If ax.MaximumScale <= ax.MinimumScale Then Exit Sub
    ppu = cht.PlotArea.InsideHeight / (ax.MaximumScale - ax.MinimumScale)
    slotPt = cht.PlotArea.InsideWidth / ns
    sec = Split(Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int")), ";")(0)
    arr = body.Cells(2, WFP_COL_LBLON).Resize(ns, 1).Value
    For i = 1 To n
        For j = 1 To 2
            If Not CBool(body.Cells(i + 1, Wfp_SideCol(j, "Big")).Value) Then
                d = Util_ToDouble(body.Cells(i + 1, Wfp_SideCol(j, "D")).Value)
                k = Wfp_SlotOf(i, j)
                If d <> 0 Then                               ' degrau zero: o formato ja mostra vazio
                    want = 1
                    If Abs(d) * ppu < WFP_LBL_MIN_H Then want = 0
                    If Labels_EstimateWidth(Format$(d, sec)) > slotPt Then want = 0
                    If Util_ToDouble(arr(k, 1)) <> want Then arr(k, 1) = want: changed = True
                End If
            End If
        Next
    Next
    If changed Then
        body.Cells(2, WFP_COL_LBLON).Resize(ns, 1).Value = arr
        Wfp_CalcIfManual blk
    End If
    Set s = Chart_FindSeries(cht, "tqq_lbltot")
    If s Is Nothing Then Exit Sub
    tp = body.Cells(2, WFP_COL_TOTPOS).Resize(ns, 1).Value
    tt = body.Cells(2, WFP_COL_TOP).Resize(ns, 1).Value
    For k = 1 To ns
        If Not IsError(tp(k, 1)) Then
            If Util_ToDouble(tp(k, 1)) < Util_ToDouble(tt(k, 1)) - 0.000000001 Then pos = xlLabelPositionBelow Else pos = xlLabelPositionAbove
            On Error Resume Next
            If s.Points(k).HasDataLabel Then
                If s.Points(k).DataLabel.Position <> pos Then s.Points(k).DataLabel.Position = pos
            End If
            If Err.Number <> 0 Then Errors_Log "Wfp_SyncLabels" & vbTab & "ponto " & k & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
            On Error GoTo 0
        End If
    Next
End Sub
```

Em `Wfp_RefreshAfterDataChange`, logo depois de `Wfp_RefreshRoom cht`: `Wfp_SyncLabels cht                  ' por ultimo: a escala final decide os rotulos finos`.
Em `Wfp_Insert`, logo depois de `Wfp_SetNames cht` (antes de `Wfp_SetConnectors cht, True`): `Wfp_SetLabels cht, True`.

`modCoreApi.ApplyToggle`: `Case "totals"`, antes do `Case "stacked", "combo"`: `Case "waterfallpair": Wfp_SetTotalPct c, value          ' cascata comparativa: o percentual do 1o total nas barras de total`; `Case "labels"`, antes do `Case "stacked", "combo"`: `Case "waterfallpair": Wfp_SetLabels c, value          ' cascata comparativa: degraus no centro, barras grandes acima`.
`modChartStyle.Style_ApplyNumberFormat`: troque a linha da Tarefa 0 por `If Kinds_Family(Registry_ChartType(cht)) = "waterfallpair" Then Wfp_ApplyNumberFormat cht: Exit Sub   ' Cascata comparativa: degraus e textos FIXED (Wfp_ApplyNumberFormat)`.
`modLabels.Labels_ApplyCellFormat`: troque a linha da Tarefa 0 por `If Kinds_Family(kind) = "waterfallpair" Then Wfp_ApplyCellFormat cht: Exit Sub    ' Cascata comparativa: celulas dos degraus e meta dos textos FIXED`.

- [ ] **Passo 4: GREEN e suíte completa** — `-Filter "TestWaterfallPair"` verde (28 testes); suíte completa verde. PNG da fixture: "27"/"29" acima das barras iniciais em negrito, "-4"/"-7" etc. no centro dos degraus, "11 (41%)" e "10 (34%)"; o da `FixtureZero`: "-7 (-70%)" abaixo da barra; registre.

- [ ] **Passo 5: Commit** — `git add src/core/modules/modChartWaterfallPair.bas src/core/modules/modCoreApi.bas src/core/modules/modChartStyle.bas src/core/modules/modLabels.bas tests/vba/core/TestWaterfallPair.bas` e `git commit -m "feat: rotulos da cascata comparativa (degraus, totais com percentual do 1o total, formato, rotulos finos)"` (+ trailer).

**Verificação no Excel real:** Formato › 1.234,5 muda degraus, totais e (com Variação em valor) a elipse; Totais desliga o "(41%)"; apagar um rótulo à mão e clicar em Atualizar: ele continua apagado.

---
### Tarefa 5: Faixa de opções (Cascata comparativa, Setas de variação, Variação em valor, Nome dentro das barras) e ícone próprio — modelo: standard

**Files:**
- Modify: `src/loader/ribbon/customUI14.xml`, `build/lib/Ribbon-Icons.ps1` (ícone `tqq_waterfallpair`), `docs/tutorial/tutorial.html` (seção `#cascata-comparativa` inicial, linha em `#tipos`, índice, `k` → `ui`), `tests/ps/Test-Docs.ps1` (seção nova)
- Test: `tests/vba/loader/TestRibbon.bas`

**Interfaces:**
- Consumes (T0..T4): `Kinds_IsValid("waterfallpair")`, recursos `pairdiff`, `pairabs`, `inbar` com `Kinds_FeatureLabel`, despachos de `Core_SetToggle`/`Core_SetOption`.
- Produces: botão `btnWaterfallPair` ("Cascata comparativa", `tag="waterfallpair"`, `image="tqq_waterfallpair16"`) no menu Cascata, depois de `btnBarWaterfall`; alternadores `tglPairDiff` ("Setas de variação", `tag="pairdiff"`, `image="tqq_arrows16"`) e `tglPairAbs` ("Variação em valor", `tag="pairabs"`, `image="tqq_waterfallpair16"`) no menu Setas, antes de `msepArrows`; item `inbar` "Nome dentro das barras" no dropdown Legenda; supertips de Totais, Legenda e Setas acrescidos; ícone `tqq_waterfallpair` (14 ícones na tabela). Contagens: 21 botões de inserção, 12 alternadores, 4 dropdowns, 8 menus, 3 botões de anotação.

- [ ] **Passo 1: Testes (RED)** — em `tests/vba/loader/TestRibbon.bas`:
  - `Test_Ribbon_Tags_MatchCoreKeys`: `AssertEqual 21, d.SelectNodes("//u:button[@onAction='Rx_OnInsert']").Length`.
  - `Test_Ribbon_Menus_Structure`: o `AssertEqual 20` (contagem de `Rx_OnInsert`) passa a `21`; na lista de tipos válidos do `InStr`, acrescente `waterfallpair,` antes da vírgula final (`...,scatter,bubble,waterfallpair,`); antes do `End Sub` acrescente
    ```vb
        ' Cascata comparativa: terceiro item do menu Cascata; as setas de variacao (2 alternadores) no menu Setas, antes do separador
        AssertEqual "btnWaterfall,btnBarWaterfall,btnWaterfallPair", AttrText(d.SelectSingleNode("//u:menu[@id='mnuWaterfall']/u:button[1]"), "id") & "," & AttrText(d.SelectSingleNode("//u:menu[@id='mnuWaterfall']/u:button[2]"), "id") & "," & AttrText(d.SelectSingleNode("//u:menu[@id='mnuWaterfall']/u:button[3]"), "id")
        AssertEqual 2, d.SelectNodes("//u:menu[@id='mnuArrows']/u:toggleButton").Length
        AssertEqual "msepArrows", AttrText(d.SelectSingleNode("//u:toggleButton[@id='tglPairAbs']/following-sibling::*[1]"), "id"), "as setas de variacao antes do separador"
    ```
  - `Test_Ribbon_Toggles_And_Dropdowns_Phase2`: `AssertEqual 12, d.SelectNodes("//u:toggleButton").Length`; a lista de tags válidas passa a `",totals,labels,connectors,ind100,vline,secondary,pieout,roundsum,trend,quad,pairdiff,pairabs,"`; o item da legenda passa a `AssertEqual "right,none,lineend,inbar", ItemIds(d.SelectSingleNode("//u:dropDown[@tag='legend']"))`.
  - `Test_Ribbon_IconFallbackMso_AreValid`: `AssertEqual 14, checked, "uma reserva por icone da tabela"`.
  - `Test_Ribbon_Enabled_FollowsCapabilityOfSelectedChart`: depois da linha `AssertFalse Rx_IsEnabledFor("pieout")` (que segue a cascata de duas séries), acrescente
    ```vb
        AssertFalse Rx_IsEnabledFor("pairdiff"), "cascata comum: sem setas de variacao"
        InsertAndSelect ws, "A40:C43", "waterfallpair"
        For Each t In Array("pairdiff", "pairabs", "inbar", "totals", "connectors", "labels", "gap", "legend", "numfmt", "")
            AssertTrue Rx_IsEnabledFor(CStr(t)), "cascata comparativa: '" & t & "' deveria estar habilitado"
        Next
        For Each t In Array("cagr", "diff", "bracket", "vline", "roundsum", "break", "pieout")
            AssertFalse Rx_IsEnabledFor(CStr(t)), "cascata comparativa: '" & t & "' deveria estar desabilitado"
        Next
    ```
  - Teste novo:
    ```vb
    ' Rotulos, tags, icones e textos de ajuda exatos dos controles da cascata comparativa (o tutorial e o Test-Docs dependem dos rotulos).
    Public Sub Test_Ribbon_PairWaterfall_LabelsAndTips()
        Dim d As Object, spec As Variant, i As Long, p() As String, n As Object
        Set d = LoadRibbon()
        spec = Array("btnWaterfallPair|Cascata comparativa|waterfallpair|tqq_waterfallpair16", "tglPairDiff|Setas de variação|pairdiff|tqq_arrows16", "tglPairAbs|Variação em valor|pairabs|tqq_waterfallpair16")
        For i = LBound(spec) To UBound(spec)
            p = Split(spec(i), "|")
            Set n = d.SelectSingleNode("//*[@id='" & p(0) & "']")
            AssertTrue Not n Is Nothing, "controle ausente: " & p(0)
            AssertEqual p(1), AttrText(n, "label"), "rotulo de " & p(0)
            AssertEqual p(2), AttrText(n, "tag"), "tag de " & p(0)
            AssertEqual p(3), AttrText(n, "image"), "icone de " & p(0)
            AssertTrue Len(AttrText(n, "supertip")) > 20, "sem supertip: " & p(0)
        Next
        AssertEqual "Nome dentro das barras", AttrText(d.SelectSingleNode("//u:dropDown[@tag='legend']/u:item[@id='inbar']"), "label")
        AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='tglTotals']"), "supertip"), "cascata comparativa") > 0, "Totais explica o percentual da comparativa"
        AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='tglPairDiff']"), "supertip"), "n/d") > 0
        AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='btnWaterfallPair']"), "supertip"), "exatamente 2 colunas de valores") > 0
    End Sub
    ```

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-cc\build\run-tests.ps1 -Target loader -Filter "TestRibbon" -TimeoutSec 400` → falhas nas contagens e no teste novo.

- [ ] **Passo 3: customUI14.xml** —
  1. Menu `mnuWaterfall`: o `supertip` do menu passa a `"Tabela: etapa na 1ª coluna e valor na 2ª. Escreva 'e' no valor para uma barra de total/subtotal calculada. A Cascata comparativa põe duas séries lado a lado."`; depois de `btnBarWaterfall` acrescente
     ```xml
            <button id="btnWaterfallPair" label="Cascata comparativa" tag="waterfallpair" onAction="Rx_OnInsert" image="tqq_waterfallpair16"
                    screentip="Cascata comparativa" supertip="Duas pontes lado a lado (uma por série), com conectores e a variação entre as duas barras de cada etapa numa seta. Tabela: etapa e exatamente 2 colunas de valores; 'e' no valor para um total."/>
     ```
  2. `tglTotals`: o `supertip` passa a `"Mostra o total de cada coluna, barra ou área empilhada ou 100%, da combinação e do Mekko (no Mekko (%), a largura da coluna). Na cascata comparativa, acrescenta o percentual do primeiro total ao rótulo das barras de total. Agrupadas, linha, cascata, pizza e Gantt não têm totais."`
  3. `ddLegend`: depois do item `lineend`, `<item id="inbar"   label="Nome dentro das barras"/>`; o `supertip` passa a `"Direita, nenhuma ou o nome de cada série no fim da linha (linha e combinação) ou dentro das barras (cascata comparativa). Na cascata a legenda só existe com várias séries."`
  4. Menu `mnuArrows`: o `supertip` passa a `"Seta de CAGR e de diferença entre duas categorias; colchete sobre atividades do Gantt; setas de variação da cascata comparativa."`; antes de `<menuSeparator id="msepArrows"/>` acrescente
     ```xml
            <toggleButton id="tglPairDiff" label="Setas de variação" image="tqq_arrows16" tag="pairdiff"
                          onAction="Rx_OnToggle" getPressed="Rx_GetPressed" getEnabled="Rx_GetEnabled"
                          screentip="Setas de variação (cascata comparativa)" supertip="Na cascata comparativa, uma seta em cotovelo da barra da 1ª série para a da 2ª em cada etapa, com a variação numa elipse: (B - A) / A em %, ou n/d quando A é zero ou os sinais são opostos."/>
            <toggleButton id="tglPairAbs" label="Variação em valor" image="tqq_waterfallpair16" tag="pairabs"
                          onAction="Rx_OnToggle" getPressed="Rx_GetPressed" getEnabled="Rx_GetEnabled"
                          screentip="Variação em valor (cascata comparativa)" supertip="As setas de variação mostram a diferença B - A no formato dos números do gráfico (com Formato em %, em pontos percentuais) em vez do percentual."/>
     ```
  (Filho de menu nunca leva `size`; ids não começam por dígito; nenhum controle com `image` e `imageMso`.)

- [ ] **Passo 4: Ícone** — em `build/lib/Ribbon-Icons.ps1`, dentro de `Get-TqqIconSpecs`, depois da linha `tqq_break`:
  ```powershell
        tqq_waterfallpair = @{ Mso = 'ChartTypeOtherInsertGallery'; Rects = @((New-TqqRect 0 3 2 12 $n), (New-TqqRect 2 2 2 13 $l), (New-TqqRect 6 3 2 4 $n), (New-TqqRect 8 2 2 3 $l), (New-TqqRect 12 7 2 8 $n), (New-TqqRect 14 9 2 6 $l)) }
  ```
  (três pares navy/azul-claro: início, degrau que desce, total; a reserva é a da cascata, já validada por `GetImageMso`).

- [ ] **Passo 5: Tutorial (mínimo para os rótulos novos)** — em `docs/tutorial/tutorial.html`:
  - Índice (`<nav>`): depois de `<li><a href="#cascata">Cascata</a></li>` acrescente `<li><a href="#cascata-comparativa">Cascata comparativa</a></li>`.
  - Tabela de `#tipos`: depois da linha de `Cascata (barras)` acrescente
    ```html
          <tr><td><span class="ui">Cascata</span> › <a href="#cascata-comparativa"><span class="ui">Cascata comparativa</span></a></td><td>Duas pontes lado a lado, uma por série, com conectores e a variação entre as duas barras de cada etapa.</td><td>Etapa e exatamente 2 colunas de valores.</td></tr>
    ```
  - Seção nova, entre `</section>` da `#cascata` e `<section id="linha-area"`:
    ```html
    <section id="cascata-comparativa" aria-labelledby="h-cascata-comparativa">
      <h2 id="h-cascata-comparativa">Cascata comparativa</h2>
      <p>Duas pontes lado a lado, uma por série, para comparar etapa por etapa (por exemplo 2024 x 2025, ou duas lojas): em cada etapa a barra da primeira série fica à esquerda e a da segunda à direita, e uma seta mostra a variação entre elas.</p>
      <dl class="ficha">
        <dt>Selecione</dt><dd>A tabela com o cabeçalho: a etapa na primeira coluna e exatamente 2 colunas de valores.</dd>
        <dt>Menu</dt><dd><span class="ui">Cascata</span> › <span class="ui">Cascata comparativa</span>.</dd>
        <dt>Marcador</dt><dd><code>e</code> (ou <code>t</code>, <code>=</code>, <code>total</code>, <code>subtotal</code>) no lugar do valor: barra de total calculada, em cada série.</dd>
        <dt>Opções</dt><dd><span class="ui">Setas</span> › <span class="ui">Setas de variação</span> e <span class="ui">Variação em valor</span>; <span class="ui">Legenda</span> › <span class="ui">Nome dentro das barras</span>; <span class="ui">Totais</span> liga e desliga o percentual do primeiro total.</dd>
      </dl>
    </section>
    ```
  - Matriz (`#matriz`): troque `<span class="k">Cascata comparativa</span>` (4 linhas), `<span class="k">Setas de variação</span>`, `<span class="k">Variação em valor</span>` e `<span class="k">Nome dentro das barras</span>` por `<span class="ui">...</span>` (os rótulos passaram a existir na faixa).
  - `tests/ps/Test-Docs.ps1`: depois do laço das seções da Fase 3 (`foreach ($id in 'borboleta', 'dispersao', ...)`), acrescente
    ```powershell
        # Cascata comparativa (1.3.0): seção própria, no índice, e a linha da tabela de #tipos leva a ela
        Check ($null -ne $xml.SelectSingleNode("//*[local-name()='section'][@id='cascata-comparativa']")) 'tutorial tem a seção #cascata-comparativa'
        Check ($null -ne $xml.SelectSingleNode("//nav//a[@href='#cascata-comparativa']")) 'o índice leva à seção #cascata-comparativa'
        Check ($null -ne $xml.SelectSingleNode("//*[@id='tipos']//a[@href='#cascata-comparativa']")) 'a tabela de #tipos leva à seção #cascata-comparativa'
    ```

- [ ] **Passo 6: GREEN e suíte completa** — `-Target loader -Filter "TestRibbon"` verde; `pwsh C:\temp\tqq-cc\build\run-tests.ps1 -TimeoutSec 900 -LoaderSwap` verde (inclusive `Test-Ribbon.ps1`, que confere o PNG `tqq_waterfallpair16` no pacote, e `Test-Docs`).

- [ ] **Passo 7: Commit** — `git add src/loader/ribbon/customUI14.xml build/lib/Ribbon-Icons.ps1 docs/tutorial/tutorial.html tests/ps/Test-Docs.ps1 tests/vba/loader/TestRibbon.bas` e `git commit -m "feat: faixa 1.3.0 (Cascata comparativa, Setas de variacao, Variacao em valor, Nome dentro das barras) e icone proprio"` (+ trailer).

**Verificação no Excel real (cópia instalada só se o controlador pedir; nunca `-Install` no perfil do usuário):** a aba ThinQQ aparece; Cascata › Cascata comparativa com ícone; com um gráfico comum selecionado, Setas de variação e Variação em valor ficam cinza; com a comparativa, ativos e pressionados conforme o estado.

---

### Tarefa 6: Integração (todos os tipos, Alterar dados) e Exemplos (27) — modelo: standard

**Files:**
- Modify: `src/core/modules/modSamples.bas`, `docs/tutorial/tutorial.html` e `README.md` (só a contagem de ThinQQ › Exemplos: 26 → 27)
- Test: `tests/vba/core/TestCoreApi.bas`, `tests/vba/core/TestSamples.bas`

**Interfaces:**
- Consumes: `Wfp_Insert` (via `Kinds_Insert`), `Wfp_TotalPctOn`, opções e séries da cascata comparativa (T1..T4).
- Produces: seção "27) Cascata comparativa: duas pontes lado a lado e a variação entre elas" na planilha "Exemplos 3" (linha 169; 9 seções); `Samples_Create` monta 28 gráficos.

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestCoreApi.bas`:
  - `FixtureFor`, antes do `Case Else`:
    ```vb
            Case "waterfallpair"
                ws.Range("A1:C1").Value = Array("Etapa", "Loja", "Online")
                ws.Range("A2:C2").Value = Array("Inicio", 100, 50): ws.Range("A3:C3").Value = Array("Ajuste", -30, 10)
                ws.Range("A4:C4").Value = Array("Fim", "e", "e")
                Set FixtureFor = ws.Range("A1:C4")
    ```
  - `Test_Core_AllKinds_InsertTogglesOptionsRebuild`: acrescente `"waterfallpair"` ao fim da lista `kinds` e `"pairdiff", "pairabs"` ao fim da lista `toggles`.
  - `ChoicesSnapshot`: a lista passa a `Array("totals", "labels", "connectors", "ind100", "vline", "secondary", "pieout", "trend", "quad", "roundsum", "pairdiff", "pairabs", "vbreak", "numfmt", "palette", "gap", "legend", "n")`.
  - `AssertRealState`: troque o bloco `If Kinds_Family(kind) = "waterfall" Then ... Else ... End If` dos conectores e a linha de totais por
    ```vb
        If Kinds_Family(kind) = "waterfall" Then
            For i = 1 To c.Shapes.Count
                If Left$(c.Shapes(i).Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then conn = True
            Next
        ElseIf Kinds_Family(kind) = "waterfallpair" Then
            conn = Not Chart_FindSeries(c, "tqq_conn_a") Is Nothing          ' cascata comparativa: conectores sao auxiliares XY
        Else
            For Each g In c.ChartGroups
                If GroupHasSeriesLines(g) Then conn = True
            Next
        End If
        ' (na cascata tqq_total sao as barras de total do proprio grafico; na comparativa, Totais e o percentual nas barras de total)
        If Kinds_Family(kind) = "waterfallpair" Then
            AssertTrue Wfp_TotalPctOn(c), tag & ": percentual do 1o total ligado (Totais)"
        ElseIf Kinds_Family(kind) <> "waterfall" Then
            AssertEqual Kinds_Supports(kind, "totals"), Not Chart_FindSeries(c, "tqq_total") Is Nothing, tag & ": totais (serie tqq_total)"
        End If
    ```
    e, antes do `End Sub`, `If Kinds_Supports(kind, "pairdiff") Then AssertTrue Not Chart_FindSeries(c, "tqq_pd") Is Nothing, tag & ": setas de variacao ligadas"`.
  - `Test_Core_RebuildFromRange_ValidationRefusals_NoticeNotGenericError`: acrescente à lista `cases` (antes do fecho) `, _` e `Array("waterfallpair", "A cascata comparativa compara exatamente 2 séries; a tabela tem 3 colunas de valores. Selecione a coluna das etapas e só as duas colunas a comparar.", 4)` (a `BadTable` de 4 colunas tem 3 de valores).

  `tests/vba/core/TestSamples.bas`:
  - `TitleRows3`: `Dim r(0 To 8) As Long` e `For k = 0 To 8`.
  - `ExpectedKinds3`: `Array("butterfly", "scatter", "bubble", "stacked100", "stacked", "barstacked", "area", "scatter", "waterfallpair")`.
  - `CheckSheet3KindsAndFeatures`: `AssertEqual 9, ws.ChartObjects.Count, "9 secoes, um grafico cada"`.
  - `CheckSheet3Layout`: `If i < 9 Then ...`, `For i = 0 To 8`, e no fim `AssertEqual "27) ", Left$(CStr(ws.Cells(169, 1).Value), 4)`.
  - `CheckSheet3Labels`, antes do `End Sub`:
    ```vb
        ' 27: a cascata comparativa da imagem do think-cell
        Set cht = ws.ChartObjects(9).Chart
        AssertEqual "waterfallpair", Registry_ChartType(cht)
        AssertEqual "+7%", Chart_FindSeries(cht, "tqq_pd").Points(3).DataLabel.Text, "27: seta do Bruto"
        AssertEqual "+75%", Chart_FindSeries(cht, "tqq_pd").Points(9).DataLabel.Text, "27: seta do Componente 1 (custo que cresce)"
        AssertEqual Format$(11, "#,##0") & " (" & Format$(41, "0") & "%)", Chart_FindSeries(cht, "tqq_lbltot").Points(13).DataLabel.Text, "27: total da serie 1"
    ```
  - O comentário de `Test_Samples_AllSheets` passa a "(Samples_Create monta 28 graficos)".

- [ ] **Passo 2: RED** — `-Target core -Filter "AllKinds|ValidationRefusals|TestSamples" -TimeoutSec 600` → `TestSamples` falha (8 gráficos); `AllKinds` e `ValidationRefusals` podem já passar (o tipo está pronto): é o esperado, eles conferem a integração.

- [ ] **Passo 3: Exemplo 27** — em `src/core/modules/modSamples.bas`, em `Samples_CreateSheet3`, antes de `Set Samples_CreateSheet3 = ws`:
  ```vb
    WriteTitle ws, 169, "27) Cascata comparativa: duas pontes lado a lado e a variação entre elas"
    WriteTable ws.Range("A170"), Array(Array("Etapa", "Segmento 1", "Segmento 2"), Array("Bruto", 27, 29), Array("Componente 1", -4, -7), Array("Componente 2", -6, -2), Array("Componente 3", -6, -10), Array("Líquido", "e", "e"))
    Set d = ChartData_FromRange(ws.Range("A170:C175"))
    Set co = Kinds_Insert(ws, d, "waterfallpair")
    Settle co
  ```
  e troque o comentário do cabeçalho da função para "os tipos e recursos da Fase 3 e a cascata comparativa (1.3.0), 9 secoes (19 a 27) a cada 21 linhas".
  Documentação (a contagem que o `Test-Docs` lê de `Samples_Create` passa a 27): no tutorial, "abre uma pasta de trabalho nova com 26 exemplos em três planilhas" → "27 exemplos"; no README, "ThinQQ > Exemplos gera 26 exemplos menores" → "27 exemplos menores" e "a planilha Exemplos 3 completa os 26 exemplos de ThinQQ > Exemplos" → "os 27 exemplos". (As citações da **demonstração**, "26 exemplos, de E01 a E26", ficam: mudam na Tarefa 7.)

- [ ] **Passo 4: GREEN e suíte completa** — filtro do passo 2 verde; suíte completa verde (o `Test-Docs` confere "27 exemplos" e "26 exemplos" nos dois documentos).

- [ ] **Passo 5: Commit** — `git add src/core/modules/modSamples.bas docs/tutorial/tutorial.html README.md tests/vba/core/TestCoreApi.bas tests/vba/core/TestSamples.bas` e `git commit -m "feat: exemplo 27 (cascata comparativa) e cascata comparativa em todos os testes de integracao"` (+ trailer).

**Verificação no Excel real:** ThinQQ › Exemplos, planilha "Exemplos 3", seção 27 igual à imagem do think-cell; Alterar dados na comparativa com uma tabela de 3 colunas de valores avisa e mantém o gráfico.

---

### Tarefa 7: Pasta de demonstração (E27, E28) e galeria do tutorial — modelo: most-capable

**Files:**
- Modify: `tools/DemoCoverage.ps1`, `tools/New-DemoWorkbook.ps1`, `tests/ps/Test-DemoWorkbook.ps1` (mapa de recursos, conferência do E27, remoção de `$pendingDemo`), `tools/Export-DemoChartImages.ps1` e `tools/Build-TutorialGallery.ps1` (só os comentários "E01..E26" → "E01..E28"), `docs/tutorial/tutorial.html` (linhas E27/E28 em `#demonstracao`, galeria regerada, contagens, links da seção nova), `README.md` (contagens da demonstração)
- Regenerate: `exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx`

**Interfaces:**
- Consumes: API real do core de Teste (`Core_Insert "waterfallpair"`, `Core_SetToggle "pairabs"`, `Core_SetOption "legend" "right"`), `Get-TqqDemoCoverage`.
- Produces: itens 27 e 28 da cobertura (`Kind = 'waterfallpair'`; Features `pairdiff`, `inbar` e `pairabs`); recursos visíveis `pairdiff`, `pairabs`, `inbar` em `Get-TqqDemoFeatures`; nome do tipo "Cascata comparativa"; galeria com 28 figuras e 29 imagens.

- [ ] **Passo 1: Cobertura** — `tools/DemoCoverage.ps1`:
  - `Get-TqqDemoFeatures`: acrescente `'pairdiff', 'pairabs', 'inbar'` ao fim da lista.
  - `Get-TqqDemoKindName`: no `@{...}` acrescente `'waterfallpair' = 'Cascata comparativa'`.
  - Depois do item 26:
    ```powershell
        @{ Num = 27; Kind = 'waterfallpair'; Features = @('pairdiff', 'inbar')
           Menu = 'Inserir gráfico > Cascata > Cascata comparativa (Setas de variação e Nome dentro das barras já vêm ligados)'
           Title = 'Da receita ao lucro, 2024 x 2025 (cascata comparativa)'
           Desc = 'As duas pontes lado a lado, 2024 à esquerda e 2025 à direita, com os anos escritos dentro das barras da receita. A seta de cada etapa mostra quanto 2025 variou sobre 2024, (2025 - 2024) / 2024 (num custo, "+" quer dizer que o custo cresceu), e o lucro mostra o percentual da receita do próprio ano. Tudo acompanha os números, com ou sem o ThinQQ.' }
        @{ Num = 28; Kind = 'waterfallpair'; Features = @('pairabs')
           Menu = 'Inserir gráfico > Cascata > Cascata comparativa; Setas > Variação em valor; Legenda > Direita'
           Title = 'As duas maiores lojas: da receita ao lucro em 2025 (variação em valor)'
           Desc = 'A mesma ponte para as duas lojas que mais venderam em 2025; a seta de cada etapa mostra a diferença em R$ milhões entre a segunda loja e a primeira, e os nomes das lojas ficam na legenda.' }
    ```

- [ ] **Passo 2: Gerador** — `tools/New-DemoWorkbook.ps1`:
  - Na seção "tabelas", depois de `$roundBody = ...`:
    ```powershell
    # cascata comparativa (E27, E28): a ponte da receita ao lucro com as proporções do E05 (despesas 21%, depreciação e juros 2,5% da receita,
    # impostos 34% do lucro antes deles), em R$ milhões com 1 casa (em R$ mil os rótulos dos degraus ficariam mais largos que a fatia); 'e' nas linhas de total
    function Get-DemoPL([int[]]$sIdx, [int[]]$mIdx) {
        $rcv = Get-DemoSum 'sales' $sIdx $allD $mIdx; $cmvY = Get-DemoSum 'cost' $sIdx $allD $mIdx
        $desp = 0.21 * $rcv; $dep = 0.025 * $rcv; $imp = 0.34 * ($rcv - $cmvY - $desp - $dep)
        return ,@(([Math]::Round($rcv / 1e6, 1)), (-([Math]::Round($cmvY / 1e6, 1))), (-([Math]::Round($desp / 1e6, 1))), (-([Math]::Round($dep / 1e6, 1))), (-([Math]::Round($imp / 1e6, 1))))
    }
    function Get-DemoPLBody($a, $b) {
        return ,@(@('Receita bruta', $a[0], $b[0]), @('Custo das mercadorias', $a[1], $b[1]), @('Despesas operacionais', $a[2], $b[2]),
                  @('EBITDA', 'e', 'e'), @('Depreciação e juros', $a[3], $b[3]), @('Impostos', $a[4], $b[4]), @('Lucro líquido', 'e', 'e'))
    }
    $plYearBody = Get-DemoPLBody (Get-DemoPL $allS $y24) (Get-DemoPL $allS $y25)
    $s2 = $storeOrder[1]                                                                             # a segunda loja em vendas de 2025
    $plStoreBody = Get-DemoPLBody (Get-DemoPL @($s1) $y25) (Get-DemoPL @($s2) $y25)
    ```
  - No `switch ($item.Num)` dos exemplos, depois do caso 26:
    ```powershell
            27 { $rng = Write-Table $ws 'A6' @('Etapa', '2024', '2025') $plYearBody '#,##0.0'
                 $co = Add-Chart $ws $rng 'waterfallpair'; Set-ChartBox $ws $co 660 340 }
            28 { $rng = Write-Table $ws 'A6' @('Etapa', $stores[$s1].Name, $stores[$s2].Name) $plStoreBody '#,##0.0'
                 $co = Add-Chart $ws $rng 'waterfallpair'
                 Invoke-CoreQuiet 'E28' 'Core_SetToggle' 'pairabs' $true | Out-Null
                 Invoke-CoreQuiet 'E28' 'Core_SetOption' 'legend' 'right' | Out-Null
                 Set-ChartBox $ws $co 700 340 }
    ```
  - No mapa `$featNames` do índice, acrescente `pairdiff = 'Setas de variação'; pairabs = 'Variação em valor'; inbar = 'Nome dentro das barras'`.
  - No `switch ($item.Num)` da unidade (`$unit`, logo antes de `New-ExampleSheet`), acrescente `{ $_ -in 27, 28 } { $unit = 'R$ milhões' }`.
  (Os dois em R$ milhões com 1 casa: em R$ mil os rótulos dos degraus ficariam mais largos que a fatia e sumiriam. O `Core_Insert` escolhe `dec1` pelo formato da 1.ª célula de valor, `ApplySourceNumFmt`.)

- [ ] **Passo 3: Teste da demonstração** — `tests/ps/Test-DemoWorkbook.ps1`:
  - Remova as 3 linhas do `$pendingDemo` (comentário e atribuição) e devolva as duas linhas de cobertura à forma original:
    ```powershell
    foreach ($k in $allKinds) { Check (@($cov | Where-Object { $_.Kind -eq $k }).Count -ge 1) "cobertura tem o tipo '$k'" }
    ```
    ```powershell
    foreach ($f in $matrixFeatures) { Check (@($cov | Where-Object { $_.Features -contains $f -or @($_.Features | Where-Object { $_ -like "$f-*" }).Count -gt 0 }).Count -ge 1) "cobertura tem o recurso da matriz '$f'" }
    ```
  - No `$featCheck`, depois de `'connectors'`:
    ```powershell
            'pairdiff'     = { param($i, $co, $ws) $n = 0; foreach ($sr in $co.Chart.SeriesCollection()) { if ([string]$sr.Name -eq 'tqq_pd') { $n++ } }; $i.Opt['pairdiff'] -eq '1' -and $n -eq 1 }
            'pairabs'      = { param($i, $co, $ws) $i.Opt['pairabs'] -eq '1' }
            'inbar'        = { param($i, $co, $ws) $i.Opt['legend'] -eq 'inbar' }
    ```
  - Depois da conferência do E05 (antes da do E02):
    ```powershell
        # E27: a seta do Lucro (7.º par, ponto 39 de tqq_pd) mostra (2025 - 2024) / 2024 do lucro, arredondado como o ROUND do Excel; as linhas
        # 'e' (EBITDA e Lucro) nao somam
        $e27 = $wb.Worksheets.Item('E27'); $c27 = $e27.ChartObjects(1).Chart
        $la = 0.0; $lb = 0.0; foreach ($rr in 7, 8, 9, 11, 12) { $la += [double]$e27.Cells.Item($rr, 2).Value2; $lb += [double]$e27.Cells.Item($rr, 3).Value2 }
        $pd = $null; foreach ($sr in $c27.SeriesCollection()) { if ([string]$sr.Name -eq 'tqq_pd') { $pd = $sr } }
        Check ($null -ne $pd) 'E27: setas de variação (tqq_pd)'
        if ($null -ne $pd -and $la -ne 0 -and $la * $lb -gt 0) {
            $r = [Math]::Round(100 * ($lb - $la) / $la, [MidpointRounding]::AwayFromZero)
            $want = $(if ($r -gt 0) { '+' } else { '' }) + ([string][int]$r) + '%'
            $got = [string]$pd.Points(39).DataLabel.Text
            Check ($got -eq $want) "E27: seta do lucro [$got] esperado [$want]"
        }
        $lblA = $null; foreach ($sr in $c27.SeriesCollection()) { if ([string]$sr.Name -eq 'tqq_lbl_a') { $lblA = $sr } }
        $nTxt = 0; if ($lblA) { for ($k = 1; $k -le $lblA.Points().Count; $k++) { try { if ([string]$lblA.Points($k).DataLabel.Text -ne '') { $nTxt++ } } catch { } } }
        Check ($nTxt -ge 2) "E27: rotulos dos degraus de 2024 visiveis (custo e despesas; $nTxt)"
    ```
    (Depreciação e impostos, ~2–3% da receita, continuam abaixo de 10–12 pt e ficam sem texto: é o comportamento previsto.)
  - Dentro do `try { ... } finally { $zip.Dispose() }` do começo do arquivo (logo antes do `finally`), a conferência dos conectores no XML do pacote:
    ```powershell
            # cascata comparativa: os conectores sao barras de erro X PERSONALIZADAS ligadas ao bloco (vivas sem o ThinQQ)
            $connXml = @($zip.Entries | Where-Object { $_.FullName -like 'xl/charts/chart*.xml' } | ForEach-Object { $sr = New-Object IO.StreamReader($_.Open()); $x = $sr.ReadToEnd(); $sr.Dispose(); $x } | Where-Object { $_ -match 'tqq_conn_a' })
            Check ($connXml.Count -ge 2) "cascata comparativa: graficos com conectores no pacote ($($connXml.Count))"
            foreach ($x in $connXml) { Check ($x -match '(?s)<c:errBars>.*?<c:errValType val="cust"/>.*?<c:plus><c:numRef><c:f>[^<]*_ThinQQ') 'conectores: barra de erro personalizada ligada a _ThinQQ' }
    ```

- [ ] **Passo 4: Gerar a pasta e conferir** —
  ```powershell
  robocopy "<repo>" C:\temp\tqq-cc /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS
  pwsh C:\temp\tqq-cc\build\build.ps1 -Target core -Configuration Test
  pwsh C:\temp\tqq-cc\tools\New-DemoWorkbook.ps1 -CoreXlam C:\temp\tqq-cc\dist\ThinQQ.core.test.xlam -OutFile "<repo>\exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx"
  ```
  Saída esperada: "Gerado: ... (N bytes) em ~45 s", código 0. Depois `pwsh "<repo>\tests\ps\Test-DemoWorkbook.ps1" -RepoRoot "<repo>"` → "ok" em tudo, código 0.

- [ ] **Passo 5: Tutorial, README e galeria** —
  - `#demonstracao`: depois da linha `data-ex="E26"`:
    ```html
          <tr data-ex="E27"><th scope="row">E27</th><td>Da receita ao lucro, 2024 x 2025 (cascata comparativa)</td><td>Cascata comparativa</td></tr>
          <tr data-ex="E28"><th scope="row">E28</th><td>As duas maiores lojas: da receita ao lucro em 2025 (variação em valor)</td><td>Cascata comparativa</td></tr>
    ```
  - Seção `#cascata-comparativa`: antes do `</section>`, `<p>Veja as figuras <a href="#ex-E27">E27</a> (percentual, nomes dentro das barras) e <a href="#ex-E28">E28</a> (variação em valor, legenda) na galeria.</p>`.
  - Contagens da demonstração: no tutorial, "e 26 exemplos, de E01 a E26, um por funcionalidade: os 20 tipos de gráfico" → "e 28 exemplos, de E01 a E28, um por funcionalidade: os 21 tipos de gráfico", o `aria-label` "Os 26 exemplos da pasta de demonstração" → "Os 28 exemplos da pasta de demonstração" e, na seção de Exemplos, "As imagens dos 26 exemplos da pasta" → "As imagens dos 28 exemplos da pasta"; no README, "um índice com 26 exemplos, de E01 a E26" → "28 exemplos, de E01 a E28", "a imagem de cada um dos 26 exemplos" → "28 exemplos", "com 26 exemplos (E01 a E26)" → "com 28 exemplos (E01 a E28)", "os 26 exemplos da pasta de demonstração" → "os 28 exemplos" e, na verificação manual 38, "os 26 exemplos mostram" → "os 28 exemplos mostram". Confira com `grep -n "26 exemplos\|E26[^0-9]" README.md docs/tutorial/tutorial.html`: nenhuma ocorrência (as do CHANGELOG ficam).
  - Comentários de `tools/Export-DemoChartImages.ps1` e `tools/Build-TutorialGallery.ps1`: "E01..E26"/"E01.png ... E26.png" → "E01..E28"/"E01.png ... E28.png".
  - Galeria:
    ```powershell
    pwsh "<repo>\tools\Export-DemoChartImages.ps1" -OutDir C:\temp\tqq-cc-img
    pwsh "<repo>\tools\Build-TutorialGallery.ps1" -ImageDir C:\temp\tqq-cc-img
    Remove-Item C:\temp\tqq-cc-img -Recurse -Force
    ```
    Saídas esperadas: "29 gráficos exportados" (E25 com dois); "galeria gravada: 28 figuras, 29 imagens, ...". Abra `C:\temp\tqq-cc-img\E27.png` e `E28.png` antes de apagar e confira as pontes, as elipses e os nomes "2024"/"2025" dentro das barras da receita (E27) e a legenda com as duas lojas (E28); registre no relatório.

- [ ] **Passo 6: Suíte completa** — `robocopy "<repo>" C:\temp\tqq-cc /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS` (o tutorial, o README e o xlsx mudaram no repositório real) e `pwsh C:\temp\tqq-cc\build\run-tests.ps1 -TimeoutSec 900` verde (Test-Docs: 28 figuras, "28 exemplos", "E01 a E28"; Test-DemoWorkbook sem SKIP).

- [ ] **Passo 7: Commit** — `git add tools/DemoCoverage.ps1 tools/New-DemoWorkbook.ps1 tools/Export-DemoChartImages.ps1 tools/Build-TutorialGallery.ps1 tests/ps/Test-DemoWorkbook.ps1 exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx docs/tutorial/tutorial.html README.md` e `git commit -m "feat: demonstracao E27 e E28 (cascata comparativa) e galeria do tutorial regerada"` (+ trailer).

**Verificação no Excel real:** abrir a pasta de demonstração sem o ThinQQ: E27 e E28 aparecem iguais à galeria; mudar um número da tabela do E27 muda a elipse e o "(x%)" do lucro.

---

### Tarefa 8: Tutorial, README, CHANGELOG e notas da 1.3.0 — modelo: standard

**Files:**
- Modify: `docs/tutorial/tutorial.html`, `README.md`, `CHANGELOG.md`, `release/notes.txt`, `tests/ps/Test-Docs.ps1` (mensagens e limites)

**Interfaces:**
- Consumes: rótulos exatos da faixa (T5), mensagens 6016/6017 (T1), comportamento das Tarefas 1 a 7.
- Produces: a seção `#cascata-comparativa` completa; README (Recursos 1.3.0, tabela da galeria, verificação manual 42, limitação); CHANGELOG `[Não lançado]` com "Adicionado"; `release/notes.txt` numa das duas variantes do Passo 4 (pela tag `v1.2.3`); `Test-Docs` com as mensagens e as frases-chave novas.

- [ ] **Passo 1: Test-Docs (RED)** — em `tests/ps/Test-Docs.ps1`, na lista `$messages` acrescente (antes do fecho `)`):
  ```powershell
    'A cascata comparativa precisa de exatamente 2 colunas de valores (uma por série); a tabela tem só 1.',
    'A cascata comparativa compara exatamente 2 séries; a tabela tem ',
    ' colunas de valores. Selecione a coluna das etapas e só as duas colunas a comparar.'
  ```
  (a vírgula vai depois do último item que já existia) e, no `$mustDocument`, antes do fecho `}`:
  ```powershell
    'comparativa: percentual (B - A) / A'                            = @('both', '(B - A) / A')
    'comparativa: n/d com zero ou sinais opostos'                    = @('both', 'sinais opostos')
    'comparativa: conectores atravessam a barra da outra série'      = @('tutorial', 'atravessam a barra da outra série')
    'comparativa: afastamento das setas só com o ThinQQ'             = @('both', 'afastamento das setas')
    'comparativa: sem cálculo de trás para frente'                   = @('tutorial', 'não calcula para trás')
    'comparativa: no máximo 2 séries'                                = @('both', 'mais de 2 séries')
  ```
  Logo depois de `Check ($matrix.Count -eq 23) ...` (dentro do `if ($parsed)`), acrescente a conferência das contagens de tipos citadas no tutorial:
  ```powershell
      $nk = $allKinds.Count
      Check ($html -match "\b$nk tipos" -and $html -notmatch "\b$($nk - 1) tipos") "tutorial cita os $nk tipos (e nenhum '$($nk - 1) tipos')"
  ```
  Rode `pwsh C:\temp\tqq-cc\tests\ps\Test-Docs.ps1 -RepoRoot C:\temp\tqq-cc` (depois do robocopy) → falhas só nas frases e mensagens novas e na conferência dos "21 tipos".

- [ ] **Passo 2: Tutorial** — em `docs/tutorial/tutorial.html`, troque o conteúdo da seção `#cascata-comparativa` (mantenha o `<h2>`, a `<dl class="ficha">` da Tarefa 5 e o parágrafo das figuras E27/E28 da Tarefa 7) acrescentando, depois da `</dl>`:
  ```html
  <h3>Passo a passo</h3>
  <ol>
    <li>Monte a tabela: etapa na primeira coluna (Bruto, Componente 1, ..., Líquido) e as duas séries nas colunas seguintes (Segmento 1, Segmento 2; ou 2024, 2025). Para uma barra de total calculada escreva <code>e</code> no lugar do valor, em cada série.</li>
    <li>Selecione a tabela com o cabeçalho e clique em <span class="ui">Cascata</span> › <span class="ui">Cascata comparativa</span>.</li>
    <li>O gráfico nasce com as duas pontes lado a lado (a primeira série à esquerda de cada etapa), conectores tracejados entre as barras da mesma série, a seta de variação em cada etapa, os rótulos e os nomes das séries escritos na vertical dentro das barras da primeira etapa.</li>
  </ol>
  <div class="tw fit">
  <table>
    <thead><tr><th scope="col">Etapa</th><th scope="col">Segmento 1</th><th scope="col">Segmento 2</th><th scope="col">Seta</th></tr></thead>
    <tbody>
      <tr><td>Bruto</td><td>27</td><td>29</td><td>+7%</td></tr>
      <tr><td>Componente 1</td><td>-4</td><td>-7</td><td>+75%</td></tr>
      <tr><td>Componente 2</td><td>-6</td><td>-2</td><td>-67%</td></tr>
      <tr><td>Componente 3</td><td>-6</td><td>-10</td><td>+67%</td></tr>
      <tr><td>Líquido</td><td>e</td><td>e</td><td>-9%</td></tr>
    </tbody>
  </table>
  </div>
  <p>É o exemplo 27 de "Exemplos 3". Os totais mostram o valor e o percentual em relação à primeira barra da própria série: <code>11 (41%)</code> e <code>10 (34%)</code>.</p>

  <h3>Como a seta calcula</h3>
  <ul>
    <li>A seta vai da barra da primeira série (A) à da segunda (B) da mesma etapa e mostra <code>(B - A) / A</code> sobre o valor da barra: o degrau nos degraus e o total nas barras de total. Num custo (valores negativos), "+" quer dizer que o custo cresceu: de -4 para -7 é <code>+75%</code>. Com A e B do mesmo sinal, é a variação do tamanho da barra: <code>(|B| - |A|) / |A|</code>.</li>
    <li>Quando A vale zero ou A e B têm sinais opostos, a seta mostra <code>n/d</code> (o percentual não tem leitura).</li>
    <li><span class="ui">Setas</span> › <span class="ui">Variação em valor</span> troca o percentual pela diferença <code>B - A</code> no formato do gráfico (com <span class="ui">Formato</span> em %, em pontos percentuais: <code>+3 p.p.</code>). <span class="ui">Setas de variação</span> liga e desliga as setas.</li>
  </ul>

  <h3>Opções</h3>
  <ul>
    <li><span class="ui">Totais</span> liga e desliga o percentual do primeiro total nos rótulos das barras de total; <span class="ui">Rótulos</span> liga e desliga todos os rótulos das barras; <span class="ui">Conectores</span>, <span class="ui">Formato</span>, <span class="ui">Paleta</span> (a primeira série na cor mais escura) e <span class="ui">Espaço</span> (entre as duas barras do par) funcionam como na cascata.</li>
    <li><span class="ui">Legenda</span>: <span class="ui">Nome dentro das barras</span> (o padrão) escreve o nome de cada série na primeira barra quando ele cabe; se não couber, o ThinQQ mostra a legenda à direita no lugar. <span class="ui">Direita</span> e <span class="ui">Nenhuma</span> funcionam como nos outros tipos.</li>
    <li><span class="ui">Mesma escala</span> funciona entre cascatas comparativas (e com outros tipos de escala livre).</li>
  </ul>

  <h3>Limites</h3>
  <ul>
    <li>Exatamente 2 séries: com 1 coluna de valores o ThinQQ avisa "A cascata comparativa precisa de exatamente 2 colunas de valores (uma por série); a tabela tem só 1."; com 3 ou mais, "A cascata comparativa compara exatamente 2 séries; a tabela tem 3 colunas de valores. Selecione a coluna das etapas e só as duas colunas a comparar." (com o número de colunas da tabela). Comparar mais de 2 séries lado a lado não existe nesta versão.</li>
    <li>Os conectores de uma série atravessam a barra da outra série que fica entre elas (é como o think-cell desenha); desligue os <span class="ui">Conectores</span> se preferir.</li>
    <li>Com <code>e</code> na primeira linha a primeira barra é zero: a cascata comparativa não calcula para trás (a cascata comum calcula).</li>
    <li>Barras, conectores, setas, o texto da elipse e os percentuais são fórmulas da planilha: acompanham os números mesmo sem o ThinQQ. O afastamento das setas em relação às barras, os rótulos de degraus finos demais (que ficam sem texto) e a troca entre o nome dentro das barras e a legenda são recalculados pelo ThinQQ ao mudar os dados ou o tamanho; sem ele, ficam como na última vez.</li>
    <li>Só na vertical (colunas); os rótulos dos degraus mostram o valor com sinal, como na cascata do ThinQQ.</li>
    <li>Os textos dos totais e das setas seguem os separadores do Excel que calculou a pasta por último: aberta num Excel em outro idioma, eles se ajustam depois de Ctrl+Alt+F9 (os rótulos dos degraus se ajustam sozinhos).</li>
    <li>Uma linha acrescentada depois da última etapa da tabela não é detectada: use <span class="ui">Alterar dados</span>.</li>
  </ul>
  ```
  Na lista de bullets abaixo da matriz (`#matriz`), acrescente `<li>A cascata comparativa usa <span class="ui">Totais</span> para o percentual do primeiro total nos rótulos das barras de total; <span class="ui">Setas de variação</span>, <span class="ui">Variação em valor</span> e <span class="ui">Nome dentro das barras</span> só existem nela.</li>`. Na seção `#limites`, acrescente um item `<li><strong>Cascata comparativa:</strong> exatamente 2 séries (mais de 2 séries lado a lado não existe); o afastamento das setas só se recalcula com o ThinQQ; com sinais opostos a seta mostra n/d. Veja a seção <a href="#cascata-comparativa">Cascata comparativa</a>.</li>`. Troque no tutorial `cinco menus com 20 tipos` → `cinco menus com 21 tipos`, `reúnem 20 tipos` → `reúnem 21 tipos`, `mostram 17 dos 20 tipos` → `mostram 18 dos 21 tipos` e `que mostra os 20 tipos` → `que mostra os 21 tipos`.

- [ ] **Passo 3: README** —
  - "Recursos": depois do bloco "**Fase 3 (1.2.0)**", acrescente
    ```markdown
    **Cascata comparativa (1.3.0)**

    - **Cascata comparativa** (ThinQQ > Cascata > Cascata comparativa): duas pontes lado a lado, uma por série (ex.: 2024 x 2025), com conectores por série, uma seta de variação em cotovelo entre as duas barras de cada etapa com `(B - A) / A` numa elipse (ou `n/d` com o primeiro valor zero ou sinais opostos; ThinQQ > Setas > Variação em valor mostra a diferença), totais com o percentual do primeiro total ("11 (41%)") e os nomes das séries dentro das barras (ThinQQ > Legenda > Nome dentro das barras) ou na legenda. Tudo em séries nativas ligadas a fórmulas: acompanha os números mesmo sem o ThinQQ; só o afastamento das setas e os rótulos finos são recalculados por ele.
    ```
  - Tabela "Galeria de gráficos": a linha "Cascata" passa a `| Cascata | Cascata (colunas), Cascata (barras), Cascata comparativa | etapa e 1 a 6 colunas de valores (a comparativa: exatamente 2); `e` marca um total |`.
  - "Verificações manuais": acrescente o item `42. Cascata comparativa com a tabela da imagem do think-cell (Bruto 27/29 até Líquido e/e): ThinQQ > Cascata > Cascata comparativa mostra as duas pontes, as elipses +7%, +75%, -67%, +67%, -9%, "11 (41%)" e "10 (34%)" e os nomes dentro das barras; ThinQQ > Setas > Variação em valor troca as elipses pela diferença; Legenda > Direita troca os nomes pela legenda; redimensionar mantém as setas acima dos rótulos; abrir a pasta sem o ThinQQ e mudar um número muda as elipses.`
  - "Limitações conhecidas": acrescente `- **Cascata comparativa:** exatamente 2 séries (mais de 2 séries lado a lado não existe); só na vertical; com o primeiro valor zero ou sinais opostos a seta mostra n/d (percentual (B - A) / A); o afastamento das setas e os rótulos finos só se recalculam com o ThinQQ; os conectores passam por cima da barra da outra série.`
  - Na verificação manual 37 ("Ícones próprios"), depois de `Cascata,` acrescente `Cascata comparativa,` e troque `Setas e PowerPoint` por `Setas (com Setas de variação e Variação em valor) e PowerPoint`.

- [ ] **Passo 4: CHANGELOG e notas** — `CHANGELOG.md`, logo abaixo de `## [Não lançado]` (antes do `### Corrigido` da 1.2.3, se ele ainda estiver lá; não o mexa):
  ```markdown
  ### Adicionado
  - Cascata comparativa (menu Cascata): duas pontes lado a lado, uma por série, com conectores tracejados por série, uma seta de variação em cotovelo entre as duas barras de cada etapa (`(B - A) / A` numa elipse; `n/d` com o primeiro valor zero ou sinais opostos; ou a diferença em valor pelo alternador Setas > Variação em valor), totais com o percentual do primeiro total da série ("11 (41%)") e os nomes das séries dentro das barras da primeira etapa (Legenda > Nome dentro das barras; quando não cabem, a legenda). Exatamente 2 colunas de valores; marcadores `e`, `t`, `=`, `total`, `subtotal`. Tudo em séries nativas ligadas ao bloco da `_ThinQQ`: barras, conectores, setas e textos acompanham os números mesmo sem o ThinQQ.
  - Exemplo 27 em ThinQQ > Exemplos (planilha Exemplos 3) e exemplos E27 e E28 na pasta de demonstração, com as imagens na galeria do tutorial.

  ### Alterado
  - Faixa de opções: Cascata ganha "Cascata comparativa" (ícone próprio); Setas ganha "Setas de variação" e "Variação em valor"; Legenda ganha "Nome dentro das barras"; o alternador Totais, na cascata comparativa, liga o percentual do primeiro total. A matriz de capacidades passa a 21 tipos x 23 recursos.
  ```
  `release/notes.txt` (limite de 1 a 8 linhas): confira primeiro se a 1.2.3 já foi publicada com `git ls-remote --tags origin` (procure `refs/tags/v1.2.3`; só leitura). Não decida pela publicação da 1.2.3. Grave o arquivo inteiro com uma das duas variantes, exatamente (UTF-8, como o arquivo atual):
  - **a tag `v1.2.3` NÃO existe** (a correção do download sai junto com a 1.3.0; também se o comando falhar), 3 linhas:
    ```text
    • Novo: Cascata comparativa (menu Cascata): duas séries lado a lado, com conectores, setas de variação entre elas e o percentual do primeiro total.
    • Atualização pelo Excel volta a baixar os arquivos do GitHub (antes parava com 'HTTP 0').
    • Quem tem a 1.2.0, 1.2.1 ou 1.2.2 precisa reinstalar uma vez pelo ZIP desta versão (o carregador antigo não consegue baixar a atualização); daí em diante a atualização pelo Excel funciona.
    ```
  - **a tag `v1.2.3` existe** (a correção já saiu na 1.2.3), 1 linha:
    ```text
    • Novo: Cascata comparativa (menu Cascata): duas séries lado a lado, com conectores, setas de variação entre elas e o percentual do primeiro total.
    ```
  Registre no relatório a saída do `git ls-remote` e a variante gravada.

- [ ] **Passo 5: GREEN e suíte completa** — `Test-Docs` verde; `pwsh C:\temp\tqq-cc\build\run-tests.ps1 -TimeoutSec 900` verde. Abra o tutorial no navegador (arquivo local) e confira a seção nova e a 4.ª tabela da matriz.

- [ ] **Passo 6: Commit** — `git add docs/tutorial/tutorial.html README.md CHANGELOG.md release/notes.txt tests/ps/Test-Docs.ps1` e `git commit -m "docs: cascata comparativa no tutorial, README, CHANGELOG e notas da 1.3.0"` (+ trailer).

---

### Tarefa 9: Verificação final — modelo: standard

**Files:** nenhum arquivo de produção (só correções pequenas que a verificação achar, cada uma com teste e commit próprio).

- [ ] **Passo 1: Suíte completa com troca de loader** — `robocopy ...` e `pwsh C:\temp\tqq-cc\build\run-tests.ps1 -TimeoutSec 900 -LoaderSwap` → "Todos os testes passaram", código 0; anote os totais (core/loader) e o tempo. Anote também a soma dos testes VBA (linhas `--- core : os 15 testes mais lentos (de N; soma S s) ---` e `--- loader ...`); se core + loader passar de 480 s, suba o padrão de `-TimeoutSec` de `build/run-tests.ps1` de 600 para 900 (commit `build:` próprio) e registre.
- [ ] **Passo 2: Conferência visual** — num Excel próprio e invisível (`New-ExcelApp`), abra uma cópia da pasta de demonstração e exporte E27, E28 e a seção 27 dos Exemplos (pelo core de Teste da cópia: `Samples_Create`) para PNG numa pasta sua (`C:\temp\tqq-cc-vis`); olhe cada PNG: pontes, nomes das etapas sob os pares, conectores, elipses, totais com percentual, nomes dentro das barras (E27) e legenda (E28); um gráfico de 300 × 220 pt depois de `Core_SafeRefresh`. Registre no relatório o que viu e apague a pasta.
- [ ] **Passo 3: Auditoria** — `grep -rn "26 exemplos\|E01 a E26" README.md docs/tutorial/tutorial.html` vazio; `git diff <base>..HEAD --stat` só com os arquivos da "Estrutura de arquivos" (e `build/run-tests.ps1`, se o Passo 1 subir o padrão); nenhum arquivo em `C:\temp` deixado pelas tarefas (`C:\temp\tqq-cc*`); nenhum processo Excel além dos que existiam no começo (compare com a lista anotada).
- [ ] **Passo 4: Relatório** — o que foi feito por tarefa, desvios do plano com o motivo, medidas (tempo da suíte, tamanho do tutorial e da pasta de demonstração), riscos que continuam abertos (spec §9) e os itens manuais do README que ficaram para o usuário.

---

## Autorrevisão (feita ao escrever o plano)

- **Cobertura do brief:** pesquisa e sondas (spec §2–§3); tipo e recusas (T0, T1); construção nativa medida (T1); conectores vivos (T2); setas com elipse, percentual/valor, sala calibrada com teto e convergência (T3); rótulos, totais com percentual, formato, rótulos finos com contabilidade sem lista (T4); faixa, ícone e `getEnabled` pela matriz (T5); Alterar dados, todos os tipos, limpeza, Exemplos (T1, T6); demonstração E27/E28 e galeria (T7); tutorial, README, CHANGELOG, notas (T5–T8); verificação (T9); fora do escopo (spec §10).
- **Placeholders:** nenhum "a definir"; todo código das tarefas de VBA está completo; os trechos de documentação trazem o texto final.
- **Consistência de nomes:** `Wfp_SetConnectors`, `Wfp_SetArrows`, `Wfp_SetArrowMode`, `Wfp_SetLabels`, `Wfp_SetTotalPct`, `Wfp_TotalPctOn`, `Wfp_ApplyNumberFormat`, `Wfp_ApplyCellFormat`, `Wfp_SetLegend`, `Wfp_SyncNames`, `Wfp_SyncGap`, `Wfp_RefreshAfterDataChange`, `Wfp_SourceRowsSpanned` são os mesmos nas tarefas que os produzem e nas que os consomem (T2–T6); séries `tqq_*` iguais em código, testes e Test-DemoWorkbook; recursos `pairdiff`, `pairabs`, `inbar` iguais na matriz, na faixa, no tutorial e na cobertura.
- **Ordem das linhas em `Wfp_Insert`** depois da T4: `Wfp_ApplyColors`, `Wfp_AddZeroLine`, `Wfp_SetNames`, `Wfp_SetLabels` (T4), `Wfp_SetConnectors` (T2), `Wfp_SetArrowMode`/`Wfp_SetArrows` (T3), `Wfp_SetLegend`, `Wfp_RefreshAfterDataChange`. Em `Wfp_RefreshAfterDataChange`: `Wfp_SyncGap`, `Wfp_SyncNames`, `Wfp_RefreshRoom` (T3), `Wfp_SyncLabels` (T4).
- **Suíte verde por tarefa:** T0 marca a demonstração como pendente (`$pendingDemo`) até a T7; os rótulos novos ficam em `k` até a T5; as contagens de exemplos mudam uma de cada vez (T6: 27, T7: 28).
