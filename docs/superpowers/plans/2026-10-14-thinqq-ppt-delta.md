# ThinQQ — Plano de Implementação: PowerPoint atualizado ao reabrir (1.4.0) e delta entre barras clicando nas barras (1.5.0)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** (Parte A, 1.4.0) gráficos do ThinQQ colados no PowerPoint como gráfico vinculado passam a ser atualizados — dados **e** anotações (setas, rótulos de seta, colchetes, conectores, marcas de corte) — ao reabrir a apresentação (suplemento `ThinQQ.ppam`) ou pelo botão "Atualizar gráficos no PowerPoint" no Excel, a colagem deixa de esticar as setas e o suplemento do PowerPoint é **atualizado automaticamente pelo atualizador do Excel** junto com o ThinQQ (emenda 1: Tarefa A7); (Parte B, 1.5.0) o usuário insere uma seta de **delta** entre quaisquer duas barras clicando nelas (origem e destino), com texto valor/percentual/ambos, várias setas sem colisão, recalculadas com os dados.

**Architecture:** A das Fases 1 a 3 e da cascata comparativa: `ThinQQ.xlam` (loader: faixa) chama `Core_*` em `ThinQQ.core.xlam`; gráficos nativos ligados a blocos da planilha muito oculta `_ThinQQ`; anotações são formas `tqq_ann_*` dentro do gráfico, recalculadas pelo `Ann_Refresh`. Novidades: módulo `modPptSync` (motor de sincronização: `LinkFormat.Update` + o gráfico do Excel no tamanho do slide + foto das formas aplicada **no lugar** no gráfico do slide), um terceiro projeto VBA `src/ppt/` (`ThinQQ.ppam`: eventos do PowerPoint chamam `Core_PptSyncPresentation` num Excel com o ThinQQ), módulos `modDelta` (referências "k:c", extensões, textos) e `modDeltaPick` (modo armado de dois cliques pelos eventos `Select`/`MouseUp` do gráfico).

**Tech Stack:** VBA 7 (Excel e PowerPoint 2016+ Windows, 32/64 bits, sem `Declare`, sem referências extras), customUI14, PowerShell 7 e 5.1, Git.

**Spec:** `docs/superpowers/specs/2026-10-14-thinqq-ppt-delta-design.md` (cópia, feita na Tarefa A0, do rascunho `.superpowers/sdd/2026-10-14-ppt-delta/design.md`: pesquisa, sondas §3, decisões §4 e §5, matriz §6, faixas §7, ordem §8, riscos §9, pré-flight §11, estimativa §12). Specs anteriores em `docs/superpowers/specs/`. Regras de ambiente: `.superpowers/sdd/2026-10-13-cascata-comparativa/implementer-rules.md` (valem aqui, trocando a cópia de trabalho por `C:\temp\tqq-pd`).

**Estado de execução:** nenhuma tarefa concluída. Duas partes **independentes**, cada uma numa branch criada pelo controlador a partir da `main` (`feat/ppt-atualizar` para a Parte A, `feat/delta-barras` para a Parte B); o implementador **não** cria nem troca de branch. Ordem recomendada: **Parte A primeiro** (spec §8.1: corrige o defeito medido da colagem atual e é a base para os deltas aparecerem certos no PowerPoint). Se a Parte B for executada antes, a Tarefa B0 promove o spec (se a A0 ainda não o fez) e o teste condicional de PowerPoint da B2 é omitido (está marcado).

**Pré-requisito do usuário (antes da Tarefa A4):** habilitar no **PowerPoint** "Arquivo > Opções > Central de Confiabilidade > Configurações da Central de Confiabilidade > Configurações de Macro > Confiar no acesso ao modelo de objeto do projeto VBA" (a chave `HKCU\Software\Microsoft\Office\16.0\PowerPoint\Security\AccessVBOM` não existe nesta máquina; medido). O implementador **nunca** grava o registro. Sem isso a A4 para com `BLOCKED` e as Tarefas A1–A3 continuam valendo (o botão do Excel é a alternativa (a) completa: spec §4.1).

## Global Constraints

- **Ambiente:** Excel 2019 32-bit pt-BR e PowerPoint 16.0 32-bit no Windows 11; testes em instâncias próprias invisíveis. O VBA compila sob demanda: o build roda o lint (`build/lib/Vba-Lint.ps1`) e uma compilação completa; erro de compilação ⇒ saída 3 (módulo:linha); travamento morto pelo watchdog ⇒ saída 2; asserts ⇒ 1; ok ⇒ 0. Fontes VBA em UTF-8 com **somente caracteres cp1252** (acentos sim; nada de `→`, `≥`, `−`, `…`, aspas tipográficas), `Option Explicit`, declarações antes dos procedimentos, `Attribute VB_Name` na primeira linha, nenhuma continuação ` _` seguida de texto, nada de `ElseIf` num `If` de uma linha, nada de `If ... Then If ...` numa linha.
- **Testes:** `Public Sub Test_x()` sem parâmetros em `tests/vba/{core,loader,ppt}/Test*.bas`, despachados por módulo gerado (nunca `Application.Run` para os testes), via `pwsh build/run-tests.ps1` (`-Target core|loader|all`, `-NoBuild`, `-TimeoutSec`, `-Filter "Modulo.Teste|outro"` — pula as suítes PowerShell —, `-LoaderSwap`, e o novo `-Ppt` da Tarefa A4). Asserts: `AssertTrue/False/Equal/Almost`, `NewTempWorkbook`, `CloseTemp`, `TempDir`, `ForceLayout`, `RepoPath`. Cada módulo de teste novo define os próprios `Fixture`, `Teardown` (`Watchers_Stop` + `CloseTemp`) e helpers privados.
- **Instâncias do usuário:** os processos Excel/PowerPoint abertos quando a tarefa começa (anote `Get-Process excel,powerpnt -ErrorAction SilentlyContinue | Select-Object Id, StartTime`) nunca são tocados, anexados (`GetObject`) nem encerrados. O PowerPoint é de **instância única**: testes que criam PowerPoint só rodam sem nenhum `POWERPNT` (padrão `StartScratchPowerPoint` de `tests/vba/core/TestPowerPoint.bas`); com um PowerPoint aberto eles saem verdes registrando "pulado" e o relatório da tarefa diz isso. Builds e testes rodam numa **CÓPIA**: `robocopy "<repo>" C:\temp\tqq-pd /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS` (repita antes de cada rodada; edite e commite no repositório real). Nunca `build.ps1 -Install`; nunca `%LOCALAPPDATA%\ThinQQ`, registro real, lojas de certificados. **Disco e memória:** ~1 GB livre de memória: nunca dois Excel de teste ao mesmo tempo; apague `C:\temp\tqq-pd` e `C:\temp\tqq-pd-*` ao terminar (em `C:\temp` use `rm -rf` do bash nos seus caminhos exatos); nada fora delas. Depois de cada execução confira `git status` (arquivos versionados já sumiram por quarentena do antivírus: restaure com `git checkout -- <arquivo>` e relate).
- **Código:** `Core_*` nunca propaga erro (`Errors_Handle`/`Errors_HandleErr`; `On Error GoTo -1` antes de `On Error Resume Next` dentro de um tratador); `MsgBox` só em `Core_Notify`/`Core_Confirm` (log-only no Test), `RxMsg`, `Loader_Msg`, formulários e, no `.ppam`, `PptTqq_Msg`/`PptTqq_Ask` (log-only no Test); ganchos de teste só agem com `BUILD_CONFIG = "Test"` e são zerados em todo caminho (o teste zera os que armou, mesmo em falha); laços de geometria com teto de passadas; fórmulas em inglês; textos de interface em pt-BR; ligação tardia com Excel/PowerPoint de outro processo (`As Object`), constantes numéricas para enums do outro aplicativo.
- **Armadilhas medidas (spec §3; não redescubra):** `LinkFormat.Update` com a pasta de origem inexistente **trava** o PowerPoint (diálogo modal invisível): confira o arquivo antes; com a pasta aberta num Excel, o `Update` lê os valores **não salvos** (60–280 ms), fechada abre um `EXCEL -Embedding` (~3,7 s); abrir o .pptx por automação **não** atualiza vínculos (nem os automáticos); `LinkFormat.AutoUpdate` padrão da colagem = 1 (manual); redimensionar o gráfico do slide **estica** as formas internas; com o mesmo tamanho a área interna do PowerPoint e a do Excel são idênticas (`ChartArea.Width` do Excel = ChartObject − 5; `Point.Left` do PowerPoint = Excel + 4; formas internas nas mesmas coordenadas); `Shapes.Paste` + `ApplyAnimation` manda a animação para o fim e muda a ordem Z (atualize **no lugar**); `Tags` persistem no .pptx; `Application.Visible = False` é recusado pelo PowerPoint (use `Presentations.Add(0)`/`Open(..., WithWindow:=0)`); `Points(i).Select` dispara `Chart_Select(3, série, ponto)` (com `Deactivate`/`Activate` antes), a série inteira dá `Arg2 = -1`, uma forma do gráfico dá `ElementID = 14`; `GetChartElement` funciona no Excel invisível com px = pt × 4/3 (zoom 100%); a geometria de rótulos e pontos só vale depois de `ForceLayout`; `IIf` avalia os dois ramos; PowerShell: `$R` = `$r`, vírgula antes da aritmética em `@(...)`, função devolve no pipeline o que não for atribuído.
- **Commits:** pequenos, em português, prefixos `feat:`/`fix:`/`test:`/`docs:`/`build:`, terminando com a linha `Co-Authored-By:` indicada no prompt do implementador ("(+ trailer)" nos comandos abaixo). **Não** fazer push, **não** criar nem trocar de branch.
- **Versões:** `version.json` não muda (o `publicar-release.bat 1.4.0`/`1.5.0` sobe loader e core). O build importa todo `.bas` de `src/<alvo>/modules` e todo `Test*.bas` de `tests/vba/<alvo>`.
- **Processo:** cada tarefa é executada por um implementador novo que só vê o texto da tarefa; cada tarefa termina com a suíte completa verde na cópia (`pwsh C:\temp\tqq-pd\build\run-tests.ps1 -TimeoutSec 900`; a partir da A4, com `-Ppt` nas tarefas da Parte A que mexem no PowerPoint) e um commit (ou mais). Se a suíte completa morrer com 0x800706BE (memória), rode o core em partes com `-Filter` e relate. Modelo sugerido: **standard** (código com testes), **most-capable** (integração com o PowerPoint, geometria, eventos), **cheap** (edições mecânicas).

## Review Focus

0. **Suplemento atualizado sozinho pelo Excel** (Parte A): só para quem já tem o `.ppam`; baixado e conferido (SHA-256) antes de trocar o core; troca atômica; com o PowerPoint segurando o arquivo, pendente e aplicado sem ação do usuário; o aviso diz o que aconteceu (nunca um par core/suplemento trocado pela metade em silêncio). Testes: `Test_Updater_Ppam_*`, `Test_Updater_Apply_CoreOk_PpamFails_IsReported` (A7).

1. **Colagem vinculada com setas, ajustada ao slide** (gráfico de colunas empilhadas com seta CAGR colado e esticado a 709,7 × 432 pelo PowerPoint): o usuário espera a seta e o rótulo exatamente sobre as colunas do slide, não esticados. Teste: `Test_PptSync_PastedChart_AnnotationsAlignedToSlide_WhenAvailable` (A1).
2. **Reabrir a apresentação depois de mudar os dados** (pasta aberta com alteração não salva, ou fechada): dados, rótulo da seta e posição novos; animação, nome, ordem Z do gráfico do slide intactos; seta nova no Excel aparece no slide; seta apagada some. Testes: `Test_PptSync_DataChangeNewArrowAndAnimation_WhenAvailable` (A1), `Test_PptSync_ClosedSource_OpenedReadOnlyAndClosed_WhenAvailable` (A2), `Test-PptAddin.ps1` (A4).
3. **Pasta de origem movida**: nada trava, o gráfico fica, o aviso diz qual arquivo e como religar. Teste: `Test_PptSync_MissingSource_DoesNotHang_WhenAvailable` (A2).
4. **Delta entre barras de séries e categorias diferentes, vários no mesmo gráfico** (colunas agrupadas, 1.ª série T1 → 2.ª série T3, e outro aninhado): cotovelo acima de todas as barras do trecho, ponta no destino, texto "−14 (−140%)", setas aninhadas em faixas sem encostar. Testes: `Test_Delta_Columns_ElbowGeometry`, `Test_Delta_Lanes_NestedArrows` (B2).
5. **Dois cliques reais**: 1.º clique já escolhe a barra (MouseUp + GetChartElement), o mesmo clique pelo `Select` não conta duas vezes, clique numa célula cancela, desarmado nada acontece. Testes: `Test_DeltaPick_*` (B3) e a verificação manual da B6.

---

## Estrutura de arquivos (novos e alterados)

```
docs/superpowers/specs/2026-10-14-thinqq-ppt-delta-design.md   NOVO (A0: cópia do design)
docs/superpowers/plans/2026-10-14-thinqq-ppt-delta.md           NOVO (A0: cópia deste plano)
--- Parte A (1.4.0) ---
src/core/modules/modPptSync.bas        NOVO  PptSync_* (A1 motor; A2 apresentação, resumo)
src/core/modules/modCoreApi.bas        ALT   colagem marca e sincroniza (A1); Core_PptSyncPresentation, Core_PptSyncOpen (A2)
src/loader/ribbon/customUI14.xml       ALT   menu PowerPoint sempre ativo + "Atualizar gráficos no PowerPoint" (A3)
src/loader/modules/modRibbon.bas       ALT   Rx_OnPptSync (A3)
src/loader/modules/modUpdater.bas, modSettings.bas, modLoaderMain.bas, forms/frmUpdate.*  ALT  atualização automática do .ppam (A7)
src/ppt/modules/modPptTqq.bas, modPptRibbon.bas; src/ppt/classes/clsPptEvents.cls; src/ppt/ribbon/customUI14.xml   NOVOS (A4)
build/build.ps1, build/lib/Build-Common.ps1, build/run-tests.ps1   ALT (A4: alvo ppt, VBOM do PowerPoint, -Ppt, AllTests no despacho)
build/release.ps1, build/lib/Release-Lib.ps1, tools/lib/Release-Publish.ps1   ALT (A6)
installer/install.ps1, uninstall.ps1, Instalar-ThinQQ.cmd, LEIA-ME.txt   ALT (A5)
tests/vba/core/TestPptSync.bas NOVO (A1, A2); TestPptFake.cls ALT (A1); tests/vba/ppt/* NOVOS (A4); tests/vba/loader/TestRibbon.bas, TestUpdater.bas ALT (A3, A7)
tests/ps/Test-PptAddin.ps1 NOVO (A4); Test-Ribbon.ps1 (A4), Test-Installer.ps1 (A5), Test-ReleaseGuard.ps1, Test-PublishRelease.ps1 (A6), Test-Docs.ps1 (A3, A8)  ALT
docs/tutorial/tutorial.html, README.md, CHANGELOG.md, release/notes.txt  ALT (A3 mínimo; A8 completo)
--- Parte B (1.5.0) ---
src/core/modules/modDelta.bas          NOVO  Delta_* (B1)
src/core/modules/modDeltaPick.bas      NOVO  DeltaPick_* (B3)
src/core/modules/modChartKinds.bas     ALT   recurso delta (B0)
src/core/modules/modAnnotations.bas    ALT   recusa (B0); Ann_AddDelta, layout, faixas, sala, órfãs (B2); moldura (B3); alternar texto (B4)
src/core/modules/modChartWaterfall.bas ALT   WF_COL_LO, WF_COL_HI (B1)
src/core/modules/modCoreApi.bas        ALT   Core_AddDelta (B2); Core_DeltaPick (B3); Core_DeltaCycleFormat (B4)
src/core/classes/clsChartWatcher.cls, clsAppEvents.cls  ALT (B3)
src/core/modules/modSamples.bas        ALT   delta no exemplo de colunas agrupadas (B5)
src/loader/ribbon/customUI14.xml, src/loader/modules/modRibbon.bas  ALT (B4)
tests/vba/core/TestKinds.bas ALT (B0); TestDelta.bas NOVO (B1, B2); TestDeltaPick.bas NOVO (B3, B4); TestSamples.bas ALT (B5); tests/vba/loader/TestRibbon.bas ALT (B4)
tools/DemoCoverage.ps1, tools/New-DemoWorkbook.ps1, tests/ps/Test-DemoWorkbook.ps1, tests/ps/Test-Docs.ps1  ALT (B0, B5)
exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx, galeria do tutorial  REGERADOS (B5)
```

## Como trabalhar (todas as tarefas)

```powershell
# 1) editar no repositório real; 2) espelhar para a cópia; 3) rodar lá
robocopy "<repo>" C:\temp\tqq-pd /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS
pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target core -Filter "TestPptSync" -TimeoutSec 400     # iteração (RED/GREEN)
pwsh C:\temp\tqq-pd\build\run-tests.ps1 -TimeoutSec 900                                         # suíte completa antes de commitar
rm -rf /c/temp/tqq-pd                                                                           # (bash) ao terminar a tarefa
```
`<repo>` é o caminho do repositório real (o diretório de trabalho do implementador).

---

# PARTE A — PowerPoint atualizado ao reabrir (1.4.0)

### Tarefa A0: Promover spec e plano — modelo: cheap

**Files:**
- Create: `docs/superpowers/specs/2026-10-14-thinqq-ppt-delta-design.md` (cópia byte a byte de `.superpowers/sdd/2026-10-14-ppt-delta/design.md`; se já existir, não sobrescreva)
- Create: `docs/superpowers/plans/2026-10-14-thinqq-ppt-delta.md` (cópia byte a byte de `.superpowers/sdd/2026-10-14-ppt-delta/plan.md`; se já existir, não sobrescreva)

**Interfaces:**
- Produces: a cópia do spec com `## 6. Matriz de capacidades 1.5.0` e `## 7. Faixas de opções` (lida pelo teste da matriz na B0).

- [ ] **Passo 1:** `Copy-Item "<repo>\.superpowers\sdd\2026-10-14-ppt-delta\design.md" "<repo>\docs\superpowers\specs\2026-10-14-thinqq-ppt-delta-design.md"` e `Copy-Item "<repo>\.superpowers\sdd\2026-10-14-ppt-delta\plan.md" "<repo>\docs\superpowers\plans\2026-10-14-thinqq-ppt-delta.md"` (só se o destino não existir). Confira `Select-String -Path <spec> -Pattern '^## 6\. Matriz de capacidades 1\.5\.0','^## 7\. Faixas'` (2 linhas).
- [ ] **Passo 2:** `pwsh <repo>\tests\ps\Test-Privacy.ps1 -RepoRoot <repo>` → saída 0 (os documentos só citam `<repo>`, `%LOCALAPPDATA%`, `%USERPROFILE%` e `C:\temp\...`).
- [ ] **Passo 3:** `git add docs/superpowers/specs/2026-10-14-thinqq-ppt-delta-design.md docs/superpowers/plans/2026-10-14-thinqq-ppt-delta.md` e `git commit -m "docs: spec e plano do PowerPoint atualizado ao reabrir e do delta entre barras" (+ trailer)`.

---

### Tarefa A1: Motor de sincronização de UM gráfico + colagem que não estica as setas — modelo: most-capable

**Files:**
- Create: `src/core/modules/modPptSync.bas`
- Modify: `src/core/modules/modCoreApi.bas` (`CopyToPowerPointImpl` + `PptSyncAfterPaste`)
- Modify: `tests/vba/core/TestPptFake.cls` (campos `Width`, `Height`)
- Test: `tests/vba/core/TestPptSync.bas` (novo)

**Interfaces:**
- Produces (`modPptSync`): constantes `PPTSYNC_TAG_CHART = "TQQ_CHART"`, `PPTSYNC_TAG_BOOK = "TQQ_BOOK"`, `PPTSYNC_TAG_SHEET = "TQQ_SHEET"`, `PPTSYNC_TAG_VER = "TQQ_VER"`, `PPTSYNC_TAG_FP = "TQQ_FP"`; `PptSync_TagShape(shp As Object, co As ChartObject)`; `PptSync_Tag(shp As Object, key As String) As String`; `PptSync_SourcePath(linkFull As String, tagBook As String) As String`; `PptSync_Snapshot(cht As Chart) As Collection`; `PptSync_Apply(pc As Object, snaps As Collection, src As Chart, dx As Double, dy As Double) As Long`; `PptSync_ForceLayout(cht As Chart)`; `PptSync_SyncShape(shp As Object, co As ChartObject)` (levanta erro); `PptSync_Fingerprint(co As ChartObject, shp As Object) As String`; ganchos `PptSync_SetFailForTest(point)` ("update", "snapshot", "apply") e `PptSync_ResetTestHooks`.
- Produces (`modCoreApi`): a colagem vinculada grava as etiquetas, sincroniza e grava `TQQ_FP`.
- Consumes: `Ppt_LinkState` (modPowerPoint), `Core_SafeRefresh`, `Registry_*`, `Block`/`Sha256_Bytes`, `ANN_PREFIX`.

- [ ] **Passo 1: Dublê** — em `tests/vba/core/TestPptFake.cls`, depois de `Public QuitCalled As Boolean`, acrescente:

```vb
Public Width As Double                          ' forma do slide (PptSync_SyncShape le Width/Height antes de tocar no grafico)
Public Height As Double
```

- [ ] **Passo 2: Testes (RED)** — crie `tests/vba/core/TestPptSync.bas`:

```vb
Attribute VB_Name = "TestPptSync"
Option Explicit
' Parte A (1.4.0): motor de sincronizacao dos graficos ThinQQ colados no PowerPoint (modPptSync). Os testes com PowerPoint real so rodam sem
' nenhum POWERPNT aberto (aplicativo de instancia unica: CreateObject anexaria ao do usuario); senao saem verdes registrando o motivo.

Private Sub Teardown(wb As Workbook)
    PptSync_ResetTestHooks
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5): ws.Range("A3:C3").Value = Array("T2", 20, 6)
    Set Fixture = ws.Range("A1:C3")
End Function

Private Function PowerPointProcessCount() As Long
    PowerPointProcessCount = GetObject("winmgmts:" & Chr(92) & Chr(92) & "." & Chr(92) & "root" & Chr(92) & "cimv2").ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'POWERPNT.EXE'").Count
End Function

Private Function ClipboardUsable(cht As Chart) As Boolean
    On Error Resume Next
    cht.CopyPicture xlScreen, xlPicture
    ClipboardUsable = (Err.Number = 0)
    Err.Clear
End Function

Private Sub WaitSeconds(s As Single)
    Dim t As Single
    t = Timer
    Do While Timer >= t And Timer - t < s
        DoEvents
    Loop
End Sub

' So cria o PowerPoint se nenhum estiver aberto (ROT nem processo). Nothing (e o motivo no log) se pulou.
Private Function StartScratchPowerPoint() As Object
    Dim app As Object
    Ppt_ResetTestHooks
    If Ppt_IsRunning() Then Errors_Log "TestPptSync: pulado (ha um PowerPoint aberto)": Exit Function
    If PowerPointProcessCount() > 0 Then Errors_Log "TestPptSync: pulado (processo POWERPNT existente)": Exit Function
    On Error Resume Next
    Set app = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If app Is Nothing Then Errors_Log "TestPptSync: pulado (PowerPoint nao instalado)": Exit Function
    Set StartScratchPowerPoint = app
End Function

Private Function UserWindowAppeared(app As Object) As Boolean
    Dim n As Long, npv As Long
    n = -1: npv = -1
    On Error Resume Next
    n = app.Windows.Count
    npv = app.ProtectedViewWindows.Count
    On Error GoTo 0
    UserWindowAppeared = (n <> 0 Or npv <> 0)
    If UserWindowAppeared Then Errors_Log "TestPptSync: resto do teste pulado (uma janela apareceu na instancia do teste: PowerPoint do usuario)"
End Function

' Fecha so as apresentacoes do teste (sem janela) e encerra o PowerPoint se nada mais ficou aberto; espera o processo sair (ate 10 s).
Private Sub StopScratchPowerPoint(app As Object, own As Collection)
    Dim p As Variant, nw As Long, left_ As Long, t As Single, npv As Long
    On Error Resume Next
    Ppt_ResetTestHooks
    If app Is Nothing Then Exit Sub
    For Each p In own
        nw = -1
        nw = p.Windows.Count
        If nw = 0 Then p.Saved = -1: p.Close
    Next
    Set p = Nothing
    Set own = New Collection
    left_ = -1: npv = -1
    left_ = app.Presentations.Count
    npv = app.ProtectedViewWindows.Count
    If left_ = 0 And npv = 0 Then
        app.Quit
        Set app = Nothing
        t = Timer
        Do While PowerPointProcessCount() > 0
            If Timer < t Or Timer - t > 10 Then Exit Do
            WaitSeconds 0.25
        Loop
    Else
        Errors_Log "TestPptSync: o PowerPoint do teste ficou aberto (" & left_ & " apresentacao(oes), " & npv & " janela(s) protegida(s))"
        Set app = Nothing
    End If
End Sub

' Regra de tres propria no grafico DO SLIDE (eixo de valores = Axes(2) do PowerPoint).
Private Function PptYOf(pc As Object, v As Double) As Double
    Dim ax As Object
    Set ax = pc.Axes(2)
    PptYOf = pc.PlotArea.InsideTop + pc.PlotArea.InsideHeight * (ax.MaximumScale - v) / (ax.MaximumScale - ax.MinimumScale)
End Function

Private Function PptHasShape(pc As Object, nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = pc.Shapes(nm)
    PptHasShape = Not s Is Nothing
End Function

Private Function PptSeriesValue(pc As Object, seriesName As String, i As Long) As Double
    Dim k As Long, v As Variant
    For k = 1 To pc.SeriesCollection.Count
        If CStr(pc.SeriesCollection(k).Name) = seriesName Then
            v = pc.SeriesCollection(k).Values
            PptSeriesValue = CDbl(v(LBound(v) + i - 1))
            Exit Function
        End If
    Next
    Err.Raise vbObjectError + 1, "TestPptSync", "serie nao encontrada no grafico do slide: " & seriesName
End Function

Public Sub Test_PptSync_SourcePath()
    AssertEqual "C:\a\b.xlsx", PptSync_SourcePath("C:\a\b.xlsx", "")
    AssertEqual "C:\a\b.xlsx", PptSync_SourcePath("C:\a\b.xlsx!Dados![b.xlsx]Dados Gráfico 1", "C:\x.xlsx"), "OLE: o caminho vai ate o primeiro '!'"
    AssertEqual "C:\x.xlsx", PptSync_SourcePath("", "C:\x.xlsx"), "sem vinculo legivel: a etiqueta TQQ_BOOK"
    AssertEqual "", PptSync_SourcePath("   ", "")
End Sub

Public Sub Test_PptSync_Fingerprint_ChangesWithDataSizeAndShapes()
    Dim wb As Workbook, rng As Range, c As Chart, shp As TestPptFake, a As String, b As String
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set shp = New TestPptFake: shp.Width = 709.7: shp.Height = 432
    a = PptSync_Fingerprint(c.Parent, shp)
    AssertEqual 64, Len(a), "SHA-256 em hexadecimal"
    AssertEqual a, PptSync_Fingerprint(c.Parent, shp), "sem mudanca: igual"
    rng.Worksheet.Range("B3").Value = 21
    Application.Calculate
    b = PptSync_Fingerprint(c.Parent, shp)
    AssertTrue a <> b, "dado novo muda a impressao digital"
    shp.Width = 500
    AssertTrue b <> PptSync_Fingerprint(c.Parent, shp), "tamanho do slide muda a impressao digital"
    b = PptSync_Fingerprint(c.Parent, shp)
    Ann_AddCagr c, 1, 2
    AssertTrue b <> PptSync_Fingerprint(c.Parent, shp), "seta nova muda a impressao digital"
    Teardown wb
End Sub

' Falha no meio (gancho): o grafico do Excel volta ao tamanho original e a pasta continua "salva"; o erro sobe.
Public Sub Test_PptSync_FailureRestoresChartSize()
    Dim wb As Workbook, rng As Range, c As Chart, shp As TestPptFake, w0 As Double, h0 As Double, pt As Variant, en As Long
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    w0 = c.Parent.Width: h0 = c.Parent.Height
    Set shp = New TestPptFake: shp.Width = 709.7: shp.Height = 432      ' sem Chart: Ppt_LinkState = -1 (sem Update)
    For Each pt In Array("snapshot", "apply")
        wb.Saved = True
        PptSync_SetFailForTest CStr(pt)
        On Error Resume Next
        PptSync_SyncShape shp, c.Parent
        en = Err.Number
        On Error GoTo 0
        AssertTrue en <> 0, "a falha injetada sobe (" & pt & ")"
        AssertAlmost w0, c.Parent.Width, 0.01, "largura restaurada (" & pt & ")"
        AssertAlmost h0, c.Parent.Height, 0.01, "altura restaurada (" & pt & ")"
        AssertTrue wb.Saved, "pasta continua salva (" & pt & ")"
    Next
    Teardown wb
End Sub

' Review Focus 1: a colagem ajusta o grafico ao slide (709,7 x 432) e a seta CAGR fica sobre as colunas do SLIDE (regra de tres com a area
' interna e a escala do grafico do PowerPoint), nao esticada; etiquetas gravadas; o grafico do Excel volta ao tamanho dele.
Public Sub Test_PptSync_PastedChart_AnnotationsAlignedToSlide_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, shp As Object, pc As Object, ln As Object
    Dim own As New Collection, en As Long, ed As String, w0 As Double, h0 As Double, extra As Double, y1 As Double, y2 As Double, x1 As Double, x2 As Double
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Ppt_PinTargetForTest slide
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\pptsync-a1.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ann_AddCagr c, 1, 2
    w0 = c.Parent.Width: h0 = c.Parent.Height
    If Not ClipboardUsable(c) Then Errors_Log "TestPptSync: colagem pulada (area de transferencia indisponivel)": GoTo Cleanup
    Ppt_SetAppForTest app
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "chart"
    AssertEqual 1, slide.Shapes.Count, "grafico colado"
    Set shp = slide.Shapes(1)
    AssertEqual Registry_IdFromChart(c), PptSync_Tag(shp, PPTSYNC_TAG_CHART)
    AssertEqual wb.FullName, PptSync_Tag(shp, PPTSYNC_TAG_BOOK)
    AssertEqual "1", PptSync_Tag(shp, PPTSYNC_TAG_VER)
    AssertEqual 64, Len(PptSync_Tag(shp, PPTSYNC_TAG_FP)), "impressao digital gravada"
    AssertAlmost 432, CDbl(shp.Height), 2, "ajustado pela altura do slide"
    Set pc = shp.Chart
    If Registry_GetOption(c, "totals", "0") = "1" Then extra = 18          ' ANN_TOTALS_EXTRA (colunas, totais ligados)
    y1 = PptYOf(pc, 15) - 10 - extra: y2 = PptYOf(pc, 26) - 10 - extra      ' topo das pilhas: 10+5 e 20+6; ANN_ARROW_GAP = 10
    If y1 < 22 Then y1 = 22
    If y2 < 22 Then y2 = 22
    x1 = pc.PlotArea.InsideLeft + pc.PlotArea.InsideWidth * 0.5 / 2: x2 = pc.PlotArea.InsideLeft + pc.PlotArea.InsideWidth * 1.5 / 2
    Set ln = pc.Shapes("tqq_ann_cagr_1")
    AssertAlmost Application.Min(y1, y2), CDbl(ln.Top), 1.5, "topo da seta no slide"
    AssertAlmost Application.Max(y1, y2), CDbl(ln.Top + ln.Height), 1.5, "base da seta no slide"
    AssertAlmost x1, CDbl(ln.Left), 1.5, "inicio da seta no centro da 1a coluna do slide"
    AssertAlmost x2, CDbl(ln.Left + ln.Width), 1.5, "fim da seta no centro da 2a coluna do slide"
    AssertEqual c.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text, CStr(pc.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text)
    AssertAlmost w0, c.Parent.Width, 0.01, "grafico do Excel volta ao tamanho original"
    AssertAlmost h0, c.Parent.Height, 0.01
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set ln = Nothing: Set pc = Nothing: Set shp = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 2: dado novo (pasta ABERTA e nao salva), seta CAGR apagada e seta de diferenca nova no Excel; a sincronizacao atualiza os
' valores no slide, apaga a seta antiga, clona a nova (com o formato) e mantem a animacao da forma do slide.
Public Sub Test_PptSync_DataChangeNewArrowAndAnimation_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, shp As Object, pc As Object
    Dim own As New Collection, en As Long, ed As String, w0 As Double, seq As Object
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Ppt_PinTargetForTest slide
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\pptsync-a2.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ann_AddCagr c, 1, 2
    w0 = c.Parent.Width
    If Not ClipboardUsable(c) Then Errors_Log "TestPptSync: colagem pulada (area de transferencia indisponivel)": GoTo Cleanup
    Ppt_SetAppForTest app
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "chart"
    Set shp = slide.Shapes(1)
    Set seq = slide.TimeLine.MainSequence
    seq.AddEffect shp, 10                                       ' msoAnimEffectFade
    rng.Worksheet.Range("B3").Value = 40                        ' sem salvar
    Application.Calculate
    Ann_RemoveAll c
    Ann_AddDifference c, 1, 2
    PptSync_SyncShape shp, c.Parent
    Set pc = shp.Chart
    AssertAlmost 40, PptSeriesValue(pc, "Sul", 2), 0.0001, "valor nao salvo lido pelo vinculo (pasta aberta)"
    AssertFalse PptHasShape(pc, "tqq_ann_cagr_1"), "seta apagada no Excel some do slide"
    AssertTrue PptHasShape(pc, "tqq_ann_diff_1") And PptHasShape(pc, "tqq_ann_diff_1_lvl") And PptHasShape(pc, "tqq_ann_diff_1_lbl"), "seta nova clonada"
    AssertEqual c.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text, CStr(pc.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text)
    AssertAlmost c.Shapes("tqq_ann_diff_1").Line.Weight, CDbl(pc.Shapes("tqq_ann_diff_1").Line.Weight), 0.01, "formato da linha clonado"
    AssertEqual CLng(c.Shapes("tqq_ann_diff_1").Line.EndArrowheadStyle), CLng(pc.Shapes("tqq_ann_diff_1").Line.EndArrowheadStyle)
    AssertEqual 1, CLng(seq.Count), "animacao mantida"
    AssertEqual CStr(shp.Name), CStr(seq.Item(1).Shape.Name)
    AssertAlmost w0, c.Parent.Width, 0.01
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set seq = Nothing: Set pc = Nothing: Set shp = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub
```

- [ ] **Passo 3: RED** — `robocopy ...` e `pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target core -Filter "TestPptSync" -TimeoutSec 400` → saída 3 (compilação: `PptSync_*` não existe). Anote como RED.

- [ ] **Passo 4: Implementar `modPptSync.bas`** — crie `src/core/modules/modPptSync.bas`:

```vb
Attribute VB_Name = "modPptSync"
Option Explicit
' Graficos ThinQQ colados no PowerPoint como grafico vinculado (Parte A, spec 2026-10-14 secao 4.2): etiquetas na forma do slide (TQQ_*) e
' sincronizacao. Dados pelo vinculo nativo (LinkFormat.Update: le a pasta ABERTA, mesmo com alteracoes nao salvas; fechada, abre um Excel
' -Embedding, ~4 s); as formas das anotacoes (tqq_ann_*) sao recalculadas pelo ThinQQ com o grafico do Excel no TAMANHO DO SLIDE e a foto
' delas e aplicada dentro do grafico do slide, NO LUGAR (animacao, ordem Z, nome e texto alternativo da forma do slide ficam). Medido (spec
' 3.1): com o mesmo tamanho a area interna dos dois aplicativos e identica; redimensionar o grafico do slide ESTICA as formas internas;
' LinkFormat.Update com a pasta de origem inexistente TRAVA o PowerPoint (o arquivo e conferido antes, em PptSync_Presentation).
' Ligacao tardia com o PowerPoint (As Object, constantes numericas); sem Declare.
Public Const PPTSYNC_TAG_CHART As String = "TQQ_CHART"      ' id do grafico (ChartObject "tqq_" & id)
Public Const PPTSYNC_TAG_BOOK As String = "TQQ_BOOK"        ' caminho completo da pasta na colagem
Public Const PPTSYNC_TAG_SHEET As String = "TQQ_SHEET"      ' planilha do grafico na colagem (informativo)
Public Const PPTSYNC_TAG_VER As String = "TQQ_VER"          ' versao do formato das etiquetas
Public Const PPTSYNC_TAG_FP As String = "TQQ_FP"            ' impressao digital da ultima sincronizacao (PptSync_Fingerprint)
Private Const MSO_AUTOSHAPE As Long = 1
Private Const MSO_LINE As Long = 9
Private Const MSO_TEXTBOX As Long = 17
Private Const SN_NAME As Long = 0                           ' campos da foto de uma forma (PptSync_Snapshot)
Private Const SN_LEFT As Long = 1
Private Const SN_TOP As Long = 2
Private Const SN_WIDTH As Long = 3
Private Const SN_HEIGHT As Long = 4
Private Const SN_VISIBLE As Long = 5
Private Const SN_HFLIP As Long = 6
Private Const SN_VFLIP As Long = 7
Private Const SN_TEXT As Long = 8
Private Const SN_HASTEXT As Long = 9
Private Const PPTSYNC_LAYOUT_TOL As Double = 2              ' pt: area interna do slide diferente da do Excel alem disso = log
Private mTestFail As String                                 ' gancho de teste: ponto de falha injetada (so no build Test)

' Gancho de teste: o proximo ponto nomeado ("update", "snapshot", "apply") levanta erro (uso unico). Inerte fora do build Test.
Public Sub PptSync_SetFailForTest(point As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestFail = point
End Sub

Public Sub PptSync_ResetTestHooks()
    mTestFail = ""
End Sub

Private Sub PptSync_Hook(point As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    If mTestFail <> point Then Exit Sub
    mTestFail = ""
    Err.Raise vbObjectError + 9992, "modPptSync", "falha injetada (" & point & ")"
End Sub

' Etiquetas da forma colada (grafico vinculado): o que o suplemento do PowerPoint e PptSync_Presentation leem para achar o grafico de origem.
Public Sub PptSync_TagShape(shp As Object, co As ChartObject)
    shp.Tags.Add PPTSYNC_TAG_CHART, Registry_IdFromChart(co.Chart)
    shp.Tags.Add PPTSYNC_TAG_BOOK, Registry_Workbook(co.Chart).FullName
    shp.Tags.Add PPTSYNC_TAG_SHEET, co.Parent.Name
    shp.Tags.Add PPTSYNC_TAG_VER, "1"
End Sub

' Valor de uma etiqueta ("" se nao existe ou a leitura falha). Nunca propaga erro.
Public Function PptSync_Tag(shp As Object, key As String) As String
    On Error Resume Next
    PptSync_Tag = CStr(shp.Tags.Item(key))
    Err.Clear
End Function

' Caminho da pasta de origem: LinkFormat.SourceFullName ate o primeiro "!" (o grafico vinculado devolve o caminho puro; objeto OLE devolve
' "caminho!planilha!..."); sem vinculo legivel, a etiqueta TQQ_BOOK. Pura.
Public Function PptSync_SourcePath(linkFull As String, tagBook As String) As String
    Dim s As String, p As Long
    s = Trim$(linkFull)
    p = InStr(s, "!")
    If p > 0 Then s = Left$(s, p - 1)
    If Len(s) = 0 Then s = Trim$(tagBook)
    PptSync_SourcePath = s
End Function

' Foto das formas das anotacoes (prefixo tqq_ann_) do grafico do Excel: nome, caixa, visivel, giros e texto. Chave da colecao = nome.
Public Function PptSync_Snapshot(cht As Chart) As Collection
    Dim out As New Collection, i As Long, s As Shape, txt As String, hasTxt As Boolean
    For i = 1 To cht.Shapes.Count
        Set s = cht.Shapes(i)
        If Left$(s.Name, Len(ANN_PREFIX)) = ANN_PREFIX Then
            txt = "": hasTxt = False
            If s.Type <> MSO_LINE Then
                On Error Resume Next
                txt = s.TextFrame2.TextRange.Text
                hasTxt = (Err.Number = 0)
                Err.Clear
                On Error GoTo 0
            End If
            out.Add Array(s.Name, CDbl(s.Left), CDbl(s.Top), CDbl(s.Width), CDbl(s.Height), CLng(s.Visible), CLng(s.HorizontalFlip), CLng(s.VerticalFlip), txt, hasTxt), s.Name
        End If
    Next
    Set PptSync_Snapshot = out
End Function

' Aplica a foto no grafico do slide (pc): apaga as tqq_ann_* do slide que nao existem mais no Excel, clona as que faltam (src = grafico do
' Excel, para o formato) e poe caixa, giros, texto e visibilidade em todas, deslocadas por (dx, dy). Devolve quantas formas clonou.
Public Function PptSync_Apply(pc As Object, snaps As Collection, src As Chart, dx As Double, dy As Double) As Long
    Dim i As Long, nm As String, t As Object, sn As Variant, keep As String
    keep = "|"
    For Each sn In snaps
        keep = keep & sn(SN_NAME) & "|"
    Next
    For i = pc.Shapes.Count To 1 Step -1
        nm = CStr(pc.Shapes(i).Name)
        If Left$(nm, Len(ANN_PREFIX)) = ANN_PREFIX And InStr(keep, "|" & nm & "|") = 0 Then pc.Shapes(i).Delete
    Next
    For Each sn In snaps
        Set t = Nothing
        On Error Resume Next
        Set t = pc.Shapes(CStr(sn(SN_NAME)))
        Err.Clear
        On Error GoTo 0
        If t Is Nothing Then
            Set t = PptSync_Clone(pc, src.Shapes(CStr(sn(SN_NAME))))
            If Not t Is Nothing Then PptSync_Apply = PptSync_Apply + 1
        End If
        If Not t Is Nothing Then PptSync_Place t, sn, dx, dy
    Next
End Function

Private Sub PptSync_Place(t As Object, sn As Variant, dx As Double, dy As Double)
    If CBool(sn(SN_HASTEXT)) Then
        If CStr(t.TextFrame2.TextRange.Text) <> CStr(sn(SN_TEXT)) Then t.TextFrame2.TextRange.Text = CStr(sn(SN_TEXT))
    End If
    If (CLng(t.HorizontalFlip) = -1) <> (CLng(sn(SN_HFLIP)) = -1) Then t.Flip 0          ' msoFlipHorizontal
    If (CLng(t.VerticalFlip) = -1) <> (CLng(sn(SN_VFLIP)) = -1) Then t.Flip 1            ' msoFlipVertical
    t.Left = CSng(sn(SN_LEFT) + dx)
    t.Top = CSng(sn(SN_TOP) + dy)
    t.Width = CSng(sn(SN_WIDTH))
    t.Height = CSng(sn(SN_HEIGHT))
    If CLng(t.Visible) <> CLng(sn(SN_VISIBLE)) Then t.Visible = CLng(sn(SN_VISIBLE))
End Sub

' Cria no grafico do slide a forma que so existe no Excel (seta inserida depois da colagem): linha, caixa de texto ou autoforma, com o
' formato copiado (o melhor esforco: um formato que nao copia so e registrado). Outros tipos: Nothing (registrado).
Private Function PptSync_Clone(pc As Object, s As Shape) As Object
    Dim t As Object
    Select Case s.Type
        Case MSO_LINE
            Set t = pc.Shapes.AddLine(s.Left, s.Top, s.Left + s.Width, s.Top + s.Height)
            PptSync_CopyLine s, t
        Case MSO_TEXTBOX
            Set t = pc.Shapes.AddTextbox(1, s.Left, s.Top, s.Width, s.Height)        ' msoTextOrientationHorizontal
            PptSync_CopyText s, t
            PptSync_CopyFillAndBorder s, t
        Case MSO_AUTOSHAPE
            Set t = pc.Shapes.AddShape(s.AutoShapeType, s.Left, s.Top, s.Width, s.Height)
            PptSync_CopyFillAndBorder s, t
            PptSync_CopyText s, t
        Case Else
            Errors_Log "PptSync_Clone: tipo de forma nao copiado (" & s.Type & ")" & vbTab & s.Name
            Exit Function
    End Select
    t.Name = s.Name
    t.AlternativeText = s.AlternativeText
    Set PptSync_Clone = t
End Function

Private Sub PptSync_CopyLine(s As Shape, t As Object)
    On Error Resume Next
    With t.Line
        .Visible = s.Line.Visible
        .ForeColor.RGB = s.Line.ForeColor.RGB
        .Weight = s.Line.Weight
        .DashStyle = s.Line.DashStyle
        .BeginArrowheadStyle = s.Line.BeginArrowheadStyle
        .EndArrowheadStyle = s.Line.EndArrowheadStyle
        .BeginArrowheadLength = s.Line.BeginArrowheadLength
        .EndArrowheadLength = s.Line.EndArrowheadLength
        .BeginArrowheadWidth = s.Line.BeginArrowheadWidth
        .EndArrowheadWidth = s.Line.EndArrowheadWidth
    End With
    If Err.Number <> 0 Then Errors_Log "PptSync_CopyLine " & s.Name & vbTab & Err.Number & vbTab & Err.Description
End Sub

Private Sub PptSync_CopyText(s As Shape, t As Object)
    On Error Resume Next
    With t.TextFrame2
        .WordWrap = s.TextFrame2.WordWrap
        .AutoSize = s.TextFrame2.AutoSize
        .MarginLeft = s.TextFrame2.MarginLeft: .MarginRight = s.TextFrame2.MarginRight
        .MarginTop = s.TextFrame2.MarginTop: .MarginBottom = s.TextFrame2.MarginBottom
        .TextRange.Text = s.TextFrame2.TextRange.Text
        .TextRange.Font.Name = s.TextFrame2.TextRange.Font.Name
        .TextRange.Font.Size = s.TextFrame2.TextRange.Font.Size
        .TextRange.Font.Bold = s.TextFrame2.TextRange.Font.Bold
        .TextRange.Font.Fill.ForeColor.RGB = s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
        .TextRange.ParagraphFormat.Alignment = s.TextFrame2.TextRange.ParagraphFormat.Alignment
    End With
    If Err.Number <> 0 Then Errors_Log "PptSync_CopyText " & s.Name & vbTab & Err.Number & vbTab & Err.Description
End Sub

Private Sub PptSync_CopyFillAndBorder(s As Shape, t As Object)
    On Error Resume Next
    t.Fill.Visible = s.Fill.Visible
    If s.Fill.Visible = msoTrue Then t.Fill.ForeColor.RGB = s.Fill.ForeColor.RGB
    t.Line.Visible = s.Line.Visible
    If s.Line.Visible = msoTrue Then
        t.Line.ForeColor.RGB = s.Line.ForeColor.RGB
        t.Line.Weight = s.Line.Weight
        t.Line.DashStyle = s.Line.DashStyle
    End If
    If Err.Number <> 0 Then Errors_Log "PptSync_CopyFillAndBorder " & s.Name & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Forca o Excel a refazer a geometria do grafico (pontos, rotulos, area interna): exportar e o unico jeito no Excel invisivel (o ForceLayout
' dos testes faz o mesmo). Arquivo temporario apagado; nunca propaga erro.
Public Sub PptSync_ForceLayout(cht As Chart)
    Dim p As String
    On Error Resume Next
    p = Environ$("TEMP") & "\thinqq_pptsync_layout.png"
    cht.Export p
    If Len(Dir$(p)) > 0 Then Kill p
    Err.Clear
End Sub

' Sincroniza UM grafico do slide (shp: forma do PowerPoint com o grafico vinculado) com o grafico ThinQQ co: (1) dados pelo vinculo nativo
' (so com o vinculo ativo); (2) o ChartObject vai para o tamanho da forma do slide, o ThinQQ recalcula as anotacoes (Core_SafeRefresh) e a foto
' delas e aplicada no grafico do slide com o deslocamento das areas internas (medido 0); (3) o ChartObject volta ao tamanho original e a pasta
' volta a "salva" se estava. Levanta erro (quem chama registra e conta); o tamanho e restaurado em qualquer caminho.
Public Sub PptSync_SyncShape(shp As Object, co As ChartObject)
    Dim cht As Chart, wb As Workbook, w0 As Double, h0 As Double, saved0 As Boolean, resized As Boolean, prevSU As Boolean
    Dim snaps As Collection, pc As Object, dx As Double, dy As Double, en As Long, ed As String
    Set cht = co.Chart: Set wb = Registry_Workbook(cht)
    saved0 = wb.Saved: w0 = co.Width: h0 = co.Height
    prevSU = Application.ScreenUpdating
    On Error GoTo Fail
    If Ppt_LinkState(shp) = 1 Then shp.LinkFormat.Update
    PptSync_Hook "update"
    Application.ScreenUpdating = False
    resized = True
    co.Width = CDbl(shp.Width): co.Height = CDbl(shp.Height)
    PptSync_ForceLayout cht
    Core_SafeRefresh cht
    PptSync_ForceLayout cht
    PptSync_Hook "snapshot"
    Set snaps = PptSync_Snapshot(cht)
    PptSync_Hook "apply"
    Set pc = shp.Chart
    dx = CDbl(pc.PlotArea.InsideLeft) - cht.PlotArea.InsideLeft
    dy = CDbl(pc.PlotArea.InsideTop) - cht.PlotArea.InsideTop
    If Abs(CDbl(pc.PlotArea.InsideWidth) - cht.PlotArea.InsideWidth) > PPTSYNC_LAYOUT_TOL Or Abs(CDbl(pc.PlotArea.InsideHeight) - cht.PlotArea.InsideHeight) > PPTSYNC_LAYOUT_TOL Then
        Errors_Log "PptSync: area interna do slide diferente da do Excel (grafico editado no PowerPoint?)" & vbTab & CStr(shp.Name)
    End If
    PptSync_Apply pc, snaps, cht, dx, dy
    PptSync_Restore co, w0, h0, wb, saved0
    Application.ScreenUpdating = prevSU
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    If resized Then PptSync_Restore co, w0, h0, wb, saved0
    Application.ScreenUpdating = prevSU
    On Error GoTo 0
    Err.Raise en, "PptSync_SyncShape", ed
End Sub

Private Sub PptSync_Restore(co As ChartObject, w0 As Double, h0 As Double, wb As Workbook, saved0 As Boolean)
    If Abs(co.Width - w0) > 0.01 Or Abs(co.Height - h0) > 0.01 Then
        co.Width = w0: co.Height = h0
        Core_SafeRefresh co.Chart
    End If
    If saved0 Then wb.Saved = True
End Sub

' Impressao digital do que o slide mostraria: SHA-256 do bloco inteiro do grafico (valores e a linha de metadados com as opcoes), dos nomes e
' textos alternativos das anotacoes e do tamanho da forma do slide. Igual = nada a atualizar (PptSync_Presentation sem "forcar").
Public Function PptSync_Fingerprint(co As ChartObject, shp As Object) As String
    Dim blk As Range, v As Variant, s As String, r As Long, c As Long, i As Long, b() As Byte, nm As String
    Set blk = Registry_Block(co.Chart)
    If Not blk Is Nothing Then
        v = blk.Value
        If IsArray(v) Then
            For r = LBound(v, 1) To UBound(v, 1)
                For c = LBound(v, 2) To UBound(v, 2)
                    If IsError(v(r, c)) Then s = s & "#" Else s = s & CStr(v(r, c))
                    s = s & Chr$(31)
                Next
            Next
        ElseIf Not IsError(v) Then
            s = CStr(v)
        End If
    End If
    For i = 1 To co.Chart.Shapes.Count
        nm = co.Chart.Shapes(i).Name
        If Left$(nm, Len(ANN_PREFIX)) = ANN_PREFIX Then s = s & nm & "=" & co.Chart.Shapes(i).AlternativeText & Chr$(30)
    Next
    s = s & "|" & Format$(CDbl(shp.Width), "0.0") & "x" & Format$(CDbl(shp.Height), "0.0")
    b = StrConv(s, vbFromUnicode)
    PptSync_Fingerprint = Sha256_Bytes(b, UBound(b) - LBound(b) + 1)
End Function
```

Confira antes de compilar: `Sha256_Bytes(data() As Byte, ByVal length As Long)` existe em `src/core/modules/modSha256.bas` e devolve 64 caracteres hexadecimais minúsculos (se devolver maiúsculos, o teste só confere o comprimento). `Registry_Block` devolve o bloco inteiro (metadados na linha 1).

- [ ] **Passo 5: Colagem marca e sincroniza** — em `src/core/modules/modCoreApi.bas`, em `CopyToPowerPointImpl`, troque

```vb
        linkState = Ppt_LinkState(shp)               ' o PowerPoint pode colar sem o vinculo: o aviso diz isso
```
por
```vb
        linkState = Ppt_LinkState(shp)               ' o PowerPoint pode colar sem o vinculo: o aviso diz isso
        If linkState = 1 Then PptSyncAfterPaste shp, co     ' etiquetas e anotacoes no tamanho do slide (o ajuste ao slide estica as formas)
```
e acrescente, logo depois do `End Sub` de `CopyToPowerPointImpl`:

```vb
' Grafico vinculado recem-colado (Parte A): etiquetas TQQ_* (para a atualizacao ao reabrir) e as anotacoes recalculadas no tamanho do slide
' (medido: Ppt_PlaceShape estica as formas internas junto com o grafico). Falha aqui NAO desfaz a colagem: fica registrada (o botao
' "Atualizar graficos no PowerPoint" refaz).
Private Sub PptSyncAfterPaste(shp As Object, co As ChartObject)
    On Error Resume Next
    PptSync_TagShape shp, co
    If Err.Number <> 0 Then Errors_Log "CopyToPowerPoint/etiquetas" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    PptSync_SyncShape shp, co
    If Err.Number <> 0 Then Errors_Log "CopyToPowerPoint/sincronizar" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    shp.Tags.Add PPTSYNC_TAG_FP, PptSync_Fingerprint(co, shp)
    Err.Clear
End Sub
```

- [ ] **Passo 6: GREEN** — `pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target core -Filter "TestPptSync|TestPowerPoint" -TimeoutSec 600` → 0 falhas. Sem `POWERPNT` aberto os dois testes `_WhenAvailable` devem ter rodado de verdade: confira no log do ThinQQ (`%LOCALAPPDATA%\..` não: o log do core de teste fica em `Errors_LogDir()`; procure "TestPptSync: pulado") e relate se pularam. Se `Test_PptSync_PastedChart_...` falhar por geometria (topo da seta com erro > 1,5 pt), investigue antes de mexer na tolerância: imprima `pc.PlotArea.Inside*` e `c.PlotArea.Inside*` com o ChartObject no tamanho do slide (devem bater, spec §3.1) e confira se o `PptSync_ForceLayout` rodou antes da foto.
- [ ] **Passo 7: Suíte completa** — `pwsh C:\temp\tqq-pd\build\run-tests.ps1 -TimeoutSec 900` → "Todos os testes passaram".
- [ ] **Passo 8: Commit** — `git add src/core/modules/modPptSync.bas src/core/modules/modCoreApi.bas tests/vba/core/TestPptSync.bas tests/vba/core/TestPptFake.cls` e `git commit -m "feat: motor de sincronizacao do PowerPoint (dados e setas no tamanho do slide); colagem vinculada deixa de esticar as setas" (+ trailer)`.

---

### Tarefa A2: Sincronizar a apresentação inteira + API para o suplemento e o botão do Excel — modelo: standard

**Files:**
- Modify: `src/core/modules/modPptSync.bas` (acrescentar ao fim)
- Modify: `src/core/modules/modCoreApi.bas` (`Core_PptSyncPresentation`, `Core_PptSyncOpen`)
- Test: `tests/vba/core/TestPptSync.bas` (acrescentar)

**Interfaces:**
- Consumes (A1): `PptSync_SyncShape`, `PptSync_Tag`, `PptSync_SourcePath`, `PptSync_Fingerprint`, `PPTSYNC_TAG_*`.
- Produces (`modPptSync`): `PPTSYNC_NONE`, `PPTSYNC_NO_PPT`; `PptSync_SummaryText(nOk, nSame, missing, nNoChart, nUnlinked, nFail) As String`; `PptSync_FileExists(path) As Boolean`; `PptSync_Presentation(pres As Object, force As Boolean) As String` (levanta erro só em falha estrutural; cada gráfico isolado).
- Produces (`modCoreApi`): `Public Function Core_PptSyncPresentation(pres As Object, Optional force As Boolean = False) As String` (nunca propaga), `Public Sub Core_PptSyncOpen()` e o contrato de versão com o suplemento: `Public Const CORE_PPT_API As Long = 1` e `Public Function Core_PptApiVersion() As Long` (spec §4.9: a assinatura de `Core_PptSyncPresentation` fica congelada; mudança incompatível = função nova e `CORE_PPT_API` 2, mantendo a antiga).

- [ ] **Passo 1: Testes (RED)** — acrescente ao fim de `tests/vba/core/TestPptSync.bas`:

```vb
Public Sub Test_PptSync_SummaryText()
    AssertEqual "Gráficos do ThinQQ atualizados: 2.", PptSync_SummaryText(2, 0, "", 0, 0, 0)
    AssertEqual "Gráficos do ThinQQ atualizados: 0. Sem mudança: 3.", PptSync_SummaryText(0, 3, "", 0, 0, 0)
    AssertEqual "Gráficos do ThinQQ atualizados: 1." & vbLf & "Pasta de origem não encontrada ou sem acesso (o gráfico ficou como estava): a.xlsx, b.xlsx. Se ela foi movida ou renomeada, use no PowerPoint Arquivo > Informações > Editar Links para Arquivos > Alterar Origem.", PptSync_SummaryText(1, 0, "a.xlsx, b.xlsx", 0, 0, 0)
    AssertEqual "Gráficos do ThinQQ atualizados: 0." & vbLf & "Gráfico não encontrado na pasta de origem (apagado ou recriado): 2." & vbLf & "Sem vínculo com o Excel (Quebrar Vínculo no PowerPoint): 1." & vbLf & "Falha ao atualizar: 1 (veja o log do ThinQQ).", PptSync_SummaryText(0, 0, "", 2, 1, 1)
End Sub

Public Sub Test_Core_PptApiVersion()
    AssertEqual 1, Core_PptApiVersion(), "contrato com o suplemento do PowerPoint (spec 4.9)"
End Sub

Public Sub Test_Core_PptSyncPresentation_NeverRaises()
    Dim s As String
    s = Core_PptSyncPresentation(Nothing, True)
    AssertTrue InStr(s, "Não foi possível atualizar os gráficos do ThinQQ") = 1, s
End Sub

Public Sub Test_Core_PptSyncOpen_WithoutPowerPoint_Notifies()
    Ppt_ResetTestHooks                                         ' build Test sem instancia injetada: Ppt_GetApp devolve Nothing
    Core_ClearLastNotice
    Core_PptSyncOpen
    AssertEqual PPTSYNC_NO_PPT, Core_LastNotice()
End Sub

' Cola um grafico de colunas empilhadas com seta CAGR da pasta salva em p (uma apresentacao do teste, sem janela). Devolve a forma colada.
Private Function PasteLinked(app As Object, pres As Object, wb As Workbook, rng As Range, p As String) As Object
    Dim slide As Object, c As Chart
    Set slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
    Ppt_PinTargetForTest slide
    wb.SaveAs p, xlOpenXMLWorkbook
    rng.Worksheet.Activate
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ann_AddCagr c, 1, 2
    Ppt_SetAppForTest app
    Core_CopyToPowerPoint "chart"
    Set PasteLinked = slide.Shapes(slide.Shapes.Count)
End Function

' Review Focus 3: pasta de origem inexistente (salva com outro nome e a original apagada): nenhum LinkFormat.Update (travaria), o grafico fica,
' o resumo diz o arquivo; tudo em menos de 10 s.
Public Sub Test_PptSync_MissingSource_DoesNotHang_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String
    Dim a As String, s As String, t As Single, v0 As Double
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    a = TempDir() & "\pptsync-orig.xlsx"
    Set shp = PasteLinked(app, pres, wb, rng, a)
    v0 = PptSeriesValue(shp.Chart, "Sul", 1)
    wb.SaveAs TempDir() & "\pptsync-outro.xlsx", xlOpenXMLWorkbook
    Kill a
    t = Timer
    s = PptSync_Presentation(pres, True)
    AssertTrue Timer - t < 10, "sem travar (" & Format$(Timer - t, "0.0") & " s)"
    AssertTrue InStr(s, "Pasta de origem não encontrada ou sem acesso") > 0 And InStr(s, "pptsync-orig.xlsx") > 0, s
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 0.") = 1, s
    AssertAlmost v0, PptSeriesValue(shp.Chart, "Sul", 1), 0.0001, "grafico ficou como estava"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 2: pasta fechada com dado novo salvo: aberta somente leitura, sincronizada e fechada (nenhuma pasta a mais aberta no fim).
Public Sub Test_PptSync_ClosedSource_OpenedReadOnlyAndClosed_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String
    Dim s As String, n0 As Long
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptsync-fechada.xlsx")
    rng.Worksheet.Range("B2").Value = 33
    wb.Save
    Watchers_Stop
    CloseTemp wb: Set wb = Nothing
    n0 = Application.Workbooks.Count
    s = PptSync_Presentation(pres, True)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertAlmost 33, PptSeriesValue(shp.Chart, "Sul", 1), 0.0001, "dado novo da pasta fechada"
    AssertEqual n0, Application.Workbooks.Count, "a pasta aberta pela sincronizacao foi fechada"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Impressao digital: sem mudanca o grafico e pulado ("Sem mudança"), com dado novo e atualizado; forcar atualiza sempre. Core_PptSyncOpen
' (instancia injetada) cita a apresentacao no aviso.
Public Sub Test_PptSync_Fingerprint_SkipsUnchanged_AndCoreOpen_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String, s As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptsync-fp.xlsx")
    s = PptSync_Presentation(pres, False)
    AssertEqual "Gráficos do ThinQQ atualizados: 0. Sem mudança: 1.", s
    rng.Worksheet.Range("C3").Value = 9
    Application.Calculate
    s = PptSync_Presentation(pres, False)
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", s
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", PptSync_Presentation(pres, True), "forcar: atualiza mesmo sem mudanca"
    Ppt_SetAppForTest app
    Core_ClearLastNotice
    Core_PptSyncOpen
    AssertTrue InStr(Core_LastNotice(), "'" & CStr(pres.Name) & "': Gráficos do ThinQQ atualizados: 1.") > 0, Core_LastNotice()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "TestPptSync"` → saída 3 (`PptSync_SummaryText` não existe).

- [ ] **Passo 3: Implementar** — acrescente ao fim de `src/core/modules/modPptSync.bas`:

```vb
Public Const CORE_PPT_API As Long = 1                     ' versao da API do core usada pelo suplemento do PowerPoint (Core_PptApiVersion)
Public Const PPTSYNC_NONE As String = "Esta apresentação não tem gráficos do ThinQQ vinculados (copiados com 'Copiar como gráfico vinculado')."
Public Const PPTSYNC_NO_PPT As String = "Abra a apresentação no PowerPoint e use de novo 'Atualizar gráficos no PowerPoint'."

' Resumo para o usuario (textos exatos da spec 4.2). missing = nomes dos arquivos ja separados por ", ". Pura.
Public Function PptSync_SummaryText(nOk As Long, nSame As Long, missing As String, nNoChart As Long, nUnlinked As Long, nFail As Long) As String
    Dim s As String
    s = "Gráficos do ThinQQ atualizados: " & nOk & "."
    If nSame > 0 Then s = s & " Sem mudança: " & nSame & "."
    If Len(missing) > 0 Then s = s & vbLf & "Pasta de origem não encontrada ou sem acesso (o gráfico ficou como estava): " & missing & ". Se ela foi movida ou renomeada, use no PowerPoint Arquivo > Informações > Editar Links para Arquivos > Alterar Origem."
    If nNoChart > 0 Then s = s & vbLf & "Gráfico não encontrado na pasta de origem (apagado ou recriado): " & nNoChart & "."
    If nUnlinked > 0 Then s = s & vbLf & "Sem vínculo com o Excel (Quebrar Vínculo no PowerPoint): " & nUnlinked & "."
    If nFail > 0 Then s = s & vbLf & "Falha ao atualizar: " & nFail & " (veja o log do ThinQQ)."
    PptSync_SummaryText = s
End Function

Public Function PptSync_FileExists(path As String) As Boolean
    On Error Resume Next
    If Len(path) = 0 Then Exit Function
    PptSync_FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
    Err.Clear
End Function

Private Function PptSync_LinkSource(shp As Object) As String
    On Error Resume Next
    PptSync_LinkSource = CStr(shp.LinkFormat.SourceFullName)
    Err.Clear
End Function

' Pasta de origem: a ja aberta neste Excel (com as alteracoes nao salvas, como o proprio vinculo le) ou aberta agora SOMENTE LEITURA, sem
' atualizar vinculos, sem MRU e com eventos desligados durante a abertura (registrada em opened para fechar no fim). Nothing se nao abre.
Private Function PptSync_OpenBook(path As String, opened As Collection) As Workbook
    Dim w As Workbook, prevAlerts As Boolean, prevEvents As Boolean
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then Set PptSync_OpenBook = w: Exit Function
    Next
    prevAlerts = Application.DisplayAlerts: prevEvents = Application.EnableEvents
    Application.DisplayAlerts = False: Application.EnableEvents = False
    On Error Resume Next
    Set w = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    If Err.Number <> 0 Then Errors_Log "PptSync_OpenBook" & vbTab & Err.Number & vbTab & Err.Description: Set w = Nothing
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = prevAlerts: Application.EnableEvents = prevEvents
    If Not w Is Nothing Then opened.Add w
    Set PptSync_OpenBook = w
End Function

Private Sub PptSync_CloseOpened(opened As Collection)
    Dim w As Variant
    On Error Resume Next
    For Each w In opened
        w.Close SaveChanges:=False
    Next
End Sub

Private Function PptSync_TrySync(shp As Object, co As ChartObject) As Boolean
    On Error GoTo Fail
    PptSync_SyncShape shp, co
    PptSync_TrySync = True
    Exit Function
Fail:
    Errors_Log "PptSync" & vbTab & Err.Number & vbTab & Err.Description
End Function

' Sincroniza todos os graficos ThinQQ vinculados (formas soltas com a etiqueta TQQ_CHART) da apresentacao. force = False pula os que nao
' mudaram (TQQ_FP). Pasta de origem inexistente ou que nao abre: o grafico fica como esta (NUNCA LinkFormat.Update: travaria o PowerPoint).
' Pastas abertas por esta chamada sao fechadas sem salvar. Devolve o resumo (PptSync_SummaryText). Erro de um grafico e contado e registrado.
Public Function PptSync_Presentation(pres As Object, force As Boolean) As String
    Dim sl As Object, shp As Object, items As Collection, opened As Collection, books As Object, path As String, k As Long
    Dim wb As Workbook, co As ChartObject, fp As String, missing As String, en As Long, ed As String
    Dim nOk As Long, nSame As Long, nNoChart As Long, nUnlinked As Long, nFail As Long
    Set items = New Collection: Set opened = New Collection
    Set books = CreateObject("Scripting.Dictionary")
    books.CompareMode = 1                                       ' texto: caminhos sem diferenciar maiusculas
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If Len(PptSync_Tag(shp, PPTSYNC_TAG_CHART)) > 0 Then items.Add shp
        Next
    Next
    If items.Count = 0 Then PptSync_Presentation = PPTSYNC_NONE: Exit Function
    On Error GoTo Fail
    For Each shp In items
        k = k + 1
        Application.StatusBar = "ThinQQ: atualizando gráfico " & k & " de " & items.Count & " do PowerPoint..."
        If Ppt_LinkState(shp) = 0 Then
            nUnlinked = nUnlinked + 1
        Else
            path = PptSync_SourcePath(PptSync_LinkSource(shp), PptSync_Tag(shp, PPTSYNC_TAG_BOOK))
            If Not books.Exists(path) Then
                Set wb = Nothing
                If PptSync_FileExists(path) Then Set wb = PptSync_OpenBook(path, opened)
                If wb Is Nothing Then missing = missing & IIf(Len(missing) > 0, ", ", "") & Util_FileName(path)
                books.Add path, wb
            End If
            Set wb = books.Item(path)
            If Not wb Is Nothing Then
                Set co = Registry_FindChartObject(wb, PptSync_Tag(shp, PPTSYNC_TAG_CHART))
                If co Is Nothing Then
                    nNoChart = nNoChart + 1
                Else
                    fp = PptSync_Fingerprint(co, shp)
                    If Not force And fp = PptSync_Tag(shp, PPTSYNC_TAG_FP) Then
                        nSame = nSame + 1
                    ElseIf PptSync_TrySync(shp, co) Then
                        shp.Tags.Add PPTSYNC_TAG_FP, fp
                        nOk = nOk + 1
                    Else
                        nFail = nFail + 1
                    End If
                End If
            End If
        End If
    Next
    PptSync_CloseOpened opened
    Application.StatusBar = False
    PptSync_Presentation = PptSync_SummaryText(nOk, nSame, missing, nNoChart, nUnlinked, nFail)
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptSync_CloseOpened opened
    Application.StatusBar = False
    On Error GoTo 0
    Err.Raise en, "PptSync_Presentation", ed
End Function
```

Em `src/core/modules/modCoreApi.bas`, logo depois de `Core_CopyToPowerPoint`:

```vb
' Parte A (1.4.0): atualiza os graficos ThinQQ vinculados de UMA apresentacao. O suplemento ThinQQ.ppam chama por Application.Run com o
' objeto da apresentacao (ao abrir e no botao do PowerPoint). Devolve o resumo para o usuario; nunca propaga erro.
Public Function Core_PptSyncPresentation(pres As Object, Optional force As Boolean = False) As String
    Dim en As Long, ed As String
    On Error GoTo Fail
    Core_PptSyncPresentation = PptSync_Presentation(pres, force)
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Errors_Log "Core_PptSyncPresentation" & vbTab & en & vbTab & ed
    Core_PptSyncPresentation = "Não foi possível atualizar os gráficos do ThinQQ: " & ed & " (" & en & ")"
End Function

' Contrato com o suplemento do PowerPoint (spec 4.9): o .ppam so chama Core_PptSyncPresentation se Core_PptApiVersion >= o que ele exige.
' Regra: a assinatura de Core_PptSyncPresentation nunca muda; mudanca incompativel = funcao nova + CORE_PPT_API maior (a antiga continua).
Public Function Core_PptApiVersion() As Long
    Core_PptApiVersion = CORE_PPT_API
End Function

' Botao "Atualizar graficos no PowerPoint": todas as apresentacoes abertas que nao sao somente leitura, forcando a atualizacao; o aviso traz
' o resumo de cada uma. Sem PowerPoint aberto (nunca abre um): aviso. Nunca propaga erro.
Public Sub Core_PptSyncOpen()
    Dim app As Object, pres As Object, msg As String, n As Long, en As Long, ed As String
    On Error GoTo Fail
    Set app = Ppt_GetApp(False)
    If app Is Nothing Then Core_Notify PPTSYNC_NO_PPT, vbInformation: Exit Sub
    For Each pres In app.Presentations
        If CLng(pres.ReadOnly) = 0 Then
            n = n + 1
            msg = msg & IIf(Len(msg) > 0, vbLf & vbLf, "") & "'" & CStr(pres.Name) & "': " & PptSync_Presentation(pres, True)
        End If
    Next
    If n = 0 Then msg = PPTSYNC_NO_PPT
    Core_Notify msg, vbInformation
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Application.StatusBar = False
    If Ppt_IsKnownFailure(en) Then
        Core_Notify Ppt_ErrorText(en, ed), vbExclamation
    Else
        Errors_HandleErr "Core_PptSyncOpen", en, ed
    End If
End Sub
```

- [ ] **Passo 4: GREEN** — `-Filter "TestPptSync|TestPowerPoint"` → 0 falhas (relate quantos `_WhenAvailable` rodaram de verdade).
- [ ] **Passo 5: Suíte completa** e **Commit** — `git commit -m "feat: atualizar todos os graficos ThinQQ de uma apresentacao (pasta ausente sem travar, pasta fechada aberta so leitura, pula o que nao mudou)" (+ trailer)`.

---

### Tarefa A3: Faixa do Excel — "Atualizar gráficos no PowerPoint" — modelo: standard

**Files:**
- Modify: `src/loader/ribbon/customUI14.xml` (menu `mnuPpt`), `src/loader/modules/modRibbon.bas` (`Rx_OnPptSync`)
- Modify: `tests/vba/loader/TestRibbon.bas`, `docs/tutorial/tutorial.html` (só a seção `#powerpoint`: o rótulo novo em `<span class="ui">`), `README.md` (item 405 das verificações manuais: o menu PowerPoint deixa de ficar cinza)

**Interfaces:**
- Consumes (A2): `Core_PptSyncOpen`.
- Produces: `btnPptSync` "Atualizar gráficos no PowerPoint" (rótulo exato usado por A8 e Test-Docs); `mnuPpt` sem `getEnabled`.

- [ ] **Passo 1: Testes (RED)** — em `tests/vba/loader/TestRibbon.bas`, `Test_Ribbon_Menus_Structure`: troque as 3 linhas

```vb
    AssertEqual "Rx_GetEnabled", d.SelectSingleNode("//u:menu[@id='mnuPpt']").Attributes.getNamedItem("getEnabled").Text
    AssertTrue d.SelectSingleNode("//u:menu[@id='mnuPpt']/@tag") Is Nothing, "PowerPoint so exige um grafico (sem tag)"
    AssertEqual 2, d.SelectNodes("//u:menu[@id='mnuPpt']/u:button").Length
```
por
```vb
    AssertTrue d.SelectSingleNode("//u:menu[@id='mnuPpt']/@getEnabled") Is Nothing, "PowerPoint sempre ativo: Atualizar graficos no PowerPoint nao exige grafico"
    AssertTrue d.SelectSingleNode("//u:menu[@id='mnuPpt']/@tag") Is Nothing
    AssertEqual 3, d.SelectNodes("//u:menu[@id='mnuPpt']/u:button").Length
    AssertEqual "Rx_GetEnabled", AttrText(d.SelectSingleNode("//u:button[@id='btnPptImage']"), "getEnabled"), "copiar exige grafico"
    AssertEqual "Rx_GetEnabled", AttrText(d.SelectSingleNode("//u:button[@id='btnPptChart']"), "getEnabled")
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:button[@id='btnPptSync'][@onAction='Rx_OnPptSync']") Is Nothing
    AssertTrue d.SelectSingleNode("//u:button[@id='btnPptSync']/@getEnabled") Is Nothing, "Atualizar graficos no PowerPoint: sempre ativo"
    AssertEqual "msepPpt", AttrText(d.SelectSingleNode("//u:button[@id='btnPptChart']/following-sibling::*[1]"), "id")
```
Em `Test_Ribbon_Phase3_LabelsAndTips`, troque `AssertEqual "Copiar para o PowerPoint", AttrText(d.SelectSingleNode("//*[@id='mnuPpt']"), "screentip")` por `AssertEqual "PowerPoint", AttrText(d.SelectSingleNode("//*[@id='mnuPpt']"), "screentip")` e acrescente ao fim do módulo:

```vb
Public Sub Test_Ribbon_PptSync_LabelAndTip()
    Dim d As Object, n As Object
    Set d = LoadRibbon()
    Set n = d.SelectSingleNode("//*[@id='btnPptSync']")
    AssertTrue Not n Is Nothing
    AssertEqual "Atualizar gráficos no PowerPoint", AttrText(n, "label")
    AssertEqual "Refresh", AttrText(n, "imageMso")
    AssertTrue InStr(AttrText(n, "supertip"), "ao abrir a apresentação") > 0, "o supertip cita a atualizacao ao abrir"
End Sub
```

- [ ] **Passo 2: RED** — `pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target loader -Filter "TestRibbon" -TimeoutSec 300` → falhas nos asserts novos.
- [ ] **Passo 3: XML** — em `src/loader/ribbon/customUI14.xml`, troque o bloco do menu PowerPoint inteiro (do comentário `<!-- PowerPoint: ... -->` até `</menu>`) por:

```xml
          <!-- PowerPoint: o menu fica sempre ativo (Atualizar graficos no PowerPoint nao exige grafico); copiar exige um grafico (getEnabled). -->
          <menu id="mnuPpt" label="PowerPoint" image="tqq_ppt16"
                screentip="PowerPoint" supertip="Cola o gráfico selecionado no slide ativo do PowerPoint (como imagem ou gráfico vinculado) e atualiza os gráficos vinculados das apresentações abertas.">
            <button id="btnPptImage" label="Copiar como imagem" onAction="Rx_OnPptImage" getEnabled="Rx_GetEnabled"
                    screentip="Copiar como imagem" supertip="Cola uma imagem do gráfico no slide ativo (abre o PowerPoint se preciso)."/>
            <button id="btnPptChart" label="Copiar como gráfico vinculado" onAction="Rx_OnPptChart" getEnabled="Rx_GetEnabled"
                    screentip="Copiar como gráfico vinculado" supertip="Cola um gráfico do PowerPoint vinculado a esta pasta de trabalho salva, com as setas no tamanho do slide. Com o suplemento do PowerPoint, dados e setas são atualizados ao abrir a apresentação."/>
            <menuSeparator id="msepPpt"/>
            <button id="btnPptSync" label="Atualizar gráficos no PowerPoint" imageMso="Refresh" onAction="Rx_OnPptSync"
                    screentip="Atualizar gráficos no PowerPoint" supertip="Atualiza dados e setas dos gráficos do ThinQQ vinculados nas apresentações abertas no PowerPoint (também acontece ao abrir a apresentação, com o suplemento do PowerPoint)."/>
          </menu>
```

- [ ] **Passo 4: Loader** — em `src/loader/modules/modRibbon.bas`, depois de `Rx_OnPptChart`:

```vb
' Atualizar graficos no PowerPoint (Parte A): nao exige grafico selecionado; core anterior: aviso de atualizar.
Public Sub Rx_OnPptSync(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_PptSyncOpen"
End Sub
```
e acrescente `Core_PptSyncOpen` ao comentário de `Rx_TestSetFeatureProcSuffix` (lista dos recursos sempre ativos).

- [ ] **Passo 5: Documentação mínima (Test-Docs exige todo rótulo da faixa no tutorial)** — em `docs/tutorial/tutorial.html`, na seção `id="powerpoint"`, depois do parágrafo que explica "Copiar como gráfico vinculado", acrescente:

```html
      <p><span class="ui">PowerPoint</span> › <span class="ui">Atualizar gráficos no PowerPoint</span> atualiza dados e setas de todos os gráficos do ThinQQ colados como gráfico vinculado nas apresentações abertas no PowerPoint.</p>
```
e no `README.md`, item 2 das verificações manuais (linha que diz que "os menus Setas e PowerPoint" ficam cinza), troque "(inclusive os menus Setas e PowerPoint)" por "(inclusive o menu Setas; no menu PowerPoint só Atualizar gráficos no PowerPoint fica ativo)".

- [ ] **Passo 6: GREEN + suíte completa** — `-Target loader -Filter "TestRibbon"` → 0 falhas; suíte completa (Test-Ribbon, Test-Docs) verde.
- [ ] **Passo 7: Commit** — `git commit -m "feat: botao Atualizar graficos no PowerPoint (menu PowerPoint sempre ativo)" (+ trailer)`.

---

### Tarefa A4: Suplemento do PowerPoint `ThinQQ.ppam` (fonte, build, testes) — modelo: most-capable

**Pré-requisito:** VBOM do PowerPoint habilitado pelo usuário (veja o topo). Confira com `(Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security' -ErrorAction SilentlyContinue).AccessVBOM` = 1; senão pare e relate `BLOCKED` (não grave o registro).

**Files:**
- Create: `src/ppt/modules/modPptTqq.bas`, `src/ppt/modules/modPptRibbon.bas`, `src/ppt/classes/clsPptEvents.cls`, `src/ppt/ribbon/customUI14.xml`
- Create: `tests/vba/ppt/modPptAssert.bas`, `tests/vba/ppt/modPptTestRunner.bas`, `tests/vba/ppt/TestPptAddin.bas`, `tests/ps/Test-PptAddin.ps1`
- Modify: `build/lib/Build-Common.ps1` (`Assert-PptVbomTrusted`, `Test-PptVbomTrusted`, `New-PptApp`, `Close-PptApp`, `AllTests` no despacho), `build/build.ps1` (alvo `ppt`), `build/run-tests.ps1` (`-Ppt`), `tests/ps/Test-Ribbon.ps1` (XML do PowerPoint), `build/lib/Vba-Lint.ps1` só se ele não varrer `src/ppt` (confira: deve varrer todo `src` e `tests`)

**Interfaces:**
- Consumes (A2): `Core_PptSyncPresentation(pres, force)` por `xl.Run("'ThinQQ.core<sufixo>.xlam'!Core_PptSyncPresentation", pres, force)`; `Core_Ping`; (A1) etiquetas `TQQ_CHART`/`TQQ_BOOK`.
- Produces: `dist\ThinQQ.ppam` (Release) e `dist\ThinQQ.ppt.test.pptm` (Test); `PptTqq_*` públicas: `PptTqq_OnOpenAction`, `PptTqq_NormalizeMode`, `PptTqq_NewestVersion`, `PptTqq_SourcePath`, `PptTqq_CountTagged`, `PptTqq_OnOpen`, `PptTqq_Sync`, `PptTqq_DataOnly`, `PptTqq_ErrorText`, `PptTqq_AskText`, `PptTqq_LastMsg`, `PptTqq_SetExcelForTest`; rótulos da faixa do PowerPoint (A8); `PPT_REQUIRED_API` e o carimbo `Comments` do `.ppam` (A6/A7).

- [ ] **Passo 1: Despacho com lista** — em `build/lib/Build-Common.ps1`, `Write-TestDispatchModule`, acrescente a função `AllTests` ao módulo gerado (o corredor do PowerPoint não tem `ThisWorkbook.VBProject`). Troque o bloco que monta `$content` por:

```powershell
    $body = ($cases -join "`r`n")
    $names = @($cases | ForEach-Object { if ($_ -match 'Case "([^"]+)"') { '    s = s & "' + $Matches[1] + '|"' } })
    $list = ($names -join "`r`n")
    $content = @"
Attribute VB_Name = "modTestDispatch"
' Gerado pelo build - nao editar
Option Explicit

Public Sub Dispatch(ByVal qualifiedName As String)
    Select Case qualifiedName
$body
        Case Else: Err.Raise vbObjectError + 514, "Dispatch", "teste nao despachado: " & qualifiedName
    End Select
End Sub

Public Function AllTests() As String
    Dim s As String
$list
    AllTests = s
End Function
"@
```

- [ ] **Passo 2: Funções de PowerPoint no build** — em `build/lib/Build-Common.ps1`, depois de `Assert-VbomTrusted`:

```powershell
function Test-PptVbomTrusted {
    $v = (Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security' -ErrorAction SilentlyContinue).AccessVBOM
    return ($v -eq 1)
}
function Assert-PptVbomTrusted {
    if (-not (Test-PptVbomTrusted)) {
        throw "Habilite no POWERPOINT 'Confiar no acesso ao modelo de objeto do projeto VBA' (Arquivo > Opções > Central de Confiabilidade > Configurações de Macro). Sem isso o suplemento ThinQQ.ppam não pode ser montado."
    }
}
# PowerPoint e de instancia unica: so cria se NENHUM POWERPNT estiver em execucao (senao CreateObject anexaria ao do usuario).
function New-PptApp {
    if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Há um PowerPoint aberto. Feche-o para montar/testar o suplemento (o PowerPoint é de instância única).' }
    $app = New-Object -ComObject PowerPoint.Application
    $script:PptPid = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
    if ($script:PptPid.Count -ne 1) { throw "PowerPoint do build não identificado ($($script:PptPid.Count) processos)." }
    $app.DisplayAlerts = 1          # ppAlertsNone
    return $app
}
function Close-PptApp($app) {
    if ($null -ne $app) {
        try { foreach ($p in @($app.Presentations)) { $p.Saved = -1; $p.Close() } } catch { }
        try { $app.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 1000
    foreach ($id in @($script:PptPid)) { if ($id -and (Get-Process -Id $id -ErrorAction SilentlyContinue)) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue } }
}
```

- [ ] **Passo 3: Alvo `ppt` no build** — em `build/build.ps1`: `[ValidateSet('loader','core','ppt','all')]`; `'all'` continua `loader` e `core` (o `.ppam` só com `-Target ppt`: as rodadas normais não exigem o PowerPoint). Antes do `foreach ($t in $targets)`, trate o alvo `ppt` à parte e depois saia:

```powershell
if ($Target -eq 'ppt') {
    Assert-PptVbomTrusted
    $srcDir = Join-Path $RepoRoot 'src\ppt'
    $outName = if ($Configuration -eq 'Test') { 'ThinQQ.ppt.test.pptm' } else { 'ThinQQ.ppam' }
    $outPath = Join-Path $DistDir $outName
    if (Test-Path $outPath) { Remove-Item $outPath -Force }
    $buildInfo = Join-Path $env:TEMP 'modBuildInfo_ppt.bas'
    Write-BuildInfoModule -Version $versions.core -Configuration $Configuration -OutPath $buildInfo
    $dispatch = Join-Path $env:TEMP 'modTestDispatch_ppt.bas'
    $pp = New-PptApp
    $marker = Join-Path ([IO.Path]::GetTempPath()) "thinqq_watchdog_ppt_$PID.flag"
    Remove-Item -LiteralPath $marker -Force -ErrorAction SilentlyContinue
    $wd = Start-PidWatchdog $script:PptPid[0] $TimeoutSec $marker
    try {
        $pres = $pp.Presentations.Add(0)
        $vbproj = $pres.VBProject
        $vbproj.Name = 'ThinQQ_ppt'
        Import-VbaFile $vbproj $buildInfo
        Get-ChildItem (Join-Path $srcDir 'modules') -Filter *.bas | ForEach-Object { Import-VbaFile $vbproj $_.FullName }
        Get-ChildItem (Join-Path $srcDir 'classes') -Filter *.cls | ForEach-Object { Import-VbaFile $vbproj $_.FullName }
        if ($Configuration -eq 'Test') {
            $tDir = Join-Path $RepoRoot 'tests\vba\ppt'
            Get-ChildItem $tDir -Filter *.bas | ForEach-Object { Import-VbaFile $vbproj $_.FullName }
            Write-TestDispatchModule -TestFiles @(Get-ChildItem $tDir -Filter 'Test*.bas' | ForEach-Object FullName) -OutPath $dispatch
            Import-VbaFile $vbproj $dispatch
        }
        $pres.BuiltInDocumentProperties.Item('Comments').Value = "ThinQQ ppt $($versions.core) ($Configuration)"   # carimbo lido sem abrir (docProps/core.xml), como o do loader
        $fmt = if ($Configuration -eq 'Test') { 25 } else { 30 }     # ppSaveAsOpenXMLPresentationMacroEnabled / ppSaveAsOpenXMLAddin
        $pres.SaveAs($outPath, $fmt)
        $pres.Close()
    } catch {
        if (Test-Path -LiteralPath $marker) { Write-Host "ERRO: o PowerPoint do build não respondeu em $TimeoutSec s (encerrado)." -ForegroundColor Red; exit 2 } else { throw }
    } finally {
        if ($wd.Proc -and -not $wd.Proc.HasExited) { Stop-Process -Id $wd.Proc.Id -Force -ErrorAction SilentlyContinue }
        Close-PptApp $pp
        Remove-Item $buildInfo, $dispatch -Force -ErrorAction SilentlyContinue
    }
    Add-RibbonToPackage -PackagePath $outPath -RibbonXmlPath (Join-Path $srcDir 'ribbon\customUI14.xml') -Icons Png
    Write-Host "OK  $outPath"
    exit 0
}
```
(`Add-RibbonToPackage` é genérico: com o XML só com `imageMso` não grava imagens. Se o `SaveAs` do formato 30 recusar uma apresentação sem slides, acrescente um slide antes de salvar e relate o desvio.)

- [ ] **Passo 4: Fonte do suplemento** — crie `src/ppt/modules/modPptTqq.bas`:

```vb
Attribute VB_Name = "modPptTqq"
Option Explicit
' Suplemento do PowerPoint do ThinQQ (ThinQQ.ppam; spec 2026-10-14 secao 4.3). Ao abrir uma apresentacao com graficos ThinQQ vinculados
' (etiqueta TQQ_CHART, gravada pelo core ao colar) pergunta (padrao), atualiza sempre ou nao faz nada (opcao "Ao abrir") e pede ao core do
' ThinQQ no Excel para atualizar dados e anotacoes (Core_PptSyncPresentation). Excel: o aberto, se ja tem o core; senao uma instancia NOVA e
' invisivel (o Excel e de varias instancias: nunca a do usuario), com o core mais novo de %LOCALAPPDATA%\ThinQQ\versions, fechada no fim.
' Sem o core: so os dados (vinculo nativo), nunca com a pasta inexistente (o Update travaria o PowerPoint). Ligacao tardia; sem Declare.
Public Const PPT_APP_NAME As String = "ThinQQ"
Public Const PPT_SECTION As String = "PowerPoint"
Public Const PPT_KEY_OPENMODE As String = "OpenMode"
Public Const PPT_TAG_CHART As String = "TQQ_CHART"
Public Const PPT_TAG_BOOK As String = "TQQ_BOOK"
Public Const PPT_MSG_NONE As String = "Esta apresentação não tem gráficos do ThinQQ vinculados."
Public Const PPT_MSG_NO_PRES As String = "Abra uma apresentação primeiro."
Public Const PPT_MSG_NO_EXCEL As String = "O Excel não está instalado (ou não pôde ser iniciado): os gráficos do ThinQQ não foram atualizados."
Public Const PPT_MSG_DATA_ONLY As String = "O ThinQQ não está instalado no Excel deste computador: só os dados foram atualizados; setas e outras anotações do ThinQQ ficaram como estavam."
Public Const PPT_MSG_OLD_CORE As String = "Atualize o ThinQQ no Excel para atualizar também as setas."
Public Const PPT_MSG_BUSY As String = "O Excel está ocupado (uma célula em edição ou uma caixa de diálogo aberta). Termine o que está fazendo no Excel e use ThinQQ > Atualizar gráficos."
Public Const PPT_REQUIRED_API As Long = 1                   ' Core_PptApiVersion minima (spec 4.9); core sem ela ou menor = so os dados
Private Const RPC_E_CALL_REJECTED As Long = -2147418111
Private Const RPC_E_RETRYLATER As Long = -2147417846
Private mEvents As clsPptEvents
Private mSeen As Collection
Private mTestXl As Object
Private mTestCoreDir As String
Private mLastMsg As String

' Carga do suplemento: liga os eventos do PowerPoint (o projeto de teste .pptm nao liga).
Public Sub Auto_Open()
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then Exit Sub
    Set mEvents = New clsPptEvents
    Set mEvents.App = Application
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Set mEvents = Nothing
End Sub

' Gancho de teste: Excel injetado (o do teste, com o core de teste aberto) e pasta das versoes do core. Inerte fora do build Test.
Public Sub PptTqq_SetExcelForTest(xl As Object, coreDir As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    Set mTestXl = xl
    mTestCoreDir = coreDir
End Sub

Public Function PptTqq_LastMsg() As String
    PptTqq_LastMsg = mLastMsg
End Function

Public Sub PptTqq_ClearLastMsg()
    mLastMsg = ""
End Sub

' Unico ponto de MsgBox de aviso (log-only no build Test).
Public Sub PptTqq_Msg(msg As String, Optional style As VbMsgBoxStyle = vbInformation)
    mLastMsg = msg
    If BUILD_CONFIG = "Test" Then Exit Sub
    MsgBox msg, style, PPT_APP_NAME
End Sub

' Unico ponto de pergunta (no build Test registra e responde Nao).
Private Function PptTqq_Ask(msg As String) As VbMsgBoxResult
    mLastMsg = msg
    If BUILD_CONFIG = "Test" Then PptTqq_Ask = vbNo: Exit Function
    PptTqq_Ask = MsgBox(msg, vbQuestion + vbYesNo, PPT_APP_NAME)
End Function

Public Function PptTqq_NormalizeMode(m As String) As String
    Select Case LCase$(Trim$(m))
        Case "always": PptTqq_NormalizeMode = "always"
        Case "never": PptTqq_NormalizeMode = "never"
        Case Else: PptTqq_NormalizeMode = "ask"
    End Select
End Function

' Modo "Ao abrir" (HKCU\Software\VB and VBA Program Settings\ThinQQ\PowerPoint\OpenMode). Testes: sempre "ask", sem ler o registro.
Public Function PptTqq_OpenMode() As String
    If BUILD_CONFIG = "Test" Then PptTqq_OpenMode = "ask": Exit Function
    PptTqq_OpenMode = PptTqq_NormalizeMode(GetSetting(PPT_APP_NAME, PPT_SECTION, PPT_KEY_OPENMODE, "ask"))
End Function

Public Sub PptTqq_SetOpenMode(m As String)
    If BUILD_CONFIG = "Test" Then Exit Sub
    SaveSetting PPT_APP_NAME, PPT_SECTION, PPT_KEY_OPENMODE, PptTqq_NormalizeMode(m)
End Sub

' O que fazer ao abrir: "" (nada), "ask" ou "sync". Pura.
Public Function PptTqq_OnOpenAction(mode As String, isReadOnly As Boolean, nTagged As Long, alreadySeen As Boolean) As String
    If nTagged = 0 Or alreadySeen Or isReadOnly Or mode = "never" Then Exit Function
    If mode = "always" Then PptTqq_OnOpenAction = "sync" Else PptTqq_OnOpenAction = "ask"
End Function

Public Function PptTqq_AskText(n As Long) As String
    PptTqq_AskText = "Esta apresentação tem " & n & " gráfico(s) do ThinQQ vinculado(s) ao Excel. Atualizar agora com os dados atuais?" & vbLf & vbLf & "Para não perguntar mais: aba ThinQQ > Ao abrir."
End Function

Private Function PptTqq_Tag(shp As Object, key As String) As String
    On Error Resume Next
    PptTqq_Tag = CStr(shp.Tags.Item(key))
    Err.Clear
End Function

' Formas soltas marcadas pelo ThinQQ na apresentacao.
Public Function PptTqq_CountTagged(pres As Object) As Long
    Dim sl As Object, shp As Object
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0 Then PptTqq_CountTagged = PptTqq_CountTagged + 1
        Next
    Next
End Function

' Mesma regra do core (PptSync_SourcePath). Pura.
Public Function PptTqq_SourcePath(linkFull As String, tagBook As String) As String
    Dim s As String, p As Long
    s = Trim$(linkFull)
    p = InStr(s, "!")
    If p > 0 Then s = Left$(s, p - 1)
    If Len(s) = 0 Then s = Trim$(tagBook)
    PptTqq_SourcePath = s
End Function

Private Function PptTqq_IsVersion(s As String) As Boolean
    Dim p() As String, i As Long
    p = Split(s, ".")
    If UBound(p) <> 2 Then Exit Function
    For i = 0 To 2
        If Len(p(i)) = 0 Or Not IsNumeric(p(i)) Or InStr(p(i), ",") > 0 Or InStr(p(i), "-") > 0 Then Exit Function
    Next
    PptTqq_IsVersion = True
End Function

' Compara x.y.z numericamente (1 se a > b, -1 se a < b, 0 iguais).
Private Function PptTqq_CompareVersion(a As String, b As String) As Long
    Dim pa() As String, pb() As String, i As Long
    pa = Split(a, "."): pb = Split(b, ".")
    For i = 0 To 2
        If CLng(pa(i)) > CLng(pb(i)) Then PptTqq_CompareVersion = 1: Exit Function
        If CLng(pa(i)) < CLng(pb(i)) Then PptTqq_CompareVersion = -1: Exit Function
    Next
End Function

' Maior versao x.y.z numa lista "a|b|c" (nomes de pastas); "" sem nenhuma valida. Pura.
Public Function PptTqq_NewestVersion(names As String) As String
    Dim p() As String, i As Long, best As String
    p = Split(names, "|")
    For i = LBound(p) To UBound(p)
        If PptTqq_IsVersion(p(i)) Then
            If Len(best) = 0 Then
                best = p(i)
            ElseIf PptTqq_CompareVersion(p(i), best) > 0 Then
                best = p(i)
            End If
        End If
    Next
    PptTqq_NewestVersion = best
End Function

Private Function PptTqq_CoreFileName() As String
    PptTqq_CoreFileName = "ThinQQ.core" & BUILD_SUFFIX & ".xlam"
End Function

Private Function PptTqq_CoreDir() As String
    If BUILD_CONFIG = "Test" Then PptTqq_CoreDir = mTestCoreDir: Exit Function
    PptTqq_CoreDir = Environ$("LOCALAPPDATA") & "\ThinQQ\versions"
End Function

' Caminho do core mais novo instalado ("" se nao ha).
Private Function PptTqq_CorePath() As String
    Dim d As String, nm As String, names As String, v As String, p As String
    d = PptTqq_CoreDir()
    If Len(d) = 0 Then Exit Function
    On Error Resume Next
    nm = Dir$(d & "\*", vbDirectory)
    Do While Len(nm) > 0
        If nm <> "." And nm <> ".." Then names = names & nm & "|"
        nm = Dir$()
    Loop
    On Error GoTo 0
    v = PptTqq_NewestVersion(names)
    If Len(v) = 0 Then Exit Function
    p = d & "\" & v & "\" & PptTqq_CoreFileName()
    If Len(Dir$(p)) > 0 Then PptTqq_CorePath = p
End Function

Private Function PptTqq_Ping(xl As Object) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = xl.Run("'" & PptTqq_CoreFileName() & "'!Core_Ping")
    PptTqq_Ping = (Err.Number = 0 And CStr(v) = "pong")
    Err.Clear
End Function

' Excel a usar: no build Test, so o injetado. Senao: o aberto (GetObject) SE ele tem o core; senao uma instancia nova e invisivel (created).
Private Function PptTqq_GetExcel(ByRef created As Boolean) As Object
    Dim xl As Object
    created = False
    If BUILD_CONFIG = "Test" Then Set PptTqq_GetExcel = mTestXl: Exit Function
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    If Not xl Is Nothing Then
        If PptTqq_Ping(xl) Then Set PptTqq_GetExcel = xl: Exit Function
    End If
    Set xl = Nothing
    Err.Clear
    Set xl = CreateObject("Excel.Application")
    If xl Is Nothing Then Exit Function
    On Error GoTo 0
    created = True
    xl.Visible = False
    xl.DisplayAlerts = False
    Set PptTqq_GetExcel = xl
End Function

' Nome do core respondendo no Excel ("" sem core). Na instancia criada por nos, abre o core mais novo instalado (somente leitura).
Private Function PptTqq_EnsureCore(xl As Object, created As Boolean) As String
    Dim p As String
    If PptTqq_Ping(xl) Then PptTqq_EnsureCore = PptTqq_CoreFileName(): Exit Function
    If Not created Then Exit Function
    p = PptTqq_CorePath()
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    xl.Workbooks.Open p, 0, True
    Err.Clear
    On Error GoTo 0
    If PptTqq_Ping(xl) Then PptTqq_EnsureCore = PptTqq_CoreFileName()
End Function

Private Sub PptTqq_Release(xl As Object, created As Boolean)
    On Error Resume Next
    If created And Not xl Is Nothing Then
        xl.DisplayAlerts = False
        xl.Quit
    End If
    Set xl = Nothing
End Sub

Public Function PptTqq_ErrorText(en As Long, ed As String) As String
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then
        PptTqq_ErrorText = PPT_MSG_BUSY
    ElseIf en = 429 Then
        PptTqq_ErrorText = PPT_MSG_NO_EXCEL
    Else
        PptTqq_ErrorText = "Não foi possível atualizar os gráficos do ThinQQ: " & ed & " (" & en & ")"
    End If
End Function

' Modo degradado (sem o core): so os dados pelo vinculo nativo, e so dos graficos cuja pasta existe.
Public Function PptTqq_DataOnly(pres As Object) As String
    Dim sl As Object, shp As Object, n As Long, nMiss As Long, p As String, lf As String
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0 Then
                lf = ""
                On Error Resume Next
                lf = CStr(shp.LinkFormat.SourceFullName)
                Err.Clear
                On Error GoTo 0
                p = PptTqq_SourcePath(lf, PptTqq_Tag(shp, PPT_TAG_BOOK))
                If PptTqq_FileExists(p) Then
                    On Error Resume Next
                    shp.LinkFormat.Update
                    If Err.Number = 0 Then n = n + 1
                    Err.Clear
                    On Error GoTo 0
                Else
                    nMiss = nMiss + 1
                End If
            End If
        Next
    Next
    PptTqq_DataOnly = "Dados atualizados em " & n & " gráfico(s)" & IIf(nMiss > 0, "; " & nMiss & " sem a pasta de origem (ficaram como estavam)", "") & "."
End Function

Private Function PptTqq_FileExists(p As String) As Boolean
    On Error Resume Next
    If Len(p) = 0 Then Exit Function
    PptTqq_FileExists = CreateObject("Scripting.FileSystemObject").FileExists(p)
    Err.Clear
End Function

' Versao da API do core (0 = core sem Core_PptApiVersion, anterior a 1.4.0). Pura quanto ao suplemento: so le.
Private Function PptTqq_CoreApi(xl As Object, core As String) As Long
    Dim v As Variant
    On Error Resume Next
    v = xl.Run("'" & core & "'!Core_PptApiVersion")
    If Err.Number = 0 And IsNumeric(v) Then PptTqq_CoreApi = CLng(v)
    Err.Clear
End Function

' Contrato (spec 4.9): so chama a API se o core a tem na versao exigida. Pura.
Public Function PptTqq_ApiCompatible(coreApi As Long) As Boolean
    PptTqq_ApiCompatible = (coreApi >= PPT_REQUIRED_API)
End Function

' Core_PptSyncPresentation no Excel. Excel ocupado: o erro sobe (aviso proprio). Core antigo ou API menor: so os dados + aviso.
Private Function PptTqq_RunCoreSync(xl As Object, core As String, pres As Object, force As Boolean) As String
    Dim v As Variant, en As Long, ed As String
    If Not PptTqq_ApiCompatible(PptTqq_CoreApi(xl, core)) Then PptTqq_RunCoreSync = PptTqq_DataOnly(pres) & vbLf & PPT_MSG_OLD_CORE: Exit Function
    On Error Resume Next
    v = xl.Run("'" & core & "'!Core_PptSyncPresentation", pres, force)
    en = Err.Number: ed = Err.Description
    Err.Clear
    On Error GoTo 0
    If en = 0 Then PptTqq_RunCoreSync = CStr(v): Exit Function
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then Err.Raise en, "PptTqq", ed
    PptTqq_RunCoreSync = PptTqq_DataOnly(pres) & vbLf & PPT_MSG_OLD_CORE
End Function

' Atualiza os graficos ThinQQ da apresentacao (force: inclusive os que nao mudaram). Aviso final com o resumo. Nunca propaga erro.
Public Sub PptTqq_Sync(pres As Object, force As Boolean)
    Dim xl As Object, created As Boolean, core As String, msg As String, en As Long, ed As String
    On Error GoTo Fail
    If PptTqq_CountTagged(pres) = 0 Then PptTqq_Msg PPT_MSG_NONE: Exit Sub
    Set xl = PptTqq_GetExcel(created)
    If xl Is Nothing Then PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation: Exit Sub
    core = PptTqq_EnsureCore(xl, created)
    If Len(core) > 0 Then
        msg = PptTqq_RunCoreSync(xl, core, pres, force)
    Else
        msg = PptTqq_DataOnly(pres) & vbLf & PPT_MSG_DATA_ONLY
    End If
    PptTqq_Release xl, created
    PptTqq_Msg msg, vbInformation
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptTqq_Release xl, created
    PptTqq_Msg PptTqq_ErrorText(en, ed), vbExclamation
End Sub

Private Function PptTqq_Seen(key As String) As Boolean
    Dim v As Variant
    If mSeen Is Nothing Then Exit Function
    On Error Resume Next
    v = mSeen(key)
    PptTqq_Seen = (Err.Number = 0)
    Err.Clear
End Function

Private Sub PptTqq_MarkSeen(key As String)
    If mSeen Is Nothing Then Set mSeen = New Collection
    On Error Resume Next
    mSeen.Add key, key
    Err.Clear
End Sub

' Gancho de teste: esquece as apresentacoes ja vistas nesta sessao.
Public Sub PptTqq_ResetSeenForTest()
    Set mSeen = Nothing
End Sub

' Eventos (AfterPresentationOpen e WindowActivate): uma vez por apresentacao (caminho) e por sessao. Nunca propaga erro.
Public Sub PptTqq_OnOpen(pres As Object)
    Dim n As Long, act As String, key As String
    On Error GoTo Fail
    key = LCase$(CStr(pres.FullName))
    n = PptTqq_CountTagged(pres)
    act = PptTqq_OnOpenAction(PptTqq_OpenMode(), (CLng(pres.ReadOnly) <> 0), n, PptTqq_Seen(key))
    If n > 0 Then PptTqq_MarkSeen key
    Select Case act
        Case "sync"
            PptTqq_Sync pres, False
        Case "ask"
            If PptTqq_Ask(PptTqq_AskText(n)) = vbYes Then PptTqq_Sync pres, False
    End Select
    Exit Sub
Fail:
    PptTqq_Msg "Não foi possível verificar os gráficos do ThinQQ desta apresentação: " & Err.Description, vbExclamation
End Sub
```

`src/ppt/modules/modPptRibbon.bas`:

```vb
Attribute VB_Name = "modPptRibbon"
Option Explicit
' Faixa do suplemento do PowerPoint (aba ThinQQ, grupo "Graficos do Excel"). Nunca propaga erro.
Private mRibbon As IRibbonUI

Public Sub PptRx_OnLoad(ribbon As IRibbonUI)
    On Error Resume Next
    Set mRibbon = ribbon
End Sub

Public Sub PptRx_OnSync(control As IRibbonControl)
    On Error Resume Next
    If Application.Presentations.Count = 0 Then PptTqq_Msg PPT_MSG_NO_PRES: Exit Sub
    PptTqq_Sync Application.ActivePresentation, True
End Sub

' Itens do dropdown "Ao abrir": modeask, modealways, modenever.
Public Sub PptRx_OnOpenMode(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    On Error Resume Next
    PptTqq_SetOpenMode Mid$(selectedId, 5)
End Sub

Public Sub PptRx_GetOpenMode(control As IRibbonControl, ByRef itemId)
    On Error Resume Next
    itemId = "mode" & PptTqq_OpenMode()
End Sub

Public Sub PptRx_OnAbout(control As IRibbonControl)
    On Error Resume Next
    PptTqq_Msg "ThinQQ para PowerPoint " & APP_VERSION & vbLf & "Atualiza dados e setas dos gráficos do ThinQQ copiados como gráfico vinculado ao abrir a apresentação (aba ThinQQ > Ao abrir).", vbInformation
End Sub
```
(Se `IRibbonUI`/`IRibbonControl` não compilarem no projeto do PowerPoint — a biblioteca do Office deveria vir por padrão —, troque os tipos por `Object` e relate.)

`src/ppt/classes/clsPptEvents.cls`:

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPptEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Eventos do PowerPoint (suplemento ThinQQ.ppam). Nunca propagam erro. WindowActivate cobre a apresentacao que sai do Modo de Exibicao
' Protegido ("Habilitar Edicao") e a que ja estava aberta quando o suplemento carregou; PptTqq_OnOpen trata cada uma uma vez por sessao.
Public WithEvents App As Application

Private Sub App_AfterPresentationOpen(ByVal Pres As Presentation)
    On Error Resume Next
    PptTqq_OnOpen Pres
End Sub

Private Sub App_WindowActivate(ByVal Pres As Presentation, ByVal Wn As DocumentWindow)
    On Error Resume Next
    PptTqq_OnOpen Pres
End Sub
```

`src/ppt/ribbon/customUI14.xml` (só `imageMso`; o `Assert-TqqRibbonXml` do build valida):

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!-- Faixa do suplemento do PowerPoint do ThinQQ (ThinQQ.ppam). So imageMso. -->
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="PptRx_OnLoad">
  <ribbon>
    <tabs>
      <tab id="tabThinQQPpt" label="ThinQQ">
        <group id="grpPptSync" label="Gráficos do Excel">
          <button id="btnPptSyncNow" label="Atualizar gráficos" size="large" imageMso="Refresh" onAction="PptRx_OnSync"
                  screentip="Atualizar gráficos do ThinQQ" supertip="Atualiza dados e setas de todos os gráficos do ThinQQ colados como gráfico vinculado nesta apresentação, pelo ThinQQ do Excel."/>
          <dropDown id="ddOpenMode" label="Ao abrir" getSelectedItemID="PptRx_GetOpenMode" onAction="PptRx_OnOpenMode"
                    screentip="Ao abrir a apresentação" supertip="O que fazer ao abrir uma apresentação com gráficos do ThinQQ vinculados: perguntar, atualizar sempre ou não atualizar.">
            <item id="modeask" label="Perguntar"/>
            <item id="modealways" label="Atualizar sempre"/>
            <item id="modenever" label="Não atualizar"/>
          </dropDown>
          <button id="btnPptAbout" label="Sobre" imageMso="Info" onAction="PptRx_OnAbout" screentip="Sobre o ThinQQ para PowerPoint"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

- [ ] **Passo 5: Testes do projeto do PowerPoint** — `tests/vba/ppt/modPptAssert.bas`:

```vb
Attribute VB_Name = "modPptAssert"
Option Explicit
' Asserts do projeto de teste do PowerPoint (o modAssert comum depende do Excel).
Public Sub AssertTrue(cond As Boolean, Optional msg As String = "esperado True")
    If Not cond Then Err.Raise vbObjectError + 100, "Assert", msg
End Sub
Public Sub AssertFalse(cond As Boolean, Optional msg As String = "esperado False")
    If cond Then Err.Raise vbObjectError + 100, "Assert", msg
End Sub
Public Sub AssertEqual(expected As Variant, actual As Variant, Optional msg As String = "")
    If expected <> actual Then Err.Raise vbObjectError + 100, "Assert", "esperado [" & expected & "] obtido [" & actual & "] " & msg
End Sub
```

`tests/vba/ppt/modPptTestRunner.bas`:

```vb
Attribute VB_Name = "modPptTestRunner"
Option Explicit
' Corredor do projeto de teste do PowerPoint: roda modTestDispatch.AllTests() (lista gerada pelo build) e grava TAP. Filtro: TQQ_TEST_FILTER.
Private mPassed As Long, mFailed As Long, mLog As String

Public Function RunAllPptTests(outputPath As String) As Long
    Dim names() As String, i As Long, filt As String
    filt = Environ$("TQQ_TEST_FILTER")
    mPassed = 0: mFailed = 0: mLog = ""
    names = Split(modTestDispatch.AllTests(), "|")
    For i = LBound(names) To UBound(names)
        If Len(names(i)) > 0 Then
            If Len(filt) = 0 Or InStr(1, names(i), filt, vbTextCompare) > 0 Then RunOne names(i)
        End If
    Next
    mLog = mLog & "1.." & (mPassed + mFailed) & vbLf & "# passed " & mPassed & " failed " & mFailed & vbLf
    WriteText outputPath, mLog
    RunAllPptTests = mFailed
End Function

Private Sub RunOne(qualified As String)
    On Error GoTo Fail
    modTestDispatch.Dispatch qualified
    mPassed = mPassed + 1
    mLog = mLog & "ok " & (mPassed + mFailed) & " - " & qualified & vbLf
    Exit Sub
Fail:
    mFailed = mFailed + 1
    mLog = mLog & "not ok " & (mPassed + mFailed) & " - " & qualified & " # " & Replace(Err.Description, vbLf, " ") & vbLf
End Sub

Private Sub WriteText(p As String, s As String)
    Dim f As Integer
    f = FreeFile
    Open p For Output As #f
    Print #f, s;
    Close #f
End Sub
```

`tests/vba/ppt/TestPptAddin.bas`:

```vb
Attribute VB_Name = "TestPptAddin"
Option Explicit
' Funcoes do suplemento do PowerPoint (projeto de teste .pptm aberto numa instancia propria e sem janela).

Private Function NewPres() As Object
    Set NewPres = Application.Presentations.Add(0)
    NewPres.Slides.Add 1, 12
End Function

Private Sub ClosePres(p As Object)
    On Error Resume Next
    p.Saved = -1
    p.Close
End Sub

Public Sub Test_PptTqq_OnOpenAction()
    AssertEqual "ask", PptTqq_OnOpenAction("ask", False, 2, False)
    AssertEqual "sync", PptTqq_OnOpenAction("always", False, 2, False)
    AssertEqual "", PptTqq_OnOpenAction("never", False, 2, False)
    AssertEqual "", PptTqq_OnOpenAction("ask", False, 0, False), "sem graficos marcados"
    AssertEqual "", PptTqq_OnOpenAction("always", True, 2, False), "somente leitura"
    AssertEqual "", PptTqq_OnOpenAction("always", False, 2, True), "ja tratada nesta sessao"
End Sub

Public Sub Test_PptTqq_NormalizeModeAndVersions()
    AssertEqual "always", PptTqq_NormalizeMode(" Always ")
    AssertEqual "ask", PptTqq_NormalizeMode("qualquer")
    AssertEqual "1.10.0", PptTqq_NewestVersion("1.2.0|1.10.0|1.3.1|x|1.2|")
    AssertEqual "", PptTqq_NewestVersion("logs|tmp|")
    AssertEqual "C:\a\b.xlsx", PptTqq_SourcePath("C:\a\b.xlsx!Dados!x", "")
    AssertEqual "C:\t.xlsx", PptTqq_SourcePath("", "C:\t.xlsx")
End Sub

Public Sub Test_PptTqq_ApiCompatible()
    AssertTrue PptTqq_ApiCompatible(1): AssertTrue PptTqq_ApiCompatible(2), "core mais novo mantem a API 1"
    AssertFalse PptTqq_ApiCompatible(0), "core sem Core_PptApiVersion (anterior a 1.4.0): so os dados"
End Sub

Public Sub Test_PptTqq_ErrorTexts()
    AssertEqual PPT_MSG_BUSY, PptTqq_ErrorText(-2147418111, "x")
    AssertEqual PPT_MSG_NO_EXCEL, PptTqq_ErrorText(429, "x")
    AssertTrue InStr(PptTqq_ErrorText(70, "Permissão negada"), "Permissão negada (70)") > 0
    AssertTrue InStr(PptTqq_AskText(3), "3 gráfico(s) do ThinQQ") > 0
End Sub

Public Sub Test_PptTqq_CountTagged_OnOpenAsksOnce()
    Dim p As Object, s As Object, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    AssertEqual 0, PptTqq_CountTagged(p)
    PptTqq_ClearLastMsg
    PptTqq_ResetSeenForTest
    PptTqq_OnOpen p
    AssertEqual "", PptTqq_LastMsg(), "sem graficos marcados: nada"
    s.Tags.Add PPT_TAG_CHART, "abc123"
    AssertEqual 1, PptTqq_CountTagged(p)
    PptTqq_OnOpen p
    AssertEqual PptTqq_AskText(1), PptTqq_LastMsg(), "modo perguntar (Test responde Nao)"
    PptTqq_ClearLastMsg
    PptTqq_OnOpen p
    AssertEqual "", PptTqq_LastMsg(), "uma vez por sessao"
Cleanup:
    en = Err.Number: ed = Err.Description          ' antes de ClosePres: o On Error dela zera o Err
    PptTqq_ResetSeenForTest
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

Public Sub Test_PptTqq_SyncWithoutTaggedAndDataOnlyMissing()
    Dim p As Object, s As Object, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ClearLastMsg
    PptTqq_Sync p, True
    AssertEqual PPT_MSG_NONE, PptTqq_LastMsg()
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    s.Tags.Add PPT_TAG_CHART, "abc123"
    s.Tags.Add PPT_TAG_BOOK, "C:\nao\existe\pasta-ausente.xlsx"
    AssertEqual "Dados atualizados em 0 gráfico(s); 1 sem a pasta de origem (ficaram como estavam).", PptTqq_DataOnly(p)
Cleanup:
    en = Err.Number: ed = Err.Description
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub
```

- [ ] **Passo 6: Ponta a ponta** — crie `tests/ps/Test-PptAddin.ps1`:

```powershell
# Suplemento do PowerPoint: roda a suite VBA do projeto de teste (.pptm) e um teste de ponta a ponta (Excel proprio com o core de teste +
# PowerPoint proprio): colar vinculado, mudar o dado, PptTqq_Sync com o Excel injetado. So roda sem POWERPNT aberto (instancia unica).
param([string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path)
$ErrorActionPreference = 'Stop'
. (Join-Path $RepoRoot 'build\lib\Build-Common.ps1')
$fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }
$pptm = Join-Path $DistDir 'ThinQQ.ppt.test.pptm'
$coreTest = Join-Path $DistDir 'ThinQQ.core.test.xlam'
if (-not (Test-Path $pptm) -or -not (Test-Path $coreTest)) { Write-Host 'ERRO: monte antes (build.ps1 -Target all -Configuration Test e -Target ppt -Configuration Test)'; exit 1 }
$tmp = Join-Path $env:TEMP ('thinqq_test_ppt_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $tmp | Out-Null
$pp = $null; $xl = $null
$pp = New-PptApp
$marker = Join-Path ([IO.Path]::GetTempPath()) "thinqq_watchdog_pptaddin_$PID.flag"
$wd = Start-PidWatchdog $script:PptPid[0] 600 $marker
try {
    $tp = $pp.Presentations.Open($pptm, -1, 0, 0)     # somente leitura, sem janela
    $tpName = [string]$tp.Name
    $out = Join-Path $RepoRoot 'tests\results\ppt.tap'
    Remove-Item $out -ErrorAction SilentlyContinue
    $failed = $pp.Run("$tpName!modPptTestRunner.RunAllPptTests", $out)
    Write-Host '=== ppt (VBA) ==='
    Get-Content $out | ForEach-Object { Write-Host $_ }
    Check ([int]$failed -eq 0) "suite VBA do suplemento ($failed falha(s))"

    Write-Host '=== ppt (ponta a ponta) ==='
    $xl = New-ExcelApp
    $xl.EnableEvents = $true
    $core = $xl.Workbooks.Open($coreTest)
    $cn = [string]$core.Name
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'Dados'
    $ws.Range('A1:C1').Value2 = @('Ano', 'Sul', 'Norte')
    $ws.Range('A2').Value2 = 'T1'; $ws.Range('B2').Value2 = 10; $ws.Range('C2').Value2 = 5
    $ws.Range('A3').Value2 = 'T2'; $ws.Range('B3').Value2 = 20; $ws.Range('C3').Value2 = 6
    $wb.SaveAs((Join-Path $tmp 'e2e.xlsx'), 51)
    $ws.Range('A1:C3').Select() | Out-Null
    [void]$xl.Run("'$cn'!Core_Insert", 'stacked')
    [void]$xl.Run("'$cn'!Core_AddAnnotation", 'cagr', 1, 2)
    $pres = $pp.Presentations.Add(0)
    $slide = $pres.Slides.Add(1, 12)
    [void]$xl.Run("'$cn'!Ppt_SetAppForTest", $pp)
    [void]$xl.Run("'$cn'!Ppt_PinTargetForTest", $slide, $false)
    [void]$xl.Run("'$cn'!Core_CopyToPowerPoint", 'chart')
    Check ($slide.Shapes.Count -eq 1) 'grafico colado'
    $ws.Range('B3').Value2 = 40
    [void]$pp.Run("$tpName!PptTqq_SetExcelForTest", $xl, '')
    [void]$pp.Run("$tpName!PptTqq_Sync", $pres, $true)
    $msg = [string]$pp.Run("$tpName!PptTqq_LastMsg")
    Check ($msg -like 'Gráficos do ThinQQ atualizados: 1.*') "resumo do core pelo suplemento: $msg"
    $vals = @($slide.Shapes.Item(1).Chart.SeriesCollection(1).Values)
    Check ($vals -contains 40) "valor novo no slide ($($vals -join ','))"
    [void]$pp.Run("$tpName!PptTqq_SetExcelForTest", $null, '')
    [void]$xl.Run("'$cn'!Ppt_ResetTestHooks")
    $pres.Saved = -1; $pres.Close()
    $wb.Close($false)
} catch {
    Write-Host "not ok - excecao: $($_.Exception.Message)"; $fail++
} finally {
    if ($wd.Proc -and -not $wd.Proc.HasExited) { Stop-Process -Id $wd.Proc.Id -Force -ErrorAction SilentlyContinue }
    if ($xl) { Close-ExcelApp $xl }
    Close-PptApp $pp
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
if (Test-Path -LiteralPath $marker) { Write-Host 'not ok - o PowerPoint do teste não respondeu (encerrado pelo watchdog)'; $fail++ }
if ($fail -gt 0) { exit 1 } else { exit 0 }
```
(Se `Application.Run` do PowerPoint não aceitar o nome com pontos `ThinQQ.ppt.test.pptm!Modulo.Macro`, tente `'ThinQQ.ppt.test.pptm'!Modulo.Macro` e relate a forma que funcionou. `New-ExcelApp` deixa `EnableEvents` falso: a linha seguinte liga.)

- [ ] **Passo 7: `run-tests -Ppt`** — em `build/run-tests.ps1`, acrescente o parâmetro `[switch]$Ppt,` e, antes do bloco `-LoaderSwap`:

```powershell
if (-not $timedOut) {
    if ($Ppt) {
        Write-Host '=== powerpoint (suplemento ThinQQ.ppam: suite VBA + ponta a ponta) ==='
        if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Write-Host 'SKIP: há um PowerPoint aberto (instância única); feche-o para rodar esta parte.' -ForegroundColor Yellow }
        elseif (-not (Test-PptVbomTrusted)) { Write-Host "SKIP: habilite no PowerPoint 'Confiar no acesso ao modelo de objeto do projeto VBA'." -ForegroundColor Yellow }
        else {
            $global:LASTEXITCODE = 0
            & "$PSScriptRoot\build.ps1" -Target ppt -Configuration Test
            if ($LASTEXITCODE -ne 0) { $totalFailed += 1 }
            else {
                $global:LASTEXITCODE = 0
                & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $RepoRoot 'tests\ps\Test-PptAddin.ps1') -RepoRoot $RepoRoot
                if ($LASTEXITCODE -ne 0) { $totalFailed += 1 }
            }
        }
    } else { Write-Host '=== powerpoint: pulado (use -Ppt) ===' }
}
```
(Com `-Filter`, o run-tests sai antes: a parte do PowerPoint só roda na suíte sem filtro.)

- [ ] **Passo 8: Test-Ribbon** — em `tests/ps/Test-Ribbon.ps1`, acrescente (antes do resumo final) a validação do XML do PowerPoint:

```powershell
# ---- Faixa do suplemento do PowerPoint (src/ppt/ribbon/customUI14.xml): bem formada, so imageMso, ids e callbacks esperados
$pptXmlPath = Join-Path $RepoRoot 'src\ppt\ribbon\customUI14.xml'
[xml]$px = [IO.File]::ReadAllText($pptXmlPath, (New-Object Text.UTF8Encoding($false)))
$pns = New-Object Xml.XmlNamespaceManager($px.NameTable); $pns.AddNamespace('u', 'http://schemas.microsoft.com/office/2009/07/customui')
Check ($px.customUI.onLoad -eq 'PptRx_OnLoad') 'ppt: onLoad PptRx_OnLoad'
Check ($null -eq $px.SelectSingleNode('//u:*[@image]', $pns)) 'ppt: nenhum image= (so imageMso)'
Check ($null -ne $px.SelectSingleNode("//u:button[@id='btnPptSyncNow'][@onAction='PptRx_OnSync']", $pns)) 'ppt: Atualizar gráficos'
Check ((@($px.SelectNodes("//u:dropDown[@id='ddOpenMode']/u:item", $pns) | ForEach-Object { $_.id }) -join ',') -eq 'modeask,modealways,modenever') 'ppt: itens de Ao abrir'
```
(Use a função `Check` que o script já tem; se o nome for outro, adapte.)

- [ ] **Passo 9: Rodar** — `robocopy ...`; `pwsh C:\temp\tqq-pd\build\build.ps1 -Target ppt -Configuration Release` → `OK ...\dist\ThinQQ.ppam` (confira com `[IO.Compression.ZipFile]::OpenRead(...)` que existem `ppt/vbaProject.bin` e `customUI/customUI14.xml`); `pwsh C:\temp\tqq-pd\build\run-tests.ps1 -TimeoutSec 900 -Ppt` → "Todos os testes passaram" com a parte do PowerPoint executada (não "SKIP"; se o usuário tiver PowerPoint aberto, relate e peça ao controlador uma janela sem PowerPoint).
- [ ] **Passo 10: Commit** — `git add src/ppt tests/vba/ppt tests/ps/Test-PptAddin.ps1 tests/ps/Test-Ribbon.ps1 build/build.ps1 build/run-tests.ps1 build/lib/Build-Common.ps1` e `git commit -m "feat: suplemento do PowerPoint ThinQQ.ppam (atualiza graficos do ThinQQ ao abrir a apresentacao)" (+ trailer)`.

---

### Tarefa A5: Instalador e desinstalador com o suplemento do PowerPoint — modelo: standard

**Files:**
- Modify: `installer/install.ps1`, `installer/uninstall.ps1`, `installer/Instalar-ThinQQ.cmd`, `installer/LEIA-ME.txt`
- Test: `tests/ps/Test-Installer.ps1`

**Interfaces:**
- Consumes (A4/A6): `ThinQQ.ppam` no pacote e `pptAddinSha256` no manifesto (A6 grava; aqui os testes criam pacotes sintéticos).
- Produces: `%LOCALAPPDATA%\ThinQQ\ThinQQ.ppam`; `PpamVersion` em `HKCU\Software\VB and VBA Program Settings\ThinQQ\Updater` (lida pela A7; pendência antiga do atualizador descartada ao reinstalar); `HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\ThinQQ` (`Path`, `AutoLoad` = 1); `...\PowerPoint\Security\Trusted Locations\LocationThinQQ`.

- [ ] **Passo 1: Testes (RED)** — em `tests/ps/Test-Installer.ps1`: (a) em `New-Pkg`, acrescente o parâmetro `[switch]$WithPpam` e, quando ligado, `Set-Content (Join-Path $p 'ThinQQ.ppam') 'ppam'` e os campos `pptAddinVersion = '1.2.3'` e `pptAddinSha256 = (Sha $p 'ThinQQ.ppam')` no hashtable do manifesto; (b) acrescente as variáveis `$pptAddSub = 'Microsoft\Office\16.0\PowerPoint\AddIns\ThinQQ'` e `$pptTlSub = 'Microsoft\Office\16.0\PowerPoint\Security\Trusted Locations\LocationThinQQ'`; (c) antes de `# ---- Caminho com acentos`, o bloco:

```powershell
    # ---- Suplemento do PowerPoint (1.4.0): copiado, registrado (AddIns\ThinQQ: Path e AutoLoad) e local confiavel; pacote sem .ppam continua valido
    $regP = New-Reg; $rootP = Join-Path $tmp 'rp\ThinQQ'
    $pkgP = New-Pkg 'pkgP' -WithPpam
    & $install -Root $rootP -RegistryRoot $regP -PackageDir $pkgP -SkipExcelCheck
    Assert (Test-Path (Join-Path $rootP 'ThinQQ.ppam')) 'ppam copiado para a raiz'
    $ak = Get-ItemProperty (Join-Path $regP $pptAddSub)
    Assert ($ak.Path -eq (Join-Path $rootP 'ThinQQ.ppam')) "AddIns\ThinQQ\Path: $($ak.Path)"
    Assert ($ak.AutoLoad -eq 1) 'AutoLoad = 1'
    Assert ((Get-ItemProperty (Join-Path $regP $pptTlSub)).Path -eq ($rootP + '\')) 'local confiavel do PowerPoint'
    Assert ((Get-ItemProperty (Join-Path $regP $upSub)).PpamVersion -eq '1.2.3') 'PpamVersion gravada (o atualizador do Excel compara com pptAddinVersion)'
    Set-ItemProperty (Join-Path $regP $upSub) -Name PendingPpam -Value '1'
    Set-Content (Join-Path $rootP 'ThinQQ.ppam.new') 'velho'
    & $install -Root $rootP -RegistryRoot $regP -PackageDir $pkgP -SkipExcelCheck
    Assert ((Get-ItemProperty (Join-Path $regP $pptAddSub)).AutoLoad -eq 1) 'reinstalacao idempotente'
    Assert (-not (Test-Path (Join-Path $rootP 'ThinQQ.ppam.new'))) 'reinstalar descarta a pendencia antiga do atualizador'
    Assert ($null -eq (Get-ItemProperty (Join-Path $regP $upSub)).PendingPpam) 'PendingPpam limpo'
    & $uninstall -Root $rootP -RegistryRoot $regP -SkipExcelCheck
    Assert (-not (Test-Path (Join-Path $regP $pptAddSub))) 'AddIns\ThinQQ removido'
    Assert (-not (Test-Path (Join-Path $regP $pptTlSub))) 'local confiavel do PowerPoint removido'
    # pacote sem .ppam (ZIP antigo): instala o Excel e nao cria chaves do PowerPoint
    $regQ = New-Reg; $rootQ = Join-Path $tmp 'rq\ThinQQ'
    & $install -Root $rootQ -RegistryRoot $regQ -PackageDir $pkg -SkipExcelCheck
    Assert (-not (Test-Path (Join-Path $rootQ 'ThinQQ.ppam'))) 'sem ppam no pacote: nada'
    Assert (-not (Test-Path (Join-Path $regQ $pptAddSub))) 'sem ppam: sem AddIns'
    & $uninstall -Root $rootQ -RegistryRoot $regQ -SkipExcelCheck
    # SHA divergente do ppam recusa e nada fica; manifesto com o campo e sem o arquivo recusa
    $pkgX = New-Pkg 'pkgX' -WithPpam
    Set-Content (Join-Path $pkgX 'ThinQQ.ppam') 'adulterado'
    $regX = New-Reg; $rootX = Join-Path $tmp 'rx\ThinQQ'
    Expect-Failure 'ppam com SHA divergente recusado' { & $install -Root $rootX -RegistryRoot $regX -PackageDir $pkgX -SkipExcelCheck } 'SHA-256 de ThinQQ.ppam'
    Assert (-not (Test-Path $rootX)) 'nada copiado com ppam adulterado'
    Remove-Item (Join-Path $pkgX 'ThinQQ.ppam')
    Expect-Failure 'manifesto cita o ppam ausente' { & $install -Root $rootX -RegistryRoot $regX -PackageDir $pkgX -SkipExcelCheck } 'ThinQQ.ppam'
    # falha injetada depois da copia: sem ppam nem chaves do PowerPoint
    $regY = New-Reg; $rootY = Join-Path $tmp 'ry\ThinQQ'
    Expect-Failure 'falha injetada (ppam)' { & $install -Root $rootY -RegistryRoot $regY -PackageDir $pkgP -SkipExcelCheck -SimulateFailureForTest } 'simulada'
    Assert (-not (Test-Path $rootY)) 'falha: nada fica'
    Assert (-not (Test-Path (Join-Path $regY $pptAddSub))) 'falha: sem AddIns'
```

- [ ] **Passo 2: RED** — `pwsh <repo>\tests\ps\Test-Installer.ps1` → "FALHA: ppam copiado para a raiz".
- [ ] **Passo 3: `install.ps1`** — (1) troque a checagem de processo por: `if (-not $SkipExcelCheck) { if (Get-Process excel -ErrorAction SilentlyContinue) { throw 'Feche o Excel antes de instalar o ThinQQ.' }; if ((Test-Path -LiteralPath (Join-Path $PackageDir 'ThinQQ.ppam')) -and (Get-Process powerpnt -ErrorAction SilentlyContinue)) { throw 'Feche o PowerPoint antes de instalar o ThinQQ (o suplemento do PowerPoint é atualizado junto).' } }` — atenção: essa checagem fica **depois** da leitura de `$PackageDir` (mova-a para logo antes de "Validacao do pacote" se preciso); (2) depois de `Test-PackageHash 'ThinQQ.xlam' 'loaderSha256'`:

```powershell
# Suplemento do PowerPoint (opcional: ZIPs anteriores a 1.4.0 nao tem). Manifesto que o cita exige o arquivo e o SHA-256.
$ppamSrc = Join-Path $PackageDir 'ThinQQ.ppam'
$hasPpamField = $manifest.PSObject.Properties['pptAddinSha256'] -and ([string]$manifest.pptAddinSha256).Trim()
$withPpam = Test-Path -LiteralPath $ppamSrc
if ($hasPpamField -and -not $withPpam) { throw "Arquivo 'ThinQQ.ppam' não encontrado em $PackageDir. Descompacte o ZIP inteiro antes de instalar." }
if ($withPpam) {
    if (-not $hasPpamField) { throw 'manifest.json sem pptAddinSha256; o suplemento do PowerPoint não pode ser verificado.' }
    $exp = ([string]$manifest.pptAddinSha256).Trim().ToLower()
    $act = (Get-FileHash -LiteralPath $ppamSrc -Algorithm SHA256).Hash.ToLower()
    if ($act -ne $exp) { throw 'Pacote corrompido ou adulterado: o SHA-256 de ThinQQ.ppam não confere com o manifest.json. Baixe o ZIP novamente.' }
}
$ppamDest = Join-Path $Root 'ThinQQ.ppam'
$hadPpam = Test-Path -LiteralPath $ppamDest
$ppamMoved = $false
```
(3) no laço que copia para o staging, acrescente `'ThinQQ.ppam'` à lista (`foreach ($f in @('ThinQQ.xlam', 'ThinQQ.core.xlam', 'tutorial.html', 'ThinQQ.ppam'))` — ele já só copia o que existe); (4) no backup do desfazer: `if ($hadPpam) { Copy-Item -LiteralPath $ppamDest -Destination (Join-Path $undoDir 'ppam') -Force }`; (5) depois de mover o tutorial: `if ($withPpam) { Move-Item -LiteralPath (Join-Path $stage 'ThinQQ.ppam') -Destination $ppamDest -Force; $ppamMoved = $true }`; (6) depois do bloco "Local confiavel" do Excel:

```powershell
    # --- PowerPoint: local confiavel e registro do suplemento (AddIns\ThinQQ: Path, AutoLoad) ---
    if ($withPpam) {
        $ptl = Join-Path $RegistryRoot 'Microsoft\Office\16.0\PowerPoint\Security\Trusted Locations\LocationThinQQ'
        if (-not (Test-Path -LiteralPath $ptl)) {
            Initialize-RegKey $ptl
            [void]$undo.Add({ Remove-Item -LiteralPath $ptl -Recurse -Force -ErrorAction SilentlyContinue }.GetNewClosure())
        }
        Set-ItemProperty -LiteralPath $ptl -Name Path -Value ($Root + '\')
        Set-ItemProperty -LiteralPath $ptl -Name AllowSubfolders -Value 1 -Type DWord
        Set-ItemProperty -LiteralPath $ptl -Name Description -Value 'ThinQQ add-in'
        $pak = Join-Path $RegistryRoot 'Microsoft\Office\16.0\PowerPoint\AddIns\ThinQQ'
        if (-not (Test-Path -LiteralPath $pak)) {
            Initialize-RegKey $pak
            [void]$undo.Add({ Remove-Item -LiteralPath $pak -Recurse -Force -ErrorAction SilentlyContinue }.GetNewClosure())
        }
        Set-ItemProperty -LiteralPath $pak -Name Path -Value $ppamDest
        Set-ItemProperty -LiteralPath $pak -Name AutoLoad -Value 1 -Type DWord
        # versao instalada do suplemento (spec 4.9: o atualizador do Excel compara com pptAddinVersion e so atualiza quem ja tem o .ppam)
        $ppVer = [string]$manifest.pptAddinVersion
        if ($ppVer -notmatch '^\d+\.\d+\.\d+$') { $ppVer = $coreVer }
        $upk = Join-Path $RegistryRoot 'VB and VBA Program Settings\ThinQQ\Updater'
        Initialize-RegKey $upk
        $oldPpVer = (Get-ItemProperty -LiteralPath $upk).PSObject.Properties['PpamVersion']
        $oldPpVal = $null; if ($oldPpVer) { $oldPpVal = [string]$oldPpVer.Value }
        Set-ItemProperty -LiteralPath $upk -Name PpamVersion -Value $ppVer
        [void]$undo.Add({ if ($null -ne $oldPpVal) { Set-ItemProperty -LiteralPath $upk -Name PpamVersion -Value $oldPpVal } else { Remove-ItemProperty -LiteralPath $upk -Name PpamVersion -ErrorAction SilentlyContinue } }.GetNewClosure())
        # pendencia antiga do atualizador (ThinQQ.ppam.new) perde o sentido: o ZIP acabou de instalar a versao dele
        foreach ($n in @('PendingPpam', 'PendingPpamSha256', 'PendingPpamVersion')) { Remove-ItemProperty -LiteralPath $upk -Name $n -ErrorAction SilentlyContinue }
        Remove-Item -LiteralPath ($ppamDest + '.new') -Force -ErrorAction SilentlyContinue
    }
```
(7) a falha simulada (`if ($SimulateFailureForTest) { throw ... }`) passa para **depois** do bloco do PowerPoint (o teste exige que as chaves novas sejam desfeitas) — confira que os testes antigos de falha continuam verdes; (8) no `catch`, junto do desfazer do tutorial: `if ($ppamMoved) { if ($hadPpam) { Copy-Item -LiteralPath (Join-Path $undoDir 'ppam') -Destination $ppamDest -Force } else { Remove-Item -LiteralPath $ppamDest -Force -ErrorAction SilentlyContinue } }`; (9) mensagem final: acrescente `if ($withPpam) { Write-Host "Suplemento do PowerPoint instalado: abra o PowerPoint e procure a aba 'ThinQQ'." }`.
- [ ] **Passo 4: `uninstall.ps1`** — a checagem vira `if (Get-Process excel -ErrorAction SilentlyContinue) { throw ... }` e `if (Get-Process powerpnt -ErrorAction SilentlyContinue) { throw 'Feche o PowerPoint antes de desinstalar o ThinQQ.' }` (ambas sob `-not $SkipExcelCheck`); antes das remoções de `Trusted Locations` do Excel: `Remove-Item -LiteralPath (Join-Path $RegistryRoot 'Microsoft\Office\16.0\PowerPoint\AddIns\ThinQQ') -Recurse -Force -ErrorAction SilentlyContinue` e `Remove-Item -LiteralPath (Join-Path $RegistryRoot 'Microsoft\Office\16.0\PowerPoint\Security\Trusted Locations\LocationThinQQ') -Recurse -Force -ErrorAction SilentlyContinue`.
- [ ] **Passo 5: Textos** — `Instalar-ThinQQ.cmd`: "Feche o Excel e tente de novo" → "Feche o Excel e o PowerPoint e tente de novo"; "Abra o Excel e procure a aba ThinQQ." → "Abra o Excel (e o PowerPoint) e procure a aba ThinQQ.". `installer/LEIA-ME.txt`: no passo de instalação, "Feche o Excel" → "Feche o Excel e o PowerPoint"; acrescente um parágrafo "PowerPoint: o instalador também instala o suplemento ThinQQ.ppam (aba ThinQQ no PowerPoint), que atualiza os gráficos do ThinQQ colados como gráfico vinculado ao abrir a apresentação. Instalação manual: PowerPoint > Arquivo > Opções > Suplementos > Gerenciar: Suplementos do PowerPoint > Ir... > Adicionar novo > %LOCALAPPDATA%\ThinQQ\ThinQQ.ppam." (se o Test-Docs conferir frases do LEIA-ME, rode-o).
- [ ] **Passo 6: GREEN** — `pwsh <repo>\tests\ps\Test-Installer.ps1` e `powershell -NoProfile -ExecutionPolicy Bypass -File <repo>\tests\ps\Test-Installer.ps1` → "Test-Installer: OK" nas duas; suíte completa.
- [ ] **Passo 7: Commit** — `git commit -m "feat: instalador registra o suplemento do PowerPoint (AddIns e local confiavel); desinstalador remove" (+ trailer)`.

---

### Tarefa A6: Release, manifesto e publicação (6 assets) com o `.ppam` — modelo: standard

**Files:**
- Modify: `build/lib/Release-Lib.ps1`, `build/release.ps1`, `tools/lib/Release-Publish.ps1`, `tools/Publish-Release.ps1` (só o `.SYNOPSIS`)
- Test: `tests/ps/Test-ReleaseGuard.ps1`, `tests/ps/Test-PublishRelease.ps1`

**Interfaces:**
- Consumes (A4): `build.ps1 -Target ppt -Configuration Release` → `dist\ThinQQ.ppam` (com o carimbo `Comments = "ThinQQ ppt x.y.z (Release)"`).
- Produces: manifesto com `pptAddinVersion` (= `coreVersion`: core e suplemento saem juntos), `pptAddinUrl`, `pptAddinSha256` (spec §4.9); ZIP com 9 arquivos soltos; `Get-ReleaseAssetNames` com 6; consumidos pela A7 (atualizador) e pela A5 (instalador grava `PpamVersion`).

- [ ] **Passo 1: Testes (RED)** — (a) `Test-ReleaseGuard.ps1`: onde ele confere `Get-ReleaseTopFiles`/`Test-ReleaseZipContent`, acrescente `ThinQQ.ppam` aos esperados e um caso de ZIP sem o `.ppam` ("ausente no ZIP: ThinQQ.ppam"); `New-ReleaseManifest ... -PptSha ('a' * 64)` gera `pptAddinVersion` igual a `coreVersion`, `pptAddinUrl` terminando em `/ThinQQ.ppam` e `pptAddinSha256` igual; sem `-PptSha`, os **três** campos não aparecem; um `.ppam` de teste (zip com `ppt/vbaProject.bin` contendo o marcador `modTestRunner`) é recusado por `Get-TestMarkersInXlam`. (b) `Test-PublishRelease.ps1`: contagens de 5 para 6 assets (procure "5 assets", `Count -eq 5`, listas de nomes) e `ThinQQ.ppam` nos arquivos falsos que ele cria.
- [ ] **Passo 2: `Release-Lib.ps1`** — `Get-ReleaseTopFiles` passa a `@('ThinQQ.xlam', 'ThinQQ.core.xlam', 'ThinQQ.ppam', 'manifest.json', 'tutorial.html') + @(Get-ReleaseInstallerFiles)`; `New-ReleaseManifest` ganha `[string]$PptSha = ''` e, quando não vazio, acrescenta ao hashtable ordenado (depois de `loaderSha256`) `pptAddinVersion = $CoreVersion`, `pptAddinUrl = Join-ReleaseUrl $BaseUrl 'ThinQQ.ppam'` e `pptAddinSha256 = $PptSha`; `Get-TestMarkersInXlamBytes` procura `xl/vbaProject.bin` **ou** `ppt/vbaProject.bin` (`$e = $z.GetEntry('xl/vbaProject.bin'); if (-not $e) { $e = $z.GetEntry('ppt/vbaProject.bin') }; if (-not $e) { throw 'vbaProject.bin ausente no pacote' }`).
- [ ] **Passo 3: `release.ps1`** — depois do build `-Target all -Configuration Release`: `& "$PSScriptRoot\build.ps1" -Target ppt -Configuration Release` (falha ⇒ `throw "Build do suplemento do PowerPoint falhou (código $LASTEXITCODE). Habilite o VBOM do PowerPoint e feche o PowerPoint."`); acrescente `'ThinQQ.ppam'` ao laço que confere marcadores de teste; a assinatura passa a assinar também `(Join-Path $DistDir 'ThinQQ.ppam')` (o SIP do Office cobre `.ppam`; antes dos SHA-256); `New-ReleaseManifest ... -PptSha (Get-FileHash (Join-Path $DistDir 'ThinQQ.ppam') -Algorithm SHA256).Hash.ToLower()`.
- [ ] **Passo 4: `Release-Publish.ps1`** — `Get-ReleaseAssetNames` (linha ~225) devolve `@("ThinQQ-$Version.zip", 'ThinQQ.xlam', 'ThinQQ.core.xlam', 'ThinQQ.ppam', 'manifest.json', 'tutorial.html')`; troque "5 assets"/"5 arquivos" por 6 nos comentários e textos (e no `.SYNOPSIS` de `tools/Publish-Release.ps1`); a preparação guarda o `.ppam` em `dist\release-<versao>\` junto com os demais.
- [ ] **Passo 5: GREEN** — suíte completa (as suítes PowerShell de release/publish rodam em pwsh e 5.1).
- [ ] **Passo 6: Commit** — `git commit -m "build: release com ThinQQ.ppam (6 assets; manifesto pptAddinVersion/Url/Sha256)" (+ trailer)`.

---

### Tarefa A7: Atualização automática do suplemento do PowerPoint pelo atualizador do Excel — modelo: most-capable

Fluxo e decisões: spec §4.9 (o que segue é a implementação). Resumo: o atualizador do loader (que já verifica na abertura do Excel a cada `CheckIntervalHours`, baixa pelo `Transport` com o mesmo proxy/TLS e confere SHA-256) passa a tratar o `.ppam` como parte do mesmo pacote do core: **só atualiza um suplemento já instalado** (`<raiz>\ThinQQ.ppam` existe), compara `pptAddinVersion` com a versão instalada (`PpamVersion` no registro do ThinQQ), **baixa e confere o `.ppam` antes de trocar o core**, troca o core e depois o `.ppam` por **renomeações atômicas** (`ThinQQ.ppam` → `.old`, baixado → `ThinQQ.ppam`, apaga `.old`). Arquivo em uso pelo PowerPoint (a renomeação falha: medido, spec §3.4) ⇒ fica `ThinQQ.ppam.new` + `PendingPpam`/`PendingPpamSha256`/`PendingPpamVersion`, aplicado **automaticamente** pelo Excel na próxima abertura (`Loader_Startup`) **e** por uma nova tentativa a cada 10 min enquanto o Excel estiver aberto (`Application.OnTime`, que se reagenda só enquanto houver pendência), sempre reconferindo o SHA-256 do `.new`. Nunca abre nem executa o `.ppam`.

**Files:**
- Modify: `src/loader/modules/modUpdater.bas`, `src/loader/modules/modSettings.bas` (chaves), `src/loader/modules/modLoaderMain.bas` (`Loader_Startup`), `src/loader/forms/frmUpdate.json` + `frmUpdate.vba` (caixa "Atualizar também o suplemento do PowerPoint"), `src/loader/modules/modRibbon.bas` (`Rx_OnSettings` pergunta a opção)
- Modify: `build/run-tests.ps1` (fixture `tests\fixtures\release\ThinQQ.ppam` + campos no manifesto de teste)
- Test: `tests/vba/loader/TestUpdater.bas`

**Interfaces:**
- Consumes (A6): manifesto com `pptAddinVersion`/`pptAddinUrl`/`pptAddinSha256`; (A5) `PpamVersion` gravado pelo instalador; (A2/A4) contrato de versão da API (`Core_PptApiVersion`, `PPT_REQUIRED_API`).
- Produces: chaves `KEY_PPAM_VERSION = "PpamVersion"`, `KEY_PPAM_AUTO = "PpamAutoUpdate"` (padrão "1"), `KEY_PENDING_PPAM = "PendingPpam"`, `KEY_PENDING_PPAM_SHA = "PendingPpamSha256"`, `KEY_PENDING_PPAM_VER = "PendingPpamVersion"`; `Updater_PpamInstalled`, `Updater_PpamNeeded(json)`, `Updater_PpamDownload(json, tmp)`, `Updater_PpamCommit(tmp, sha, ver)` (devolve `"replaced"`, `"replaced-restart"` ou `"staged"`), `Updater_ApplyPendingPpam`, `Updater_PpamStatusText(state)`; ganchos `Updater_TestSetPptRunning`, `Updater_TestSetPpamAuto`, contador `Updater_PpamRetryScheduleCount`.

- [ ] **Passo 1: Fixture de teste** — em `build/run-tests.ps1`, logo depois de copiar `ThinQQ.core.test.xlam` para `tests\fixtures\release`: crie um `.ppam` falso (`[IO.File]::WriteAllBytes((Join-Path $relDir 'ThinQQ.ppam'), [Text.Encoding]::ASCII.GetBytes('fake-ppam-9.9.9'))`), calcule `$ppSha` e acrescente ao hashtable do manifesto de teste `pptAddinVersion = '9.9.9'; pptAddinUrl = "http://127.0.0.1:$httpPort/release/ThinQQ.ppam"; pptAddinSha256 = $ppSha`.
- [ ] **Passo 2: Testes (RED)** — em `tests/vba/loader/TestUpdater.bas`: acrescente ao `SettingKeys()` as 5 chaves novas (e aumente `mSaved(0 To 14)`); em `Setup`/`Cleanup`, zere `Updater_TestSetPptRunning Empty`, `Updater_TestSetPpamAuto Empty` e `Updater_PpamRetryScheduleCount = 0`. Acrescente:

```vb
' Raiz temporaria com um suplemento "instalado" (conteudo antigo) e a versao instalada 1.0.0.
Private Function PpamRoot(Optional withPpam As Boolean = True) As String
    Dim root As String
    root = TempDir()
    If withPpam Then Transport_WriteTextFile root & "\ThinQQ.ppam", "ppam-antigo"
    Loader_SetRootOverride root
    Settings_Set KEY_PPAM_VERSION, "1.0.0"
    PpamRoot = root
End Function

Private Function PpamJson(url As String, sha As String, ver As String) As String
    PpamJson = "{""schemaVersion"":1,""coreVersion"":""0.0.1"",""loaderVersion"":""0.0.1"",""pptAddinVersion"":""" & ver & """,""pptAddinUrl"":""" & Replace(url, "\", "\\") & """,""pptAddinSha256"":""" & sha & """}"
End Function

Private Function FixturePpam() As String
    FixturePpam = RepoPath() & "\tests\fixtures\release\ThinQQ.ppam"
End Function

' (5) Nunca instala o suplemento para quem nao o tem: sem <raiz>\ThinQQ.ppam nada e baixado nem criado.
Public Sub Test_Updater_Ppam_NotInstalled_NeverInstalls()
    Dim root As String, json As String
    On Error GoTo Fail
    Setup
    root = PpamRoot(False)
    json = PpamJson(FixturePpam(), Transport_Sha256File(FixturePpam()), "9.9.9")
    AssertFalse Updater_PpamInstalled(), "sem o arquivo: nao instalado"
    AssertFalse Updater_PpamNeeded(json)
    AssertEqual "", Updater_PpamUpdateNow(json), "nada a fazer"
    AssertEqual 0, Len(Dir(root & "\ThinQQ.ppam")), "nada criado"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (1)(2)(7) Instalado e mais antigo, PowerPoint fechado: baixa pelo servidor HTTP de teste, confere o SHA-256 e troca atomicamente; grava a
' versao; nao sobra .old, .new nem temporario.
Public Sub Test_Updater_Ppam_ReplacesWhenUnlocked_FromHttp()
    Dim root As String, json As String, st As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Updater_TestSetPptRunning False
    json = Transport_FetchText(HttpBase() & "release/manifest.json")
    AssertEqual "9.9.9", JsonGetString(json, "pptAddinVersion")
    AssertTrue Updater_PpamNeeded(json)
    st = Updater_PpamUpdateNow(json)
    AssertEqual "replaced", st
    AssertEqual "fake-ppam-9.9.9", Transport_ReadTextFile(root & "\ThinQQ.ppam")
    AssertEqual "9.9.9", Settings_Get(KEY_PPAM_VERSION)
    AssertEqual 0, Len(Dir(root & "\ThinQQ.ppam.old")) + Len(Dir(root & "\ThinQQ.ppam.new")) + Len(Dir(Loader_TmpDir() & "\ThinQQ.ppam"))
    AssertFalse Updater_PpamNeeded(json), "ja atualizado"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (2) Arquivo em uso (aberto com trava exclusiva, como o PowerPoint faz com um arquivo aberto para edicao: medido): fica .new pendente,
' agenda nova tentativa; liberado o arquivo, a aplicacao pendente reconfere o SHA-256 e troca.
Public Sub Test_Updater_Ppam_Locked_StagesThenAppliesPending()
    Dim root As String, json As String, f As Integer, st As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Updater_TestSetPptRunning True
    json = PpamJson(FixturePpam(), Transport_Sha256File(FixturePpam()), "9.9.9")
    f = FreeFile
    Open root & "\ThinQQ.ppam" For Binary Access Read Lock Read Write As #f
    st = Updater_PpamUpdateNow(json)
    AssertEqual "staged", st
    AssertTrue Len(Dir(root & "\ThinQQ.ppam.new")) > 0
    AssertEqual "1", Settings_Get(KEY_PENDING_PPAM)
    AssertEqual "9.9.9", Settings_Get(KEY_PENDING_PPAM_VER)
    AssertEqual 1, Updater_PpamRetryScheduleCount, "nova tentativa agendada"
    AssertTrue InStr(Updater_PpamStatusText(st), "quando o PowerPoint for fechado") > 0
    Updater_ApplyPendingPpam
    AssertEqual "1", Settings_Get(KEY_PENDING_PPAM), "ainda em uso: continua pendente"
    Close #f
    AssertEqual "ppam-antigo", Transport_ReadTextFile(root & "\ThinQQ.ppam"), "o arquivo em uso nao foi tocado"
    Updater_ApplyPendingPpam
    AssertEqual "fake-ppam-9.9.9", Transport_ReadTextFile(root & "\ThinQQ.ppam")
    AssertEqual "", Settings_Get(KEY_PENDING_PPAM): AssertEqual "9.9.9", Settings_Get(KEY_PPAM_VERSION)
    AssertEqual 0, Len(Dir(root & "\ThinQQ.ppam.new"))
    Cleanup
    Exit Sub
Fail:
    On Error Resume Next
    Close #f
    On Error GoTo 0
    FailCleanup
End Sub

' (7) SHA-256 divergente: nada trocado, temporario apagado. .new adulterado depois de pendente: recusado e apagado.
Public Sub Test_Updater_Ppam_HashMismatch_AndTamperedPending()
    Dim root As String, json As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Updater_TestSetPptRunning False
    json = PpamJson(FixturePpam(), String$(64, "0"), "9.9.9")
    AssertEqual "", Updater_PpamUpdateNow(json), "SHA errado: nada"
    AssertEqual "ppam-antigo", Transport_ReadTextFile(root & "\ThinQQ.ppam")
    AssertEqual 0, Len(Dir(Loader_TmpDir() & "\ThinQQ.ppam")), "temporario apagado"
    Transport_WriteTextFile root & "\ThinQQ.ppam.new", "adulterado"
    Settings_Set KEY_PENDING_PPAM, "1": Settings_Set KEY_PENDING_PPAM_SHA, Transport_Sha256File(FixturePpam()): Settings_Set KEY_PENDING_PPAM_VER, "9.9.9"
    Updater_ApplyPendingPpam
    AssertEqual "ppam-antigo", Transport_ReadTextFile(root & "\ThinQQ.ppam"), "pendente adulterado nao entra"
    AssertEqual 0, Len(Dir(root & "\ThinQQ.ppam.new")), "e apagado"
    AssertEqual "", Settings_Get(KEY_PENDING_PPAM)
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (4) Opcao desligada: nada; (2) PowerPoint aberto mas arquivo livre (suplemento carregado sem trava): troca e pede reinicio.
Public Sub Test_Updater_Ppam_AutoOff_And_RestartNotice()
    Dim root As String, json As String, st As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    json = PpamJson(FixturePpam(), Transport_Sha256File(FixturePpam()), "9.9.9")
    Updater_TestSetPpamAuto "0"
    AssertFalse Updater_PpamNeeded(json), "opcao desligada"
    Updater_TestSetPpamAuto Empty
    Updater_TestSetPptRunning True
    st = Updater_PpamUpdateNow(json)
    AssertEqual "replaced-restart", st
    AssertTrue InStr(Updater_PpamStatusText(st), "feche e abra o PowerPoint") > 0
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (3) Ordem: com core + suplemento, o .ppam e baixado e conferido ANTES da troca do core; .ppam indisponivel nao derruba o core, mas o
' resultado diz que o suplemento fica para a proxima verificacao (nunca em silencio).
Public Sub Test_Updater_Apply_CoreOk_PpamFails_IsReported()
    Dim root As String, json As String, ok As Boolean
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    json = Transport_FetchText(HttpBase() & "release/manifest.json")
    json = Replace(json, "/release/ThinQQ.ppam", "/release/nao-existe.ppam")
    ok = Updater_Apply(json, "core", True)
    AssertTrue ok, "o core e atualizado"
    AssertTrue Len(Dir(root & "\versions\9.9.9\" & Loader_CoreFileName())) > 0
    AssertEqual "ppam-antigo", Transport_ReadTextFile(root & "\ThinQQ.ppam"), "suplemento intocado"
    AssertEqual "1.0.0", Settings_Get(KEY_PPAM_VERSION)
    AssertTrue InStr(Updater_LastPpamNote(), "próxima verificação") > 0, Updater_LastPpamNote()
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (1)(4) So o suplemento atrasado (core em dia): a verificacao silenciosa o atualiza sem abrir o formulario; a linha de versao do formulario
' cita o suplemento quando ele entra junto.
Public Sub Test_Updater_Check_PpamOnly_Silent_And_VersionLine()
    Dim root As String, json As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Updater_TestSetPptRunning False
    json = PpamJson(FixturePpam(), Transport_Sha256File(FixturePpam()), "9.9.9")
    json = Replace(json, """coreVersion"":""0.0.1""", """coreVersion"":""" & Loader_CurrentCoreVersion() & """")
    UseManifestFile json
    Updater_FormShownCount = 0
    Updater_Check True
    AssertEqual 0, Updater_FormShownCount, "so o suplemento: sem formulario"
    AssertEqual "fake-ppam-9.9.9", Transport_ReadTextFile(root & "\ThinQQ.ppam")
    Settings_Set KEY_PPAM_VERSION, "1.0.0"
    AssertTrue InStr(Updater_VersionLine("core", "1.0.0", Replace(json, """coreVersion"":""" & Loader_CurrentCoreVersion() & """", """coreVersion"":""9.9.9""")), "Suplemento do PowerPoint: 1.0.0 > 9.9.9") > 0
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub
```

- [ ] **Passo 3: RED** — `pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target loader -Filter "TestUpdater" -TimeoutSec 400` → saída 3.
- [ ] **Passo 4: Chaves** — em `src/loader/modules/modSettings.bas`, no formato das outras:

```vb
Public Const KEY_PPAM_VERSION As String = "PpamVersion"             ' versao do ThinQQ.ppam instalado (instalador e atualizador gravam)
Public Const KEY_PPAM_AUTO As String = "PpamAutoUpdate"             ' "0" = nao atualizar o suplemento do PowerPoint (padrao "1")
Public Const KEY_PENDING_PPAM As String = "PendingPpam"             ' "1" = ThinQQ.ppam.new aguardando o PowerPoint liberar o arquivo
Public Const KEY_PENDING_PPAM_SHA As String = "PendingPpamSha256"
Public Const KEY_PENDING_PPAM_VER As String = "PendingPpamVersion"
```
- [ ] **Passo 5: Atualizador** — em `src/loader/modules/modUpdater.bas`, acrescente (depois de `Updater_SchedulePendingLoader`):

```vb
' ---------- Suplemento do PowerPoint (ThinQQ.ppam; spec 2026-10-14 secao 4.9) ----------
Public Updater_PpamRetryScheduleCount As Long          ' gancho de teste: novas tentativas que seriam agendadas (Test nao agenda OnTime)
Private mTestPptRunning As Variant                     ' gancho de teste: Empty = consulta real
Private mTestPpamAuto As Variant                       ' gancho de teste: Empty = configuracao real
Private mLastPpamNote As String                        ' ultima frase sobre o suplemento (acrescentada ao aviso do atualizador)
Private Const PPAM_RETRY_MINUTES As Long = 10

Public Sub Updater_TestSetPptRunning(v As Variant)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestPptRunning = v
End Sub

Public Sub Updater_TestSetPpamAuto(v As Variant)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestPpamAuto = v
End Sub

Public Function Updater_LastPpamNote() As String
    Updater_LastPpamNote = mLastPpamNote
End Function

Private Function PpamPath() As String
    PpamPath = Loader_Root() & "\ThinQQ.ppam"
End Function

' PowerPoint em execucao (WMI; nunca GetObject). Na duvida (WMI indisponivel): em execucao.
Private Function Updater_PptRunning() As Boolean
    If BUILD_CONFIG = "Test" And Not IsEmpty(mTestPptRunning) Then Updater_PptRunning = CBool(mTestPptRunning): Exit Function
    On Error Resume Next
    Updater_PptRunning = True
    Updater_PptRunning = (GetObject("winmgmts:" & Chr(92) & Chr(92) & "." & Chr(92) & "root" & Chr(92) & "cimv2").ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'POWERPNT.EXE'").Count > 0)
End Function

Private Function Updater_PpamAutoOn() As Boolean
    If BUILD_CONFIG = "Test" And Not IsEmpty(mTestPpamAuto) Then Updater_PpamAutoOn = (CStr(mTestPpamAuto) <> "0"): Exit Function
    Updater_PpamAutoOn = (Settings_Get(KEY_PPAM_AUTO, "1") <> "0")
End Function

' (5) Instalado = o instalador deixou <raiz>\ThinQQ.ppam. O atualizador NUNCA instala o suplemento para quem nao o tem.
Public Function Updater_PpamInstalled() As Boolean
    On Error Resume Next
    Updater_PpamInstalled = (Len(Dir(PpamPath())) > 0)
End Function

' (1) Ha suplemento mais novo para este usuario: opcao ligada, instalado, manifesto com versao/url/SHA validos, versao maior que a
' instalada (PpamVersion; ausente = 0.0.0) e diferente da que ja esta pendente.
Public Function Updater_PpamNeeded(json As String) As Boolean
    Dim v As String
    If Not Updater_PpamAutoOn() Then Exit Function
    If Not Updater_PpamInstalled() Then Exit Function
    v = JsonGetString(json, "pptAddinVersion")
    If Not IsValidVersion(v) Then Exit Function
    If Len(JsonGetString(json, "pptAddinUrl")) = 0 Or Len(Trim$(JsonGetString(json, "pptAddinSha256"))) <> 64 Then Exit Function
    If Settings_Get(KEY_PENDING_PPAM) = "1" And Settings_Get(KEY_PENDING_PPAM_VER) = v Then Exit Function
    Updater_PpamNeeded = (CompareVersions(v, Settings_Get(KEY_PPAM_VERSION, "0.0.0")) > 0)
End Function

' (7) Baixa para tmp e confere o SHA-256 do manifesto (obrigatorio: AllowUnverified nao vale para o suplemento). False = nada fica em tmp.
Private Function Updater_PpamDownload(json As String, tmp As String) As Boolean
    On Error GoTo Fail
    If Len(Dir(tmp)) > 0 Then Kill tmp
    Transport_DownloadFile JsonGetString(json, "pptAddinUrl"), tmp
    If LCase$(Transport_Sha256File(tmp)) <> LCase$(Trim$(JsonGetString(json, "pptAddinSha256"))) Then
        Kill tmp
        Loader_Log "Updater_PpamDownload: SHA-256 nao confere; suplemento nao atualizado"
        Exit Function
    End If
    Updater_PpamDownload = True
    Exit Function
Fail:
    Loader_Log "Updater_PpamDownload: " & Err.Number & " " & Err.Description
    On Error Resume Next
    If Len(Dir(tmp)) > 0 Then Kill tmp
End Function

' (2) Troca atomica (mesma pasta: renomeacoes): ThinQQ.ppam -> .old, src -> ThinQQ.ppam, apaga .old. Arquivo em uso (a renomeacao
' falha: medido com o PowerPoint) -> False sem mudar nada. Falha no 2o passo -> o .old volta.
Private Function Updater_PpamSwap(src As String) As Boolean
    Dim dest As String, old As String
    dest = PpamPath(): old = dest & ".old"
    On Error Resume Next
    If Len(Dir(old)) > 0 Then Kill old
    Err.Clear
    Name dest As old
    If Err.Number <> 0 Then Err.Clear: Exit Function            ' em uso
    Name src As dest
    If Err.Number <> 0 Then
        Err.Clear
        Name old As dest
        Exit Function
    End If
    Kill old
    Err.Clear
    Updater_PpamSwap = True
End Function

' (2) Depois do download conferido: troca agora ou deixa pendente (.new + chaves) e agenda nova tentativa. Devolve "replaced",
' "replaced-restart" (trocado com o PowerPoint aberto: vale ao reabrir) ou "staged".
Private Function Updater_PpamCommit(tmp As String, sha As String, ver As String) As String
    Dim nw As String
    If Updater_PpamSwap(tmp) Then
        Settings_Set KEY_PPAM_VERSION, ver
        Settings_Delete KEY_PENDING_PPAM: Settings_Delete KEY_PENDING_PPAM_SHA: Settings_Delete KEY_PENDING_PPAM_VER
        If Len(Dir(PpamPath() & ".new")) > 0 Then Kill PpamPath() & ".new"
        If Updater_PptRunning() Then Updater_PpamCommit = "replaced-restart" Else Updater_PpamCommit = "replaced"
        Exit Function
    End If
    nw = PpamPath() & ".new"
    If Len(Dir(nw)) > 0 Then Kill nw
    Name tmp As nw
    Settings_Set KEY_PENDING_PPAM, "1": Settings_Set KEY_PENDING_PPAM_SHA, LCase$(sha): Settings_Set KEY_PENDING_PPAM_VER, ver
    Updater_SchedulePpamRetry
    Updater_PpamCommit = "staged"
End Function

' So o suplemento (core em dia, ou depois da troca do core): baixa, confere e aplica. "" = nada feito (nao precisava ou falhou: registrado).
Public Function Updater_PpamUpdateNow(json As String) As String
    Dim tmp As String
    On Error GoTo Fail
    If Not Updater_PpamNeeded(json) Then Exit Function
    tmp = Loader_TmpDir() & "\ThinQQ.ppam"
    If Not Updater_PpamDownload(json, tmp) Then Exit Function
    Updater_PpamUpdateNow = Updater_PpamCommit(tmp, JsonGetString(json, "pptAddinSha256"), JsonGetString(json, "pptAddinVersion"))
    Exit Function
Fail:
    Loader_Log "Updater_PpamUpdateNow: " & Err.Number & " " & Err.Description
End Function

' (4) Frase para o usuario (acrescentada ao aviso de atualizacao).
Public Function Updater_PpamStatusText(state As String) As String
    Select Case state
        Case "replaced": Updater_PpamStatusText = "Suplemento do PowerPoint atualizado."
        Case "replaced-restart": Updater_PpamStatusText = "Suplemento do PowerPoint atualizado: feche e abra o PowerPoint para usar a versão nova."
        Case "staged": Updater_PpamStatusText = "O suplemento do PowerPoint será atualizado automaticamente quando o PowerPoint for fechado (o Excel tenta de novo a cada 10 minutos e na próxima abertura)."
        Case "failed": Updater_PpamStatusText = "O suplemento do PowerPoint não foi atualizado agora; nova tentativa na próxima verificação."
    End Select
End Function

' (2) Aplicacao da pendencia: na abertura do Excel (Loader_Startup) e pela nova tentativa agendada. Reconfere o SHA-256 do .new (arquivo em
' pasta do usuario: nunca confia no que esta la); adulterado ou ausente = descarta. Ainda em uso = continua pendente (e reagenda). Nunca propaga.
Public Sub Updater_ApplyPendingPpam()
    Dim nw As String
    On Error GoTo Fail
    If Settings_Get(KEY_PENDING_PPAM) <> "1" Then Exit Sub
    nw = PpamPath() & ".new"
    If Len(Dir(nw)) = 0 Or LCase$(Transport_Sha256File(nw)) <> Settings_Get(KEY_PENDING_PPAM_SHA) Then
        Loader_Log "Updater_ApplyPendingPpam: .new ausente ou com SHA-256 diferente; descartado"
        If Len(Dir(nw)) > 0 Then Kill nw
        Settings_Delete KEY_PENDING_PPAM: Settings_Delete KEY_PENDING_PPAM_SHA: Settings_Delete KEY_PENDING_PPAM_VER
        Exit Sub
    End If
    If Updater_PpamSwap(nw) Then
        Settings_Set KEY_PPAM_VERSION, Settings_Get(KEY_PENDING_PPAM_VER)
        Settings_Delete KEY_PENDING_PPAM: Settings_Delete KEY_PENDING_PPAM_SHA: Settings_Delete KEY_PENDING_PPAM_VER
        Loader_Log "Suplemento do PowerPoint pendente aplicado"
    Else
        Updater_SchedulePpamRetry
    End If
    Exit Sub
Fail:
    Loader_Log "Updater_ApplyPendingPpam: " & Err.Number & " " & Err.Description
End Sub

' Nova tentativa em 10 min (so enquanto houver pendencia; o Excel fechado leva a tentativa para a proxima abertura).
Private Sub Updater_SchedulePpamRetry()
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then Updater_PpamRetryScheduleCount = Updater_PpamRetryScheduleCount + 1: Exit Sub
    Application.OnTime Now + TimeSerial(0, PPAM_RETRY_MINUTES, 0), "'" & ThisWorkbook.Name & "'!Updater_ApplyPendingPpam"
End Sub
```
Integração, ainda em `modUpdater.bas`:
  1. `Updater_Apply`: declare `Dim ppamTmp As String, ppamOk As Boolean, ppamState As String, wantPpam As Boolean`; **antes** de `Application.StatusBar = "ThinQQ: baixando atualização..."`, `mLastPpamNote = ""`; logo depois dela: `wantPpam = doCore And Updater_PpamNeeded(json)` e `If wantPpam Then ppamTmp = Loader_TmpDir() & "\ThinQQ.ppam": ppamOk = Updater_PpamDownload(json, ppamTmp)` (o `.ppam` é baixado e conferido ANTES de trocar o core); depois de `committed = True` (core trocado e carregado): `If wantPpam Then` → `If ppamOk Then ppamState = Updater_PpamCommit(ppamTmp, JsonGetString(json, "pptAddinSha256"), JsonGetString(json, "pptAddinVersion")) Else ppamState = "failed"` → `mLastPpamNote = Updater_PpamStatusText(ppamState)` (envolva o `Commit` em `On Error Resume Next` com registro: erro aqui vira `"failed"`, nunca desfaz o core). No `Fail:` (core não trocado), apague `ppamTmp` se existir (o suplemento só é trocado com o core já trocado: o par nunca fica "suplemento novo + core velho"). A mensagem final de sucesso passa a `"ThinQQ atualizado com sucesso." & ... & IIf(Len(mLastPpamNote) > 0, vbLf & mLastPpamNote, "")`.
  2. `Updater_Check`: depois de `what = Updater_Evaluate(...)`: se `what = "none"`, então `If Updater_PpamNeeded(json) Then` → `mLastPpamNote = Updater_PpamStatusText(Updater_PpamUpdateNow(json))`, `Loader_Log mLastPpamNote` e, se `Not silent`, `Loader_Msg mLastPpamNote, vbInformation` (só o suplemento: sem formulário, porque o core da mesma versão já foi aceito); senão a mensagem "já está na versão mais recente" como hoje.
  3. `Updater_VersionLine`: acrescente ao fim do texto, quando `Updater_PpamNeeded(json)`, `"   |   Suplemento do PowerPoint: " & Settings_Get(KEY_PPAM_VERSION, "0.0.0") & " > " & JsonGetString(json, "pptAddinVersion")`.
  4. Em `src/loader/modules/modLoaderMain.bas`, `Loader_Startup`, logo depois de `Loader_EnsureFolders`: `Updater_ApplyPendingPpam` (com o PowerPoint fechado, a primeira abertura do Excel aplica).
- [ ] **Passo 6: Controle do usuário (4)** — `frmUpdate.json`: acrescente `{ "type": "CheckBox", "name": "chkPpam", "Left": 16, "Top": 228, "Width": 396, "Height": 16, "Caption": "Atualizar também o suplemento do PowerPoint (feche o PowerPoint para aplicar na hora)", "fontName": "Segoe UI", "fontSize": 9 }` e reduza `txtNotes` para `"Height": 140`; `frmUpdate.vba`: `Prepare` ganha `Optional showPpam As Boolean = False` → `chkPpam.Visible = showPpam: chkPpam.Value = (Settings_Get(KEY_PPAM_AUTO, "1") <> "0")`; em `btnUpdate_Click`, antes de `Me.Hide`: `If chkPpam.Visible Then Settings_Set KEY_PPAM_AUTO, IIf(chkPpam.Value, "1", "0")`. `ShowUpdateForm` passa `showPpam:=Updater_PpamInstalled()` (só para quem tem o suplemento). `Rx_OnSettings` (modRibbon), depois de gravar a fonte e só se `Updater_PpamInstalled()`: `Loader_Msg("Atualizar também o suplemento do PowerPoint automaticamente?", vbYesNo + vbQuestion)` → grava `KEY_PPAM_AUTO` "1"/"0" (no build Test o `Rx_OnSettings` já sai antes).
- [ ] **Passo 7: GREEN** — `-Target loader -Filter "TestUpdater"` → 0 falhas; suíte completa (o `Test-LoaderSwap` e as suítes de release continuam verdes).
- [ ] **Passo 8: Commit** — `git commit -m "feat: atualizador do Excel atualiza o suplemento do PowerPoint (troca atomica; pendente aplicado quando o PowerPoint libera o arquivo)" (+ trailer)`.

---

### Tarefa A8: Documentação da Parte A — modelo: standard

**Files:**
- Modify: `docs/tutorial/tutorial.html` (seção `#powerpoint`, mensagens, limites), `README.md` (Recursos, Instalação, Verificações manuais, Limitações), `CHANGELOG.md` (`[Não lançado]`), `release/notes.txt`, `tests/ps/Test-Docs.ps1`

**Interfaces:**
- Consumes: rótulos exatos das faixas (A3: "Atualizar gráficos no PowerPoint"; A4: "Atualizar gráficos", "Ao abrir", "Perguntar", "Atualizar sempre", "Não atualizar", "Sobre", aba "ThinQQ", grupo "Gráficos do Excel"); textos de §4.2/§4.3 do spec; textos da atualização automática do suplemento (A7: `Updater_PpamStatusText`, caixa "Atualizar também o suplemento do PowerPoint").

- [ ] **Passo 1: Test-Docs (RED)** — (a) carregue também `src/ppt/ribbon/customUI14.xml` na lista de rótulos conhecidos (onde ele lê `$ribbon.SelectNodes('//*[@label]')`, repita para o XML do PowerPoint), mas **não** na regra "todo rótulo da faixa aparece no tutorial" para os itens de lista; (b) acrescente ao `$messages` (mensagens que o tutorial lista) as frases: as 4 de `Updater_PpamStatusText`, `PPTSYNC_NONE`, "Pasta de origem não encontrada ou sem acesso (o gráfico ficou como estava)", "O ThinQQ não está instalado no Excel deste computador: só os dados foram atualizados", "O Excel está ocupado (uma célula em edição ou uma caixa de diálogo aberta)"; (c) acrescente ao mapa de limites (`mustDocument`): `'powerpoint: sem o suplemento so os dados atualizam' = @('both', 'sem o suplemento do PowerPoint')`, `'powerpoint: grafico dentro de grupo fica de fora' = @('tutorial', 'dentro de um grupo')`, `'powerpoint: Modo de Exibicao Protegido' = @('both', 'Habilitar Edição')`, `'powerpoint: pasta movida' = @('both', 'Alterar Origem')`.
- [ ] **Passo 2: Tutorial** — reescreva a seção `id="powerpoint"` (mantendo o `id` e o `h-powerpoint`) com: o que é colado (imagem x gráfico vinculado); **as setas no tamanho do slide**; o suplemento do PowerPoint (aba <span class="ui">ThinQQ</span>, grupo <span class="ui">Gráficos do Excel</span>, <span class="ui">Atualizar gráficos</span>, <span class="ui">Ao abrir</span> com <span class="ui">Perguntar</span>/<span class="ui">Atualizar sempre</span>/<span class="ui">Não atualizar</span>, <span class="ui">Sobre</span>); o botão do Excel <span class="ui">Atualizar gráficos no PowerPoint</span>; qual Excel é usado (o aberto com o ThinQQ, senão um invisível só para isso) e que a pasta aberta é lida com as alterações não salvas; a tabela de mensagens (§4.2 e §4.3, textos exatos); os limites: sem o suplemento do PowerPoint (ou sem o ThinQQ no Excel) só os dados atualizam; Modo de Exibição Protegido: clique em "Habilitar Edição" e a pergunta aparece; pasta movida/renomeada: "Alterar Origem" no PowerPoint; gráfico dentro de um grupo no slide fica de fora; formatação feita no PowerPoint fica, mas mudar a legenda ou o tamanho da área do gráfico no PowerPoint pode desalinhar as setas (refaça "Copiar como gráfico vinculado"); a instalação manual do suplemento (LEIA-ME). Atualize a nota da seção "Instalação" (três arquivos: `ThinQQ.xlam`, `ThinQQ.core.xlam`, `ThinQQ.ppam`) e a seção de atualização: o suplemento do PowerPoint é atualizado **automaticamente** pelo Excel junto com o ThinQQ (só para quem o instalou pelo ZIP; opção "Atualizar também o suplemento do PowerPoint" no aviso de atualização e em ThinQQ > Configurações); com o PowerPoint aberto a troca fica para quando ele fechar (o Excel tenta a cada 10 minutos e na próxima abertura) e, se o arquivo estava livre, basta fechar e abrir o PowerPoint; um core mais antigo que o suplemento faz o suplemento atualizar só os dados.
- [ ] **Passo 3: README/CHANGELOG/notes** — README: em Recursos (linha do menu PowerPoint: acrescente "Atualizar gráficos no PowerPoint" e o suplemento do PowerPoint), Instalação (o instalador registra o suplemento do PowerPoint; feche o Excel e o PowerPoint), Verificações manuais (itens novos numerados no fim: abrir pela interface nos 3 modos; anexo em Modo de Exibição Protegido + Habilitar Edição; pasta movida; pasta em rede; 30 gráficos; Excel aberto com a pasta alterada; PowerPoint sem o ThinQQ no Excel; instalação pelo ZIP mostra a aba ThinQQ no PowerPoint; atualização pelo Excel com o PowerPoint fechado e aberto), Limitações. CHANGELOG `[Não lançado]`: `### Adicionado` "Suplemento do PowerPoint (ThinQQ.ppam): ao abrir uma apresentação, atualiza dados e setas dos gráficos do ThinQQ colados como gráfico vinculado (perguntar, sempre ou nunca)." e "Excel: PowerPoint > Atualizar gráficos no PowerPoint."; "Atualização: o Excel também atualiza o suplemento do PowerPoint (com o PowerPoint aberto, quando ele fechar)."; `### Corrigido` "Copiar como gráfico vinculado: as setas e rótulos de seta não ficam mais esticados e fora do lugar quando o PowerPoint ajusta o gráfico ao slide.". `release/notes.txt`: acrescente "• Novo: o PowerPoint atualiza os gráficos do ThinQQ ao abrir a apresentação (dados e setas); reinstale pelo ZIP para ganhar o suplemento do PowerPoint." (limite de 1 a 8 linhas; mantenha as existentes que ainda valem conforme a regra do controlador).
- [ ] **Passo 4: GREEN** — `pwsh <repo>\tests\ps\Test-Docs.ps1 -RepoRoot <repo>` e `powershell ... Test-Docs.ps1` → 0 falhas; `Test-Privacy` verde; suíte completa.
- [ ] **Passo 5: Commit** — `git commit -m "docs: tutorial, README e CHANGELOG do PowerPoint atualizado ao reabrir" (+ trailer)`.

---

### Tarefa A9: Suíte final, revisão do ramo e verificação manual (humano) — modelo: standard (controlador + humano)

- [ ] **Passo 1:** sem PowerPoint aberto, `pwsh C:\temp\tqq-pd\build\run-tests.ps1 -TimeoutSec 900 -Ppt -LoaderSwap` → "Todos os testes passaram" (se cair por memória: core em 3 partes com `-Filter` e o resto sem o core, relatando).
- [ ] **Passo 2:** revisão do ramo inteiro (modelo mais capaz) + **uma** onda de correção + revisão de escopo, como na cascata comparativa.
- [ ] **Passo 3 (humano):** `build.ps1 -Target all -Configuration Release` + `-Target ppt -Configuration Release` numa cópia; instalar pelo ZIP local (ou `release.ps1` com `-PublishTo` numa pasta) **numa conta de teste** ou na própria máquina; percorrer a lista de verificações manuais nova do README (spec §4.7): colar vinculado com setas (alinhadas), fechar tudo, mudar o dado, reabrir o .pptx pelo Explorer nos modos Perguntar/Sempre/Nunca; anexo baixado (Modo de Exibição Protegido) + Habilitar Edição; pasta renomeada (aviso, sem travar); pasta em rede; 30 gráficos (tempo); Excel aberto com a pasta alterada e não salva; sem o ThinQQ no Excel (só dados); PowerPoint de 64 bits se houver. **Atualização automática do suplemento** (com um manifesto local de teste numa pasta, `ThinQQ > Configurações` apontando para ele, versão maior): (a) PowerPoint fechado → "Verificar atualizações" troca o `ThinQQ.ppam` e o PowerPoint seguinte mostra "Sobre" com a versão nova; (b) PowerPoint **aberto com o suplemento carregado** → mede se o arquivo fica travado (a spec §3.4 só mediu apresentação aberta): travado = aviso "quando o PowerPoint for fechado" e `ThinQQ.ppam.new`; fechar o PowerPoint e esperar a nova tentativa (≤ 10 min) ou reabrir o Excel aplica; livre = aviso "feche e abra o PowerPoint"; (c) desmarcar "Atualizar também o suplemento do PowerPoint" → não troca; (d) conta sem o suplemento instalado → o atualizador não cria o `.ppam`. Restaure a fonte de atualização no fim.
- [ ] **Passo 4:** PR, merge, `publicar-release.bat 1.4.0` (2 etapas).

---

# PARTE B — Delta entre barras inserido clicando nas barras (1.5.0)

### Tarefa B0: Recurso `delta` na matriz (21 × 24) — modelo: standard

**Files:**
- Modify: `src/core/modules/modChartKinds.bas` (`Kinds_Supports`, `Kinds_FeatureLabel`), `src/core/modules/modAnnotations.bas` (`Ann_RefusalText`, constante da recusa)
- Modify: `docs/tutorial/tutorial.html` (`#matriz`: 5.ª tabela), `tests/ps/Test-Docs.ps1` (24 recursos), `tests/ps/Test-DemoWorkbook.ps1` (pendência temporária, se ele exigir demonstração de todo recurso)
- Create (só se a Parte A ainda não rodou): `docs/superpowers/specs/2026-10-14-thinqq-ppt-delta-design.md` (Passo 1 da Tarefa A0)
- Test: `tests/vba/core/TestKinds.bas`

**Interfaces:**
- Produces: `Kinds_Supports(k, "delta")` para `stacked, clustered, line, waterfall, barstacked, barclustered, barwaterfall, combo, waterfallpair`; `Kinds_FeatureLabel("delta") = "Delta clicando nas barras"`; `Public Const ANN_DELTA_REFUSAL As String = "O delta entre barras não está disponível para este tipo de gráfico."`; `Ann_RefusalText(cht, "delta")`.

- [ ] **Passo 1: Spec** — se `docs/superpowers/specs/2026-10-14-thinqq-ppt-delta-design.md` não existir, faça o Passo 1 da Tarefa A0 (copiar spec e plano).
- [ ] **Passo 2: Testes (RED)** — ao fim de `tests/vba/core/TestKinds.bas` (usa `SpecSectionOf` e `SpecKinds` que o módulo já tem):

```vb
Public Sub Test_Kinds_Delta_LabelAndRefusal()
    Dim wb As Workbook, ws As Worksheet, c As Chart
    AssertEqual "Delta clicando nas barras", Kinds_FeatureLabel("delta")
    AssertTrue Kinds_Supports("clustered", "delta"): AssertTrue Kinds_Supports("waterfallpair", "delta")
    AssertFalse Kinds_Supports("stacked100", "delta"): AssertFalse Kinds_Supports("area", "delta"): AssertFalse Kinds_Supports("butterfly", "delta")
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = "Cat": ws.Range("B1").Value = "Valor": ws.Range("A2").Value = "a": ws.Range("B2").Value = 1: ws.Range("A3").Value = "b": ws.Range("B3").Value = 2
    ws.Range("A1:B3").Select
    Core_Insert "pie"
    Set c = Core_ActiveChart()
    AssertEqual ANN_DELTA_REFUSAL, Ann_RefusalText(c, "delta")
    Watchers_Stop
    CloseTemp wb
End Sub

' Parte B (1.5.0): a matriz inteira (21 tipos x 24 recursos) conferida contra o TEXTO da tabela da secao 6 do spec 2026-10-14.
Public Sub Test_Kinds_Delta_Matrix_ExhaustiveFromSpec()
    Dim sec As String, lines() As String, i As Long, cells() As String, feat As String, expect As String, k As Variant, allKinds As Variant
    Dim nFeat As Long, nPairs As Long, want As Boolean
    allKinds = Array("stacked", "clustered", "stacked100", "line", "waterfall", "barstacked", "barclustered", "barstacked100", "barwaterfall", "area", "area100", "combo", "mekko", "mekkounits", "gantt", "pie", "doughnut", "butterfly", "scatter", "bubble", "waterfallpair")
    sec = SpecSectionOf("2026-10-14-thinqq-ppt-delta-design.md", "## 6. ", "## 7. ")
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
    AssertEqual 24, nFeat, "24 recursos na tabela do spec"
    AssertEqual 504, nPairs
End Sub
```
- [ ] **Passo 3: RED** — `-Target core -Filter "TestKinds" -TimeoutSec 300` → falhas (recurso desconhecido) ou saída 3 (`ANN_DELTA_REFUSAL`).
- [ ] **Passo 4: Implementar** — `modChartKinds.Kinds_Supports`, antes do `Case Else` (uma linha, lida pelo Test-Docs por expressão regular):

```vb
        Case "delta":      list = "stacked,clustered,line,waterfall,barstacked,barclustered,barwaterfall,combo,waterfallpair"
```
`Kinds_FeatureLabel`, antes do `End Select`:
```vb
        Case "delta":      Kinds_FeatureLabel = "Delta clicando nas barras"
```
`modAnnotations`: depois de `Public Const ANN_BRACKET_REFUSAL ...`:
```vb
Public Const ANN_DELTA_REFUSAL As String = "O delta entre barras não está disponível para este tipo de gráfico."
```
e em `Ann_RefusalText`, depois do `Case "bracket"`:
```vb
        Case "delta"
            Ann_RefusalText = ANN_DELTA_REFUSAL
```
- [ ] **Passo 5: Tutorial e Test-Docs** — `Test-Docs.ps1`: `23 recursos` → `24 recursos` (linha ~201: `$matrix.Count -eq 24`). Tutorial `#matriz`: depois do `</div>` que fecha a tabela "Cascata comparativa", acrescente `<h4>Delta entre barras</h4>` e uma 5.ª tabela `class="mx"` no mesmo formato (mesmo `div class="tw"` com `aria-label="Tabela do que cada tipo aceita: delta entre barras (em telas estreitas, role para os lados)"`), cabeçalho `<th scope="col">Tipo</th><th scope="col" data-f="delta"><span class="k">Delta clicando nas barras</span></th>` e as **mesmas 21 linhas** `data-kind` da 4.ª tabela (copie os `<th scope="row">`), cada uma com uma célula: `<td class="y">Sim</td>` para stacked, clustered, line, waterfall, barstacked, barclustered, barwaterfall, combo e waterfallpair, `<td>—</td>` nas outras. (`span class="k"` vira `ui` na B4, quando o botão existir.) Se `Test-DemoWorkbook.ps1` exigir que todo recurso da matriz apareça na demonstração, acrescente `'delta'` a uma lista de pendências temporária no formato da `$pendingDemo` da cascata comparativa (a B5 remove).
- [ ] **Passo 6: GREEN + suíte completa** e **Commit** — `git commit -m "feat: recurso delta na matriz de capacidades (21 x 24)" (+ trailer)`.

---

### Tarefa B1: `modDelta` — referências, extensões, textos — modelo: most-capable

**Files:**
- Create: `src/core/modules/modDelta.bas`
- Modify: `src/core/modules/modChartWaterfall.bas` (constantes `WF_COL_LO = 7`, `WF_COL_HI = 8`)
- Test: `tests/vba/core/TestDelta.bas` (novo)

**Interfaces:**
- Consumes (B0): `Kinds_Supports(..., "delta")`. Do core: `Style_IsHelperSeries`, `Geom_*`, `Registry_*`, `Block_Body`, `Wfp_SideCol`, `Wfp_SlotOf`, `Wfp_StepCount`, `Util_ToDouble`.
- Produces: `DELTA_GAP = 8`, `DELTA_LABEL_ROOM = 18`, `DELTA_BAD_PICK = "Escolha duas barras diferentes do gráfico."`; `Delta_Text(a, b, mode, fmt)`, `Delta_NextFormat(fmt)`, `Delta_ParseRef(ref, k, c)`, `Delta_SeriesIndex(cht, k)`, `Delta_DataOrdinal(cht, s)`, `Delta_DataSeriesCount(cht)`, `Delta_IsColumnSeries(s)`, `Delta_ResolvePick(cht, s, p, ref)`, `Delta_Extent(cht, ref, v, lo, hi)`, `Delta_AxisIndex(cht, ref)`, `Delta_SpanEdge(cht, ia, ib, above)`, `Delta_CatCenter(cht, ref)`, `Delta_BarHalf(cht, ref)`.

- [ ] **Passo 1: Testes (RED)** — crie `tests/vba/core/TestDelta.bas` (a B2 acrescenta a geometria no mesmo módulo):

```vb
Attribute VB_Name = "TestDelta"
Option Explicit
' Parte B (1.5.0): delta entre barras escolhidas com o mouse. Referencias "k:c" (k = ordinal da serie de DADOS; 0 = barra da cascata; na
' comparativa k = lado e c = etapa), extensoes, textos e (B2) geometria independente das formulas de producao.

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

' Sul 10, 20, 30; Norte 5, 6, -4 (categorias T1..T3).
Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Tri", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5)
    ws.Range("A3:C3").Value = Array("T2", 20, 6)
    ws.Range("A4:C4").Value = Array("T3", 30, -4)
    Set Fixture = ws.Range("A1:C4")
End Function

Private Function WaterfallFixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Receita", 100)
    ws.Range("A3:B3").Value = Array("Custo", -30)
    ws.Range("A4:B4").Value = Array("Lucro", "e")
    Set WaterfallFixture = ws.Range("A1:B4")
End Function

Private Function PairFixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Par"
    ws.Range("A1:C1").Value = Array("Etapa", "Seg 1", "Seg 2")
    ws.Range("A2:C2").Value = Array("Bruto", 27, 29)
    ws.Range("A3:C3").Value = Array("Comp 1", -4, -7)
    ws.Range("A4:C4").Value = Array("Líquido", "e", "e")
    Set PairFixture = ws.Range("A1:C4")
End Function

Private Function InsertOn(rng As Range, kind As String) As Chart
    rng.Worksheet.Activate
    rng.Select
    Core_Insert kind
    Set InsertOn = Core_ActiveChart()
End Function

Private Function FirstSeries(c As Chart, wantHelper As Boolean, wantColumn As Boolean) As Long
    Dim i As Long
    For i = 1 To c.SeriesCollection.Count
        If Style_IsHelperSeries(c.SeriesCollection(i)) = wantHelper And Delta_IsColumnSeries(c.SeriesCollection(i)) = wantColumn Then FirstSeries = i: Exit Function
    Next
End Function

Public Sub Test_Delta_Text()
    Dim f As String
    f = "0"
    AssertEqual "+5 (+50%)", Delta_Text(10, 15, "both", f)
    AssertEqual "+5", Delta_Text(10, 15, "abs", f)
    AssertEqual "+50%", Delta_Text(10, 15, "pct", f)
    AssertEqual "-14 (-140%)", Delta_Text(10, -4, "both", f)
    AssertEqual "+5", Delta_Text(0, 5, "both", f), "origem zero: so o valor (como a Seta diferenca)"
    AssertEqual "n/d", Delta_Text(0, 5, "pct", f)
    AssertEqual "-75%", Delta_Text(-4, -7, "pct", f), "(B - A) / |A|, como a Seta diferenca (a seta de variacao da comparativa daria +75%)"
    f = Style_NumberFormat("int")
    AssertEqual Format$(1500, "+" & Split(f & ";", ";")(0) & ";-" & Split(f & ";", ";")(0) & ";0") & " (+150%)", Delta_Text(1000, 2500, "both", f), "formato do grafico com a localidade"
    AssertEqual "abs", Delta_NextFormat("both"): AssertEqual "pct", Delta_NextFormat("abs"): AssertEqual "both", Delta_NextFormat("pct"): AssertEqual "abs", Delta_NextFormat("")
End Sub

Public Sub Test_Delta_ParseRef()
    Dim k As Long, c As Long
    AssertTrue Delta_ParseRef("2:3", k, c): AssertEqual 2, k: AssertEqual 3, c
    AssertTrue Delta_ParseRef("0:1", k, c), "barra da cascata"
    AssertFalse Delta_ParseRef("2", k, c): AssertFalse Delta_ParseRef("a:1", k, c): AssertFalse Delta_ParseRef("1:0", k, c): AssertFalse Delta_ParseRef("-1:2", k, c)
End Sub

Public Sub Test_Delta_ResolvePick_PerFamily()
    Dim wb As Workbook, c As Chart, ref As String, s As Long
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    AssertTrue Delta_ResolvePick(c, Delta_SeriesIndex(c, 2), 3, ref): AssertEqual "2:3", ref
    AssertFalse Delta_ResolvePick(c, Delta_SeriesIndex(c, 1), 4, ref), "categoria inexistente"
    s = FirstSeries(c, True, False)
    If s > 0 Then AssertFalse Delta_ResolvePick(c, s, 1, ref), "auxiliar tqq_* nao e barra"
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(WaterfallFixture(wb), "waterfall")
    AssertTrue Delta_ResolvePick(c, FirstSeries(c, True, True), 2, ref): AssertEqual "0:2", ref, "cascata: a barra da categoria"
    AssertFalse Delta_ResolvePick(c, FirstSeries(c, True, True), 4, ref)
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(PairFixture(wb), "waterfallpair")
    s = FirstSeries(c, True, True)
    If s = 0 Then s = FirstSeries(c, False, True)
    AssertTrue Delta_ResolvePick(c, s, 1, ref): AssertEqual "1:1", ref, "fatia A da etapa 1"
    AssertTrue Delta_ResolvePick(c, s, 2, ref): AssertEqual "2:1", ref, "fatia B da etapa 1"
    AssertFalse Delta_ResolvePick(c, s, 3, ref), "vao entre etapas"
    AssertTrue Delta_ResolvePick(c, s, 4, ref): AssertEqual "1:2", ref
    AssertFalse Delta_ResolvePick(c, Delta_FindSeriesByName(c, "tqq_pd"), 1, ref), "seta de variacao (XY) nao e barra"
    Teardown wb
End Sub

Public Sub Test_Delta_Extent_PerFamily()
    Dim wb As Workbook, c As Chart, v As Double, lo As Double, hi As Double
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "stacked")
    AssertTrue Delta_Extent(c, "2:1", v, lo, hi): AssertEqual 5, v: AssertEqual 10, lo: AssertEqual 15, hi
    AssertTrue Delta_Extent(c, "2:3", v, lo, hi): AssertEqual -4, v: AssertEqual -4, lo: AssertEqual 0, hi, "segmento negativo desce do zero"
    AssertTrue Delta_Extent(c, "1:3", v, lo, hi): AssertEqual 30, v: AssertEqual 0, lo: AssertEqual 30, hi
    AssertFalse Delta_Extent(c, "3:1", v, lo, hi), "serie inexistente"
    AssertEqual 26, Delta_SpanEdge(c, 1, 2, True), "maior pilha positiva de T1..T2 (20 + 6)"
    AssertEqual -4, Delta_SpanEdge(c, 3, 3, False)
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    AssertTrue Delta_Extent(c, "2:3", v, lo, hi): AssertEqual -4, lo: AssertEqual 0, hi
    AssertEqual 30, Delta_SpanEdge(c, 1, 3, True)
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "line")
    AssertTrue Delta_Extent(c, "2:3", v, lo, hi): AssertEqual -4, lo: AssertEqual -4, hi, "linha: o proprio ponto"
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(WaterfallFixture(wb), "waterfall")
    AssertTrue Delta_Extent(c, "0:2", v, lo, hi): AssertEqual 70, v, "valor acumulado (End)": AssertEqual 70, lo: AssertEqual 100, hi
    AssertTrue Delta_Extent(c, "0:3", v, lo, hi): AssertEqual 70, v: AssertEqual 0, lo, "total vai do zero"
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(PairFixture(wb), "waterfallpair")
    AssertTrue Delta_Extent(c, "2:2", v, lo, hi): AssertEqual -7, v, "Val do lado B na etapa 2": AssertEqual 22, lo: AssertEqual 29, hi
    AssertEqual 5, Delta_AxisIndex(c, "2:2"), "fatia B da etapa 2 = 3*(2-1)+2"
    AssertFalse Delta_Extent(c, "3:1", v, lo, hi)
    Teardown wb
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "TestDelta"` → saída 3.

- [ ] **Passo 3: Implementar** — em `src/core/modules/modChartWaterfall.bas`, depois de `Public Const WF_COL_END As Long = 6`:

```vb
Public Const WF_COL_LO As Long = 7                  ' extensao da barra (delta entre barras, Parte B)
Public Const WF_COL_HI As Long = 8
```
Crie `src/core/modules/modDelta.bas`:

```vb
Attribute VB_Name = "modDelta"
Option Explicit
' Delta entre duas barras escolhidas com o mouse (Parte B, spec 2026-10-14 secao 5). Referencia "k:c": k = ordinal da serie de DADOS (as
' auxiliares tqq_* nao contam) e c = categoria; nas cascatas k = 0 (a barra da categoria); na cascata comparativa k = lado (1 = A, 2 = B) e
' c = etapa. Valor (texto) e extensao (lo..hi, onde a seta encosta) por familia: agrupadas e linha = o ponto; empilhadas e combinacao = o
' segmento (positivos sobem e negativos descem na ordem das series); cascata = End e Lo/Hi do bloco; comparativa = Val e Lo/Hi do lado.
Public Const DELTA_GAP As Double = 8            ' pt: da borda mais externa das barras do trecho ate o rotulo
Public Const DELTA_LABEL_ROOM As Double = 18    ' pt a mais com Rotulos ou Totais ligados (como ANN_TOTALS_EXTRA)
Public Const DELTA_BAD_PICK As String = "Escolha duas barras diferentes do gráfico."

' Texto do delta de a para b: "both" = "+5 (+50%)" (como a Seta diferenca), "abs" = "+5", "pct" = "+50%". Percentual (b - a) / |a|; a = 0:
' "both" mostra so o valor e "pct" mostra "n/d". fmt = codigo do formato do grafico (Style_NumberFormat; a 1a secao). Pura.
Public Function Delta_Text(a As Double, b As Double, mode As String, fmt As String) As String
    Dim d As Double, sec As String, absTxt As String, pctTxt As String
    d = b - a
    sec = Split(fmt & ";", ";")(0)
    absTxt = Format$(d, "+" & sec & ";-" & sec & ";0")
    If a <> 0 Then pctTxt = Format$(d / Abs(a), "+0%;-0%;0%") Else pctTxt = "n/d"
    Select Case mode
        Case "abs": Delta_Text = absTxt
        Case "pct": Delta_Text = pctTxt
        Case Else
            Delta_Text = absTxt
            If a <> 0 Then Delta_Text = absTxt & " (" & pctTxt & ")"
    End Select
End Function

' Ciclo do botao "Delta: valor / % / ambos": ambos -> valor -> % -> ambos. Pura.
Public Function Delta_NextFormat(fmt As String) As String
    Select Case fmt
        Case "both": Delta_NextFormat = "abs"
        Case "abs": Delta_NextFormat = "pct"
        Case "pct": Delta_NextFormat = "both"
        Case Else: Delta_NextFormat = "abs"
    End Select
End Function

Public Function Delta_ParseRef(ref As String, ByRef k As Long, ByRef c As Long) As Boolean
    Dim p() As String
    p = Split(ref, ":")
    If UBound(p) <> 1 Then Exit Function
    If Not IsNumeric(p(0)) Or Not IsNumeric(p(1)) Then Exit Function
    If InStr(p(0), ",") > 0 Or InStr(p(1), ",") > 0 Or InStr(p(0), ".") > 0 Or InStr(p(1), ".") > 0 Then Exit Function
    k = CLng(p(0)): c = CLng(p(1))
    Delta_ParseRef = (k >= 0 And c >= 1)
End Function

' Series de colunas/barras (as de linha e XY nao sao barras).
Public Function Delta_IsColumnSeries(s As Series) As Boolean
    Select Case s.ChartType
        Case xlColumnClustered, xlColumnStacked, xlColumnStacked100, xlBarClustered, xlBarStacked, xlBarStacked100
            Delta_IsColumnSeries = True
    End Select
End Function

Public Function Delta_DataSeriesCount(cht As Chart) As Long
    Dim i As Long
    For i = 1 To cht.SeriesCollection.Count
        If Not Style_IsHelperSeries(cht.SeriesCollection(i)) Then Delta_DataSeriesCount = Delta_DataSeriesCount + 1
    Next
End Function

' Indice em SeriesCollection da k-esima serie de dados (0 se nao existe).
Public Function Delta_SeriesIndex(cht As Chart, k As Long) As Long
    Dim i As Long, n As Long
    For i = 1 To cht.SeriesCollection.Count
        If Not Style_IsHelperSeries(cht.SeriesCollection(i)) Then
            n = n + 1
            If n = k Then Delta_SeriesIndex = i: Exit Function
        End If
    Next
End Function

' Ordinal de dados da serie de indice s (0 = auxiliar ou inexistente).
Public Function Delta_DataOrdinal(cht As Chart, s As Long) As Long
    Dim i As Long, n As Long
    If s < 1 Or s > cht.SeriesCollection.Count Then Exit Function
    If Style_IsHelperSeries(cht.SeriesCollection(s)) Then Exit Function
    For i = 1 To s
        If Not Style_IsHelperSeries(cht.SeriesCollection(i)) Then n = n + 1
    Next
    Delta_DataOrdinal = n
End Function

' Indice da serie pelo nome (0 se nao existe). Series.Name devolve o texto.
Public Function Delta_FindSeriesByName(cht As Chart, nm As String) As Long
    Dim i As Long
    For i = 1 To cht.SeriesCollection.Count
        If cht.SeriesCollection(i).Name = nm Then Delta_FindSeriesByName = i: Exit Function
    Next
End Function

Private Function Delta_PointValue(s As Series, c As Long) As Double
    Dim vals As Variant
    vals = s.Values
    If Not IsArray(vals) Then Exit Function
    If LBound(vals) + c - 1 > UBound(vals) Then Exit Function
    Delta_PointValue = Util_ToDouble(vals(LBound(vals) + c - 1))
End Function

Private Function Delta_IsStackedKind(kind As String) As Boolean
    Delta_IsStackedKind = (kind = "stacked" Or kind = "barstacked" Or kind = "combo")
End Function

' Clique do Excel (Arg1 = indice da serie, Arg2 = ponto) -> referencia. False: tipo sem o recurso, auxiliar que nao e barra, vao da
' comparativa, ponto fora das categorias.
Public Function Delta_ResolvePick(cht As Chart, s As Long, p As Long, ByRef ref As String) As Boolean
    Dim kind As String, k As Long
    kind = Registry_ChartType(cht)
    If Not Kinds_Supports(kind, "delta") Then Exit Function
    If s < 1 Or s > cht.SeriesCollection.Count Or p < 1 Then Exit Function
    Select Case Kinds_Family(kind)
        Case "waterfall"
            If Not Delta_IsColumnSeries(cht.SeriesCollection(s)) Then Exit Function
            If p > Geom_ChartCategoryCount(cht) Then Exit Function
            ref = "0:" & p
        Case "waterfallpair"
            If Not Delta_IsColumnSeries(cht.SeriesCollection(s)) Then Exit Function      ' setas, conectores, rotulos e nomes (XY/linhas)
            If p Mod 3 = 0 Then Exit Function                                            ' vao entre as etapas
            If (p + 2) \ 3 > Wfp_StepCount(cht) Then Exit Function
            If p Mod 3 = 1 Then ref = "1:" & ((p + 2) \ 3) Else ref = "2:" & ((p + 2) \ 3)
        Case Else
            k = Delta_DataOrdinal(cht, s)
            If k = 0 Then Exit Function
            If p > Geom_ChartCategoryCount(cht) Then Exit Function
            ref = k & ":" & p
    End Select
    Delta_ResolvePick = True
End Function

' Valor (texto) e extensao lo..hi (no eixo de valores) da referencia. False se ela nao existe mais.
Public Function Delta_Extent(cht As Chart, ref As String, ByRef v As Double, ByRef lo As Double, ByRef hi As Double) As Boolean
    Dim k As Long, c As Long, kind As String, bb As Range, s As Long, i As Long, x As Double, sumPos As Double, sumNeg As Double
    If Not Delta_ParseRef(ref, k, c) Then Exit Function
    kind = Registry_ChartType(cht)
    Select Case Kinds_Family(kind)
        Case "waterfall"
            If k <> 0 Then Exit Function
            Set bb = Block_Body(Registry_Block(cht))
            If c > bb.Rows.Count - 1 Then Exit Function
            v = Util_ToDouble(bb.Cells(c + 1, WF_COL_END).Value)
            lo = Util_ToDouble(bb.Cells(c + 1, WF_COL_LO).Value)
            hi = Util_ToDouble(bb.Cells(c + 1, WF_COL_HI).Value)
        Case "waterfallpair"
            If k < 1 Or k > 2 Or c > Wfp_StepCount(cht) Then Exit Function
            Set bb = Block_Body(Registry_Block(cht))
            v = Util_ToDouble(bb.Cells(c + 1, Wfp_SideCol(k, "Val")).Value)
            lo = Util_ToDouble(bb.Cells(c + 1, Wfp_SideCol(k, "Lo")).Value)
            hi = Util_ToDouble(bb.Cells(c + 1, Wfp_SideCol(k, "Hi")).Value)
        Case Else
            If k < 1 Then Exit Function
            s = Delta_SeriesIndex(cht, k)
            If s = 0 Or c > Geom_ChartCategoryCount(cht) Then Exit Function
            v = Delta_PointValue(cht.SeriesCollection(s), c)
            If Not Delta_IsColumnSeries(cht.SeriesCollection(s)) Then
                lo = v: hi = v                                            ' linha: o proprio ponto
            ElseIf Delta_IsStackedKind(kind) Then
                For i = 1 To s
                    If Not Style_IsHelperSeries(cht.SeriesCollection(i)) And Delta_IsColumnSeries(cht.SeriesCollection(i)) Then
                        x = Delta_PointValue(cht.SeriesCollection(i), c)
                        If x > 0 Then sumPos = sumPos + x Else sumNeg = sumNeg + x
                    End If
                Next
                If v >= 0 Then
                    lo = sumPos - v: hi = sumPos
                Else
                    lo = sumNeg: hi = sumNeg - v
                End If
            Else
                If v >= 0 Then lo = 0: hi = v Else lo = v: hi = 0
            End If
    End Select
    Delta_Extent = True
End Function

' Indice no eixo de categorias (na comparativa, a fatia da barra; nos demais, a categoria).
Public Function Delta_AxisIndex(cht As Chart, ref As String) As Long
    Dim k As Long, c As Long
    If Not Delta_ParseRef(ref, k, c) Then Exit Function
    If Kinds_Family(Registry_ChartType(cht)) = "waterfallpair" Then Delta_AxisIndex = Wfp_SlotOf(c, k) Else Delta_AxisIndex = c
End Function

' Borda mais externa (maior hi acima; menor lo abaixo; nunca aquem do zero) de TODAS as barras do trecho ia..ib do eixo de categorias.
Public Function Delta_SpanEdge(cht As Chart, ia As Long, ib As Long, above As Boolean) As Double
    Dim a As Long, k As Long, fam As String, nData As Long, edge As Double
    fam = Kinds_Family(Registry_ChartType(cht))
    nData = Delta_DataSeriesCount(cht)
    For a = Application.WorksheetFunction.Min(ia, ib) To Application.WorksheetFunction.Max(ia, ib)
        If fam = "waterfall" Then
            Delta_EdgeAdd cht, "0:" & a, above, edge
        ElseIf fam = "waterfallpair" Then
            If a Mod 3 = 1 Then Delta_EdgeAdd cht, "1:" & ((a + 2) \ 3), above, edge
            If a Mod 3 = 2 Then Delta_EdgeAdd cht, "2:" & ((a + 2) \ 3), above, edge
        Else
            For k = 1 To nData
                Delta_EdgeAdd cht, k & ":" & a, above, edge
            Next
        End If
    Next
    Delta_SpanEdge = edge
End Function

Private Sub Delta_EdgeAdd(cht As Chart, ref As String, above As Boolean, ByRef edge As Double)
    Dim v As Double, lo As Double, hi As Double
    If Not Delta_Extent(cht, ref, v, lo, hi) Then Exit Sub
    If above Then
        If hi > edge Then edge = hi
    Else
        If lo < edge Then edge = lo
    End If
End Sub

' Centro da barra no sentido das categorias (pt). Agrupadas: o centro do PONTO da serie (cada serie tem a sua barra; geometria do Excel,
' valida depois do layout); demais: o centro da categoria (Geom_ChartCatCoord; na comparativa a fatia).
Public Function Delta_CatCenter(cht As Chart, ref As String) As Double
    Dim k As Long, c As Long, kind As String, pt As Point
    Delta_ParseRef ref, k, c
    kind = Registry_ChartType(cht)
    If kind = "clustered" Or kind = "barclustered" Then
        Set pt = cht.SeriesCollection(Delta_SeriesIndex(cht, k)).Points(c)
        If Geom_ChartIsHorizontal(cht) Then Delta_CatCenter = pt.Top + pt.Height / 2 Else Delta_CatCenter = pt.Left + pt.Width / 2
    Else
        Delta_CatCenter = Geom_ChartCatCoord(cht, Delta_AxisIndex(cht, ref))
    End If
End Function

' Meia espessura da barra no sentido das categorias (moldura da origem no modo de dois cliques).
Public Function Delta_BarHalf(cht As Chart, ref As String) As Double
    Dim k As Long, c As Long, kind As String, pt As Point, s As Long
    Delta_ParseRef ref, k, c
    kind = Registry_ChartType(cht)
    If kind = "clustered" Or kind = "barclustered" Then
        Set pt = cht.SeriesCollection(Delta_SeriesIndex(cht, k)).Points(c)
        If Geom_ChartIsHorizontal(cht) Then Delta_BarHalf = pt.Height / 2 Else Delta_BarHalf = pt.Width / 2
    ElseIf kind = "line" Then
        Delta_BarHalf = 5
    Else
        s = Delta_SeriesIndex(cht, k)
        If s > 0 Then
            If Not Delta_IsColumnSeries(cht.SeriesCollection(s)) Then Delta_BarHalf = 5: Exit Function
        End If
        Delta_BarHalf = Geom_ChartBarHalfWidth(cht)
    End If
End Function
```
Confira antes: `Wfp_SideCol(j, "Lo")` e `"Hi"` existem em `Wfp_PartIndex` (spec da comparativa §4.2: `Src, Tot, D, Start, End, Lo, Hi, Base, PP, PN, Val, Big`); na fixture da comparativa a etapa 2 do lado B vai de 29 a 22 (Lo 22, Hi 29) — se o bloco disser outra coisa, o teste está errado, não o código: corrija a expectativa a partir do bloco e relate.

- [ ] **Passo 4: GREEN + suíte completa** e **Commit** — `git commit -m "feat: referencias, extensoes e textos do delta entre barras (modDelta)" (+ trailer)`.

---

### Tarefa B2: Seta de delta: desenho em cotovelo, faixas sem colisão, espaço, recálculo — modelo: most-capable

**Files:**
- Modify: `src/core/modules/modAnnotations.bas` (`Ann_AddDelta`, `Ann_LayoutDeltas`, `Ann_DeltaPrepare`, `Ann_DrawDelta`, `Ann_DeltaCleanup`, `Ann_EnsureDeltaRoom`, `Ann_SetParam`, integração em `Ann_Refresh` e `Ann_EnsureRoomAfterChange`)
- Modify: `src/core/modules/modCoreApi.bas` (`Core_AddDelta`)
- Test: `tests/vba/core/TestDelta.bas` (acrescentar)

**Interfaces:**
- Consumes (B1): `Delta_*`, `DELTA_GAP`, `DELTA_LABEL_ROOM`, `DELTA_BAD_PICK`; (B0) `ANN_DELTA_REFUSAL`.
- Produces: `Public Function Ann_AddDelta(cht As Chart, fromRef As String, toRef As String, Optional fmt As String = "both") As Shape` (erros 7001/7003; transacional); `Public Sub Core_AddDelta(fromRef As String, toRef As String, Optional fmt As String = "both", Optional cht As Chart = Nothing)`; formas `tqq_ann_delta_<n>` (linha principal, `kind=delta;v=1;from=..;to=..;fmt=..`), `_t1`, `_t2` (`kind=deltaleg;owner=..`), `_lbl`.

- [ ] **Passo 1: Testes (RED)** — acrescente a `tests/vba/core/TestDelta.bas`:

```vb
' Regra de tres propria no grafico (independente de Geom_*): y de um valor nas colunas, x nas barras.
Private Function YOf(c As Chart, v As Double) As Double
    YOf = c.PlotArea.InsideTop + c.PlotArea.InsideHeight * (c.Axes(xlValue).MaximumScale - v) / (c.Axes(xlValue).MaximumScale - c.Axes(xlValue).MinimumScale)
End Function

Private Function XOf(c As Chart, v As Double) As Double
    XOf = c.PlotArea.InsideLeft + c.PlotArea.InsideWidth * (v - c.Axes(xlValue).MinimumScale) / (c.Axes(xlValue).MaximumScale - c.Axes(xlValue).MinimumScale)
End Function

Private Function HasShape(c As Chart, nm As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = c.Shapes(nm)
    HasShape = Not s Is Nothing
End Function

Private Function Relayout(c As Chart) As Chart
    ForceLayout c
    Ann_Refresh c
    ForceLayout c
    Set Relayout = c
End Function

Private Function Overlaps(a As Shape, b As Shape) As Boolean
    Overlaps = Not (a.Left + a.Width <= b.Left Or b.Left + b.Width <= a.Left Or a.Top + a.Height <= b.Top Or b.Top + b.Height <= a.Top)
End Function

' Review Focus 4: 1a serie T1 (10) -> 2a serie T3 (-4) em colunas agrupadas: perna da origem 3 pt acima da barra, ponta no destino 3 pt
' acima do zero (barra negativa), linha principal acima de TODAS as barras do trecho (a maior e 30) e do rotulo, texto "-14 (-140%)".
Public Sub Test_Delta_Columns_ElbowGeometry()
    Dim wb As Workbook, c As Chart, t1 As Shape, t2 As Shape, ln As Shape, lbl As Shape, pa As Point, pb As Point, xA As Double, xB As Double
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    Core_ClearLastNotice
    Core_AddDelta "1:1", "2:3"
    AssertEqual "", Core_LastNotice()
    AssertEqual "kind=delta;v=1;from=1:1;to=2:3;fmt=both", c.Shapes("tqq_ann_delta_1").AlternativeText
    Relayout c
    Set ln = c.Shapes("tqq_ann_delta_1"): Set t1 = c.Shapes("tqq_ann_delta_1_t1"): Set t2 = c.Shapes("tqq_ann_delta_1_t2"): Set lbl = c.Shapes("tqq_ann_delta_1_lbl")
    Set pa = c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(1): Set pb = c.SeriesCollection(Delta_SeriesIndex(c, 2)).Points(3)
    xA = pa.Left + pa.Width / 2: xB = pb.Left + pb.Width / 2
    AssertAlmost xA, t1.Left, 1, "perna da origem no centro da barra Sul T1"
    AssertAlmost YOf(c, 10) - 3, t1.Top + t1.Height, 1, "perna da origem comeca 3 pt acima da barra"
    AssertAlmost xB, t2.Left, 1, "perna do destino no centro da barra Norte T3"
    AssertAlmost YOf(c, 0) - 3, t2.Top + t2.Height, 1, "ponta no destino: 3 pt acima do zero (barra negativa)"
    AssertEqual CLng(msoArrowheadTriangle), CLng(t2.Line.EndArrowheadStyle)
    AssertTrue (t2.VerticalFlip = msoFalse), "a ponta (fim da linha) fica embaixo, no destino"
    AssertAlmost ln.Top, t1.Top, 0.5, "linha principal na altura do topo das pernas"
    AssertTrue lbl.Top + lbl.Height <= YOf(c, 30) - 8 + 0.5, "rotulo acima da maior barra do trecho (30) com folga"
    AssertAlmost (xA + xB) / 2, lbl.Left + lbl.Width / 2, 1, "rotulo centrado"
    AssertEqual Delta_Text(10, -4, "both", Style_NumberFormat(Registry_GetOption(c, "numfmt", "int"))), CStr(lbl.TextFrame2.TextRange.Text)
    Teardown wb
End Sub

' Duas setas com trechos sobrepostos: a mais estreita fica mais perto das barras e a larga sobe ate os rotulos nao se tocarem.
Public Sub Test_Delta_Lanes_NestedArrows()
    Dim wb As Workbook, c As Chart, l1 As Shape, l2 As Shape
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    Core_AddDelta "1:1", "1:3"          ' larga (T1 -> T3)
    Core_AddDelta "1:2", "2:2"          ' estreita (dentro de T2)
    Relayout c
    Set l1 = c.Shapes("tqq_ann_delta_1_lbl"): Set l2 = c.Shapes("tqq_ann_delta_2_lbl")
    AssertTrue c.Shapes("tqq_ann_delta_2").Top > c.Shapes("tqq_ann_delta_1").Top, "a estreita fica mais baixa (perto das barras)"
    AssertTrue c.Shapes("tqq_ann_delta_2").Top - c.Shapes("tqq_ann_delta_1").Top >= l1.Height / 2 + l2.Height / 2 + 6 - 0.5, "faixas separadas"
    AssertFalse Overlaps(l1, l2), "rotulos nao se tocam"
    Teardown wb
End Sub

' Barras horizontais: tudo girado (pernas horizontais, linha principal vertical a direita das barras).
Public Sub Test_Delta_Bars_Horizontal()
    Dim wb As Workbook, c As Chart, t1 As Shape, ln As Shape
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "barclustered")
    Core_AddDelta "1:1", "1:3"
    Relayout c
    Set t1 = c.Shapes("tqq_ann_delta_1_t1"): Set ln = c.Shapes("tqq_ann_delta_1")
    AssertAlmost XOf(c, 10) + 3, t1.Left, 1, "perna da origem comeca 3 pt depois do fim da barra"
    AssertTrue ln.Width < 0.5, "linha principal vertical"
    AssertTrue ln.Left >= XOf(c, 30) + 8 - 0.5, "a direita da maior barra do trecho"
    Teardown wb
End Sub

' Duas barras inteiramente negativas: cotovelo ABAIXO.
Public Sub Test_Delta_NegativeBelow()
    Dim wb As Workbook, c As Chart, rng As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Worksheet.Range("B2").Value = -5: rng.Worksheet.Range("B3").Value = -8
    Set c = InsertOn(rng, "clustered")
    Core_AddDelta "1:1", "1:2"
    Relayout c
    AssertTrue c.Shapes("tqq_ann_delta_1").Top >= YOf(c, -8) + 8 - 0.5, "linha principal abaixo da barra mais funda"
    AssertAlmost YOf(c, -5) + 3, c.Shapes("tqq_ann_delta_1_t1").Top, 1, "perna da origem 3 pt abaixo da barra"
    Teardown wb
End Sub

' Escala fixa no maior valor: sem sala, a area de plotagem desce e o rotulo fica dentro do grafico.
Public Sub Test_Delta_RoomTop_FixedMax()
    Dim wb As Workbook, c As Chart, top0 As Double
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    c.Axes(xlValue).MaximumScale = 30
    ForceLayout c
    top0 = c.PlotArea.InsideTop
    Core_AddDelta "1:1", "1:3"
    Relayout c
    AssertTrue c.Shapes("tqq_ann_delta_1_lbl").Top >= 1.5, "rotulo inteiro dentro do grafico (" & c.Shapes("tqq_ann_delta_1_lbl").Top & ")"
    AssertTrue c.PlotArea.InsideTop > top0, "a area de plotagem desceu"
    Teardown wb
End Sub

' Dado novo: texto e geometria acompanham (sem refazer a seta).
Public Sub Test_Delta_DataChange_TextAndGeometry()
    Dim wb As Workbook, c As Chart, rng As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    Set c = InsertOn(rng, "clustered")
    Core_AddDelta "1:1", "1:3"
    rng.Worksheet.Range("B4").Value = 40
    Application.Calculate
    Relayout c
    AssertEqual Delta_Text(10, 40, "both", Style_NumberFormat(Registry_GetOption(c, "numfmt", "int"))), CStr(c.Shapes("tqq_ann_delta_1_lbl").TextFrame2.TextRange.Text)
    AssertTrue c.Shapes("tqq_ann_delta_1_lbl").Top + c.Shapes("tqq_ann_delta_1_lbl").Height <= YOf(c, 40) - 8 + 0.5
    Teardown wb
End Sub

' Referencia que nao existe mais: seta oculta (nao apagada).
Public Sub Test_Delta_InvalidRefHidden()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    Core_AddDelta "1:1", "1:3"
    c.Shapes("tqq_ann_delta_1").AlternativeText = "kind=delta;v=1;from=1:1;to=1:9;fmt=both"
    Ann_Refresh c
    AssertEqual CLng(msoFalse), CLng(c.Shapes("tqq_ann_delta_1").Visible)
    AssertEqual CLng(msoFalse), CLng(c.Shapes("tqq_ann_delta_1_lbl").Visible)
    AssertEqual CLng(msoFalse), CLng(c.Shapes("tqq_ann_delta_1_t2").Visible)
    Teardown wb
End Sub

' Recusas: tipo sem o recurso (7001) e mesma barra / referencia invalida (7003) viram aviso; nada criado.
Public Sub Test_Delta_Refusals()
    Dim wb As Workbook, c As Chart, ws As Worksheet
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    Core_ClearLastNotice
    Core_AddDelta "1:2", "1:2"
    AssertEqual DELTA_BAD_PICK, Core_LastNotice()
    Core_AddDelta "1:2", "5:1"
    AssertEqual DELTA_BAD_PICK, Core_LastNotice()
    AssertFalse HasShape(c, "tqq_ann_delta_1")
    Teardown wb
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = "Cat": ws.Range("B1").Value = "V": ws.Range("A2").Value = "a": ws.Range("B2").Value = 1: ws.Range("A3").Value = "b": ws.Range("B3").Value = 2
    Set c = InsertOn(ws.Range("A1:B3"), "pie")
    Core_ClearLastNotice
    Core_AddDelta "1:1", "1:2"
    AssertEqual ANN_DELTA_REFUSAL, Core_LastNotice()
    Teardown wb
End Sub

' Falha no meio da criacao (gancho): nenhuma forma da seta sobra.
Public Sub Test_Delta_AddFailure_RollsBack()
    Dim wb As Workbook, c As Chart, i As Long, n As Long
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    Util_FailThinSyncForTest = "delta-add:1"
    Core_ClearLastNotice
    Core_AddDelta "1:1", "1:3"
    Util_FailThinSyncForTest = ""
    For i = 1 To c.Shapes.Count
        If Left$(c.Shapes(i).Name, 14) = "tqq_ann_delta_" Then n = n + 1
    Next
    AssertEqual 0, n, "nada sobra"
    Teardown wb
End Sub

' Apagar a linha principal a mao: pernas e rotulo orfaos somem no proximo refresh; Remover setas apaga tudo.
Public Sub Test_Delta_Orphans_And_RemoveAll()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    Core_AddDelta "1:1", "1:3"
    Core_AddDelta "2:1", "2:2"
    c.Shapes("tqq_ann_delta_1").Delete
    Ann_Refresh c
    AssertFalse HasShape(c, "tqq_ann_delta_1_lbl"): AssertFalse HasShape(c, "tqq_ann_delta_1_t1")
    AssertTrue HasShape(c, "tqq_ann_delta_2_lbl")
    c.Parent.Select
    Core_RemoveAnnotations
    AssertFalse HasShape(c, "tqq_ann_delta_2")
    Teardown wb
End Sub

' Cascata: texto pelo valor acumulado (Receita 100 -> Lucro 70) e perna da origem no topo da barra (Hi 100).
Public Sub Test_Delta_Waterfall_EndValues()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Set c = InsertOn(WaterfallFixture(wb), "waterfall")
    Core_AddDelta "0:1", "0:3"
    Relayout c
    AssertEqual Delta_Text(100, 70, "both", Style_NumberFormat(Registry_GetOption(c, "numfmt", "int"))), CStr(c.Shapes("tqq_ann_delta_1_lbl").TextFrame2.TextRange.Text)
    AssertAlmost YOf(c, 100) - 3, c.Shapes("tqq_ann_delta_1_t1").Top + c.Shapes("tqq_ann_delta_1_t1").Height, 1
    Teardown wb
End Sub

' Comparativa: delta entre etapas diferentes (A1 -> A3, fora do escopo das setas de variacao da 1.3.0).
Public Sub Test_Delta_WaterfallPair_BetweenSteps()
    Dim wb As Workbook, c As Chart, v1 As Double, v3 As Double, lo As Double, hi As Double
    Set wb = NewTempWorkbook()
    Set c = InsertOn(PairFixture(wb), "waterfallpair")
    Core_AddDelta "1:1", "1:3"
    Relayout c
    AssertTrue Delta_Extent(c, "1:1", v1, lo, hi) And Delta_Extent(c, "1:3", v3, lo, hi)
    AssertEqual Delta_Text(v1, v3, "both", Style_NumberFormat(Registry_GetOption(c, "numfmt", "int"))), CStr(c.Shapes("tqq_ann_delta_1_lbl").TextFrame2.TextRange.Text)
    AssertAlmost c.PlotArea.InsideLeft + c.PlotArea.InsideWidth * (1 - 0.5) / 8, c.Shapes("tqq_ann_delta_1_t1").Left, 1, "fatia A da etapa 1 (8 fatias)"
    AssertAlmost c.PlotArea.InsideLeft + c.PlotArea.InsideWidth * (7 - 0.5) / 8, c.Shapes("tqq_ann_delta_1_t2").Left, 1, "fatia A da etapa 3"
    Teardown wb
End Sub
```
**Teste condicional à Parte A** (só se `modPptSync.bas` existir na branch; senão não o acrescente): copie de `TestPptSync.bas` os helpers privados `StartScratchPowerPoint`, `StopScratchPowerPoint`, `UserWindowAppeared`, `PowerPointProcessCount`, `WaitSeconds` e acrescente `Test_Delta_InPowerPoint_WhenAvailable`: cola (como em `Test_PptSync_PastedChart_...`) um gráfico de colunas agrupadas com `Core_AddDelta "1:1", "2:3"` e confere que o gráfico do slide tem `tqq_ann_delta_1`, `_t1`, `_t2`, `_lbl` com o mesmo texto do Excel e que o topo da perna `_t1` no slide fica 3 pt acima da barra Sul T1 medida com a regra de três **do gráfico do slide** (`PptYOf`), tolerância 1,5 pt.

- [ ] **Passo 2: RED** — `-Filter "TestDelta"` → saída 3 (`Core_AddDelta` não existe).

- [ ] **Passo 3: Implementar em `modAnnotations.bas`** — (a) depois de `Private Const ANN_BRACKET_MAX_FRACTION ...`:

```vb
Private Const ANN_DELTA As String = "tqq_ann_delta_"      ' prefixo das setas de delta (linha principal nm; pernas nm_t1/nm_t2; rotulo nm_lbl)
```
(b) depois de `Ann_AddBracket`:

```vb
' Delta entre duas barras (Parte B): linha principal nm (kind=delta;v=1;from;to;fmt), pernas nm_t1 (da origem) e nm_t2 (com a ponta no
' destino; kind=deltaleg;owner=nm) e o rotulo nm_lbl. Ann_LayoutDeltas posiciona; Ann_EnsureDeltaRoom abre a sala. Recusas: 7001 (matriz),
' 7003 (mesma barra ou referencia inexistente). Transacional: falha no meio apaga as formas criadas.
Public Function Ann_AddDelta(cht As Chart, fromRef As String, toRef As String, Optional fmt As String = "both") As Shape
    Dim nm As String, ln As Shape, t As Shape, sfx As Variant, refusal As String, v As Double, lo As Double, hi As Double, en As Long, ed As String
    refusal = Ann_RefusalText(cht, "delta")
    If Len(refusal) > 0 Then Err.Raise vbObjectError + 7001, "Ann_AddDelta", refusal
    If fromRef = toRef Then Err.Raise vbObjectError + 7003, "Ann_AddDelta", DELTA_BAD_PICK
    If Not Delta_Extent(cht, fromRef, v, lo, hi) Then Err.Raise vbObjectError + 7003, "Ann_AddDelta", DELTA_BAD_PICK
    If Not Delta_Extent(cht, toRef, v, lo, hi) Then Err.Raise vbObjectError + 7003, "Ann_AddDelta", DELTA_BAD_PICK
    nm = Ann_NewName(cht, "delta")
    On Error GoTo Fail
    Set ln = cht.Shapes.AddLine(0, 0, 10, 0)
    ln.Name = nm
    ln.AlternativeText = "kind=delta;v=1;from=" & fromRef & ";to=" & toRef & ";fmt=" & fmt
    Ann_StyleDeltaLine ln, False
    For Each sfx In Array("_t1", "_t2")
        Set t = cht.Shapes.AddLine(0, 0, 0, 10)
        t.Name = nm & sfx
        t.AlternativeText = "kind=deltaleg;owner=" & nm
        Ann_StyleDeltaLine t, (sfx = "_t2")
    Next
    Ann_AddLabel cht, nm
    If BUILD_CONFIG = "Test" Then
        If Util_FailThinSyncForTest = "delta-add:1" Then Util_FailThinSyncForTest = "": Err.Raise vbObjectError + 9993, "Ann_AddDelta", "falha simulada (teste)"
    End If
    Ann_Refresh cht
    Ann_EnsureDeltaRoom cht
    Set Ann_AddDelta = ln
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    For Each sfx In Array("_lbl", "_t2", "_t1", "")
        cht.Shapes(nm & sfx).Delete
    Next
    On Error GoTo 0
    Err.Raise en, "Ann_AddDelta", ed
End Function

Private Sub Ann_StyleDeltaLine(ln As Shape, withHead As Boolean)
    With ln.Line
        .ForeColor.RGB = RGB(89, 89, 89)
        .Weight = 1
        If withHead Then
            .EndArrowheadStyle = msoArrowheadTriangle
            .EndArrowheadLength = msoArrowheadShort
            .EndArrowheadWidth = msoArrowheadNarrow
        End If
    End With
End Sub

' Troca (ou acrescenta) o valor de uma chave do texto alternativo "k=v;k=v". Pura.
Private Function Ann_SetParam(altText As String, key As String, value As String) As String
    Dim parts() As String, i As Long, found As Boolean
    parts = Split(altText, ";")
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), Len(key) + 1) = key & "=" Then parts(i) = key & "=" & value: found = True
    Next
    Ann_SetParam = Join(parts, ";")
    If Not found Then Ann_SetParam = Ann_SetParam & ";" & key & "=" & value
End Function

Private Function Ann_Overlap(a1 As Double, a2 As Double, b1 As Double, b2 As Double, margin As Double) As Boolean
    Dim loA As Double, hiA As Double, loB As Double, hiB As Double
    If a1 < a2 Then loA = a1: hiA = a2 Else loA = a2: hiA = a1
    If b1 < b2 Then loB = b1: hiB = b2 Else loB = b2: hiB = b1
    Ann_Overlap = (loA <= hiB + margin And loB <= hiA + margin)
End Function

' Uma seta de delta pronta para o layout (rec): 0 nome, 1 centro da origem e 2 do destino (sentido das categorias, pt), 3 e 4 ancoras da
' origem e do destino (sentido dos valores, pt; 3 pt para fora da borda), 5 sentido para fora (-1 ou +1, em pt no sentido dos valores),
' 6 linha principal (pt; a altura base, ajustada pelas faixas), 7 meia espessura do rotulo no sentido dos valores. False = oculta.
Private Function Ann_DeltaPrepare(cht As Chart, ln As Shape, fmt As String, ByRef rec As Variant) As Boolean
    Dim fromRef As String, toRef As String, mode As String, vA As Double, loA As Double, hiA As Double, vB As Double, loB As Double, hiB As Double
    Dim above As Boolean, horiz As Boolean, sgn As Double, edgeV As Double, half As Double, lbl As Shape, extra As Double, aA As Double, aB As Double
    If Not Kinds_Supports(Registry_ChartType(cht), "delta") Then Exit Function
    fromRef = Ann_ParseParam(ln.AlternativeText, "from"): toRef = Ann_ParseParam(ln.AlternativeText, "to")
    mode = Ann_ParseParam(ln.AlternativeText, "fmt")
    If fromRef = toRef Then Exit Function
    If Not Delta_Extent(cht, fromRef, vA, loA, hiA) Then Exit Function
    If Not Delta_Extent(cht, toRef, vB, loB, hiB) Then Exit Function
    horiz = Geom_ChartIsHorizontal(cht)
    above = Not (hiA <= 0 And hiB <= 0 And (loA < 0 Or loB < 0))
    If horiz Then sgn = 1 Else sgn = -1
    If Not above Then sgn = -sgn
    Set lbl = Ann_Shape(cht, ln.Name & "_lbl")
    If Not lbl Is Nothing Then
        lbl.TextFrame2.TextRange.Text = Delta_Text(vA, vB, mode, fmt)
        If horiz Then half = lbl.Width / 2 Else half = lbl.Height / 2
    End If
    If Registry_GetOption(cht, "labels", "1") = "1" Or Registry_GetOption(cht, "totals", "0") = "1" Then extra = DELTA_LABEL_ROOM
    edgeV = Delta_SpanEdge(cht, Delta_AxisIndex(cht, fromRef), Delta_AxisIndex(cht, toRef), above)
    If above Then aA = hiA: aB = hiB Else aA = loA: aB = loB
    rec = Array(ln.Name, Delta_CatCenter(cht, fromRef), Delta_CatCenter(cht, toRef), Geom_ChartValCoord(cht, aA) + sgn * 3, Geom_ChartValCoord(cht, aB) + sgn * 3, sgn, Geom_ChartValCoord(cht, edgeV) + sgn * (DELTA_GAP + extra + half), half)
    Ann_DeltaPrepare = True
End Function

' Pernas, linha principal e rotulo de uma seta ja com a linha principal decidida (rec(6)).
Private Sub Ann_DrawDelta(cht As Chart, rec As Variant)
    Dim nm As String, ln As Shape, t1 As Shape, t2 As Shape, lbl As Shape, cA As Double, cB As Double, aA As Double, aB As Double, lv As Double
    nm = rec(0): cA = rec(1): cB = rec(2): aA = rec(3): aB = rec(4): lv = rec(6)
    Set ln = Ann_Shape(cht, nm): Set t1 = Ann_Shape(cht, nm & "_t1"): Set t2 = Ann_Shape(cht, nm & "_t2"): Set lbl = Ann_Shape(cht, nm & "_lbl")
    If Geom_ChartIsHorizontal(cht) Then
        If Not t1 Is Nothing Then Ann_MoveLineIfMoved t1, aA, cA, lv, cA
        If Not ln Is Nothing Then Ann_MoveLineIfMoved ln, lv, cA, lv, cB
        If Not t2 Is Nothing Then Ann_MoveLineIfMoved t2, lv, cB, aB, cB
        If Not lbl Is Nothing Then lbl.Left = lv - lbl.Width / 2: lbl.Top = (cA + cB) / 2 - lbl.Height / 2
    Else
        If Not t1 Is Nothing Then Ann_MoveLineIfMoved t1, cA, aA, cA, lv
        If Not ln Is Nothing Then Ann_MoveLineIfMoved ln, cA, lv, cB, lv
        If Not t2 Is Nothing Then Ann_MoveLineIfMoved t2, cB, lv, cB, aB
        If Not lbl Is Nothing Then lbl.Left = (cA + cB) / 2 - lbl.Width / 2: lbl.Top = lv - lbl.Height / 2
    End If
    If Not lbl Is Nothing Then Ann_ClampShape cht, lbl
End Sub

' Formas de delta orfas (perna ou rotulo cuja linha principal foi apagada a mao) sao apagadas.
Private Sub Ann_DeltaCleanup(cht As Chart)
    Dim i As Long, nm As String, k As String
    For i = cht.Shapes.Count To 1 Step -1
        nm = cht.Shapes(i).Name
        If Left$(nm, Len(ANN_DELTA)) = ANN_DELTA Then
            k = Ann_ParseParam(cht.Shapes(i).AlternativeText, "kind")
            If k = "deltaleg" Or k = "label" Then
                If Ann_Shape(cht, Ann_ParseParam(cht.Shapes(i).AlternativeText, "owner")) Is Nothing Then cht.Shapes(i).Delete
            End If
        End If
    Next
End Sub

' Todas as setas de delta juntas (spec 5.3): ordem pela largura do trecho (menor primeiro); cada uma comeca na altura base e se afasta das ja
' postas do mesmo lado cujo trecho se sobrepoe (2 pt de margem) ate ficar a (meia espessura dela + da outra + 6 pt). Erro numa seta: oculta.
Private Sub Ann_LayoutDeltas(cht As Chart)
    Dim i As Long, j As Long, k As Long, tmp As Long, shp As Shape, recs As Collection, rec As Variant, p As Variant, fmt As String
    Dim order() As Long, w() As Double, placed() As Variant, nPlaced As Long, lineV As Double, stp As Double, moved As Boolean, passes As Long, valid As Boolean
    Ann_DeltaCleanup cht
    Set recs = New Collection
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    For i = 1 To cht.Shapes.Count
        Set shp = cht.Shapes(i)
        If Left$(shp.Name, Len(ANN_DELTA)) = ANN_DELTA Then
            If Ann_ParseParam(shp.AlternativeText, "kind") = "delta" Then
                valid = False
                On Error Resume Next
                valid = Ann_DeltaPrepare(cht, shp, fmt, rec)
                If Err.Number <> 0 Then
                    valid = False
                    mLastError = "Ann_LayoutDeltas " & shp.Name & " " & Err.Number & " " & Err.Description
                    Errors_Log mLastError
                End If
                Err.Clear
                On Error GoTo 0
                Ann_SetShapeVisible cht, shp.Name, valid
                If valid Then recs.Add rec
            End If
        End If
    Next
    If recs.Count = 0 Then Exit Sub
    ReDim order(1 To recs.Count): ReDim w(1 To recs.Count): ReDim placed(1 To recs.Count)
    For i = 1 To recs.Count
        order(i) = i: w(i) = Abs(recs(i)(2) - recs(i)(1))
    Next
    For i = 2 To recs.Count                                   ' insercao: menor largura primeiro (estavel)
        j = i
        Do While j > 1
            If w(order(j)) < w(order(j - 1)) Then
                tmp = order(j): order(j) = order(j - 1): order(j - 1) = tmp: j = j - 1
            Else
                Exit Do
            End If
        Loop
    Next
    For k = 1 To recs.Count
        rec = recs(order(k))
        lineV = rec(6)
        passes = 0
        Do
            moved = False
            For j = 1 To nPlaced
                p = placed(j)
                If p(5) = rec(5) And Ann_Overlap(CDbl(rec(1)), CDbl(rec(2)), CDbl(p(1)), CDbl(p(2)), 2) Then
                    stp = p(7) + rec(7) + 6
                    If Abs(lineV - p(6)) < stp - 0.001 Then
                        lineV = p(6) + rec(5) * stp
                        moved = True
                    End If
                End If
            Next
            passes = passes + 1
        Loop While moved And passes <= nPlaced + 1
        rec(6) = lineV
        nPlaced = nPlaced + 1
        placed(nPlaced) = rec
        On Error Resume Next
        Ann_DrawDelta cht, rec
        If Err.Number <> 0 Then
            mLastError = "Ann_LayoutDeltas " & rec(0) & " " & Err.Number & " " & Err.Description
            Errors_Log mLastError
            Ann_SetShapeVisible cht, CStr(rec(0)), False
        End If
        Err.Clear
        On Error GoTo 0
    Next
End Sub

' Sala das setas de delta (spec 5.3): se alguma forma de delta visivel passa da borda livre (2 pt), a area de plotagem encolhe pelo deficit
' e o layout e refeito; ate 3 passadas. Colunas: topo e base (Ann_BottomEdge); barras: direita (Geom_ChartRightEdge) e esquerda. Nunca propaga.
Private Sub Ann_EnsureDeltaRoom(cht As Chart)
    Dim k As Long, i As Long, s As Shape, minTop As Double, maxBottom As Double, minLeft As Double, maxRight As Double, d As Double, found As Boolean
    On Error GoTo Done
    For k = 1 To 3
        minTop = 1E+30: maxBottom = -1E+30: minLeft = 1E+30: maxRight = -1E+30: found = False
        For i = 1 To cht.Shapes.Count
            Set s = cht.Shapes(i)
            If Left$(s.Name, Len(ANN_DELTA)) = ANN_DELTA And s.Visible = msoTrue Then
                found = True
                If s.Top < minTop Then minTop = s.Top
                If s.Top + s.Height > maxBottom Then maxBottom = s.Top + s.Height
                If s.Left < minLeft Then minLeft = s.Left
                If s.Left + s.Width > maxRight Then maxRight = s.Left + s.Width
            End If
        Next
        If Not found Then Exit Sub
        d = 0
        If Geom_ChartIsHorizontal(cht) Then
            If maxRight > Geom_ChartRightEdge(cht) - 2 Then d = maxRight - (Geom_ChartRightEdge(cht) - 2): cht.PlotArea.Width = cht.PlotArea.Width - d
            If minLeft < 2 Then cht.PlotArea.Left = cht.PlotArea.Left + (2 - minLeft): cht.PlotArea.Width = cht.PlotArea.Width - (2 - minLeft): d = d + 2 - minLeft
        Else
            If minTop < 2 Then cht.PlotArea.Height = cht.PlotArea.Height - (2 - minTop): cht.PlotArea.Top = cht.PlotArea.Top + (2 - minTop): d = 2 - minTop
            If maxBottom > Ann_BottomEdge(cht) - 2 Then cht.PlotArea.Height = cht.PlotArea.Height - (maxBottom - (Ann_BottomEdge(cht) - 2)): d = d + maxBottom - (Ann_BottomEdge(cht) - 2)
        End If
        If d <= 0.01 Then Exit Sub
        Ann_Refresh cht
    Next
    Exit Sub
Done:
    Errors_Log "Ann_EnsureDeltaRoom" & vbTab & Err.Number & vbTab & Err.Description
End Sub
```
(c) em `Ann_Refresh`, troque o fim

```vb
NextShape:
    Next
Done:
    mRefreshing = False
End Sub
```
por
```vb
NextShape:
    Next
    ' delta entre barras: todas juntas (faixas), depois das demais anotacoes; erro isolado
    On Error Resume Next
    Ann_LayoutDeltas cht
    If Err.Number <> 0 Then mLastError = "Ann_Refresh delta " & Err.Number & " " & Err.Description: Errors_Log mLastError
    Err.Clear
Done:
    mRefreshing = False
End Sub
```
(d) em `Ann_EnsureRoomAfterChange`, troque a condição e acrescente a sala do delta:

```vb
Public Sub Ann_EnsureRoomAfterChange(cht As Chart)
    Dim nDelta As Long
    nDelta = Ann_CountPrefix(cht, ANN_DELTA)
    If Ann_CountPrefix(cht, ANN_PREFIX & "cagr_") + Ann_CountPrefix(cht, ANN_PREFIX & "diff_") + Ann_CountPrefix(cht, ANN_BRACKET) + nDelta = 0 Then Exit Sub
    Ann_EnsureRoom cht
    If nDelta > 0 Then Ann_EnsureDeltaRoom cht
End Sub
```
- [ ] **Passo 4: `Core_AddDelta`** — em `src/core/modules/modCoreApi.bas`, depois de `Core_AddAnnotation`:

```vb
' Delta entre duas barras (Parte B): referencias "k:c" (modDelta). Chamado pelo modo de dois cliques (com o grafico) e pelo gerador da
' demonstracao (Application.Run, grafico ativo). Recusas 7001 (tipo) e 7003 (mesma barra/referencia invalida) viram aviso. Nunca propaga.
Public Sub Core_AddDelta(fromRef As String, toRef As String, Optional fmt As String = "both", Optional cht As Chart = Nothing)
    Dim c As Chart
    On Error GoTo Fail
    If cht Is Nothing Then Set c = Core_ActiveChart() Else Set c = cht
    If c Is Nothing Then Exit Sub
    If fmt <> "abs" And fmt <> "pct" Then fmt = "both"
    Ann_AddDelta c, fromRef, toRef, fmt
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7001 Or Err.Number = vbObjectError + 7003 Then
        Core_Notify Err.Description, vbInformation
    Else
        Errors_Handle "Core_AddDelta"
    End If
End Sub
```
- [ ] **Passo 5: GREEN** — `-Filter "TestDelta|TestAnnotations"` → 0 falhas. Se `Test_Delta_Columns_ElbowGeometry` falhar só no `x` (centro do ponto), confira se `ForceLayout` rodou antes de `Ann_Refresh` (o `Relayout` faz) e se `Points(c).Left` está no mesmo referencial das formas (deve: ambos relativos à área do gráfico); relate a medida antes de mudar a regra.
- [ ] **Passo 6: Suíte completa** e **Commit** — `git commit -m "feat: seta de delta entre barras (cotovelo, faixas sem colisao, sala, recalculo com os dados)" (+ trailer)`.

---

### Tarefa B3: Modo de dois cliques (eventos do gráfico) — modelo: most-capable

**Files:**
- Create: `src/core/modules/modDeltaPick.bas`
- Modify: `src/core/classes/clsChartWatcher.cls` (`Cht_Select`, `Cht_MouseUp`), `src/core/classes/clsAppEvents.cls` (`App_SheetSelectionChange`), `src/core/modules/modAnnotations.bas` (`Ann_DeltaCleanup` apaga a moldura de gráfico não armado), `src/core/modules/modCoreApi.bas` (`Core_DeltaPick`)
- Test: `tests/vba/core/TestDeltaPick.bas` (novo)

**Interfaces:**
- Consumes (B1/B2): `Delta_ResolvePick`, `Delta_Extent`, `Delta_CatCenter`, `Delta_BarHalf`, `Core_AddDelta(from, to, fmt, cht)`, `ANN_DELTA_REFUSAL`.
- Produces: `DELTAPICK_MARK = "tqq_pickmark"`, `DELTAPICK_MSG_FROM/TO/SAME/NOTBAR/SERIES/CANCEL`; `DeltaPick_Arm`, `DeltaPick_Disarm(msg)`, `DeltaPick_IsArmed`, `DeltaPick_IsArmedFor(cht)`, `DeltaPick_FromRef`, `DeltaPick_OnSelect(cht, id, a1, a2)`, `DeltaPick_OnMouseUp(cht, x, y)`, `DeltaPick_OnSheetClick`, `DeltaPick_EscKey`, `DeltaPick_Timeout`, `DeltaPick_ExpireForTest`; `Public Sub Core_DeltaPick()`.

- [ ] **Passo 1: Testes (RED)** — crie `tests/vba/core/TestDeltaPick.bas`:

```vb
Attribute VB_Name = "TestDeltaPick"
Option Explicit
' Parte B (1.5.0): modo "Delta clicando nas barras". Eventos REAIS por Points(i).Select (medido: dispara Chart_Select(3, serie, ponto));
' o caminho do MouseUp e chamado com pixels calculados (pt x 4/3 a 100%) depois de conferir GetChartElement.

Private Sub Teardown(wb As Workbook)
    DeltaPick_Disarm ""
    Application.StatusBar = False
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Tri", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5)
    ws.Range("A3:C3").Value = Array("T2", 20, 6)
    ws.Range("A4:C4").Value = Array("T3", 30, -4)
    Set Fixture = ws.Range("A1:C4")
End Function

Private Function ArmedChart(wb As Workbook, kind As String) As Chart
    Dim rng As Range, c As Chart
    Set rng = Fixture(wb)
    rng.Select
    Core_Insert kind
    Set c = Core_ActiveChart()
    c.Parent.Activate
    Core_DeltaPick
    Set ArmedChart = c
End Function

Private Function HasShape(c As Chart, nm As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = c.Shapes(nm)
    HasShape = Not s Is Nothing
End Function

Private Function StatusText() As String
    If VarType(Application.StatusBar) = vbBoolean Then StatusText = "" Else StatusText = CStr(Application.StatusBar)
End Function

Public Sub Test_DeltaPick_TwoClicksBySelectEvents()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    AssertTrue DeltaPick_IsArmedFor(c), "armado no grafico ativo"
    AssertEqual DELTAPICK_MSG_FROM, StatusText()
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(1).Select
    AssertEqual "1:1", DeltaPick_FromRef()
    AssertEqual DELTAPICK_MSG_TO, StatusText()
    AssertTrue HasShape(c, DELTAPICK_MARK), "moldura na origem"
    c.SeriesCollection(Delta_SeriesIndex(c, 2)).Points(3).Select
    AssertFalse DeltaPick_IsArmed(), "desarmado depois do destino"
    AssertTrue HasShape(c, "tqq_ann_delta_1")
    AssertEqual "kind=delta;v=1;from=1:1;to=2:3;fmt=both", c.Shapes("tqq_ann_delta_1").AlternativeText
    AssertFalse HasShape(c, DELTAPICK_MARK), "moldura removida"
    AssertEqual "", StatusText(), "barra de status devolvida"
    Teardown wb
End Sub

Public Sub Test_DeltaPick_SeriesSelectionAsksAgain_SameBar_NotBar()
    Dim wb As Workbook, c As Chart, s As Long, i As Long
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Select
    AssertEqual DELTAPICK_MSG_SERIES, StatusText(), "serie inteira (1o clique do Excel): pede o 2o clique"
    AssertEqual "", DeltaPick_FromRef()
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(2).Select
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(2).Select
    AssertEqual DELTAPICK_MSG_SAME, StatusText(), "mesma barra (pelo mesmo caminho) nao vira destino"
    AssertTrue DeltaPick_IsArmed()
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "stacked")
    For i = 1 To c.SeriesCollection.Count
        If Style_IsHelperSeries(c.SeriesCollection(i)) Then s = i: Exit For
    Next
    If s > 0 Then
        c.SeriesCollection(s).Points(1).Select
        AssertEqual DELTAPICK_MSG_NOTBAR, StatusText(), "auxiliar tqq_*: nao e barra"
        AssertEqual "", DeltaPick_FromRef()
    Else
        Errors_Log "TestDeltaPick: sem serie auxiliar na fixture empilhada (trecho pulado)"
    End If
    Teardown wb
End Sub

' 1o clique: MouseUp + GetChartElement acha o ponto; o mesmo clique chegando pelo Select nao conta de novo.
Public Sub Test_DeltaPick_MouseUpFindsPoint_AndDedupe()
    Dim wb As Workbook, c As Chart, pt As Point, x As Long, y As Long, id As Long, a1 As Long, a2 As Long, s As Long
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    ForceLayout c
    s = Delta_SeriesIndex(c, 2)
    Set pt = c.SeriesCollection(s).Points(2)
    x = CLng((pt.Left + pt.Width / 2) * 4 / 3): y = CLng((pt.Top + pt.Height / 2) * 4 / 3)
    c.GetChartElement x, y, id, a1, a2
    If id <> xlSeries Or a1 <> s Or a2 <> 2 Then
        Errors_Log "TestDeltaPick: GetChartElement(" & x & "," & y & ") deu " & id & "/" & a1 & "/" & a2 & " (pixels diferentes de pt x 4/3 aqui): trecho do MouseUp pulado"
    Else
        DeltaPick_OnMouseUp c, x, y
        AssertEqual "2:2", DeltaPick_FromRef(), "o MouseUp ja escolhe a barra"
        c.SeriesCollection(s).Points(2).Select
        AssertEqual DELTAPICK_MSG_TO, StatusText(), "o Select do mesmo clique nao vira 'mesma barra'"
    End If
    Teardown wb
End Sub

Public Sub Test_DeltaPick_CancelPaths()
    Dim wb As Workbook, c As Chart, c2 As Chart, rng As Range
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    Core_DeltaPick
    AssertFalse DeltaPick_IsArmed(), "o botao de novo cancela"
    AssertEqual DELTAPICK_MSG_CANCEL, StatusText()
    Core_DeltaPick
    DeltaPick_EscKey
    AssertFalse DeltaPick_IsArmed(), "Esc cancela"
    Core_DeltaPick
    DeltaPick_ExpireForTest
    DeltaPick_Timeout
    AssertFalse DeltaPick_IsArmed(), "60 s sem clique cancela"
    Watchers_Start
    c.Parent.Activate
    Core_DeltaPick
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(1).Select
    AssertTrue HasShape(c, DELTAPICK_MARK)
    wb.Worksheets(1).Range("H20").Select
    AssertFalse DeltaPick_IsArmed(), "clique numa celula cancela"
    AssertFalse HasShape(c, DELTAPICK_MARK), "moldura removida no cancelamento"
    Set rng = wb.Worksheets(1).Range("A1:C4")
    rng.Select
    Core_Insert "clustered"
    Set c2 = Core_ActiveChart()
    c.Parent.Activate
    Core_DeltaPick
    c2.SeriesCollection(1).Points(1).Select
    AssertFalse DeltaPick_IsArmed(), "clique noutro grafico cancela"
    Teardown wb
End Sub

Public Sub Test_DeltaPick_NotArmed_NoEffect_AndRefusals()
    Dim wb As Workbook, rng As Range, c As Chart, i As Long, t As Single, ws As Worksheet
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    c.Parent.Activate
    t = Timer
    For i = 1 To 30
        c.SeriesCollection(1).Points(1 + (i Mod 3)).Select
    Next
    AssertTrue Timer - t < 3, "desarmado: eventos baratos"
    AssertFalse HasShape(c, "tqq_ann_delta_1"), "desarmado: nada inserido"
    AssertEqual "", StatusText()
    Teardown wb
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = "Cat": ws.Range("B1").Value = "V": ws.Range("A2").Value = "a": ws.Range("B2").Value = 1: ws.Range("A3").Value = "b": ws.Range("B3").Value = 2
    ws.Range("A1:B3").Select
    Core_Insert "pie"
    Core_ClearLastNotice
    Core_DeltaPick
    AssertEqual ANN_DELTA_REFUSAL, Core_LastNotice()
    AssertFalse DeltaPick_IsArmed()
    Teardown wb
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "clustered"
    rng.Worksheet.Protect DrawingObjects:=True
    Core_ClearLastNotice
    Core_DeltaPick
    AssertEqual DELTAPICK_MSG_PROTECTED, Core_LastNotice()
    AssertFalse DeltaPick_IsArmed()
    rng.Worksheet.Unprotect
    Teardown wb
End Sub
```

- [ ] **Passo 2: RED** — `-Filter "TestDeltaPick"` → saída 3.

- [ ] **Passo 3: Implementar** — crie `src/core/modules/modDeltaPick.bas`:

```vb
Attribute VB_Name = "modDeltaPick"
Option Explicit
' Modo "Delta clicando nas barras" (Parte B, spec 2026-10-14 secao 5.1): o botao arma o modo no grafico ativo; o 1o clique numa barra
' escolhe a origem (moldura laranja temporaria tqq_pickmark), o 2o o destino e a seta e inserida (Core_AddDelta). Caminhos do clique:
' Chart.MouseUp + GetChartElement (acha o ponto ja no 1o clique) e Chart.Select (ElementID = xlSeries, Arg1 = serie, Arg2 = ponto; -1
' quando o Excel selecionou a serie inteira). O mesmo clique pelos dois caminhos conta uma vez. Cancela: Esc, clique numa celula, outro
' grafico, o botao de novo, 60 s sem clique. Desarmado, os manipuladores saem na primeira linha.
Public Const DELTAPICK_MARK As String = "tqq_pickmark"
Public Const DELTAPICK_MSG_FROM As String = "ThinQQ - Delta: clique na barra de ORIGEM (Esc ou um clique numa célula cancela)."
Public Const DELTAPICK_MSG_TO As String = "ThinQQ - Delta: agora clique na barra de DESTINO."
Public Const DELTAPICK_MSG_SAME As String = "ThinQQ - Delta: escolha uma barra diferente da origem."
Public Const DELTAPICK_MSG_NOTBAR As String = "ThinQQ - Delta: clique numa barra (não em rótulos, setas ou linhas auxiliares)."
Public Const DELTAPICK_MSG_SERIES As String = "ThinQQ - Delta: clique de novo na mesma barra para escolher só ela."
Public Const DELTAPICK_MSG_CANCEL As String = "ThinQQ - Delta cancelado."
Public Const DELTAPICK_MSG_PROTECTED As String = "A planilha está protegida. Desproteja a planilha para inserir o delta."
Private Const DELTAPICK_TIMEOUT_S As Long = 60
Private Const DELTAPICK_DEDUPE_S As Single = 0.6
Private mArmed As Boolean
Private mChartName As String                    ' ChartObject e planilha do grafico armado
Private mSheetName As String
Private mBookName As String
Private mFromRef As String
Private mLastKey As String                      ' ultimo clique tratado (serie:ponto), de onde veio e quando
Private mLastSrc As String
Private mLastAt As Single
Private mTimeoutAt As Date

Public Function DeltaPick_IsArmed() As Boolean
    DeltaPick_IsArmed = mArmed
End Function

Public Function DeltaPick_IsArmedFor(cht As Chart) As Boolean
    On Error Resume Next
    If Not mArmed Then Exit Function
    DeltaPick_IsArmedFor = (cht.Parent.Name = mChartName And cht.Parent.Parent.Name = mSheetName And cht.Parent.Parent.Parent.Name = mBookName)
End Function

Public Function DeltaPick_FromRef() As String
    DeltaPick_FromRef = mFromRef
End Function

Public Sub DeltaPick_Arm(cht As Chart)
    DeltaPick_Disarm ""
    mArmed = True
    mChartName = cht.Parent.Name: mSheetName = cht.Parent.Parent.Name: mBookName = cht.Parent.Parent.Parent.Name
    mFromRef = "": mLastKey = "": mLastSrc = ""
    Application.StatusBar = DELTAPICK_MSG_FROM
    On Error Resume Next
    Application.OnKey "{ESC}", "'" & ThisWorkbook.Name & "'!DeltaPick_EscKey"
    mTimeoutAt = Now + TimeSerial(0, 0, DELTAPICK_TIMEOUT_S)
    If BUILD_CONFIG <> "Test" Then Application.OnTime mTimeoutAt, "'" & ThisWorkbook.Name & "'!DeltaPick_Timeout"
    Err.Clear
End Sub

' Desarma (msg "" = barra de status devolvida ao Excel; senao mostra msg). Remove a moldura, devolve o Esc e cancela o tempo. Nunca propaga.
Public Sub DeltaPick_Disarm(msg As String)
    Dim co As ChartObject
    On Error Resume Next
    If mArmed Then
        Set co = Workbooks(mBookName).Worksheets(mSheetName).ChartObjects(mChartName)
        If Not co Is Nothing Then DeltaPick_RemoveMark co.Chart
        Application.OnKey "{ESC}"
        If BUILD_CONFIG <> "Test" Then Application.OnTime mTimeoutAt, "'" & ThisWorkbook.Name & "'!DeltaPick_Timeout", , False
    End If
    mArmed = False
    mFromRef = "": mLastKey = "": mLastSrc = ""
    If Len(msg) > 0 Then Application.StatusBar = msg Else Application.StatusBar = False
    If Len(msg) > 0 And BUILD_CONFIG <> "Test" Then Application.OnTime Now + TimeSerial(0, 0, 4), "'" & ThisWorkbook.Name & "'!DeltaPick_ClearStatus"
    Err.Clear
End Sub

Public Sub DeltaPick_ClearStatus()
    On Error Resume Next
    If Not mArmed Then Application.StatusBar = False
End Sub

Public Sub DeltaPick_EscKey()
    On Error Resume Next
    If mArmed Then DeltaPick_Disarm DELTAPICK_MSG_CANCEL
End Sub

Public Sub DeltaPick_Timeout()
    On Error Resume Next
    If mArmed And Now >= mTimeoutAt Then DeltaPick_Disarm "ThinQQ - Delta cancelado (60 s sem clique)."
End Sub

' Gancho de teste: faz o prazo do modo armado vencer agora. Inerte fora do build Test.
Public Sub DeltaPick_ExpireForTest()
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTimeoutAt = Now - TimeSerial(0, 0, 1)
End Sub

' Clique numa celula (App_SheetSelectionChange).
Public Sub DeltaPick_OnSheetClick()
    On Error Resume Next
    If mArmed Then DeltaPick_Disarm DELTAPICK_MSG_CANCEL
End Sub

' Chart.Select de QUALQUER grafico observado. Outro grafico cancela; area/eixo/legenda/seta: segue armado.
Public Sub DeltaPick_OnSelect(cht As Chart, elementId As Long, a1 As Long, a2 As Long)
    If Not mArmed Then Exit Sub
    On Error Resume Next
    If Not DeltaPick_IsArmedFor(cht) Then DeltaPick_Disarm "ThinQQ - Delta cancelado (outro gráfico).": Exit Sub
    If elementId <> xlSeries Then Exit Sub
    If a2 < 1 Then
        If Not DeltaPick_IsDuplicate(a1 & ":", "select") Then Application.StatusBar = DELTAPICK_MSG_SERIES
        Exit Sub
    End If
    DeltaPick_Take cht, a1, a2, "select"
End Sub

' Chart.MouseUp: acha o elemento sob o ponteiro (coordenadas do proprio evento, as mesmas do GetChartElement).
Public Sub DeltaPick_OnMouseUp(cht As Chart, x As Long, y As Long)
    Dim id As Long, a1 As Long, a2 As Long
    If Not mArmed Then Exit Sub
    On Error Resume Next
    If Not DeltaPick_IsArmedFor(cht) Then Exit Sub
    cht.GetChartElement x, y, id, a1, a2
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    If id = xlSeries And a2 >= 1 Then DeltaPick_Take cht, a1, a2, "mouse"
End Sub

' O mesmo clique chega pelos dois caminhos (MouseUp e Select): mesma chave, OUTRA fonte, em menos de 0,6 s = repetido.
Private Function DeltaPick_IsDuplicate(key As String, src As String) As Boolean
    Dim t As Single
    t = Timer
    If Len(mLastKey) > 0 And Left$(mLastKey, Len(key)) = key And src <> mLastSrc And Abs(t - mLastAt) < DELTAPICK_DEDUPE_S Then DeltaPick_IsDuplicate = True
End Function

Private Sub DeltaPick_Take(cht As Chart, s As Long, p As Long, src As String)
    Dim ref As String, key As String, fromRef As String
    key = s & ":" & p
    If key = mLastKey And src <> mLastSrc And Abs(Timer - mLastAt) < DELTAPICK_DEDUPE_S Then Exit Sub
    mLastKey = key: mLastSrc = src: mLastAt = Timer
    If Not Delta_ResolvePick(cht, s, p, ref) Then Application.StatusBar = DELTAPICK_MSG_NOTBAR: Exit Sub
    If Len(mFromRef) = 0 Then
        mFromRef = ref
        DeltaPick_DrawMark cht, ref
        Application.StatusBar = DELTAPICK_MSG_TO
        Exit Sub
    End If
    If ref = mFromRef Then Application.StatusBar = DELTAPICK_MSG_SAME: Exit Sub
    fromRef = mFromRef
    DeltaPick_Disarm ""
    Core_AddDelta fromRef, ref, "both", cht
End Sub

' Moldura laranja de 2 pt em volta da barra de origem (cosmetica: falha so registra).
Private Sub DeltaPick_DrawMark(cht As Chart, ref As String)
    Dim v As Double, lo As Double, hi As Double, c As Double, half As Double, a As Double, b As Double, m As Shape, lft As Double, tp As Double
    On Error GoTo Done
    DeltaPick_RemoveMark cht
    If Not Delta_Extent(cht, ref, v, lo, hi) Then Exit Sub
    c = Delta_CatCenter(cht, ref)
    half = Delta_BarHalf(cht, ref) + 2
    a = Geom_ChartValCoord(cht, hi): b = Geom_ChartValCoord(cht, lo)
    If a < b Then lft = a Else lft = b
    If Geom_ChartIsHorizontal(cht) Then
        Set m = cht.Shapes.AddShape(msoShapeRectangle, lft - 2, c - half, Abs(b - a) + 4, 2 * half)
    Else
        Set m = cht.Shapes.AddShape(msoShapeRectangle, c - half, lft - 2, 2 * half, Abs(b - a) + 4)
    End If
    m.Name = DELTAPICK_MARK
    m.AlternativeText = "kind=pickmark"
    m.Fill.Visible = msoFalse
    m.Line.ForeColor.RGB = RGB(237, 125, 49)
    m.Line.Weight = 2
    Exit Sub
Done:
    Errors_Log "DeltaPick_DrawMark" & vbTab & Err.Number & vbTab & Err.Description
End Sub

Public Sub DeltaPick_RemoveMark(cht As Chart)
    On Error Resume Next
    cht.Shapes(DELTAPICK_MARK).Delete
    Err.Clear
End Sub
```
(Na `DeltaPick_OnSelect`, o `DeltaPick_IsDuplicate(a1 & ":", "select")` evita sobrescrever o status quando o mesmo clique já escolheu o ponto pelo `MouseUp`.)

`src/core/classes/clsChartWatcher.cls`, depois de `Cht_Resize`:

```vb
' Delta clicando nas barras (Parte B): so age com o modo armado (modDeltaPick sai na primeira linha quando desarmado).
Private Sub Cht_Select(ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long)
    On Error Resume Next
    DeltaPick_OnSelect Cht, ElementID, Arg1, Arg2
End Sub

Private Sub Cht_MouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    On Error Resume Next
    DeltaPick_OnMouseUp Cht, x, y
End Sub
```
`src/core/classes/clsAppEvents.cls`, em `App_SheetSelectionChange`, antes de `Watchers_NotifySelection`: `DeltaPick_OnSheetClick`.

`modAnnotations.Ann_DeltaCleanup`: acrescente, no início do laço (antes do `If Left$(nm, Len(ANN_DELTA))`), o ramo da moldura:

```vb
        If nm = DELTAPICK_MARK Then
            If Not DeltaPick_IsArmedFor(cht) Then cht.Shapes(i).Delete
        ElseIf Left$(nm, Len(ANN_DELTA)) = ANN_DELTA Then
```
(e ajuste o `End If` correspondente). `Ann_DeltaCleanup` precisa rodar mesmo sem setas de delta (a moldura pode estar sozinha): ela já é a primeira linha de `Ann_LayoutDeltas`, chamada em todo `Ann_Refresh`.

`modCoreApi`, depois de `Core_AddDelta`:

```vb
' Botao "Delta clicando nas barras": arma o modo de dois cliques no grafico ativo (ou cancela, se ja estava armado nele). Nunca propaga.
Public Sub Core_DeltaPick()
    Dim c As Chart, refusal As String
    On Error GoTo Fail
    Set c = Core_ActiveChart()
    If c Is Nothing Then Core_Notify "Selecione um gráfico do ThinQQ antes de inserir o delta.", vbInformation: Exit Sub
    If DeltaPick_IsArmedFor(c) Then DeltaPick_Disarm DELTAPICK_MSG_CANCEL: Exit Sub
    refusal = Ann_RefusalText(c, "delta")
    If Len(refusal) > 0 Then Core_Notify refusal, vbInformation: Exit Sub
    If c.Parent.Parent.ProtectDrawingObjects Then Core_Notify DELTAPICK_MSG_PROTECTED, vbExclamation: Exit Sub
    DeltaPick_Arm c
    Exit Sub
Fail:
    Errors_Handle "Core_DeltaPick"
End Sub
```
Em `Core_OnUnload`, acrescente `DeltaPick_Disarm ""` (o Esc do Excel nunca fica preso ao core descarregado).
- [ ] **Passo 4: GREEN** — `-Filter "TestDeltaPick|TestDelta|TestCoreApi"` → 0 falhas. Relate se o trecho do `MouseUp` foi pulado (log "pixels diferentes").
- [ ] **Passo 5: Suíte completa** e **Commit** — `git commit -m "feat: delta clicando nas barras (modo de dois cliques pelos eventos do grafico)" (+ trailer)`.

---

### Tarefa B4: Alternar o texto do delta e a faixa — modelo: standard

**Files:**
- Modify: `src/core/modules/modAnnotations.bas` (`Ann_CycleDeltaFormat`, `Ann_SelectedDeltaName`), `src/core/modules/modCoreApi.bas` (`Core_DeltaCycleFormat`)
- Modify: `src/loader/ribbon/customUI14.xml` (menu Setas), `src/loader/modules/modRibbon.bas` (`Rx_OnDelta`, `Rx_OnDeltaFormat`)
- Modify: `docs/tutorial/tutorial.html` (`#setas` mínimo; matriz `k` → `ui`), `tests/vba/loader/TestRibbon.bas`
- Test: `tests/vba/core/TestDeltaPick.bas` (acrescentar)

**Interfaces:**
- Consumes (B2/B3): `Delta_NextFormat`, `Ann_SetParam`, `Ann_Refresh`, `Ann_EnsureDeltaRoom`, `Core_DeltaPick`.
- Produces: `Public Function Ann_CycleDeltaFormat(cht As Chart, onlyName As String) As Long`; `Public Function Ann_SelectedDeltaName() As String`; `Public Sub Core_DeltaCycleFormat()`; botões `btnDelta` "Delta clicando nas barras" e `btnDeltaFmt` "Delta: valor / % / ambos" (rótulos exatos usados na B5).

- [ ] **Passo 1: Testes (RED)** — ao fim de `tests/vba/core/TestDeltaPick.bas`:

```vb
Public Sub Test_Delta_CycleFormat_AllOrSelected()
    Dim wb As Workbook, rng As Range, c As Chart
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    Core_AddDelta "1:1", "1:3"
    Core_AddDelta "2:1", "2:2"
    c.Parent.Activate
    c.ChartArea.Select
    Core_DeltaCycleFormat
    AssertEqual "abs", Ann_ParseParam(c.Shapes("tqq_ann_delta_1").AlternativeText, "fmt"), "sem seta selecionada: todas"
    AssertEqual "abs", Ann_ParseParam(c.Shapes("tqq_ann_delta_2").AlternativeText, "fmt")
    AssertEqual Delta_Text(10, 30, "abs", Style_NumberFormat(Registry_GetOption(c, "numfmt", "int"))), CStr(c.Shapes("tqq_ann_delta_1_lbl").TextFrame2.TextRange.Text)
    c.Shapes("tqq_ann_delta_2_lbl").Select
    AssertEqual "tqq_ann_delta_2", Ann_SelectedDeltaName()
    Core_DeltaCycleFormat
    AssertEqual "abs", Ann_ParseParam(c.Shapes("tqq_ann_delta_1").AlternativeText, "fmt"), "so a selecionada muda"
    AssertEqual "pct", Ann_ParseParam(c.Shapes("tqq_ann_delta_2").AlternativeText, "fmt")
    Ann_RemoveAll c
    c.ChartArea.Select
    Core_ClearLastNotice
    Core_DeltaCycleFormat
    AssertEqual DELTA_NONE, Core_LastNotice(), "sem delta: aviso"
    Teardown wb
End Sub
```
Em `tests/vba/loader/TestRibbon.bas`, acrescente:

```vb
Public Sub Test_Ribbon_Delta_LabelsAndTips()
    Dim d As Object, n As Object, m As Object
    Set d = LoadRibbon()
    Set n = d.SelectSingleNode("//u:menu[@id='mnuArrows']/u:button[@id='btnDelta']")
    AssertTrue Not n Is Nothing, "btnDelta no menu Setas"
    AssertEqual "Delta clicando nas barras", AttrText(n, "label"): AssertEqual "delta", AttrText(n, "tag")
    AssertEqual "Rx_OnDelta", AttrText(n, "onAction"): AssertEqual "Rx_GetEnabled", AttrText(n, "getEnabled")
    AssertEqual "btnDelta", AttrText(d.SelectSingleNode("//u:button[@id='btnDiff']/following-sibling::*[1]"), "id"), "logo depois da Seta diferenca"
    Set m = d.SelectSingleNode("//u:menu[@id='mnuArrows']/u:button[@id='btnDeltaFmt']")
    AssertTrue Not m Is Nothing
    AssertEqual "Delta: valor / % / ambos", AttrText(m, "label"): AssertEqual "delta", AttrText(m, "tag"): AssertEqual "Rx_OnDeltaFormat", AttrText(m, "onAction")
    AssertTrue InStr(AttrText(n, "supertip"), "Esc") > 0, "supertip ensina a cancelar"
    AssertEqual 3, d.SelectNodes("//u:button[@onAction='Rx_OnAnnotation']").Length, "as 3 setas antigas continuam"
End Sub
```

- [ ] **Passo 2: RED** — core `-Filter "TestDeltaPick"` (saída 3) e loader `-Filter "TestRibbon"`.
- [ ] **Passo 3: Core** — em `modAnnotations.bas` (depois de `Ann_AddDelta`):

```vb
' Alterna o texto (ambos -> valor -> % -> ambos) da seta onlyName ou, com "", de todas as setas de delta do grafico. Devolve quantas mudou.
Public Function Ann_CycleDeltaFormat(cht As Chart, onlyName As String) As Long
    Dim i As Long, shp As Shape
    For i = 1 To cht.Shapes.Count
        Set shp = cht.Shapes(i)
        If Left$(shp.Name, Len(ANN_DELTA)) = ANN_DELTA Then
            If Ann_ParseParam(shp.AlternativeText, "kind") = "delta" And (Len(onlyName) = 0 Or shp.Name = onlyName) Then
                shp.AlternativeText = Ann_SetParam(shp.AlternativeText, "fmt", Delta_NextFormat(Ann_ParseParam(shp.AlternativeText, "fmt")))
                Ann_CycleDeltaFormat = Ann_CycleDeltaFormat + 1
            End If
        End If
    Next
    If Ann_CycleDeltaFormat > 0 Then
        Ann_Refresh cht
        Ann_EnsureDeltaRoom cht
    End If
End Function

' Nome da seta de delta selecionada (linha, perna ou rotulo dentro do grafico ativo); "" se a selecao nao e uma delas.
Public Function Ann_SelectedDeltaName() As String
    Dim nm As String, sfx As Variant
    On Error Resume Next
    nm = CStr(Selection.Name)
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    If Left$(nm, Len(ANN_DELTA)) <> ANN_DELTA Then Exit Function
    For Each sfx In Array("_lbl", "_t1", "_t2")
        If Right$(nm, Len(sfx)) = sfx Then nm = Left$(nm, Len(nm) - Len(sfx)): Exit For
    Next
    Ann_SelectedDeltaName = nm
End Function
```
Em `modDelta.bas`: `Public Const DELTA_NONE As String = "Este gráfico não tem delta entre barras. Use Setas > Delta clicando nas barras."`. Em `modCoreApi.bas`, depois de `Core_DeltaPick`:

```vb
' Botao "Delta: valor / % / ambos": alterna o texto da seta de delta selecionada (ou de todas do grafico). Nunca propaga.
Public Sub Core_DeltaCycleFormat()
    Dim c As Chart
    On Error GoTo Fail
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    If Ann_CycleDeltaFormat(c, Ann_SelectedDeltaName()) = 0 Then Core_Notify DELTA_NONE, vbInformation
    Exit Sub
Fail:
    Errors_Handle "Core_DeltaCycleFormat"
End Sub
```
- [ ] **Passo 4: Faixa** — em `src/loader/ribbon/customUI14.xml`, menu `mnuArrows`: troque o supertip por `"Seta de CAGR e de diferença entre duas categorias; delta entre duas barras escolhidas com o mouse; colchete sobre atividades do Gantt; setas de variação da cascata comparativa."` e acrescente logo depois de `btnDiff`:

```xml
            <button id="btnDelta" label="Delta clicando nas barras" image="tqq_arrows16" tag="delta" onAction="Rx_OnDelta"
                    getEnabled="Rx_GetEnabled" screentip="Delta clicando nas barras"
                    supertip="Clique aqui e depois em duas barras do gráfico (origem e destino): a seta mostra a diferença em valor e em % entre elas. Esc ou um clique numa célula cancela."/>
            <button id="btnDeltaFmt" label="Delta: valor / % / ambos" image="tqq_arrows16" tag="delta" onAction="Rx_OnDeltaFormat"
                    getEnabled="Rx_GetEnabled" screentip="Texto do delta"
                    supertip="Alterna o texto da seta de delta selecionada (ou de todas do gráfico, sem nenhuma selecionada): valor e %, só o valor, só o %."/>
```
`src/loader/modules/modRibbon.bas`, depois de `Rx_OnAnnotation`:

```vb
' Delta clicando nas barras (Parte B): arma o modo de dois cliques no core; core anterior: aviso de atualizar.
Public Sub Rx_OnDelta(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_DeltaPick"
End Sub

Public Sub Rx_OnDeltaFormat(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_DeltaCycleFormat"
End Sub
```
- [ ] **Passo 5: Tutorial mínimo** — na seção `id="setas"`, acrescente uma subseção `<h3 id="delta">Delta entre barras</h3>` com 3 frases: <span class="ui">Setas</span> › <span class="ui">Delta clicando nas barras</span>, clique na barra de origem e depois na de destino (a barra de status orienta; Esc ou um clique numa célula cancela); <span class="ui">Delta: valor / % / ambos</span> alterna o texto. Na 5.ª tabela da matriz (B0), troque `<span class="k">Delta clicando nas barras</span>` por `<span class="ui">Delta clicando nas barras</span>`.
- [ ] **Passo 6: GREEN + suíte completa** (Test-Ribbon, Test-Docs) e **Commit** — `git commit -m "feat: faixa do delta entre barras (inserir clicando e alternar valor/%/ambos)" (+ trailer)`.

---

### Tarefa B5: Exemplos, demonstração, galeria e documentação do delta — modelo: standard

**Files:**
- Modify: `src/core/modules/modSamples.bas` (exemplo "3) Colunas agrupadas"), `tests/vba/core/TestSamples.bas`
- Modify: `tools/DemoCoverage.ps1` (E02 com `delta`; lista de recursos), `tools/New-DemoWorkbook.ps1` (E02 e rótulo do recurso), `tests/ps/Test-DemoWorkbook.ps1` (mapa `delta`; remove a pendência da B0, se criada)
- Regenerate: `exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx`, imagens da galeria do tutorial (`tools/Export-DemoChartImages.ps1`, `tools/Build-TutorialGallery.ps1`)
- Modify: `docs/tutorial/tutorial.html` (`#delta` completa, mensagens, limites), `README.md`, `CHANGELOG.md`, `release/notes.txt`, `tests/ps/Test-Docs.ps1`

**Interfaces:**
- Consumes: `Ann_AddDelta`, `Core_AddDelta` (via `Invoke-CoreQuiet` do gerador), rótulos da B4, mensagens `ANN_DELTA_REFUSAL`, `DELTA_BAD_PICK`, `DELTA_NONE`, `DELTAPICK_MSG_*`.

- [ ] **Passo 1: Exemplos (RED → GREEN)** — em `TestSamples.bas`, no teste que confere a planilha "Exemplos 3", acrescente que o gráfico do item "3) Colunas agrupadas" (o primeiro `clustered` dessa planilha) tem a forma `tqq_ann_delta_1` com `AlternativeText = "kind=delta;v=1;from=1:1;to=2:4;fmt=both"`. Em `modSamples.bas`, troque `Stacked_Insert ws, d, "clustered"` (logo depois do `WriteTitle ws, 43, "3) Colunas agrupadas"`) por `Set co = Stacked_Insert(ws, d, "clustered")` e acrescente `Ann_AddDelta co.Chart, "1:1", "2:4"       ' Loja T1 -> Online T4: delta entre series e trimestres diferentes` (confira que `Stacked_Insert` é `Function ... As ChartObject`; se for `Sub`, use `Set co = ws.ChartObjects(ws.ChartObjects.Count)` logo depois). As contagens de exemplos não mudam.
- [ ] **Passo 2: Demonstração** — `tools/DemoCoverage.ps1`: E02 `Features = @('gap', 'delta')` e `'delta'` na lista de `Get-TqqDemoFeatures` (depois de `'diff'`); `tools/New-DemoWorkbook.ps1`, no `2 { ... }` do E02, depois dos `Core_SetOption`: `Invoke-CoreQuiet 'E02' 'Core_AddDelta' '1:1' ('2:' + $storeYearMiBody.Count) | Out-Null       # 2024 da 1a loja -> 2025 da ultima: delta entre series e lojas diferentes` e, no mapa de rótulos de recurso (linha ~454), `delta = 'Delta entre barras'`; `tests/ps/Test-DemoWorkbook.ps1`: no mapa de verificações, `'delta' = { param($i, $co, $ws) Test-HasShape $co 'tqq_ann_delta_' }` e remova a pendência temporária da B0. Regere: `pwsh <repo>\tools\New-DemoWorkbook.ps1` (num Excel próprio, como na cascata comparativa), depois `tools/Export-DemoChartImages.ps1` e `tools/Build-TutorialGallery.ps1` (siga os parâmetros que o cabeçalho de cada script documenta). Confira `Test-DemoWorkbook.ps1` verde.
- [ ] **Passo 3: Documentação** — tutorial `#delta` completo: o fluxo (botão, barra de status, moldura laranja, 1.º/2.º clique, "clique de novo na mesma barra" quando o Excel seleciona a série inteira), o que conta como barra em cada tipo (tabela da spec §5.2), o texto (valor/%/ambos; `(B − A) / |A|`; n/d com origem zero; a diferença para as setas de variação da cascata comparativa com o exemplo −4 → −7: −75% aqui, +75% lá), faixas automáticas, a seta não é arrastável (o refresh a recoloca), Alterar dados não preserva setas, gráfico em grupo de formas não responde aos cliques, no PowerPoint o delta é atualizado pelo suplemento (Parte A); mensagens novas na tabela de mensagens (`ANN_DELTA_REFUSAL`, `DELTA_BAD_PICK`, `DELTA_NONE`, `DELTAPICK_MSG_PROTECTED`, `"Selecione um gráfico do ThinQQ antes de inserir o delta."`). README: Recursos (menu Setas), Verificações manuais (itens novos: cliques reais nos 9 tipos, 1.º clique num gráfico não ativado, Esc com o gráfico ativo, zoom 75%/150%, monitor a 125%, delta no PowerPoint), Limitações. CHANGELOG `[Não lançado]` `### Adicionado`: "Delta entre barras: Setas > Delta clicando nas barras insere uma seta com a diferença (valor e %) entre quaisquer duas barras, escolhidas com dois cliques; Delta: valor / % / ambos alterna o texto.". `release/notes.txt`: "• Novo: delta entre quaisquer duas barras, escolhidas clicando nelas (menu Setas)." `Test-Docs.ps1`: mensagens novas no `$messages`; limites no `mustDocument`: `'delta: seta nao arrastavel' = @('tutorial', 'não é arrastável')`, `'delta: grafico em grupo de formas' = @('tutorial', 'grupo de formas')`, `'delta: percentual diferente da seta de variacao' = @('tutorial', '-75%')`.
- [ ] **Passo 4: GREEN + suíte completa** e **Commit** (dois commits: `feat: delta no exemplo de colunas agrupadas e na demonstracao (E02)`; `docs: tutorial, README e CHANGELOG do delta entre barras`) (+ trailer).

---

### Tarefa B6: Suíte final, revisão do ramo e verificação manual (humano) — modelo: standard (controlador + humano)

- [ ] **Passo 1:** `pwsh C:\temp\tqq-pd\build\run-tests.ps1 -TimeoutSec 900 -LoaderSwap` (com `-Ppt` se a Parte A já estiver na `main` e nenhum PowerPoint aberto) → "Todos os testes passaram".
- [ ] **Passo 2:** revisão do ramo inteiro (modelo mais capaz) + uma onda de correção + revisão de escopo.
- [ ] **Passo 3 (humano):** no Excel visível com o ThinQQ instalado a partir do build Release: para cada um dos 9 tipos, inserir um delta com o mouse (inclusive o 1.º clique num gráfico **não** selecionado, séries diferentes na mesma categoria, categorias diferentes, barras negativas); Esc com o gráfico ativo; clique numa célula; zoom 75% e 150%; monitor a 125% (se houver); "Delta: valor / % / ambos" com uma seta selecionada; copiar como gráfico vinculado para o PowerPoint e reabrir (Parte A).
- [ ] **Passo 4:** PR, merge, `publicar-release.bat 1.5.0`.

---

## Resumo da estimativa (detalhe e suposições: spec §12)

| Parte | Tarefas | Agêntico corrido (otimista / provável / pessimista) | Humano (otimista / provável / pessimista) | Confiança |
|---|---|---|---|---|
| A — PowerPoint atualizado ao reabrir + suplemento atualizado pelo Excel (1.4.0) | A0–A9 (10) | ~8,5 / 13 / 19,5 h | 2,25 / 4 / 6,25 h | média-baixa |
| B — Delta clicando nas barras (1.5.0) | B0–B6 (7) | ~6 / 9,5 / 14 h | 1,5 / 2,5 / 4 h | média-alta |
| **Total** | **17** | **~14,5 / 22,5 / 33,5 h (2 / 3 / 4–5 dias de trabalho agêntico)** | **3,75 / 6,5 / 10,25 h** | |

A emenda "atualização automática do suplemento pelo Excel" (Tarefa A7 nova; A2, A4, A5, A6, A8, A9 ampliadas) acrescentou à Parte A **+75 / +120 / +180 min** agênticos (≈ +1,5 / +2 / +2,5 h corridas) e **+15 / +30 / +45 min** humanos.

Maiores incertezas: montar o `.ppam` por automação do PowerPoint e os eventos ao abrir pela interface/Modo de Exibição Protegido (Parte A); `MouseUp` no clique que ativa o gráfico, `OnKey` com o gráfico ativo e DPI acima de 100% (Parte B).
