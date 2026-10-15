# ThinQQ — Plano de Implementação: paridade com o think-cell na integração Excel ↔ PowerPoint (1.6.0 a 1.10.0)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** (1.6.0, Parte C) elementos do ThinQQ no PowerPoint com **"Atualizar automaticamente"** (por elemento, como o "Update Automatically" do think-cell): com o Excel e o PowerPoint abertos, mudar a pasta atualiza dados e setas no slide em ~2–3 s; "Atualizar uma vez" só nos selecionados. (1.7.0, Parte D) **"Inserir também no slide"**: inserir qualquer tipo a partir do intervalo selecionado já cola o gráfico vinculado no slide, com uma **moldura verde** no intervalo de origem. (1.8.0, Parte G-I — emenda 1, manchete) **criar o gráfico pelo PowerPoint**, como o "Insert > Elements" do think-cell: aba ThinQQ do PowerPoint > "Inserir gráfico" com os 21 tipos; o gráfico nasce no slide, vinculado a uma **pasta de dados** ao lado da apresentação (o Excel é o datasheet), automático, com "Editar dados". (1.9.0, Partes G-II + E) **comandos do gráfico pelo PowerPoint** (rótulos, totais, conectores, formato, legenda, CAGR, diferença, remover setas, empacotar) e **texto vinculado a célula** (caixa de texto do slide; título do gráfico). (1.10.0, Parte F) **tabelas do PowerPoint vinculadas** a intervalos do Excel.

**Architecture:** a mesma de 2026-10-14: `ThinQQ.xlam` (loader, faixa) chama `Core_*` em `ThinQQ.core.xlam`; o motor `modPptSync` sincroniza formas marcadas (`Shape.Tags` `TQQ_*`) usando a pasta de origem aberta (macros desligadas); `ThinQQ.ppam` (fonte `src/ppt/`) chama o core por `Application.Run`. Novidades: índice das formas automáticas na etiqueta da apresentação (`TQQ_AUTOIDX`), `PptSync_Items` com rota por tipo de elemento (gráfico, texto, tabela), `modPptAuto` (eventos do Excel + `OnTime`), `modLinkFrame` (moldura), `modPptData` (pasta de dados dos gráficos criados pelo PowerPoint), `modPptText`, `modPptTable`.

**Tech Stack:** VBA 7 (Excel e PowerPoint 2016+ Windows, 32/64 bits, sem `Declare`, sem referências extras), customUI14, PowerShell 7 e 5.1, Git.

**Spec:** `docs/superpowers/specs/2026-10-15-thinqq-thinkcell-parity-design.md` (cópia, feita na Tarefa C0, do rascunho `.superpowers/sdd/2026-10-15-thinkcell-parity/design.md`: pesquisa §2, sondas §3 e **pendentes §3.1**, decisões §5, matriz §6, faixas §7, ordem §8, riscos §9, pré-flight §11, estimativa §12). Spec anterior: `docs/superpowers/specs/2026-10-14-thinqq-ppt-delta-design.md`. Regras de ambiente: `.superpowers/sdd/2026-10-14-ppt-delta/implementer-rules.md` (valem aqui sem exceção, inclusive a seção do EDR).

**Estado de execução:** nenhuma tarefa concluída. Cinco partes em cinco versões, cada versão numa branch criada pelo controlador a partir da `main` (`feat/ppt-auto` 1.6.0, `feat/ppt-inserir` 1.7.0, `feat/ppt-criar` 1.8.0 (fase G-I: tarefas G1–G6), `feat/ppt-comandos-texto` 1.9.0 (fase G-II: tarefas G7–G11, depois E1–E5), `feat/ppt-tabela` 1.10.0); D pode começar em paralelo com a verificação manual da C (arquivos diferentes, salvo `modCoreApi.bas`/faixa: D parte da ponta da C se a C já estiver mesclada). O implementador **não** cria nem troca de branch. **Antes da C2:** o controlador agenda as sondas pendentes P1 e P3 do spec §3.1 (P3 é manual) e anota o resultado no ledger; a C2 traz o caminho alternativo já descrito.

**Ordem e por quê:** C (maior valor e infraestrutura: índice, `PptSync_Items`, API 2, motor de eventos) → D (pequena, independente; a moldura D1 serve à G) → G-I (manchete pedida na emenda 1; precisa do automático da C para o slide acompanhar a digitação e da moldura D1) → G-II + E (G-II reaproveita a pasta de dados; E introduz o modelo "nome definido + texto exibido + elemento não-gráfico" que a F reaproveita) → F (maior volume de formatação). Alternativa com pressa pela manchete: 1.6.0 = C + D e 1.7.0 = G-I (mesmo esforço, um release a menos). Detalhe: spec §8 e §14. **API do core para o `.ppam`:** 2 = C1, 3 = G2/G3 (fase G-I), 4 = G7 (fase G-II), 5 = E1.

## Global Constraints

- **Ambiente:** Excel 2019 32-bit pt-BR e PowerPoint 16.0 32-bit no Windows 11; testes em instâncias próprias invisíveis. Build roda o lint (`build/lib/Vba-Lint.ps1`) e compilação completa; compilação ⇒ saída 3 (módulo:linha); travamento ⇒ 2; asserts ⇒ 1; ok ⇒ 0. Fontes VBA em UTF-8 com **somente caracteres cp1252** (nada de `→`, `≥`, `−`, `…`, aspas tipográficas), `Option Explicit`, declarações antes dos procedimentos, `Attribute VB_Name` na 1ª linha, nenhuma continuação ` _` seguida de texto, nada de `ElseIf` em `If` de uma linha nem `If ... Then If ...`.
- **EDR:** NUNCA iniciar processo filho a partir de VBA (`Shell`, `WScript.Shell.Run/Exec`, `powershell`, `cmd`): o SentinelOne mata o Excel e põe arquivos em quarentena; `Test-VbaLint` barra. Excel morreu (0x800706BE/0x800A9C68)? Olhe primeiro `Microsoft-Windows-SentinelOne/Operational` (eventos 31/32).
- **Testes:** `Public Sub Test_x()` sem parâmetros em `tests/vba/{core,loader,ppt}/Test*.bas`, via `pwsh build/run-tests.ps1` (`-Target core|loader|all`, `-Filter "Modulo.Teste|outro"`, `-LoaderSwap`, `-Ppt`, `-TimeoutSec 1200`). Asserts: `AssertTrue/False/Equal/Almost`, `NewTempWorkbook`, `CloseTemp`, `TempDir`, `ForceLayout`. Testes de PowerPoint real só rodam sem nenhum `POWERPNT` (helpers `StartScratchPowerPoint`, `UserWindowAppeared`, `StopScratchPowerPoint`, `PasteLinked`, `PasteAgain`, `PptSeriesValue`, `Fixture`, `Teardown` de `tests/vba/core/TestPptSync.bas`: os testes novos de PowerPoint da C entram NESSE módulo para reaproveitá-los).
- **Instâncias do usuário:** anote `Get-Process excel,powerpnt -ErrorAction SilentlyContinue | Select-Object Id, StartTime` ao começar; esses processos nunca são tocados, anexados (`GetObject`) nem encerrados. Com um PowerPoint do usuário aberto os testes reais saem verdes registrando "pulado" e o relatório diz isso. Builds e testes numa **CÓPIA**: script `.ps1` com `robocopy "<repo>" C:\temp\tqq-pd /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS` (rodado por `pwsh -File`; nunca `/MIR` em direção ao repo). Nunca `build.ps1 -Install`, `%LOCALAPPDATA%\ThinQQ`, registro real, lojas de certificados. Apague `C:\temp\tqq-pd*` ao terminar (`rm -rf` do bash, caminho exato). Nunca toque `C:\temp\tqq-video*`. `git status --short` ao começar e terminar (quarentena: `git checkout -- <arquivo>` e relate).
- **Código:** `Core_*` nunca propaga erro (`Errors_Handle`/`Errors_HandleErr`; `On Error GoTo -1` antes de `Resume Next` dentro de tratador); `MsgBox` só em `Core_Notify`/`Core_Confirm`/`RxMsg`/`Loader_Msg`/`PptTqq_Msg`/`PptTqq_Ask` (log-only no Test); ganchos de teste só com `BUILD_CONFIG = "Test"`, zerados em todo caminho; ligação tardia com o outro aplicativo (`As Object`, constantes numéricas); registro (`GetSetting/SaveSetting`) nunca no build Test (memória). Pastas de origem abertas pelo motor: só por `PptSync_OpenBook` (somente leitura, `AutomationSecurity = 3`, eventos desligados). O motor só usa um PowerPoint **já aberto** (`Ppt_GetApp(False)`).
- **Armadilhas medidas (spec §3 e 2026-10-14 §3; não redescubra):** `LinkFormat.Update` com origem inexistente trava o PowerPoint (confira antes); com a pasta aberta lê valores não salvos (~60 ms); `Tags.Item` de chave inexistente devolve "" sem erro; varrer formas custa ~6,4 ms por forma de fora do processo (use o índice + `FindBySlideID`, ~5 ms); `ChartTitle.Formula` aceita `=Planilha!$E$1` e segue pelo vínculo; `AddTable` + preencher 55 células ~1,5 s, trocar texto mantém a formatação da célula, `Rows.Add` copia a formatação da última linha; retângulo com `Placement = 1` acompanha linhas inseridas; hiperlink NÃO serve de marcador de campo; **nunca chame `ChartData.BreakLink`** (derruba o PowerPoint ao salvar); `Application.Visible = False` é recusado pelo PowerPoint (`Presentations.Add(0)`); objetos COM de outro processo não se comparam com `Is` (compare `Shape.Id`).
- **Commits:** pequenos, em português, `feat:`/`fix:`/`test:`/`docs:`/`build:`, terminando com a linha `Co-Authored-By:` indicada no prompt do implementador. **Não** fazer push, **não** criar nem trocar de branch.
- **Versões:** `version.json` não muda nas tarefas. O build importa todo `.bas`/`.cls` de `src/<alvo>/{modules,classes}` e todo `Test*.bas` de `tests/vba/<alvo>`.
- **Processo:** cada tarefa termina com a suíte completa verde na cópia (`pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target all -LoaderSwap -Ppt -TimeoutSec 1200` nas tarefas que mexem no PowerPoint; sem `-Ppt` nas demais) e commit(s). Modelo sugerido: **standard**, **most-capable** (eventos, PowerPoint, geometria), **cheap** (mecânico).

## Review Focus

1. **Automático de verdade, sem custo escondido** (C2): com dois gráficos colados da mesma pasta, só o marcado "Atualizar automaticamente" muda depois de alterar B2 e do ciclo; mudança numa pasta sem `_ThinQQ` ou não salva não agenda nada; o próprio refresh do ThinQQ não reagenda em laço; PowerPoint ocupado ⇒ nova tentativa, nunca "falha" silenciosa. Testes: `Test_PptAuto_Tick_UpdatesOnlyAutoShapes_WhenAvailable`, `Test_PptAuto_OnChange_Filters`, `Test_PptAuto_NoReentryDuringTick`, `Test_PptAuto_Busy_Retries`.
2. **Índice consistente nos dois projetos** (C1/C3): a cadeia gravada pelo `.ppam` e pelo core é a mesma (`Test_PptIdx_Literal` nos dois); entrada velha sai; sincronização completa reconstrói.
3. **Inserir no slide sem deixar lixo** (D2): pasta não salva ⇒ nada criado; sucesso ⇒ gráfico no Excel, moldura no intervalo, forma vinculada no slide com as etiquetas. Teste: `Test_InsertToPpt_*`.
4. **Texto vinculado nunca sobrescreve edição do usuário** (E1): `Test_PptText_Render_*` (puro) e `Test_PptText_Sync_EditedOutsideField_NotTouched_WhenAvailable`.
5. **Segurança**: textos e tabelas abrem a pasta de origem só por `PptSync_OpenBook` (macros desligadas); nada novo abre pasta no modo automático. Teste: `Test_PptText_HostileSourceMacros_NeverRun_Fake` (E1).
6. **Criar pelo PowerPoint** (G2–G4): com uma apresentação salva, "Inserir gráfico > Colunas agrupadas" cria `<nome> - dados ThinQQ.xlsx` ao lado, a planilha "Gráfico 1" com os dados de exemplo e o gráfico, e põe no slide a forma vinculada, automática e marcada `TQQ_DATA`; mudar B2 na planilha e rodar o ciclo atualiza o slide e salva a pasta; apresentação não salva ⇒ mensagem e nada criado; os 21 tipos têm dados de exemplo válidos. Testes: `Test_PptInsertChart_FromPpt_WhenAvailable`, `Test_PptInsertChart_UnsavedPres_NothingCreated`, `Test_Samples_TableFor_AllKindsInsert`, `Test_PptData_*`.

---

## Estrutura de arquivos (novos e alterados)

```
docs/superpowers/specs/2026-10-15-thinqq-thinkcell-parity-design.md   NOVO (C0)
docs/superpowers/plans/2026-10-15-thinqq-thinkcell-parity.md          NOVO (C0)
--- Parte C (1.6.0) ---
src/core/modules/modPptSync.bas    ALT  PptIdx_*, PptSync_SetAuto, PptSync_Items, PptSync_ItemKind/ItemBook, ResolveEntries, índice (C1); ocupado propagado (C2)
src/core/modules/modPptAuto.bas    NOVO opções AutoUpdate/AutoNew (C1); motor de eventos + OnTime (C2)
src/core/modules/modPowerPoint.bas ALT  Ppt_IsBusyError (C1)
src/core/modules/modCoreApi.bas    ALT  Core_PptSyncShapes, Core_PptGetFlag/SetFlag, colagem com AutoNew (C1); Core_OnUnload cancela o ciclo (C2)
src/core/classes/clsAppEvents.cls  ALT  SheetChange, SheetCalculate, WorkbookAfterSave (C2)
src/ppt/modules/modPptTqq.bas, modPptRibbon.bas, classes/clsPptEvents.cls, ribbon/customUI14.xml   ALT (C3)
src/loader/ribbon/customUI14.xml, src/loader/modules/modRibbon.bas   ALT (C4)
tests/vba/core/TestPptSync.bas ALT (C1, C2); tests/vba/ppt/TestPptAddin.bas ALT (C3); tests/vba/loader/TestRibbon.bas, tests/ps/Test-Ribbon.ps1 ALT (C3, C4)
docs/tutorial/tutorial.html, README.md, CHANGELOG.md, release/notes.txt, tests/ps/Test-Docs.ps1   ALT (C5)
--- Parte D (1.7.0) ---
src/core/modules/modLinkFrame.bas  NOVO (D1);  modCoreApi.bas ALT Core_InsertToPowerPoint, Core_ToggleLinkFrames, moldura em Core_RebuildFromRange (D1/D2); modWatchers.bas ALT órfãs (D1)
src/loader/ribbon/customUI14.xml, modRibbon.bas ALT (D2); tests/vba/core/TestLinkFrame.bas NOVO (D1); TestPptSync.bas ALT (D2); TestRibbon.bas ALT (D2)
--- Parte G (1.8.0 G-I; 1.9.0 G-II) ---
src/core/modules/modSamples.bas ALT Samples_TableFor (G1);  modPptData.bas NOVO (G2, G3);  modCoreApi.bas ALT Core_PptInsertChart, Core_PptEditData, Core_PptDataSave (G2, G3), Core_PptCommand, Core_PptCategories (G7)
src/core/modules/modPptAuto.bas ALT salva pastas de dados no ciclo (G3);  modPptSync.bas ALT TQQ_KIND/TQQ_OPTS (G7)
src/ppt/* ALT menu Inserir grafico, Editar dados, reapontar (PptTqq_RelinkData), eventos de salvar/abrir (G4), grupo Grafico do ThinQQ e Empacotar (G8, G9)
tests/vba/core/TestSamples.bas ALT (G1); TestPptData.bas NOVO (G2, G3); TestPptSync.bas ALT (G2, G3, G7); tests/vba/ppt/TestPptAddin.bas ALT (G4, G8, G9); tests/ps/Test-Ribbon.ps1 ALT (G4, G8)
--- Parte E (1.9.0) ---
src/core/modules/modPptText.bas    NOVO (E1);  modPptSync.bas ALT rota "text" (E1);  modPptAuto.bas ALT candidatos por nome (E1)
src/core/modules/modCoreApi.bas    ALT Core_PptLinkText, Core_PptLinkTextFromPpt, Core_LinkTitle (E1–E3)
src/ppt/* ALT (E3);  src/loader/* ALT (E3);  tests/vba/core/TestPptText.bas NOVO (E1, E2)
--- Parte F (1.10.0) ---
src/core/modules/modPptTable.bas   NOVO (F1, F2);  modPptSync.bas ALT rota "table" (F2);  modCoreApi.bas ALT Core_PptInsertTable (F2)
src/loader/* ALT (F3);  tests/vba/core/TestPptTable.bas NOVO (F1, F2)
```

## Como trabalhar (todas as tarefas)

Crie (uma vez por tarefa) o arquivo `C:\temp\tqq-pd-sync.ps1` com o conteúdo abaixo e rode `pwsh -File C:\temp\tqq-pd-sync.ps1 -Repo "<repo>"` antes de cada rodada:

```powershell
param([Parameter(Mandatory)][string]$Repo)
$dst = 'C:\temp\tqq-pd'
if ($dst -notlike 'C:\temp\tqq-pd*') { throw 'destino inesperado' }
robocopy $Repo $dst /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS | Out-Null
if ($LASTEXITCODE -ge 8) { throw "robocopy falhou: $LASTEXITCODE" }
'sincronizado'
```

```powershell
pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target core -Filter "TestPptSync" -TimeoutSec 600      # iteração
pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target all -LoaderSwap -Ppt -TimeoutSec 1200            # suíte completa
```
Ao terminar (bash): `rm -rf /c/temp/tqq-pd /c/temp/tqq-pd-sync.ps1`.

---

# PARTE C — Atualização automática (1.6.0)

### Tarefa C0: Promover spec e plano — modelo: cheap

**Files:** Create `docs/superpowers/specs/2026-10-15-thinqq-thinkcell-parity-design.md` e `docs/superpowers/plans/2026-10-15-thinqq-thinkcell-parity.md` (cópias byte a byte de `design.md` e `plan.md` desta pasta; não sobrescreva se existirem).

- [ ] **Passo 1:** `Copy-Item` dos dois arquivos (só se o destino não existir). Confira `Select-String -Path <spec> -Pattern '^## 6\. Matriz de paridade','^### 3\.1 Sondas pendentes'` (2 linhas).
- [ ] **Passo 2:** `pwsh <repo>\tests\ps\Test-Privacy.ps1 -RepoRoot <repo>` → saída 0.
- [ ] **Passo 3:** `git add` dos dois e `git commit -m "docs: spec e plano da paridade com o think-cell no PowerPoint"` (+ trailer).

---

### Tarefa C1: Índice das formas automáticas, `PptSync_Items`, API 2 e opções — modelo: most-capable

**Files:**
- Modify: `src/core/modules/modPptSync.bas` (constantes, `PptIdx_*`, `PptSync_SetAuto`, `PptSync_EnsureUniqueName`, `PptSync_ItemKind`, `PptSync_ItemBook`, `PptSync_ResolveEntries`, `PptSync_Items`, `PptSync_RebuildIndex`; `PptSync_Presentation` passa a delegar)
- Create: `src/core/modules/modPptAuto.bas` (só as opções nesta tarefa)
- Modify: `src/core/modules/modPowerPoint.bas` (`Ppt_IsBusyError`)
- Modify: `src/core/modules/modCoreApi.bas` (`PptSyncAfterPaste`, `Core_PptSyncShapes`, `Core_PptGetFlag`, `Core_PptSetFlag`)
- Test: `tests/vba/core/TestPptSync.bas` (acrescentar)

**Interfaces:**
- Produces: `PPTSYNC_TAG_AUTO = "TQQ_AUTO"`, `PPTSYNC_TAG_INDEX = "TQQ_AUTOIDX"`, `CORE_PPT_API = 2`; `PptIdx_Entry(slideId As Long, shapeName As String, book As String) As String`; `PptIdx_Parse(idx As String) As Collection` (itens `Array(slideId, nome, pasta)`); `PptIdx_Join(entries As Collection) As String`; `PptIdx_Set(idx, slideId, shapeName, book, turnOn As Boolean) As String`; `PptSync_SetAuto(shp As Object, turnOn As Boolean)`; `PptSync_EnsureUniqueName(shp As Object)`; `PptSync_ItemKind(shp As Object) As String` ("chart" ou ""); `PptSync_ItemBook(shp As Object) As String`; `PptSync_ResolveEntries(pres As Object, entries As String) As Collection`; `PptSync_Items(items As Collection, force As Boolean) As String`; `PptSync_TagDelete(o As Object, key As String)`; `Ppt_IsBusyError(errNum As Long) As Boolean`; `PptAuto_Enabled() As Boolean`, `PptAuto_SetEnabled(v As Boolean)`, `PptAuto_NewLinksAuto() As Boolean`, `PptAuto_SetNewLinksAuto(v As Boolean)`, `PptAuto_SetFlagsForTest(enabled As Variant, newAuto As Variant)`, `PptAuto_ResetTestHooks()`; `Core_PptSyncShapes(pres As Object, entries As String, Optional force As Boolean = True) As String`; `Core_PptGetFlag(key As String) As Boolean`; `Core_PptSetFlag(key As String, value As Boolean)`.
- Consumes: `PptSync_Tag`, `PptSync_SourcePath`, `PptSync_LinkSource` (privada, mesmo módulo), `PptSync_OpenBook`, `PptSync_SyncOne`, `PptSync_CloseOpened`, `PptSync_MissingText`, `PptSync_SummaryText`, `PptSync_Classify`.

- [ ] **Passo 1: Testes (RED)** — acrescente ao fim de `tests/vba/core/TestPptSync.bas`:

```vb
' ---- Parte C (1.6.0): indice das formas automaticas, PptSync_Items, API 2 ----

' Contrato com o .ppam (TestPptAddin.Test_PptIdx_Literal tem as MESMAS cadeias): entradas por Chr(30), campos por Chr(31).
Public Sub Test_PptIdx_Literal()
    Dim s As String, a As String, b As String
    a = "256" & Chr$(31) & "tqq_abc" & Chr$(31) & "C:\a\b.xlsx"
    b = "257" & Chr$(31) & "tqq_def" & Chr$(31) & "C:\a\c.xlsx"
    s = PptIdx_Set("", 256, "tqq_abc", "C:\a\b.xlsx", True)
    AssertEqual a, s
    s = PptIdx_Set(s, 257, "tqq_def", "C:\a\c.xlsx", True)
    AssertEqual a & Chr$(30) & b, s
    AssertEqual a & Chr$(30) & b, PptIdx_Set(s, 257, "tqq_def", "C:\a\c.xlsx", True), "sem repeticao"
    AssertEqual b, PptIdx_Set(s, 256, "tqq_abc", "", False), "desligar tira a entrada"
    AssertEqual "", PptIdx_Set(a, 256, "tqq_abc", "", False)
    AssertEqual 2, PptIdx_Parse(s & Chr$(30) & "x" & Chr$(31) & "y" & Chr$(31) & "z" & Chr$(30) & Chr$(30)).Count, "entrada invalida e vazia ignoradas"
    AssertEqual a, PptIdx_Entry(256, "tqq_abc", "C:\a\b.xlsx")
End Sub

Public Sub Test_PptSync_ItemKind_ByTags()
    Dim f As TestPptFake
    Set f = New TestPptFake
    AssertEqual "", PptSync_ItemKind(f), "duble sem Tags: nenhum tipo (sem erro)"
End Sub

Public Sub Test_PptAuto_Flags_TestBuildInMemory()
    PptAuto_ResetTestHooks
    AssertTrue PptAuto_Enabled(), "padrao: ligado"
    AssertFalse PptAuto_NewLinksAuto(), "padrao: vinculo nasce manual (como o think-cell)"
    Core_PptSetFlag "autonew", True
    AssertTrue Core_PptGetFlag("autonew")
    Core_PptSetFlag "autoon", False
    AssertFalse Core_PptGetFlag("autoon")
    AssertFalse Core_PptGetFlag("outra"), "chave desconhecida: False"
    PptAuto_ResetTestHooks
    AssertTrue PptAuto_Enabled()
End Sub

Public Sub Test_PptSync_SetAuto_IndexAndRebuild_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String, idx As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    PptAuto_ResetTestHooks
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptauto-idx.xlsx")
    AssertEqual "", PptSync_Tag(shp, PPTSYNC_TAG_AUTO), "nasce manual"
    PptSync_SetAuto shp, True
    AssertEqual "1", PptSync_Tag(shp, PPTSYNC_TAG_AUTO)
    idx = PptSync_Tag(pres, PPTSYNC_TAG_INDEX)
    AssertEqual PptIdx_Entry(CLng(shp.Parent.SlideID), CStr(shp.Name), PptSync_ItemBook(shp)), idx
    PptSync_TagDelete pres, PPTSYNC_TAG_INDEX
    PptSync_Presentation pres, True
    AssertEqual idx, PptSync_Tag(pres, PPTSYNC_TAG_INDEX), "sincronizacao completa reconstroi o indice"
    PptSync_SetAuto shp, False
    AssertEqual "", PptSync_Tag(shp, PPTSYNC_TAG_AUTO)
    AssertEqual "", PptSync_Tag(pres, PPTSYNC_TAG_INDEX), "desligar tira do indice"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptAuto_ResetTestHooks
    Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' "Atualizar uma vez" so nos selecionados: duas colagens da mesma pasta, dado novo, so a primeira entrada sincronizada.
Public Sub Test_PptSync_SyncShapes_OnlyListed_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, s2 As Object, own As New Collection, en As Long, ed As String
    Dim v2 As Double, msg As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\pptauto-sel.xlsx")
    Set s2 = PasteAgain(app, pres)
    v2 = PptSeriesValue(s2.Chart, "Sul", 1)
    rng.Worksheet.Range("B2").Value = 42
    Application.Calculate
    msg = Core_PptSyncShapes(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), True)
    AssertTrue InStr(msg, "Gráficos do ThinQQ atualizados: 1.") = 1, msg
    AssertAlmost 42, PptSeriesValue(s1.Chart, "Sul", 1), 0.0001, "selecionado atualizado"
    AssertAlmost v2, PptSeriesValue(s2.Chart, "Sul", 1), 0.0001, "o outro ficou como estava"
    AssertEqual PPTSYNC_NONE, Core_PptSyncShapes(pres, PptIdx_Entry(999999, "nada", ""), True), "entrada que nao existe"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set s1 = Nothing: Set s2 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' AutoNew ligado: a colagem ja nasce automatica e no indice.
Public Sub Test_PptSync_PasteWithAutoNew_TagsAndIndexes_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    PptAuto_SetFlagsForTest Empty, True
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptauto-new.xlsx")
    AssertEqual "1", PptSync_Tag(shp, PPTSYNC_TAG_AUTO)
    AssertTrue InStr(PptSync_Tag(pres, PPTSYNC_TAG_INDEX), CStr(shp.Name)) > 0
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptAuto_ResetTestHooks
    Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub
```

- [ ] **Passo 2:** rode `-Target core -Filter "TestPptSync"` → saída 3 (compilação: `PptIdx_Set` não existe). Anote como RED.

- [ ] **Passo 3: `Ppt_IsBusyError`** — em `src/core/modules/modPowerPoint.bas`, depois de `Ppt_IsKnownFailure`:

```vb
' PowerPoint ocupado (caixa de dialogo aberta, texto em edicao): a chamada pode ser repetida mais tarde.
Public Function Ppt_IsBusyError(errNum As Long) As Boolean
    Ppt_IsBusyError = (errNum = PPT_RPC_E_CALL_REJECTED Or errNum = PPT_RPC_E_RETRYLATER)
End Function
```

- [ ] **Passo 4: Opções** — crie `src/core/modules/modPptAuto.bas`:

```vb
Attribute VB_Name = "modPptAuto"
Option Explicit
' Atualizacao automatica dos elementos do ThinQQ no PowerPoint (spec 2026-10-15 secao 5.1). Opcoes em
' HKCU\Software\VB and VBA Program Settings\ThinQQ\PowerPoint: AutoUpdate (padrao "1": o motor age nas formas marcadas TQQ_AUTO) e AutoNew
' (padrao "0": o vinculo nasce manual, como no think-cell). No build Test so memoria (nunca o registro).
Private Const PPTAUTO_SECTION As String = "PowerPoint"
Private mTestEnabled As Variant                 ' gancho: Empty = padrao
Private mTestNewAuto As Variant

Public Sub PptAuto_SetFlagsForTest(enabled As Variant, newAuto As Variant)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestEnabled = enabled
    mTestNewAuto = newAuto
End Sub

Public Sub PptAuto_ResetTestHooks()
    mTestEnabled = Empty
    mTestNewAuto = Empty
End Sub

Public Function PptAuto_Enabled() As Boolean
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then
        If IsEmpty(mTestEnabled) Then PptAuto_Enabled = True Else PptAuto_Enabled = CBool(mTestEnabled)
        Exit Function
    End If
    PptAuto_Enabled = (GetSetting(APP_NAME, PPTAUTO_SECTION, "AutoUpdate", "1") <> "0")
End Function

Public Sub PptAuto_SetEnabled(v As Boolean)
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then mTestEnabled = v: Exit Sub
    SaveSetting APP_NAME, PPTAUTO_SECTION, "AutoUpdate", IIf(v, "1", "0")
End Sub

Public Function PptAuto_NewLinksAuto() As Boolean
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then
        If IsEmpty(mTestNewAuto) Then PptAuto_NewLinksAuto = False Else PptAuto_NewLinksAuto = CBool(mTestNewAuto)
        Exit Function
    End If
    PptAuto_NewLinksAuto = (GetSetting(APP_NAME, PPTAUTO_SECTION, "AutoNew", "0") = "1")
End Function

Public Sub PptAuto_SetNewLinksAuto(v As Boolean)
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then mTestNewAuto = v: Exit Sub
    SaveSetting APP_NAME, PPTAUTO_SECTION, "AutoNew", IIf(v, "1", "0")
End Sub
```

- [ ] **Passo 5: Índice e itens** — em `src/core/modules/modPptSync.bas`:
  - troque `Public Const CORE_PPT_API As Long = 1` por `Public Const CORE_PPT_API As Long = 2                       ' 1 = Core_PptSyncPresentation; 2 = Core_PptSyncShapes, Core_PptGetFlag/SetFlag (1.6.0)`;
  - depois de `PPTSYNC_TAG_FP` acrescente:

```vb
Public Const PPTSYNC_TAG_AUTO As String = "TQQ_AUTO"        ' "1" = atualizar automaticamente (spec 2026-10-15 secao 5.1)
Public Const PPTSYNC_TAG_INDEX As String = "TQQ_AUTOIDX"   ' etiqueta da APRESENTACAO: indice das formas automaticas (PptIdx_*)
```

  - acrescente, antes de `' ---- Sincronizacao da apresentacao inteira (Tarefa A2) ----`:

```vb
' ---- Indice das formas automaticas (1.6.0). Formato (igual no .ppam: PptTqq_Idx*): entradas separadas por Chr(30), campos por Chr(31):
' SlideID, nome da forma, pasta de origem. Puras. ----

Public Function PptIdx_Entry(slideId As Long, shapeName As String, book As String) As String
    PptIdx_Entry = CStr(slideId) & Chr$(31) & shapeName & Chr$(31) & book
End Function

Public Function PptIdx_Parse(idx As String) As Collection
    Dim out As New Collection, e As Variant, f() As String
    For Each e In Split(idx, Chr$(30))
        If Len(CStr(e)) > 0 Then
            f = Split(CStr(e), Chr$(31))
            If UBound(f) = 2 Then
                If Len(f(0)) > 0 And Len(f(0)) < 10 And Len(f(1)) > 0 Then
                    If IsNumeric(f(0)) And InStr(f(0), ",") = 0 And InStr(f(0), ".") = 0 And InStr(f(0), "-") = 0 Then out.Add Array(CLng(f(0)), f(1), f(2))
                End If
            End If
        End If
    Next
    Set PptIdx_Parse = out
End Function

Public Function PptIdx_Join(entries As Collection) As String
    Dim e As Variant, s As String
    For Each e In entries
        If Len(s) > 0 Then s = s & Chr$(30)
        s = s & PptIdx_Entry(CLng(e(0)), CStr(e(1)), CStr(e(2)))
    Next
    PptIdx_Join = s
End Function

' Tira a forma (slideId, nome) do indice e, com turnOn, a acrescenta no fim. Pura.
Public Function PptIdx_Set(idx As String, slideId As Long, shapeName As String, book As String, turnOn As Boolean) As String
    Dim out As New Collection, e As Variant
    For Each e In PptIdx_Parse(idx)
        If Not (CLng(e(0)) = slideId And StrComp(CStr(e(1)), shapeName, vbBinaryCompare) = 0) Then out.Add e
    Next
    If turnOn Then out.Add Array(slideId, shapeName, book)
    PptIdx_Set = PptIdx_Join(out)
End Function

' Apaga uma etiqueta (forma ou apresentacao); inexistente nao e erro. Nunca propaga erro.
Public Sub PptSync_TagDelete(o As Object, key As String)
    On Error Resume Next
    o.Tags.Delete key
    Err.Clear
End Sub

' Tipo do elemento do ThinQQ na forma do slide: "chart" (TQQ_CHART); "" = nao e do ThinQQ. As versoes seguintes acrescentam "text" e
' "table". Nunca propaga erro.
Public Function PptSync_ItemKind(shp As Object) As String
    If Len(PptSync_Tag(shp, PPTSYNC_TAG_CHART)) > 0 Then PptSync_ItemKind = "chart"
End Function

' Pasta de origem do elemento (vinculo nativo ou TQQ_BOOK). Nunca propaga erro.
Public Function PptSync_ItemBook(shp As Object) As String
    PptSync_ItemBook = PptSync_SourcePath(PptSync_LinkSource(shp), PptSync_Tag(shp, PPTSYNC_TAG_BOOK))
End Function

' Nome unico no slide (o indice acha a forma pelo nome): outra forma com o mesmo nome (Shape.Id diferente: objetos de outro processo nao se
' comparam com Is) faz esta virar "<nome>_2", "_3"... O nome da forma do slide nao e usado pelo Excel.
Public Sub PptSync_EnsureUniqueName(shp As Object)
    Dim sl As Object, s As Object, base As String, n As Long, clash As Boolean, myId As Long
    Set sl = shp.Parent
    base = CStr(shp.Name): myId = CLng(shp.Id)
    For Each s In sl.Shapes
        If CLng(s.Id) <> myId And StrComp(CStr(s.Name), base, vbBinaryCompare) = 0 Then clash = True: Exit For
    Next
    If Not clash Then Exit Sub
    For n = 2 To 999
        If Not PptSync_SlideHasName(sl, base & "_" & n) Then shp.Name = base & "_" & n: Exit Sub
    Next
End Sub

Private Function PptSync_SlideHasName(sl As Object, nm As String) As Boolean
    Dim s As Object
    For Each s In sl.Shapes
        If StrComp(CStr(s.Name), nm, vbBinaryCompare) = 0 Then PptSync_SlideHasName = True: Exit Function
    Next
End Function

' Liga/desliga "Atualizar automaticamente" de uma forma do ThinQQ e mantem o indice da apresentacao (TQQ_AUTOIDX). Propaga erro do PowerPoint.
Public Sub PptSync_SetAuto(shp As Object, turnOn As Boolean)
    Dim sl As Object, pres As Object, idx As String
    Set sl = shp.Parent: Set pres = sl.Parent
    If turnOn Then PptSync_EnsureUniqueName shp
    If turnOn Then shp.Tags.Add PPTSYNC_TAG_AUTO, "1" Else PptSync_TagDelete shp, PPTSYNC_TAG_AUTO
    idx = PptIdx_Set(PptSync_Tag(pres, PPTSYNC_TAG_INDEX), CLng(sl.SlideID), CStr(shp.Name), PptSync_ItemBook(shp), turnOn)
    If Len(idx) = 0 Then PptSync_TagDelete pres, PPTSYNC_TAG_INDEX Else pres.Tags.Add PPTSYNC_TAG_INDEX, idx
End Sub

' Formas do ThinQQ apontadas pelas entradas (formato do indice; a pasta da entrada e ignorada): slide por SlideID, forma pelo nome. Entrada
' que nao acha nada e ignorada; PowerPoint ocupado sobe (nunca vira "nao achou").
Public Function PptSync_ResolveEntries(pres As Object, entries As String) As Collection
    Dim out As New Collection, e As Variant, sl As Object, shp As Object, en As Long
    For Each e In PptIdx_Parse(entries)
        Set sl = Nothing: Set shp = Nothing
        On Error Resume Next
        Set sl = pres.Slides.FindBySlideID(CLng(e(0)))
        If Not sl Is Nothing Then Set shp = sl.Shapes(CStr(e(1)))
        en = Err.Number
        Err.Clear
        On Error GoTo 0
        If Ppt_IsBusyError(en) Then Err.Raise en, "PptSync_ResolveEntries", "PowerPoint ocupado"
        If Not shp Is Nothing Then
            If Len(PptSync_ItemKind(shp)) > 0 Then out.Add shp
        End If
    Next
    Set PptSync_ResolveEntries = out
End Function

' Reconstroi TQQ_AUTOIDX com as formas automaticas de items (todas as formas do ThinQQ da apresentacao); so grava se mudou. Nunca propaga erro.
Private Sub PptSync_RebuildIndex(pres As Object, items As Collection)
    Dim shp As Object, out As New Collection, s As String
    On Error GoTo Done
    For Each shp In items
        If PptSync_Tag(shp, PPTSYNC_TAG_AUTO) = "1" Then
            PptSync_EnsureUniqueName shp
            out.Add Array(CLng(shp.Parent.SlideID), CStr(shp.Name), PptSync_ItemBook(shp))
        End If
    Next
    s = PptIdx_Join(out)
    If s <> PptSync_Tag(pres, PPTSYNC_TAG_INDEX) Then
        If Len(s) = 0 Then PptSync_TagDelete pres, PPTSYNC_TAG_INDEX Else pres.Tags.Add PPTSYNC_TAG_INDEX, s
    End If
Done:
End Sub
```

  - substitua a função `PptSync_Presentation` INTEIRA (da linha do comentário `' Sincroniza todos os graficos ThinQQ vinculados (formas soltas...` até o `End Function` dela) por:

```vb
' Sincroniza todos os elementos do ThinQQ (formas soltas marcadas) da apresentacao e reconstroi o indice das automaticas. force = False pula
' os que nao mudaram (TQQ_FP). Regras de origem, aviso e erros: PptSync_Items. Devolve o resumo (PptSync_SummaryText).
Public Function PptSync_Presentation(pres As Object, force As Boolean) As String
    Dim sl As Object, shp As Object, items As Collection
    PptSync_Hook "presentation"
    Set items = New Collection
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If Len(PptSync_ItemKind(shp)) > 0 Then items.Add shp
        Next
    Next
    If items.Count = 0 Then PptSync_Presentation = PPTSYNC_NONE: Exit Function
    PptSync_RebuildIndex pres, items
    PptSync_Presentation = PptSync_Items(items, force)
End Function

' Sincroniza as formas dadas (todas da mesma apresentacao ou nao). Pasta de origem inexistente ou que nao abre (senha, corrompida): o elemento
' fica como esta e a pasta entra no aviso (NUNCA LinkFormat.Update com a origem ausente: travaria o PowerPoint). Update ignorado por origem
' inalcancavel: nao conta como atualizado e a TQQ_FP nao e gravada. Forma sem vinculo ou com estado ilegivel nunca e sincronizada. Pastas
' abertas por esta chamada sao fechadas sem salvar, com a tela congelada (restaurada em todo caminho). Erro de um elemento e contado
' (PptSync_SyncOne); PowerPoint ocupado e falha estrutural sobem.
Public Function PptSync_Items(items As Collection, force As Boolean) As String
    Dim shp As Object, opened As Collection, books As Object, missingSet As Object, path As String, k As Long
    Dim wb As Workbook, missing As String, en As Long, ed As String, prevSU As Boolean, ls As Long
    Dim nOk As Long, nSame As Long, nNoChart As Long, nUnlinked As Long, nFail As Long
    Set opened = New Collection
    Set books = CreateObject("Scripting.Dictionary")
    books.CompareMode = 1
    Set missingSet = CreateObject("Scripting.Dictionary")
    missingSet.CompareMode = 1
    prevSU = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False
    For Each shp In items
        k = k + 1
        Application.StatusBar = "ThinQQ: atualizando gráfico " & k & " de " & items.Count & " do PowerPoint..."
        ls = Ppt_LinkState(shp)
        path = ""
        If ls = 1 Then path = PptSync_SourcePath(PptSync_LinkSource(shp), PptSync_Tag(shp, PPTSYNC_TAG_BOOK))
        Select Case PptSync_Classify(ls, path)
            Case PPTSYNC_C_UNLINKED
                nUnlinked = nUnlinked + 1
            Case PPTSYNC_C_UNKNOWN
                nFail = nFail + 1
                Errors_Log "PptSync: estado do vinculo ilegivel; forma nao sincronizada" & vbTab & PptSync_Tag(shp, PPTSYNC_TAG_CHART)
            Case PPTSYNC_C_NOPATH
                If Not missingSet.Exists("") Then missingSet.Add "", True
            Case Else
                If Not books.Exists(path) Then
                    Set wb = Nothing
                    If PptSync_FileExists(path) Then Set wb = PptSync_OpenBook(path, opened)
                    If wb Is Nothing Then missingSet(path) = True
                    books.Add path, wb
                End If
                Set wb = books.Item(path)
                If Not wb Is Nothing Then
                    Select Case PptSync_SyncOne(shp, wb, force)
                        Case PPTSYNC_R_OK: nOk = nOk + 1
                        Case PPTSYNC_R_SAME: nSame = nSame + 1
                        Case PPTSYNC_R_NOCHART: nNoChart = nNoChart + 1
                        Case PPTSYNC_R_SKIPPED: missingSet(path) = True
                        Case Else: nFail = nFail + 1
                    End Select
                End If
        End Select
    Next
    PptSync_CloseOpened opened
    Application.ScreenUpdating = prevSU
    Application.StatusBar = False
    If missingSet.Count > 0 Then missing = PptSync_MissingText(missingSet.Keys)
    PptSync_Items = PptSync_SummaryText(nOk, nSame, missing, nNoChart, nUnlinked, nFail)
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptSync_CloseOpened opened
    Application.ScreenUpdating = prevSU
    Application.StatusBar = False
    On Error GoTo 0
    Err.Raise en, "PptSync_Items", ed
End Function
```

  - em `PptSync_SyncOne`, troque o bloco `Fail:` por (PowerPoint ocupado sobe para quem pode tentar de novo):

```vb
Fail:
    If Ppt_IsBusyError(Err.Number) Then
        Dim bn As Long, bd As String
        bn = Err.Number: bd = Err.Description
        On Error GoTo -1
        Err.Raise bn, "PptSync_SyncOne", bd
    End If
    Errors_Log "PptSync" & vbTab & Err.Number & vbTab & Err.Description
    PptSync_SyncOne = PPTSYNC_R_FAIL
```
  (o `Dim` dentro do bloco é válido em VBA; se o lint recusar, mova `bn`/`bd` para o `Dim` do topo da função.)

- [ ] **Passo 6: API** — em `src/core/modules/modCoreApi.bas`:
  - em `PptSyncAfterPaste`, depois da linha `If ok Then shp.Tags.Add PPTSYNC_TAG_FP, PptSync_Fingerprint(co, shp)`, acrescente:

```vb
    If PptAuto_NewLinksAuto() Then
        PptSync_SetAuto shp, True                     ' opcao "Novos vinculos atualizam automaticamente" (1.6.0)
        If Err.Number <> 0 Then Errors_Log "CopyToPowerPoint/automatico" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    End If
```
  - depois de `Core_PptApiVersion`, acrescente:

```vb
' API 2 (1.6.0): "Atualizar uma vez" do .ppam so nas formas listadas (entries no formato do indice, PptIdx_*). Devolve o resumo; nunca propaga.
Public Function Core_PptSyncShapes(pres As Object, entries As String, Optional force As Boolean = True) As String
    Dim items As Collection, en As Long, ed As String
    On Error GoTo Fail
    Set items = PptSync_ResolveEntries(pres, entries)
    If items.Count = 0 Then Core_PptSyncShapes = PPTSYNC_NONE: Exit Function
    Core_PptSyncShapes = PptSync_Items(items, force)
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Errors_Log "Core_PptSyncShapes" & vbTab & en & vbTab & ed
    If Ppt_IsBusyError(en) Then Core_PptSyncShapes = PPT_ERR_BUSY Else Core_PptSyncShapes = "Não foi possível atualizar os gráficos do ThinQQ: " & ed & " (" & en & ")"
End Function

' Opcoes do PowerPoint para a faixa do Excel: "autoon" (atualizacao automatica ligada) e "autonew" (novos vinculos ja automaticos).
Public Function Core_PptGetFlag(key As String) As Boolean
    On Error Resume Next
    Select Case key
        Case "autoon": Core_PptGetFlag = PptAuto_Enabled()
        Case "autonew": Core_PptGetFlag = PptAuto_NewLinksAuto()
    End Select
End Function

Public Sub Core_PptSetFlag(key As String, value As Boolean)
    On Error Resume Next
    Select Case key
        Case "autoon": PptAuto_SetEnabled value
        Case "autonew": PptAuto_SetNewLinksAuto value
    End Select
End Sub
```

- [ ] **Passo 7 (GREEN):** sincronize e rode `-Target core -Filter "TestPptSync|TestCoreApi"` → 0 falhas (com o PowerPoint livre, os testes `_WhenAvailable` rodam de verdade: confira no log que nenhum foi "pulado"). Depois a suíte completa (`-Target all -LoaderSwap -TimeoutSec 1200`) → "Todos os testes passaram". Se o teste de contenção da API (`TestCoreApi`, todos os `Core_*` nunca propagam) listar as funções novas, ele já as cobre pelo `On Error`.
- [ ] **Passo 8:** commits `feat: indice das formas automaticas e PptSync_Items no motor do PowerPoint` (modPptSync, modPowerPoint, modPptAuto, testes) e `feat: API 2 do PowerPoint (atualizar selecionados, opcoes de atualizacao automatica)` (modCoreApi) (+ trailer).

---

### Tarefa C2: Motor automático (eventos do Excel + `OnTime`) — modelo: most-capable

**Pré-requisito:** resultado das sondas P1/P3 no ledger. P1 > 100 ms por gráfico ⇒ acrescente o Passo 6b (comparação barata antes da impressão digital, abaixo). P3 com caixa "aguardando ação OLE" ⇒ a guarda `PptAuto_PptEditingText` já está no código; relate.

**Files:**
- Modify: `src/core/modules/modPptAuto.bas` (motor)
- Modify: `src/core/classes/clsAppEvents.cls` (3 eventos)
- Modify: `src/core/modules/modCoreApi.bas` (`Core_OnUnload` cancela o ciclo)
- Test: `tests/vba/core/TestPptSync.bas` (acrescentar)

**Interfaces:**
- Produces: `PPTAUTO_DEBOUNCE_S = 2`, `PPTAUTO_MAX_RETRIES = 5`; `PptAuto_OnChange(wb As Workbook)`; `PptAuto_IsCandidate(wb As Workbook) As Boolean`; `PptAuto_Tick()` (alvo do `OnTime`); `PptAuto_Cancel()`; `PptAuto_ItemsFor(pres As Object, dirty As Object) As Collection`; ganchos `PptAuto_SetNoTimerForTest(v As Boolean)`, `PptAuto_PendingCountForTest() As Long`, `PptAuto_IsScheduledForTest() As Boolean`, `PptAuto_LastResultForTest() As String`, `PptAuto_SetBusyForTest(n As Long)`, `PptAuto_RetriesForTest() As Long`.
- Consumes: `Ppt_GetApp(False)`, `Ppt_IsBusyError`, `PptSync_Items`, `PptSync_Tag`, `PptSync_TagDelete`, `PptIdx_Parse/Join`, `PPTSYNC_TAG_AUTO/INDEX`, `CALC_SHEET`.

- [ ] **Passo 1: Testes (RED)** — acrescente ao fim de `tests/vba/core/TestPptSync.bas`:

```vb
' ---- Parte C2: motor automatico ----

Private Sub AutoReset()
    PptAuto_Cancel
    PptAuto_ResetTestHooks
    PptAuto_SetNoTimerForTest True
End Sub

' So pastas SALVAS com graficos do ThinQQ (_ThinQQ) agendam; desligado nada agenda.
Public Sub Test_PptAuto_OnChange_Filters()
    Dim wb As Workbook, wb2 As Workbook, rng As Range
    AutoReset
    Set wb2 = NewTempWorkbook()
    PptAuto_OnChange wb2
    AssertEqual 0, PptAuto_PendingCountForTest(), "pasta nova sem ThinQQ e sem arquivo"
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    PptAuto_OnChange wb
    AssertEqual 0, PptAuto_PendingCountForTest(), "pasta nao salva (o vinculo exige o arquivo)"
    wb.SaveAs TempDir() & "\pptauto-filter.xlsx", xlOpenXMLWorkbook
    PptAuto_SetFlagsForTest False, Empty
    PptAuto_OnChange wb
    AssertEqual 0, PptAuto_PendingCountForTest(), "desligado"
    PptAuto_SetFlagsForTest Empty, Empty
    PptAuto_OnChange wb
    AssertEqual 1, PptAuto_PendingCountForTest()
    AssertTrue PptAuto_IsScheduledForTest(), "ciclo agendado"
    PptAuto_OnChange wb
    AssertEqual 1, PptAuto_PendingCountForTest(), "a mesma pasta conta uma vez"
    AutoReset
    AssertFalse PptAuto_IsScheduledForTest(), "Cancel desagenda"
    CloseTemp wb2
    Teardown wb
End Sub

' Eventos reais do Excel: digitar na pasta salva agenda o ciclo.
Public Sub Test_PptAuto_SheetChangeEvent_Schedules()
    Dim wb As Workbook, rng As Range
    AutoReset
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    wb.SaveAs TempDir() & "\pptauto-evt.xlsx", xlOpenXMLWorkbook
    Watchers_Start
    PptAuto_Cancel: PptAuto_SetNoTimerForTest True
    rng.Worksheet.Range("B2").Value = 77
    AssertEqual 1, PptAuto_PendingCountForTest(), "SheetChange/SheetCalculate marcaram a pasta"
    AutoReset
    Teardown wb
End Sub

' Sem PowerPoint injetado (build Test: Ppt_GetApp devolve so a instancia injetada) o ciclo esvazia a fila sem erro.
Public Sub Test_PptAuto_Tick_NoPowerPoint_ClearsQueue()
    Dim wb As Workbook, rng As Range
    AutoReset
    Ppt_ResetTestHooks
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    wb.SaveAs TempDir() & "\pptauto-noppt.xlsx", xlOpenXMLWorkbook
    PptAuto_OnChange wb
    PptAuto_Tick
    AssertEqual 0, PptAuto_PendingCountForTest()
    AssertEqual 0, PptAuto_RetriesForTest()
    AutoReset
    Teardown wb
End Sub

' PowerPoint ocupado (gancho): nova tentativa agendada, fila mantida; depois de PPTAUTO_MAX_RETRIES desiste e esvazia.
Public Sub Test_PptAuto_Busy_Retries()
    Dim wb As Workbook, rng As Range, i As Long
    AutoReset
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    wb.SaveAs TempDir() & "\pptauto-busy.xlsx", xlOpenXMLWorkbook
    PptAuto_OnChange wb
    PptAuto_SetBusyForTest 99
    PptAuto_Tick
    AssertEqual 1, PptAuto_RetriesForTest()
    AssertEqual 1, PptAuto_PendingCountForTest(), "fila mantida para a nova tentativa"
    AssertTrue PptAuto_IsScheduledForTest()
    For i = 2 To PPTAUTO_MAX_RETRIES + 1
        PptAuto_Tick
    Next
    AssertEqual 0, PptAuto_PendingCountForTest(), "desistiu"
    AssertEqual 0, PptAuto_RetriesForTest()
    AutoReset
    Teardown wb
End Sub

' Review Focus 1: so a forma automatica muda; entrada de forma apagada sai do indice; o refresh do proprio ciclo nao reagenda.
Public Sub Test_PptAuto_Tick_UpdatesOnlyAutoShapes_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, s2 As Object, s3 As Object, own As New Collection
    Dim en As Long, ed As String, v2 As Double
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoReset
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\pptauto-tick.xlsx")
    Set s2 = PasteAgain(app, pres)
    Set s3 = PasteAgain(app, pres)
    PptSync_SetAuto s1, True
    PptSync_SetAuto s3, True
    s3.Delete: Set s3 = Nothing
    v2 = PptSeriesValue(s2.Chart, "Sul", 1)
    Watchers_Start
    rng.Worksheet.Range("B2").Value = 55                 ' evento real: marca a pasta (pasta aberta, nao salva: o vinculo le o valor novo)
    AssertEqual 1, PptAuto_PendingCountForTest()
    Ppt_SetAppForTest app
    PptAuto_Tick
    AssertAlmost 55, PptSeriesValue(s1.Chart, "Sul", 1), 0.0001, "automatica atualizada"
    AssertAlmost v2, PptSeriesValue(s2.Chart, "Sul", 1), 0.0001, "manual ficou"
    AssertEqual 1, PptIdx_Parse(PptSync_Tag(pres, PPTSYNC_TAG_INDEX)).Count, "entrada da forma apagada saiu do indice"
    AssertEqual 0, PptAuto_PendingCountForTest(), "o refresh do proprio ciclo nao marcou a pasta de novo"
    AssertTrue InStr(PptAuto_LastResultForTest(), "Gráficos do ThinQQ atualizados: 1.") = 1, PptAuto_LastResultForTest()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set s1 = Nothing: Set s2 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Sem mudanca nos dados (so recalculo): a impressao digital pula; nada reescrito.
Public Sub Test_PptAuto_Tick_SkipsUnchanged_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, own As New Collection, en As Long, ed As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoReset
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\pptauto-same.xlsx")
    PptSync_SetAuto s1, True
    PptAuto_OnChange wb
    Ppt_SetAppForTest app
    PptSync_ResetTestHooks
    PptAuto_Tick
    AssertTrue InStr(PptAuto_LastResultForTest(), "Sem mudança: 1.") > 0, PptAuto_LastResultForTest()
    AssertEqual 0, PptSync_UpdateCountForTest(), "nenhum LinkFormat.Update"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set s1 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub
```

- [ ] **Passo 2:** `-Target core -Filter "TestPptSync"` → saída 3 (`PptAuto_OnChange` não existe). RED.

- [ ] **Passo 3: Motor** — acrescente ao fim de `src/core/modules/modPptAuto.bas` (as declarações `Private` novas vão para o topo, junto das existentes):

```vb
' (topo do modulo, depois de mTestNewAuto)
Public Const PPTAUTO_DEBOUNCE_S As Long = 2                 ' espera depois da ULTIMA mudanca (cada evento reagenda)
Public Const PPTAUTO_MAX_RETRIES As Long = 5                ' PowerPoint ocupado: novas tentativas antes de desistir
Private mDue As Date                                        ' horario do OnTime agendado (0 = nenhum): o cancelamento exige o mesmo valor
Private mDirty As Object                                    ' Scripting.Dictionary (texto): FullName das pastas que mudaram
Private mInTick As Boolean                                  ' reentrada e eventos do proprio ciclo
Private mRetries As Long
Private mLastResult As String
Private mTestNoTimer As Boolean                             ' gancho: marca o horario sem chamar Application.OnTime (o teste chama PptAuto_Tick)
Private mTestBusy As Long                                   ' gancho: quantos ciclos simulam PowerPoint ocupado
```

```vb
Public Sub PptAuto_SetNoTimerForTest(v As Boolean)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestNoTimer = v
End Sub

Public Sub PptAuto_SetBusyForTest(n As Long)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestBusy = n
End Sub

Public Function PptAuto_PendingCountForTest() As Long
    If mDirty Is Nothing Then Exit Function
    PptAuto_PendingCountForTest = mDirty.Count
End Function

Public Function PptAuto_IsScheduledForTest() As Boolean
    PptAuto_IsScheduledForTest = (mDue <> 0)
End Function

Public Function PptAuto_LastResultForTest() As String
    PptAuto_LastResultForTest = mLastResult
End Function

Public Function PptAuto_RetriesForTest() As Long
    PptAuto_RetriesForTest = mRetries
End Function

' Pasta que pode ter elementos no PowerPoint: salva (o vinculo aponta para o arquivo) e com a planilha _ThinQQ (graficos do ThinQQ). As
' versoes 1.8.0/1.9.0 acrescentam os nomes definidos TQQ_X_/TQQ_T_. Nunca propaga erro.
Public Function PptAuto_IsCandidate(wb As Workbook) As Boolean
    Dim ws As Object
    On Error Resume Next
    If wb Is Nothing Then Exit Function
    If Len(wb.Path) = 0 Then Exit Function
    Set ws = wb.Worksheets(CALC_SHEET)
    PptAuto_IsCandidate = Not ws Is Nothing
    Err.Clear
End Function

' Eventos do Excel (SheetChange, SheetCalculate, WorkbookAfterSave): marca a pasta e (re)agenda o ciclo para daqui a PPTAUTO_DEBOUNCE_S.
' Barato quando nada se aplica (desligado, pasta sem ThinQQ). Ignora os eventos do proprio ciclo. Nunca propaga erro.
Public Sub PptAuto_OnChange(wb As Workbook)
    On Error GoTo Done
    If mInTick Then Exit Sub
    If Not PptAuto_Enabled() Then Exit Sub
    If Not PptAuto_IsCandidate(wb) Then Exit Sub
    If mDirty Is Nothing Then
        Set mDirty = CreateObject("Scripting.Dictionary")
        mDirty.CompareMode = 1
    End If
    mDirty(wb.FullName) = True
    mRetries = 0
    PptAuto_Schedule PPTAUTO_DEBOUNCE_S
Done:
End Sub

Private Function PptAuto_Proc() As String
    PptAuto_Proc = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PptAuto_Tick"
End Function

Private Sub PptAuto_Schedule(secs As Long)
    PptAuto_CancelTimer
    mDue = Now + TimeSerial(0, 0, secs)
    If BUILD_CONFIG = "Test" And mTestNoTimer Then Exit Sub
    On Error Resume Next
    Application.OnTime mDue, PptAuto_Proc()
    If Err.Number <> 0 Then Errors_Log "PptAuto_Schedule" & vbTab & Err.Number & vbTab & Err.Description: mDue = 0
    Err.Clear
End Sub

Private Sub PptAuto_CancelTimer()
    If mDue = 0 Then Exit Sub
    If Not (BUILD_CONFIG = "Test" And mTestNoTimer) Then
        On Error Resume Next
        Application.OnTime mDue, PptAuto_Proc(), , False
        Err.Clear
    End If
    mDue = 0
End Sub

' Desliga tudo (Core_OnUnload, opcao desligada, testes): desagenda e esvazia a fila.
Public Sub PptAuto_Cancel()
    On Error Resume Next
    PptAuto_CancelTimer
    Set mDirty = Nothing
    mRetries = 0
    mInTick = False
    mTestBusy = 0
    mTestNoTimer = False
    mLastResult = ""
End Sub

' Texto em edicao no PowerPoint (o ciclo adia em vez de disputar com o usuario). Nunca propaga erro.
Private Function PptAuto_PptEditingText(app As Object) As Boolean
    On Error Resume Next
    PptAuto_PptEditingText = (CLng(app.ActiveWindow.Selection.Type) = 3)     ' ppSelectionText
    Err.Clear
End Function

' Formas automaticas de pres cuja pasta de origem esta em dirty, pelo indice TQQ_AUTOIDX (nunca varre todas as formas). Entradas velhas
' (slide ou forma apagados, TQQ_AUTO ausente) saem do indice. PowerPoint ocupado sobe.
Public Function PptAuto_ItemsFor(pres As Object, dirty As Object) As Collection
    Dim out As New Collection, keep As New Collection, e As Variant, sl As Object, shp As Object, stale As Boolean, idx As String, en As Long
    Set PptAuto_ItemsFor = out
    idx = PptSync_Tag(pres, PPTSYNC_TAG_INDEX)
    If Len(idx) = 0 Then Exit Function
    For Each e In PptIdx_Parse(idx)
        Set sl = Nothing: Set shp = Nothing
        On Error Resume Next
        Set sl = pres.Slides.FindBySlideID(CLng(e(0)))
        If Not sl Is Nothing Then Set shp = sl.Shapes(CStr(e(1)))
        en = Err.Number
        Err.Clear
        On Error GoTo 0
        If Ppt_IsBusyError(en) Then Err.Raise en, "PptAuto_ItemsFor", "PowerPoint ocupado"
        If shp Is Nothing Then
            stale = True
        ElseIf PptSync_Tag(shp, PPTSYNC_TAG_AUTO) <> "1" Then
            stale = True
        Else
            keep.Add e
            If dirty.Exists(CStr(e(2))) Then out.Add shp
        End If
    Next
    If stale Then
        On Error Resume Next
        If keep.Count = 0 Then PptSync_TagDelete pres, PPTSYNC_TAG_INDEX Else pres.Tags.Add PPTSYNC_TAG_INDEX, PptIdx_Join(keep)
        Err.Clear
    End If
End Function

' Alvo do OnTime: sincroniza, nas apresentacoes editaveis do PowerPoint JA aberto, as formas automaticas das pastas que mudaram (sem forcar:
' a impressao digital pula o que nao mudou). Sem PowerPoint: esvazia a fila. Ocupado: nova tentativa em 2 * (n + 1) s, ate PPTAUTO_MAX_RETRIES.
' Sem caixa de mensagem (so log e barra de status). Nunca propaga erro.
Public Sub PptAuto_Tick()
    Dim app As Object, pres As Object, items As Collection, s As String, en As Long, ed As String
    mDue = 0
    If mInTick Then Exit Sub
    If mDirty Is Nothing Then Exit Sub
    If mDirty.Count = 0 Then Exit Sub
    mInTick = True
    On Error GoTo Fail
    If BUILD_CONFIG = "Test" And mTestBusy > 0 Then mTestBusy = mTestBusy - 1: Err.Raise -2147418111, "PptAuto_Tick", "ocupado (teste)"
    Set app = Ppt_GetApp(False)
    If Not app Is Nothing Then
        If PptAuto_PptEditingText(app) Then Err.Raise -2147418111, "PptAuto_Tick", "texto em edicao no PowerPoint"
        For Each pres In app.Presentations
            If CLng(pres.ReadOnly) = 0 Then
                Set items = PptAuto_ItemsFor(pres, mDirty)
                If items.Count > 0 Then
                    s = PptSync_Items(items, False)
                    mLastResult = s
                    Errors_Log "PptAuto" & vbTab & CStr(pres.Name) & vbTab & Replace(s, vbLf, " | ")
                End If
            End If
        Next
    End If
    mDirty.RemoveAll
    mRetries = 0
    mInTick = False
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    mInTick = False
    Application.StatusBar = False
    If Ppt_IsBusyError(en) And mRetries < PPTAUTO_MAX_RETRIES Then
        mRetries = mRetries + 1
        Errors_Log "PptAuto: PowerPoint ocupado (" & ed & "); nova tentativa " & mRetries
        PptAuto_Schedule PPTAUTO_DEBOUNCE_S * (mRetries + 1)
    Else
        Errors_Log "PptAuto: desistiu" & vbTab & en & vbTab & ed
        mDirty.RemoveAll
        mRetries = 0
    End If
End Sub
```

- [ ] **Passo 4: Eventos** — em `src/core/classes/clsAppEvents.cls`, acrescente ao fim:

```vb
' Atualizacao automatica no PowerPoint (1.6.0): qualquer mudanca numa pasta salva com graficos do ThinQQ agenda o ciclo (sai cedo quando nao se aplica).
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    PptAuto_OnChange Sh.Parent
End Sub

Private Sub App_SheetCalculate(ByVal Sh As Object)
    On Error Resume Next
    PptAuto_OnChange Sh.Parent
End Sub

Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    On Error Resume Next
    If Success Then PptAuto_OnChange Wb
End Sub
```

- [ ] **Passo 5:** em `Core_OnUnload` (`modCoreApi.bas`), depois de `DeltaPick_Disarm ""`, acrescente `PptAuto_Cancel                       ' o OnTime do ciclo automatico nunca fica preso ao core descarregado`.

- [ ] **Passo 6 (GREEN):** `-Target core -Filter "TestPptSync|TestCoreApi|TestDeltaPick"` → 0 falhas; confira no log que os `_WhenAvailable` rodaram. **Passo 6b (só se P1 > 100 ms):** em `PptSync_SyncOne`, antes de `fp = PptSync_Fingerprint(co, shp)`, compare uma assinatura barata (`PptSync_QuickSig`: `Join` dos valores do bloco + tamanho, sem SHA) gravada em `TQQ_QS`; igual e `Not force` ⇒ `PPTSYNC_R_SAME`; acrescente teste que muda um dado e espera "atualizados: 1".
- [ ] **Passo 7:** suíte completa `-Target all -LoaderSwap -Ppt -TimeoutSec 1200` → verde, `POWERPNT` 0 antes/depois.
- [ ] **Passo 8:** commit `feat: atualizacao automatica dos graficos do ThinQQ no PowerPoint (eventos do Excel)` (+ trailer).

---

### Tarefa C3: Suplemento do PowerPoint — "Atualizar automaticamente" e "Atualizar uma vez" — modelo: most-capable

**Files:**
- Modify: `src/ppt/modules/modPptTqq.bas`, `src/ppt/modules/modPptRibbon.bas`, `src/ppt/classes/clsPptEvents.cls`, `src/ppt/ribbon/customUI14.xml`
- Test: `tests/vba/ppt/TestPptAddin.bas`, `tests/ps/Test-Ribbon.ps1`

**Interfaces:**
- Produces: `PPT_TAG_AUTO`, `PPT_TAG_INDEX`, `PPT_API_SHAPES = 2`, `PPT_MSG_SELECT`, `PPT_MSG_AUTO_OLDCORE`; `PptTqq_IdxEntry/IdxParse/IdxJoin/IdxSet` (cópias exatas das `PptIdx_*`); `PptTqq_SetAutoShapes(shapes As Collection, turnOn As Boolean)`; `PptTqq_ShapesAuto(shapes As Collection) As Boolean`; `PptTqq_EntriesOf(shapes As Collection) As String`; `PptTqq_SelectedTagged() As Collection`; `PptTqq_SyncEntries(pres As Object, entries As String, force As Boolean)`; callbacks `PptRx_GetAuto`, `PptRx_OnAuto`, `PptRx_GetElemEnabled`, `PptRx_OnUpdateSel`, `PptRx_InvalidateElem`.
- Consumes: `Core_PptSyncShapes` (API 2), `Core_PptApiVersion`, funções existentes do `.ppam`.

- [ ] **Passo 1: Testes (RED)** — acrescente ao fim de `tests/vba/ppt/TestPptAddin.bas`:

```vb
' Contrato com o core (TestPptSync.Test_PptIdx_Literal tem as MESMAS cadeias).
Public Sub Test_PptIdx_Literal()
    Dim s As String, a As String, b As String
    a = "256" & Chr$(31) & "tqq_abc" & Chr$(31) & "C:\a\b.xlsx"
    b = "257" & Chr$(31) & "tqq_def" & Chr$(31) & "C:\a\c.xlsx"
    s = PptTqq_IdxSet("", 256, "tqq_abc", "C:\a\b.xlsx", True)
    AssertEqual a, s
    s = PptTqq_IdxSet(s, 257, "tqq_def", "C:\a\c.xlsx", True)
    AssertEqual a & Chr$(30) & b, s
    AssertEqual a & Chr$(30) & b, PptTqq_IdxSet(s, 257, "tqq_def", "C:\a\c.xlsx", True), "sem repeticao"
    AssertEqual b, PptTqq_IdxSet(s, 256, "tqq_abc", "", False)
    AssertEqual "", PptTqq_IdxSet(a, 256, "tqq_abc", "", False)
    AssertEqual 2, PptTqq_IdxParse(s & Chr$(30) & "x" & Chr$(31) & "y" & Chr$(31) & "z" & Chr$(30) & Chr$(30)).Count
    AssertEqual a, PptTqq_IdxEntry(256, "tqq_abc", "C:\a\b.xlsx")
End Sub

' Ligar/desligar sem Excel: etiqueta na forma, indice na apresentacao, nome repetido no slide renomeado.
Public Sub Test_PptTqq_SetAutoShapes_TagsAndIndex()
    Dim p As Object, s1 As Object, s2 As Object, c As New Collection, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    Set s1 = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): s1.Name = "tqq_x": s1.Tags.Add PPT_TAG_CHART, "abc123": s1.Tags.Add PPT_TAG_BOOK, "C:\t\a.xlsx"
    Set s2 = p.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50): s2.Name = "tqq_x": s2.Tags.Add PPT_TAG_CHART, "abc123": s2.Tags.Add PPT_TAG_BOOK, "C:\t\a.xlsx"
    c.Add s1: c.Add s2
    AssertFalse PptTqq_ShapesAuto(c)
    PptTqq_SetAutoShapes c, True
    AssertEqual "1", s1.Tags(PPT_TAG_AUTO): AssertEqual "1", s2.Tags(PPT_TAG_AUTO)
    AssertTrue s1.Name <> s2.Name, "nomes unicos no slide"
    AssertEqual 2, PptTqq_IdxParse(p.Tags(PPT_TAG_INDEX)).Count
    AssertTrue PptTqq_ShapesAuto(c)
    PptTqq_SetAutoShapes c, False
    AssertEqual "", p.Tags(PPT_TAG_INDEX)
    AssertFalse PptTqq_ShapesAuto(c)
    AssertEqual PptTqq_IdxEntry(p.Slides(1).SlideID, s1.Name, "") & Chr$(30) & PptTqq_IdxEntry(p.Slides(1).SlideID, s2.Name, ""), PptTqq_EntriesOf(c)
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' "Atualizar uma vez" sem Excel (gancho com Nothing): aviso de Excel ausente, nada quebra.
Public Sub Test_PptTqq_SyncEntries_NoExcel()
    Dim p As Object
    Set p = NewTagged()
    PptTqq_SetExcelForTest Nothing, ""
    PptTqq_ClearLastMsg
    PptTqq_SyncEntries p, PptTqq_IdxEntry(p.Slides(1).SlideID, p.Slides(1).Shapes(1).Name, ""), True
    AssertEqual PPT_MSG_NO_EXCEL, PptTqq_LastMsg()
    PptTqq_ResetTestHooks
    ClosePres p
End Sub
```

  Em `tests/ps/Test-Ribbon.ps1`, depois da linha que confere `btnPptSyncNow`, acrescente:

```powershell
Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptElem']/u:toggleButton[@id='tglPptAuto'][@getPressed='PptRx_GetAuto'][@onAction='PptRx_OnAuto'][@getEnabled='PptRx_GetElemEnabled']", $pns)) 'ppt: Atualizar automaticamente'
Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptElem']/u:button[@id='btnPptUpdateSel'][@onAction='PptRx_OnUpdateSel'][@getEnabled='PptRx_GetElemEnabled']", $pns)) 'ppt: Atualizar uma vez'
```

- [ ] **Passo 2:** `pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target core -Filter "TestSmoke" -Ppt -TimeoutSec 900` (monta o `.pptm` de teste) → falha de compilação do projeto do PowerPoint (`PptTqq_IdxSet`). RED.

- [ ] **Passo 3: Código do `.ppam`** — em `src/ppt/modules/modPptTqq.bas`, depois de `PPT_REQUIRED_API`:

```vb
Public Const PPT_TAG_AUTO As String = "TQQ_AUTO"
Public Const PPT_TAG_INDEX As String = "TQQ_AUTOIDX"
Public Const PPT_API_SHAPES As Long = 2                     ' Core_PptSyncShapes (core 1.6.0)
Public Const PPT_MSG_SELECT As String = "Selecione um ou mais gráficos do ThinQQ vinculados ao Excel."
Public Const PPT_MSG_AUTO_OLDCORE As String = "Atualize o ThinQQ no Excel para a atualização automática funcionar (a marca foi gravada)."
```

  e, ao fim do módulo:

```vb
' ---- 1.6.0: atualizacao automatica por elemento. Indice: COPIA EXATA das PptIdx_* do core (mesmos testes de cadeia). ----

Public Function PptTqq_IdxEntry(slideId As Long, shapeName As String, book As String) As String
    PptTqq_IdxEntry = CStr(slideId) & Chr$(31) & shapeName & Chr$(31) & book
End Function

Public Function PptTqq_IdxParse(idx As String) As Collection
    Dim out As New Collection, e As Variant, f() As String
    For Each e In Split(idx, Chr$(30))
        If Len(CStr(e)) > 0 Then
            f = Split(CStr(e), Chr$(31))
            If UBound(f) = 2 Then
                If Len(f(0)) > 0 And Len(f(0)) < 10 And Len(f(1)) > 0 Then
                    If IsNumeric(f(0)) And InStr(f(0), ",") = 0 And InStr(f(0), ".") = 0 And InStr(f(0), "-") = 0 Then out.Add Array(CLng(f(0)), f(1), f(2))
                End If
            End If
        End If
    Next
    Set PptTqq_IdxParse = out
End Function

Public Function PptTqq_IdxJoin(entries As Collection) As String
    Dim e As Variant, s As String
    For Each e In entries
        If Len(s) > 0 Then s = s & Chr$(30)
        s = s & PptTqq_IdxEntry(CLng(e(0)), CStr(e(1)), CStr(e(2)))
    Next
    PptTqq_IdxJoin = s
End Function

Public Function PptTqq_IdxSet(idx As String, slideId As Long, shapeName As String, book As String, turnOn As Boolean) As String
    Dim out As New Collection, e As Variant
    For Each e In PptTqq_IdxParse(idx)
        If Not (CLng(e(0)) = slideId And StrComp(CStr(e(1)), shapeName, vbBinaryCompare) = 0) Then out.Add e
    Next
    If turnOn Then out.Add Array(slideId, shapeName, book)
    PptTqq_IdxSet = PptTqq_IdxJoin(out)
End Function

Private Sub PptTqq_TagDelete(o As Object, key As String)
    On Error Resume Next
    o.Tags.Delete key
    Err.Clear
End Sub

Private Function PptTqq_SlideHasName(sl As Object, nm As String) As Boolean
    Dim s As Object
    For Each s In sl.Shapes
        If StrComp(CStr(s.Name), nm, vbBinaryCompare) = 0 Then PptTqq_SlideHasName = True: Exit Function
    Next
End Function

Private Sub PptTqq_EnsureUniqueName(shp As Object)
    Dim sl As Object, s As Object, base As String, n As Long, clash As Boolean
    Set sl = shp.Parent
    base = CStr(shp.Name)
    For Each s In sl.Shapes
        If s.Id <> shp.Id And StrComp(CStr(s.Name), base, vbBinaryCompare) = 0 Then clash = True: Exit For
    Next
    If Not clash Then Exit Sub
    For n = 2 To 999
        If Not PptTqq_SlideHasName(sl, base & "_" & n) Then shp.Name = base & "_" & n: Exit Sub
    Next
End Sub

' Formas do ThinQQ (TQQ_CHART) da selecao da janela ativa (formas ou texto dentro de uma forma). Vazia sem janela/selecao.
Public Function PptTqq_SelectedTagged() As Collection
    Dim out As New Collection, sel As Object, sr As Object, i As Long
    Set PptTqq_SelectedTagged = out
    On Error Resume Next
    Set sel = Application.ActiveWindow.Selection
    If sel Is Nothing Then Exit Function
    If sel.Type <> 2 And sel.Type <> 3 Then Exit Function          ' ppSelectionShapes, ppSelectionText
    Set sr = sel.ShapeRange
    If sr Is Nothing Then Exit Function
    For i = 1 To sr.Count
        If Len(PptTqq_Tag(sr(i), PPT_TAG_CHART)) > 0 Then out.Add sr(i)
    Next
    Err.Clear
End Function

' Todas as formas dadas sao automaticas (e ha pelo menos uma).
Public Function PptTqq_ShapesAuto(shapes As Collection) As Boolean
    Dim s As Object
    If shapes.Count = 0 Then Exit Function
    For Each s In shapes
        If PptTqq_Tag(s, PPT_TAG_AUTO) <> "1" Then Exit Function
    Next
    PptTqq_ShapesAuto = True
End Function

' Liga/desliga "Atualizar automaticamente" (etiqueta + indice da apresentacao), sem o Excel.
Public Sub PptTqq_SetAutoShapes(shapes As Collection, turnOn As Boolean)
    Dim s As Object, pres As Object, idx As String, book As String
    For Each s In shapes
        Set pres = s.Parent.Parent
        If turnOn Then PptTqq_EnsureUniqueName s
        If turnOn Then s.Tags.Add PPT_TAG_AUTO, "1" Else PptTqq_TagDelete s, PPT_TAG_AUTO
        book = PptTqq_SourcePath(PptTqq_LinkSource(s), PptTqq_Tag(s, PPT_TAG_BOOK))
        idx = PptTqq_IdxSet(PptTqq_Tag(pres, PPT_TAG_INDEX), CLng(s.Parent.SlideID), CStr(s.Name), book, turnOn)
        If Len(idx) = 0 Then PptTqq_TagDelete pres, PPT_TAG_INDEX Else pres.Tags.Add PPT_TAG_INDEX, idx
    Next
End Sub

' Entradas (formato do indice, pasta vazia) das formas dadas: o que Core_PptSyncShapes recebe.
Public Function PptTqq_EntriesOf(shapes As Collection) As String
    Dim c As New Collection, s As Object
    For Each s In shapes
        c.Add Array(CLng(s.Parent.SlideID), CStr(s.Name), "")
    Next
    PptTqq_EntriesOf = PptTqq_IdxJoin(c)
End Function

' "Atualizar uma vez" das entradas dadas: core com API 2 => Core_PptSyncShapes; core mais velho => a apresentacao inteira (aviso). Nunca propaga.
Public Sub PptTqq_SyncEntries(pres As Object, entries As String, force As Boolean)
    Dim xl As Object, created As Boolean, core As String, msg As String, v As Variant, en As Long, ed As String
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fail
    Set xl = PptTqq_GetExcel(created)
    If xl Is Nothing Then mBusy = False: PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation: Exit Sub
    core = PptTqq_EnsureCore(xl, created)
    If Len(core) = 0 Then
        msg = PptTqq_DataOnly(pres, PptTqq_WorkbookNames(xl)) & vbLf & PPT_MSG_DATA_ONLY
    ElseIf PptTqq_CoreApi(xl, core) >= PPT_API_SHAPES Then
        v = xl.Run("'" & core & "'!Core_PptSyncShapes", pres, entries, force)
        msg = CStr(v)
    Else
        msg = PptTqq_RunCoreSync(xl, core, pres, force)
    End If
    PptTqq_Release xl, created
    mBusy = False
    PptTqq_Msg msg, vbInformation
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptTqq_Release xl, created
    mBusy = False
    PptTqq_Msg PptTqq_ErrorText(en, ed), vbExclamation
End Sub
```

  Em `src/ppt/modules/modPptRibbon.bas`, acrescente:

```vb
Public Sub PptRx_InvalidateElem()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.InvalidateControl "tglPptAuto": mRibbon.InvalidateControl "btnPptUpdateSel"
End Sub

Public Sub PptRx_GetElemEnabled(control As IRibbonControl, ByRef enabled)
    On Error Resume Next
    enabled = (PptTqq_SelectedTagged().Count > 0)
End Sub

Public Sub PptRx_GetAuto(control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    pressed = PptTqq_ShapesAuto(PptTqq_SelectedTagged())
End Sub

' Ligar ja atualiza uma vez (como o "Update Automatically" do think-cell).
Public Sub PptRx_OnAuto(control As IRibbonControl, pressed As Boolean)
    Dim sel As Collection
    On Error Resume Next
    Set sel = PptTqq_SelectedTagged()
    If sel.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    PptTqq_SetAutoShapes sel, pressed
    If pressed Then PptTqq_SyncEntries Application.ActivePresentation, PptTqq_EntriesOf(sel), True
    PptRx_InvalidateElem
End Sub

Public Sub PptRx_OnUpdateSel(control As IRibbonControl)
    Dim sel As Collection
    On Error Resume Next
    Set sel = PptTqq_SelectedTagged()
    If sel.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    PptTqq_SyncEntries Application.ActivePresentation, PptTqq_EntriesOf(sel), True
End Sub
```

  Em `src/ppt/classes/clsPptEvents.cls`, acrescente:

```vb
' Selecao mudou: o estado de "Atualizar automaticamente" segue o grafico selecionado.
Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    On Error Resume Next
    PptRx_InvalidateElem
End Sub
```

  Em `src/ppt/ribbon/customUI14.xml`, depois do `</group>` de `grpPptSync`:

```xml
        <group id="grpPptElem" label="Gráfico selecionado">
          <toggleButton id="tglPptAuto" label="Atualizar automaticamente" size="large" imageMso="Refresh"
                        getPressed="PptRx_GetAuto" onAction="PptRx_OnAuto" getEnabled="PptRx_GetElemEnabled"
                        screentip="Atualizar automaticamente" supertip="Com o Excel aberto, os gráficos do ThinQQ selecionados passam a acompanhar sozinhos as mudanças da pasta de origem (dados e setas), cerca de 2 segundos depois de cada mudança. Clique de novo para voltar ao manual."/>
          <button id="btnPptUpdateSel" label="Atualizar uma vez" imageMso="Refresh" onAction="PptRx_OnUpdateSel" getEnabled="PptRx_GetElemEnabled"
                  screentip="Atualizar uma vez" supertip="Atualiza agora dados e setas só dos gráficos do ThinQQ selecionados."/>
        </group>
```

  Em `PptRx_OnAbout`, troque o texto final por `"Atualiza dados e setas dos gráficos do ThinQQ copiados como gráfico vinculado: ao abrir a apresentação (aba ThinQQ > Ao abrir), pelo botão ou automaticamente (Gráfico selecionado > Atualizar automaticamente)."`.

- [ ] **Passo 4 (GREEN):** `-Target core -Filter "TestSmoke" -Ppt -TimeoutSec 900` → projeto do PowerPoint 0 falhas e e2e ok; `pwsh <repo>\tests\ps\Test-Ribbon.ps1 -RepoRoot <repo>` → 0 falhas.
- [ ] **Passo 5:** suíte completa com `-Ppt` → verde. Commit `feat: suplemento do PowerPoint com Atualizar automaticamente e Atualizar uma vez` (+ trailer).

---

### Tarefa C4: Faixa do Excel — opções da atualização automática — modelo: standard

**Files:** Modify `src/loader/ribbon/customUI14.xml`, `src/loader/modules/modRibbon.bas`; Test `tests/vba/loader/TestRibbon.bas`.

**Interfaces:** Produces `Rx_GetPptFlag(control, ByRef pressed)`, `Rx_OnPptFlag(control, pressed As Boolean)`. Consumes `Core_PptGetFlag`, `Core_PptSetFlag`.

- [ ] **Passo 1: Teste (RED)** — acrescente a `tests/vba/loader/TestRibbon.bas`:

```vb
Public Sub Test_Ribbon_PptAutoCheckBoxes()
    Dim d As Object
    Set d = LoadRibbon()
    AssertEqual 2, d.SelectNodes("//u:menu[@id='mnuPpt']/u:checkBox").Length
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:checkBox[@id='chkPptAutoOn'][@tag='autoon'][@getPressed='Rx_GetPptFlag'][@onAction='Rx_OnPptFlag']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:checkBox[@id='chkPptAutoNew'][@tag='autonew'][@getPressed='Rx_GetPptFlag'][@onAction='Rx_OnPptFlag']") Is Nothing
    AssertEqual "msepPptAuto", AttrText(d.SelectSingleNode("//u:button[@id='btnPptSync']/following-sibling::*[1]"), "id")
    AssertTrue Len(AttrText(d.SelectSingleNode("//*[@id='chkPptAutoOn']"), "supertip")) > 20
    AssertTrue Len(AttrText(d.SelectSingleNode("//*[@id='chkPptAutoNew']"), "supertip")) > 20
End Sub

' Core sem os procedimentos (sufixo de teste): a caixa fica desmarcada e o clique avisa "Atualize".
Public Sub Test_Ribbon_PptFlag_OldCore()
    Dim p As Variant
    Rx_TestSetFeatureProcSuffix "_naoexiste"
    Loader_LastMsg = ""
    Rx_OnPptFlagCore "autoon", True
    AssertEqual RX_MSG_UPDATE, Loader_LastMsg
    Rx_TestSetFeatureProcSuffix ""
End Sub
```

- [ ] **Passo 2: XML** — em `src/loader/ribbon/customUI14.xml`, depois do `btnPptSync` (antes de `</menu>` do `mnuPpt`):

```xml
            <menuSeparator id="msepPptAuto"/>
            <checkBox id="chkPptAutoOn" label="Atualização automática ligada" tag="autoon" getPressed="Rx_GetPptFlag" onAction="Rx_OnPptFlag"
                      screentip="Atualização automática ligada" supertip="Com o Excel e o PowerPoint abertos, atualiza sozinho os gráficos marcados como 'Atualizar automaticamente' cerca de 2 segundos depois que um valor da pasta muda. Desmarque para pausar."/>
            <checkBox id="chkPptAutoNew" label="Novos vínculos atualizam automaticamente" tag="autonew" getPressed="Rx_GetPptFlag" onAction="Rx_OnPptFlag"
                      screentip="Novos vínculos atualizam automaticamente" supertip="Gráficos colados ou inseridos no PowerPoint a partir de agora já nascem com 'Atualizar automaticamente' ligado (sem esta opção nascem manuais, como no think-cell)."/>
```

  Atualize o supertip do `mnuPpt` para: "Cola o gráfico selecionado no slide ativo do PowerPoint (como imagem ou gráfico vinculado), atualiza os gráficos vinculados das apresentações abertas e controla a atualização automática."

- [ ] **Passo 3: Callbacks** — em `src/loader/modules/modRibbon.bas`, depois de `Rx_OnPptSync`:

```vb
' Opcoes da atualizacao automatica (1.6.0): "autoon" e "autonew" (tag do controle). Core anterior: desmarcada; o clique avisa "Atualize".
Public Sub Rx_GetPptFlag(control As IRibbonControl, ByRef pressed)
    Dim v As Variant, p As String
    On Error Resume Next
    pressed = False
    If Loader_CoreWorkbook() Is Nothing Then Exit Sub
    p = "Core_PptGetFlag"
    If BUILD_CONFIG = "Test" Then p = p & mTestProcSuffix
    v = Loader_RunCore(p, control.Tag)
    If Not IsEmpty(v) Then pressed = CBool(v)
End Sub

Public Sub Rx_OnPptFlag(control As IRibbonControl, pressed As Boolean)
    On Error Resume Next
    Rx_OnPptFlagCore control.Tag, pressed
End Sub

' Parte testavel de Rx_OnPptFlag (IRibbonControl nao e criavel nos testes).
Public Sub Rx_OnPptFlagCore(tag As String, pressed As Boolean)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_PptSetFlag", Array(tag, pressed)
    Loader_InvalidateRibbon
End Sub
```

  `RunCoreFeature` hoje passa um único argumento: troque o seu corpo por

```vb
Private Sub RunCoreFeature(procName As String, Optional arg As Variant)
    Dim p As String
    p = procName
    If BUILD_CONFIG = "Test" Then p = p & mTestProcSuffix
    If IsMissing(arg) Then
        Loader_RunCore p
    ElseIf IsArray(arg) Then
        Loader_RunCore p, arg(0), arg(1)
    Else
        Loader_RunCore p, arg
    End If
    If Loader_RunCoreMissed() Then RxMsg RX_MSG_UPDATE, vbInformation
End Sub
```
  (confira em `modLoaderMain.bas` que `Loader_RunCore` aceita dois argumentos opcionais; se tiver `ParamArray`, o mesmo código vale.)

- [ ] **Passo 4 (GREEN):** `-Target loader -Filter "TestRibbon"` → 0 falhas; `Test-Ribbon.ps1` → 0 falhas.
- [ ] **Passo 5:** suíte completa sem `-Ppt`. Commit `feat: opcoes da atualizacao automatica no menu PowerPoint do Excel` (+ trailer).

---

### Tarefa C5: Documentação da 1.6.0 — modelo: standard

**Files:** `docs/tutorial/tutorial.html` (seção `#powerpoint`: subseção nova `#ppt-auto` "Atualização automática"), `README.md` (Recursos; Limitações; Verificações manuais novos itens 61–68), `CHANGELOG.md` (`[Não lançado]` > Adicionado), `release/notes.txt` (uma linha), `tests/ps/Test-Docs.ps1` (rótulos e frases).

- [ ] **Passo 1:** Tutorial `#ppt-auto`: o que é (como o "Update Automatically" do think-cell), como ligar (PowerPoint: aba ThinQQ > Gráfico selecionado > Atualizar automaticamente; Excel: PowerPoint > Novos vínculos atualizam automaticamente), "Atualizar uma vez", pausa (Atualização automática ligada), quando acontece (Excel e PowerPoint abertos, pasta salva com o gráfico, ~2 s depois da última mudança; inclusive valores não salvos), limites (só a pasta que mudou; apresentação somente leitura não; texto em edição no PowerPoint adia; slides copiados de outra apresentação voltam a ser automáticos depois de uma "Atualizar gráficos"; sem o Excel aberto não há atualização automática — vale a pergunta ao abrir).
- [ ] **Passo 2:** README itens manuais (numerar a partir do último existente): (61) digitar valores com o slide visível ao lado e ver dados e setas mudarem em ≤ 3 s; (62) duas apresentações abertas com gráficos da mesma pasta; (63) PowerPoint com texto em edição durante a mudança (sem caixa "aguardando ação OLE" no Excel; anote se aparecer: sonda P3); (64) apresentação somente leitura não muda; (65) pausa global e "Novos vínculos" pelo menu; (66) 20 gráficos automáticos de uma pasta: tempo do Excel ocupado; (67) salvar, fechar e reabrir a apresentação: continua automática; (68) `.ppam` 1.6.0 com core 1.5.x: aviso de atualizar.
- [ ] **Passo 3:** CHANGELOG/notes: "Atualização automática dos gráficos do ThinQQ no PowerPoint (por gráfico, como no think-cell) e Atualizar uma vez só nos selecionados."
- [ ] **Passo 4:** Test-Docs: rótulos `Atualizar automaticamente`, `Atualizar uma vez`, `Atualização automática ligada`, `Novos vínculos atualizam automaticamente` no tutorial e no README; âncora `id="ppt-auto"`. Rode `Test-Docs.ps1`, `Test-Privacy.ps1` (pwsh 7 e 5.1) → 0 falhas.
- [ ] **Passo 5:** commit `docs: atualizacao automatica no PowerPoint (tutorial, README, CHANGELOG)` (+ trailer).

### Tarefa C6: Suíte final, revisão do ramo e verificação manual — modelo: standard (controlador + humano)
- [ ] Suíte completa `-Target all -LoaderSwap -Ppt -TimeoutSec 1200` verde, `POWERPNT` 0/0, 0 testes de PowerPoint pulados.
- [ ] Revisão final do ramo (modelo mais capaz) com foco: Review Focus 1–2 e 5; uma onda de correção; re-revisão escopada.
- [ ] Humano: itens manuais 61–68 do README (inclui as sondas P2/P3 como observação), PR, `publicar-release 1.6.0`.

---
# PARTE D — Inserir a partir do intervalo direto no slide + moldura (1.7.0)

### Tarefa D1: Moldura do intervalo vinculado (`modLinkFrame`) — modelo: standard

**Files:** Create `src/core/modules/modLinkFrame.bas`; Modify `src/core/modules/modWatchers.bas` (órfãs), `src/core/modules/modCoreApi.bas` (`Core_ToggleLinkFrames`; moldura refeita em `Core_RebuildFromRange`); Test `tests/vba/core/TestLinkFrame.bas` (novo).

**Interfaces:** Produces `LINKFRAME_PREFIX = "tqqlink_"`; `LinkFrame_Name(cht As Chart) As String`; `LinkFrame_Ensure(cht As Chart, rng As Range) As Shape`; `LinkFrame_Find(cht As Chart) As Shape`; `LinkFrame_Remove(cht As Chart)`; `LinkFrame_CleanOrphans(wb As Workbook)`; `LinkFrame_ToggleAll(wb As Workbook) As Long` (−1 sem molduras, 0 ocultas, 1 visíveis); `Core_ToggleLinkFrames()`. Consumes `Registry_IdFromChart`, `Registry_Workbook`, `Registry_FindChartObject`.

- [ ] **Passo 1: Testes (RED)** — crie `tests/vba/core/TestLinkFrame.bas`:

```vb
Attribute VB_Name = "TestLinkFrame"
Option Explicit
' 1.7.0: moldura verde em volta do intervalo de origem de um grafico que foi ao PowerPoint (spec 2026-10-15 secao 5.2).

Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5): ws.Range("A3:C3").Value = Array("T2", 20, 6)
    Set Fixture = ws.Range("A1:C3")
End Function

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

Public Sub Test_LinkFrame_DrawnAroundRangeAndFollowsRows()
    Dim wb As Workbook, rng As Range, c As Chart, s As Shape
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set s = LinkFrame_Ensure(c, rng)
    AssertTrue Not s Is Nothing
    AssertEqual LINKFRAME_PREFIX & Registry_IdFromChart(c), s.Name
    AssertAlmost rng.Top, s.Top, 0.5: AssertAlmost rng.Left, s.Left, 0.5
    AssertAlmost rng.Height, s.Height, 0.5: AssertAlmost rng.Width, s.Width, 0.5
    AssertEqual msoFalse, s.Fill.Visible, "sem preenchimento"
    AssertEqual RGB(0, 176, 80), s.Line.ForeColor.RGB
    AssertEqual xlMoveAndSize, s.Placement
    rng.Worksheet.Rows(1).Insert
    AssertAlmost rng.Worksheet.Range("A2:C4").Top, s.Top, 0.5, "linha acima: a moldura desce"
    rng.Worksheet.Rows(3).Insert
    AssertAlmost rng.Worksheet.Range("A2:C5").Height, s.Height, 0.5, "linha dentro: a moldura cresce"
    Set s = LinkFrame_Ensure(c, rng.Worksheet.Range("A2:B5"))
    AssertEqual 1, CountFrames(wb), "refazer nao duplica"
    LinkFrame_Remove c
    AssertEqual 0, CountFrames(wb)
    Teardown wb
End Sub

Public Sub Test_LinkFrame_OrphansAndToggle()
    Dim wb As Workbook, rng As Range, c As Chart, s As Shape
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set s = LinkFrame_Ensure(c, rng)
    AssertEqual 0, LinkFrame_ToggleAll(wb), "visivel -> oculta"
    AssertEqual msoFalse, s.Visible
    AssertEqual 1, LinkFrame_ToggleAll(wb), "oculta -> visivel"
    c.Parent.Delete
    LinkFrame_CleanOrphans wb
    AssertEqual 0, CountFrames(wb), "grafico apagado: moldura orfa apagada"
    AssertEqual -1, LinkFrame_ToggleAll(wb), "sem molduras"
    Teardown wb
End Sub

Private Function CountFrames(wb As Workbook) As Long
    Dim ws As Worksheet, s As Shape
    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            If Left$(s.Name, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then CountFrames = CountFrames + 1
        Next
    Next
End Function
```

- [ ] **Passo 2:** `-Target core -Filter "TestLinkFrame"` → saída 3. RED.
- [ ] **Passo 3:** crie `src/core/modules/modLinkFrame.bas`:

```vb
Attribute VB_Name = "modLinkFrame"
Option Explicit
' Moldura do intervalo de origem de um grafico do ThinQQ que foi ao PowerPoint (como a "data link frame" verde do think-cell; spec 2026-10-15
' secao 5.2): retangulo sem preenchimento, linha verde de 2 pt, que acompanha linhas/colunas inseridas (Placement = xlMoveAndSize, medido),
' nao impresso e bloqueado. Nome tqqlink_<id do grafico> (fora de ID_PREFIX + id: nao e ChartObject). Arrastar a moldura NAO muda o vinculo
' (mudar o intervalo = Alterar dados, que redesenha a moldura). Nada aqui propaga erro.
Public Const LINKFRAME_PREFIX As String = "tqqlink_"

Public Function LinkFrame_Name(cht As Chart) As String
    LinkFrame_Name = LINKFRAME_PREFIX & Registry_IdFromChart(cht)
End Function

' Desenho isolado (sonda pendente P4: se o retangulo atrapalhar cliques nas celulas, trocar por 4 linhas AddLine aqui).
Private Function LinkFrame_Draw(ws As Worksheet, rng As Range) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = RGB(0, 176, 80)
    s.Line.Weight = 2
    s.Placement = xlMoveAndSize
    s.Locked = True
    On Error Resume Next
    s.DrawingObject.PrintObject = False
    Err.Clear
    Set LinkFrame_Draw = s
End Function

' A moldura do grafico (em qualquer planilha da pasta do grafico) ou Nothing.
Public Function LinkFrame_Find(cht As Chart) As Shape
    Dim ws As Worksheet, s As Shape, nm As String
    On Error Resume Next
    nm = LinkFrame_Name(cht)
    For Each ws In Registry_Workbook(cht).Worksheets
        Set s = Nothing
        Set s = ws.Shapes(nm)
        Err.Clear
        If Not s Is Nothing Then Set LinkFrame_Find = s: Exit Function
    Next
End Function

Public Sub LinkFrame_Remove(cht As Chart)
    Dim s As Shape, guard As Long
    On Error Resume Next
    Do
        Set s = LinkFrame_Find(cht)
        If s Is Nothing Then Exit Do
        s.Delete
        guard = guard + 1
    Loop While guard < 20
End Sub

' Desenha (ou redesenha) a moldura em volta de rng. Devolve a forma (Nothing se falhou: so registrado).
Public Function LinkFrame_Ensure(cht As Chart, rng As Range) As Shape
    Dim s As Shape
    On Error GoTo Fail
    LinkFrame_Remove cht
    Set s = LinkFrame_Draw(rng.Worksheet, rng)
    s.Name = LinkFrame_Name(cht)
    s.AlternativeText = "ThinQQ: intervalo do gráfico " & Registry_IdFromChart(cht) & " vinculado ao PowerPoint"
    Set LinkFrame_Ensure = s
    Exit Function
Fail:
    Errors_Log "LinkFrame_Ensure" & vbTab & Err.Number & vbTab & Err.Description
End Function

' Molduras cujo grafico nao existe mais na pasta.
Public Sub LinkFrame_CleanOrphans(wb As Workbook)
    Dim ws As Worksheet, i As Long, nm As String
    On Error Resume Next
    For Each ws In wb.Worksheets
        For i = ws.Shapes.Count To 1 Step -1
            nm = ws.Shapes(i).Name
            If Left$(nm, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then
                If Registry_FindChartObject(wb, Mid$(nm, Len(LINKFRAME_PREFIX) + 1)) Is Nothing Then ws.Shapes(i).Delete
            End If
        Next
    Next
    Err.Clear
End Sub

' Mostrar/ocultar todas as molduras da pasta: alguma visivel => oculta todas (0); todas ocultas => mostra (1); nenhuma moldura => -1.
Public Function LinkFrame_ToggleAll(wb As Workbook) As Long
    Dim ws As Worksheet, s As Shape, anyVis As Boolean, n As Long
    On Error Resume Next
    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            If Left$(s.Name, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then
                n = n + 1
                If s.Visible = msoTrue Then anyVis = True
            End If
        Next
    Next
    If n = 0 Then LinkFrame_ToggleAll = -1: Exit Function
    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            If Left$(s.Name, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then s.Visible = IIf(anyVis, msoFalse, msoTrue)
        Next
    Next
    LinkFrame_ToggleAll = IIf(anyVis, 0, 1)
End Function
```

- [ ] **Passo 4:** em `modWatchers.bas`, no fim de `Watchers_Rehook` (depois do `Next` do laço), acrescente `LinkFrame_CleanOrphans ws.Parent` (dentro de `On Error Resume Next` local). Em `modCoreApi.bas`:

```vb
' Mostrar/ocultar as molduras dos intervalos vinculados da pasta ativa (1.7.0). Nunca propaga erro.
Public Sub Core_ToggleLinkFrames()
    Dim r As Long
    On Error GoTo Fail
    If ActiveWorkbook Is Nothing Then Exit Sub
    r = LinkFrame_ToggleAll(ActiveWorkbook)
    If r = -1 Then Core_Notify "Esta pasta não tem molduras de intervalos vinculados ao PowerPoint (elas aparecem ao usar 'Inserir também no slide')."
    Exit Sub
Fail:
    Errors_Handle "Core_ToggleLinkFrames"
End Sub
```
  e, em `Core_RebuildFromRange`, no caminho de sucesso (logo antes do `Core_RebuildFromRange = True` final), acrescente `If Not LinkFrame_Find(c) Is Nothing Then LinkFrame_Ensure c, rng      ' Alterar dados: a moldura acompanha a tabela nova`. Acrescente a `TestLinkFrame.bas` o teste `Test_LinkFrame_FollowsChangeSource` (inserir, moldura, `Core_RebuildFromRange c, ws.Range("A1:B3")` ⇒ largura da moldura = largura de A1:B3).
- [ ] **Passo 5 (GREEN):** `-Filter "TestLinkFrame|TestCoreApi"` → 0 falhas; suíte completa. Commit `feat: moldura verde no intervalo de origem dos graficos vinculados ao PowerPoint` (+ trailer).

---

### Tarefa D2: "Inserir também no slide" — modelo: most-capable

**Files:** Modify `src/core/modules/modCoreApi.bas` (`Core_InsertToPowerPoint`, `mPptPasteOk`), `src/loader/ribbon/customUI14.xml`, `src/loader/modules/modRibbon.bas`; Test `tests/vba/core/TestPptSync.bas`, `tests/vba/loader/TestRibbon.bas`.

**Interfaces:** Produces `PPT_ERR_INSERT_UNSAVED`; `Core_InsertToPowerPoint(kind As String)`; `Core_LastPptPasteOk() As Boolean`; loader `Rx_GetPptInsert`, `Rx_OnPptInsert`, `Rx_SetPptInsertForTest(v As Boolean)`, `Rx_OnPptFrames`. Consumes `Core_Insert`, `Core_CopyToPowerPoint`, `ChartData_FromRange`, `LinkFrame_Ensure/Remove`.

- [ ] **Passo 1: Testes (RED)** — em `TestPptSync.bas`:

```vb
' ---- Parte D2: inserir a partir do intervalo direto no slide ----

Public Sub Test_InsertToPpt_Unsaved_NothingCreated()
    Dim wb As Workbook, rng As Range, n0 As Long
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    n0 = rng.Worksheet.ChartObjects.Count
    rng.Select
    Core_ClearLastNotice
    Core_InsertToPowerPoint "stacked"
    AssertEqual PPT_ERR_INSERT_UNSAVED, Core_LastNotice()
    AssertEqual n0, rng.Worksheet.ChartObjects.Count, "nenhum grafico criado"
    AssertEqual 0, rng.Worksheet.Shapes.Count - rng.Worksheet.ChartObjects.Count, "nenhuma moldura"
    Teardown wb
End Sub

Public Sub Test_InsertToPpt_NoPowerPoint_FrameRemoved()
    Dim wb As Workbook, rng As Range
    Ppt_ResetTestHooks                                   ' build Test sem instancia injetada: Ppt_GetApp devolve Nothing
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\pptins-noppt.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_InsertToPowerPoint "clustered"
    AssertEqual 1, rng.Worksheet.ChartObjects.Count, "o grafico fica no Excel"
    AssertFalse Core_LastPptPasteOk()
    AssertTrue LinkFrame_Find(rng.Worksheet.ChartObjects(1).Chart) Is Nothing, "sem colagem, sem moldura"
    AssertEqual PPT_ERR_NOT_INSTALLED, Core_LastNotice()
    Teardown wb
End Sub

Public Sub Test_InsertToPpt_ChartFrameAndSlide_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String, shp As Object, c As Chart
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Ppt_PinTargetForTest slide
    Ppt_SetAppForTest app
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    wb.SaveAs TempDir() & "\pptins-ok.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_InsertToPowerPoint "waterfall"
    Set c = rng.Worksheet.ChartObjects(1).Chart
    AssertTrue Core_LastPptPasteOk()
    AssertEqual 1, slide.Shapes.Count
    Set shp = slide.Shapes(1)
    AssertEqual Registry_IdFromChart(c), PptSync_Tag(shp, PPTSYNC_TAG_CHART)
    AssertEqual 1, Ppt_LinkState(shp)
    AssertTrue Not LinkFrame_Find(c) Is Nothing, "moldura no intervalo"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub
```
  (A cascata usa A1:C3 com 2 colunas de valores: se o tipo recusar a tabela, use `"stacked"` — o teste confere a integração, não o tipo.)

  Em `TestRibbon.bas`: troque `AssertEqual 3, d.SelectNodes("//u:menu[@id='mnuPpt']/u:button").Length` por `AssertEqual 4, ...`; troque o `AssertEqual 2, ...checkBox...` da C4 por `3`; acrescente:

```vb
Public Sub Test_Ribbon_PptInsertAndFrames()
    Dim d As Object
    Set d = LoadRibbon()
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:checkBox[@id='chkPptInsert'][@getPressed='Rx_GetPptInsert'][@onAction='Rx_OnPptInsert']") Is Nothing
    AssertEqual "chkPptInsert", AttrText(d.SelectSingleNode("//u:menu[@id='mnuPpt']/*[1]"), "id"), "primeiro item do menu"
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:button[@id='btnPptFrames'][@onAction='Rx_OnPptFrames']") Is Nothing
    AssertTrue d.SelectSingleNode("//u:button[@id='btnPptFrames']/@getEnabled") Is Nothing, "sempre ativo"
End Sub

Public Sub Test_Ribbon_OnInsert_RoutesToPpt()
    Rx_SetPptInsertForTest True
    AssertEqual "Core_InsertToPowerPoint", Rx_InsertProcFor()
    Rx_SetPptInsertForTest False
    AssertEqual "Core_Insert", Rx_InsertProcFor()
End Sub
```

- [ ] **Passo 2:** RED (compilação).
- [ ] **Passo 3: Core** — em `modCoreApi.bas`: declare `Private mPptPasteOk As Boolean     ' a ultima Core_CopyToPowerPoint colou (1.7.0)` no topo; em `CopyToPowerPointImpl`, primeira linha depois do `On Error GoTo Fail`: `mPptPasteOk = False`; logo antes de `Core_Notify msg, vbInformation` (sucesso): `mPptPasteOk = True`. Acrescente:

```vb
Public Const PPT_ERR_INSERT_UNSAVED As String = "Salve a pasta de trabalho antes de inserir no PowerPoint: o gráfico vinculado aponta para o arquivo salvo. Nada foi criado."

Public Function Core_LastPptPasteOk() As Boolean
    Core_LastPptPasteOk = mPptPasteOk
End Function

' "Inserir também no slide" (1.7.0; o "Link To PowerPoint" do think-cell): com a tabela selecionada numa pasta SALVA, insere o grafico do tipo
' no Excel, desenha a moldura verde no intervalo de origem e cola o grafico vinculado no slide ativo (Core_CopyToPowerPoint "chart": etiquetas,
' setas no tamanho do slide e "Novos vinculos atualizam automaticamente"). Pasta nao salva: aviso e NADA criado. Colagem que falhou: o grafico
' fica no Excel, sem moldura, com o aviso da colagem. Nunca propaga erro.
Public Sub Core_InsertToPowerPoint(kind As String)
    Dim rng As Range, c As Chart, data As clsChartData, src As Range
    On Error GoTo Fail
    If TypeName(Selection) <> "Range" Then Core_Insert kind: Exit Sub        ' o aviso de selecao e o do Core_Insert
    Set rng = Selection
    If Len(rng.Worksheet.Parent.Path) = 0 Then Core_Notify PPT_ERR_INSERT_UNSAVED, vbExclamation: Exit Sub
    Set data = ChartData_FromRange(rng)
    Set src = data.Source
    Core_Insert kind
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub                                            ' Core_Insert recusou e avisou
    LinkFrame_Ensure c, src
    c.Parent.Activate
    Core_CopyToPowerPoint "chart"
    If Not mPptPasteOk Then LinkFrame_Remove c
    Exit Sub
Fail:
    Errors_Handle "Core_InsertToPowerPoint(" & kind & ")"
End Sub
```
  (Confira em `clsChartData` o nome da propriedade do intervalo de origem — `Source` é usada por `Core_Insert` (`data.Source.Worksheet`).)

- [ ] **Passo 4: Faixa** — no `mnuPpt`, como PRIMEIRO filho:

```xml
            <checkBox id="chkPptInsert" label="Inserir também no slide" getPressed="Rx_GetPptInsert" onAction="Rx_OnPptInsert"
                      screentip="Inserir também no slide" supertip="Marcada, cada gráfico inserido pelos menus Colunas, Barras, Cascata, Linha e área e Mais tipos também vai para o slide ativo do PowerPoint como gráfico vinculado, e o intervalo de origem ganha uma moldura verde. A pasta precisa estar salva."/>
```
  e depois de `chkPptAutoNew`:

```xml
            <button id="btnPptFrames" label="Mostrar/ocultar molduras" image="tqq_ppt16" onAction="Rx_OnPptFrames"
                    screentip="Mostrar/ocultar molduras" supertip="Mostra ou oculta as molduras verdes dos intervalos de origem dos gráficos inseridos no slide (a moldura não é impressa)."/>
```
  Em `modRibbon.bas`: `Private mPptInsert As Boolean   ' "Inserir tambem no slide" (sessao)` no topo e:

```vb
Public Sub Rx_SetPptInsertForTest(v As Boolean)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mPptInsert = v
End Sub

Public Function Rx_InsertProcFor() As String
    If mPptInsert Then Rx_InsertProcFor = "Core_InsertToPowerPoint" Else Rx_InsertProcFor = "Core_Insert"
End Function

Public Sub Rx_GetPptInsert(control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    pressed = mPptInsert
End Sub

Public Sub Rx_OnPptInsert(control As IRibbonControl, pressed As Boolean)
    On Error Resume Next
    mPptInsert = pressed
End Sub

Public Sub Rx_OnPptFrames(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_ToggleLinkFrames"
End Sub
```
  e troque o corpo de `Rx_OnInsert` por:

```vb
Public Sub Rx_OnInsert(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then
        If mPptInsert Then RunCoreFeature "Core_InsertToPowerPoint", control.Tag Else Loader_RunCore "Core_Insert", control.Tag
    End If
    Loader_InvalidateRibbon                 ' o grafico novo fica selecionado: o tipo mudou
End Sub
```
- [ ] **Passo 5 (GREEN):** `-Target all -Filter "TestPptSync|TestLinkFrame|TestRibbon"` → 0 falhas; `Test-Ribbon.ps1` ok; suíte completa com `-Ppt`. Commit `feat: inserir o grafico direto no slide a partir do intervalo selecionado` (+ trailer).

### Tarefa D3: Documentação da 1.7.0 — modelo: standard
Tutorial `#ppt-inserir` (fluxo, moldura, pasta salva, mostrar/ocultar, limites: a moldura não muda o vínculo; mudar intervalo = Alterar dados; só gráficos inseridos por este caminho ganham moldura); README (Recursos, Limitações, manuais 69–73: clicar dentro da moldura (P4), imprimir, inserir/apagar linhas, copiar a planilha, PowerPoint fechado); CHANGELOG/notes; Test-Docs (rótulos `Inserir também no slide`, `Mostrar/ocultar molduras`, âncora). Commit `docs: inserir direto no slide e moldura do intervalo` (+ trailer).

### Tarefa D4: Suíte final, revisão do ramo, verificação manual — modelo: standard (controlador + humano)
Suíte completa com `-Ppt`; revisão final (Review Focus 3); onda de correção; humano: itens 69–73, inserir 5 tipos variados, PR, `publicar-release 1.7.0`.

---

# PARTE G — Criar gráficos pelo PowerPoint (emenda 1; fase G-I = 1.8.0, tarefas G1–G6; fase G-II = 1.9.0, tarefas G7–G11)

Spec §14. Pré-requisitos: Parte C mesclada (automático, índice, `PptSync_Items`, `PptSync_ResolveEntries`, API 2) e D1 (moldura). **Antes da G2:** o controlador agenda as sondas pendentes PG3 e PG4 (spec §3.1); **antes da G4:** PG1 e PG5. Cada resultado vai ao ledger; os caminhos alternativos estão descritos em cada tarefa.

### Tarefa G1: Dados de exemplo dos 21 tipos (`Samples_TableFor`) — modelo: standard

**Files:** Modify `src/core/modules/modSamples.bas`; Test `tests/vba/core/TestSamples.bas`.

**Interfaces:** Produces `Samples_TableFor(kind As String) As Variant` (Array de linhas; `Empty` para tipo desconhecido); `Samples_WriteTable(topLeft As Range, tbl As Variant) As Range`. Consumes `YearTable`, `WriteTable` (privadas do módulo).

- [ ] **Passo 1: Teste (RED)** — acrescente a `tests/vba/core/TestSamples.bas`:

```vb
' Parte G (1.8.0): cada um dos 21 tipos tem uma tabela de exemplo valida e o Core_Insert cria o grafico a partir dela.
Public Sub Test_Samples_TableFor_AllKindsInsert()
    Dim wb As Workbook, ws As Worksheet, k As Variant, rng As Range, n As Long, c As Chart
    Set wb = NewTempWorkbook()
    For Each k In Split("stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,gantt,pie,doughnut,butterfly,scatter,bubble,waterfallpair", ",")
        AssertTrue Kinds_IsValid(CStr(k)), "tipo da lista: " & k
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        Set rng = Samples_WriteTable(ws.Range("A1"), Samples_TableFor(CStr(k)))
        AssertEqual "", ChartData_FromRange(rng).Validate(), "tabela valida: " & k
        ws.Activate
        rng.Select
        Core_ClearLastNotice
        Core_Insert CStr(k)
        Set c = Core_ActiveChart()
        AssertTrue Not c Is Nothing, "grafico criado: " & k & " (" & Core_LastNotice() & ")"
        AssertEqual CStr(k), Registry_ChartType(c)
        n = n + 1
    Next
    AssertEqual 21, n
    AssertTrue IsEmpty(Samples_TableFor("naoexiste"))
    Watchers_Stop
    CloseTemp wb
End Sub
```

- [ ] **Passo 2:** `-Target core -Filter "TestSamples"` → saída 3. RED.
- [ ] **Passo 3:** acrescente a `src/core/modules/modSamples.bas`:

```vb
' Escreve a tabela (Array de linhas; a 1a e o cabecalho, em negrito) a partir de topLeft e devolve o intervalo escrito.
Public Function Samples_WriteTable(topLeft As Range, tbl As Variant) As Range
    Dim r As Variant
    WriteTable topLeft, tbl
    r = tbl(LBound(tbl))
    Set Samples_WriteTable = topLeft.Resize(UBound(tbl) - LBound(tbl) + 1, UBound(r) - LBound(r) + 1)
End Function

' Tabela de exemplo de cada tipo, no formato que o ChartData do tipo exige (spec 2026-10-15 secao 14.7): o "Inserir grafico" do PowerPoint
' cria o grafico com ela e o usuario troca os numeros. Empty para tipo desconhecido. Pura.
Public Function Samples_TableFor(kind As String) As Variant
    Select Case kind
        Case "stacked", "clustered", "stacked100", "barstacked", "barclustered", "barstacked100", "area", "area100", "line"
            Samples_TableFor = YearTable()
        Case "waterfall", "barwaterfall"
            Samples_TableFor = Array(Array("Etapa", "Valor"), Array("Receita", 1200), Array("Custos", -700), Array("Despesas", -250), _
                                     Array("EBITDA", "e"), Array("Impostos", -80), Array("Lucro", "e"))
        Case "waterfallpair"
            Samples_TableFor = Array(Array("Etapa", "2025", "2026"), Array("Receita", 1200, 1350), Array("Custos", -700, -760), _
                                     Array("Despesas", -250, -240), Array("Lucro", "e", "e"))
        Case "combo"
            Samples_TableFor = Array(Array("Ano", "Loja", "Online", "Margem %"), Array(2023, 120, 40, 0.18), Array(2024, 130, 55, 0.2), _
                                     Array(2025, 138, 72, 0.22))
        Case "mekko"
            Samples_TableFor = Array(Array("Região", "Eletro", "Móveis"), Array("Sul", 40, 10), Array("Sudeste", 20, 10), Array("Norte", 10, 10))
        Case "mekkounits"
            Samples_TableFor = Array(Array("Região", "Largura", "Eletro", "Móveis"), Array("Sul", 50, 40, 10), Array("Sudeste", 30, 20, 10), _
                                     Array("Norte", 20, 10, 10))
        Case "gantt"
            Samples_TableFor = Array(Array("Atividade", "Início", "Fim"), Array("Planejamento", DateSerial(2026, 1, 5), DateSerial(2026, 1, 30)), _
                                     Array("Execução", DateSerial(2026, 2, 2), DateSerial(2026, 3, 27)), Array("Entrega", DateSerial(2026, 3, 30), Empty))
        Case "pie", "doughnut"
            Samples_TableFor = Array(Array("Canal", "Vendas"), Array("Loja", 55), Array("Online", 30), Array("Atacado", 15))
        Case "butterfly"
            Samples_TableFor = Array(Array("Faixa", "Mulheres", "Homens"), Array("18-24", 12, 10), Array("25-34", 20, 18), _
                                     Array("35-44", 16, 17), Array("45+", 9, 11))
        Case "scatter"
            Samples_TableFor = Array(Array("Loja", "Área (m²)", "Vendas"), Array("A", 120, 300), Array("B", 200, 410), Array("C", 90, 180), _
                                     Array("D", 260, 520))
        Case "bubble"
            Samples_TableFor = Array(Array("Loja", "Área (m²)", "Vendas", "Clientes"), Array("A", 120, 300, 40), Array("B", 200, 410, 65), _
                                     Array("C", 90, 180, 25), Array("D", 260, 520, 80))
    End Select
End Function
```
  (As continuações ` _` acima terminam a linha — permitidas pelo lint; o que é proibido é texto depois do ` _`. Se a "Entrega" com fim vazio (marco) for recusada pelo `Validate` do Gantt, use uma data de fim e relate.)
- [ ] **Passo 4 (GREEN):** `-Filter "TestSamples"` → 0 falhas; suíte core. Commit `feat: dados de exemplo para os 21 tipos de grafico` (+ trailer).

---

### Tarefa G2: Pasta de dados e `Core_PptInsertChart` (API 3) — modelo: most-capable

**Pré-requisito:** sondas PG3 (colagem dentro de um `xl.Run` vindo do PowerPoint) e PG6 no ledger. PG3 negativa ⇒ ver o Passo 6b.

**Files:** Create `src/core/modules/modPptData.bas`; Modify `src/core/modules/modCoreApi.bas` (`Core_PptInsertChart`), `src/core/modules/modPptSync.bas` (`CORE_PPT_API = 3`); Test `tests/vba/core/TestPptData.bas` (novo), `tests/vba/core/TestPptSync.bas`.

**Interfaces:** Produces `PPTDATA_SUFFIX`, `PPTDATA_MARK = "TQQ_DATABOOK"`, `PPTDATA_TAG = "TQQ_DATA"`, `PPTDATA_SHEET_PREFIX`, `PPTDATA_MSG_UNSAVED`, `PPTDATA_MSG_WEB`, `PPTDATA_MSG_RO`; `PptData_PathFor(presFullName As String) As String`; `PptData_NextSheetName(names As String) As String`; `PptData_IsDataBook(wb As Workbook) As Boolean`; `PptData_OpenForEdit(path As String) As Workbook`; `PptData_OpenOrCreate(path As String, presName As String) As Workbook`; `PptData_NewChart(wb As Workbook, kind As String, ByRef rng As Range) As ChartObject`; `Core_PptInsertChart(slide As Object, kind As String) As String`. Consumes `Samples_TableFor/WriteTable`, `Core_Insert`, `Ppt_PasteChart`, `PptSyncAfterPaste` (privada do `modCoreApi`), `PptSync_SetAuto`, `LinkFrame_Ensure`.

- [ ] **Passo 1: Testes (RED)** — crie `tests/vba/core/TestPptData.bas`:

```vb
Attribute VB_Name = "TestPptData"
Option Explicit
' Parte G (1.8.0): pasta de dados dos graficos criados pelo PowerPoint (sem PowerPoint).

Private Function SheetExists(wb As Workbook, nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = wb.Sheets(nm)
    SheetExists = Not s Is Nothing
End Function

Public Sub Test_PptData_PathFor()
    AssertEqual "C:\a\Vendas - dados ThinQQ.xlsx", PptData_PathFor("C:\a\Vendas.pptx")
    AssertEqual "C:\a\b.c - dados ThinQQ.xlsx", PptData_PathFor("C:\a\b.c.pptx"), "so a ultima extensao sai"
    AssertEqual "\\srv\x\R - dados ThinQQ.xlsx", PptData_PathFor("\\srv\x\R.pptx")
    AssertEqual "", PptData_PathFor("https://empresa.sharepoint.com/x/R.pptx"), "endereco da Web: recusado"
    AssertEqual "", PptData_PathFor("Apresentação1"), "nunca salva: sem pasta"
End Sub

Public Sub Test_PptData_NextSheetName()
    AssertEqual "Gráfico 1", PptData_NextSheetName("|Leia-me|")
    AssertEqual "Gráfico 3", PptData_NextSheetName("|Leia-me|Gráfico 1|Gráfico 2|")
    AssertEqual "Gráfico 2", PptData_NextSheetName("|Gráfico 1|Gráfico 3|"), "primeiro livre"
End Sub

Public Sub Test_PptData_CreateReopenAndNewChart()
    Dim p As String, wb As Workbook, co As ChartObject, rng As Range
    p = TempDir() & "\Teste G - dados ThinQQ.xlsx"
    If Len(Dir(p)) > 0 Then Kill p
    Set wb = PptData_OpenOrCreate(p, "Teste G.pptx")
    AssertTrue PptData_IsDataBook(wb), "marcada"
    AssertEqual "Leia-me", wb.Worksheets(1).Name
    AssertTrue Len(Dir(p)) > 0, "salva no caminho"
    Set co = PptData_NewChart(wb, "stacked", rng)
    AssertEqual "Gráfico 1", co.Parent.Name
    AssertEqual "$A$1:$D$5", rng.Address, "tabela de exemplo (Ano x 3 regioes, 4 anos)"
    AssertEqual "stacked", Registry_ChartType(co.Chart)
    Set co = PptData_NewChart(wb, "pie", rng)
    AssertEqual "Gráfico 2", co.Parent.Name
    wb.Save
    AssertTrue PptData_OpenOrCreate(p, "Teste G.pptx") Is wb, "ja aberta: a mesma"
    Watchers_Stop
    wb.Close SaveChanges:=False
    Set wb = PptData_OpenForEdit(p)
    AssertFalse wb.ReadOnly
    AssertTrue SheetExists(wb, "Gráfico 1") And SheetExists(wb, "Gráfico 2"), "reaberta com os dois graficos"
    wb.Close SaveChanges:=False
    Kill p
End Sub
```
  E em `TestPptSync.bas` (depois dos testes da C2, que definem `AutoReset`):

```vb
' ---- Parte G: grafico criado pelo PowerPoint ----

Public Sub Test_PptInsertChart_UnsavedPres_NothingCreated_WhenAvailable()
    Dim app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    If UserWindowAppeared(app) Then GoTo Cleanup
    AssertEqual PPTDATA_MSG_UNSAVED, Core_PptInsertChart(slide, "stacked")
    AssertEqual 0, slide.Shapes.Count
    AssertEqual "Tipo de gráfico desconhecido: xyz", Core_PptInsertChart(slide, "xyz")
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 6: criar pelo PowerPoint, digitar no "datasheet" (planilha de dados), o ciclo leva ao slide e salva a pasta.
Public Sub Test_PptInsertChart_FromPpt_WhenAvailable()
    Dim app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String, msg As String
    Dim shp As Object, s2 As Object, p As String, pp As String, wb As Workbook, ws As Worksheet
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoReset
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    If UserWindowAppeared(app) Then GoTo Cleanup
    pp = TempDir() & "\Vendas G.pptx": p = TempDir() & "\Vendas G - dados ThinQQ.xlsx"
    If Len(Dir(p)) > 0 Then Kill p
    pres.SaveAs pp
    msg = Core_PptInsertChart(slide, "clustered")
    AssertEqual "", msg
    AssertEqual 1, slide.Shapes.Count
    Set shp = slide.Shapes(1)
    AssertEqual "1", PptSync_Tag(shp, PPTDATA_TAG)
    AssertEqual "1", PptSync_Tag(shp, PPTSYNC_TAG_AUTO), "nasce automatico (datasheet ao vivo)"
    AssertEqual 1, Ppt_LinkState(shp)
    AssertEqual LCase$(p), LCase$(PptSync_ItemBook(shp))
    AssertEqual "Gráfico 1", PptSync_Tag(shp, PPTSYNC_TAG_SHEET)
    AssertEqual LCase$(p), LCase$(PptSync_Tag(pres, PPTDATA_MARK))
    Set wb = Workbooks(Dir(p))
    AssertTrue wb.Saved, "pasta de dados salva"
    Set ws = wb.Worksheets("Gráfico 1")
    AssertTrue Not LinkFrame_Find(ws.ChartObjects(1).Chart) Is Nothing, "dados destacados pela moldura"
    Watchers_Start
    ws.Range("B2").Value = 999
    AssertTrue PptAuto_PendingCountForTest() >= 1
    Ppt_SetAppForTest app
    PptAuto_Tick
    AssertAlmost 999, PptSeriesValue(shp.Chart, "Sul", 1), 0.0001, "o slide acompanhou a planilha"
    AssertTrue wb.Saved, "o ciclo salvou a pasta de dados (G3)"
    AssertEqual "", Core_PptInsertChart(slide, "pie")
    Set s2 = slide.Shapes(2)
    AssertEqual "Gráfico 2", PptSync_Tag(s2, PPTSYNC_TAG_SHEET)
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set shp = Nothing: Set s2 = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    Watchers_Stop
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Kill p: Kill pp
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub
```
  (A asserção "o ciclo salvou a pasta" só passa depois da G3; na G2 marque-a com `' G3:` e ligue-a na G3.)

- [ ] **Passo 2:** RED (compilação).
- [ ] **Passo 3:** crie `src/core/modules/modPptData.bas`:

```vb
Attribute VB_Name = "modPptData"
Option Explicit
' Graficos criados PELO POWERPOINT (1.8.0; "Insert > Elements" do think-cell; spec 2026-10-15 secao 14). Pasta de dados ao lado da
' apresentacao ("<nome> - dados ThinQQ.xlsx", sem macros, marcada pelo nome oculto TQQ_DATABOOK), uma planilha por grafico ("Gráfico N":
' tabela em A1 e o grafico ThinQQ ao lado). O grafico vai ao slide VINCULADO (o motor da 1.4.0) e nasce automatico (1.6.0). Aberta para
' edicao sempre com as macros desligadas (AutomationSecurity = 3). Sem processo filho.
Public Const PPTDATA_SUFFIX As String = " - dados ThinQQ.xlsx"
Public Const PPTDATA_MARK As String = "TQQ_DATABOOK"
Public Const PPTDATA_TAG As String = "TQQ_DATA"
Public Const PPTDATA_SHEET_PREFIX As String = "Gráfico "
Public Const PPTDATA_MSG_UNSAVED As String = "Salve a apresentação primeiro: os dados do gráfico ficam numa pasta do Excel ao lado dela (""<nome> - dados ThinQQ.xlsx"")."
Public Const PPTDATA_MSG_WEB As String = "A apresentação está aberta por um endereço da Web (OneDrive/SharePoint). Salve-a numa pasta do computador (ou na pasta sincronizada do OneDrive) para inserir gráficos do ThinQQ."
Public Const PPTDATA_MSG_RO As String = "A pasta de dados está aberta por outra pessoa ou é somente leitura: "

' Caminho da pasta de dados de uma apresentacao (FullName): mesma pasta, nome sem a ultima extensao + PPTDATA_SUFFIX. "" para endereco
' http(s) ou apresentacao sem pasta (nunca salva). Pura.
Public Function PptData_PathFor(presFullName As String) As String
    Dim p As Long, k As Long, base As String, f As String
    f = Trim$(presFullName)
    If LCase$(Left$(f, 7)) = "http://" Or LCase$(Left$(f, 8)) = "https://" Then Exit Function
    p = InStrRev(f, "\")
    If p < 2 Then Exit Function
    base = Mid$(f, p + 1)
    k = InStrRev(base, ".")
    If k > 1 Then base = Left$(base, k - 1)
    PptData_PathFor = Left$(f, p - 1) & "\" & base & PPTDATA_SUFFIX
End Function

' Primeiro "Gráfico N" livre; names = nomes das planilhas como "|a|b|". Pura.
Public Function PptData_NextSheetName(names As String) As String
    Dim i As Long
    For i = 1 To 999
        If InStr(1, names, "|" & PPTDATA_SHEET_PREFIX & i & "|", vbTextCompare) = 0 Then PptData_NextSheetName = PPTDATA_SHEET_PREFIX & i: Exit Function
    Next
End Function

Private Function PptData_SheetNames(wb As Workbook) As String
    Dim ws As Object, s As String
    s = "|"
    For Each ws In wb.Sheets
        s = s & ws.Name & "|"
    Next
    PptData_SheetNames = s
End Function

' A pasta tem a marca TQQ_DATABOOK. Nunca propaga erro.
Public Function PptData_IsDataBook(wb As Workbook) As Boolean
    Dim n As Name
    On Error Resume Next
    Set n = wb.Names(PPTDATA_MARK)
    PptData_IsDataBook = Not n Is Nothing
    Err.Clear
End Function

' Abre para EDICAO (macros desligadas e eventos desligados durante o Open, sem atualizar vinculos, sem MRU, sem a caixa de senha) ou devolve a
' ja aberta neste Excel. Somente leitura (aberta por outra pessoa) = erro 7130 (fechada antes). Propaga erro.
Public Function PptData_OpenForEdit(path As String) As Workbook
    Dim w As Workbook, prevSec As Long, prevAlerts As Boolean, prevEv As Boolean, en As Long, ed As String
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then Set PptData_OpenForEdit = w: Exit Function
    Next
    prevAlerts = Application.DisplayAlerts: prevEv = Application.EnableEvents: prevSec = Application.AutomationSecurity
    Application.DisplayAlerts = False: Application.EnableEvents = False
    Application.AutomationSecurity = 3                          ' msoAutomationSecurityForceDisable
    On Error Resume Next
    Set w = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=False, Password:="x", IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    en = Err.Number: ed = Err.Description
    Err.Clear
    If en = 0 And Not w Is Nothing Then
        If w.ReadOnly Then w.Close SaveChanges:=False: en = vbObjectError + 7130: ed = PPTDATA_MSG_RO & path
    End If
    On Error GoTo 0
    Application.AutomationSecurity = prevSec
    Application.DisplayAlerts = prevAlerts: Application.EnableEvents = prevEv
    If en <> 0 Then Err.Raise en, "PptData_OpenForEdit", ed
    Set PptData_OpenForEdit = w
End Function

' A pasta de dados da apresentacao: aberta para edicao se existe; senao criada ("Leia-me" + marca) e salva como .xlsx. Propaga erro.
Public Function PptData_OpenOrCreate(path As String, presName As String) As Workbook
    Dim w As Workbook, prevAlerts As Boolean
    If CBool(CreateObject("Scripting.FileSystemObject").FileExists(path)) Then Set PptData_OpenOrCreate = PptData_OpenForEdit(path): Exit Function
    Set w = Application.Workbooks.Add(xlWBATWorksheet)
    w.Worksheets(1).Name = "Leia-me"
    w.Worksheets(1).Range("A1").Value = "Dados dos gráficos do ThinQQ da apresentação '" & presName & "'."
    w.Worksheets(1).Range("A2").Value = "Cada planilha 'Gráfico N' alimenta um gráfico do slide. Envie este arquivo junto com a apresentação (PowerPoint: ThinQQ > Empacotar para enviar)."
    w.Names.Add Name:=PPTDATA_MARK, RefersTo:="=""1""", Visible:=False
    prevAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False
    w.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook, AddToMru:=False
    Application.DisplayAlerts = prevAlerts
    Set PptData_OpenOrCreate = w
End Function

' Planilha nova "Gráfico N" com a tabela de exemplo do tipo e o grafico ThinQQ inserido pelo MESMO caminho do usuario (Core_Insert: formato
' inicial, avisos, observadores). rng volta com o intervalo dos dados. A planilha criada e apagada se o grafico nao nascer. Propaga erro
' (7131 tipo sem exemplo; 7132 o Core_Insert nao criou o grafico).
Public Function PptData_NewChart(wb As Workbook, kind As String, ByRef rng As Range) As ChartObject
    Dim ws As Worksheet, tbl As Variant, c As Chart, prevAlerts As Boolean, msg As String
    tbl = Samples_TableFor(kind)
    If IsEmpty(tbl) Then Err.Raise vbObjectError + 7131, "PptData", "Tipo de gráfico desconhecido: " & kind
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = PptData_NextSheetName(PptData_SheetNames(wb))
    Set rng = Samples_WriteTable(ws.Range("A1"), tbl)
    ws.Columns(1).ColumnWidth = 16
    wb.Activate
    ws.Activate
    rng.Select
    Core_ClearLastNotice
    Core_Insert kind
    Set c = Core_ActiveChart()
    If c Is Nothing Then
        msg = Core_LastNotice()
        prevAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = prevAlerts
        Err.Raise vbObjectError + 7132, "PptData", "Não foi possível criar o gráfico na pasta de dados: " & msg
    End If
    Set PptData_NewChart = c.Parent
End Function
```
  (`Core_ClearLastNotice` zera também os contadores usados pelos testes — aceitável: é a mesma chamada que os testes fazem.)

- [ ] **Passo 4: API** — `modPptSync.bas`: `CORE_PPT_API = 3` (comentário "3 = Core_PptInsertChart, Core_PptEditData, Core_PptDataSave (1.8.0)"). `modCoreApi.bas`, depois de `Core_PptGetFlag/SetFlag`:

```vb
' API 3 (1.8.0): grafico do ThinQQ criado PELO POWERPOINT (spec 14). slide = slide de destino (objeto do PowerPoint, passado pelo .ppam).
' Cria a planilha na pasta de dados da apresentacao, insere o grafico, destaca os dados (moldura), salva, cola VINCULADO no slide (o mesmo
' caminho do "Copiar como grafico vinculado": etiquetas e setas no tamanho do slide), marca TQQ_DATA e liga "Atualizar automaticamente".
' Devolve "" (ok) ou a mensagem; falha desfaz a forma e a planilha criadas. Nunca propaga erro.
Public Function Core_PptInsertChart(slide As Object, kind As String) As String
    Dim pres As Object, path As String, wb As Workbook, co As ChartObject, rng As Range, shp As Object, ws As Worksheet
    Dim en As Long, ed As String, prevAlerts As Boolean
    On Error GoTo Fail
    If Not Kinds_IsValid(kind) Then Core_PptInsertChart = "Tipo de gráfico desconhecido: " & kind: Exit Function
    Set pres = slide.Parent
    If Len(CStr(pres.Path)) = 0 Then Core_PptInsertChart = PPTDATA_MSG_UNSAVED: Exit Function
    path = PptData_PathFor(CStr(pres.FullName))
    If Len(path) = 0 Then Core_PptInsertChart = PPTDATA_MSG_WEB: Exit Function
    Set wb = PptData_OpenOrCreate(path, CStr(pres.Name))
    Set co = PptData_NewChart(wb, kind, rng)
    Set ws = co.Parent
    LinkFrame_Ensure co.Chart, rng
    wb.Save
    Set shp = Ppt_PasteChart(co, slide)
    On Error Resume Next: Application.CutCopyMode = False: Err.Clear: On Error GoTo Fail
    If Ppt_LinkState(shp) <> 1 Then Err.Raise vbObjectError + 7133, "PptData", "O PowerPoint colou o gráfico sem o vínculo com a pasta de dados; tente de novo."
    PptSyncAfterPaste shp, co
    shp.Tags.Add PPTDATA_TAG, "1"
    PptSync_SetAuto shp, True
    pres.Tags.Add PPTDATA_MARK, path
    Core_PptInsertChart = ""
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Application.CutCopyMode = False
    If Not shp Is Nothing Then shp.Delete
    If Not ws Is Nothing Then
        prevAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = prevAlerts
        wb.Save
    End If
    Errors_Log "Core_PptInsertChart(" & kind & ")" & vbTab & en & vbTab & ed
    If en >= vbObjectError + 7130 And en <= vbObjectError + 7133 Then
        Core_PptInsertChart = ed
    ElseIf Ppt_IsKnownFailure(en) Then
        Core_PptInsertChart = Ppt_ErrorText(en, ed)
    Else
        Core_PptInsertChart = "Não foi possível criar o gráfico: " & ed & " (" & en & ")"
    End If
End Function
```
- [ ] **Passo 5 (GREEN):** `-Filter "TestPptData|TestPptSync|TestCoreApi|TestSamples"` → 0 falhas (PowerPoint livre: os `_WhenAvailable` rodam). **Passo 6b (só se PG3 negativa):** `Core_PptInsertChart` deixa de colar: devolve `"=" & path & Chr(31) & ws.Name & Chr(31) & co.Name` e o `.ppam` (G4) faz `xl.Workbooks(...).Worksheets(...).ChartObjects(...).Copy` + `slide.Shapes.Paste` e chama `Core_PptAfterInsert(shp)` (nova API que faz `PptSyncAfterPaste` + etiquetas); ajuste os testes.
- [ ] **Passo 7:** suíte completa com `-Ppt`. Commit `feat: grafico do ThinQQ criado pelo PowerPoint numa pasta de dados ao lado da apresentacao` (+ trailer).

---

### Tarefa G3: "Editar dados" e salvar a pasta de dados — modelo: standard

**Files:** Modify `src/core/modules/modCoreApi.bas` (`Core_PptEditData`, `Core_PptDataSave`), `src/core/modules/modPptAuto.bas` (salvar no ciclo); Test `tests/vba/core/TestPptSync.bas`.

**Interfaces:** Produces `Core_PptEditData(pres As Object, entries As String) As String` (legenda da janela do Excel, ou `"!" & mensagem`); `Core_PptDataSave(pres As Object) As Long`; `PptAuto_SaveDataBooks` (privada). Consumes `PptSync_ResolveEntries`, `PptSync_ItemBook`, `PptSync_FileExists`, `PptData_OpenForEdit`, `PptData_IsDataBook`, `LinkFrame_Find`, `Registry_FindChartObject`.

- [ ] **Passo 1: Teste (RED)** — em `TestPptSync.bas`, ligue a asserção `' G3:` da G2 e acrescente ao fim de `Test_PptInsertChart_FromPpt_WhenAvailable` (antes de `Cleanup:`):

```vb
    msg = Core_PptEditData(pres, PptIdx_Entry(CLng(slide.SlideID), CStr(shp.Name), ""))
    AssertTrue Left$(msg, 1) <> "!", msg
    AssertEqual "Gráfico 1", ActiveSheet.Name, "Editar dados ativa a planilha do grafico"
    AssertEqual "$A$1:$D$5", Selection.Address, "e seleciona os dados"
    ws.Range("C2").Value = 7
    AssertEqual 1, Core_PptDataSave(pres), "Salvar no PowerPoint salva a pasta de dados"
    AssertTrue wb.Saved
    AssertEqual "!" & PPTSYNC_NONE, Core_PptEditData(pres, PptIdx_Entry(999999, "nada", ""))
```
- [ ] **Passo 2: Código** — `modCoreApi.bas`:

```vb
' API 3: "Editar dados" (spec 14.3). Abre a pasta da forma (a 1a das entradas) PARA EDICAO com as macros desligadas (de dados ou do usuario; ja
' aberta = como esta), ativa a planilha do grafico e seleciona os dados (o intervalo da moldura ou a regiao de A1); fora do build Test deixa o
' Excel visivel. Devolve a legenda da janela (o .ppam faz AppActivate) ou "!" + a mensagem. Nunca propaga erro.
Public Function Core_PptEditData(pres As Object, entries As String) As String
    Dim items As Collection, shp As Object, path As String, wb As Workbook, co As ChartObject, rng As Range, fr As Shape, en As Long, ed As String
    On Error GoTo Fail
    Set items = PptSync_ResolveEntries(pres, entries)
    If items.Count = 0 Then Core_PptEditData = "!" & PPTSYNC_NONE: Exit Function
    Set shp = items(1)
    path = PptSync_ItemBook(shp)
    If Not PptSync_FileExists(path) Then Core_PptEditData = "!Pasta de dados não encontrada: " & path & ". Se ela foi movida, ponha-a na mesma pasta da apresentação.": Exit Function
    Set wb = PptData_OpenForEdit(path)
    Set co = Registry_FindChartObject(wb, PptSync_Tag(shp, PPTSYNC_TAG_CHART))
    If co Is Nothing Then Core_PptEditData = "!Gráfico não encontrado na pasta de dados (apagado ou recriado).": Exit Function
    If BUILD_CONFIG <> "Test" Then
        Application.Visible = True
        If Application.WindowState = xlMinimized Then Application.WindowState = xlNormal
    End If
    wb.Activate
    co.Parent.Activate
    Set fr = LinkFrame_Find(co.Chart)
    If Not fr Is Nothing Then
        Set rng = fr.TopLeftCell.CurrentRegion
    ElseIf PptData_IsDataBook(wb) Then
        Set rng = co.Parent.Range("A1").CurrentRegion
    End If
    If Not rng Is Nothing Then rng.Select
    If LCase$(Right$(path, 5)) = ".xlsm" And Not PptData_IsDataBook(wb) Then Core_Notify "As macros desta pasta não rodam quando ela é aberta pelo ThinQQ; abra-a pelo Excel se precisar delas."
    Core_PptEditData = CStr(Application.ActiveWindow.Caption)
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Errors_Log "Core_PptEditData" & vbTab & en & vbTab & ed
    If en = vbObjectError + 7130 Then Core_PptEditData = "!" & ed Else Core_PptEditData = "!Não foi possível abrir os dados: " & ed & " (" & en & ")"
End Function

' API 3: "Salvar" no PowerPoint (evento do .ppam): salva a pasta de DADOS da apresentacao (etiqueta TQQ_DATABOOK) se estiver aberta e
' alterada neste Excel. Nunca abre nada; pastas do usuario nunca. Devolve 1 se salvou. Nunca propaga erro.
Public Function Core_PptDataSave(pres As Object) As Long
    Dim p As String, w As Workbook
    On Error Resume Next
    p = PptSync_Tag(pres, PPTDATA_MARK)
    If Len(p) = 0 Then Exit Function
    For Each w In Application.Workbooks
        If StrComp(w.FullName, p, vbTextCompare) = 0 Then
            If Not w.Saved And Not w.ReadOnly And PptData_IsDataBook(w) Then w.Save: Core_PptDataSave = 1
        End If
    Next
    Err.Clear
End Function
```
  `modPptAuto.bas`: em `PptAuto_Tick`, logo antes de `mDirty.RemoveAll` do caminho normal, acrescente `PptAuto_SaveDataBooks mDirty` e, no fim do módulo:

```vb
' Pastas de DADOS (graficos criados pelo PowerPoint, marca TQQ_DATABOOK) que mudaram: salvas no fim do ciclo, para o arquivo nunca ficar atras
' do slide (spec 14.2). Pastas do usuario nunca sao salvas pelo ThinQQ. O AfterSave cai em PptAuto_OnChange com mInTick = True (ignorado).
Private Sub PptAuto_SaveDataBooks(dirty As Object)
    Dim w As Workbook
    On Error Resume Next
    For Each w In Application.Workbooks
        If dirty.Exists(w.FullName) Then
            If Not w.Saved And Not w.ReadOnly And PptData_IsDataBook(w) Then w.Save
        End If
    Next
    Err.Clear
End Sub
```
- [ ] **Passo 3 (GREEN):** `-Filter "TestPptSync|TestPptData|TestCoreApi"`; suíte completa com `-Ppt`. Commit `feat: editar dados e salvar a pasta de dados dos graficos criados pelo PowerPoint` (+ trailer).

---

### Tarefa G4: Suplemento do PowerPoint — "Inserir gráfico" (21 tipos), "Editar dados", reapontar e salvar — modelo: most-capable

**Pré-requisito:** PG1 (Excel entregue ao usuário) e PG5 (`PresentationSave`) no ledger. PG1 negativa ⇒ manter `mEditXl` (já no código) e relatar; PG5 negativa ⇒ remover o evento (o ciclo automático continua salvando).

**Files:** `src/ppt/modules/modPptTqq.bas`, `src/ppt/modules/modPptRibbon.bas`, `src/ppt/classes/clsPptEvents.cls`, `src/ppt/ribbon/customUI14.xml`; testes `tests/vba/ppt/TestPptAddin.bas`, `tests/ps/Test-Ribbon.ps1`.

**Interfaces:** Produces `PPT_TAG_DATA`, `PPT_TAG_DATABOOK`, `PPT_DATA_SUFFIX`, `PPT_API_DATA = 3`, `PPT_MSG_SAVE_FIRST`, `PPT_MSG_WEB_PRES`, `PPT_MSG_DATA_NOCORE`, `PPT_MSG_OLD_CORE_DATA`; `PptTqq_DataPathFor` (cópia de `PptData_PathFor`), `PptTqq_ActiveSlide(pres)`, `PptTqq_GetEditExcel(ByRef core As String) As Object`, `PptTqq_InsertChartIn(pres, slide, kind)`, `PptTqq_EditShapes(pres, shapes)`, `PptTqq_RelinkData(pres) As Long`, `PptTqq_OnSave(pres)`; callbacks `PptRx_OnInsertChart`, `PptRx_OnEditData`. Consumes core API 3.

- [ ] **Passo 1: Testes (RED)** — `TestPptAddin.bas`:

```vb
' Mesmas cadeias de TestPptData.Test_PptData_PathFor (contrato core x .ppam).
Public Sub Test_PptTqq_DataPathFor()
    AssertEqual "C:\a\Vendas - dados ThinQQ.xlsx", PptTqq_DataPathFor("C:\a\Vendas.pptx")
    AssertEqual "C:\a\b.c - dados ThinQQ.xlsx", PptTqq_DataPathFor("C:\a\b.c.pptx")
    AssertEqual "", PptTqq_DataPathFor("https://empresa.sharepoint.com/x/R.pptx")
    AssertEqual "", PptTqq_DataPathFor("Apresentação1")
End Sub

Public Sub Test_PptTqq_InsertChartIn_UnsavedAndNoExcel()
    Dim p As Object, f As String, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptTqq_ClearLastMsg
    PptTqq_InsertChartIn p, p.Slides(1), "stacked"
    AssertEqual PPT_MSG_SAVE_FIRST, PptTqq_LastMsg(), "apresentacao nunca salva"
    f = Environ$("TEMP") & "\tqq-ppam-g4.pptx"
    p.SaveAs f
    PptTqq_SetExcelForTest Nothing, ""
    PptTqq_ClearLastMsg
    PptTqq_InsertChartIn p, p.Slides(1), "stacked"
    AssertEqual PPT_MSG_NO_EXCEL, PptTqq_LastMsg(), "sem Excel"
    AssertEqual 0, p.Slides(1).Shapes.Count
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    If Len(f) > 0 Then
        If Len(Dir$(f)) > 0 Then Kill f
    End If
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Reapontar: sem a etiqueta ou com a pasta de dados no lugar, nada; forma que nao e grafico vinculado: nada reapontado (erro so registrado).
Public Sub Test_PptTqq_RelinkData_Guards()
    Dim p As Object, s As Object, f As String, d As String, k As Integer, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    AssertEqual 0, PptTqq_RelinkData(p), "sem TQQ_DATABOOK"
    f = Environ$("TEMP") & "\tqq-relink.pptx": p.SaveAs f
    d = Environ$("TEMP") & "\tqq-relink - dados ThinQQ.xlsx"
    k = FreeFile: Open d For Output As #k: Close #k
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    s.Tags.Add PPT_TAG_DATA, "1": s.Tags.Add PPT_TAG_BOOK, "C:\nao\existe\tqq-relink - dados ThinQQ.xlsx"
    p.Tags.Add PPT_TAG_DATABOOK, "C:\nao\existe\tqq-relink - dados ThinQQ.xlsx"
    AssertEqual 0, PptTqq_RelinkData(p), "forma sem vinculo: nao reaponta"
    p.Tags.Add PPT_TAG_DATABOOK, d
    AssertEqual 0, PptTqq_RelinkData(p), "pasta de dados no lugar: barato, nada"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    ClosePres p
    If Len(Dir$(f)) > 0 Then Kill f
    If Len(Dir$(d)) > 0 Then Kill d
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub
```
  Em `tests/ps/Test-Ribbon.ps1`, depois das conferências do PowerPoint:

```powershell
$xx = New-Object Xml.XmlDocument; $xx.Load((Join-Path $RepoRoot 'src\loader\ribbon\customUI14.xml'))
$xns = New-Object Xml.XmlNamespaceManager($xx.NameTable); $xns.AddNamespace('u', 'http://schemas.microsoft.com/office/2009/07/customui')
$excelTags = @($xx.SelectNodes("//u:group[@id='grpInsert']//u:button[@onAction='Rx_OnInsert']", $xns) | ForEach-Object { $_.GetAttribute('tag') })
$pptTags = @($px.SelectNodes("//u:group[@id='grpPptInsert']//u:button[@onAction='PptRx_OnInsertChart']", $pns) | ForEach-Object { $_.GetAttribute('tag') })
Check ($pptTags.Count -eq 21) "ppt: 21 tipos em Inserir gráfico ($($pptTags.Count))"
Check (($excelTags -join ',') -eq ($pptTags -join ',')) 'ppt: mesmos tipos e mesma ordem dos menus do Excel'
Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptElem']/u:button[@id='btnPptEditData'][@onAction='PptRx_OnEditData'][@getEnabled='PptRx_GetElemEnabled']", $pns)) 'ppt: Editar dados'
Check ('grpPptInsert' -eq $px.SelectSingleNode("//u:tab[@id='tabThinQQPpt']/u:group[1]", $pns).GetAttribute('id')) 'ppt: Inserir gráfico é o primeiro grupo'
```
  (Use o nome do parâmetro de raiz que o script já tem — `$RepoRoot` nos demais testes PS.)

- [ ] **Passo 2:** RED (`-Ppt`: compilação do projeto do PowerPoint).
- [ ] **Passo 3: Código do `.ppam`** — `modPptTqq.bas`, constantes (depois das da C3):

```vb
Public Const PPT_TAG_DATA As String = "TQQ_DATA"
Public Const PPT_TAG_DATABOOK As String = "TQQ_DATABOOK"
Public Const PPT_DATA_SUFFIX As String = " - dados ThinQQ.xlsx"
Public Const PPT_API_DATA As Long = 3                       ' Core_PptInsertChart, Core_PptEditData, Core_PptDataSave (core 1.8.0)
Public Const PPT_MSG_SAVE_FIRST As String = "Salve a apresentação primeiro: os dados do gráfico ficam numa pasta do Excel ao lado dela (""<nome> - dados ThinQQ.xlsx"")."
Public Const PPT_MSG_WEB_PRES As String = "A apresentação está aberta por um endereço da Web (OneDrive/SharePoint). Salve-a numa pasta do computador (ou na pasta sincronizada do OneDrive) para inserir gráficos do ThinQQ."
Public Const PPT_MSG_DATA_NOCORE As String = "O ThinQQ não está instalado no Excel deste computador: não é possível criar nem editar gráficos do ThinQQ."
Public Const PPT_MSG_OLD_CORE_DATA As String = "Atualize o ThinQQ no Excel para criar gráficos pelo PowerPoint."
```
  e `Private mEditXl As Object` e `Private mHintShown As Boolean` no topo; ao fim do módulo:

```vb
' ---- 1.8.0 (Parte G): criar graficos pelo PowerPoint. spec 2026-10-15 secao 14. ----

' Copia exata de PptData_PathFor do core (mesmos testes de cadeia). Pura.
Public Function PptTqq_DataPathFor(presFullName As String) As String
    Dim p As Long, k As Long, base As String, f As String
    f = Trim$(presFullName)
    If LCase$(Left$(f, 7)) = "http://" Or LCase$(Left$(f, 8)) = "https://" Then Exit Function
    p = InStrRev(f, "\")
    If p < 2 Then Exit Function
    base = Mid$(f, p + 1)
    k = InStrRev(base, ".")
    If k > 1 Then base = Left$(base, k - 1)
    PptTqq_DataPathFor = Left$(f, p - 1) & "\" & base & PPT_DATA_SUFFIX
End Function

' Slide ativo da apresentacao (modo Normal; senao o selecionado). Nothing em modos de mestre ou sem janela.
Public Function PptTqq_ActiveSlide(pres As Object) As Object
    Dim s As Object, k As Long
    On Error Resume Next
    Set s = Application.ActiveWindow.View.Slide
    k = 0: k = s.SlideIndex
    If k = 0 Then Set s = Application.ActiveWindow.Selection.SlideRange(1): k = s.SlideIndex
    If k = 0 Then Set s = Nothing
    If Not s Is Nothing Then
        If StrComp(CStr(s.Parent.FullName), CStr(pres.FullName), vbTextCompare) <> 0 Then Set s = Nothing
    End If
    Err.Clear
    Set PptTqq_ActiveSlide = s
End Function

Private Sub PptTqq_OpenLoader(xl As Object)
    Dim p As String
    On Error Resume Next
    p = Environ$("LOCALAPPDATA") & "\ThinQQ\ThinQQ.xlam"
    If Len(Dir$(p)) = 0 Then Exit Sub
    xl.Workbooks.Open p
    Err.Clear
End Sub

' Excel para criar/editar dados, que FICA com o usuario (spec 14.3): (1) o criado antes por aqui, se ainda responde; (2) o Excel aberto COM o
' core; (3) senao um novo, VISIVEL, UserControl, com o loader instalado aberto (carrega core e faixa) ou, sem loader, o core pelo caminho de
' sempre. Nunca carrega o ThinQQ num Excel do usuario que nao o tem; nunca da Quit no Excel criado aqui. Ocupado: erro de ocupado.
Public Function PptTqq_GetEditExcel(ByRef core As String) As Object
    Dim xl As Object, st As Long
    core = ""
    If BUILD_CONFIG = "Test" Then
        If mTestXl Is Nothing Then Exit Function
        core = PptTqq_EnsureCore(mTestXl, mTestCreated)
        Set PptTqq_GetEditExcel = mTestXl
        Exit Function
    End If
    If Not mEditXl Is Nothing Then
        If PptTqq_PingState(mEditXl) = 1 Then core = PptTqq_CoreFileName(): Set PptTqq_GetEditExcel = mEditXl: Exit Function
        Set mEditXl = Nothing                                       ' o usuario fechou
    End If
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    Err.Clear
    On Error GoTo 0
    If Not xl Is Nothing Then
        st = PptTqq_PingState(xl)
        If st = 1 Then core = PptTqq_CoreFileName(): Set PptTqq_GetEditExcel = xl: Exit Function
        If st = -1 Then PptTqq_RaiseBusy
    End If
    Set xl = Nothing
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    Err.Clear
    On Error GoTo 0
    If xl Is Nothing Then Exit Function
    Set mEditXl = xl                                                ' sonda PG1: a referencia fica enquanto o PowerPoint estiver aberto
    On Error Resume Next
    xl.Visible = True
    xl.UserControl = True
    Err.Clear
    On Error GoTo 0
    PptTqq_OpenLoader xl
    If PptTqq_PingState(xl) = 1 Then core = PptTqq_CoreFileName() Else core = PptTqq_EnsureCore(xl, True)
    Set PptTqq_GetEditExcel = xl
End Function

' Abre a planilha dos dados da forma no Excel dado e traz a janela a frente (AppActivate; sonda PG2). Mensagem se nao der.
Private Sub PptTqq_EditShape(shp As Object, xl As Object, core As String)
    Dim cap As String, c As New Collection
    c.Add shp
    cap = CStr(xl.Run("'" & core & "'!Core_PptEditData", shp.Parent.Parent, PptTqq_EntriesOf(c)))
    If Left$(cap, 1) = "!" Then PptTqq_Msg Mid$(cap, 2), vbExclamation: Exit Sub
    If BUILD_CONFIG = "Test" Then mLastMsg = "edit:" & cap: Exit Sub
    On Error Resume Next
    AppActivate cap
    If Err.Number <> 0 Then Err.Clear: PptTqq_Msg "O Excel abriu a planilha de dados do gráfico (veja a barra de tarefas)."
End Sub

' Aviso unico por sessao depois do primeiro grafico criado.
Private Sub PptTqq_FirstHint(pres As Object)
    If mHintShown Then Exit Sub
    mHintShown = True
    PptTqq_Msg "Digite os dados na planilha do Excel que abriu: o slide acompanha em cerca de 2 segundos." & vbLf & _
               "Os dados ficam em '" & Mid$(PptTqq_DataPathFor(CStr(pres.FullName)), InStrRev(PptTqq_DataPathFor(CStr(pres.FullName)), "\") + 1) & "', ao lado da apresentação: envie os dois arquivos juntos (ThinQQ > Empacotar para enviar)."
End Sub

' "Inserir grafico" (1.8.0): apresentacao salva (senao a mensagem), slide de destino (o dado; Nothing = um slide novo no fim), Excel para editar
' (fica com o usuario), core com a API 3; o core cria o grafico na pasta de dados e cola no slide; depois abre a planilha dos dados. Nunca propaga.
Public Sub PptTqq_InsertChartIn(pres As Object, slide As Object, kind As String)
    Dim xl As Object, core As String, msg As String, sl As Object, n0 As Long
    If mBusy Then Exit Sub
    On Error GoTo Fail
    If Len(CStr(pres.Path)) = 0 Then PptTqq_Msg PPT_MSG_SAVE_FIRST: Exit Sub
    If Len(PptTqq_DataPathFor(CStr(pres.FullName))) = 0 Then PptTqq_Msg PPT_MSG_WEB_PRES: Exit Sub
    mBusy = True
    Set xl = PptTqq_GetEditExcel(core)
    If xl Is Nothing Then mBusy = False: PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation: Exit Sub
    If Len(core) = 0 Then mBusy = False: PptTqq_Msg PPT_MSG_DATA_NOCORE, vbExclamation: Exit Sub
    If PptTqq_CoreApi(xl, core) < PPT_API_DATA Then mBusy = False: PptTqq_Msg PPT_MSG_OLD_CORE_DATA: Exit Sub
    Set sl = slide
    If sl Is Nothing Then Set sl = pres.Slides.Add(pres.Slides.Count + 1, 12)
    n0 = sl.Shapes.Count
    msg = CStr(xl.Run("'" & core & "'!Core_PptInsertChart", sl, kind))
    mBusy = False
    If Len(msg) > 0 Then PptTqq_Msg msg, vbExclamation: Exit Sub
    If sl.Shapes.Count > n0 Then PptTqq_EditShape sl.Shapes(sl.Shapes.Count), xl, core
    PptRx_InvalidateElem
    PptTqq_FirstHint pres
    Exit Sub
Fail:
    mBusy = False
    PptTqq_Msg PptTqq_ErrorText(Err.Number, Err.Description), vbExclamation
End Sub

' "Editar dados" das formas dadas (a 1a): reaponta antes (apresentacao movida), Excel para editar, Core_PptEditData. Nunca propaga.
Public Sub PptTqq_EditShapes(pres As Object, shapes As Collection)
    Dim xl As Object, core As String
    If mBusy Then Exit Sub
    If shapes.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    On Error GoTo Fail
    PptTqq_RelinkData pres
    mBusy = True
    Set xl = PptTqq_GetEditExcel(core)
    If xl Is Nothing Then mBusy = False: PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation: Exit Sub
    If Len(core) = 0 Then mBusy = False: PptTqq_Msg PPT_MSG_DATA_NOCORE, vbExclamation: Exit Sub
    If PptTqq_CoreApi(xl, core) < PPT_API_DATA Then mBusy = False: PptTqq_Msg PPT_MSG_OLD_CORE_DATA: Exit Sub
    PptTqq_EditShape shapes(1), xl, core
    mBusy = False
    Exit Sub
Fail:
    mBusy = False
    PptTqq_Msg PptTqq_ErrorText(Err.Number, Err.Description), vbExclamation
End Sub

' Apresentacao movida/renomeada (spec 14.2): formas TQQ_DATA cuja pasta de dados sumiu sao reapontadas para a pasta de dados NA PASTA DA
' APRESENTACAO (nome atual da apresentacao, ou o nome antigo do arquivo), se existir; o indice das automaticas acompanha. Sem Excel, sem
' perguntar. Devolve quantas formas. Barato quando TQQ_DATABOOK existe. Nunca propaga.
Public Function PptTqq_RelinkData(pres As Object) As Long
    Dim fso As Object, mark As String, sl As Object, shp As Object, oldP As String, cand As String, dirP As String, idx As String
    On Error Resume Next
    mark = PptTqq_Tag(pres, PPT_TAG_DATABOOK)
    If Len(mark) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If CBool(fso.FileExists(mark)) Then Exit Function
    dirP = CStr(pres.Path)
    If Len(dirP) = 0 Then Exit Function
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If PptTqq_Tag(shp, PPT_TAG_DATA) = "1" Then
                oldP = PptTqq_SourcePath(PptTqq_LinkSource(shp), PptTqq_Tag(shp, PPT_TAG_BOOK))
                If Not CBool(fso.FileExists(oldP)) Then
                    cand = PptTqq_DataPathFor(CStr(pres.FullName))
                    If Not CBool(fso.FileExists(cand)) Then cand = dirP & "\" & fso.GetFileName(oldP)
                    If CBool(fso.FileExists(cand)) Then
                        Err.Clear
                        shp.LinkFormat.SourceFullName = cand              ' sonda PG4: se o PowerPoint recusar, ver G4 Passo 5b
                        If Err.Number = 0 Then
                            shp.Tags.Add PPT_TAG_BOOK, cand
                            pres.Tags.Add PPT_TAG_DATABOOK, cand
                            If PptTqq_Tag(shp, PPT_TAG_AUTO) = "1" Then
                                idx = PptTqq_IdxSet(PptTqq_Tag(pres, PPT_TAG_INDEX), CLng(sl.SlideID), CStr(shp.Name), cand, True)
                                pres.Tags.Add PPT_TAG_INDEX, idx
                            End If
                            PptTqq_RelinkData = PptTqq_RelinkData + 1
                        Else
                            PptTqq_Log "PptTqq_RelinkData: " & Err.Number & " " & Err.Description
                            Err.Clear
                        End If
                    End If
                End If
            End If
        Next
    Next
End Function

' Evento PresentationSave (sonda PG5): salva a pasta de dados aberta e alterada no Excel com o ThinQQ (o de edicao, senao o aberto); nunca
' abre nem cria um Excel. Nunca propaga.
Public Sub PptTqq_OnSave(pres As Object)
    Dim xl As Object
    On Error Resume Next
    If Len(PptTqq_Tag(pres, PPT_TAG_DATABOOK)) = 0 Then Exit Sub
    If BUILD_CONFIG = "Test" Then
        Set xl = mTestXl
    ElseIf Not mEditXl Is Nothing Then
        Set xl = mEditXl
    Else
        Set xl = GetObject(, "Excel.Application")
    End If
    If xl Is Nothing Then Exit Sub
    If PptTqq_PingState(xl) <> 1 Then Exit Sub
    If PptTqq_CoreApi(xl, PptTqq_CoreFileName()) < PPT_API_DATA Then Exit Sub
    xl.Run "'" & PptTqq_CoreFileName() & "'!Core_PptDataSave", pres
    Err.Clear
End Sub
```
  Em `PptTqq_HandleOpen`, logo depois de `If Not seen Then`, acrescente `PptTqq_RelinkData pres                    ' apresentacao movida: graficos criados pelo PowerPoint voltam a achar os dados`. Em `clsPptEvents.cls`:

```vb
' Salvar a apresentacao tambem salva a pasta de dados dos graficos criados pelo PowerPoint (1.8.0).
Private Sub App_PresentationSave(ByVal Pres As Presentation)
    On Error Resume Next
    PptTqq_OnSave Pres
End Sub
```
  `modPptRibbon.bas`:

```vb
Public Sub PptRx_OnInsertChart(control As IRibbonControl)
    Dim pres As Object
    On Error Resume Next
    If Application.Presentations.Count = 0 Then PptTqq_Msg PPT_MSG_NO_PRES: Exit Sub
    Set pres = Application.ActivePresentation
    PptTqq_InsertChartIn pres, PptTqq_ActiveSlide(pres), control.Tag
End Sub

Public Sub PptRx_OnEditData(control As IRibbonControl)
    On Error Resume Next
    If Application.Presentations.Count = 0 Then PptTqq_Msg PPT_MSG_NO_PRES: Exit Sub
    PptTqq_EditShapes Application.ActivePresentation, PptTqq_SelectedTagged()
End Sub
```
  `customUI14.xml` do `.ppam`: como PRIMEIRO grupo da aba `tabThinQQPpt`:

```xml
        <group id="grpPptInsert" label="Inserir gráfico">
          <menu id="pmnuColumns" label="Colunas" size="large" image="tqq_columns" screentip="Colunas" supertip="Colunas empilhadas, agrupadas ou 100%. O gráfico nasce no slide com dados de exemplo; digite os seus na planilha do Excel que abre.">
            <button id="pbtnStacked" label="Colunas empilhadas" tag="stacked" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnClustered" label="Colunas agrupadas" tag="clustered" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnStacked100" label="Colunas 100%" tag="stacked100" onAction="PptRx_OnInsertChart"/>
          </menu>
          <menu id="pmnuBars" label="Barras" size="large" image="tqq_bars" screentip="Barras" supertip="Como as colunas, com as categorias no eixo vertical.">
            <button id="pbtnBarStacked" label="Barras empilhadas" tag="barstacked" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnBarClustered" label="Barras agrupadas" tag="barclustered" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnBarStacked100" label="Barras 100%" tag="barstacked100" onAction="PptRx_OnInsertChart"/>
          </menu>
          <menu id="pmnuWaterfall" label="Cascata" size="large" image="tqq_waterfall" screentip="Cascata (ponte)" supertip="Etapa e valor; 'e' no valor para uma barra de total/subtotal calculada. A Cascata comparativa põe duas séries lado a lado.">
            <button id="pbtnWaterfall" label="Cascata (colunas)" tag="waterfall" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnBarWaterfall" label="Cascata (barras)" tag="barwaterfall" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnWaterfallPair" label="Cascata comparativa" tag="waterfallpair" onAction="PptRx_OnInsertChart" image="tqq_waterfallpair16"/>
          </menu>
          <menu id="pmnuLines" label="Linha e área" size="large" image="tqq_lines" screentip="Linha e área" supertip="Linha, áreas empilhadas ou 100% e a combinação de colunas com linha.">
            <button id="pbtnLine" label="Linha" tag="line" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnArea" label="Área empilhada" tag="area" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnArea100" label="Área 100%" tag="area100" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnCombo" label="Combinação (colunas + linha)" tag="combo" onAction="PptRx_OnInsertChart"/>
          </menu>
          <menu id="pmnuOthers" label="Mais tipos" size="large" image="tqq_more" screentip="Mais tipos" supertip="Mekko, Gantt, pizza, rosca, borboleta, dispersão e bolhas.">
            <button id="pbtnMekko" label="Mekko (%)" tag="mekko" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnMekkoUnits" label="Mekko (unidades)" tag="mekkounits" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnGantt" label="Gantt" tag="gantt" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnPie" label="Pizza" tag="pie" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnDoughnut" label="Rosca" tag="doughnut" onAction="PptRx_OnInsertChart"/>
            <button id="pbtnButterfly" label="Borboleta" tag="butterfly" onAction="PptRx_OnInsertChart" image="tqq_butterfly16"/>
            <button id="pbtnScatter" label="Dispersão" tag="scatter" onAction="PptRx_OnInsertChart" image="tqq_scatter16"/>
            <button id="pbtnBubble" label="Bolhas" tag="bubble" onAction="PptRx_OnInsertChart" image="tqq_bubble16"/>
          </menu>
        </group>
```
  e, no `grpPptElem`, depois de `btnPptUpdateSel`:

```xml
          <button id="btnPptEditData" label="Editar dados" imageMso="TableInsert" onAction="PptRx_OnEditData" getEnabled="PptRx_GetElemEnabled"
                  screentip="Editar dados" supertip="Abre no Excel a planilha com os dados do gráfico selecionado (para gráficos criados aqui, a pasta de dados ao lado da apresentação). O slide acompanha o que você digita."/>
```
  (Ordem dos 21 `tag` = a do Excel, conferida pelo `Test-Ribbon.ps1`. Filhos de menu sem `size`. Nenhum controle com `image` e `imageMso` juntos.)
- [ ] **Passo 4 (GREEN):** `-Target core -Filter "TestSmoke" -Ppt -TimeoutSec 900` → projeto do PowerPoint 0 falhas e e2e ok; `Test-Ribbon.ps1` 0 falhas.
- [ ] **Passo 5b (só se PG4 negativa — `SourceFullName` não gravável):** `PptTqq_RelinkData` passa a só marcar as formas (`TQQ_RELINK = cand`) e o `Core_PptEditData`/sincronização recolam o gráfico do caminho novo (`Ppt_PasteChart` + `PptSyncAfterPaste` + apagar a forma velha, mantendo posição e tamanho); documentar que a animação vai para o fim (medido em 2026-10-14).
- [ ] **Passo 6:** suíte completa com `-Ppt`. Commit `feat: aba ThinQQ do PowerPoint com Inserir grafico (21 tipos) e Editar dados` (+ trailer).

### Tarefa G5: Documentação da 1.8.0 — modelo: standard
Tutorial `#ppt-criar` (inserir pelo PowerPoint; os 5 menus; apresentação salva antes; a pasta de dados ao lado: nome, uma planilha por gráfico, não renomear as planilhas; digitar e ver o slide mudar; "Editar dados"; salvar salva os dados; mover a pasta com os dois arquivos; o que acontece sem a pasta de dados; duplo clique abre o "Editar Dados" do PowerPoint, não o do ThinQQ; Excel que abre sozinho fica com o usuário; limitações frente ao think-cell — spec §14.8); README (Recursos; Limitações; manuais 86–96 = spec §14.9); CHANGELOG/notes ("Criar gráficos do ThinQQ direto pelo PowerPoint, com os 21 tipos e o Excel como planilha de dados."); Test-Docs (rótulos `Inserir gráfico`, `Editar dados`, sufixo ` - dados ThinQQ.xlsx`, âncora). Commit `docs: criar graficos pelo PowerPoint` (+ trailer).

### Tarefa G6: Suíte final, revisão do ramo, verificação manual — modelo: standard (controlador + humano)
Suíte completa com `-Ppt`, `POWERPNT` 0/0; revisão final (Review Focus 5–6; segurança da abertura para edição; nenhum processo filho); uma onda; humano: itens 86–96 (inclui PG2, PG6, PG8, PG9), PR, `publicar-release 1.8.0`.

---

### Tarefa G7: Comandos do gráfico pelo PowerPoint no core (API 4) — modelo: most-capable

**Files:** `src/core/modules/modCoreApi.bas` (`Core_PptCommand`, `Core_PptCategories`, `PptCmd_One`, `PptCmd_Book`), `src/core/modules/modPptSync.bas` (`PPTSYNC_TAG_KIND/OPTS`, `PptSync_OptsText`, `PptSync_TagOpts`, chamada em `PptSync_SyncOne`; `CORE_PPT_API = 4`); `PptSyncAfterPaste` grava as opções; testes `TestPptData.bas`, `TestPptSync.bas`.

**Interfaces:** `PptSync_OptsText(c As Chart) As String` (`labels=1;totals=0;connectors=1;numfmt=int;legend=right;caps=labels,totals,...,`); `PptSync_TagOpts(shp As Object, c As Chart)`; `Core_PptCommand(pres As Object, entries As String, cmd As String, arg As String) As String`; `Core_PptCategories(pres As Object, entries As String) As String` (nomes por `Chr(31)`); `PPTCMD_MSG_OPEN_SOURCE`.

- [ ] **Passo 1: Testes (RED)** — `TestPptData.bas`:

```vb
Public Sub Test_PptSync_OptsText()
    Dim wb As Workbook, ws As Worksheet, c As Chart, s As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Samples_WriteTable(ws.Range("A1"), Samples_TableFor("stacked")).Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    s = PptSync_OptsText(c)
    AssertTrue InStr(s, "numfmt=") > 0 And InStr(s, "legend=") > 0, s
    AssertTrue InStr(s, "caps=") > 0 And InStr(s, "totals,") > 0 And InStr(s, "cagr,") > 0, "empilhadas: totais e CAGR (" & s & ")"
    Core_SetToggle "totals", True
    AssertTrue InStr(PptSync_OptsText(c), "totals=1") > 0
    Watchers_Stop
    CloseTemp wb
End Sub
```
  `TestPptSync.bas` (reaproveita o arranjo de `Test_PptInsertChart_FromPpt_WhenAvailable`):

```vb
Public Sub Test_PptCommand_FromPpt_WhenAvailable()
    Dim app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String, msg As String
    Dim shp As Object, p As String, pp As String, wb As Workbook, e As String, n0 As Long
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoReset
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    If UserWindowAppeared(app) Then GoTo Cleanup
    pp = TempDir() & "\Cmd G.pptx": p = TempDir() & "\Cmd G - dados ThinQQ.xlsx"
    If Len(Dir(p)) > 0 Then Kill p
    pres.SaveAs pp
    AssertEqual "", Core_PptInsertChart(slide, "clustered")
    Set shp = slide.Shapes(1)
    Set wb = Workbooks(Dir(p))
    e = PptIdx_Entry(CLng(slide.SlideID), CStr(shp.Name), "")
    Ppt_SetAppForTest app
    AssertEqual "clustered", PptSync_Tag(shp, PPTSYNC_TAG_KIND)
    AssertEqual "", Core_PptCommand(pres, e, "numfmt", "pct0")
    AssertTrue InStr(PptSync_Tag(shp, PPTSYNC_TAG_OPTS), "numfmt=pct0") > 0, PptSync_Tag(shp, PPTSYNC_TAG_OPTS)
    AssertTrue wb.Saved, "pasta de dados salva depois do comando"
    AssertEqual "2022" & Chr$(31) & "2023" & Chr$(31) & "2024" & Chr$(31) & "2025", Core_PptCategories(pres, e)
    n0 = shp.Chart.Shapes.Count
    AssertEqual "", Core_PptCommand(pres, e, "cagr", "1,4")
    AssertTrue shp.Chart.Shapes.Count > n0, "seta CAGR clonada no slide"
    AssertEqual "", Core_PptCommand(pres, e, "clear", "")
    AssertEqual 0, CountAnn(shp.Chart), "setas removidas no slide"
    AssertTrue Len(Core_PptCommand(pres, e, "toggle", "totals")) > 0, "agrupadas nao tem totais: recusa da matriz em texto"
    AssertEqual "Comando desconhecido: xyz", Core_PptCommand(pres, e, "xyz", "")
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set shp = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    Watchers_Stop
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Kill p: Kill pp
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

Private Function CountAnn(pc As Object) As Long
    Dim i As Long
    For i = 1 To pc.Shapes.Count
        If Left$(CStr(pc.Shapes(i).Name), Len(ANN_PREFIX)) = ANN_PREFIX Then CountAnn = CountAnn + 1
    Next
End Function
```
- [ ] **Passo 2: Código** — `modPptSync.bas`: constantes `Public Const PPTSYNC_TAG_KIND As String = "TQQ_KIND"` e `Public Const PPTSYNC_TAG_OPTS As String = "TQQ_OPTS"`; `CORE_PPT_API = 4` (comentário "4 = Core_PptCommand, Core_PptCategories (1.9.0)"); funções:

```vb
' Estado das opcoes do grafico para a faixa do PowerPoint ler SEM chamar o Excel (spec 14.5). Pura sobre o grafico.
Public Function PptSync_OptsText(c As Chart) As String
    Dim k As Variant, s As String, caps As String, kind As String
    kind = Registry_ChartType(c)
    For Each k In Array("labels", "totals", "connectors")
        s = s & k & "=" & IIf(Registry_GetOption(c, CStr(k), "0") = "1", "1", "0") & ";"
    Next
    s = s & "numfmt=" & Registry_GetOption(c, "numfmt", "int") & ";legend=" & Registry_GetOption(c, "legend", "right") & ";"
    For Each k In Array("labels", "totals", "connectors", "numfmt", "legend", "cagr", "diff")
        If Kinds_Supports(kind, CStr(k)) Then caps = caps & k & ","
    Next
    PptSync_OptsText = s & "caps=" & caps
End Function

Public Sub PptSync_TagOpts(shp As Object, c As Chart)
    On Error Resume Next
    shp.Tags.Add PPTSYNC_TAG_KIND, Registry_ChartType(c)
    shp.Tags.Add PPTSYNC_TAG_OPTS, PptSync_OptsText(c)
    Err.Clear
End Sub
```
  Em `PptSync_SyncOne`, logo depois de `PptSync_SyncOne = PPTSYNC_R_OK`: `PptSync_TagOpts shp, co.Chart`. Em `PptSyncAfterPaste` (`modCoreApi`), depois de `PptSync_TagShape`: `PptSync_TagOpts shp, co.Chart`. `modCoreApi.bas`:

```vb
Public Const PPTCMD_MSG_OPEN_SOURCE As String = "Abra a pasta de origem no Excel para mudar este gráfico (o ThinQQ não altera pastas fechadas do usuário)."

' Pasta para um comando do PowerPoint: a ja aberta neste Excel; senao so a pasta de DADOS (forma TQQ_DATA + marca) e aberta para edicao.
' Nothing = pasta do usuario fechada.
Private Function PptCmd_Book(shp As Object, path As String) As Workbook
    Dim w As Workbook
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then Set PptCmd_Book = w: Exit Function
    Next
    If PptSync_Tag(shp, PPTDATA_TAG) <> "1" Then Exit Function
    If Not PptSync_FileExists(path) Then Exit Function
    Set w = PptData_OpenForEdit(path)
    If Not PptData_IsDataBook(w) Then w.Close SaveChanges:=False: Exit Function
    Set PptCmd_Book = w
End Function

' Um comando numa forma: o mesmo codigo da faixa do Excel, Core_SafeRefresh e, se for pasta de dados, salvar. Pastas do usuario ficam alteradas
' e nao salvas (o usuario decide). "" ou a mensagem.
Private Function PptCmd_One(shp As Object, cmd As String, arg As String) As String
    Dim wb As Workbook, co As ChartObject, c As Chart, p() As String, msg As String, kind As String
    On Error GoTo Fail
    Set wb = PptCmd_Book(shp, PptSync_ItemBook(shp))
    If wb Is Nothing Then PptCmd_One = PPTCMD_MSG_OPEN_SOURCE: Exit Function
    Set co = Registry_FindChartObject(wb, PptSync_Tag(shp, PPTSYNC_TAG_CHART))
    If co Is Nothing Then PptCmd_One = "Gráfico não encontrado na pasta de origem (apagado ou recriado).": Exit Function
    Set c = co.Chart
    kind = Registry_ChartType(c)
    Select Case cmd
        Case "toggle"
            If Not Kinds_Supports(kind, arg) Then PptCmd_One = Kinds_NotApplicableText(arg): Exit Function
            ApplyToggle c, arg, Not (Registry_GetOption(c, arg, "0") = "1")
        Case "numfmt", "legend"
            msg = ValidateOption(cmd, arg)
            If Len(msg) > 0 Then PptCmd_One = msg: Exit Function
            ApplyOption c, cmd, arg
        Case "cagr", "diff"
            If Not Kinds_Supports(kind, cmd) Then PptCmd_One = Kinds_NotApplicableText(cmd): Exit Function
            p = Split(arg, ",")
            If UBound(p) <> 1 Then PptCmd_One = "Escolha duas categorias.": Exit Function
            If cmd = "cagr" Then Ann_AddCagr c, CLng(Val(p(0))), CLng(Val(p(1))) Else Ann_AddDifference c, CLng(Val(p(0))), CLng(Val(p(1)))
        Case "clear"
            Ann_RemoveAll c, False
        Case Else
            PptCmd_One = "Comando desconhecido: " & cmd: Exit Function
    End Select
    Core_SafeRefresh c
    If PptData_IsDataBook(wb) Then wb.Save
    Exit Function
Fail:
    PptCmd_One = Err.Description
    Errors_Log "PptCmd_One(" & cmd & ")" & vbTab & Err.Number & vbTab & Err.Description
End Function

' API 4 (1.9.0): comando da faixa do PowerPoint nos graficos listados (spec 14.5): "toggle" (arg = labels|totals|connectors: inverte),
' "numfmt"/"legend" (arg = valor), "cagr"/"diff" (arg = "de,ate", indices de categoria a partir de 1), "clear" (remove as setas). Depois
' sincroniza (forcado) so essas formas. Devolve "" ou a 1a mensagem. Nunca propaga erro.
Public Function Core_PptCommand(pres As Object, entries As String, cmd As String, arg As String) As String
    Dim items As Collection, shp As Object, msg As String, first As String, en As Long, ed As String
    On Error GoTo Fail
    Set items = PptSync_ResolveEntries(pres, entries)
    If items.Count = 0 Then Core_PptCommand = PPTSYNC_NONE: Exit Function
    For Each shp In items
        msg = PptCmd_One(shp, cmd, arg)
        If Len(msg) > 0 And Len(first) = 0 Then first = msg
    Next
    If Len(first) = 0 Then PptSync_Items items, True
    Core_PptCommand = first
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Errors_Log "Core_PptCommand" & vbTab & en & vbTab & ed
    If Ppt_IsBusyError(en) Then Core_PptCommand = PPT_ERR_BUSY Else Core_PptCommand = "Não foi possível aplicar: " & ed & " (" & en & ")"
End Function

' API 4: categorias do (1o) grafico listado, separadas por Chr(31), para o .ppam montar a escolha de/ate. "" se a pasta nao estiver
' disponivel. Nunca propaga erro.
Public Function Core_PptCategories(pres As Object, entries As String) As String
    Dim items As Collection, shp As Object, wb As Workbook, co As ChartObject, v As Variant, i As Long, s As String
    On Error GoTo Done
    Set items = PptSync_ResolveEntries(pres, entries)
    If items.Count = 0 Then Exit Function
    Set shp = items(1)
    Set wb = PptCmd_Book(shp, PptSync_ItemBook(shp))
    If wb Is Nothing Then Exit Function
    Set co = Registry_FindChartObject(wb, PptSync_Tag(shp, PPTSYNC_TAG_CHART))
    If co Is Nothing Then Exit Function
    v = co.Chart.SeriesCollection(1).XValues
    If Not IsArray(v) Then Exit Function
    For i = LBound(v) To UBound(v)
        If i > LBound(v) Then s = s & Chr$(31)
        s = s & CStr(v(i))
    Next
    Core_PptCategories = s
Done:
End Function
```
  (Com uma recusa, nada é sincronizado: o slide fica como estava. `Kinds_NotApplicableText` é o texto das recusas 7002 da faixa do Excel.)
- [ ] **Passo 3 (GREEN):** `-Filter "TestPptData|TestPptSync|TestCoreApi"`; suíte com `-Ppt`. Commit `feat: comandos do grafico do ThinQQ pelo PowerPoint (API 4)` (+ trailer).

### Tarefa G8: Grupo "Gráfico do ThinQQ" no PowerPoint — modelo: standard

**Files:** `src/ppt/modules/modPptTqq.bas`, `modPptRibbon.bas`, `ribbon/customUI14.xml`; testes `TestPptAddin.bas`, `Test-Ribbon.ps1`.

**Interfaces:** `PPT_API_CMD = 4`, `PPT_MSG_OLD_CORE_CMD`, `PPT_TAG_OPTS`; `PptTqq_OptValue(opts, key) As String`, `PptTqq_OptCap(opts, feature) As Boolean`, `PptTqq_SelectedOpts() As String`, `PptTqq_CommandShapes(pres, shapes, cmd, arg)`, `PptTqq_PickTwo(cats As String, ByRef a As Long, ByRef b As Long) As Boolean` (gancho `PptTqq_SetPickForTest`); callbacks `PptRx_GetOpt`, `PptRx_OnToggleOpt`, `PptRx_GetOptEnabled`, `PptRx_GetOptItem`, `PptRx_OnOptItem`, `PptRx_OnArrow`, `PptRx_OnClearAnn`.

- [ ] **Passo 1: Testes (RED)** — `TestPptAddin.bas`:

```vb
Public Sub Test_PptTqq_OptValueAndCap()
    Dim o As String
    o = "labels=1;totals=0;connectors=1;numfmt=pct0;legend=right;caps=labels,totals,cagr,"
    AssertEqual "1", PptTqq_OptValue(o, "labels")
    AssertEqual "pct0", PptTqq_OptValue(o, "numfmt")
    AssertEqual "", PptTqq_OptValue(o, "nada")
    AssertTrue PptTqq_OptCap(o, "cagr")
    AssertFalse PptTqq_OptCap(o, "diff")
    AssertFalse PptTqq_OptCap(o, "tal"), "nao confunde prefixo (totals)"
End Sub

Public Sub Test_PptTqq_PickTwo_TestHook()
    Dim a As Long, b As Long
    PptTqq_SetPickForTest 1, 3
    AssertTrue PptTqq_PickTwo("A" & Chr$(31) & "B" & Chr$(31) & "C", a, b)
    AssertEqual 1, a: AssertEqual 3, b
    PptTqq_SetPickForTest 3, 3
    AssertFalse PptTqq_PickTwo("A" & Chr$(31) & "B" & Chr$(31) & "C", a, b), "de = ate: recusado"
    PptTqq_ResetTestHooks
End Sub
```
- [ ] **Passo 2: Código** — `modPptTqq.bas`:

```vb
Public Const PPT_API_CMD As Long = 4
Public Const PPT_TAG_OPTS As String = "TQQ_OPTS"
Public Const PPT_MSG_OLD_CORE_CMD As String = "Atualize o ThinQQ no Excel para mudar o gráfico pelo PowerPoint."
Private mTestPickA As Long                                  ' gancho: respostas de PptTqq_PickTwo
Private mTestPickB As Long

Public Sub PptTqq_SetPickForTest(a As Long, b As Long)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestPickA = a: mTestPickB = b
End Sub

' Valor de uma chave em TQQ_OPTS ("a=1;b=x;caps=..."). Pura.
Public Function PptTqq_OptValue(opts As String, key As String) As String
    Dim p As Variant, k As Long
    For Each p In Split(opts, ";")
        k = InStr(CStr(p), "=")
        If k > 1 Then
            If Left$(CStr(p), k - 1) = key Then PptTqq_OptValue = Mid$(CStr(p), k + 1): Exit Function
        End If
    Next
End Function

Public Function PptTqq_OptCap(opts As String, feature As String) As Boolean
    PptTqq_OptCap = InStr("," & PptTqq_OptValue(opts, "caps"), "," & feature & ",") > 0
End Function

' TQQ_OPTS da 1a forma do ThinQQ selecionada ("" sem selecao).
Public Function PptTqq_SelectedOpts() As String
    Dim sel As Collection
    Set sel = PptTqq_SelectedTagged()
    If sel.Count > 0 Then PptTqq_SelectedOpts = PptTqq_Tag(sel(1), PPT_TAG_OPTS)
End Function

' Escolha de/ate pela lista numerada de categorias (InputBox do VBA). False se cancelado ou invalido.
Public Function PptTqq_PickTwo(cats As String, ByRef a As Long, ByRef b As Long) As Boolean
    Dim c() As String, i As Long, lst As String, n As Long, v As String
    c = Split(cats, Chr$(31)): n = UBound(c) + 1
    If n < 2 Then Exit Function
    If BUILD_CONFIG = "Test" Then
        a = mTestPickA: b = mTestPickB
    Else
        For i = 0 To n - 1
            lst = lst & (i + 1) & " - " & c(i) & vbLf
            If Len(lst) > 700 Then lst = lst & "...": Exit For
        Next
        v = InputBox(lst & vbLf & "Da categoria número:", PPT_APP_NAME, "1")
        If Len(v) = 0 Or Not IsNumeric(v) Then Exit Function
        a = CLng(v)
        v = InputBox(lst & vbLf & "Até a categoria número:", PPT_APP_NAME, CStr(n))
        If Len(v) = 0 Or Not IsNumeric(v) Then Exit Function
        b = CLng(v)
    End If
    PptTqq_PickTwo = (a >= 1 And b <= n And a < b)
End Function

' Comando no core para as formas dadas (Excel com o core; uma instancia criada aqui e fechada no fim, como a sincronizacao). Nunca propaga.
Public Sub PptTqq_CommandShapes(pres As Object, shapes As Collection, cmd As String, arg As String)
    Dim xl As Object, created As Boolean, core As String, msg As String, cats As String, a As Long, b As Long, a2 As String
    If mBusy Then Exit Sub
    If shapes.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    mBusy = True
    On Error GoTo Fail
    Set xl = PptTqq_GetExcel(created)
    If xl Is Nothing Then mBusy = False: PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation: Exit Sub
    core = PptTqq_EnsureCore(xl, created)
    a2 = arg
    If Len(core) = 0 Then
        msg = PPT_MSG_DATA_NOCORE
    ElseIf PptTqq_CoreApi(xl, core) < PPT_API_CMD Then
        msg = PPT_MSG_OLD_CORE_CMD
    Else
        If cmd = "cagr" Or cmd = "diff" Then
            cats = CStr(xl.Run("'" & core & "'!Core_PptCategories", pres, PptTqq_EntriesOf(shapes)))
            If Len(cats) = 0 Then
                msg = "Abra a pasta de origem no Excel para mudar este gráfico (o ThinQQ não altera pastas fechadas do usuário)."
            ElseIf Not PptTqq_PickTwo(cats, a, b) Then
                msg = ""
                cmd = ""
            Else
                a2 = a & "," & b
            End If
        End If
        If Len(cmd) > 0 And Len(msg) = 0 Then msg = CStr(xl.Run("'" & core & "'!Core_PptCommand", pres, PptTqq_EntriesOf(shapes), cmd, a2))
    End If
    PptTqq_Release xl, created
    mBusy = False
    If Len(msg) > 0 Then PptTqq_Msg msg, vbExclamation
    PptRx_InvalidateElem
    Exit Sub
Fail:
    PptTqq_Release xl, created
    mBusy = False
    PptTqq_Msg PptTqq_ErrorText(Err.Number, Err.Description), vbExclamation
End Sub
```
  `PptTqq_ResetTestHooks` passa a zerar `mTestPickA/B`. `modPptRibbon.bas`:

```vb
Public Sub PptRx_GetOpt(control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    pressed = (PptTqq_OptValue(PptTqq_SelectedOpts(), control.Tag) = "1")
End Sub

Public Sub PptRx_GetOptEnabled(control As IRibbonControl, ByRef enabled)
    On Error Resume Next
    enabled = PptTqq_OptCap(PptTqq_SelectedOpts(), control.Tag)
End Sub

Public Sub PptRx_OnToggleOpt(control As IRibbonControl, pressed As Boolean)
    On Error Resume Next
    PptTqq_CommandShapes Application.ActivePresentation, PptTqq_SelectedTagged(), "toggle", control.Tag
End Sub

Public Sub PptRx_GetOptItem(control As IRibbonControl, ByRef itemId)
    On Error Resume Next
    itemId = control.Tag & PptTqq_OptValue(PptTqq_SelectedOpts(), control.Tag)
End Sub

' Itens dos dropdowns: id = tag + valor ("numfmtpct0", "legendright").
Public Sub PptRx_OnOptItem(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    On Error Resume Next
    PptTqq_CommandShapes Application.ActivePresentation, PptTqq_SelectedTagged(), control.Tag, Mid$(selectedId, Len(control.Tag) + 1)
End Sub

Public Sub PptRx_OnArrow(control As IRibbonControl)
    On Error Resume Next
    PptTqq_CommandShapes Application.ActivePresentation, PptTqq_SelectedTagged(), control.Tag, ""
End Sub

Public Sub PptRx_OnClearAnn(control As IRibbonControl)
    On Error Resume Next
    PptTqq_CommandShapes Application.ActivePresentation, PptTqq_SelectedTagged(), "clear", ""
End Sub
```
  `PptRx_InvalidateElem` passa a invalidar a faixa inteira (`mRibbon.Invalidate`). XML (depois do `grpPptElem`):

```xml
        <group id="grpPptChart" label="Gráfico do ThinQQ">
          <toggleButton id="tglPptLabels" label="Rótulos" imageMso="TextBoxInsert" tag="labels" getPressed="PptRx_GetOpt" onAction="PptRx_OnToggleOpt" getEnabled="PptRx_GetOptEnabled"/>
          <toggleButton id="tglPptTotals" label="Totais" imageMso="AutoSum" tag="totals" getPressed="PptRx_GetOpt" onAction="PptRx_OnToggleOpt" getEnabled="PptRx_GetOptEnabled"/>
          <toggleButton id="tglPptConnectors" label="Conectores" imageMso="ShapeStraightConnector" tag="connectors" getPressed="PptRx_GetOpt" onAction="PptRx_OnToggleOpt" getEnabled="PptRx_GetOptEnabled"/>
          <dropDown id="ddPptNumFmt" label="Formato" tag="numfmt" getSelectedItemID="PptRx_GetOptItem" onAction="PptRx_OnOptItem" getEnabled="PptRx_GetOptEnabled">
            <item id="numfmtint" label="1.234"/>
            <item id="numfmtdec1" label="1.234,5"/>
            <item id="numfmtdec2" label="1.234,56"/>
            <item id="numfmtpct0" label="12%"/>
            <item id="numfmtpct1" label="12,3%"/>
            <item id="numfmtmil" label="Milhares (1.234 → 1)"/>
            <item id="numfmtmi" label="Milhões (1,2)"/>
          </dropDown>
          <dropDown id="ddPptLegend" label="Legenda" tag="legend" getSelectedItemID="PptRx_GetOptItem" onAction="PptRx_OnOptItem" getEnabled="PptRx_GetOptEnabled">
            <item id="legendright" label="Direita"/>
            <item id="legendnone" label="Nenhuma"/>
          </dropDown>
          <button id="btnPptCagr" label="Seta CAGR" imageMso="ArrowsMore" tag="cagr" onAction="PptRx_OnArrow" getEnabled="PptRx_GetOptEnabled"/>
          <button id="btnPptDiff" label="Seta diferença" imageMso="ShapeUpArrow" tag="diff" onAction="PptRx_OnArrow" getEnabled="PptRx_GetOptEnabled"/>
          <button id="btnPptClearAnn" label="Remover setas" imageMso="Delete" onAction="PptRx_OnClearAnn" getEnabled="PptRx_GetElemEnabled"/>
        </group>
```
  (O item "Milhares (1.234 → 1)" já existe no XML do Excel: a seta é caractere do XML, não do VBA — o lint de cp1252 vale só para `.bas/.cls`. `Test-Ribbon.ps1`: acrescente a conferência do grupo, dos 3 `toggleButton`, dos 2 `dropDown` e dos 3 botões com os callbacks acima.)
- [ ] **Passo 3 (GREEN)** e suíte com `-Ppt`. Commit `feat: grupo Grafico do ThinQQ no PowerPoint (rotulos, totais, formato, setas)` (+ trailer).

### Tarefa G9: "Empacotar para enviar" — modelo: standard

**Files:** `src/ppt/modules/modPptTqq.bas` (`PptTqq_Package`), `modPptRibbon.bas` (`PptRx_OnPackage`), `ribbon/customUI14.xml` (`btnPptPackage` no `grpPptSync`); teste `TestPptAddin.bas`.

- [ ] **Passo 1: Teste (RED):**

```vb
Public Sub Test_PptTqq_Package_CopiesBoth()
    Dim p As Object, f As String, d As String, dest As String, k As Integer, en As Long, ed As String, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set p = NewPres()
    On Error GoTo Cleanup
    f = Environ$("TEMP") & "\tqq-pack.pptx": p.SaveAs f
    d = Environ$("TEMP") & "\tqq-pack - dados ThinQQ.xlsx"
    k = FreeFile: Open d For Output As #k: Print #k, "x": Close #k
    p.Tags.Add PPT_TAG_DATABOOK, d
    dest = PptTqq_Package(p)
    AssertEqual Environ$("TEMP") & "\tqq-pack (ThinQQ)", dest
    AssertTrue fso.FileExists(dest & "\tqq-pack.pptx"), "apresentacao copiada"
    AssertTrue fso.FileExists(dest & "\tqq-pack - dados ThinQQ.xlsx"), "pasta de dados copiada"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    ClosePres p
    On Error Resume Next
    fso.DeleteFolder Environ$("TEMP") & "\tqq-pack (ThinQQ)", True
    Kill f: Kill d
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub
```
- [ ] **Passo 2: Código:**

```vb
' "Empacotar para enviar" (1.9.0; spec 14.2): salva (apresentacao e pasta de dados aberta) e copia os dois para "<pasta>\<nome> (ThinQQ)\".
' Sem Excel obrigatorio, sem ZIP, sem processo filho. Devolve a pasta ("" se nao foi possivel; o aviso diz por que). Nunca propaga.
Public Function PptTqq_Package(pres As Object) As String
    Dim fso As Object, dest As String, data As String
    On Error GoTo Fail
    If Len(CStr(pres.Path)) = 0 Then PptTqq_Msg PPT_MSG_SAVE_FIRST: Exit Function
    PptTqq_OnSave pres
    If CLng(pres.Saved) = 0 Then pres.Save
    Set fso = CreateObject("Scripting.FileSystemObject")
    dest = CStr(pres.Path) & "\" & fso.GetBaseName(CStr(pres.FullName)) & " (ThinQQ)"
    If Not CBool(fso.FolderExists(dest)) Then fso.CreateFolder dest
    fso.CopyFile CStr(pres.FullName), dest & "\", True
    data = PptTqq_Tag(pres, PPT_TAG_DATABOOK)
    If Len(data) > 0 Then
        If CBool(fso.FileExists(data)) Then fso.CopyFile data, dest & "\", True
    End If
    PptTqq_Package = dest
    PptTqq_Msg "Pronto: a apresentação e a pasta de dados foram copiadas para '" & dest & "'. Envie a pasta inteira (por exemplo, compactada): quem abrir a apresentação dessa pasta terá os gráficos ligados aos dados ao lado."
    Exit Function
Fail:
    PptTqq_Msg "Não foi possível empacotar: " & Err.Description & " (" & Err.Number & ")", vbExclamation
End Function
```
  Ribbon: `<button id="btnPptPackage" label="Empacotar para enviar" image="tqq_ppt16" onAction="PptRx_OnPackage" screentip="Empacotar para enviar" supertip="Copia a apresentação e a pasta de dados dos gráficos do ThinQQ para uma pasta nova ao lado, pronta para compactar e enviar."/>` no `grpPptSync`; `PptRx_OnPackage` chama `PptTqq_Package Application.ActivePresentation` (sem apresentação: `PPT_MSG_NO_PRES`). Sonda manual: copiar a pasta de dados aberta no Excel (G11).
- [ ] **Passo 3 (GREEN)**, suíte com `-Ppt`, commit `feat: empacotar apresentacao e dados para enviar` (+ trailer).

### Tarefa G10: Documentação da fase G-II — modelo: standard
Tutorial `#ppt-criar` ganha "Mudar o gráfico pelo PowerPoint" (o que cada botão faz; pastas do usuário precisam estar abertas; delta por cliques continua no Excel via "Editar dados") e "Empacotar para enviar"; README (manuais 97–101: comandos em 5 tipos, pasta do usuário fechada, Empacotar com a pasta de dados aberta, abrir o pacote em outro computador); CHANGELOG/notes; Test-Docs. Commit `docs: comandos do grafico pelo PowerPoint e empacotar` (+ trailer).

### Tarefa G11: Suíte final da fase G-II — modelo: standard (controlador + humano)
Suíte com `-Ppt`; revisão escopada de G7–G10 (segurança: `PptCmd_Book` nunca abre pasta do usuário); humano: itens 97–101. O PR/release é o mesmo da Parte E (1.9.0, depois de E5).

---

# PARTE E — Texto vinculado a célula (1.9.0)

### Tarefa E1: Modelo do texto vinculado e rota "text" no motor — modelo: most-capable

**Files:** Create `src/core/modules/modPptText.bas`; Modify `src/core/modules/modPptSync.bas` (constantes de resultado públicas, `PPTSYNC_R_EDITED`, `PptSync_SummaryText` com `nEdited`, `PptSync_ItemKind` "text", rota em `PptSync_Items`), `src/core/modules/modPptAuto.bas` (`PptAuto_IsCandidate` por nomes), `src/core/modules/modCoreApi.bas` (`Core_PptLinkText`, `CORE_PPT_API = 3`); Test `tests/vba/core/TestPptText.bas` (novo, puro) e `TestPptSync.bas` (PowerPoint real).

**Interfaces:** Produces `PPTTEXT_TAG = "TQQ_TEXT"`, `PPTLINK_TAG_NAME = "TQQ_TNAME"`, `PPTLINK_TAG_REF = "TQQ_TREF"`, `PPTTEXT_TAG_PRE/SUF/LAST`; `PptText_Locate(cur, pre, suf, last, ByRef startPos As Long, ByRef length As Long) As Boolean`; `PptText_CellText(c As Range) As String`; `PptLink_DefineName(rng As Range, prefix As String) As String`; `PptLink_RefOf(rng As Range) As String`; `PptLink_ResolveRange(wb As Workbook, nm As String, ref As String) As Range`; `PptText_Link(shp As Object, startPos As Long, length As Long, cell As Range) As String` ("" = ok, senão a mensagem); `PptText_Sync(shp As Object, wb As Workbook, force As Boolean) As Long`; `Core_PptLinkText(shp As Object, startPos As Long, length As Long) As String`; `PPTSYNC_R_OK/SAME/NOCHART/FAIL/SKIPPED/EDITED` públicas.

- [ ] **Passo 1: Testes puros (RED)** — crie `tests/vba/core/TestPptText.bas`:

```vb
Attribute VB_Name = "TestPptText"
Option Explicit
' 1.9.0: texto vinculado a celula (spec 2026-10-15 secao 5.3). Partes puras e o modelo de nome definido (sem PowerPoint).

Public Sub Test_PptText_Locate_Cases()
    Dim st As Long, ln As Long
    AssertTrue PptText_Locate("Receita de 123 mil", "Receita de ", " mil", "123", st, ln)
    AssertEqual 12, st: AssertEqual 3, ln
    AssertTrue PptText_Locate("Receita de 99999 mil", "Receita de ", " mil", "123", st, ln), "usuario editou so o campo"
    AssertEqual 12, st: AssertEqual 5, ln
    AssertFalse PptText_Locate("Lucro de 123 mil", "Receita de ", " mil", "123", st, ln), "editou o prefixo: nao tocar"
    AssertFalse PptText_Locate("x", "abc", "def", "", st, ln), "texto menor que prefixo+sufixo"
    AssertTrue PptText_Locate("123", "", "", "123", st, ln), "forma inteira"
    AssertEqual 1, st: AssertEqual 3, ln
    AssertTrue PptText_Locate("ab", "ab", "", "", st, ln), "campo vazio no fim"
    AssertEqual 3, st: AssertEqual 0, ln
End Sub

Public Sub Test_PptLink_NameFollowsRowsAndRef()
    Dim wb As Workbook, ws As Worksheet, nm As String, r As Range
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("E1").Value = 1234.5: ws.Range("E1").NumberFormat = "#,##0.0"
    nm = PptLink_DefineName(ws.Range("E1"), "TQQ_X_")
    AssertTrue Left$(nm, 6) = "TQQ_X_"
    AssertFalse wb.Names(nm).Visible, "nome oculto"
    ws.Rows(1).Insert
    Set r = PptLink_ResolveRange(wb, nm, "")
    AssertEqual "$E$2", r.Address, "o nome acompanha a linha inserida"
    Set r = PptLink_ResolveRange(wb, "TQQ_X_naoexiste", PptLink_RefOf(ws.Range("E2")))
    AssertEqual "$E$2", r.Address, "reserva pelo endereco"
    AssertTrue PptLink_ResolveRange(wb, "TQQ_X_naoexiste", "") Is Nothing
    AssertEqual ws.Range("E2").Text, PptText_CellText(ws.Range("E2")), "o texto exibido (formato do Excel)"
    ws.Columns(5).ColumnWidth = 1
    AssertTrue InStr(PptText_CellText(ws.Range("E2")), "#") = 0, "coluna estreita nao vira ####"
    CloseTemp wb
End Sub

Public Sub Test_PptAuto_IsCandidate_ByLinkName()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    wb.SaveAs TempDir() & "\ppttext-cand.xlsx", xlOpenXMLWorkbook
    AssertFalse PptAuto_IsCandidate(wb), "sem _ThinQQ e sem nomes"
    PptLink_DefineName ws.Range("A1"), "TQQ_X_"
    AssertTrue PptAuto_IsCandidate(wb), "texto vinculado"
    CloseTemp wb
End Sub
```
  E em `TestPptSync.bas` (PowerPoint real):

```vb
' ---- Parte E1: texto vinculado ----

Private Function NewTextBox(pres As Object, txt As String) As Object
    Dim sl As Object
    Set sl = pres.Slides.Add(pres.Slides.Count + 1, 12)
    Set NewTextBox = sl.Shapes.AddTextbox(1, 20, 20, 400, 40)
    NewTextBox.TextFrame.TextRange.Text = txt
End Function

Public Sub Test_PptText_LinkAndSync_WhenAvailable()
    Dim wb As Workbook, ws As Worksheet, app As Object, pres As Object, tb As Object, own As New Collection, en As Long, ed As String, s As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    If UserWindowAppeared(app) Then GoTo Cleanup
    ws.Range("E1").Value = 123
    wb.SaveAs TempDir() & "\ppttext-sync.xlsx", xlOpenXMLWorkbook
    Set tb = NewTextBox(pres, "Receita de X mil")
    tb.TextFrame.TextRange.Characters(12, 1).Font.Bold = -1
    AssertEqual "", PptText_Link(tb, 12, 1, ws.Range("E1"))
    AssertEqual "Receita de 123 mil", tb.TextFrame.TextRange.Text
    AssertEqual "text", PptSync_ItemKind(tb)
    ws.Range("E1").Value = 4567
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertEqual "Receita de 4567 mil", tb.TextFrame.TextRange.Text
    AssertEqual -1, CLng(tb.TextFrame.TextRange.Characters(12, 4).Font.Bold), "formatacao do campo mantida"
    AssertEqual 0, CLng(tb.TextFrame.TextRange.Characters(1, 7).Font.Bold), "resto do texto intacto"
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Sem mudança: 1.") > 0, s
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set tb = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 4: o usuario editou o texto fora do campo: nada e sobrescrito; o resumo conta.
Public Sub Test_PptText_Sync_EditedOutsideField_NotTouched_WhenAvailable()
    Dim wb As Workbook, ws As Worksheet, app As Object, pres As Object, tb As Object, own As New Collection, en As Long, ed As String, s As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    If UserWindowAppeared(app) Then GoTo Cleanup
    ws.Range("A1").Value = 10
    wb.SaveAs TempDir() & "\ppttext-edit.xlsx", xlOpenXMLWorkbook
    Set tb = NewTextBox(pres, "Total: X")
    AssertEqual "", PptText_Link(tb, 8, 1, ws.Range("A1"))
    tb.TextFrame.TextRange.Text = "Outro texto do usuario"
    ws.Range("A1").Value = 20
    s = PptSync_Presentation(pres, True)
    AssertEqual "Outro texto do usuario", tb.TextFrame.TextRange.Text
    AssertTrue InStr(s, "Texto editado no PowerPoint fora do campo vinculado (vincule de novo): 1.") > 0, s
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set tb = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 5 (mesmo padrao de Test_PptSync_HostileSourceMacros_NeverRun_Fake): uma forma de TEXTO vinculado (duble) citando uma .xlsm
' hostil fechada com Workbook_BeforeClose que grava um marcador: a sincronizacao abre e fecha a pasta sem rodar a macro.
Public Sub Test_PptText_HostileSourceMacros_NeverRun_Fake()
    Dim wb As Workbook, dir_ As String, xlsm As String, marker As String, cm As Object, pres As TestPptFake, sl As TestPptFake, shp As TestPptFake
    Dim tags As Object, sec0 As Long, prevEv As Boolean, s As String, n0 As Long, code As String
    dir_ = TempDir()
    xlsm = dir_ & "\ppttext-hostil.xlsm": marker = dir_ & "\marcador-texto.txt"
    Set wb = NewTempWorkbook()
    code = "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf
    code = code & "    Dim f As Integer: f = FreeFile: Open """ & marker & """ For Output As #f: Print #f, ""macro rodou"": Close #f" & vbCrLf
    code = code & "End Sub"
    Set cm = wb.VBProject.VBComponents(wb.CodeName).CodeModule
    cm.AddFromString code
    Set cm = Nothing
    wb.SaveAs xlsm, 52
    prevEv = Application.EnableEvents: Application.EnableEvents = False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Set wb = Nothing
    If Len(Dir(marker)) > 0 Then Kill marker
    Set tags = CreateObject("Scripting.Dictionary")
    tags.Add PPTTEXT_TAG, "1"
    tags.Add PPTSYNC_TAG_BOOK, xlsm
    tags.Add PPTLINK_TAG_REF, "Planilha1" & Chr$(31) & "$A$1"
    Set shp = New TestPptFake: Set shp.Tags = tags
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add shp
    Set pres = New TestPptFake: Set pres.Slides = New Collection: pres.Slides.Add sl
    Application.EnableEvents = True
    sec0 = Application.AutomationSecurity
    n0 = Application.Workbooks.Count
    s = PptSync_Presentation(pres, True)
    Application.EnableEvents = prevEv
    AssertEqual 0, Len(Dir(marker)), "a macro da pasta de origem NAO rodou (" & s & ")"
    AssertEqual sec0, CLng(Application.AutomationSecurity), "AutomationSecurity restaurada"
    AssertEqual n0, Application.Workbooks.Count, "a pasta aberta pela sincronizacao foi fechada"
End Sub
```
  (O dublê não tem `TextFrame`: `PptText_Sync` cai em `PPTSYNC_R_FAIL` depois de abrir a pasta — o que o teste prova é a abertura/fechamento sem macro. Se `TestPptFake` não aceitar `Tags` como `Dictionary`, use o mesmo arranjo do teste da Parte A, que já o faz.)

- [ ] **Passo 2:** RED.
- [ ] **Passo 3:** crie `src/core/modules/modPptText.bas`:

```vb
Attribute VB_Name = "modPptText"
Option Explicit
' Texto vinculado a celula (1.9.0; "Text Linked to Excel" do think-cell; spec 2026-10-15 secao 5.3). UM campo por forma: o texto e
' prefixo + valor exibido da celula + sufixo. A celula e lembrada por um nome definido OCULTO na pasta de origem (acompanha linhas inseridas)
' e por um endereco de reserva. Atualizar troca so os caracteres do campo (a formatacao do resto fica); edicao do usuario fora do campo nunca
' e sobrescrita. A pasta de origem e aberta so pelo motor (PptSync_OpenBook: somente leitura, macros desligadas).
Public Const PPTTEXT_TAG As String = "TQQ_TEXT"
Public Const PPTLINK_TAG_NAME As String = "TQQ_TNAME"
Public Const PPTLINK_TAG_REF As String = "TQQ_TREF"
Public Const PPTTEXT_TAG_PRE As String = "TQQ_TPRE"
Public Const PPTTEXT_TAG_SUF As String = "TQQ_TSUF"
Public Const PPTTEXT_TAG_LAST As String = "TQQ_TLAST"
Public Const PPTTEXT_MSG_UNSAVED As String = "Salve a pasta de trabalho antes de vincular: o vínculo aponta para o arquivo salvo."
Public Const PPTTEXT_MSG_NOTEXT As String = "Selecione uma caixa de texto (ou parte do texto dela) no PowerPoint."
Public Const PPTTEXT_MSG_NOCELL As String = "Selecione no Excel a célula com o valor e tente de novo."

' Onde esta o campo no texto atual: texto = prefixo + ultimo + sufixo, ou ainda comeca pelo prefixo e termina pelo sufixo (o usuario editou so
' o campo). False = edicao incompativel (nao tocar). Pura.
Public Function PptText_Locate(cur As String, pre As String, suf As String, last As String, ByRef startPos As Long, ByRef length As Long) As Boolean
    If cur = pre & last & suf Then
        startPos = Len(pre) + 1: length = Len(last): PptText_Locate = True
        Exit Function
    End If
    If Len(cur) < Len(pre) + Len(suf) Then Exit Function
    If Left$(cur, Len(pre)) <> pre Then Exit Function
    If Right$(cur, Len(suf)) <> suf Then Exit Function
    startPos = Len(pre) + 1: length = Len(cur) - Len(pre) - Len(suf)
    PptText_Locate = True
End Function

' Texto exibido da celula (formato do Excel). Coluna estreita ("####"): o valor formatado pela funcao TEXTO com o formato da celula; erro =>
' o valor cru. Sonda pendente P7 confirma datas e formatos com localidade.
Public Function PptText_CellText(c As Range) As String
    Dim t As String, v As Variant
    On Error Resume Next
    t = CStr(c.Cells(1, 1).Text)
    If Len(t) > 0 And Len(Replace(t, "#", "")) = 0 Then
        v = c.Cells(1, 1).Value2
        t = ""
        t = CStr(Application.WorksheetFunction.Text(v, c.Cells(1, 1).NumberFormat))
        If Err.Number <> 0 Or Len(t) = 0 Then Err.Clear: t = CStr(v)
    End If
    PptText_CellText = t
End Function

' Endereco de reserva: planilha + Chr(31) + endereco absoluto.
Public Function PptLink_RefOf(rng As Range) As String
    PptLink_RefOf = rng.Worksheet.Name & Chr$(31) & rng.Address
End Function

' Nome definido OCULTO na pasta de rng (prefixo TQQ_X_ texto, TQQ_T_ tabela). A pasta fica "nao salva" (o usuario e avisado por quem chama).
Public Function PptLink_DefineName(rng As Range, prefix As String) As String
    Dim wb As Workbook, nm As String, i As Long, n As Name
    Set wb = rng.Worksheet.Parent
    Do
        nm = prefix & Registry_NewId(wb)
        Set n = Nothing
        On Error Resume Next
        Set n = wb.Names(nm)
        Err.Clear
        On Error GoTo 0
        i = i + 1
    Loop While Not n Is Nothing And i < 50
    wb.Names.Add Name:=nm, RefersTo:="='" & Replace(rng.Worksheet.Name, "'", "''") & "'!" & rng.Address, Visible:=False
    PptLink_DefineName = nm
End Function

' Intervalo do vinculo: pelo nome (acompanha insercoes) ou, sem ele, pelo endereco de reserva. Nothing se nenhum resolve. Nunca propaga erro.
Public Function PptLink_ResolveRange(wb As Workbook, nm As String, ref As String) As Range
    Dim p As Long, r As Range
    On Error Resume Next
    If Len(nm) > 0 Then Set r = wb.Names(nm).RefersToRange
    Err.Clear
    If r Is Nothing And Len(ref) > 0 Then
        p = InStr(ref, Chr$(31))
        If p > 1 Then Set r = wb.Worksheets(Left$(ref, p - 1)).Range(Mid$(ref, p + 1))
        Err.Clear
    End If
    Set PptLink_ResolveRange = r
End Function

Private Sub PptText_InsertAt(tr As Object, st As Long, v As String)
    If st > Len(CStr(tr.Text)) Then tr.InsertAfter v Else tr.Characters(st, 0).InsertBefore v
End Sub

' Vincula a forma do slide a celula: o trecho [startPos, startPos + length) do texto vira o valor exibido da celula (length 0 = inserir no
' cursor; a forma inteira = 1 e Len). Pasta nao salva ou forma sem texto: devolve a mensagem e nao muda nada. Propaga erro do PowerPoint.
Public Function PptText_Link(shp As Object, startPos As Long, length As Long, cell As Range) As String
    Dim wb As Workbook, tr As Object, cur As String, pre As String, suf As String, v As String, st As Long, ln As Long, hasTf As Boolean
    Set wb = cell.Worksheet.Parent
    If Len(wb.Path) = 0 Then PptText_Link = PPTTEXT_MSG_UNSAVED: Exit Function
    On Error Resume Next
    hasTf = CBool(shp.HasTextFrame)
    On Error GoTo 0
    If Not hasTf Then PptText_Link = PPTTEXT_MSG_NOTEXT: Exit Function
    Set tr = shp.TextFrame.TextRange
    cur = CStr(tr.Text)
    st = startPos: ln = length
    If st < 1 Then st = 1
    If st > Len(cur) + 1 Then st = Len(cur) + 1
    If ln < 0 Then ln = 0
    If st + ln - 1 > Len(cur) Then ln = Len(cur) - st + 1
    pre = Left$(cur, st - 1): suf = Mid$(cur, st + ln)
    v = PptText_CellText(cell)
    If ln > 0 Then tr.Characters(st, ln).Text = v Else PptText_InsertAt tr, st, v
    shp.Tags.Add PPTTEXT_TAG, "1"
    shp.Tags.Add PPTSYNC_TAG_BOOK, wb.FullName
    shp.Tags.Add PPTLINK_TAG_NAME, PptLink_DefineName(cell.Cells(1, 1), "TQQ_X_")
    shp.Tags.Add PPTLINK_TAG_REF, PptLink_RefOf(cell.Cells(1, 1))
    shp.Tags.Add PPTTEXT_TAG_PRE, pre
    shp.Tags.Add PPTTEXT_TAG_SUF, suf
    shp.Tags.Add PPTTEXT_TAG_LAST, v
    If PptAuto_NewLinksAuto() Then PptSync_SetAuto shp, True
End Function

' Atualiza o campo da forma pela pasta wb (aberta). PPTSYNC_R_SAME sem mudanca (e sem force); R_NOCHART celula nao encontrada; R_EDITED
' edicao incompativel (nada tocado); R_OK; R_FAIL (registrado). PowerPoint ocupado sobe.
Public Function PptText_Sync(shp As Object, wb As Workbook, force As Boolean) As Long
    Dim rng As Range, v As String, cur As String, pre As String, suf As String, last As String, st As Long, ln As Long, bn As Long, bd As String
    On Error GoTo Fail
    Set rng = PptLink_ResolveRange(wb, PptSync_Tag(shp, PPTLINK_TAG_NAME), PptSync_Tag(shp, PPTLINK_TAG_REF))
    If rng Is Nothing Then PptText_Sync = PPTSYNC_R_NOCHART: Exit Function
    v = PptText_CellText(rng)
    pre = PptSync_Tag(shp, PPTTEXT_TAG_PRE): suf = PptSync_Tag(shp, PPTTEXT_TAG_SUF): last = PptSync_Tag(shp, PPTTEXT_TAG_LAST)
    cur = CStr(shp.TextFrame.TextRange.Text)
    If Not force And v = last And cur = pre & last & suf Then PptText_Sync = PPTSYNC_R_SAME: Exit Function
    If Not PptText_Locate(cur, pre, suf, last, st, ln) Then PptText_Sync = PPTSYNC_R_EDITED: Exit Function
    If Mid$(cur, st, ln) <> v Then
        If ln > 0 Then shp.TextFrame.TextRange.Characters(st, ln).Text = v Else PptText_InsertAt shp.TextFrame.TextRange, st, v
    End If
    shp.Tags.Add PPTTEXT_TAG_LAST, v
    PptText_Sync = PPTSYNC_R_OK
    Exit Function
Fail:
    bn = Err.Number: bd = Err.Description
    On Error GoTo -1
    If Ppt_IsBusyError(bn) Then Err.Raise bn, "PptText_Sync", bd
    Errors_Log "PptText_Sync" & vbTab & bn & vbTab & bd
    PptText_Sync = PPTSYNC_R_FAIL
End Function
```

- [ ] **Passo 4: Motor** — em `modPptSync.bas`: torne `Public` as constantes `PPTSYNC_R_OK..PPTSYNC_R_SKIPPED` e acrescente `Public Const PPTSYNC_R_EDITED As Long = 6`; em `PptSync_ItemKind` acrescente `ElseIf PptSync_Tag(shp, PPTTEXT_TAG) = "1" Then PptSync_ItemKind = "text"` (reescreva como `If ... ElseIf ... End If` em bloco); em `PptSync_SummaryText` acrescente o parâmetro `Optional nEdited As Long = 0` e, antes da linha de falha, `If nEdited > 0 Then s = s & vbLf & "Texto editado no PowerPoint fora do campo vinculado (vincule de novo): " & nEdited & "."`; em `PptSync_Items`:
  - declare `Dim kind As String, r As Long, nEdited As Long`;
  - troque `ls = Ppt_LinkState(shp)` por `kind = PptSync_ItemKind(shp)` + `If kind = "chart" Then ls = Ppt_LinkState(shp) Else ls = 1` e `path = PptSync_SourcePath(PptSync_LinkSource(shp), ...)` por `path = PptSync_ItemBook(shp)`;
  - troque `Select Case PptSync_SyncOne(shp, wb, force)` por `If kind = "text" Then r = PptText_Sync(shp, wb, force) Else r = PptSync_SyncOne(shp, wb, force)` + `Select Case r` com o caso novo `Case PPTSYNC_R_EDITED: nEdited = nEdited + 1`;
  - passe `nEdited` ao `PptSync_SummaryText`.
  Em `modPptAuto.bas`, `PptAuto_IsCandidate`: depois de `PptAuto_IsCandidate = Not ws Is Nothing`, acrescente:

```vb
    If Not PptAuto_IsCandidate Then
        Dim n As Object
        For Each n In wb.Names
            If Left$(n.Name, 6) = "TQQ_X_" Or Left$(n.Name, 6) = "TQQ_T_" Then PptAuto_IsCandidate = True: Exit For
        Next
    End If
```
  (mova o `Dim n` para o topo da função se o lint exigir.)
- [ ] **Passo 5: API** — `CORE_PPT_API = 5` (comentário: "5 = Core_PptLinkText (1.9.0)") e em `modCoreApi.bas`:

```vb
' API 5 (1.9.0): o .ppam vincula a forma do slide a celula ATIVA deste Excel. Devolve "" (ok) ou a mensagem; nunca propaga.
Public Function Core_PptLinkText(shp As Object, startPos As Long, length As Long) As String
    Dim en As Long, ed As String
    On Error GoTo Fail
    If TypeName(Selection) <> "Range" Then Core_PptLinkText = PPTTEXT_MSG_NOCELL: Exit Function
    Core_PptLinkText = PptText_Link(shp, startPos, length, ActiveCell)
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Errors_Log "Core_PptLinkText" & vbTab & en & vbTab & ed
    If Ppt_IsBusyError(en) Then Core_PptLinkText = PPT_ERR_BUSY Else Core_PptLinkText = "Não foi possível vincular o texto: " & ed & " (" & en & ")"
End Function
```
- [ ] **Passo 6 (GREEN):** `-Filter "TestPptText|TestPptSync|TestCoreApi"` → 0 falhas; suíte completa com `-Ppt`. Commits `feat: texto do slide vinculado a celula do Excel` e `test: texto vinculado com pasta hostil fechada nao roda macros` (+ trailer).

### Tarefa E2: Título do gráfico ligado a uma célula — modelo: standard

**Pré-requisito:** sonda P6 (o `Core_SafeRefresh` mantém o título? o rebuild o apaga?). O plano assume "mantém no refresh, apaga no rebuild".

**Files:** Modify `src/core/modules/modPptText.bas` (`LinkTitle_Apply`, `LinkTitle_Reapply`), `src/core/modules/modCoreApi.bas` (`Core_LinkTitle`, `Core_TestSetTitleCell`, reaplicar em `Core_RebuildFromRange`); Test `TestPptText.bas`.

**Interfaces:** `LinkTitle_Apply(c As Chart, cell As Range)` (Nothing = remove); `LinkTitle_Reapply(c As Chart)`; opção do gráfico `titleref` (planilha + Chr(31) + endereço); `Core_LinkTitle()`; gancho `Core_TestSetTitleCell(cell As Range, Optional cancel As Boolean)`.

- [ ] **Passo 1: Teste (RED)** — em `TestPptText.bas`:

```vb
Public Sub Test_LinkTitle_FollowsCellRefreshAndRebuild()
    Dim wb As Workbook, ws As Worksheet, c As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5): ws.Range("A3:C3").Value = Array("T2", 20, 6)
    ws.Range("E1").Value = "Receita por região"
    ws.Range("A1:C3").Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Core_TestSetTitleCell ws.Range("E1")
    Core_LinkTitle
    AssertTrue c.HasTitle
    AssertEqual "Receita por região", c.ChartTitle.Text
    ws.Range("E1").Value = "Novo título"
    AssertEqual "Novo título", c.ChartTitle.Text, "acompanha a celula"
    Core_SafeRefresh c
    AssertEqual "Novo título", c.ChartTitle.Text, "refresh mantem"
    Core_RebuildFromRange c, ws.Range("A1:C3")
    AssertTrue c.HasTitle, "Alterar dados reaplica"
    AssertEqual "Novo título", c.ChartTitle.Text
    Core_TestSetTitleCell Nothing, True
    Core_LinkTitle
    AssertEqual "", Registry_GetOption(c, "titleref", ""), "cancelar nao muda nada"
    Core_TestSetTitleCell Nothing
    Core_LinkTitle
    AssertFalse c.HasTitle, "sem celula: remove o titulo ligado"
    Watchers_Stop
    CloseTemp wb
End Sub
```
- [ ] **Passo 2: Código** — em `modPptText.bas`:

```vb
' Titulo do grafico ligado a uma celula da mesma pasta (nativo: ChartTitle.Formula; segue para o PowerPoint pelo vinculo do grafico, medido).
' cell Nothing = remove. Grava a opcao "titleref" (reaplicada por LinkTitle_Reapply depois de Alterar dados).
Public Sub LinkTitle_Apply(c As Chart, cell As Range)
    If cell Is Nothing Then
        c.HasTitle = False
        Registry_SetOption c, "titleref", ""
        Exit Sub
    End If
    If Not cell.Worksheet.Parent Is Registry_Workbook(c) Then Err.Raise vbObjectError + 7110, "LinkTitle", "Escolha uma célula da mesma pasta de trabalho do gráfico."
    c.HasTitle = True
    c.ChartTitle.Formula = "='" & Replace(cell.Worksheet.Name, "'", "''") & "'!" & cell.Cells(1, 1).Address
    Registry_SetOption c, "titleref", PptLink_RefOf(cell.Cells(1, 1))
End Sub

' Reaplica o titulo ligado gravado (depois de reconstruir o grafico). Nunca propaga erro.
Public Sub LinkTitle_Reapply(c As Chart)
    Dim ref As String, r As Range
    On Error Resume Next
    ref = Registry_GetOption(c, "titleref", "")
    If Len(ref) = 0 Then Exit Sub
    Set r = PptLink_ResolveRange(Registry_Workbook(c), "", ref)
    If Not r Is Nothing Then LinkTitle_Apply c, r
    Err.Clear
End Sub
```
  Em `modCoreApi.bas`: `Private mTestTitleCell As Range`, `Private mTestTitleCancel As Boolean`, `Private mTestTitleSet As Boolean` no topo e:

```vb
Public Sub Core_TestSetTitleCell(cell As Range, Optional cancel As Boolean = False)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    Set mTestTitleCell = cell
    mTestTitleCancel = cancel
    mTestTitleSet = True
End Sub

' Titulo ligado a uma celula (1.9.0): pergunta a celula (Cancelar = nada; caixa vazia = remove o titulo ligado). Nunca propaga erro.
Public Sub Core_LinkTitle()
    Dim c As Chart, v As Variant, cell As Range
    On Error GoTo Fail
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    If BUILD_CONFIG = "Test" And mTestTitleSet Then
        mTestTitleSet = False
        If mTestTitleCancel Then Exit Sub
        Set cell = mTestTitleCell
    Else
        On Error Resume Next
        Set v = Application.InputBox("Clique na célula com o título (deixe vazio e OK para remover o título ligado):", APP_NAME, Type:=8)
        If Err.Number <> 0 Then Err.Clear: On Error GoTo Fail: LinkTitle_Apply c, Nothing: Core_SafeRefresh c: Exit Sub
        On Error GoTo Fail
        If TypeName(v) <> "Range" Then Exit Sub                       ' Cancelar
        Set cell = v
    End If
    LinkTitle_Apply c, cell
    Core_SafeRefresh c
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7110 Then Core_Notify Err.Description, vbExclamation Else Errors_Handle "Core_LinkTitle"
End Sub
```
  (Comportamento do `InputBox Type:=8`: Cancelar devolve `False`; OK com a caixa vazia levanta erro — tratado como "remover". Confirme manualmente na E5.) Em `Core_RebuildFromRange`, no caminho de sucesso, acrescente `LinkTitle_Reapply c` antes da linha da moldura (D1).
- [ ] **Passo 3 (GREEN)** e commit `feat: titulo do grafico ligado a uma celula` (+ trailer).

### Tarefa E3: Entradas de "Texto vinculado" (PowerPoint e Excel) e título na faixa — modelo: standard

**Files:** `src/ppt/modules/modPptTqq.bas`, `modPptRibbon.bas`, `ribbon/customUI14.xml`; `src/core/modules/modCoreApi.bas` (`Core_PptLinkTextFromPpt`); `src/loader/ribbon/customUI14.xml`, `modRibbon.bas`; testes `TestPptAddin.bas`, `TestRibbon.bas`, `Test-Ribbon.ps1`.

**Interfaces:** `.ppam`: `PPT_API_TEXT = 5`, `PPT_MSG_OPEN_EXCEL`, `PptTqq_SelectionTextRange(ByRef shp As Object, ByRef st As Long, ByRef ln As Long) As Boolean`, `PptTqq_LinkTextTo(xl As Object, core As String, shp As Object, st As Long, ln As Long) As String`, `PptRx_OnLinkText`. Excel: `Core_PptLinkTextFromPpt()`, `Rx_OnPptLinkText`, `Rx_OnLinkTitle`.

- [ ] **Passo 1:** `.ppam` — constantes e funções:

```vb
Public Const PPT_API_TEXT As Long = 5
Public Const PPT_MSG_OPEN_EXCEL As String = "Abra a pasta no Excel (com o ThinQQ), selecione a célula com o valor e clique de novo em Texto vinculado ao Excel."

' Selecao de texto (parte ou cursor) ou de UMA forma com texto (o texto inteiro). False se nao houver.
Public Function PptTqq_SelectionTextRange(ByRef shp As Object, ByRef st As Long, ByRef ln As Long) As Boolean
    Dim sel As Object
    On Error GoTo Done
    Set sel = Application.ActiveWindow.Selection
    If sel.Type = 3 Then                                         ' ppSelectionText
        Set shp = sel.ShapeRange(1)
        st = sel.TextRange.Start: ln = sel.TextRange.Length
        PptTqq_SelectionTextRange = True
    ElseIf sel.Type = 2 Then                                     ' ppSelectionShapes
        If sel.ShapeRange.Count <> 1 Then Exit Function
        Set shp = sel.ShapeRange(1)
        If Not CBool(shp.HasTextFrame) Then Exit Function
        st = 1: ln = Len(CStr(shp.TextFrame.TextRange.Text))
        PptTqq_SelectionTextRange = True
    End If
Done:
End Function

' Chama Core_PptLinkText no Excel dado (so com API >= 5). Devolve a mensagem para o usuario ("" = vinculado).
Public Function PptTqq_LinkTextTo(xl As Object, core As String, shp As Object, st As Long, ln As Long) As String
    If PptTqq_CoreApi(xl, core) < PPT_API_TEXT Then PptTqq_LinkTextTo = PPT_MSG_OLD_CORE: Exit Function
    PptTqq_LinkTextTo = CStr(xl.Run("'" & core & "'!Core_PptLinkText", shp, st, ln))
End Function
```
  `modPptRibbon.bas`:

```vb
' Texto vinculado ao Excel (1.9.0): so com um Excel JA aberto com o ThinQQ (a celula ativa dele e a origem); nunca cria um Excel.
Public Sub PptRx_OnLinkText(control As IRibbonControl)
    Dim shp As Object, st As Long, ln As Long, xl As Object, created As Boolean, core As String, msg As String
    On Error GoTo Fail
    If Not PptTqq_SelectionTextRange(shp, st, ln) Then PptTqq_Msg PPTTEXT_MSG_NOTEXT_PPT: Exit Sub
    Set xl = PptTqq_GetExcel(created)
    If xl Is Nothing Or created Then PptTqq_Release xl, created: PptTqq_Msg PPT_MSG_OPEN_EXCEL: Exit Sub
    core = PptTqq_EnsureCore(xl, False)
    If Len(core) = 0 Then PptTqq_Msg PPT_MSG_OPEN_EXCEL: Exit Sub
    msg = PptTqq_LinkTextTo(xl, core, shp, st, ln)
    If Len(msg) = 0 Then msg = "Texto vinculado à célula do Excel."
    PptTqq_Msg msg
    Exit Sub
Fail:
    PptTqq_Msg PptTqq_ErrorText(Err.Number, Err.Description), vbExclamation
End Sub
```
  (`PptTqq_GetExcel`, `PptTqq_EnsureCore`, `PptTqq_Release` são `Private` hoje: torne-as `Public`; acrescente `Public Const PPTTEXT_MSG_NOTEXT_PPT As String = "Selecione uma caixa de texto (ou parte do texto dela)."` em `modPptTqq`.) XML: no `grpPptElem` (rótulo passa a "Elemento selecionado"):

```xml
          <button id="btnPptLinkText" label="Texto vinculado ao Excel" imageMso="TextBoxInsert" onAction="PptRx_OnLinkText"
                  screentip="Texto vinculado ao Excel" supertip="Selecione no Excel a célula e, aqui, a caixa de texto (ou a parte do texto, ou ponha o cursor): o trecho passa a mostrar o valor da célula e acompanha as atualizações."/>
```
  Teste `.ppam` (`TestPptAddin.bas`): `Test_PptTqq_LinkTextTo_OldCore` com um Excel injetado sem API 5 não é possível sem core; teste a parte pura: `PptTqq_SelectionTextRange` devolve False sem janela (`Application.ActiveWindow` inexistente no projeto de teste): `AssertFalse PptTqq_SelectionTextRange(s, a, b)`.
- [ ] **Passo 2:** Excel — `modCoreApi.bas`:

```vb
' Excel: vincula a celula ativa ao texto selecionado na janela ativa do PowerPoint ja aberto (sonda pendente P5: se a selecao do PowerPoint
' nao for legivel com o Excel ativo, este botao avisa para usar o botao do PowerPoint). Nunca propaga erro.
Public Sub Core_PptLinkTextFromPpt()
    Dim app As Object, sel As Object, shp As Object, st As Long, ln As Long, msg As String
    On Error GoTo Fail
    If TypeName(Selection) <> "Range" Then Core_Notify PPTTEXT_MSG_NOCELL: Exit Sub
    Set app = Ppt_GetApp(False)
    If app Is Nothing Then Core_Notify PPTSYNC_NO_PPT: Exit Sub
    Set sel = app.ActiveWindow.Selection
    If sel.Type = 3 Then
        Set shp = sel.ShapeRange(1): st = sel.TextRange.Start: ln = sel.TextRange.Length
    ElseIf sel.Type = 2 Then
        Set shp = sel.ShapeRange(1): st = 1: ln = Len(CStr(shp.TextFrame.TextRange.Text))
    Else
        Core_Notify PPTTEXT_MSG_NOTEXT: Exit Sub
    End If
    msg = PptText_Link(shp, st, ln, ActiveCell)
    If Len(msg) = 0 Then msg = "Texto do slide vinculado à célula " & ActiveCell.Address(False, False) & "."
    Core_Notify msg
    Exit Sub
Fail:
    If Ppt_IsKnownFailure(Err.Number) Then Core_Notify Ppt_ErrorText(Err.Number, Err.Description), vbExclamation Else Errors_Handle "Core_PptLinkTextFromPpt"
End Sub
```
  Faixa do Excel: no `mnuPpt`, depois de `btnPptChart`: `<button id="btnPptLinkText" label="Vincular célula ao texto selecionado" image="tqq_ppt16" onAction="Rx_OnPptLinkText" screentip="Vincular célula ao texto selecionado" supertip="A célula selecionada passa a alimentar o texto selecionado no PowerPoint (a caixa inteira, a parte selecionada ou o ponto do cursor). Também há o botão Texto vinculado ao Excel no PowerPoint."/>` (sem `getEnabled`); no `mnuMore`, depois de `btnResetScale`: `<button id="btnLinkTitle" label="Título ligado a uma célula..." imageMso="ChartTitle" onAction="Rx_OnLinkTitle" getEnabled="Rx_GetEnabled" screentip="Título ligado a uma célula" supertip="O título do gráfico passa a mostrar o texto de uma célula e acompanha as mudanças dela (também no PowerPoint, pelo vínculo do gráfico)."/>` (se `imageMso="ChartTitle"` não existir, o Excel descarta a aba: use `image="tqq_options16"` — o Test-Ribbon só aceita nomes da tabela de ícones ou `imageMso` já usados; prefira `image="tqq_options16"`). Callbacks `Rx_OnPptLinkText` (`RunCoreFeature "Core_PptLinkTextFromPpt"`) e `Rx_OnLinkTitle` (`RunCoreFeature "Core_LinkTitle"` + `Loader_InvalidateRibbon`). Atualize as contagens de `TestRibbon` (botões do `mnuPpt` 5; botões do `mnuMore` +1) e acrescente asserts dos ids/callbacks; `Test-Ribbon.ps1` com o `btnPptLinkText` do PowerPoint.
- [ ] **Passo 3 (GREEN):** suíte completa com `-Ppt`. Commit `feat: botoes de texto vinculado no PowerPoint e no Excel e titulo ligado na faixa` (+ trailer).

### Tarefa E4: Documentação da 1.9.0 — modelo: standard
Tutorial `#ppt-texto` (vincular pelos dois caminhos; um campo por forma; edição do texto em volta e do campo; "Texto editado..."; título ligado; pasta salva; nome definido oculto na pasta (não apagar no Gerenciador de Nomes)); **seção "Think-cell: o que fica de fora"** com Tableau (caminho: exportar a view para uma planilha e usar o ThinQQ, §5.5) e datasheet (Quebrar vínculo pela interface congela as setas; nunca pelo ThinQQ; §5.6); README (manuais 74–80: vincular inteira/parte/cursor pelos dois botões (P5), editar em volta e o campo, data/%/moeda/coluna estreita (P7), título + Alterar dados (P6), Quebrar vínculo pela interface (sonda S10, pela UI)); CHANGELOG/notes; Test-Docs. Commit `docs: texto vinculado a celula e limites da paridade com o think-cell` (+ trailer).

### Tarefa E5: Suíte final, revisão do ramo, verificação manual — modelo: standard (controlador + humano)
Como C6, com Review Focus 4–5; humano: itens 74–80, PR, `publicar-release 1.9.0`.

---

# PARTE F — Tabelas vinculadas (1.10.0)

### Tarefa F1: Modelo puro da tabela — modelo: standard

**Files:** Create `src/core/modules/modPptTable.bas` (parte pura); Test `tests/vba/core/TestPptTable.bas`.

**Interfaces:** `PPTTABLE_TAG = "TQQ_TABLE"`, `PPTTABLE_TAG_LAST = "TQQ_TBLAST"`, `PPTTABLE_MAX_ROWS = 75`, `PPTTABLE_MAX_COLS = 25`; `PptTable_Serialize(m As Variant) As String`; `PptTable_Parse(s As String) As Variant` (matriz 1..r × 1..c de String, ou `Empty`); `PptTable_Diff(oldM As Variant, newM As Variant) As Collection` (itens `Array(r, c)`); `PptTable_TextMatrix(rng As Range) As Variant`; `PptTable_RefusalText(rng As Range) As String`.

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestPptTable.bas`:

```vb
Attribute VB_Name = "TestPptTable"
Option Explicit
' 1.10.0: tabela do PowerPoint vinculada a intervalo (spec 2026-10-15 secao 5.4). Partes puras.

Private Function M2(a As Variant, rows As Long, cols As Long) As Variant
    Dim m() As String, r As Long, c As Long
    ReDim m(1 To rows, 1 To cols)
    For r = 1 To rows
        For c = 1 To cols
            m(r, c) = CStr(a((r - 1) * cols + c - 1))
        Next
    Next
    M2 = m
End Function

Public Sub Test_PptTable_SerializeParseRoundTrip()
    Dim m As Variant, p As Variant
    m = M2(Array("Ano", "Sul", "T1", "1.234,5"), 2, 2)
    p = PptTable_Parse(PptTable_Serialize(m))
    AssertEqual 2, UBound(p, 1): AssertEqual 2, UBound(p, 2)
    AssertEqual "1.234,5", p(2, 2)
    AssertTrue IsEmpty(PptTable_Parse("")), "vazio"
    m = M2(Array("", ""), 1, 2)
    p = PptTable_Parse(PptTable_Serialize(m))
    AssertEqual 2, UBound(p, 2), "celulas vazias contam"
End Sub

Public Sub Test_PptTable_Diff()
    Dim a As Variant, b As Variant, d As Collection
    a = M2(Array("x", "1", "y", "2"), 2, 2)
    b = M2(Array("x", "1", "y", "3"), 2, 2)
    Set d = PptTable_Diff(a, b)
    AssertEqual 1, d.Count: AssertEqual 2, d(1)(0): AssertEqual 2, d(1)(1)
    b = M2(Array("x", "1", "y", "2", "z", "5"), 3, 2)
    Set d = PptTable_Diff(a, b)
    AssertEqual 2, d.Count, "linha nova: as duas celulas"
    Set d = PptTable_Diff(Empty, a)
    AssertEqual 4, d.Count, "sem historico: tudo"
    Set d = PptTable_Diff(a, a)
    AssertEqual 0, d.Count
End Sub

Public Sub Test_PptTable_TextMatrixAndRefusals()
    Dim wb As Workbook, ws As Worksheet, m As Variant
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B2").Value = 1234.5: ws.Range("A1:B2").NumberFormat = "#,##0.0"
    m = PptTable_TextMatrix(ws.Range("A1:B2"))
    AssertEqual ws.Range("A1").Text, m(1, 1)
    AssertEqual "", PptTable_RefusalText(ws.Range("A1:B2"))
    ws.Range("D1:E1").Merge
    AssertTrue Len(PptTable_RefusalText(ws.Range("D1:E2"))) > 0, "mescladas recusadas"
    AssertTrue Len(PptTable_RefusalText(ws.Range("A1:A80"))) > 0, "mais de 75 linhas"
    CloseTemp wb
End Sub
```

- [ ] **Passo 2:** RED; **Passo 3:** crie `src/core/modules/modPptTable.bas` (parte pura):

```vb
Attribute VB_Name = "modPptTable"
Option Explicit
' Tabela nativa do PowerPoint vinculada a um intervalo do Excel (1.10.0; "Create tables with Excel formatting" do think-cell; spec 2026-10-15
' secao 5.4). A ultima matriz de textos escrita fica na etiqueta TQQ_TBLAST (linhas por Chr(30), celulas por Chr(31)): a atualizacao escreve so
' as celulas cuja ORIGEM mudou (a formatacao feita no PowerPoint fica). Limites do PowerPoint: 75 linhas e 75 colunas (o ThinQQ limita a 25).
Public Const PPTTABLE_TAG As String = "TQQ_TABLE"
Public Const PPTTABLE_TAG_LAST As String = "TQQ_TBLAST"
Public Const PPTTABLE_MAX_ROWS As Long = 75
Public Const PPTTABLE_MAX_COLS As Long = 25

Public Function PptTable_Serialize(m As Variant) As String
    Dim r As Long, c As Long, s As String
    For r = 1 To UBound(m, 1)
        If r > 1 Then s = s & Chr$(30)
        For c = 1 To UBound(m, 2)
            If c > 1 Then s = s & Chr$(31)
            s = s & Replace(Replace(CStr(m(r, c)), Chr$(30), " "), Chr$(31), " ")
        Next
    Next
    PptTable_Serialize = s
End Function

Public Function PptTable_Parse(s As String) As Variant
    Dim rows() As String, cells() As String, m() As String, r As Long, c As Long, nc As Long
    If Len(s) = 0 Then Exit Function
    rows = Split(s, Chr$(30))
    nc = UBound(Split(rows(0), Chr$(31))) + 1
    ReDim m(1 To UBound(rows) + 1, 1 To nc)
    For r = 0 To UBound(rows)
        cells = Split(rows(r), Chr$(31))
        For c = 0 To nc - 1
            If c <= UBound(cells) Then m(r + 1, c + 1) = cells(c)
        Next
    Next
    PptTable_Parse = m
End Function

' Celulas de newM a escrever: diferentes de oldM ou fora das dimensoes dele (oldM Empty = todas). Pura.
Public Function PptTable_Diff(oldM As Variant, newM As Variant) As Collection
    Dim out As New Collection, r As Long, c As Long, inOld As Boolean
    For r = 1 To UBound(newM, 1)
        For c = 1 To UBound(newM, 2)
            inOld = False
            If Not IsEmpty(oldM) Then inOld = (r <= UBound(oldM, 1) And c <= UBound(oldM, 2))
            If Not inOld Then
                out.Add Array(r, c)
            ElseIf CStr(oldM(r, c)) <> CStr(newM(r, c)) Then
                out.Add Array(r, c)
            End If
        Next
    Next
    Set PptTable_Diff = out
End Function

Public Function PptTable_TextMatrix(rng As Range) As Variant
    Dim m() As String, r As Long, c As Long
    ReDim m(1 To rng.Rows.Count, 1 To rng.Columns.Count)
    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            m(r, c) = PptText_CellText(rng.Cells(r, c))
        Next
    Next
    PptTable_TextMatrix = m
End Function

' "" = o intervalo serve; senao o motivo (mescladas, varias areas, tamanho).
Public Function PptTable_RefusalText(rng As Range) As String
    If rng.Areas.Count > 1 Then PptTable_RefusalText = "Selecione um único intervalo retangular.": Exit Function
    If rng.Rows.Count > PPTTABLE_MAX_ROWS Or rng.Columns.Count > PPTTABLE_MAX_COLS Then PptTable_RefusalText = "A tabela vinculada aceita até " & PPTTABLE_MAX_ROWS & " linhas e " & PPTTABLE_MAX_COLS & " colunas.": Exit Function
    If IsNull(rng.MergeCells) Or rng.MergeCells Then PptTable_RefusalText = "A tabela vinculada não aceita células mescladas no intervalo."
End Function
```
- [ ] **Passo 4 (GREEN):** `-Filter "TestPptTable"`; commit `feat: modelo da tabela vinculada (serializacao e diferenca)` (+ trailer).

### Tarefa F2: Criar e atualizar a tabela no slide — modelo: most-capable

**Pré-requisito opcional:** sonda P8 (preencher pelo `.ppam`); o plano preenche pelo core.

**Files:** `src/core/modules/modPptTable.bas` (`PptTable_Insert`, `PptTable_Sync`, `PptTable_CopyFormat`), `modPptSync.bas` (`ItemKind` "table", rota), `modCoreApi.bas` (`Core_PptInsertTable`); testes `TestPptSync.bas`.

**Interfaces:** `PptTable_Insert(rng As Range, slide As Object) As Object`; `PptTable_Sync(shp As Object, wb As Workbook, force As Boolean) As Long`; `Core_PptInsertTable()`.

- [ ] **Passo 1: Testes (RED)** — `TestPptSync.bas`:

```vb
' ---- Parte F2: tabela vinculada ----

Public Sub Test_PptTable_InsertAndSync_WhenAvailable()
    Dim wb As Workbook, ws As Worksheet, app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String
    Dim shp As Object, t As Object, s As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    If UserWindowAppeared(app) Then GoTo Cleanup
    ws.Range("A1:C1").Value = Array("Região", "2025", "2026"): ws.Range("A1:C1").Font.Bold = True
    ws.Range("A2:C2").Value = Array("Sul", 10, 12): ws.Range("A3:C3").Value = Array("Norte", 5, 7)
    wb.SaveAs TempDir() & "\ppttable.xlsx", xlOpenXMLWorkbook
    Set shp = PptTable_Insert(ws.Range("A1:C3"), slide)
    Set t = shp.Table
    AssertEqual 3, t.Rows.Count: AssertEqual 3, t.Columns.Count
    AssertEqual "Norte", t.Cell(3, 1).Shape.TextFrame.TextRange.Text
    AssertEqual -1, CLng(t.Cell(1, 2).Shape.TextFrame.TextRange.Font.Bold), "negrito do Excel"
    AssertEqual "table", PptSync_ItemKind(shp)
    t.Cell(2, 2).Shape.TextFrame.TextRange.Font.Italic = -1          ' formatacao feita no PowerPoint
    ws.Range("B2").Value = 99
    ws.Rows(4).Insert: ws.Range("A4:C4").Value = Array("Leste", 1, 2)   ' fora do nome: nao entra
    ws.Rows(3).Insert: ws.Range("A3:C3").Value = Array("Oeste", 3, 4)   ' dentro: entra
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertEqual "99", t.Cell(2, 2).Shape.TextFrame.TextRange.Text
    AssertEqual -1, CLng(t.Cell(2, 2).Shape.TextFrame.TextRange.Font.Italic), "formatacao do PowerPoint mantida"
    AssertEqual 4, t.Rows.Count, "linha inserida dentro do intervalo"
    AssertEqual "Oeste", t.Cell(3, 1).Shape.TextFrame.TextRange.Text
    ws.Rows(3).Delete
    s = PptSync_Presentation(pres, False)
    AssertEqual 3, t.Rows.Count, "linha apagada"
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Sem mudança: 1.") > 0, s
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set t = Nothing: Set shp = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

Public Sub Test_PptInsertTable_Unsaved_NothingCreated()
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    wb.Worksheets(1).Range("A1:B2").Value = 1
    wb.Worksheets(1).Range("A1:B2").Select
    Core_ClearLastNotice
    Core_PptInsertTable
    AssertEqual PPTTEXT_MSG_UNSAVED, Core_LastNotice()
    CloseTemp wb
End Sub
```
- [ ] **Passo 2: Código** — acrescente a `modPptTable.bas`:

```vb
' Alinhamento do Excel para o do PowerPoint (ppAlignLeft 1, ppAlignCenter 2, ppAlignRight 3). Geral: numero a direita, texto a esquerda.
Private Function PptTable_Align(src As Range) As Long
    Select Case src.HorizontalAlignment
        Case xlRight: PptTable_Align = 3
        Case xlCenter, xlCenterAcrossSelection: PptTable_Align = 2
        Case xlLeft: PptTable_Align = 1
        Case Else: If IsNumeric(src.Value2) And Not IsEmpty(src.Value2) Then PptTable_Align = 3 Else PptTable_Align = 1
    End Select
End Function

' Formatacao inicial da celula (so na criacao e nas linhas/colunas novas): negrito, italico, cor e nome da fonte, tamanho, preenchimento,
' alinhamento. Melhor esforco: falha so registrada.
Private Sub PptTable_CopyFormat(src As Range, cell As Object)
    On Error Resume Next
    With cell.Shape.TextFrame.TextRange
        .Font.Bold = IIf(src.Font.Bold, -1, 0)
        .Font.Italic = IIf(src.Font.Italic, -1, 0)
        .Font.Color.RGB = src.Font.Color
        .Font.Name = src.Font.Name
        .Font.Size = src.Font.Size
        .ParagraphFormat.Alignment = PptTable_Align(src)
    End With
    If src.Interior.ColorIndex = xlNone Then
        cell.Shape.Fill.Visible = 0
    Else
        cell.Shape.Fill.Visible = -1
        cell.Shape.Fill.Solid
        cell.Shape.Fill.ForeColor.RGB = src.Interior.Color
    End If
    If Err.Number <> 0 Then Errors_Log "PptTable_CopyFormat" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Cria a tabela no slide a partir de rng (pasta salva; quem chama confere). 80% da largura do slide, centralizada; etiquetas do vinculo.
Public Function PptTable_Insert(rng As Range, slide As Object) As Object
    Dim m As Variant, shp As Object, t As Object, r As Long, c As Long, sw As Double, sh As Double, w As Double, h As Double, msg As String
    msg = PptTable_RefusalText(rng)
    If Len(msg) > 0 Then Err.Raise vbObjectError + 7120, "PptTable", msg
    m = PptTable_TextMatrix(rng)
    sw = slide.Parent.PageSetup.SlideWidth: sh = slide.Parent.PageSetup.SlideHeight
    If sw <= 0 Or sh <= 0 Then sw = 960: sh = 540
    w = 0.8 * sw: h = rng.Rows.Count * 22
    If h > 0.8 * sh Then h = 0.8 * sh
    Set shp = slide.Shapes.AddTable(rng.Rows.Count, rng.Columns.Count, (sw - w) / 2, (sh - h) / 2, w, h)
    Set t = shp.Table
    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            t.Cell(r, c).Shape.TextFrame.TextRange.Text = m(r, c)
            PptTable_CopyFormat rng.Cells(r, c), t.Cell(r, c)
        Next
    Next
    shp.Tags.Add PPTTABLE_TAG, "1"
    shp.Tags.Add PPTSYNC_TAG_BOOK, rng.Worksheet.Parent.FullName
    shp.Tags.Add PPTLINK_TAG_NAME, PptLink_DefineName(rng, "TQQ_T_")
    shp.Tags.Add PPTLINK_TAG_REF, PptLink_RefOf(rng)
    shp.Tags.Add PPTTABLE_TAG_LAST, PptTable_Serialize(m)
    If PptAuto_NewLinksAuto() Then PptSync_SetAuto shp, True
    Set PptTable_Insert = shp
End Function

' Atualiza a tabela pela pasta wb: ajusta linhas/colunas (novas herdam a formatacao da ultima; as do fim saem) e escreve so as celulas cuja
' origem mudou (force: todas). Resultados como PptText_Sync. PowerPoint ocupado sobe.
Public Function PptTable_Sync(shp As Object, wb As Workbook, force As Boolean) As Long
    Dim rng As Range, newM As Variant, oldM As Variant, t As Object, d As Collection, e As Variant, resized As Boolean, bn As Long, bd As String
    On Error GoTo Fail
    Set rng = PptLink_ResolveRange(wb, PptSync_Tag(shp, PPTLINK_TAG_NAME), PptSync_Tag(shp, PPTLINK_TAG_REF))
    If rng Is Nothing Then PptTable_Sync = PPTSYNC_R_NOCHART: Exit Function
    If Len(PptTable_RefusalText(rng)) > 0 Then Errors_Log "PptTable_Sync: intervalo recusado" & vbTab & PptTable_RefusalText(rng): PptTable_Sync = PPTSYNC_R_FAIL: Exit Function
    newM = PptTable_TextMatrix(rng)
    If force Then oldM = Empty Else oldM = PptTable_Parse(PptSync_Tag(shp, PPTTABLE_TAG_LAST))
    Set t = shp.Table
    Do While t.Rows.Count < rng.Rows.Count: t.Rows.Add: resized = True: Loop
    Do While t.Rows.Count > rng.Rows.Count: t.Rows(t.Rows.Count).Delete: resized = True: Loop
    Do While t.Columns.Count < rng.Columns.Count: t.Columns.Add: resized = True: Loop
    Do While t.Columns.Count > rng.Columns.Count: t.Columns(t.Columns.Count).Delete: resized = True: Loop
    If resized Then oldM = Empty                                  ' linhas no meio mudam a posicao de tudo abaixo: reescreve os textos
    Set d = PptTable_Diff(oldM, newM)
    If d.Count = 0 Then PptTable_Sync = PPTSYNC_R_SAME: Exit Function
    For Each e In d
        If CStr(t.Cell(e(0), e(1)).Shape.TextFrame.TextRange.Text) <> CStr(newM(e(0), e(1))) Then t.Cell(e(0), e(1)).Shape.TextFrame.TextRange.Text = newM(e(0), e(1))
    Next
    shp.Tags.Add PPTTABLE_TAG_LAST, PptTable_Serialize(newM)
    PptTable_Sync = PPTSYNC_R_OK
    Exit Function
Fail:
    bn = Err.Number: bd = Err.Description
    On Error GoTo -1
    If Ppt_IsBusyError(bn) Then Err.Raise bn, "PptTable_Sync", bd
    Errors_Log "PptTable_Sync" & vbTab & bn & vbTab & bd
    PptTable_Sync = PPTSYNC_R_FAIL
End Function
```
  (`Do While ... : ... : Loop` numa linha é válido em VBA; se o lint recusar, quebre em linhas.) Em `modPptSync.bas`: `PptSync_ItemKind` ganha `ElseIf PptSync_Tag(shp, PPTTABLE_TAG) = "1" Then PptSync_ItemKind = "table"`; em `PptSync_Items`, `If kind = "text" Then ... ElseIf kind = "table" Then r = PptTable_Sync(shp, wb, force) Else ...` (em bloco). Em `modCoreApi.bas`:

```vb
' "Inserir tabela vinculada no slide" (1.10.0): tabela do PowerPoint a partir das celulas selecionadas (pasta salva), no slide ativo (ou numa
' apresentacao nova), com a formatacao inicial do Excel. Falha: desfaz o que criou. Nunca propaga erro.
Public Sub Core_PptInsertTable()
    Dim rng As Range, app As Object, slide As Object, shp As Object, newPres As Object, newSlide As Object, msg As String, en As Long, ed As String
    On Error GoTo Fail
    If TypeName(Selection) <> "Range" Then Core_Notify "Selecione as células da tabela.": Exit Sub
    Set rng = Selection
    If Len(rng.Worksheet.Parent.Path) = 0 Then Core_Notify PPTTEXT_MSG_UNSAVED, vbExclamation: Exit Sub
    msg = PptTable_RefusalText(rng)
    If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Sub
    Set app = Ppt_GetApp(True)
    If app Is Nothing Then Core_Notify PPT_ERR_NOT_INSTALLED, vbExclamation: Exit Sub
    Set slide = Ppt_TargetSlide(app, newPres, newSlide)
    Set shp = PptTable_Insert(rng, slide)
    Core_Notify "Tabela vinculada criada no slide " & slide.SlideIndex & " de '" & slide.Parent.Name & "'. Ela acompanha a pasta nas atualizações (manual, ao abrir ou automática)."
    If BUILD_CONFIG <> "Test" Then
        On Error Resume Next
        app.Visible = -1: app.Activate
        Err.Clear
    End If
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    If Not shp Is Nothing Then shp.Delete
    If Not newSlide Is Nothing And newPres Is Nothing Then newSlide.Delete
    If Not newPres Is Nothing Then newPres.Saved = -1: newPres.Close
    On Error GoTo 0
    If en = vbObjectError + 7120 Then
        Core_Notify ed, vbExclamation
    ElseIf Ppt_IsKnownFailure(en) Then
        Core_Notify Ppt_ErrorText(en, ed), vbExclamation
    Else
        Errors_HandleErr "Core_PptInsertTable", en, ed
    End If
End Sub
```
- [ ] **Passo 3 (GREEN):** `-Filter "TestPptTable|TestPptSync|TestCoreApi"` → 0 falhas; suíte completa com `-Ppt`. Commit `feat: tabela do PowerPoint vinculada a intervalo do Excel` (+ trailer).

### Tarefa F3: Faixa — "Inserir tabela vinculada no slide" — modelo: standard
`mnuPpt`, depois de `btnPptLinkText`: `<button id="btnPptTable" label="Inserir tabela vinculada no slide" image="tqq_ppt16" onAction="Rx_OnPptTable" screentip="Inserir tabela vinculada no slide" supertip="Cria no slide ativo uma tabela do PowerPoint com as células selecionadas (com a formatação do Excel) que acompanha a pasta nas atualizações. A pasta precisa estar salva; sem células mescladas; até 75 linhas e 25 colunas."/>`; `Rx_OnPptTable` → `RunCoreFeature "Core_PptInsertTable"`; `TestRibbon` (botões do `mnuPpt` 6; asserts do id/callback; sem `getEnabled`). Suíte; commit `feat: botao de tabela vinculada no menu PowerPoint` (+ trailer).

### Tarefa F4: Documentação da 1.10.0 — modelo: standard
Tutorial `#ppt-tabela` (criar, o que é copiado da formatação, o que a atualização muda e o que preserva, linhas/colunas, limites: mescladas, bordas, larguras, nome definido oculto, 75 × 25); README (manuais 81–85: formatação inicial com tema e cores, formatar no PowerPoint e atualizar, inserir/apagar linhas e colunas, 30 × 10 (tempo), automático); matriz de paridade no tutorial (a do spec §6, como tabela); CHANGELOG/notes; Test-Docs. Commit `docs: tabela vinculada e matriz de paridade com o think-cell` (+ trailer).

### Tarefa F5: Suíte final, revisão do ramo, verificação manual — modelo: standard (controlador + humano)
Como C6; humano: itens 81–85, PR, `publicar-release 1.10.0`.

---

## Resumo da estimativa (detalhe, suposições e confiança: spec §12)

| Versão | Agêntico (impl. + revisão + correções + suítes), min | Corrido de agente | Humano (manual + PR + release) |
|---|---|---|---|
| 1.6.0 Atualização automática (C0–C6) | 385 / 620 / 950 | ~5,8 / 9,3 / 14,3 h | 2,5 / 3,75 / 5,5 h |
| 1.7.0 Inserir no slide + moldura (D1–D4) | 195 / 315 / 470 | ~3 / 4,8 / 7 h | 1,5 / 2,4 / 3,7 h |
| 1.8.0 Criar gráficos pelo PowerPoint — fase G-I (G1–G6) | 440 / 695 / 1 060 | ~6,6 / 10,4 / 15,9 h | 2,5 / 3,75 / 5,5 h |
| 1.9.0 Comandos pelo PowerPoint (G7–G11) + texto vinculado (E1–E5) | 565 / 885 / 1 360 | ~8,5 / 13,3 / 20,4 h | 2,5 / 4,1 / 6 h |
| 1.10.0 Tabelas vinculadas (F1–F5) | 295 / 460 / 695 | ~4,4 / 6,9 / 10,4 h | 1,5 / 2,4 / 3,5 h |
| **Total** | **1 880 / 2 975 / 4 535** | **~28 / 44,5 / 68 h ⇒ ≈ 4 / 6,5 / 10 dias corridos** | **10,5 / 16,5 / 24 h** |
| *Parte G acrescenta (emenda 1)* | *+715 / +1 125 / +1 720* | *~+10,7 / +16,9 / +25,8 h* | *+3,5 / +5,5 / +8 h* |

Fora das contas: as sondas pendentes P1–P8 e PG1–PG9 (~1–2 h de agente, depois da gravação de tela; PG2, PG6, PG8, PG9 são manuais) e este planejamento (~3 h com a emenda). Confiança: média-baixa na 1.6.0 (eventos entre aplicativos), média nas 1.8.0/1.9.0, média-alta na 1.7.0.
