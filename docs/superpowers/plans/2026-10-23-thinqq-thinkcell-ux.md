# ThinQQ — Plano de Implementação: UX e funcionalidades como o think-cell, rodada 2 (1.11.0 a 1.13.0)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** (1.11.0, Parte U) o vínculo Excel → PowerPoint passa a funcionar **como o think-cell**: tudo nasce **manual**; quando a pasta muda, o slide
não muda sozinho — o elemento ganha a **bandeira "Dados novos"** (etiqueta `TQQ_STATE`, gravada pelo motor do Excel **sem alterar nada no Excel**) e
um clique em **Atualizar** traz os dados; "Atualizar automaticamente" continua por elemento (opt-in); o diálogo não modal **Vínculos** (o Data Links)
lista todos os elementos com estado e age em lote; a aba contextual **ThinQQ** aparece ao selecionar um gráfico; sucesso não abre caixa; abrir a
apresentação não pergunta. (1.12.0, Parte S) **Estilos de gráfico**: padrões nomeados (cores por papel, hachura, conectores, rótulos, fonte,
espaço), salvos pelo usuário e na pasta, importar/exportar, padrão para novos gráficos, reaplicados pelo core e levados ao PowerPoint, com os
embutidos "Cascata MC" (o exemplo do usuário) e "Cascata DDE". (1.13.0, Parte X) **Pontos de entrada como o think-cell**: "Vincular ao
PowerPoint" (galeria na guia Inserir e na aba ThinQQ do Excel), moldura com estado, alterar intervalo e tipo **sem perder o slide**, botão direito
no gráfico do Excel.

**Architecture:** a de sempre (`ThinQQ.xlam` loader e faixa → `Core_*` em `ThinQQ.core.xlam`; `ThinQQ.ppam` chama o core por `xl.Run`; formas do
slide marcadas por `Shape.Tags` `TQQ_*`; motor `modPptAuto` com eventos do Excel + `OnTime`). Novos: `modPptState` (estado/bandeira e índice
`TQQ_LINKIDX`), `modPptLinks` + `frmTqqLinks` no `.ppam`, `modStyleDef`/`modStyleStore`/`modStyleApply` + `frmStyles` no core, `modChartMenu` (1.13.0).

**Tech Stack:** VBA 7 (Excel/PowerPoint 2016+, 32/64 bits, sem `Declare`, sem referências extras), customUI14, PowerShell 7 e 5.1, Git.

**Spec:** `docs/superpowers/specs/2026-10-23-thinqq-thinkcell-ux-design.md` (cópia, na U0, de `.superpowers/sdd/2026-10-23-thinkcell-ux/design.md`:
pesquisa §2–3, lacunas §4, viabilidade §5, sondas §6 (**pendentes §6.3**), decisões §8, matriz §9, versões §10, riscos §11, pré-flight §12,
estimativa §13, verificação manual §14). Specs anteriores: `docs/superpowers/specs/2026-10-15-thinqq-thinkcell-parity-design.md` e
`2026-10-14-thinqq-ppt-delta-design.md`. Regras de ambiente: `.superpowers/sdd/2026-10-15-thinkcell-parity/implementer-rules.md` (sem exceção).

**Estado de execução:** nenhuma tarefa concluída. Três versões em três branches empilhadas criadas pelo controlador: `feat/ux-vinculos` (1.11.0, de
`main`), `feat/estilos` (1.12.0, da ponta da 1.11.0), `feat/ux-entrada` (1.13.0, da ponta da 1.12.0). O implementador **não** cria nem troca de
branch. Sondas pendentes (spec §6.3): **P-U1** antes/dentro da U1 (passo de sonda), **P-U2/P-U3** na U2/U6 (P-U3 é manual, com o usuário),
**P-U4/P-U5/P-U6** na tarefa E0 antes da X2/X3.

**Ordem e por quê:** U (as duas queixas: "atualizar a cada alteração" e UX; cria estado, índice, formulários no `.ppam` e aba contextual) → S
(pedido explícito; independe do PowerPoint, só a S5 toca o `.ppam`) → X (depende de sondas manuais). **API do core para o `.ppam`:** 6 =
`Core_PptLinkStates` (U1); 7 = `Core_StyleNames` + comando `style` (S5).

## Global Constraints

- **Ambiente:** Excel 2019 32-bit pt-BR e PowerPoint 16.0 32-bit no Windows 11; testes em instâncias próprias invisíveis. Build roda o lint
  (`build/lib/Vba-Lint.ps1`) e compilação completa; compilação ⇒ saída 3 (módulo:linha); travamento ⇒ 2; asserts ⇒ 1; ok ⇒ 0. Fontes VBA em UTF-8 com
  **somente caracteres cp1252** (nada de `→`, `≥`, `−`, `…`, `✓`, aspas tipográficas), `Option Explicit`, declarações antes dos procedimentos,
  `Attribute VB_Name` na 1ª linha, nenhuma continuação ` _` seguida de texto, nada de `ElseIf` em `If` de uma linha nem `If ... Then If ...`,
  nada de `Sub X(): ...: End Sub` numa linha.
- **EDR:** NUNCA iniciar processo filho a partir de VBA (`Shell`, `WScript.Shell.Run/Exec`, `powershell`, `cmd`); `Test-VbaLint` barra. Excel morreu
  (0x800706BE/0x800A9C68)? Olhe primeiro `Microsoft-Windows-SentinelOne/Operational` (eventos 31/32). `ADODB.Stream` e `Scripting.FileSystemObject`
  são objetos COM em processo (permitidos; já usados).
- **Testes:** `Public Sub Test_x()` sem parâmetros em `tests/vba/{core,loader,ppt}/Test*.bas`, via `pwsh build/run-tests.ps1` (`-Target
  core|loader|all`, `-Filter "Modulo.Teste|outro"`, `-LoaderSwap`, `-Ppt`, `-TimeoutSec 1200`). Asserts (core/loader): `AssertTrue/False/Equal/
  Almost`, `NewTempWorkbook`, `CloseTemp`, `TempDir`; `.ppam`: `modPptAssert` (`AssertTrue/False/Equal`) e o dublê `TestXlFake`. Testes de
  PowerPoint real só rodam sem nenhum `POWERPNT` (helpers de `tests/vba/core/TestPptSync.bas`: `StartScratchPowerPoint`, `UserWindowAppeared`,
  `StopScratchPowerPoint`, `PasteLinked`, `PasteAgain`, `PptSeriesValue`, `Fixture`, `Teardown`, `AutoBegin`/`AutoEnd`/`AutoReset`: os testes
  novos de PowerPoint real do core entram NESSE módulo).
- **Instâncias do usuário:** ao começar, `Get-CimInstance Win32_Process | Where-Object Name -match '^(EXCEL|POWERPNT|LogonUI)\.exe$'` (nunca
  `tasklist` com vários `/FI`); esses processos nunca são tocados, anexados (`GetObject`) nem encerrados. Com um PowerPoint do usuário aberto os
  testes reais saem verdes registrando "pulado" e o relatório diz quantos (o controlador reagenda). LogonUI presente = sessão bloqueada: testes de
  área de transferência falham; pare e reporte. Builds e testes numa **CÓPIA** (`C:\temp\tqq-pd`, script de sincronização abaixo). Nunca
  `build.ps1 -Install`, `%LOCALAPPDATA%\ThinQQ`, registro real, lojas de certificados. Apague `C:\temp\tqq-pd*` ao terminar (`rm -rf` do bash, caminho
  exato). `git status --short` ao começar e terminar (quarentena: `git checkout -- <arquivo>` e relate).
- **Código:** `Core_*` nunca propaga erro (`Errors_Handle`/`Errors_HandleErr`; `On Error GoTo -1` antes de `Resume Next` em tratador); `MsgBox` só em
  `Core_Notify`/`Core_Confirm`/`RxMsg`/`Loader_Msg`/`PptTqq_Msg`/`PptTqq_Ask` (log-only no Test); formulários só fora do build Test (os testes chamam
  a lógica pura dos módulos `modPptLinks`/`modStyle*`); ganchos de teste só com `BUILD_CONFIG = "Test"`, zerados em todo caminho; ligação tardia
  com o outro aplicativo (`As Object`, constantes numéricas); registro (`GetSetting/SaveSetting`) nunca no build Test. Pastas de origem abertas pelo
  motor: só por `PptSync_OpenBook` (somente leitura, macros desligadas). O motor só usa um PowerPoint **já aberto** (`Ppt_GetApp(False)`).
- **Armadilhas medidas (spec 2026-10-15 §3 e 2026-10-23 §6; não redescubra):** `LinkFormat.Update` com origem inexistente trava o PowerPoint
  (confira antes); `Tags.Item` de chave inexistente devolve "" sem erro; varrer formas de fora do processo custa ~6,4 ms/forma (use índice +
  `FindBySlideID`, ~5 ms); **nunca chame `ChartData.BreakLink`**; `Application.Visible = False` é recusado pelo PowerPoint; objetos COM de outro
  processo não se comparam com `Is` (compare `Shape.Id`); etiquetas no `.pptx` salvo: `Chr(30)`/`Chr(31)` somem, Tab vira espaço, CR vira LF —
  separadores novos só `vbLf` e `|`; `Shapes(nome)` ignora maiúsculas; **o Excel invisível de automação não carrega a faixa** (`onLoad` nunca é
  chamado, PX2): o `customUI` só se valida estaticamente (Test-Ribbon) e numa sessão visível; **um atributo/`idMso` inválido faz o Office descartar a
  aba inteira em silêncio**; hachura e cor de ponto sobrevivem a mudar valores e a `SetSourceData` (PX1), não a recriar o gráfico.
- **Commits:** pequenos, em português, `feat:`/`fix:`/`test:`/`docs:`/`build:`, terminando com a linha `Co-Authored-By:` indicada no prompt do
  implementador. **Não** fazer push, **não** criar nem trocar de branch.
- **Versões:** `version.json` não muda nas tarefas (o controlador/`publicar-release.bat` cuida). O build importa todo `.bas`/`.cls` de
  `src/<alvo>/{modules,classes}`, os formulários de `src/<alvo>/forms` (o alvo `ppt` passa a importar na U3) e todo `Test*.bas` de `tests/vba/<alvo>`.
- **Docs:** cada versão termina com uma tarefa de docs (tutorial `docs/tutorial/tutorial.html` com teto **655.360 B** — meça antes e depois e
  condense o que a versão substituiu; README "Recursos", "Arquitetura", "Verificações manuais" (continua do item 110) e "Limitações"; `CHANGELOG.md`
  `[Não lançado]`; `release/notes.txt` reescrito para a versão; `tests/ps/Test-Docs.ps1` com os rótulos, âncoras e mensagens novas, sem pinos presos
  a uma versão antiga: os blocos do CHANGELOG lidos por `[Não lançado]` com recuo para a versão) e uma tarefa de suíte + revisão final.
- **Processo:** cada tarefa termina com a suíte completa verde na cópia (`pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target all -LoaderSwap -Ppt
  -TimeoutSec 1500` nas tarefas que mexem no PowerPoint; sem `-Ppt` nas demais) e commit(s); nunca duas suítes ao mesmo tempo nem suíte enquanto
  outro agente trabalha no repositório. Modelo sugerido por tarefa: **cheap** (mecânico), **standard**, **most-capable** (motor, PowerPoint,
  formulários, geometria).

## Review Focus

1. **Manual de verdade** (U1): com um gráfico colado (manual) e a pasta salva, mudar B2 e rodar o ciclo deixa o slide **como estava**, grava
   `TQQ_STATE = "new"`, **não** altera a pasta do Excel (`wb.Saved` continua True, o `ChartObject` não é redimensionado, nenhum `LinkFormat.Update`) e
   a apresentação salva continua `Saved`; atualizar tira a bandeira; o dado voltando tira a bandeira. Teste:
   `Test_PptAuto_Tick_FlagsManual_NoExcelChange_KeepsSaved_WhenAvailable`.
2. **Nada nasce automático** (U1): colar, inserir no slide, criar pelo PowerPoint, texto e tabela vinculados nascem sem `TQQ_AUTO` (salvo a opção
   "Novos vínculos nascem automáticos"), entram no `TQQ_LINKIDX`. Teste: `Test_PptInsertChart_FromPpt_WhenAvailable` (ajustado).
3. **Silêncio no sucesso e bandeira na faixa** (U2): `PptTqq_ShowResult` não abre caixa para o resumo de sucesso puro; `PptRx_GetNewVisible`
   segue a etiqueta; nenhuma chamada ao Excel nos callbacks da faixa. Testes: `Test_PptTqq_ShowResult_SilentOnSuccess`,
   `Test_PptTqq_AnyNew_AndCallbacks`.
4. **Vínculos sem efeitos colaterais** (U3): abrir/filtrar/listar não chama o Excel; "Verificar" é a única ação que abre pastas fechadas (somente
   leitura); "Desvincular" tira só etiquetas `TQQ_*`. Testes `Test_PptLinks_*`.
5. **Estilo hostil** (S1/S2): arquivo com chaves desconhecidas, valores fora da lista, nomes com `; = | < > \`, 20 KB, 500 linhas, BOM, CRLF/LF:
   nada é aplicado fora das regras, nada executa. Testes `Test_StyleDef_*`, `Test_StyleStore_Import_*`.
6. **Estilo sobrevive** (S3): Alterar dados, ligar conectores, mudar formato, nova linha na tabela, salvar/reabrir mantêm cores/hachura. Testes
   `Test_StyleApply_*`.
7. **Faixa válida** (todas): Test-Ribbon confere ids únicos, callbacks existentes, `idMso` só da lista permitida (`TabSetChartTools`, `TabInsert` e os
   `imageMso` em uso), nada de `size` em filho de menu.

---

## Estrutura de arquivos (novos e alterados)

```
docs/superpowers/specs/2026-10-23-thinqq-thinkcell-ux-design.md   NOVO (U0)
docs/superpowers/plans/2026-10-23-thinqq-thinkcell-ux.md          NOVO (U0)
--- Parte U (1.11.0) ---
src/core/modules/modPptState.bas   NOVO  estado/bandeira (TQQ_STATE), TQQ_LINKIDX, ManualFor, MarkItems, Lines (U1)
src/core/modules/modPptAuto.bas    ALT   ciclo marca os manuais (U1)
src/core/modules/modPptSync.bas    ALT   CORE_PPT_API 6; OpenBook/FindOpenBook/CloseOpened Public; Items tira a bandeira; Presentation reconstroi LINKIDX (U1)
src/core/modules/modPptTable.bas, modPptText.bas  ALT  LINKIDX na criacao; PPTTABLE_NAME_PREFIX Public (U1)
src/core/modules/modCoreApi.bas    ALT   Core_PptLinkStates; colagem/insercao pelo PowerPoint entram no LINKIDX e nascem manuais (U1)
src/ppt/modules/modPptTqq.bas, modPptRibbon.bas, ribbon/customUI14.xml   ALT (U2, U3)
src/ppt/modules/modPptLinks.bas    NOVO  logica do dialogo Vinculos (U3)
src/ppt/forms/frmTqqLinks.json, .vba     NOVO (U3)
build/build.ps1                    ALT   alvo ppt importa src/ppt/forms (U3)
build/lib/Ribbon-Icons.ps1         ALT   icones tqq_flag, tqq_links (U2/U3)
src/loader/ribbon/customUI14.xml   ALT   menu PowerPoint sem as caixas do automatico (U4)
tests/vba/core/TestPptSync.bas ALT (U1); tests/vba/ppt/TestPptAddin.bas, TestXlFake.cls ALT (U2, U3); tests/ps/Test-Ribbon.ps1 ALT (U2-U4); tests/vba/loader/TestRibbon.bas ALT (U4)
docs/tutorial/tutorial.html, README.md, CHANGELOG.md, release/notes.txt, tests/ps/Test-Docs.ps1   ALT (U5)
--- Parte S (1.12.0) ---
src/core/modules/modStyleDef.bas   NOVO  modelo puro (S1)
src/core/modules/modStyleStore.bas NOVO  depositos, importar/exportar, padrao (S2)
src/core/modules/modStyleApply.bas NOVO  aplicar por papel, carimbo, ler do grafico (S3)
src/core/modules/modCoreApi.bas    ALT   Core_Style*, ganchos de reaplicacao, PptCmd "style" (S2, S3, S5); modChartKinds.bas ALT "style" (S3)
src/core/forms/frmStyles.json, .vba NOVO (S4); src/loader/ribbon/customUI14.xml, modRibbon.bas ALT (S4)
src/core/modules/modPptSync.bas    ALT   PptSync_ApplyFormat, CORE_PPT_API 7 (S5); src/ppt/* ALT menu Estilo (S5)
tests/vba/core/TestStyles.bas NOVO (S1-S3); TestPptSync.bas ALT (S5); tests/vba/loader/TestRibbon.bas ALT (S4); tests/vba/ppt/TestPptAddin.bas ALT (S5)
--- Parte X (1.13.0) ---
src/loader/ribbon/customUI14.xml, modRibbon.bas ALT galeria Vincular ao PowerPoint e guia Inserir (X1)
src/core/modules/modCoreApi.bas, modLinkFrame.bas, modPptAuto.bas ALT moldura com estado, alterar intervalo/tipo mantendo o id (X2, X4)
src/core/modules/modChartMenu.bas NOVO; src/core/classes/clsChartWatcher.cls ALT botao direito/duplo clique (X3)
```

## Como trabalhar (todas as tarefas)

Crie (uma vez por tarefa) `C:\temp\tqq-pd-sync.ps1` e rode `pwsh -File C:\temp\tqq-pd-sync.ps1 -Repo "<repo>"` antes de cada rodada:

```powershell
param([Parameter(Mandatory)][string]$Repo)
$dst = 'C:\temp\tqq-pd'
if ($dst -notlike 'C:\temp\tqq-pd*') { throw 'destino inesperado' }
robocopy $Repo $dst /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS | Out-Null
if ($LASTEXITCODE -ge 8) { throw "robocopy falhou: $LASTEXITCODE" }
'sincronizado'
```

```powershell
pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target core -Filter "TestPptSync" -TimeoutSec 900       # iteracao
pwsh C:\temp\tqq-pd\build\run-tests.ps1 -Target all -LoaderSwap -Ppt -TimeoutSec 1500            # suite completa
pwsh C:\temp\tqq-pd\tests\ps\Test-Ribbon.ps1 -RepoRoot C:\temp\tqq-pd; pwsh C:\temp\tqq-pd\tests\ps\Test-Docs.ps1 -RepoRoot C:\temp\tqq-pd
```
Log da suíte FORA da cópia (pasta scratchpad). Ao terminar (bash): `rm -rf /c/temp/tqq-pd /c/temp/tqq-pd-sync.ps1`. Relatório:
`task-<ID>-report.md` nesta pasta (status, commits, contagens da suíte, testes de PowerPoint pulados, desvios do plano, preocupações).

---

# PARTE U — Atualizar como o think-cell: manual com bandeira, Vínculos e aba contextual (1.11.0)

### Tarefa U0: Promover spec e plano — modelo: cheap

**Files:** Create `docs/superpowers/specs/2026-10-23-thinqq-thinkcell-ux-design.md` e `docs/superpowers/plans/2026-10-23-thinqq-thinkcell-ux.md` (cópias
byte a byte de `design.md` e `plan.md` desta pasta; não sobrescreva se existirem).

- [ ] **Passo 1:** `Copy-Item` dos dois (só se o destino não existir). Confira `Select-String -Path <spec> -Pattern '^## 9\. Matriz de paridade','^### 6\.3 Sondas pendentes'` (2 linhas).
- [ ] **Passo 2:** `pwsh <repo>\tests\ps\Test-Privacy.ps1 -RepoRoot <repo>` → saída 0.
- [ ] **Passo 3:** `git add` dos dois e `git commit -m "docs: spec e plano da UX como o think-cell (1.11.0 a 1.13.0)"` (+ trailer).

---

### Tarefa U1: Estado (bandeira), índice de todos os vínculos, motor marca os manuais, nada nasce automático, API 6 — modelo: most-capable

**Files:**
- Create: `src/core/modules/modPptState.bas`
- Modify: `src/core/modules/modPptAuto.bas` (`PptAuto_Tick`, `PptAuto_CollectManual`, `mLastMarked`, `PptAuto_LastMarkedForTest`, `PptAuto_Cancel`)
- Modify: `src/core/modules/modPptSync.bas` (`CORE_PPT_API = 6`; `PptSync_OpenBook`/`PptSync_FindOpenBook`/`PptSync_CloseOpened` → `Public`; `PptSync_Items` tira a bandeira; `PptSync_Presentation` reconstrói `TQQ_LINKIDX`)
- Modify: `src/core/modules/modPptTable.bas` (`PPTTABLE_NAME_PREFIX` → `Public`; `PptLinkIdx_Add` na criação), `src/core/modules/modPptText.bas` (`PptLinkIdx_Add` ao vincular)
- Modify: `src/core/modules/modCoreApi.bas` (`PptSyncAfterPaste`, `PptInsertChartImpl`, `Core_PptLinkStates`)
- Test: `tests/vba/core/TestPptSync.bas` (acrescentar; ajustar `Test_Core_PptApiVersion` e `Test_PptInsertChart_FromPpt_WhenAvailable`)

**Interfaces:**
- Produces: constantes `PPTSTATE_TAG = "TQQ_STATE"`, `PPTSTATE_TAG_LINKIDX = "TQQ_LINKIDX"`, `PPTSTATE_NEW/OK/UNKNOWN/MISSING/NOSOURCE/UNLINKED`;
  `PptState_Of(shp As Object, wb As Workbook) As String`; `PptState_SetTag(shp As Object, isNew As Boolean) As Boolean`; `PptState_SavedOf(pres As
  Object) As Boolean`; `PptState_RestoreSaved(pres As Object, wasSaved As Boolean)`; `PptLinkIdx_Add(shp As Object)`; `PptLinkIdx_Rebuild(pres As
  Object, items As Collection)`; `PptState_ManualFor(pres As Object, dirty As Object) As Collection`; `PptState_MarkItems(items As Collection) As
  Long`; `PptState_Lines(pres As Object, entries As String, verifyClosed As Boolean) As String`; `PptState_MarksForTest() As Long`;
  `PptState_ResetTestHooks()`; `PptAuto_LastMarkedForTest() As Long`; `Core_PptLinkStates(pres As Object, entries As String, verifyClosed As Boolean)
  As String` (linhas `SlideID|nome|pasta|tipo|modo|estado` por `vbLf`; tipo `chart:<kind>`|`text`|`table`; modo `auto`|`manual`; ou `"!"` +
  mensagem); `CORE_PPT_API = 6`.
- Consumes: `PptSync_Tag`, `PptSync_ItemKind`, `PptSync_ItemBook`, `PptSync_Fingerprint`, `PptSync_EnsureUniqueName`, `PptSync_ResolveEntries`,
  `PptSync_FileExists`, `PptSync_OptValue`, `PptIdx_*`, `PptText_SafeValue`, `PptText_CellText`, `PptText_TagDec`, `PptLink_HasMark`,
  `PptLink_ResolveRange`, `PptTable_Parse`, `PptTable_TextMatrix`, `PptTable_Diff`, `Ppt_IsBusyError`, `Ppt_LinkState`, `Registry_FindChartObject`.

- [ ] **Passo 0: Sonda P-U1 (se não houver `POWERPNT` aberto; senão registre "P-U1 bloqueada" e siga: o código já preserva `Saved`).** Crie
  `C:\temp\tqq-pu1\pu1.ps1` (padrão `p1.ps1` da rodada anterior: conferir `Get-CimInstance` sem POWERPNT, `New-Object -ComObject PowerPoint.Application`,
  PIDs novos registrados, `Presentations.Add(0)`, `SaveAs C:\temp\tqq-pu1\a.pptx`, `AddShape`, `Save`; ler `Saved`; `Shapes(1).Tags.Add 'TQQ_STATE','new'`;
  ler `Saved`; `Saved = -1`; ler `Saved`; `Close`; encerrar pelo PID no `finally`). Anote os três valores no relatório. Apague `C:\temp\tqq-pu1`.

- [ ] **Passo 1: Testes (RED)** — acrescente ao fim de `tests/vba/core/TestPptSync.bas`:

```vb
' ---- 1.11.0 (U1): estado (bandeira TQQ_STATE), indice de todos os vinculos (TQQ_LINKIDX), motor marca os manuais, API 6 ----

Public Sub Test_PptState_SetTag_OnlyWhenChanged_Fake()
    Dim f As TestPptFake, t As TestPptTags
    Set f = New TestPptFake: Set t = New TestPptTags: Set f.Tags = t
    PptState_ResetTestHooks
    AssertFalse PptState_SetTag(f, False), "sem bandeira e sem dados novos: nada gravado"
    AssertTrue PptState_SetTag(f, True)
    AssertEqual PPTSTATE_NEW, t.Item(PPTSTATE_TAG)
    AssertFalse PptState_SetTag(f, True), "ja marcada: nao regrava"
    AssertEqual 1, t.Writes
    AssertTrue PptState_SetTag(f, False)
    AssertEqual "", t.Item(PPTSTATE_TAG)
    AssertEqual 1, t.Deletes
    t.Preset PPTSTATE_TAG, "lixo"
    AssertFalse PptState_SetTag(f, False), "valor estranho (deck alheio) nao e bandeira: nada a apagar"
    t.FailOnKey = PPTSTATE_TAG
    AssertFalse PptState_SetTag(f, True), "falha comum ao gravar: registrada, sem erro"
    AssertEqual 2, PptState_MarksForTest()
End Sub

' Sem indice TQQ_LINKIDX: nenhum manual (nunca varre as formas); com o indice, so as formas da pasta suja, nao automaticas, do ThinQQ.
Public Sub Test_PptState_ManualFor_IndexOnly_Fake()
    Dim shp As TestPptFake, sl As TestPptFake, slides As TestPptFake, pres As TestPptFake, dirty As Object, lk As TestPptFake
    Set shp = New TestPptFake: Set shp.Tags = CreateObject("Scripting.Dictionary")
    shp.Tags.Add PPTSYNC_TAG_CHART, "abc"
    shp.Tags.Add PPTSYNC_TAG_BOOK, "C:\x\a.xlsx"
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add shp, "tqq_m"
    Set slides = New TestPptFake: Set slides.Slide = sl: slides.SlideIdValue = 300
    Set pres = New TestPptFake: Set pres.Slides = slides: Set pres.Tags = CreateObject("Scripting.Dictionary")
    Set dirty = CreateObject("Scripting.Dictionary"): dirty.CompareMode = 1
    dirty.Add "C:\x\a.xlsx", True
    AssertEqual 0, PptState_ManualFor(pres, dirty).Count, "sem indice"
    AssertEqual 0, slides.FindBySlideIDCalls, "sem indice: nenhuma chamada ao PowerPoint"
    pres.Tags.Add PPTSTATE_TAG_LINKIDX, PptIdx_Entry(300, "tqq_m", "C:\x\a.xlsx")
    AssertEqual 1, PptState_ManualFor(pres, dirty).Count
    shp.Tags(PPTSYNC_TAG_AUTO) = "1"
    AssertEqual 0, PptState_ManualFor(pres, dirty).Count, "automatica fica com o ciclo automatico"
    shp.Tags.Remove PPTSYNC_TAG_AUTO
    dirty.RemoveAll: dirty.Add "C:\x\outra.xlsx", True
    AssertEqual 0, PptState_ManualFor(pres, dirty).Count, "pasta que nao mudou"
    AssertEqual 2, slides.FindBySlideIDCalls, "entrada de pasta limpa nao chega ao PowerPoint"
End Sub

' Bandeira (Review Focus 1): colado manual, B2 muda, o ciclo marca TQQ_STATE sem tocar no slide nem no Excel, com a apresentacao salva
' continuando salva; Atualizar tira a bandeira; o dado voltando ao valor do slide tira a bandeira sozinho.
Public Sub Test_PptAuto_Tick_FlagsManual_NoExcelChange_KeepsSaved_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, own As New Collection, en As Long, ed As String
    Dim v0 As Double, pp As String, w0 As Double, msg As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoBegin
    PptState_ResetTestHooks
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\pptstate-flag.xlsx")
    AssertEqual "", PptSync_Tag(s1, PPTSYNC_TAG_AUTO), "nasce manual"
    AssertTrue InStr(1, PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX), CStr(s1.Name), vbBinaryCompare) > 0, "colagem entra no indice de todos os vinculos"
    pp = TempDir() & "\pptstate-flag.pptx"
    pres.SaveAs pp
    AssertTrue CLng(pres.Saved) <> 0, "apresentacao salva"
    v0 = PptSeriesValue(s1.Chart, "Sul", 1)
    w0 = rng.Worksheet.ChartObjects(1).Width
    Watchers_Start
    rng.Worksheet.Range("B2").Value = 77
    AssertEqual 1, PptAuto_PendingCountForTest()
    wb.Saved = True                                     ' qualquer coisa que o ciclo fizer no Excel apareceria como alteracao
    Ppt_SetAppForTest app
    PptSync_ResetTestHooks
    PptAuto_Tick
    AssertEqual PPTSTATE_NEW, PptSync_Tag(s1, PPTSTATE_TAG), "bandeira"
    AssertEqual 1, PptAuto_LastMarkedForTest()
    AssertAlmost v0, PptSeriesValue(s1.Chart, "Sul", 1), 0.0001, "manual: o slide nao mudou"
    AssertEqual 0, PptSync_UpdateCountForTest(), "nenhum LinkFormat.Update"
    AssertTrue wb.Saved, "o ciclo nao alterou a pasta do Excel (o Desfazer do usuario fica)"
    AssertAlmost w0, rng.Worksheet.ChartObjects(1).Width, 0.01, "o grafico do Excel nao foi redimensionado"
    AssertTrue CLng(pres.Saved) <> 0, "a bandeira nao deixa a apresentacao alterada (P-U1)"
    msg = Core_PptSyncShapes(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), True)
    AssertTrue InStr(msg, "Gráficos do ThinQQ atualizados: 1.") = 1, msg
    AssertAlmost 77, PptSeriesValue(s1.Chart, "Sul", 1), 0.0001
    AssertEqual "", PptSync_Tag(s1, PPTSTATE_TAG), "atualizar tira a bandeira"
    rng.Worksheet.Range("B2").Value = 78
    PptAuto_Tick
    AssertEqual PPTSTATE_NEW, PptSync_Tag(s1, PPTSTATE_TAG), "de novo"
    rng.Worksheet.Range("B2").Value = 77
    PptAuto_Tick
    AssertEqual "", PptSync_Tag(s1, PPTSTATE_TAG), "o dado voltou ao valor do slide: a bandeira sai"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set s1 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    If Len(pp) > 0 Then Kill pp
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' API 6: linhas de estado (manual e automatica), bandeira gravada, pasta fechada "unknown" sem verificar e verificada com abertura somente
' leitura (fechada no fim), indice reconstruido com entries = "".
Public Sub Test_Core_PptLinkStates_Lines_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, s2 As Object, own As New Collection, en As Long, ed As String
    Dim lines As String, a() As String, f() As String, p As String, nBooks As Long, i As Long, l1 As String, l2 As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoBegin
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    p = TempDir() & "\pptstate-lines.xlsx"
    Set s1 = PasteLinked(app, pres, wb, rng, p)
    Set s2 = PasteAgain(app, pres)
    PptSync_SetAuto s2, True
    PptSync_TagDelete pres, PPTSTATE_TAG_LINKIDX
    rng.Worksheet.Range("B2").Value = 66
    lines = Core_PptLinkStates(pres, "", False)
    AssertTrue Left$(lines, 1) <> "!", lines
    a = Split(lines, vbLf)
    AssertEqual 2, UBound(a) + 1, lines
    For i = 0 To 1
        f = Split(a(i), "|")
        AssertEqual 5, UBound(f), a(i)
        AssertTrue Left$(f(3), 6) = "chart:", a(i)
        If f(1) = CStr(s1.Name) Then l1 = a(i)
        If f(1) = CStr(s2.Name) Then l2 = a(i)
    Next
    AssertTrue Right$(l1, Len("|manual|new")) = "|manual|new", l1
    AssertTrue Right$(l2, Len("|auto|new")) = "|auto|new", l2
    AssertEqual PPTSTATE_NEW, PptSync_Tag(s1, PPTSTATE_TAG)
    AssertEqual 2, PptIdx_Parse(PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX)).Count, "indice de todos os vinculos reconstruido"
    Core_PptSyncShapes pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), True
    lines = Core_PptLinkStates(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), False)
    AssertTrue Right$(lines, Len("|manual|ok")) = "|manual|ok", lines
    wb.Save
    nBooks = Workbooks.Count
    wb.Close SaveChanges:=False: Set wb = Nothing
    lines = Core_PptLinkStates(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), False)
    AssertTrue Right$(lines, Len("|unknown")) = "|unknown", "pasta fechada sem verificar: " & lines
    lines = Core_PptLinkStates(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), True)
    AssertTrue Right$(lines, Len("|ok")) = "|ok", "verificada (aberta somente leitura): " & lines
    AssertEqual nBooks - 1, Workbooks.Count, "a pasta aberta para verificar foi fechada"
    AssertFalse Core_IsQuiet(), "modo silencioso fechado"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set s1 = Nothing: Set s2 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    If Len(p) > 0 Then Kill p
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

Public Sub Test_Core_PptLinkStates_NeverRaises()
    Dim s As String
    s = Core_PptLinkStates(Nothing, "", False)
    AssertTrue Left$(s, 1) = "!", s
    AssertFalse Core_IsQuiet()
End Sub
```

  Ajustes nos testes existentes do mesmo módulo:
  - `Test_Core_PptApiVersion`: `AssertEqual 6, Core_PptApiVersion(), "contrato com o suplemento do PowerPoint (spec 4.9; 2 = 1.6.0; 3 = 1.8.0; 4 = 1.9.0 comandos; 5 = 1.9.0 texto vinculado; 6 = 1.11.0 estados)"`.
  - `Test_PptInsertChart_FromPpt_WhenAvailable` (linhas 2202, 2207 e 2213–2219 em 3a30f38): troque por

```vb
    AssertEqual "", PptSync_Tag(shp, PPTSYNC_TAG_AUTO), "nasce manual, como o think-cell (1.11.0)"
    AssertEqual 1, Ppt_LinkState(shp)
    AssertEqual LCase$(p), LCase$(PptSync_ItemBook(shp))
    AssertEqual "Gráfico 1", PptSync_Tag(shp, PPTSYNC_TAG_SHEET)
    AssertEqual LCase$(p), LCase$(PptSync_Tag(pres, PPTDATA_MARK))
    AssertTrue InStr(1, PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX), CStr(shp.Name), vbBinaryCompare) > 0, "no indice de todos os vinculos"
    AssertEqual "", PptSync_Tag(pres, PPTSYNC_TAG_INDEX), "nenhum automatico"
    Set wb = Workbooks(Dir(p))
    AssertTrue PptData_IsDataBook(wb), "pasta de dados marcada"
    AssertTrue wb.Saved, "pasta de dados salva"
    Set ws = wb.Worksheets("Gráfico 1")
    AssertTrue Not LinkFrame_Find(ws.ChartObjects(1).Chart) Is Nothing, "dados destacados pela moldura"
    Watchers_Start
    ws.Range("B2").Value = 999
    AssertTrue PptAuto_PendingCountForTest() >= 1
    Ppt_SetAppForTest app
    PptAuto_Tick
    AssertEqual PPTSTATE_NEW, PptSync_Tag(shp, PPTSTATE_TAG), "bandeira: o slide nao acompanha sozinho"
    AssertTrue wb.Saved, "o ciclo salvou a pasta de dados (G3)"
    AssertTrue InStr(Core_PptSyncShapes(pres, PptIdx_Entry(CLng(slide.SlideID), CStr(shp.Name), ""), True), "Gráficos do ThinQQ atualizados: 1.") = 1
    AssertAlmost 999, PptSeriesValue(shp.Chart, "Sul", 1), 0.0001, "Atualizar traz o dado"
```
  (mantendo o resto do teste como está).

- [ ] **Passo 2:** `-Target core -Filter "TestPptSync"` → saída 3 (compilação: `PptState_SetTag` não existe). Anote como RED.

- [ ] **Passo 3: `modPptState`** — crie `src/core/modules/modPptState.bas`:

```vb
Attribute VB_Name = "modPptState"
Option Explicit
' Estado dos elementos do ThinQQ no PowerPoint (1.11.0; spec 2026-10-23 secao 8.1). Como no think-cell, o vinculo e MANUAL por padrao e o slide
' mostra quando a pasta tem dados novos (a bandeira). O core compara, SEM ALTERAR NADA no Excel (o Desfazer do usuario fica: nada de
' redimensionar, de Core_SafeRefresh nem de salvar), o que o slide mostra (TQQ_FP do grafico, TQQ_TLAST do texto, TQQ_TBLAST da tabela) com o
' que a pasta aberta daria agora, e grava na forma TQQ_STATE = "new" ou a apaga. TQQ_LINKIDX (etiqueta da apresentacao, formato PptIdx_*)
' lista TODOS os elementos vinculados; TQQ_AUTOIDX continua so com os automaticos. Gravar estas etiquetas preserva Presentation.Saved (sao
' estado da sessao, nao alteracao do usuario).
Public Const PPTSTATE_TAG As String = "TQQ_STATE"
Public Const PPTSTATE_TAG_LINKIDX As String = "TQQ_LINKIDX"
Public Const PPTSTATE_NEW As String = "new"              ' dados novos na pasta (a bandeira)
Public Const PPTSTATE_OK As String = "ok"                ' o slide mostra o que a pasta tem
Public Const PPTSTATE_UNKNOWN As String = "unknown"      ' sem referencia gravada (TQQ_FP ausente) ou pasta fechada nao verificada
Public Const PPTSTATE_MISSING As String = "missing"      ' pasta de origem nao encontrada
Public Const PPTSTATE_NOSOURCE As String = "nosource"    ' grafico, celula ou intervalo nao encontrado na pasta
Public Const PPTSTATE_UNLINKED As String = "unlinked"    ' grafico sem o vinculo nativo
Private mTestMarks As Long                               ' gancho de teste: etiquetas TQQ_STATE gravadas/apagadas por PptState_SetTag

Public Function PptState_MarksForTest() As Long
    PptState_MarksForTest = mTestMarks
End Function

Public Sub PptState_ResetTestHooks()
    mTestMarks = 0
End Sub

' Etiqueta lida sem engolir o PowerPoint ocupado (sobe); outro erro ou etiqueta ausente = "".
Private Function PptState_TagStrict(o As Object, key As String) As String
    Dim v As Variant, en As Long, ed As String
    On Error Resume Next
    v = o.Tags.Item(key)
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptState_TagStrict", ed
    If en = 0 And Not IsNull(v) And Not IsObject(v) Then PptState_TagStrict = CStr(v)
End Function

' Tipo do elemento (a regra de PptSync_ItemKind) sem engolir o ocupado.
Private Function PptState_KindStrict(shp As Object) As String
    If Len(PptState_TagStrict(shp, PPTSYNC_TAG_CHART)) > 0 Then PptState_KindStrict = "chart": Exit Function
    If PptState_TagStrict(shp, PPTTEXT_TAG) = "1" Then PptState_KindStrict = "text": Exit Function
    If PptState_TagStrict(shp, PPTTABLE_TAG) = "1" Then PptState_KindStrict = "table"
End Function

' Intervalo de um texto/tabela vinculado na pasta wb com as MESMAS regras de seguranca de PptText_Sync/PptTable_Sync (revisao final E, I1): so
' um nome com o prefixo dado (criado pela acao de vincular) e seguido; o endereco de reserva so numa pasta com a marca PPTLINK_MARK. Nothing se
' nao acha. PowerPoint ocupado sobe.
Private Function PptState_LinkRange(shp As Object, wb As Workbook, prefix As String) As Range
    Dim nm As String, ref As String
    nm = PptState_TagStrict(shp, PPTLINK_TAG_NAME)
    If Left$(nm, Len(prefix)) <> prefix Then nm = ""
    If PptLink_HasMark(wb) Then ref = PptState_TagStrict(shp, PPTLINK_TAG_REF)
    If Len(nm) = 0 And Len(ref) = 0 Then Exit Function
    On Error Resume Next
    Set PptState_LinkRange = PptLink_ResolveRange(wb, nm, ref)
    Err.Clear
End Function

' Estado do elemento pela pasta wb (aberta). So LE: nada e gravado no Excel nem no slide. PowerPoint ocupado sobe; outro erro = UNKNOWN (registrado).
Public Function PptState_Of(shp As Object, wb As Workbook) As String
    Dim co As ChartObject, fp As String, rng As Range, v As String, cut As Boolean, oldM As Variant, newM As Variant, en As Long, ed As String
    On Error GoTo Fail
    Select Case PptState_KindStrict(shp)
        Case "chart"
            Set co = Registry_FindChartObject(wb, PptState_TagStrict(shp, PPTSYNC_TAG_CHART))
            If co Is Nothing Then PptState_Of = PPTSTATE_NOSOURCE: Exit Function
            fp = PptState_TagStrict(shp, PPTSYNC_TAG_FP)
            If Len(fp) = 0 Then PptState_Of = PPTSTATE_UNKNOWN: Exit Function
            If PptSync_Fingerprint(co, shp) = fp Then PptState_Of = PPTSTATE_OK Else PptState_Of = PPTSTATE_NEW
        Case "text"
            Set rng = PptState_LinkRange(shp, wb, "TQQ_X_")
            If rng Is Nothing Then PptState_Of = PPTSTATE_NOSOURCE: Exit Function
            v = PptText_SafeValue(PptText_CellText(rng), cut)
            If v = PptText_TagDec(PptState_TagStrict(shp, PPTTEXT_TAG_LAST)) Then PptState_Of = PPTSTATE_OK Else PptState_Of = PPTSTATE_NEW
        Case "table"
            Set rng = PptState_LinkRange(shp, wb, PPTTABLE_NAME_PREFIX)
            If rng Is Nothing Then PptState_Of = PPTSTATE_NOSOURCE: Exit Function
            oldM = PptTable_Parse(PptState_TagStrict(shp, PPTTABLE_TAG_LAST))
            If IsEmpty(oldM) Then PptState_Of = PPTSTATE_UNKNOWN: Exit Function
            newM = PptTable_TextMatrix(rng)
            If UBound(oldM, 1) <> UBound(newM, 1) Or UBound(oldM, 2) <> UBound(newM, 2) Then PptState_Of = PPTSTATE_NEW: Exit Function
            If PptTable_Diff(oldM, newM).Count = 0 Then PptState_Of = PPTSTATE_OK Else PptState_Of = PPTSTATE_NEW
        Case Else
            PptState_Of = PPTSTATE_UNKNOWN
    End Select
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error GoTo 0                                     ' sem isto o GoTo -1 rearma o Fail e o Raise voltaria para ca (laco)
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptState_Of", ed
    Errors_Log "PptState_Of" & vbTab & en & vbTab & ed
    PptState_Of = PPTSTATE_UNKNOWN
End Function

' Grava (isNew) ou apaga TQQ_STATE so quando muda (um valor diferente de "new", vindo de um deck alheio, nao e bandeira e fica). True se gravou
' ou apagou. PowerPoint ocupado sobe; outro erro e registrado (False).
Public Function PptState_SetTag(shp As Object, isNew As Boolean) As Boolean
    Dim cur As String, en As Long, ed As String
    cur = PptState_TagStrict(shp, PPTSTATE_TAG)
    If isNew = (cur = PPTSTATE_NEW) Then Exit Function
    If Not isNew Then
        If cur <> PPTSTATE_NEW Then Exit Function
    End If
    On Error Resume Next
    If isNew Then shp.Tags.Add PPTSTATE_TAG, PPTSTATE_NEW Else shp.Tags.Delete PPTSTATE_TAG
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptState_SetTag", ed
    If en <> 0 Then Errors_Log "PptState_SetTag" & vbTab & en & vbTab & ed: Exit Function
    PptState_SetTag = True
    mTestMarks = mTestMarks + 1
End Function

' Presentation.Saved antes de gravar etiquetas de estado (False se ilegivel). Nunca propaga erro.
Public Function PptState_SavedOf(pres As Object) As Boolean
    On Error Resume Next
    PptState_SavedOf = (CLng(pres.Saved) <> 0)
    Err.Clear
End Function

' Devolve Saved = True se a apresentacao estava salva antes (P-U1). Nunca propaga erro.
Public Sub PptState_RestoreSaved(pres As Object, wasSaved As Boolean)
    On Error Resume Next
    If wasSaved Then
        If CLng(pres.Saved) = 0 Then pres.Saved = -1
    End If
    Err.Clear
End Sub

' Acrescenta a forma (elemento do ThinQQ recem-criado, colado ou vinculado) ao TQQ_LINKIDX da apresentacao, com o nome unico no slide (o indice
' acha a forma pelo nome). Preserva Saved. Nunca propaga erro (falha registrada: a sincronizacao completa e a abertura no .ppam reconstroem).
Public Sub PptLinkIdx_Add(shp As Object)
    Dim sl As Object, pres As Object, idx As String, was As Boolean
    On Error GoTo Fail
    Set sl = shp.Parent: Set pres = sl.Parent
    was = PptState_SavedOf(pres)
    PptSync_EnsureUniqueName shp
    idx = PptIdx_Set(PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX), CLng(sl.SlideID), CStr(shp.Name), PptSync_ItemBook(shp), True)
    pres.Tags.Add PPTSTATE_TAG_LINKIDX, idx
    PptState_RestoreSaved pres, was
    Exit Sub
Fail:
    Errors_Log "PptLinkIdx_Add" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Reconstroi TQQ_LINKIDX com todos os itens dados (todas as formas do ThinQQ da apresentacao); so grava se mudou; preserva Saved. Cada forma
' isolada (a que falha fica fora e e registrada). Nunca propaga erro.
Public Sub PptLinkIdx_Rebuild(pres As Object, items As Collection)
    Dim shp As Object, out As New Collection, s As String, was As Boolean
    On Error GoTo Done
    For Each shp In items
        On Error Resume Next
        PptSync_EnsureUniqueName shp
        out.Add Array(CLng(shp.Parent.SlideID), CStr(shp.Name), PptSync_ItemBook(shp))
        If Err.Number <> 0 Then Errors_Log "PptLinkIdx_Rebuild: forma fora do indice" & vbTab & Err.Number & vbTab & Err.Description
        Err.Clear
        On Error GoTo Done
    Next
    s = PptIdx_Join(out)
    If s = PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX) Then Exit Sub
    was = PptState_SavedOf(pres)
    If Len(s) = 0 Then PptSync_TagDelete pres, PPTSTATE_TAG_LINKIDX Else pres.Tags.Add PPTSTATE_TAG_LINKIDX, s
    PptState_RestoreSaved pres, was
    Exit Sub
Done:
    Errors_Log "PptLinkIdx_Rebuild" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Elementos MANUAIS (sem TQQ_AUTO = "1") de pres cuja pasta de origem esta em dirty, pelo TQQ_LINKIDX (nunca varre as formas; sem o indice:
' vazia). A pasta da entrada pre-filtra antes de qualquer chamada ao PowerPoint; a pasta REAL da forma tambem tem de estar em dirty. Entradas
' velhas sao ignoradas (a sincronizacao completa e a abertura reconstroem o indice). PowerPoint ocupado sobe.
Public Function PptState_ManualFor(pres As Object, dirty As Object) As Collection
    Dim out As New Collection, e As Variant, sl As Object, shp As Object, en As Long, seen As Object, k As String, idx As String
    Set PptState_ManualFor = out
    idx = PptState_TagStrict(pres, PPTSTATE_TAG_LINKIDX)
    If Len(idx) = 0 Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For Each e In PptIdx_Parse(idx)
        If dirty.Exists(CStr(e(2))) Then
            Set sl = Nothing: Set shp = Nothing
            On Error Resume Next
            Set sl = pres.Slides.FindBySlideID(CLng(e(0)))
            If Not sl Is Nothing Then Set shp = sl.Shapes(CStr(e(1)))
            en = Err.Number
            Err.Clear
            On Error GoTo 0
            If Ppt_IsBusyError(en) Then Err.Raise en, "PptState_ManualFor", "PowerPoint ocupado"
            If Not shp Is Nothing Then
                If PptState_TagStrict(shp, PPTSYNC_TAG_AUTO) <> "1" Then
                    If Len(PptState_KindStrict(shp)) > 0 Then
                        k = CStr(e(0)) & "|" & LCase$(CStr(e(1)))
                        If Not seen.Exists(k) Then
                            seen.Add k, True
                            If dirty.Exists(PptSync_ItemBook(shp)) Then out.Add shp
                        End If
                    End If
                End If
            End If
        End If
    Next
End Function

Private Function PptState_PresKey(pres As Object) As String
    On Error Resume Next
    PptState_PresKey = LCase$(CStr(pres.FullName))
    If Len(PptState_PresKey) = 0 Then PptState_PresKey = "?" & CStr(ObjPtr(pres))
    Err.Clear
End Function

Private Sub PptState_RestoreAll(was As Object)
    Dim k As Variant, a As Variant
    On Error Resume Next
    For Each k In was.Keys
        a = was(k)
        PptState_RestoreSaved a(0), CBool(a(1))
    Next
    Err.Clear
End Sub

' Marca (TQQ_STATE) os elementos dados pela pasta ABERTA neste Excel (nunca abre pasta): "new" se o slide esta atras, apagada se esta em dia;
' estado desconhecido ou sem origem nao mexe na etiqueta. Saved de cada apresentacao preservado. Devolve quantos estao com dados novos.
' PowerPoint ocupado (ou falha estrutural) sobe depois de devolver o Saved.
Public Function PptState_MarkItems(items As Collection) As Long
    Dim shp As Object, wb As Workbook, st As String, pres As Object, was As Object, k As String, en As Long, ed As String
    Set was = CreateObject("Scripting.Dictionary")
    On Error GoTo Fail
    For Each shp In items
        Set wb = PptSync_FindOpenBook(PptSync_ItemBook(shp))
        If Not wb Is Nothing Then
            st = PptState_Of(shp, wb)
            If st = PPTSTATE_NEW Or st = PPTSTATE_OK Then
                Set pres = shp.Parent.Parent
                k = PptState_PresKey(pres)
                If Not was.Exists(k) Then was.Add k, Array(pres, PptState_SavedOf(pres))
                PptState_SetTag shp, (st = PPTSTATE_NEW)
                If st = PPTSTATE_NEW Then PptState_MarkItems = PptState_MarkItems + 1
            End If
        End If
    Next
    PptState_RestoreAll was
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptState_RestoreAll was
    On Error GoTo 0
    Err.Raise en, "PptState_MarkItems", ed
End Function

' Estado de cada elemento (entries no formato do indice; "" = todos os elementos da apresentacao, por varredura, e o TQQ_LINKIDX e
' reconstruido). Linha: SlideID|nome|pasta|tipo|modo|estado (campos sem os separadores, PptIdx_SafeField; linhas por vbLf); tipo = "chart:" +
' TQQ_KIND, "text" ou "table"; modo = auto|manual; estado = PPTSTATE_*. Pasta fechada: com verifyClosed, aberta SOMENTE LEITURA com as macros
' desligadas (PptSync_OpenBook) e fechada no fim; sem, "unknown" (o arquivo existe) ou "missing". Atualiza TQQ_STATE (Saved preservado).
' PowerPoint ocupado sobe.
Public Function PptState_Lines(pres As Object, entries As String, verifyClosed As Boolean) As String
    Dim items As Collection, sl As Object, shp As Object, opened As New Collection, books As Object, out() As String, n As Long
    Dim path As String, wb As Workbook, st As String, kind As String, mode As String, was As Boolean, en As Long, ed As String
    On Error GoTo Fail
    was = PptState_SavedOf(pres)
    If Len(entries) = 0 Then
        Set items = New Collection
        For Each sl In pres.Slides
            For Each shp In sl.Shapes
                If Len(PptSync_ItemKind(shp)) > 0 Then items.Add shp
            Next
        Next
        PptLinkIdx_Rebuild pres, items
    Else
        Set items = PptSync_ResolveEntries(pres, entries)
    End If
    If items.Count > 0 Then
        Set books = CreateObject("Scripting.Dictionary"): books.CompareMode = 1
        ReDim out(1 To items.Count)
        For Each shp In items
            kind = PptSync_ItemKind(shp)
            If PptSync_Tag(shp, PPTSYNC_TAG_AUTO) = "1" Then mode = "auto" Else mode = "manual"
            path = PptSync_ItemBook(shp)
            st = ""
            If kind = "chart" Then
                If Ppt_LinkState(shp) = 0 Then st = PPTSTATE_UNLINKED
            End If
            If Len(st) = 0 Then
                If Not books.Exists(path) Then
                    Set wb = PptSync_FindOpenBook(path)
                    If wb Is Nothing And verifyClosed And Len(path) > 0 Then
                        If PptSync_FileExists(path) Then Set wb = PptSync_OpenBook(path, opened)
                    End If
                    books.Add path, wb
                End If
                Set wb = books(path)
                If wb Is Nothing Then
                    If Len(path) > 0 And PptSync_FileExists(path) Then st = PPTSTATE_UNKNOWN Else st = PPTSTATE_MISSING
                Else
                    st = PptState_Of(shp, wb)
                End If
            End If
            If st = PPTSTATE_NEW Or st = PPTSTATE_OK Then PptState_SetTag shp, (st = PPTSTATE_NEW)
            If kind = "chart" Then kind = "chart:" & PptSync_OptValue(PptSync_Tag(shp, PPTSYNC_TAG_KIND))
            n = n + 1
            out(n) = CStr(shp.Parent.SlideID) & "|" & PptIdx_SafeField(CStr(shp.Name)) & "|" & PptIdx_SafeField(path) & "|" & kind & "|" & mode & "|" & st
        Next
        PptState_Lines = Join(out, vbLf)
    End If
    PptSync_CloseOpened opened
    PptState_RestoreSaved pres, was
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptSync_CloseOpened opened
    PptState_RestoreSaved pres, was
    On Error GoTo 0
    Err.Raise en, "PptState_Lines", ed
End Function
```

- [ ] **Passo 4: `modPptSync`** — (a) `Public Const CORE_PPT_API As Long = 6` com o comentário acrescido de `; 6 = Core_PptLinkStates (1.11.0)`; (b) troque
  `Private Function PptSync_OpenBook`, `Private Function PptSync_FindOpenBook` e `Private Sub PptSync_CloseOpened` por `Public` (mesmas
  assinaturas); (c) em `PptSync_Items`, no `Select Case r`, troque as duas primeiras linhas por

```vb
                        Case PPTSYNC_R_OK: nOk = nOk + 1: PptState_SetTag shp, False          ' 1.11.0: atualizado = sem bandeira
                        Case PPTSYNC_R_SAME: nSame = nSame + 1: PptState_SetTag shp, False
```
  (d) em `PptSync_Presentation`, no ramo `items.Count = 0` acrescente `If Len(PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX)) > 0 Then PptSync_TagDelete pres, PPTSTATE_TAG_LINKIDX`
  e, depois de `PptSync_RebuildIndex pres, items`, `PptLinkIdx_Rebuild pres, items`.

- [ ] **Passo 5: criação entra no índice e nasce manual** —
  - `modCoreApi.bas`, fim de `PptSyncAfterPaste` (depois do `End If` do `PptAuto_NewLinksAuto`): `PptLinkIdx_Add shp                              ' 1.11.0: indice de todos os vinculos (nunca propaga erro)`.
  - `modCoreApi.bas`, `PptInsertChartImpl`: troque `PptSync_SetAuto shp, True                        ' por ultimo: ...` por
    `PptLinkIdx_Add shp                               ' 1.11.0: nasce MANUAL (como o think-cell); a colagem ja ligou o automatico se "Novos vinculos" estiver marcada`
    e atualize o comentário do cabeçalho de `Core_PptInsertChart` ("marca TQQ_DATA e liga 'Atualizar automaticamente'" → "marca TQQ_DATA; nasce manual (1.11.0)").
  - `modPptText.bas`, em `PptText_Link`, depois de `If wasAuto Or PptAuto_NewLinksAuto() Then PptSync_SetAuto shp, True`: `PptLinkIdx_Add shp`.
  - `modPptTable.bas`: `Public Const PPTTABLE_NAME_PREFIX As String = "TQQ_T_"` (era `Private`); em `PptTable_Insert`, depois de
    `If PptAuto_NewLinksAuto() Then PptSync_SetAuto shp, True`: `PptLinkIdx_Add shp`.

- [ ] **Passo 6: motor** — em `src/core/modules/modPptAuto.bas`:
  - declare `Private mLastMarked As Long                                ' elementos manuais com bandeira no ultimo ciclo (lido pelos testes)` e acrescente

```vb
Public Function PptAuto_LastMarkedForTest() As Long
    PptAuto_LastMarkedForTest = mLastMarked
End Function

' Elementos manuais (1.11.0) de uma apresentacao editavel, para a bandeira. Erro comum: registrado e a apresentacao pulada; ocupado sobe.
Private Sub PptAuto_CollectManual(pres As Object, out As Collection)
    Dim items As Collection, shp As Object, en As Long, ed As String
    On Error GoTo Fail
    If CLng(pres.ReadOnly) <> 0 Then Exit Sub
    Set items = PptState_ManualFor(pres, mDirty)
    For Each shp In items
        out.Add shp
    Next
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error GoTo 0
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptAuto_CollectManual", ed
    PptAuto_Log "PptAuto: apresentacao pulada (bandeiras)" & vbTab & PptAuto_PresName(pres) & vbTab & en & vbTab & ed
End Sub
```
  - em `PptAuto_Cancel` acrescente `mLastMarked = 0`; em `PptAuto_Tick` declare `manual As Collection` e troque o bloco `If Not app Is Nothing Then ... End If` por

```vb
    If Not app Is Nothing Then
        Set work = New Collection
        For Each pres In app.Presentations
            PptAuto_Collect pres, work
        Next
        Set pres = Nothing
        If work.Count > 0 Then
            If PptAuto_PptEditingText(app) Then Err.Raise PPTAUTO_ERR_BUSY, "PptAuto_Tick", "texto em edicao no PowerPoint"
            PptAuto_StatusSave
            For Each w In work
                PptAuto_SyncWork w(0), w(1)
            Next
            PptAuto_StatusRestore
        End If
        ' 1.11.0: os MANUAIS da pasta que mudou ganham (ou perdem) a bandeira TQQ_STATE, sem nada mudar no Excel nem no slide
        Set manual = New Collection
        For Each pres In app.Presentations
            PptAuto_CollectManual pres, manual
        Next
        Set pres = Nothing
        mLastMarked = 0
        If manual.Count > 0 Then mLastMarked = PptState_MarkItems(manual)
    End If
```
  e acrescente `Set manual = Nothing` nas duas linhas que soltam `work`/`app` (fim normal e `Fail:`). Atualize o comentário de `PptAuto_Tick`
  ("...sincroniza as formas automaticas...; desde a 1.11.0 tambem marca a bandeira dos manuais (PptState_MarkItems), sem alterar o Excel").

- [ ] **Passo 7: API** — em `modCoreApi.bas`, depois de `Core_PptSyncShapes`:

```vb
' API 6 (1.11.0): estados dos elementos para o .ppam (dialogo Vinculos e abertura da apresentacao; PptState_Lines). Sem caixas (modo
' silencioso) e com o ciclo automatico esperando, equilibrados em todo caminho. Devolve as linhas, ou "!" + mensagem. Nunca propaga erro.
Public Function Core_PptLinkStates(pres As Object, entries As String, verifyClosed As Boolean) As String
    Dim en As Long, ed As String
    On Error GoTo Fail
    Core_QuietBegin
    Ppt_BeginOp
    Core_PptLinkStates = PptState_Lines(pres, entries, verifyClosed)
    Ppt_EndOp
    Core_QuietEnd
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Ppt_EndOp
    Core_QuietEnd
    Errors_Log "Core_PptLinkStates" & vbTab & en & vbTab & ed
    If Ppt_IsBusyError(en) Then Core_PptLinkStates = "!" & PPT_ERR_BUSY Else Core_PptLinkStates = "!Não foi possível verificar os vínculos: " & ed & " (" & en & ")"
End Function
```

- [ ] **Passo 8:** sincronize e rode `-Target core -Filter "TestPptSync|TestPptText|TestPptTable|TestPptData|TestCoreApi"` (sem POWERPNT do usuário:
  os testes `_WhenAvailable` rodam; com: anote quantos pularam) → verde. Depois a suíte completa `-Target all -LoaderSwap -Ppt -TimeoutSec 1500` →
  saída 0 (os testes do `.ppam` ainda não mudam).
- [ ] **Passo 9:** commits: `feat: estado dos vinculos (bandeira TQQ_STATE) e indice TQQ_LINKIDX; tudo nasce manual (API 6)` e `test: bandeira, estados e criacao manual`.

---

### Tarefa U2: `.ppam` — bandeira na faixa, todos os tipos no "Elemento selecionado", silêncio no sucesso, abertura sem pergunta, aba contextual — modelo: most-capable

**Files:**
- Modify: `src/ppt/modules/modPptTqq.bas`, `src/ppt/modules/modPptRibbon.bas`, `src/ppt/ribbon/customUI14.xml`, `build/lib/Ribbon-Icons.ps1`
- Test: `tests/vba/ppt/TestPptAddin.bas`, `tests/vba/ppt/TestXlFake.cls`, `tests/ps/Test-Ribbon.ps1`, `tests/ps/Test-Docs.ps1` (só os pinos que esta tarefa quebra)

**Interfaces:**
- Produces: `PPT_TAG_STATE = "TQQ_STATE"`, `PPT_TAG_LINKIDX = "TQQ_LINKIDX"`, `PPT_API_STATES = 6`, `PPT_DEFAULT_OPENMODE = "never"`, `PPT_SUM_PREFIX`,
  `PPT_MSG_OLD_CORE_STATES`, `PPT_MSG_NOCORE_STATES`; `PptTqq_IsLinkedShape(shp) As Boolean`; `PptTqq_SelectedLinked() As Collection`;
  `PptTqq_AnyNew(shapes As Collection) As Boolean`; `PptTqq_IsPlainSuccess(msg As String) As Boolean`; `PptTqq_ShowResult(msg As String)`;
  `PptTqq_LastResult() As String`; `PptTqq_SyncEntriesMsg(pres, entries, force, Optional fromAuto) As String`; `PptTqq_PrepareOnOpen(pres) As Long`;
  `PptTqq_RebuildLinkIdx(pres As Object) As Long`; `PptTqq_LinkStates(pres, entries, verifyClosed, Optional allowCreate) As String`;
  `PptTqq_Ask` passa a `Public`; callbacks `PptRx_GetLinkedEnabled`, `PptRx_GetNewVisible`, `PptRx_GetUpdVisible`.
- Consumes: `Core_PptLinkStates` (API 6), `Core_PptSyncShapes` (API 2).

- [ ] **Passo 0: Sonda P-U2** (sem POWERPNT do usuário): no próprio teste `Test_PptTqq_AnyNew_AndCallbacks` (abaixo) meça com `Timer` 200 leituras de
  `PptTqq_AnyNew` sobre um slide com 200 formas e registre no relatório (meta < 5 ms por leitura; se passar, troque `PptTqq_AnyNew` para ler só a
  1ª forma da seleção e anote).

- [ ] **Passo 1: Testes (RED)** — em `tests/vba/ppt/TestXlFake.cls` acrescente `Public StatesResult As String                  ' resposta de Core_PptLinkStates ("" = "1|a|b|chart:stacked|manual|ok")`
  e `Public LastVerify As Variant` e, no `Select Case nm` de `Run`:

```vb
        Case "Core_PptLinkStates"
            LastEntries = CStr(a2)
            LastVerify = a3
            If Len(StatesResult) > 0 Then Run = StatesResult Else Run = "1|a|b|chart:stacked|manual|ok"
```
  Acrescente ao fim de `tests/vba/ppt/TestPptAddin.bas`:

```vb
' ---- 1.11.0 (U2): bandeira, todos os tipos, silencio no sucesso, abertura ----

Public Sub Test_PptTqq_IsPlainSuccess()
    AssertTrue PptTqq_IsPlainSuccess("Gráficos do ThinQQ atualizados: 1.")
    AssertTrue PptTqq_IsPlainSuccess("Gráficos do ThinQQ atualizados: 0. Sem mudança: 2.")
    AssertFalse PptTqq_IsPlainSuccess("Gráficos do ThinQQ atualizados: 0." & vbLf & "Pasta de origem não encontrada ou sem acesso: a.xlsx.")
    AssertFalse PptTqq_IsPlainSuccess(PPT_MSG_BUSY)
    AssertFalse PptTqq_IsPlainSuccess("")
End Sub

Public Sub Test_PptTqq_ShowResult_SilentOnSuccess()
    PptTqq_ClearLastMsg
    PptTqq_ShowResult "Gráficos do ThinQQ atualizados: 1."
    AssertEqual "", PptTqq_LastMsg(), "sucesso puro: sem caixa (como o think-cell)"
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", PptTqq_LastResult()
    PptTqq_ShowResult "Gráficos do ThinQQ atualizados: 0." & vbLf & "Falha ao atualizar: 1 (veja o log do ThinQQ)."
    AssertTrue InStr(PptTqq_LastMsg(), "Falha ao atualizar") > 0, "problema: caixa"
    PptTqq_ShowResult "!" & PPT_MSG_BUSY
    AssertEqual PPT_MSG_BUSY, PptTqq_LastMsg(), "'!' = aviso de exclamacao, sem o '!'"
End Sub

Public Sub Test_PptTqq_LinkedShapes_AndAnyNew()
    Dim p As Object, sl As Object, a As Object, b As Object, c As Object, d As Object, col As New Collection, t As Single, i As Long
    Set p = NewPres()
    On Error GoTo Cleanup
    Set sl = p.Slides(1)
    Set a = sl.Shapes.AddShape(1, 10, 10, 50, 50): a.Tags.Add PPT_TAG_CHART, "abc"
    Set b = sl.Shapes.AddShape(1, 70, 10, 50, 50): b.Tags.Add PPT_TAG_TEXT, "1"
    Set c = sl.Shapes.AddShape(1, 130, 10, 50, 50): c.Tags.Add PPT_TAG_TABLE, "1"
    Set d = sl.Shapes.AddShape(1, 190, 10, 50, 50)
    AssertTrue PptTqq_IsLinkedShape(a) And PptTqq_IsLinkedShape(b) And PptTqq_IsLinkedShape(c)
    AssertFalse PptTqq_IsLinkedShape(d)
    col.Add a: col.Add b
    AssertFalse PptTqq_AnyNew(col)
    b.Tags.Add PPT_TAG_STATE, "new"
    AssertTrue PptTqq_AnyNew(col), "texto vinculado com dados novos"
    b.Tags.Add PPT_TAG_STATE, "x"
    AssertFalse PptTqq_AnyNew(col), "so 'new' e bandeira"
    AssertEqual 0, PptTqq_SelectedLinked().Count, "sem janela: nada selecionado"
    For i = 1 To 196
        sl.Shapes.AddShape 1, 1, 1, 2, 2
    Next
    t = Timer
    For i = 1 To 200
        PptTqq_AnyNew col
    Next
    PptTqq_Log "P-U2: AnyNew x200 = " & Format$((Timer - t) * 1000, "0") & " ms"
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Abertura (1.11.0): bandeiras gravadas no arquivo saem, TQQ_LINKIDX e reconstruido com os elementos vinculados (formato do core), Saved fica.
Public Sub Test_PptTqq_PrepareOnOpen_ClearsFlags_BuildsIndex_KeepsSaved()
    Dim p As Object, a As Object, b As Object, c As Object, idx As String
    Set p = NewPres()
    On Error GoTo Cleanup
    Set a = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_a": a.Tags.Add PPT_TAG_CHART, "abc": a.Tags.Add PPT_TAG_STATE, "new"
    a.Tags.Add PPT_TAG_BOOK, "C:\x\a.xlsx"
    Set b = p.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50): b.Name = "txt_b": b.Tags.Add PPT_TAG_TEXT, "1": b.Tags.Add PPT_TAG_BOOK, "C:\x\b.xlsx"
    Set c = p.Slides(1).Shapes.AddShape(1, 130, 10, 50, 50)
    p.Saved = -1
    AssertEqual 2, PptTqq_PrepareOnOpen(p)
    AssertEqual "", PptTqq_Tag(a, PPT_TAG_STATE), "bandeira do arquivo sai"
    idx = PptTqq_Tag(p, PPT_TAG_LINKIDX)
    AssertEqual PptTqq_IdxEntry(CLng(p.Slides(1).SlideID), "tqq_a", "C:\x\a.xlsx") & vbLf & PptTqq_IdxEntry(CLng(p.Slides(1).SlideID), "txt_b", "C:\x\b.xlsx"), idx
    AssertTrue CLng(p.Saved) <> 0, "Saved preservado"
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

Public Sub Test_PptTqq_LinkStates_ApiAndCore()
    Dim p As Object, f As TestXlFake, en As Long, ed As String
    Set p = NewTagged()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    Set f = NewFakeXl(5)
    PptTqq_SetExcelForTest f, ""
    AssertEqual "!" & PPT_MSG_OLD_CORE_STATES, PptTqq_LinkStates(p, "", False)
    Set f = NewFakeXl(6)
    PptTqq_SetExcelForTest f, ""
    AssertEqual "1|a|b|chart:stacked|manual|ok", PptTqq_LinkStates(p, "E", True)
    AssertEqual "E", f.LastEntries
    AssertTrue CBool(f.LastVerify)
    f.NoCore = True
    AssertEqual "!" & PPT_MSG_NOCORE_STATES, PptTqq_LinkStates(p, "", False)
    PptTqq_SetExcelForTest Nothing, ""
    AssertEqual "!" & PPT_MSG_NO_EXCEL, PptTqq_LinkStates(p, "", False)
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

Public Sub Test_PptRx_LinkedCallbacks_NoSelection()
    Dim v As Variant
    v = True: PptRx_GetLinkedEnabled Nothing, v
    AssertFalse CBool(v)
    v = True: PptRx_GetNewVisible Nothing, v
    AssertFalse CBool(v), "sem selecao: sem bandeira"
    v = False: PptRx_GetUpdVisible Nothing, v
    AssertTrue CBool(v), "sem selecao: o Atualizar normal aparece (desligado)"
    AssertEqual "never", PPT_DEFAULT_OPENMODE
End Sub
```
  (`NewFakeXl` e `NewTagged` já existem no módulo; `PptTqq_Tag` passa a `Public` para o teste.)

- [ ] **Passo 2:** monte o `.ppam` de teste (`build.ps1 -Target ppt -Configuration Test`, só sem POWERPNT do usuário) e rode `Test-PptAddin.ps1` → falha de
  compilação (watchdog/exit) = RED; anote.

- [ ] **Passo 3: constantes e estado** — em `modPptTqq.bas`, depois de `PPT_MSG_TEXT_SAVE`:

```vb
Public Const PPT_TAG_STATE As String = "TQQ_STATE"          ' = PPTSTATE_TAG do core (1.11.0): "new" = dados novos (a bandeira)
Public Const PPT_TAG_LINKIDX As String = "TQQ_LINKIDX"      ' = PPTSTATE_TAG_LINKIDX do core: indice de todos os elementos vinculados
Public Const PPT_API_STATES As Long = 6                    ' Core_PptLinkStates (core 1.11.0)
Public Const PPT_DEFAULT_OPENMODE As String = "never"      ' 1.11.0: como o think-cell, abrir nao pergunta (quem gravou outra escolha a mantem)
Public Const PPT_SUM_PREFIX As String = "Gráficos do ThinQQ atualizados: "
Public Const PPT_MSG_OLD_CORE_STATES As String = "Atualize o ThinQQ no Excel para ver o estado dos vínculos."
Public Const PPT_MSG_NOCORE_STATES As String = "O ThinQQ não está carregado no Excel: abra o Excel com o ThinQQ para ver o estado dos vínculos."
```
  declare `Private mLastResult As String` e troque, em `PptTqq_OpenMode`, `GetSetting(..., "ask")` por `GetSetting(PPT_APP_NAME, PPT_SECTION,
  PPT_KEY_OPENMODE, PPT_DEFAULT_OPENMODE)`; `Private Function PptTqq_Tag` e `Private Function PptTqq_Ask` → `Public`. Acrescente (antes de
  `' ---- 1.8.0 (Parte G)`):

```vb
' ---- 1.11.0: bandeira, todos os tipos de elemento, silencio no sucesso, abertura (spec 2026-10-23 secao 8.1) ----

' Elemento do ThinQQ vinculado: grafico (TQQ_CHART), texto (TQQ_TEXT = "1") ou tabela (TQQ_TABLE = "1"). Nunca propaga erro.
Public Function PptTqq_IsLinkedShape(shp As Object) As Boolean
    If Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0 Then PptTqq_IsLinkedShape = True: Exit Function
    If PptTqq_Tag(shp, PPT_TAG_TEXT) = "1" Then PptTqq_IsLinkedShape = True: Exit Function
    PptTqq_IsLinkedShape = (PptTqq_Tag(shp, PPT_TAG_TABLE) = "1")
End Function

' Alguma das formas tem dados novos (TQQ_STATE = "new", gravada pelo core). So etiquetas, em processo. Nunca propaga erro.
Public Function PptTqq_AnyNew(shapes As Collection) As Boolean
    Dim s As Object
    On Error Resume Next
    For Each s In shapes
        If PptTqq_Tag(s, PPT_TAG_STATE) = "new" Then PptTqq_AnyNew = True: Exit Function
    Next
End Function

' Resumo de sucesso sem nada a avisar: "Gráficos do ThinQQ atualizados: N." [" Sem mudança: M."] numa linha. Pura.
Public Function PptTqq_IsPlainSuccess(msg As String) As Boolean
    PptTqq_IsPlainSuccess = (Left$(msg, Len(PPT_SUM_PREFIX)) = PPT_SUM_PREFIX And InStr(msg, vbLf) = 0)
End Function

Public Function PptTqq_LastResult() As String
    PptTqq_LastResult = mLastResult
End Function

' Resultado de uma acao: sucesso puro em silencio (como o think-cell: a bandeira some); "!" no inicio = aviso de exclamacao; o resto,
' informacao. A faixa e redesenhada (bandeira). Nunca propaga erro.
Public Sub PptTqq_ShowResult(msg As String)
    On Error Resume Next
    mLastResult = msg
    If Len(msg) > 0 And Not PptTqq_IsPlainSuccess(msg) Then
        If Left$(msg, 1) = "!" Then PptTqq_Msg Mid$(msg, 2), vbExclamation Else PptTqq_Msg msg, vbInformation
    End If
    PptRx_InvalidateElem
End Sub

' TQQ_LINKIDX (formato do core) com todos os elementos vinculados das formas soltas, nomes unicos; so grava se mudou. Devolve quantos ha.
' Propaga erro (quem chama trata).
Public Function PptTqq_RebuildLinkIdx(pres As Object) As Long
    Dim sl As Object, s As Object, c As New Collection, idx As String
    For Each sl In pres.Slides
        For Each s In sl.Shapes
            If PptTqq_IsLinkedShape(s) Then
                PptTqq_EnsureUniqueName s
                c.Add Array(CLng(sl.SlideID), CStr(s.Name), PptTqq_SourcePath(PptTqq_LinkSource(s), PptTqq_Tag(s, PPT_TAG_BOOK)))
            End If
        Next
    Next
    idx = PptTqq_IdxJoin(c)
    If idx <> PptTqq_Tag(pres, PPT_TAG_LINKIDX) Then
        If Len(idx) = 0 Then PptTqq_TagDelete pres, PPT_TAG_LINKIDX Else pres.Tags.Add PPT_TAG_LINKIDX, idx
    End If
    PptTqq_RebuildLinkIdx = c.Count
End Function

' Ao abrir: tira as bandeiras gravadas no arquivo (TQQ_STATE salva pode nao valer mais) e reconstroi TQQ_LINKIDX; preserva Saved; somente
' leitura: nada. Devolve quantos elementos vinculados ha. Nunca propaga erro.
Public Function PptTqq_PrepareOnOpen(pres As Object) As Long
    Dim sl As Object, s As Object, was As Boolean
    On Error GoTo Done
    If CLng(pres.ReadOnly) <> 0 Then Exit Function
    was = (CLng(pres.Saved) <> 0)
    For Each sl In pres.Slides
        For Each s In sl.Shapes
            If Len(PptTqq_Tag(s, PPT_TAG_STATE)) > 0 Then PptTqq_TagDelete s, PPT_TAG_STATE
        Next
    Next
    PptTqq_PrepareOnOpen = PptTqq_RebuildLinkIdx(pres)
    If was And CLng(pres.Saved) = 0 Then pres.Saved = -1
    Exit Function
Done:
    PptTqq_Log "PptTqq_PrepareOnOpen: " & Err.Number & " " & Err.Description
    On Error Resume Next
    If was And CLng(pres.Saved) = 0 Then pres.Saved = -1
End Function

' Estados pelo core (API 6): so o Excel ja aberto (allowCreate False; com allowCreate, como "Atualizar uma vez", cria um invisivel e o fecha).
' Devolve as linhas do core ou "!" + mensagem. Reentrada: "". Nunca propaga erro.
Public Function PptTqq_LinkStates(pres As Object, entries As String, verifyClosed As Boolean, Optional allowCreate As Boolean = False) As String
    Dim xl As Object, created As Boolean, core As String, v As Variant, en As Long, ed As String
    If mBusy Then Exit Function
    mBusy = True
    On Error GoTo Fail
    Set xl = PptTqq_GetExcel(created, Not allowCreate)
    If xl Is Nothing Then
        PptTqq_LinkStates = "!" & PPT_MSG_NO_EXCEL
    Else
        core = PptTqq_EnsureCore(xl, created)
        If Len(core) = 0 Then
            PptTqq_LinkStates = "!" & PPT_MSG_NOCORE_STATES
        ElseIf PptTqq_CoreApi(xl, core) < PPT_API_STATES Then
            PptTqq_LinkStates = "!" & PPT_MSG_OLD_CORE_STATES
        Else
            v = xl.Run("'" & core & "'!Core_PptLinkStates", pres, entries, verifyClosed)
            PptTqq_LinkStates = CStr(v)
        End If
    End If
    PptTqq_Release xl, created
    mBusy = False
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptTqq_Release xl, created
    mBusy = False
    PptTqq_LinkStates = "!" & PptTqq_ErrorText(en, ed)
End Function
```
  Nota: no build Test `PptTqq_GetExcel(created, True)` devolve o Excel injetado (ou Nothing): o teste cobre os quatro caminhos.

- [ ] **Passo 4: seleção de qualquer tipo** — substitua `PptTqq_SelectedTagged` por:

```vb
' Formas do ThinQQ da selecao da janela ativa (formas ou texto dentro de uma forma), so as de um slide. anyKind: grafico, texto ou tabela
' vinculados (1.11.0: Atualizar e Automatico); senao so graficos (Editar dados e os comandos do grafico). Vazia sem janela/selecao. Nunca propaga.
Private Function PptTqq_SelectedOf(anyKind As Boolean) As Collection
    Dim out As New Collection, sel As Object, sr As Object, shp As Object, i As Long, sid As Long, ok As Boolean
    Set PptTqq_SelectedOf = out
    On Error Resume Next
    If Application.Windows.Count = 0 Then Exit Function
    Set sel = Application.ActiveWindow.Selection
    If sel Is Nothing Then Exit Function
    If sel.Type <> 2 And sel.Type <> 3 Then Exit Function          ' ppSelectionShapes, ppSelectionText
    Set sr = sel.ShapeRange
    If sr Is Nothing Then Exit Function
    For i = 1 To sr.Count
        Set shp = Nothing
        Set shp = sr(i)
        If Not shp Is Nothing Then
            If anyKind Then ok = PptTqq_IsLinkedShape(shp) Else ok = (Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0)
            If ok Then
                Err.Clear
                sid = CLng(shp.Parent.SlideID)
                If Err.Number = 0 Then out.Add shp
            End If
        End If
        Err.Clear
    Next
    Err.Clear
End Function

Public Function PptTqq_SelectedTagged() As Collection
    Set PptTqq_SelectedTagged = PptTqq_SelectedOf(False)
End Function

Public Function PptTqq_SelectedLinked() As Collection
    Set PptTqq_SelectedLinked = PptTqq_SelectedOf(True)
End Function
```

- [ ] **Passo 5: silêncio** — transforme `PptTqq_SyncEntries` em `Public Function PptTqq_SyncEntriesMsg(pres As Object, entries As String, force As
  Boolean, Optional fromAuto As Boolean = False) As String` com o MESMO corpo, trocando cada `PptTqq_Msg X, vbExclamation` por
  `PptTqq_SyncEntriesMsg = "!" & X` e cada `PptTqq_Msg X, vbInformation` / `PptTqq_Msg msg, vbInformation` por `PptTqq_SyncEntriesMsg = X` (o `Exit
  Sub` vira `Exit Function`; reentrada devolve ""), e acrescente

```vb
Public Sub PptTqq_SyncEntries(pres As Object, entries As String, force As Boolean, Optional fromAuto As Boolean = False)
    PptTqq_ShowResult PptTqq_SyncEntriesMsg(pres, entries, force, fromAuto)
End Sub
```
  Em `PptTqq_DoSync` troque `PptTqq_Msg msg, vbInformation` por `PptTqq_ShowResult msg` e, na falha, `PptTqq_Msg PptTqq_ErrorText(en, ed), vbExclamation`
  por `PptTqq_ShowResult "!" & PptTqq_ErrorText(en, ed)`. O aviso `PPT_MSG_NONE` de `PptTqq_DoSync` continua (não é sucesso). Os testes existentes
  que leem `PptTqq_LastMsg()` depois de `SyncEntries` com o dublê ("selecionados", "inteira", mensagens de erro) continuam valendo (não são sucesso
  puro); ajuste só os que esperavam caixa num resumo de sucesso puro (se algum quebrar: troque para `PptTqq_LastResult()`).

- [ ] **Passo 6: abertura** — em `PptTqq_HandleOpen`, depois de `n = PptTqq_CountTagged(pres, srcs)` acrescente `PptTqq_PrepareOnOpen pres` e troque o
  bloco `If act = "sync" ... End If` por

```vb
        If act = "sync" Then
            PptTqq_DoSync pres, False
        ElseIf act = "ask" Then
            If PptTqq_Ask(PptTqq_AskText(n, srcs)) = vbYes Then PptTqq_DoSync pres, False
        ElseIf n > 0 And BUILD_CONFIG <> "Test" Then
            mBusy = False                                     ' PptTqq_LinkStates tem a propria guarda de reentrada
            PptTqq_LinkStates pres, "", False                 ' 1.11.0: so com o Excel JA aberto com o core: as bandeiras aparecem; sem caixa
            mBusy = True
        End If
```
  Em `PptTqq_ToggleAuto` e `PptTqq_UpdateSelected` nada muda (o `SyncEntries` já silencia o sucesso).

- [ ] **Passo 7: faixa** — em `modPptRibbon.bas`: `PptRx_GetAuto`, `PptRx_OnAuto` e `PptRx_OnUpdateSel` passam a usar `PptTqq_SelectedLinked()`;
  acrescente

```vb
' 1.11.0: Atualizar e Automatico valem para grafico, texto e tabela vinculados.
Public Sub PptRx_GetLinkedEnabled(control As IRibbonControl, ByRef enabled)
    On Error Resume Next
    enabled = False
    enabled = (PptTqq_SelectedLinked().Count > 0)
End Sub

' Bandeira (1.11.0): com dados novos em alguma forma selecionada aparece "Dados novos: atualizar" (icone tqq_flag) no lugar de "Atualizar".
Public Sub PptRx_GetNewVisible(control As IRibbonControl, ByRef visible)
    On Error Resume Next
    visible = False
    visible = PptTqq_AnyNew(PptTqq_SelectedLinked())
End Sub

Public Sub PptRx_GetUpdVisible(control As IRibbonControl, ByRef visible)
    On Error Resume Next
    visible = True
    visible = Not PptTqq_AnyNew(PptTqq_SelectedLinked())
End Sub

' Item das listas da aba contextual: ids com o prefixo "c" (ids sao unicos no customUI inteiro). Pura.
Public Function PptRx_ItemPrefix(controlId As String) As String
    If Left$(controlId, 1) = "c" Then PptRx_ItemPrefix = "c"
End Function
```
  `PptRx_GetOptItem`: `itemId = PptRx_ItemPrefix(PptRx_IdOf(control)) & PptTqq_OptItemId(t, ...)`; `PptRx_OnOptItem`: se `PptRx_ItemPrefix(PptRx_IdOf(control)) = "c"`
  use `Mid$(selectedId, 2)` antes de `Mid$(..., Len(t) + 1)`; acrescente `Private Function PptRx_IdOf(control As IRibbonControl) As String` (como
  `PptRx_TagOf`, lendo `control.Id`). `PptRx_InvalidateElem` passa a invalidar também `"btnPptUpdateNew", "cbtnUpdate", "cbtnUpdateNew",
  "ctglAuto", "cbtnEditData", "ctglLabels", "ctglTotals", "ctglConnectors", "cddNumFmt", "cddLegend", "cbtnCagr", "cbtnDiff", "cbtnClearAnn"`.
  Ícone: em `build/lib/Ribbon-Icons.ps1` acrescente `$g2 = @(0, 153, 68)` dentro de `Get-TqqIconSpecs` e a entrada
  `tqq_flag = @{ Mso = 'Refresh'; Rects = @((New-TqqRect 3 1 2 14 $n), (New-TqqRect 5 1 9 6 $g2), (New-TqqRect 5 7 5 2 $g2)) }`.
  No `customUI14.xml` do `.ppam`: substitua o grupo `grpPptElem` por

```xml
        <group id="grpPptElem" label="Elemento selecionado">
          <button id="btnPptUpdateSel" label="Atualizar" size="large" imageMso="Refresh" onAction="PptRx_OnUpdateSel" getEnabled="PptRx_GetLinkedEnabled" getVisible="PptRx_GetUpdVisible"
                  screentip="Atualizar" supertip="Traz agora os dados do Excel para os gráficos, textos e tabelas do ThinQQ selecionados (dados e setas). Quando o Excel tem dados novos, este botão vira 'Dados novos: atualizar', com a bandeira verde."/>
          <button id="btnPptUpdateNew" label="Dados novos: atualizar" size="large" image="tqq_flag" onAction="PptRx_OnUpdateSel" getVisible="PptRx_GetNewVisible"
                  screentip="Dados novos no Excel" supertip="A pasta do Excel mudou depois da última atualização do elemento selecionado (a bandeira do think-cell). Clique para atualizar."/>
          <toggleButton id="tglPptAuto" label="Atualizar automaticamente" imageMso="Refresh" getPressed="PptRx_GetAuto" onAction="PptRx_OnAuto" getEnabled="PptRx_GetLinkedEnabled"
                        screentip="Atualizar automaticamente" supertip="Liga a atualização automática só nos elementos selecionados: com o Excel e o PowerPoint abertos, eles acompanham as mudanças da pasta (cerca de 2 segundos depois). Por padrão tudo é manual, como no think-cell."/>
          <button id="btnPptEditData" label="Editar dados" imageMso="TableInsert" onAction="PptRx_OnEditData" getEnabled="PptRx_GetElemEnabled"
                  screentip="Editar dados" supertip="Abre no Excel a planilha com os dados do gráfico selecionado (para gráficos criados aqui, a pasta de dados ao lado da apresentação). Depois de editar, clique em Atualizar."/>
          <button id="btnPptLinkText" label="Texto vinculado ao Excel" imageMso="TextBoxInsert" onAction="PptRx_OnLinkText"
                  screentip="Texto vinculado ao Excel" supertip="Selecione no Excel a célula e, aqui, a caixa de texto (ou a parte do texto, ou ponha o cursor): o trecho passa a mostrar o valor da célula. Exige o Excel aberto com o ThinQQ 1.9.0 ou mais novo."/>
        </group>
```
  e acrescente, depois de `</tabs>` e antes de `</ribbon>`:

```xml
    <!-- 1.11.0: aba contextual ThinQQ ao selecionar um grafico (o mini toolbar do think-cell). UNICO idMso de contexto permitido: TabSetChartTools
         (Test-Ribbon); um idMso invalido faz o PowerPoint descartar a faixa inteira em silencio (spec 2026-10-23 secao 6.1, PX2). Ids com prefixo "c". -->
    <contextualTabs>
      <tabSet idMso="TabSetChartTools">
        <tab id="tabTqqChartCtx" label="ThinQQ">
          <group id="cgrpLink" label="Vínculo">
            <button id="cbtnUpdate" label="Atualizar" size="large" imageMso="Refresh" onAction="PptRx_OnUpdateSel" getEnabled="PptRx_GetLinkedEnabled" getVisible="PptRx_GetUpdVisible" screentip="Atualizar"/>
            <button id="cbtnUpdateNew" label="Dados novos: atualizar" size="large" image="tqq_flag" onAction="PptRx_OnUpdateSel" getVisible="PptRx_GetNewVisible" screentip="Dados novos no Excel"/>
            <toggleButton id="ctglAuto" label="Automático" imageMso="Refresh" getPressed="PptRx_GetAuto" onAction="PptRx_OnAuto" getEnabled="PptRx_GetLinkedEnabled" screentip="Atualizar automaticamente"/>
            <button id="cbtnEditData" label="Editar dados" imageMso="TableInsert" onAction="PptRx_OnEditData" getEnabled="PptRx_GetElemEnabled" screentip="Editar dados"/>
          </group>
          <group id="cgrpChart" label="Gráfico do ThinQQ">
            <toggleButton id="ctglLabels" label="Rótulos" imageMso="TextBoxInsert" tag="labels" getPressed="PptRx_GetOpt" onAction="PptRx_OnToggleOpt" getEnabled="PptRx_GetOptEnabled"/>
            <toggleButton id="ctglTotals" label="Totais" imageMso="AutoSum" tag="totals" getPressed="PptRx_GetOpt" onAction="PptRx_OnToggleOpt" getEnabled="PptRx_GetOptEnabled"/>
            <toggleButton id="ctglConnectors" label="Conectores" imageMso="ShapeStraightConnector" tag="connectors" getPressed="PptRx_GetOpt" onAction="PptRx_OnToggleOpt" getEnabled="PptRx_GetOptEnabled"/>
            <dropDown id="cddNumFmt" label="Formato" tag="numfmt" getSelectedItemID="PptRx_GetOptItem" onAction="PptRx_OnOptItem" getEnabled="PptRx_GetOptEnabled">
              <item id="cnumfmtint" label="1.234"/>
              <item id="cnumfmtdec1" label="1.234,5"/>
              <item id="cnumfmtdec2" label="1.234,56"/>
              <item id="cnumfmtpct0" label="12%"/>
              <item id="cnumfmtpct1" label="12,3%"/>
              <item id="cnumfmtmil" label="Milhares (1.234 → 1)"/>
              <item id="cnumfmtmi" label="Milhões (1,2)"/>
            </dropDown>
            <dropDown id="cddLegend" label="Legenda" tag="legend" getSelectedItemID="PptRx_GetOptItem" onAction="PptRx_OnOptItem" getEnabled="PptRx_GetOptEnabled">
              <item id="clegendright" label="Direita"/>
              <item id="clegendnone" label="Nenhuma"/>
            </dropDown>
            <button id="cbtnCagr" label="Seta CAGR" imageMso="ArrowsMore" tag="cagr" onAction="PptRx_OnArrow" getEnabled="PptRx_GetOptEnabled"/>
            <button id="cbtnDiff" label="Seta diferença" imageMso="ShapeUpArrow" tag="diff" onAction="PptRx_OnArrow" getEnabled="PptRx_GetOptEnabled"/>
            <button id="cbtnClearAnn" label="Remover setas" imageMso="Delete" onAction="PptRx_OnClearAnn" getEnabled="PptRx_GetElemEnabled"/>
          </group>
        </tab>
      </tabSet>
    </contextualTabs>
```
  (O `→` dentro do XML é permitido: a regra cp1252 vale para fontes VBA; o XML já tem esse caractere no item `numfmtmil`.)

- [ ] **Passo 8: Test-Ribbon / Test-Docs** — em `tests/ps/Test-Ribbon.ps1`: troque `getEnabled='PptRx_GetElemEnabled'` por `'PptRx_GetLinkedEnabled'`
  nas checagens de `tglPptAuto` e `btnPptUpdateSel` (rótulo novo "Atualizar"); troque a regra "image= so no grupo Inserir gráfico" por
  `Check ($null -eq $px.SelectSingleNode("//u:*[@image and not(ancestor-or-self::u:group[@id='grpPptInsert']) and not(@image='tqq_flag') and not(@image='tqq_links')]", $pns)) 'ppt: image= so em Inserir gráfico e nas bandeiras/Vínculos'`;
  acrescente `Check ((@($px.SelectNodes('//u:tabSet', $pns) | ForEach-Object { $_.GetAttribute('idMso') }) -join ',') -eq 'TabSetChartTools') 'ppt: unico tabSet = TabSetChartTools'`,
  `Check ($null -ne $px.SelectSingleNode("//u:tab[@id='tabTqqChartCtx']/u:group[@id='cgrpLink']/u:button[@id='cbtnUpdateNew'][@image='tqq_flag'][@getVisible='PptRx_GetNewVisible']", $pns)) 'ppt: bandeira na aba contextual'`
  e `Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptElem']/u:button[@id='btnPptUpdateNew'][@getVisible='PptRx_GetNewVisible']", $pns)) 'ppt: bandeira no grupo Elemento selecionado'`.
  Em `tests/ps/Test-Docs.ps1` rode e, para cada pino que esta tarefa quebrar (rótulo "Atualizar uma vez" do `btnPptUpdateSel` ~linha 889, lista de
  rótulos conhecidos ~linha 1498), troque o texto esperado pelo novo **e** acrescente o rótulo novo ao tutorial/README só se o próprio Test-Docs exigir
  (a reescrita completa da documentação é a U5); liste no relatório cada pino tocado.

- [ ] **Passo 9:** suíte completa com `-Ppt` → saída 0; sem POWERPNT do usuário, 0 testes de PowerPoint pulados (senão: relatório diz quantos).
- [ ] **Passo 10:** commits: `feat: ppam com bandeira de dados novos, aba contextual ThinQQ e silencio no sucesso` e `test: bandeira, abertura e estados no ppam`.

---

### Tarefa U3: Diálogo "Vínculos" (Data Links) não modal no `.ppam` — modelo: most-capable

**Files:**
- Create: `src/ppt/modules/modPptLinks.bas`, `src/ppt/forms/frmTqqLinks.json`, `src/ppt/forms/frmTqqLinks.vba`
- Modify: `build/build.ps1` (alvo `ppt` importa formulários), `build/lib/Ribbon-Icons.ps1` (`tqq_links`), `src/ppt/ribbon/customUI14.xml` (grupo "Dados do
  Excel": `btnPptLinks`, sem `ddOpenMode`; `cbtnLinks` na aba contextual), `src/ppt/modules/modPptRibbon.bas` (`PptRx_OnLinks`), `src/ppt/modules/modPptTqq.bas`
  (`PptTqq_OnClose` fecha o diálogo da apresentação; `PptTqq_CoreFlag`/`PptTqq_SetCoreFlag`)
- Test: `tests/vba/ppt/TestPptAddin.bas`, `tests/ps/Test-Ribbon.ps1`, `tests/ps/Test-Docs.ps1` (pinos de "Ao abrir")

**Interfaces:**
- Produces: `PptLinks_Collect(pres As Object) As Collection` (itens `Array(slideIndex, slideId, nome, tipo, modo, estado, pasta)`; tipo `chart`|`text`|`table`;
  modo `auto`|`manual`; estado: `new` pela etiqueta, senão o do cache da sessão (`PptLinks_Remember`), senão `""`); `PptLinks_Remember(lines As String) As
  Long`; `PptLinks_Filter(rows As Collection, filterIndex As Long) As Collection` (0 todos, 1 dados novos, 2 atualizados, 3 automáticos, 4 com
  problema); `PptLinks_StateLabel(code) As String`; `PptLinks_KindLabel(kind) As String`; `PptLinks_ModeLabel(mode) As String`; `PptLinks_FileName(p)
  As String`; `PptLinks_HeadText(rows) As String`; `PptLinks_OpenModeIndex(mode) As Long`; `PptLinks_OpenModeOf(index) As String`; `PptLinks_Shapes(pres,
  rows) As Collection`; `PptLinks_Do(pres, what, rows) As String` (`update`, `auto`, `manual`, `unlink`, `edit`, `goto`, `check`); `PptLinks_Unlink(shapes)
  As Long`; `PptLinks_Show()`; `PptLinks_OnClose(pres)`; `PptLinks_ResetTestHooks()`; `PptTqq_CoreFlag(key, ByRef available) As Boolean`;
  `PptTqq_SetCoreFlag(key, v) As String`; callback `PptRx_OnLinks(control)`.

- [ ] **Passo 1: Testes (RED)** — acrescente a `tests/vba/ppt/TestPptAddin.bas`:

```vb
' ---- 1.11.0 (U3): dialogo Vinculos (modPptLinks; o formulario so e conferido pela legenda) ----

Public Sub Test_PptForms_LinksCaption()
    Dim f As Object
    Set f = VBA.UserForms.Add("frmTqqLinks")
    AssertEqual "Vínculos do ThinQQ", f.Caption
    AssertEqual 5, CLng(f.Controls("lstLinks").ColumnCount)
    Unload f
End Sub

Public Sub Test_PptLinks_Labels()
    AssertEqual "Dados novos", PptLinks_StateLabel("new")
    AssertEqual "Atualizado", PptLinks_StateLabel("ok")
    AssertEqual "Pasta não encontrada", PptLinks_StateLabel("missing")
    AssertEqual "Não encontrado na pasta", PptLinks_StateLabel("nosource")
    AssertEqual "Sem vínculo", PptLinks_StateLabel("unlinked")
    AssertEqual "Não verificado (pasta fechada)", PptLinks_StateLabel("unknown")
    AssertEqual "Não verificado", PptLinks_StateLabel("")
    AssertEqual "Gráfico", PptLinks_KindLabel("chart")
    AssertEqual "Texto", PptLinks_KindLabel("text")
    AssertEqual "Tabela", PptLinks_KindLabel("table")
    AssertEqual "Automático", PptLinks_ModeLabel("auto")
    AssertEqual "Manual", PptLinks_ModeLabel("manual")
    AssertEqual "b.xlsx", PptLinks_FileName("C:\a\b.xlsx")
    AssertEqual "c.xlsx", PptLinks_FileName("https://x/y/c.xlsx")
    AssertEqual 0, PptLinks_OpenModeIndex("never")
    AssertEqual 1, PptLinks_OpenModeIndex("ask")
    AssertEqual 2, PptLinks_OpenModeIndex("always")
    AssertEqual "never", PptLinks_OpenModeOf(0)
    AssertEqual "always", PptLinks_OpenModeOf(2)
End Sub

Public Sub Test_PptLinks_CollectFilterAndCache()
    Dim p As Object, a As Object, b As Object, c As Object, rows As Collection, sid As Long
    Set p = NewPres()
    On Error GoTo Cleanup
    PptLinks_ResetTestHooks
    sid = CLng(p.Slides(1).SlideID)
    Set a = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_a": a.Tags.Add PPT_TAG_CHART, "abc": a.Tags.Add PPT_TAG_BOOK, "C:\x\a.xlsx"
    a.Tags.Add PPT_TAG_STATE, "new"
    Set b = p.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50): b.Name = "txt_b": b.Tags.Add PPT_TAG_TEXT, "1": b.Tags.Add PPT_TAG_BOOK, "C:\x\b.xlsx"
    b.Tags.Add PPT_TAG_AUTO, "1"
    Set c = p.Slides(1).Shapes.AddShape(1, 130, 10, 50, 50)
    Set rows = PptLinks_Collect(p)
    AssertEqual 2, rows.Count
    AssertEqual "tqq_a", CStr(rows(1)(2))
    AssertEqual "chart", CStr(rows(1)(3))
    AssertEqual "manual", CStr(rows(1)(4))
    AssertEqual "new", CStr(rows(1)(5))
    AssertEqual "auto", CStr(rows(2)(4))
    AssertEqual "", CStr(rows(2)(5)), "sem verificar"
    AssertEqual 1, PptLinks_Filter(rows, 1).Count, "dados novos"
    AssertEqual 1, PptLinks_Filter(rows, 3).Count, "automaticos"
    AssertEqual 0, PptLinks_Filter(rows, 4).Count, "com problema"
    AssertEqual 1, PptLinks_Remember(sid & "|txt_b|C:\x\b.xlsx|text|auto|missing")
    Set rows = PptLinks_Collect(p)
    AssertEqual "missing", CStr(rows(2)(5)), "estado do core lembrado na sessao"
    AssertEqual 1, PptLinks_Filter(rows, 4).Count
    AssertEqual 0, PptLinks_Remember("lixo" & vbLf & "1|2"), "linha invalida ignorada"
    AssertTrue InStr(PptLinks_HeadText(rows), "2 elemento(s)") > 0, PptLinks_HeadText(rows)
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptLinks_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Acoes sem Excel: manual/automatico gravam a marca e o indice; desvincular tira so as etiquetas TQQ_* (a forma fica); sem selecao: aviso.
Public Sub Test_PptLinks_Do_ModesAndUnlink_NoExcel()
    Dim p As Object, a As Object, rows As Collection, msg As String
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks: PptLinks_ResetTestHooks
    Set a = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_a": a.Tags.Add PPT_TAG_CHART, "abc": a.Tags.Add PPT_TAG_BOOK, "C:\x\a.xlsx"
    Set rows = PptLinks_Collect(p)
    AssertEqual PPTLINKS_MSG_SELECT, PptLinks_Do(p, "update", New Collection)
    msg = PptLinks_Do(p, "manual", rows)
    AssertEqual "", PptTqq_Tag(a, PPT_TAG_AUTO)
    msg = PptLinks_Do(p, "auto", rows)
    AssertEqual "1", PptTqq_Tag(a, PPT_TAG_AUTO)
    AssertTrue InStr(PptTqq_Tag(p, PPT_TAG_INDEX), "tqq_a") > 0, "no indice dos automaticos"
    AssertTrue Len(msg) > 0, "sem Excel: a dica de abrir o Excel"
    msg = PptLinks_Do(p, "unlink", rows)
    AssertEqual "", PptTqq_Tag(a, PPT_TAG_CHART), "desvinculado"
    AssertEqual "", PptTqq_Tag(a, PPT_TAG_AUTO)
    AssertEqual "", PptTqq_Tag(p, PPT_TAG_INDEX), "fora do indice dos automaticos"
    AssertEqual "", PptTqq_Tag(p, PPT_TAG_LINKIDX), "fora do indice de todos"
    AssertEqual 1, p.Slides(1).Shapes.Count, "a forma continua no slide"
    AssertEqual "Desvinculados: 1.", msg
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks: PptLinks_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

Public Sub Test_PptLinks_Do_Check_UsesCoreAndRemembers()
    Dim p As Object, a As Object, f As TestXlFake, rows As Collection, msg As String, sid As Long
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks: PptLinks_ResetTestHooks
    sid = CLng(p.Slides(1).SlideID)
    Set a = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_a": a.Tags.Add PPT_TAG_CHART, "abc"
    Set f = NewFakeXl(6)
    f.StatesResult = sid & "|tqq_a|C:\x\a.xlsx|chart:stacked|manual|nosource"
    PptTqq_SetExcelForTest f, ""
    msg = PptLinks_Do(p, "check", PptLinks_Collect(p))
    AssertTrue CBool(f.LastVerify), "Verificar abre pastas fechadas (somente leitura, no core)"
    AssertEqual "", f.LastEntries, "todos"
    AssertEqual "nosource", CStr(PptLinks_Collect(p)(1)(5))
    AssertTrue InStr(msg, "Verificados: 1") = 1, msg
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks: PptLinks_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub
```

- [ ] **Passo 2:** RED (compilação: `modPptLinks` não existe).

- [ ] **Passo 3: build** — em `build/build.ps1`, alvo `ppt`, depois da linha que importa `classes`:

```powershell
        $formsDir = Join-Path $srcDir 'forms'
        if (Test-Path $formsDir) {
            Get-ChildItem $formsDir -Filter *.json | ForEach-Object { Add-FormFromJson $vbproj $_.FullName (Join-Path $formsDir ($_.BaseName + '.vba')) }
        }
```
  (mesma função do Excel; se o designer do PowerPoint recusar alguma propriedade, o erro sobe no build: registre e ajuste o JSON, nunca o
  `Add-FormFromJson` para o Excel).

- [ ] **Passo 4: formulário** — `src/ppt/forms/frmTqqLinks.json`:

```json
{
  "name": "frmTqqLinks",
  "caption": "Vínculos do ThinQQ",
  "width": 650,
  "height": 420,
  "controls": [
    { "type": "Label", "name": "lblHead", "Left": 12, "Top": 8, "Width": 440, "Height": 14, "Caption": "Elementos do ThinQQ vinculados ao Excel", "fontName": "Segoe UI", "fontSize": 9, "fontBold": true },
    { "type": "Label", "name": "lblFilter", "Left": 460, "Top": 8, "Width": 40, "Height": 14, "Caption": "Mostrar:", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "ComboBox", "name": "cboFilter", "Left": 500, "Top": 5, "Width": 136, "Height": 20, "Style": 2, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "Label", "name": "lblColSlide", "Left": 14, "Top": 32, "Width": 34, "Height": 12, "Caption": "Slide", "fontName": "Segoe UI", "fontSize": 8 },
    { "type": "Label", "name": "lblColElem", "Left": 50, "Top": 32, "Width": 150, "Height": 12, "Caption": "Elemento", "fontName": "Segoe UI", "fontSize": 8 },
    { "type": "Label", "name": "lblColMode", "Left": 202, "Top": 32, "Width": 60, "Height": 12, "Caption": "Modo", "fontName": "Segoe UI", "fontSize": 8 },
    { "type": "Label", "name": "lblColState", "Left": 264, "Top": 32, "Width": 140, "Height": 12, "Caption": "Estado", "fontName": "Segoe UI", "fontSize": 8 },
    { "type": "Label", "name": "lblColSrc", "Left": 406, "Top": 32, "Width": 220, "Height": 12, "Caption": "Origem", "fontName": "Segoe UI", "fontSize": 8 },
    { "type": "ListBox", "name": "lstLinks", "Left": 12, "Top": 46, "Width": 624, "Height": 214, "ColumnCount": 5, "ColumnWidths": "36 pt;152 pt;62 pt;142 pt;220 pt", "MultiSelect": 2, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnUpdate", "Left": 12, "Top": 268, "Width": 90, "Height": 24, "Caption": "Atualizar", "fontName": "Segoe UI", "fontSize": 9, "fontBold": true },
    { "type": "CommandButton", "name": "btnUpdateNew", "Left": 106, "Top": 268, "Width": 150, "Height": 24, "Caption": "Atualizar com dados novos", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnAuto", "Left": 260, "Top": 268, "Width": 80, "Height": 24, "Caption": "Automático", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnManual", "Left": 344, "Top": 268, "Width": 70, "Height": 24, "Caption": "Manual", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnGoto", "Left": 418, "Top": 268, "Width": 70, "Height": 24, "Caption": "Ir para", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnEdit", "Left": 492, "Top": 268, "Width": 144, "Height": 24, "Caption": "Abrir no Excel", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnCheck", "Left": 12, "Top": 298, "Width": 90, "Height": 24, "Caption": "Verificar", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnUnlink", "Left": 106, "Top": 298, "Width": 90, "Height": 24, "Caption": "Desvincular", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnAll", "Left": 200, "Top": 298, "Width": 100, "Height": 24, "Caption": "Selecionar tudo", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnClose", "Left": 556, "Top": 298, "Width": 80, "Height": 24, "Caption": "Fechar", "Cancel": true, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CheckBox", "name": "chkEngine", "Left": 12, "Top": 332, "Width": 330, "Height": 18, "Caption": "Acompanhar mudanças do Excel (bandeiras e automáticos)", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CheckBox", "name": "chkNewAuto", "Left": 350, "Top": 332, "Width": 286, "Height": 18, "Caption": "Novos vínculos nascem automáticos", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "Label", "name": "lblOpen", "Left": 12, "Top": 358, "Width": 140, "Height": 14, "Caption": "Ao abrir a apresentação:", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "ComboBox", "name": "cboOpen", "Left": 154, "Top": 355, "Width": 150, "Height": 20, "Style": 2, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "Label", "name": "lblStatus", "Left": 12, "Top": 382, "Width": 624, "Height": 14, "Caption": "", "fontName": "Segoe UI", "fontSize": 9 }
  ]
}
```
  `src/ppt/forms/frmTqqLinks.vba`:

```vb
Option Explicit
' Vinculos do ThinQQ (1.11.0; o Data Links do think-cell): formulario NAO modal. A logica e testavel em modPptLinks; aqui so a tela.
Private mPres As Object
Private mRows As Collection
Private mShown As Collection
Private mLoading As Boolean

Public Property Get Pres() As Object
    Set Pres = mPres
End Property

Public Sub Prepare(p As Object)
    Dim avail As Boolean
    Set mPres = p
    mLoading = True
    cboFilter.Clear
    cboFilter.AddItem "Todos"
    cboFilter.AddItem "Dados novos"
    cboFilter.AddItem "Atualizados"
    cboFilter.AddItem "Automáticos"
    cboFilter.AddItem "Com problema"
    cboFilter.ListIndex = 0
    cboOpen.Clear
    cboOpen.AddItem "Não perguntar"
    cboOpen.AddItem "Perguntar"
    cboOpen.AddItem "Atualizar sempre"
    cboOpen.ListIndex = PptLinks_OpenModeIndex(PptTqq_OpenMode())
    chkEngine.Value = PptTqq_CoreFlag("autoon", avail)
    chkEngine.Enabled = avail
    chkNewAuto.Value = PptTqq_CoreFlag("autonew", avail)
    chkNewAuto.Enabled = avail
    If Not avail Then lblStatus.Caption = "Abra o Excel com o ThinQQ para ver o estado e mudar as opções de acompanhamento."
    mLoading = False
    Reload
End Sub

Public Sub Reload()
    Set mRows = PptLinks_Collect(mPres)
    Fill
End Sub

Private Sub Fill()
    Dim r As Variant, i As Long
    Set mShown = PptLinks_Filter(mRows, cboFilter.ListIndex)
    lstLinks.Clear
    For Each r In mShown
        lstLinks.AddItem CStr(r(0))
        i = lstLinks.ListCount - 1
        lstLinks.List(i, 1) = PptLinks_KindLabel(CStr(r(3))) & " " & CStr(r(2))
        lstLinks.List(i, 2) = PptLinks_ModeLabel(CStr(r(4)))
        lstLinks.List(i, 3) = PptLinks_StateLabel(CStr(r(5)))
        lstLinks.List(i, 4) = PptLinks_FileName(CStr(r(6)))
    Next
    lblHead.Caption = PptLinks_HeadText(mRows)
End Sub

Private Function SelectedRows() As Collection
    Dim c As New Collection, i As Long
    For i = 0 To lstLinks.ListCount - 1
        If lstLinks.Selected(i) Then c.Add mShown(i + 1)
    Next
    Set SelectedRows = c
End Function

Private Sub Act(what As String, rows As Collection)
    lblStatus.Caption = PptLinks_Do(mPres, what, rows)
    Reload
End Sub

Private Sub btnUpdate_Click()
    Act "update", SelectedRows()
End Sub

Private Sub btnUpdateNew_Click()
    Act "update", PptLinks_Filter(mRows, 1)
End Sub

Private Sub btnAuto_Click()
    Act "auto", SelectedRows()
End Sub

Private Sub btnManual_Click()
    Act "manual", SelectedRows()
End Sub

Private Sub btnGoto_Click()
    Act "goto", SelectedRows()
End Sub

Private Sub btnEdit_Click()
    Act "edit", SelectedRows()
End Sub

Private Sub btnCheck_Click()
    Act "check", mRows
End Sub

Private Sub btnUnlink_Click()
    Act "unlink", SelectedRows()
End Sub

Private Sub btnAll_Click()
    Dim i As Long
    For i = 0 To lstLinks.ListCount - 1
        lstLinks.Selected(i) = True
    Next
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cboFilter_Change()
    If Not mLoading Then Fill
End Sub

Private Sub cboOpen_Change()
    If Not mLoading Then PptTqq_SetOpenMode PptLinks_OpenModeOf(cboOpen.ListIndex)
End Sub

Private Sub chkEngine_Click()
    If Not mLoading Then lblStatus.Caption = PptTqq_SetCoreFlag("autoon", CBool(chkEngine.Value))
End Sub

Private Sub chkNewAuto_Click()
    If Not mLoading Then lblStatus.Caption = PptTqq_SetCoreFlag("autonew", CBool(chkNewAuto.Value))
End Sub

Private Sub lstLinks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Act "goto", SelectedRows()
End Sub
```

- [ ] **Passo 5: `modPptLinks`** — `src/ppt/modules/modPptLinks.bas`:

```vb
Attribute VB_Name = "modPptLinks"
Option Explicit
' Logica do dialogo Vinculos (1.11.0; spec 2026-10-23 secao 8.2). Listar e filtrar so leem etiquetas (em processo, sem Excel); os estados que
' nao estao na etiqueta (pasta nao encontrada, nao encontrado na pasta, sem vinculo, fechada) vem do core (Verificar; abertura) e ficam
' lembrados na sessao (mStates). Acoes reaproveitam o .ppam: Atualizar = PptTqq_SyncEntriesMsg; Automatico/Manual = PptTqq_SetAutoShapes;
' Abrir no Excel = PptTqq_EditShapes. Desvincular tira as etiquetas TQQ_* (o vinculo nativo do PowerPoint fica: nunca BreakLink, S10).
Public Const PPTLINKS_MSG_SELECT As String = "Selecione um ou mais elementos na lista."
Private Const PPTLINKS_TAGS As String = "TQQ_CHART,TQQ_BOOK,TQQ_SHEET,TQQ_VER,TQQ_FP,TQQ_TITLE,TQQ_KIND,TQQ_OPTS,TQQ_DATA,TQQ_TEXT,TQQ_TNAME,TQQ_TREF,TQQ_TPRE,TQQ_TSUF,TQQ_TLAST,TQQ_TABLE,TQQ_TBLAST,TQQ_STATE,TQQ_AUTO"
Private mStates As Object                       ' "slideId|nome em minusculas" -> estado vindo do core nesta sessao

Public Sub PptLinks_ResetTestHooks()
    Set mStates = Nothing
End Sub

Public Function PptLinks_StateLabel(code As String) As String
    Select Case code
        Case "new": PptLinks_StateLabel = "Dados novos"
        Case "ok": PptLinks_StateLabel = "Atualizado"
        Case "missing": PptLinks_StateLabel = "Pasta não encontrada"
        Case "nosource": PptLinks_StateLabel = "Não encontrado na pasta"
        Case "unlinked": PptLinks_StateLabel = "Sem vínculo"
        Case "unknown": PptLinks_StateLabel = "Não verificado (pasta fechada)"
        Case Else: PptLinks_StateLabel = "Não verificado"
    End Select
End Function

Public Function PptLinks_KindLabel(kind As String) As String
    Select Case kind
        Case "text": PptLinks_KindLabel = "Texto"
        Case "table": PptLinks_KindLabel = "Tabela"
        Case Else: PptLinks_KindLabel = "Gráfico"
    End Select
End Function

Public Function PptLinks_ModeLabel(mode As String) As String
    If mode = "auto" Then PptLinks_ModeLabel = "Automático" Else PptLinks_ModeLabel = "Manual"
End Function

Public Function PptLinks_FileName(p As String) As String
    Dim k As Long
    k = InStrRev(p, "\")
    If InStrRev(p, "/") > k Then k = InStrRev(p, "/")
    PptLinks_FileName = Mid$(p, k + 1)
End Function

Public Function PptLinks_OpenModeIndex(mode As String) As Long
    Select Case mode
        Case "ask": PptLinks_OpenModeIndex = 1
        Case "always": PptLinks_OpenModeIndex = 2
        Case Else: PptLinks_OpenModeIndex = 0
    End Select
End Function

Public Function PptLinks_OpenModeOf(index As Long) As String
    Select Case index
        Case 1: PptLinks_OpenModeOf = "ask"
        Case 2: PptLinks_OpenModeOf = "always"
        Case Else: PptLinks_OpenModeOf = "never"
    End Select
End Function

Private Function PptLinks_Key(slideId As Variant, nm As Variant) As String
    PptLinks_Key = CStr(slideId) & "|" & LCase$(CStr(nm))
End Function

' Guarda os estados devolvidos pelo core (linhas de Core_PptLinkStates). Devolve quantas linhas validas. Pura sobre o cache da sessao.
Public Function PptLinks_Remember(lines As String) As Long
    Dim a() As String, f() As String, i As Long
    If mStates Is Nothing Then Set mStates = CreateObject("Scripting.Dictionary")
    If Len(lines) = 0 Or Left$(lines, 1) = "!" Then Exit Function
    a = Split(lines, vbLf)
    For i = 0 To UBound(a)
        f = Split(a(i), "|")
        If UBound(f) = 5 Then
            If Len(f(0)) > 0 And Len(f(0)) < 10 And Not f(0) Like "*[!0-9]*" Then
                mStates(PptLinks_Key(f(0), f(1))) = f(5)
                PptLinks_Remember = PptLinks_Remember + 1
            End If
        End If
    Next
End Function

' Linhas do dialogo: Array(slideIndex, slideId, nome, tipo, modo, estado, pasta) de cada forma vinculada solta, na ordem dos slides. Nunca propaga.
Public Function PptLinks_Collect(pres As Object) As Collection
    Dim out As New Collection, sl As Object, s As Object, kind As String, mode As String, st As String, k As String
    Set PptLinks_Collect = out
    On Error Resume Next
    If mStates Is Nothing Then Set mStates = CreateObject("Scripting.Dictionary")
    For Each sl In pres.Slides
        For Each s In sl.Shapes
            If PptTqq_IsLinkedShape(s) Then
                If Len(PptTqq_Tag(s, PPT_TAG_CHART)) > 0 Then
                    kind = "chart"
                ElseIf PptTqq_Tag(s, PPT_TAG_TEXT) = "1" Then
                    kind = "text"
                Else
                    kind = "table"
                End If
                If PptTqq_Tag(s, PPT_TAG_AUTO) = "1" Then mode = "auto" Else mode = "manual"
                k = PptLinks_Key(sl.SlideID, s.Name)
                If PptTqq_Tag(s, PPT_TAG_STATE) = "new" Then
                    st = "new"
                ElseIf mStates.Exists(k) Then
                    st = CStr(mStates(k))
                    If st = "new" Then st = "ok"            ' o core marcou, a etiqueta saiu depois (atualizado)
                Else
                    st = ""
                End If
                out.Add Array(CLng(sl.SlideIndex), CLng(sl.SlideID), CStr(s.Name), kind, mode, st, PptTqq_SourcePath(PptTqq_LinkSourceOf(s), PptTqq_Tag(s, PPT_TAG_BOOK)))
            End If
        Next
    Next
    Err.Clear
End Function

Public Function PptLinks_Filter(rows As Collection, filterIndex As Long) As Collection
    Dim out As New Collection, r As Variant, ok As Boolean
    For Each r In rows
        Select Case filterIndex
            Case 1: ok = (r(5) = "new")
            Case 2: ok = (r(5) = "ok")
            Case 3: ok = (r(4) = "auto")
            Case 4: ok = (r(5) = "missing" Or r(5) = "nosource" Or r(5) = "unlinked")
            Case Else: ok = True
        End Select
        If ok Then out.Add r
    Next
    Set PptLinks_Filter = out
End Function

Public Function PptLinks_HeadText(rows As Collection) As String
    Dim r As Variant, nNew As Long
    For Each r In rows
        If r(5) = "new" Then nNew = nNew + 1
    Next
    PptLinks_HeadText = rows.Count & " elemento(s) vinculado(s) ao Excel; " & nNew & " com dados novos."
End Function

' Formas das linhas (slide pelo SlideID, forma pelo nome). Linhas que nao acham mais a forma sao ignoradas. Nunca propaga.
Public Function PptLinks_Shapes(pres As Object, rows As Collection) As Collection
    Dim out As New Collection, r As Variant, sl As Object, s As Object
    On Error Resume Next
    For Each r In rows
        Set sl = Nothing: Set s = Nothing
        Set sl = pres.Slides.FindBySlideID(CLng(r(1)))
        If Not sl Is Nothing Then Set s = sl.Shapes(CStr(r(2)))
        If Not s Is Nothing Then out.Add s
        Err.Clear
    Next
    Set PptLinks_Shapes = out
End Function

' Desvincular: tira do indice dos automaticos e apaga as etiquetas TQQ_* das formas (a forma e o vinculo nativo ficam); reconstroi TQQ_LINKIDX.
Public Function PptLinks_Unlink(shapes As Collection) As Long
    Dim s As Object, k As Variant, pres As Object
    If shapes.Count = 0 Then Exit Function
    Set pres = shapes(1).Parent.Parent
    PptTqq_SetAutoShapes shapes, False
    For Each s In shapes
        For Each k In Split(PPTLINKS_TAGS, ",")
            On Error Resume Next
            s.Tags.Delete CStr(k)
            Err.Clear
            On Error GoTo 0
        Next
        PptLinks_Unlink = PptLinks_Unlink + 1
    Next
    PptTqq_RebuildLinkIdx pres
End Function

' Uma acao do dialogo sobre as linhas dadas. Devolve o texto da linha de status (vazio quando a acao ja mostrou a propria mensagem). Nunca propaga.
Public Function PptLinks_Do(pres As Object, what As String, rows As Collection) As String
    Dim shapes As Collection, s As Object, msg As String, entries As String, n As Long, nNew As Long, nBad As Long, r As Variant
    On Error GoTo Fail
    If what <> "check" And rows.Count = 0 Then PptLinks_Do = PPTLINKS_MSG_SELECT: Exit Function
    Select Case what
        Case "check"
            msg = PptTqq_LinkStates(pres, "", True, True)
            If Left$(msg, 1) = "!" Then PptLinks_Do = Mid$(msg, 2): Exit Function
            n = PptLinks_Remember(msg)
            For Each r In PptLinks_Collect(pres)
                If r(5) = "new" Then nNew = nNew + 1
                If r(5) = "missing" Or r(5) = "nosource" Or r(5) = "unlinked" Then nBad = nBad + 1
            Next
            PptLinks_Do = "Verificados: " & n & ". Com dados novos: " & nNew & ". Com problema: " & nBad & "."
            Exit Function
    End Select
    Set shapes = PptLinks_Shapes(pres, rows)
    If shapes.Count = 0 Then PptLinks_Do = PPTLINKS_MSG_SELECT: Exit Function
    Select Case what
        Case "update"
            For Each s In shapes
                PptTqq_EnsureUniqueNamePublic s
            Next
            entries = PptTqq_EntriesOf(shapes)
            msg = PptTqq_SyncEntriesMsg(pres, entries, True)
            PptLinks_Remember PptTqq_LinkStates(pres, entries, False)
            If Left$(msg, 1) = "!" Then msg = Mid$(msg, 2)
            PptLinks_Do = Replace(msg, vbLf, " ")
        Case "auto"
            PptTqq_SetAutoShapes shapes, True
            msg = PptTqq_SyncEntriesMsg(pres, PptTqq_EntriesOf(shapes), True, True)
            If Left$(msg, 1) = "!" Then msg = Mid$(msg, 2)
            PptLinks_Do = "Automático: " & shapes.Count & " elemento(s). " & Replace(msg, vbLf, " ")
        Case "manual"
            PptTqq_SetAutoShapes shapes, False
            PptLinks_Do = "Manual: " & shapes.Count & " elemento(s)."
        Case "unlink"
            If PptTqq_Ask("Desvincular " & shapes.Count & " elemento(s)? O ThinQQ deixa de atualizá-los (eles ficam no slide como estão).") <> vbYes Then
                If BUILD_CONFIG <> "Test" Then Exit Function
            End If
            PptLinks_Do = "Desvinculados: " & PptLinks_Unlink(shapes) & "."
        Case "edit"
            PptTqq_EditShapes pres, shapes
        Case "goto"
            PptLinks_Goto shapes(1)
    End Select
    PptRx_InvalidateElem
    Exit Function
Fail:
    PptLinks_Do = "Não foi possível concluir: " & Err.Description & " (" & Err.Number & ")"
End Function

' Mostra o slide da forma na janela ativa e seleciona a forma. Nunca propaga.
Private Sub PptLinks_Goto(s As Object)
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then Exit Sub
    Application.ActiveWindow.View.GotoSlide CLng(s.Parent.SlideIndex)
    s.Select
    Err.Clear
End Sub

' Abre (ou atualiza) o dialogo para a apresentacao ativa. Nunca propaga erro.
Public Sub PptLinks_Show()
    On Error GoTo Fail
    If Application.Presentations.Count = 0 Then PptTqq_Msg PPT_MSG_NO_PRES: Exit Sub
    If BUILD_CONFIG = "Test" Then Exit Sub
    frmTqqLinks.Prepare Application.ActivePresentation
    frmTqqLinks.Show vbModeless
    Exit Sub
Fail:
    PptTqq_Msg "Não foi possível abrir os Vínculos: " & Err.Description & " (" & Err.Number & ")", vbExclamation
End Sub

' Apresentacao fechada: o dialogo dela sai (nao fica apontando para um objeto morto). Nunca propaga erro.
Public Sub PptLinks_OnClose(pres As Object)
    Dim f As Object
    On Error Resume Next
    For Each f In VBA.UserForms
        If f.Name = "frmTqqLinks" Then
            If LCase$(CStr(f.Pres.FullName)) = LCase$(CStr(pres.FullName)) Then Unload f
        End If
    Next
    Err.Clear
End Sub
```
  Em `modPptTqq.bas`: acrescente `Public Function PptTqq_LinkSourceOf(shp As Object) As String` (chama a privada `PptTqq_LinkSource`), `Public Sub
  PptTqq_EnsureUniqueNamePublic(shp As Object)` (chama `PptTqq_EnsureUniqueName`), e

```vb
' Opcao do core (API 2: "autoon"/"autonew") no Excel JA aberto; available = False sem Excel, sem o core ou com o core antigo. Nunca propaga.
Public Function PptTqq_CoreFlag(key As String, ByRef available As Boolean) As Boolean
    Dim xl As Object, created As Boolean, core As String
    available = False
    On Error GoTo Done
    Set xl = PptTqq_GetExcel(created, True)
    If xl Is Nothing Then Exit Function
    core = PptTqq_EnsureCore(xl, created)
    If Len(core) > 0 Then
        If PptTqq_CoreApi(xl, core) >= PPT_API_SHAPES Then
            PptTqq_CoreFlag = CBool(xl.Run("'" & core & "'!Core_PptGetFlag", key))
            available = True
        End If
    End If
Done:
    On Error Resume Next
    PptTqq_Release xl, created
End Function

' Grava a opcao no core (Excel ja aberto). Devolve a linha de status. Nunca propaga.
Public Function PptTqq_SetCoreFlag(key As String, v As Boolean) As String
    Dim xl As Object, created As Boolean, core As String
    On Error GoTo Fail
    Set xl = PptTqq_GetExcel(created, True)
    If xl Is Nothing Then PptTqq_SetCoreFlag = PPT_MSG_NOCORE_STATES: Exit Function
    core = PptTqq_EnsureCore(xl, created)
    If Len(core) = 0 Then
        PptTqq_SetCoreFlag = PPT_MSG_NOCORE_STATES
    ElseIf PptTqq_CoreApi(xl, core) < PPT_API_SHAPES Then
        PptTqq_SetCoreFlag = PPT_MSG_OLD_CORE_STATES
    Else
        xl.Run "'" & core & "'!Core_PptSetFlag", key, v
        PptTqq_SetCoreFlag = "Opção gravada no Excel."
    End If
    PptTqq_Release xl, created
    Exit Function
Fail:
    On Error Resume Next
    PptTqq_Release xl, created
    PptTqq_SetCoreFlag = PptTqq_ErrorText(Err.Number, Err.Description)
End Function
```
  e, no início de `PptTqq_OnClose`, `PptLinks_OnClose pres`. Em `TestXlFake.Run`, acrescente `Case "Core_PptSetFlag": Run = Empty` (o `Core_PptGetFlag` já existe).

- [ ] **Passo 6: faixa** — ícone `tqq_links = @{ Mso = 'HyperlinkInsert'; Rects = @((New-TqqRect 1 2 6 5 $n), (New-TqqRect 9 9 6 5 $b), (New-TqqRect 6 4 4 2 $g), (New-TqqRect 6 10 4 2 $g), (New-TqqRect 7 5 2 6 $g)) }`
  em `Get-TqqIconSpecs`; `customUI14.xml` do `.ppam`: grupo `grpPptSync` com `label="Dados do Excel"` e, como 1º controle,
  `<button id="btnPptLinks" label="Vínculos" size="large" image="tqq_links" onAction="PptRx_OnLinks" screentip="Vínculos do ThinQQ" supertip="Lista os gráficos, textos e tabelas do ThinQQ vinculados ao Excel nesta apresentação, com o estado de cada um (dados novos, atualizado, automático, pasta não encontrada), e atualiza, liga o automático ou desvincula em lote (o Data Links do think-cell). Também tem as opções de acompanhamento e do que fazer ao abrir."/>`;
  `btnPptSyncNow` passa a `label="Atualizar tudo"` (supertip: "Traz agora os dados do Excel para todos os elementos do ThinQQ desta apresentação. Sem caixa quando tudo dá certo.");
  **remova** o `dropDown` `ddOpenMode` (a escolha foi para o diálogo). Na aba contextual, grupo `cgrpLink`, acrescente
  `<button id="cbtnLinks" label="Vínculos" image="tqq_links" onAction="PptRx_OnLinks" screentip="Vínculos do ThinQQ"/>`. Em `modPptRibbon.bas`:

```vb
' Vinculos (1.11.0): o dialogo nao modal da apresentacao ativa.
Public Sub PptRx_OnLinks(control As IRibbonControl)
    On Error Resume Next
    PptLinks_Show
End Sub
```
  `Test-Ribbon.ps1`: troque a checagem dos itens de `ddOpenMode` por `Check ($null -eq $px.SelectSingleNode("//u:dropDown[@id='ddOpenMode']", $pns)) 'ppt: Ao abrir saiu da faixa (Vínculos > Opções)'`
  e acrescente a de `btnPptLinks` (1º controle de `grpPptSync`, `image='tqq_links'`, `onAction='PptRx_OnLinks'`). `Test-Docs.ps1`: os pinos de "Ao abrir"
  (~linhas 46, 78 e seguintes, 503) passam a procurar o texto no diálogo (README/tutorial dizem "Vínculos > Ao abrir a apresentação"); se o
  conjunto de rótulos conhecidos for montado a partir do XML, inclua as legendas do `frmTqqLinks.json` nele (liste no relatório).

- [ ] **Passo 7:** suíte com `-Ppt` → 0; commits `build: alvo ppt importa formularios`, `feat: dialogo Vinculos (Data Links) no PowerPoint`,
  `test: Vinculos`.

---

### Tarefa U4: Excel — menu PowerPoint sem as opções do automático — modelo: standard

**Files:** Modify `src/loader/ribbon/customUI14.xml`, `tests/vba/loader/TestRibbon.bas`, `tests/ps/Test-Docs.ps1` (pinos de `chkPptAuto*`).

- [ ] **Passo 1 (RED):** em `TestRibbon.bas` troque as asserções das linhas 889–910 e 970–973 (em 3a30f38) por

```vb
    AssertTrue d.SelectSingleNode("//*[@id='chkPptAutoOn']") Is Nothing, "1.11.0: opcoes do automatico no dialogo Vinculos do PowerPoint"
    AssertTrue d.SelectSingleNode("//*[@id='chkPptAutoNew']") Is Nothing
    AssertTrue d.SelectSingleNode("//*[@id='msepPptAuto']") Is Nothing
    AssertEqual "btnPptFrames", AttrText(d.SelectSingleNode("//u:button[@id='btnPptSync']/following-sibling::*[1]"), "id")
    AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='btnPptSync']"), "supertip"), "Vínculos") > 0, "aponta para o dialogo do PowerPoint"
```
  (mantenha os testes de `Rx_PptFlagValue`/`Rx_OnPptFlagCore`: os callbacks continuam existindo para um `.xlam` antigo não quebrar; a lista
  `HasAllCallbacks` não muda se os nomes continuam presentes no módulo). Rode `-Target loader -Filter TestRibbon` → falha.
- [ ] **Passo 2:** no `customUI14.xml` do loader, remova `msepPptAuto`, `chkPptAutoOn` e `chkPptAutoNew`; `btnPptSync`:
  `supertip="Traz agora os dados do Excel para os gráficos, textos e tabelas do ThinQQ das apresentações abertas no PowerPoint. Como no think-cell, nada muda sozinho: no PowerPoint, o elemento com dados novos mostra 'Dados novos: atualizar'; a lista de todos, as opções de atualização automática e o que fazer ao abrir ficam em ThinQQ > Vínculos (no PowerPoint)."`;
  `mnuPpt` supertip: troque "e controla a atualização automática" por "(as opções de atualização ficam em Vínculos, no PowerPoint)". Ajuste os pinos
  do `Test-Docs.ps1` ~linhas 885–888 (`$xlOnNode`, `$xlNewNode`): passam a exigir `$null -eq` (os controles saíram) e o README/tutorial a citar
  "Vínculos" (a U5 escreve o texto; aqui só o pino).
- [ ] **Passo 3:** suíte `-Target all -LoaderSwap` → 0 (Test-Ribbon e Test-Docs incluídos); commit `feat: faixa do Excel sem as opcoes do automatico (foram para Vinculos)`.

---

### Tarefa U5: Documentação da 1.11.0 — modelo: standard

**Files:** `docs/tutorial/tutorial.html`, `README.md`, `CHANGELOG.md`, `release/notes.txt`, `tests/ps/Test-Docs.ps1`.

- [ ] **Passo 1:** meça o tutorial (`(Get-Item docs\tutorial\tutorial.html).Length`; hoje 607.118 B). Reescreva a subseção `#ppt-auto` (mantenha a
  âncora; título "Atualizar como o think-cell: manual, bandeira e automático") e crie `#ppt-vinculos` ("Vínculos (o Data Links)"), substituindo o
  texto da 1.6.0 sobre "atualiza sozinho ~2 s depois" por: tudo nasce manual; a bandeira "Dados novos: atualizar" (com a figura do ícone em CSS, sem
  PNG novo); Automático por elemento; o diálogo (colunas, filtros, botões, opções, "Ao abrir"); a aba contextual ThinQQ; silêncio no sucesso;
  migração ("decks da 1.6–1.10 continuam como estavam; para voltar tudo ao manual: Vínculos > Selecionar tudo > Manual"); limites (bandeira aparece ao
  selecionar o elemento, não sobre o gráfico; "Desvincular" não quebra o vínculo nativo; sem Revert: use Ctrl+Z). Atualize `#ppt-paridade` (linhas da
  matriz §9 da spec marcadas 1.11.0), `#ppt-mensagens` (as mensagens novas: `PPT_MSG_OLD_CORE_STATES`, `PPT_MSG_NOCORE_STATES`,
  `PPTLINKS_MSG_SELECT`, "Desvincular N elemento(s)? ...", "Verificados: ...", "Não foi possível verificar os vínculos: ...") e a referência da faixa
  (Excel: menu PowerPoint sem as caixas; PowerPoint: grupos "Dados do Excel", "Elemento selecionado", aba contextual). Condense o que a 1.6.0
  descrevia e ficou obsoleto para caber no teto (655.360 B).
- [ ] **Passo 2:** README: "Recursos" (Vínculos, bandeira, manual por padrão), "Arquitetura" (`modPptState`, `TQQ_STATE`, `TQQ_LINKIDX`, API 6,
  `modPptLinks`/`frmTqqLinks`, formulários no build do `.ppam`), "Verificações manuais" itens **110–118** (texto da spec §14) e "Limitações".
- [ ] **Passo 3:** `CHANGELOG.md` `[Não lançado]`: "Adicionado" (bandeira, Vínculos, aba contextual, API 6), "Alterado" (tudo nasce manual, inclusive
  os criados pelo PowerPoint; abrir não pergunta por padrão; sucesso sem caixa; opções do automático foram do Excel para Vínculos; o Desfazer do Excel
  não é mais limpo com elementos manuais), "Não verificado em PowerPoint real" (P-U1/P-U3 se não rodaram). `release/notes.txt` reescrito (5–6
  marcadores, em linguagem do usuário).
- [ ] **Passo 4:** `Test-Docs.ps1`: rótulos novos da faixa e do diálogo, âncoras `#ppt-auto`/`#ppt-vinculos`, mensagens novas nos lugares certos,
  itens 110–118; os blocos do CHANGELOG lidos de `[Não lançado]` com recuo para `[1.11.0]` (padrão das versões anteriores). Rode com pwsh 7 e 5.1 (e
  `Set-StrictMode`) → 0. Commit `docs: 1.11.0 (manual com bandeira, Vinculos, aba contextual)`.

---

### Tarefa U6: Suíte final, revisão do ramo, verificação manual — modelo: standard (controlador + humano)

- [ ] Controlador: com o PowerPoint do usuário **fechado** (combinado com ele) e sem LogonUI, suíte `-Target all -LoaderSwap -Ppt -TimeoutSec 1500` → 0
  e 0 testes de PowerPoint pulados; registrar duração e memória livre.
- [ ] Revisão final do ramo (most-capable, somente leitura) sobre `main..HEAD`: Review Focus 1–4 e 7; segurança (Verificar só abre somente leitura;
  nada salva pasta do usuário; tags de deck alheio); onda única de correções; suíte de novo.
- [ ] Verificação manual (usuário, itens 110–118 da spec §14; **P-U3**: a aba ThinQQ continua aparecendo depois de instalar o `.ppam` — se sumir,
  remova o bloco `contextualTabs` num commit e registre). Ledger e push de `feat/ux-vinculos`.

# PARTE S — Estilos de gráfico (1.12.0)

Spec §8.5. Ordem interna: S1 (modelo puro) → S2 (depósitos + API de nomes) → S3 (aplicar por papel + sobrevivência) → S4 (janela e faixa do Excel)
→ S5 (PowerPoint) → S6 (docs) → S7 (suíte). Todos os testes do core novos em `tests/vba/core/TestStyles.bas` (novo).

### Tarefa S1: Modelo puro do estilo (`modStyleDef`) — modelo: standard

**Files:** Create `src/core/modules/modStyleDef.bas`, `tests/vba/core/TestStyles.bas`.

**Interfaces (Produces):** `STYLEDEF_HEADER = "ThinQQ-Estilo=1"`, `STYLEDEF_MAX_BYTES = 16384`, `STYLEDEF_MAX_LINES = 200`, `STYLEDEF_EXT = ".tqqestilo"`;
`StyleDef_Validate(key As String, value As String) As String` (valor normalizado; "" = inválido); `StyleDef_Parse(text As String, ByRef ignored As Long) As
Object` (Dictionary ou Nothing); `StyleDef_Serialize(d As Object) As String`; `StyleDef_Hex(v) As String`; `StyleDef_ColorOf(hexValue) As Long` (-1 se
inválido); `StyleDef_HexOf(rgbValue As Long) As String`; `StyleDef_PatternOf(nm) As Long` (0 = sem hachura); `StyleDef_PatternName(pattern As Long) As
String`; `StyleDef_SafeName(s) As String`; `StyleDef_Get(d, key) As String`; `StyleDef_SeriesCount(d) As Long`; `StyleDef_Version(text) As String` (8 hex);
`StyleDef_Keys() As Variant`; `StyleDef_BuiltinNames() As Variant`; `StyleDef_IsBuiltin(name) As Boolean`; `StyleDef_BuiltinPalette(name) As String`;
`StyleDef_Builtin(name) As String`.

- [ ] **Passo 1: Testes (RED)** — `tests/vba/core/TestStyles.bas`:

```vb
Attribute VB_Name = "TestStyles"
Option Explicit
' Estilos de grafico (1.12.0; spec 2026-10-23 secao 8.5): modelo puro, depositos, aplicar por papel e sobrevivencia.

Public Sub Test_StyleDef_Validate_Rules()
    AssertEqual "#70AD47", StyleDef_Validate("aumento", " #70ad47 ")
    AssertEqual "", StyleDef_Validate("aumento", "#70AD4"), "hex curto"
    AssertEqual "", StyleDef_Validate("aumento", "70AD47"), "sem #"
    AssertEqual "", StyleDef_Validate("aumento", "#GGGGGG")
    AssertEqual "auto", StyleDef_Validate("rotulo.cor", "AUTO")
    AssertEqual "#FFFFFF", StyleDef_Validate("rotulo.cor", "#ffffff")
    AssertEqual "nenhuma", StyleDef_Validate("primeira.borda", "nenhuma")
    AssertEqual "diagonal_larga_para_cima", StyleDef_Validate("primeira.padrao", "Diagonal_Larga_Para_Cima")
    AssertEqual "", StyleDef_Validate("primeira.padrao", "nenhum,xadrez_pequeno"), "virgula nunca casa dois itens"
    AssertEqual "", StyleDef_Validate("primeira.padrao", "qualquer")
    AssertEqual "9", StyleDef_Validate("fonte.tamanho", "9")
    AssertEqual "", StyleDef_Validate("fonte.tamanho", "5")
    AssertEqual "", StyleDef_Validate("fonte.tamanho", "9.5")
    AssertEqual "", StyleDef_Validate("espaco", "-1")
    AssertEqual "500", StyleDef_Validate("espaco", "500")
    AssertEqual "0.75", StyleDef_Validate("conector.espessura", "0,75")
    AssertEqual "", StyleDef_Validate("conector.espessura", "3.5")
    AssertEqual "", StyleDef_Validate("conector.espessura", "1.2.3")
    AssertEqual "pontos", StyleDef_Validate("conector.traco", "Pontos")
    AssertEqual "", StyleDef_Validate("conector.traco", "zigue")
    AssertEqual "1", StyleDef_Validate("rotulo.negrito", "1")
    AssertEqual "", StyleDef_Validate("rotulo.negrito", "sim")
    AssertEqual "Calibri Light", StyleDef_Validate("fonte", "Calibri Light")
    AssertEqual "", StyleDef_Validate("fonte", "Calibri;=x")
    AssertEqual "", StyleDef_Validate("desconhecida", "#000000")
    AssertEqual "", StyleDef_Validate("serie9", "#000000"), "so serie1..serie8"
End Sub

Public Sub Test_StyleDef_SafeName()
    AssertEqual "Cascata MC", StyleDef_SafeName("  Cascata MC  ")
    AssertEqual "ab", StyleDef_SafeName("a;=|""\/:*?<>&b")
    AssertEqual "Padrão ThinQQ", StyleDef_SafeName("Padrão ThinQQ")
    AssertEqual 40, Len(StyleDef_SafeName(String$(60, "x")))
    AssertEqual "", StyleDef_SafeName(vbTab & vbCr)
End Sub

Public Sub Test_StyleDef_ParseSerialize_RoundTrip()
    Dim d As Object, ign As Long, t As String, d2 As Object
    t = ChrW(65279) & STYLEDEF_HEADER & vbLf & "# comentario" & vbCrLf & "nome=Meu" & vbLf & "AUMENTO=#70ad47" & vbLf & "reducao=vermelho" & vbLf & "lixo" & vbLf & "outra=1" & vbLf & vbLf & "espaco=40"
    Set d = StyleDef_Parse(t, ign)
    AssertTrue Not d Is Nothing
    AssertEqual 3, ign, "reducao invalida, linha sem '=', chave desconhecida"
    AssertEqual "#70AD47", StyleDef_Get(d, "aumento")
    AssertEqual "40", StyleDef_Get(d, "espaco")
    AssertEqual "", StyleDef_Get(d, "reducao")
    t = StyleDef_Serialize(d)
    AssertEqual STYLEDEF_HEADER & vbCrLf & "nome=Meu" & vbCrLf & "aumento=#70AD47" & vbCrLf & "espaco=40", t, "ordem canonica"
    Set d2 = StyleDef_Parse(t, ign)
    AssertEqual 0, ign
    AssertEqual t, StyleDef_Serialize(d2), "ida e volta"
    AssertEqual 8, Len(StyleDef_Version(t))
    AssertTrue StyleDef_Version(t) <> StyleDef_Version(t & vbCrLf & "total=#000000")
End Sub

Public Sub Test_StyleDef_Parse_RefusesHostile()
    Dim ign As Long, i As Long, s As String
    AssertTrue StyleDef_Parse("", ign) Is Nothing
    AssertTrue StyleDef_Parse("nome=x", ign) Is Nothing, "sem cabecalho"
    AssertTrue StyleDef_Parse(STYLEDEF_HEADER & vbLf & String$(STYLEDEF_MAX_BYTES, "a"), ign) Is Nothing, "grande demais"
    s = STYLEDEF_HEADER
    For i = 1 To STYLEDEF_MAX_LINES + 1
        s = s & vbLf & "x=1"
    Next
    AssertTrue StyleDef_Parse(s, ign) Is Nothing, "linhas demais"
    AssertTrue Not StyleDef_Parse(STYLEDEF_HEADER & vbLf & "nome==<script>" & vbLf & "fonte=a&b", ign) Is Nothing
    AssertEqual 2, ign, "nome e fonte invalidos ignorados, nada executado"
End Sub

Public Sub Test_StyleDef_ColorsPatternsAndBuiltins()
    Dim n As Variant, d As Object, ign As Long
    AssertEqual RGB(112, 173, 71), StyleDef_ColorOf("#70AD47")
    AssertEqual -1, StyleDef_ColorOf("x")
    AssertEqual "#70AD47", StyleDef_HexOf(RGB(112, 173, 71))
    AssertEqual msoPatternWideUpwardDiagonal, StyleDef_PatternOf("diagonal_larga_para_cima")
    AssertEqual 0, StyleDef_PatternOf("nenhum")
    AssertEqual "diagonal_escura_para_cima", StyleDef_PatternName(msoPatternDarkUpwardDiagonal)
    AssertEqual "", StyleDef_PatternName(999)
    AssertEqual 5, UBound(StyleDef_BuiltinNames()) + 1
    For Each n In StyleDef_BuiltinNames()
        Set d = StyleDef_Parse(StyleDef_Builtin(CStr(n)), ign)
        AssertTrue Not d Is Nothing, CStr(n)
        AssertEqual 0, ign, "embutido sem linha invalida: " & CStr(n)
        AssertEqual CStr(n), StyleDef_Get(d, "nome")
        AssertTrue StyleDef_IsBuiltin(CStr(n))
    Next
    AssertEqual "padrao", StyleDef_BuiltinPalette("Padrão ThinQQ")
    AssertEqual "", StyleDef_BuiltinPalette("Cascata MC")
    Set d = StyleDef_Parse(StyleDef_Builtin("Cascata MC"), ign)
    AssertEqual "#70AD47", StyleDef_Get(d, "aumento")
    AssertEqual "#E8173C", StyleDef_Get(d, "reducao")
    AssertEqual "#D9D9D9", StyleDef_Get(d, "total")
    AssertEqual "diagonal_larga_para_cima", StyleDef_Get(d, "primeira.padrao")
    AssertEqual 3, StyleDef_SeriesCount(d)
    AssertFalse StyleDef_IsBuiltin("Meu")
End Sub
```

- [ ] **Passo 2:** `-Target core -Filter TestStyles` → saída 3 (RED).
- [ ] **Passo 3:** `src/core/modules/modStyleDef.bas`:

```vb
Attribute VB_Name = "modStyleDef"
Option Explicit
' Estilos de grafico (1.12.0; spec 2026-10-23 secao 8.5): modelo PURO. Texto "chave=valor" (uma linha por chave, a primeira util =
' STYLEDEF_HEADER; linhas vazias e comecando por "#" ignoradas; BOM ignorado), validado chave a chave contra listas fechadas: valor invalido ou
' chave desconhecida sao ignorados e contados. Nada e executado (so Split, Like e comparacoes). Teto: STYLEDEF_MAX_BYTES caracteres e
' STYLEDEF_MAX_LINES linhas (texto maior e recusado inteiro).
Public Const STYLEDEF_HEADER As String = "ThinQQ-Estilo=1"
Public Const STYLEDEF_MAX_BYTES As Long = 16384
Public Const STYLEDEF_MAX_LINES As Long = 200
Public Const STYLEDEF_EXT As String = ".tqqestilo"
Private Const STYLEDEF_KEYS As String = ",nome,fonte,fonte.tamanho,serie1,serie2,serie3,serie4,serie5,serie6,serie7,serie8,total,aumento,reducao,primeira,primeira.padrao,primeira.padrao.cor,primeira.borda,ultima,ultima.padrao,ultima.padrao.cor,ultima.borda,rotulo.cor,rotulo.negrito,rotulo.tamanho,conector.cor,conector.traco,conector.espessura,seta.cor,seta.rotulo.fundo,seta.rotulo.borda,eixo.categoria.linha,espaco,"
Private Const STYLEDEF_PATTERNS As String = ",nenhum,diagonal_clara_para_cima,diagonal_escura_para_cima,diagonal_larga_para_cima,diagonal_clara_para_baixo,diagonal_escura_para_baixo,diagonal_larga_para_baixo,horizontal_escura,vertical_escura,xadrez_pequeno,pontilhado_50,grade_pequena,"

Public Function StyleDef_Keys() As Variant
    StyleDef_Keys = Split(Mid$(STYLEDEF_KEYS, 2, Len(STYLEDEF_KEYS) - 2), ",")
End Function

Public Function StyleDef_Hex(v As String) As String
    Dim s As String
    s = UCase$(Trim$(v))
    If Len(s) = 7 And Left$(s, 1) = "#" Then
        If Not Mid$(s, 2) Like "*[!0-9A-F]*" Then StyleDef_Hex = s
    End If
End Function

Public Function StyleDef_ColorOf(hexValue As String) As Long
    Dim s As String
    s = StyleDef_Hex(hexValue)
    If Len(s) = 0 Then StyleDef_ColorOf = -1: Exit Function
    StyleDef_ColorOf = RGB(CLng("&H" & Mid$(s, 2, 2)), CLng("&H" & Mid$(s, 4, 2)), CLng("&H" & Mid$(s, 6, 2)))
End Function

Public Function StyleDef_HexOf(rgbValue As Long) As String
    StyleDef_HexOf = "#" & Right$("0" & Hex$(rgbValue And &HFF&), 2) & Right$("0" & Hex$((rgbValue \ &H100&) And &HFF&), 2) & Right$("0" & Hex$((rgbValue \ &H10000) And &HFF&), 2)
End Function

Private Function StyleDef_IntIn(v As String, lo As Long, hi As Long) As Boolean
    If Len(v) = 0 Or Len(v) > 4 Then Exit Function
    If v Like "*[!0-9]*" Then Exit Function
    StyleDef_IntIn = (CLng(v) >= lo And CLng(v) <= hi)
End Function

' Numero com ponto ou virgula (no maximo um separador e 6 caracteres); -1 se invalido. Pura.
Private Function StyleDef_Num(v As String) As Double
    Dim s As String
    s = Replace(v, ",", ".")
    StyleDef_Num = -1
    If Len(s) = 0 Or Len(s) > 6 Then Exit Function
    If s Like "*[!0-9.]*" Then Exit Function
    If Len(s) - Len(Replace(s, ".", "")) > 1 Then Exit Function
    StyleDef_Num = Val(s)
End Function

' Nome de estilo seguro (arquivo, opcao do grafico, XML): sem ; = | " \ / : * ? < > & e sem controles; espacos das pontas fora; no maximo 40. Pura.
Public Function StyleDef_SafeName(s As String) As String
    Dim i As Long, ch As String, t As String, c As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1): c = AscW(ch)
        If (c >= 32 Or c < 0) And c <> 127 And InStr(";=|""\/:*?<>&", ch) = 0 Then t = t & ch
    Next
    t = Trim$(t)
    If Len(t) > 40 Then t = Trim$(Left$(t, 40))
    StyleDef_SafeName = t
End Function

' Valor normalizado da chave ou "" (invalido/chave desconhecida). Pura.
Public Function StyleDef_Validate(key As String, value As String) As String
    Dim v As String, d As Double
    v = Trim$(value)
    If InStr(key, ",") > 0 Or InStr(STYLEDEF_KEYS, "," & key & ",") = 0 Or Len(key) = 0 Then Exit Function
    Select Case key
        Case "nome"
            If Len(v) > 0 And StyleDef_SafeName(v) = v Then StyleDef_Validate = v
        Case "fonte"
            If Len(v) > 0 And Len(v) <= 40 Then
                If Not v Like "*[!A-Za-z0-9 -]*" Then StyleDef_Validate = v
            End If
        Case "fonte.tamanho", "rotulo.tamanho"
            If StyleDef_IntIn(v, 6, 24) Then StyleDef_Validate = CStr(CLng(v))
        Case "espaco"
            If StyleDef_IntIn(v, 0, 500) Then StyleDef_Validate = CStr(CLng(v))
        Case "rotulo.negrito", "eixo.categoria.linha"
            If v = "0" Or v = "1" Then StyleDef_Validate = v
        Case "primeira.padrao", "ultima.padrao"
            If Len(v) > 0 And InStr(v, ",") = 0 Then
                If InStr(STYLEDEF_PATTERNS, "," & LCase$(v) & ",") > 0 Then StyleDef_Validate = LCase$(v)
            End If
        Case "conector.traco"
            Select Case LCase$(v)
                Case "solido", "pontos", "tracos": StyleDef_Validate = LCase$(v)
            End Select
        Case "conector.espessura"
            d = StyleDef_Num(v)
            If d >= 0.25 And d <= 3 Then StyleDef_Validate = Replace(Format$(d, "0.00"), ",", ".")
        Case "rotulo.cor"
            If LCase$(v) = "auto" Then StyleDef_Validate = "auto" Else StyleDef_Validate = StyleDef_Hex(v)
        Case "primeira.borda", "ultima.borda", "seta.rotulo.borda"
            If LCase$(v) = "nenhuma" Then StyleDef_Validate = "nenhuma" Else StyleDef_Validate = StyleDef_Hex(v)
        Case Else
            StyleDef_Validate = StyleDef_Hex(v)                ' as chaves restantes da lista sao cores
    End Select
End Function

Public Function StyleDef_Get(d As Object, key As String) As String
    If d Is Nothing Then Exit Function
    If d.Exists(key) Then StyleDef_Get = CStr(d(key))
End Function

' Quantas cores de serie seguidas (serie1, serie2, ...) o estilo define. Pura.
Public Function StyleDef_SeriesCount(d As Object) As Long
    Dim i As Long
    For i = 1 To 8
        If Len(StyleDef_Get(d, "serie" & i)) = 0 Then Exit For
        StyleDef_SeriesCount = i
    Next
End Function

' Texto -> Dictionary (chave -> valor normalizado) das chaves validas; ignored = linhas desconhecidas ou invalidas. Nothing se o texto e
' recusado inteiro (vazio, maior que o teto, linhas demais, sem o cabecalho como 1a linha util). Pura.
Public Function StyleDef_Parse(text As String, ByRef ignored As Long) As Object
    Dim d As Object, lines() As String, i As Long, ln As String, p As Long, k As String, v As String, gotHeader As Boolean
    ignored = 0
    If Len(text) = 0 Or Len(text) > STYLEDEF_MAX_BYTES Then Exit Function
    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    If UBound(lines) + 1 > STYLEDEF_MAX_LINES Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Left$(ln, 1) = ChrW(65279) Then ln = Mid$(ln, 2)
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            If Not gotHeader Then
                If ln <> STYLEDEF_HEADER Then Exit Function
                gotHeader = True
            Else
                p = InStr(ln, "=")
                If p > 1 Then
                    k = LCase$(Trim$(Left$(ln, p - 1)))
                    v = StyleDef_Validate(k, Mid$(ln, p + 1))
                    If Len(v) > 0 Then d(k) = v Else ignored = ignored + 1
                Else
                    ignored = ignored + 1
                End If
            End If
        End If
    Next
    If gotHeader Then Set StyleDef_Parse = d
End Function

' Dictionary -> texto canonico: cabecalho e as chaves validas na ordem de STYLEDEF_KEYS, linhas por vbCrLf. Pura.
Public Function StyleDef_Serialize(d As Object) As String
    Dim k As Variant, s As String, v As String
    s = STYLEDEF_HEADER
    For Each k In StyleDef_Keys()
        If d.Exists(CStr(k)) Then
            v = StyleDef_Validate(CStr(k), CStr(d(CStr(k))))
            If Len(v) > 0 Then s = s & vbCrLf & CStr(k) & "=" & v
        End If
    Next
    StyleDef_Serialize = s
End Function

' Versao curta do estilo: 8 primeiros hex do SHA-256 do texto canonico. Pura.
Public Function StyleDef_Version(text As String) As String
    Dim d As Object, ign As Long, s As String, b() As Byte
    Set d = StyleDef_Parse(text, ign)
    If d Is Nothing Then s = text Else s = StyleDef_Serialize(d)
    b = StrConv(s, vbFromUnicode)
    StyleDef_Version = Left$(Sha256_Bytes(b, UBound(b) - LBound(b) + 1), 8)
End Function

Public Function StyleDef_PatternOf(nm As String) As Long
    Select Case LCase$(nm)
        Case "diagonal_clara_para_cima": StyleDef_PatternOf = msoPatternLightUpwardDiagonal
        Case "diagonal_escura_para_cima": StyleDef_PatternOf = msoPatternDarkUpwardDiagonal
        Case "diagonal_larga_para_cima": StyleDef_PatternOf = msoPatternWideUpwardDiagonal
        Case "diagonal_clara_para_baixo": StyleDef_PatternOf = msoPatternLightDownwardDiagonal
        Case "diagonal_escura_para_baixo": StyleDef_PatternOf = msoPatternDarkDownwardDiagonal
        Case "diagonal_larga_para_baixo": StyleDef_PatternOf = msoPatternWideDownwardDiagonal
        Case "horizontal_escura": StyleDef_PatternOf = msoPatternDarkHorizontal
        Case "vertical_escura": StyleDef_PatternOf = msoPatternDarkVertical
        Case "xadrez_pequeno": StyleDef_PatternOf = msoPatternSmallCheckerBoard
        Case "pontilhado_50": StyleDef_PatternOf = msoPattern50Percent
        Case "grade_pequena": StyleDef_PatternOf = msoPatternSmallGrid
    End Select
End Function

Public Function StyleDef_PatternName(pattern As Long) As String
    Dim k As Variant
    For Each k In Split(Mid$(STYLEDEF_PATTERNS, 2, Len(STYLEDEF_PATTERNS) - 2), ",")
        If StyleDef_PatternOf(CStr(k)) = pattern And pattern <> 0 Then StyleDef_PatternName = CStr(k): Exit Function
    Next
End Function

Public Function StyleDef_BuiltinNames() As Variant
    StyleDef_BuiltinNames = Array("Padrão ThinQQ", "Tons de cinza", "Cores do tema", "Cascata MC", "Cascata DDE")
End Function

Public Function StyleDef_IsBuiltin(name As String) As Boolean
    Dim n As Variant
    For Each n In StyleDef_BuiltinNames()
        If StrComp(CStr(n), name, vbTextCompare) = 0 Then StyleDef_IsBuiltin = True: Exit Function
    Next
End Function

' Paleta antiga (opcao "palette") dos embutidos que sao so paleta; "" para os demais. Pura.
Public Function StyleDef_BuiltinPalette(name As String) As String
    Select Case LCase$(name)
        Case LCase$("Padrão ThinQQ"): StyleDef_BuiltinPalette = "padrao"
        Case "tons de cinza": StyleDef_BuiltinPalette = "cinza"
        Case "cores do tema": StyleDef_BuiltinPalette = "tema"
    End Select
End Function

' Texto dos embutidos (validos por construcao: Test_StyleDef_ColorsPatternsAndBuiltins). "" para nome desconhecido. Pura.
Public Function StyleDef_Builtin(name As String) As String
    Dim s As String, nl As String
    nl = vbCrLf
    Select Case LCase$(name)
        Case LCase$("Padrão ThinQQ"): s = "nome=Padrão ThinQQ"
        Case "tons de cinza": s = "nome=Tons de cinza"
        Case "cores do tema": s = "nome=Cores do tema"
        Case "cascata mc"
            s = "nome=Cascata MC" & nl & "serie1=#70AD47" & nl & "serie2=#D9D9D9" & nl & "serie3=#E8173C" & nl & "total=#D9D9D9" & nl & "aumento=#70AD47" & nl & "reducao=#E8173C"
            s = s & nl & "primeira=#FFFFFF" & nl & "primeira.padrao=diagonal_larga_para_cima" & nl & "primeira.padrao.cor=#000000" & nl & "primeira.borda=#000000"
            s = s & nl & "ultima=#D9D9D9" & nl & "rotulo.cor=auto" & nl & "rotulo.negrito=1" & nl & "conector.cor=#7F7F7F" & nl & "conector.traco=pontos"
            s = s & nl & "conector.espessura=0.75" & nl & "seta.cor=#404040" & nl & "seta.rotulo.fundo=#FFFFFF" & nl & "seta.rotulo.borda=#1F3864" & nl & "eixo.categoria.linha=1"
        Case "cascata dde"
            s = "nome=Cascata DDE" & nl & "serie1=#4472C4" & nl & "serie2=#ED7D31" & nl & "serie3=#70AD47" & nl & "total=#4472C4" & nl & "aumento=#ED7D31" & nl & "reducao=#70AD47"
            s = s & nl & "primeira=#FFFFFF" & nl & "primeira.padrao=diagonal_escura_para_cima" & nl & "primeira.padrao.cor=#4472C4" & nl & "primeira.borda=#4472C4"
            s = s & nl & "ultima=#4472C4" & nl & "rotulo.cor=auto" & nl & "conector.cor=#A6A6A6" & nl & "conector.traco=solido" & nl & "conector.espessura=0.75"
        Case Else: Exit Function
    End Select
    StyleDef_Builtin = STYLEDEF_HEADER & nl & s
End Function
```
- [ ] **Passo 4:** `-Target core -Filter TestStyles` → verde; suíte `-Target all -LoaderSwap` → 0; commit `feat: modelo dos estilos de grafico (validacao, embutidos)`.

---

### Tarefa S2: Depósitos, importar/exportar, padrão para novos e API de nomes — modelo: standard

**Files:** Create `src/core/modules/modStyleStore.bas`; Modify `src/core/modules/modCoreApi.bas` (`Core_StyleNames`, `Core_StyleImport`, `Core_StyleExport`,
`Core_StyleSetDefault`, `Core_StyleDefault`); Test `tests/vba/core/TestStyles.bas`.

**Interfaces (Produces):** `STYLESTORE_NS = "urn:thinqq:styles:1"`; `StyleStore_SetStoreDirForTest(dir As String)`; `StyleStore_ResetTestHooks()`;
`StyleStore_UserDir() As String`; `StyleStore_ReadFile(path, ByRef msg) As String`; `StyleStore_WriteFile(path, text) As String` ("" ok);
`StyleStore_UserText(name) As String`; `StyleStore_UserNames() As Collection`; `StyleStore_SaveUser(name, text) As String`; `StyleStore_DeleteUser(name) As
String`; `StyleStore_BookAll(wb) As Object`; `StyleStore_BookText(wb, name) As String`; `StyleStore_SaveBook(wb, name, text)`; `StyleStore_EnsureInBook(wb,
name, text)`; `StyleStore_Get(name, Optional wb As Workbook, Optional preferUser As Boolean = False) As String`; `StyleStore_Names(Optional wb As Workbook) As
String` ("|"); `StyleStore_Import(path, ByRef name As String, ByRef ignored As Long) As String` (texto canônico ou "!" + mensagem); `StyleStore_Export(name, path,
Optional wb) As String`; `StyleStore_DefaultName() As String`; `StyleStore_SetDefaultName(n As String)`; `Core_StyleNames() As String`; `Core_StyleImport()`;
`Core_StyleExport(name As String)`; `Core_StyleSetDefault(name As String)`; `Core_StyleDefault() As String`.

- [ ] **Passo 1: Testes (RED)** — acrescente a `TestStyles.bas`:

```vb
Private Function StyleTestDir() As String
    Dim d As String
    d = TempDir() & "\styles"
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
    StyleStore_SetStoreDirForTest d
    StyleTestDir = d
End Function

Public Sub Test_StyleStore_UserSaveGetDelete_TestDirOnly()
    Dim d As String, t As String
    StyleStore_ResetTestHooks
    AssertEqual "", StyleStore_UserDir(), "build Test sem gancho: nunca a pasta real do usuario"
    AssertTrue Left$(StyleStore_SaveUser("X", STYLEDEF_HEADER), 1) = "!", "sem deposito: recusa"
    d = StyleTestDir()
    t = STYLEDEF_HEADER & vbCrLf & "nome=Meu Estilo" & vbCrLf & "aumento=#00FF00"
    AssertEqual "", StyleStore_SaveUser("Meu Estilo", t)
    AssertTrue Len(Dir(d & "\Meu Estilo" & STYLEDEF_EXT)) > 0
    AssertEqual t, StyleStore_UserText("meu estilo"), "sem diferenciar maiusculas"
    AssertTrue InStr("|" & StyleStore_Names() & "|", "|Meu Estilo|") > 0
    AssertTrue InStr("|" & StyleStore_Names() & "|", "|Cascata MC|") > 0, "embutidos na lista"
    AssertTrue Left$(StyleStore_SaveUser("Cascata MC", t), 1) = "!", "nome de embutido recusado"
    AssertEqual "", StyleStore_DeleteUser("Meu Estilo")
    AssertEqual "", StyleStore_UserText("Meu Estilo")
    StyleStore_ResetTestHooks
End Sub

Public Sub Test_StyleStore_Book_CustomXml_SaveReopen()
    Dim wb As Workbook, p As String, t As String
    Set wb = NewTempWorkbook()
    t = STYLEDEF_HEADER & vbCrLf & "nome=Da Pasta" & vbCrLf & "total=#123456"
    StyleStore_SaveBook wb, "Da Pasta", t
    StyleStore_SaveBook wb, "Outro <&>", STYLEDEF_HEADER
    AssertEqual t, Replace(StyleStore_BookText(wb, "da pasta"), vbLf, vbCrLf), "texto de volta (o XML normaliza CRLF)"
    AssertEqual 1, wb.CustomXMLParts.SelectByNamespace(STYLESTORE_NS).Count, "uma parte so"
    p = TempDir() & "\style-book.xlsx"
    wb.SaveAs p, xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Set wb = Workbooks.Open(p)
    AssertTrue InStr(StyleStore_BookText(wb, "Da Pasta"), "total=#123456") > 0, "sobrevive a salvar e reabrir"
    AssertEqual 2, StyleStore_BookAll(wb).Count
    AssertTrue InStr(StyleStore_Get("Da Pasta", wb), "#123456") > 0, "a pasta acha o estilo que o usuario nao tem"
    CloseTemp wb
    Kill p
End Sub

Public Sub Test_StyleStore_ImportExport_Validates()
    Dim d As String, p As String, nm As String, ign As Long, r As String, f As Integer
    StyleStore_ResetTestHooks
    d = StyleTestDir()
    p = d & "\entrada.tqqestilo"
    f = FreeFile
    Open p For Output As #f
    Print #f, STYLEDEF_HEADER
    Print #f, "nome=Recebido;x"
    Print #f, "aumento=#ABCDEF"
    Print #f, "comando=Shell(""calc"")"
    Close #f
    r = StyleStore_Import(p, nm, ign)
    AssertTrue Left$(r, 1) <> "!", r
    AssertEqual "entrada", nm, "nome invalido no arquivo: vale o nome do arquivo"
    AssertEqual 2, ign
    AssertTrue InStr(StyleStore_UserText("entrada"), "aumento=#ABCDEF") > 0
    AssertEqual "", StyleStore_Export("entrada", d & "\saida.tqqestilo")
    AssertTrue InStr(StyleStore_ReadFile(d & "\saida.tqqestilo", r), "aumento=#ABCDEF") > 0
    f = FreeFile
    Open d & "\ruim.tqqestilo" For Output As #f
    Print #f, "<xml>nada</xml>"
    Close #f
    AssertTrue Left$(StyleStore_Import(d & "\ruim.tqqestilo", nm, ign), 1) = "!", "sem cabecalho"
    StyleStore_ResetTestHooks
End Sub

Public Sub Test_StyleStore_Default_TestMemory()
    StyleStore_ResetTestHooks
    AssertEqual "", StyleStore_DefaultName()
    StyleStore_SetDefaultName "Cascata MC"
    AssertEqual "Cascata MC", StyleStore_DefaultName()
    StyleStore_ResetTestHooks
    AssertEqual "", StyleStore_DefaultName()
End Sub
```

- [ ] **Passo 2:** RED. **Passo 3:** `src/core/modules/modStyleStore.bas`:

```vb
Attribute VB_Name = "modStyleStore"
Option Explicit
' Depositos dos estilos (1.12.0; spec 2026-10-23 secao 8.5): embutidos (StyleDef_Builtin), do usuario (%LOCALAPPDATA%\ThinQQ\styles\<nome>
' .tqqestilo, UTF-8) e da pasta (CustomXMLParts no namespace STYLESTORE_NS: quem recebe a pasta ve o mesmo visual; medido PX1c). No build Test o
' deposito do usuario so existe com o gancho StyleStore_SetStoreDirForTest (nunca a pasta real) e o padrao para novos graficos fica em memoria.
Public Const STYLESTORE_NS As String = "urn:thinqq:styles:1"
Private Const STYLESTORE_SECTION As String = "Estilos"
Private mTestDir As String
Private mTestDefault As String

Public Sub StyleStore_SetStoreDirForTest(dir As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestDir = dir
End Sub

Public Sub StyleStore_ResetTestHooks()
    mTestDir = ""
    mTestDefault = ""
End Sub

Public Function StyleStore_UserDir() As String
    If BUILD_CONFIG = "Test" Then StyleStore_UserDir = mTestDir: Exit Function
    StyleStore_UserDir = Environ$("LOCALAPPDATA") & "\ThinQQ\styles"
End Function

' Cria a pasta do deposito do usuario se faltar (ThinQQ\styles). "" ok ou a mensagem. Nunca propaga.
Private Function StyleStore_EnsureDir() As String
    Dim d As String, parent As String
    On Error GoTo Fail
    d = StyleStore_UserDir()
    If Len(d) = 0 Then StyleStore_EnsureDir = "O depósito de estilos do usuário não está disponível.": Exit Function
    If Len(Dir(d, vbDirectory)) > 0 Then Exit Function
    parent = Left$(d, InStrRev(d, "\") - 1)
    If Len(Dir(parent, vbDirectory)) = 0 Then MkDir parent
    MkDir d
    Exit Function
Fail:
    StyleStore_EnsureDir = "Não foi possível criar a pasta de estilos: " & Err.Description
End Function

Public Function StyleStore_ReadFile(path As String, ByRef msg As String) As String
    Dim stm As Object
    On Error GoTo Fail
    msg = ""
    If FileLen(path) > STYLEDEF_MAX_BYTES * 3 Then msg = "O arquivo é grande demais para um estilo do ThinQQ.": Exit Function
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.LoadFromFile path
    StyleStore_ReadFile = stm.ReadText
    stm.Close
    If Left$(StyleStore_ReadFile, 1) = ChrW(65279) Then StyleStore_ReadFile = Mid$(StyleStore_ReadFile, 2)   ' BOM (o ADODB grava UTF-8 com BOM)
    Exit Function
Fail:
    msg = "Não foi possível ler o arquivo: " & Err.Description
End Function

Public Function StyleStore_WriteFile(path As String, text As String) As String
    Dim stm As Object
    On Error GoTo Fail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2                                          ' adSaveCreateOverWrite
    stm.Close
    Exit Function
Fail:
    StyleStore_WriteFile = "Não foi possível gravar o arquivo: " & Err.Description
End Function

Private Function StyleStore_UserPath(name As String) As String
    If Len(StyleStore_UserDir()) = 0 Or Len(StyleDef_SafeName(name)) = 0 Then Exit Function
    StyleStore_UserPath = StyleStore_UserDir() & "\" & StyleDef_SafeName(name) & STYLEDEF_EXT
End Function

Public Function StyleStore_UserText(name As String) As String
    Dim p As String, msg As String
    On Error Resume Next
    p = StyleStore_UserPath(name)
    If Len(p) = 0 Then Exit Function
    If Len(Dir(p)) = 0 Then Exit Function
    StyleStore_UserText = StyleStore_ReadFile(p, msg)
    Err.Clear
End Function

Public Function StyleStore_UserNames() As Collection
    Dim out As New Collection, f As String, d As String
    Set StyleStore_UserNames = out
    On Error Resume Next
    d = StyleStore_UserDir()
    If Len(d) = 0 Then Exit Function
    f = Dir(d & "\*" & STYLEDEF_EXT)
    Do While Len(f) > 0
        out.Add Left$(f, Len(f) - Len(STYLEDEF_EXT))
        f = Dir()
    Loop
    Err.Clear
End Function

' Grava o estilo do usuario (texto canonico, validado). Nome de embutido e recusado. "" ok ou "!" + mensagem.
Public Function StyleStore_SaveUser(name As String, text As String) As String
    Dim d As Object, ign As Long, msg As String, nm As String
    nm = StyleDef_SafeName(name)
    If Len(nm) = 0 Then StyleStore_SaveUser = "!Dê um nome ao estilo.": Exit Function
    If StyleDef_IsBuiltin(nm) Then StyleStore_SaveUser = "!'" & nm & "' é um estilo embutido do ThinQQ: use outro nome (Duplicar).": Exit Function
    Set d = StyleDef_Parse(text, ign)
    If d Is Nothing Then StyleStore_SaveUser = "!O estilo não é válido.": Exit Function
    d("nome") = nm
    msg = StyleStore_EnsureDir()
    If Len(msg) > 0 Then StyleStore_SaveUser = "!" & msg: Exit Function
    msg = StyleStore_WriteFile(StyleStore_UserPath(nm), StyleDef_Serialize(d))
    If Len(msg) > 0 Then StyleStore_SaveUser = "!" & msg
End Function

Public Function StyleStore_DeleteUser(name As String) As String
    Dim p As String
    On Error GoTo Fail
    If StyleDef_IsBuiltin(name) Then StyleStore_DeleteUser = "!Estilos embutidos não podem ser excluídos.": Exit Function
    p = StyleStore_UserPath(name)
    If Len(p) > 0 Then
        If Len(Dir(p)) > 0 Then Kill p
    End If
    Exit Function
Fail:
    StyleStore_DeleteUser = "!Não foi possível excluir: " & Err.Description
End Function

Private Function StyleStore_Esc(s As String) As String
    StyleStore_Esc = Replace(Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;")
End Function

' Estilos da pasta: Dictionary (texto) nome -> texto, lidos das partes no namespace (filhos <style name="...">texto</style>). Nunca propaga.
Public Function StyleStore_BookAll(wb As Workbook) As Object
    Dim d As Object, p As Object, nd As Object, a As Object, nm As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    Set StyleStore_BookAll = d
    On Error Resume Next
    For Each p In wb.CustomXMLParts.SelectByNamespace(STYLESTORE_NS)
        For Each nd In p.DocumentElement.ChildNodes
            If nd.BaseName = "style" Then
                nm = ""
                For Each a In nd.Attributes
                    If a.BaseName = "name" Then nm = StyleDef_SafeName(CStr(a.NodeValue))
                Next
                If Len(nm) > 0 And Len(nd.Text) <= STYLEDEF_MAX_BYTES Then d(nm) = CStr(nd.Text)
            End If
        Next
    Next
    Err.Clear
End Function

Public Function StyleStore_BookText(wb As Workbook, name As String) As String
    Dim d As Object
    Set d = StyleStore_BookAll(wb)
    If d.Exists(name) Then StyleStore_BookText = CStr(d(name))
End Function

' Grava/atualiza o estilo na pasta (uma unica parte no namespace, reescrita). Propaga erro (quem chama trata).
Public Sub StyleStore_SaveBook(wb As Workbook, name As String, text As String)
    Dim all As Object, k As Variant, xml As String, p As Object, old As New Collection
    Set all = StyleStore_BookAll(wb)
    all(StyleDef_SafeName(name)) = text
    xml = "<tqqStyles xmlns=""" & STYLESTORE_NS & """>"
    For Each k In all.Keys
        xml = xml & "<style name=""" & StyleStore_Esc(CStr(k)) & """>" & StyleStore_Esc(CStr(all(k))) & "</style>"
    Next
    xml = xml & "</tqqStyles>"
    For Each p In wb.CustomXMLParts.SelectByNamespace(STYLESTORE_NS)
        old.Add p
    Next
    For Each p In old
        p.Delete
    Next
    wb.CustomXMLParts.Add xml
End Sub

' A pasta passa a levar o estilo (so grava se o texto mudou). Nunca propaga (registrado).
Public Sub StyleStore_EnsureInBook(wb As Workbook, name As String, text As String)
    On Error GoTo Fail
    If Replace(StyleStore_BookText(wb, name), vbCrLf, vbLf) = Replace(text, vbCrLf, vbLf) Then Exit Sub   ' o XML devolve LF: nada a regravar
    StyleStore_SaveBook wb, name, text
    Exit Sub
Fail:
    Errors_Log "StyleStore_EnsureInBook" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Texto do estilo: preferUser (escolha explicita do usuario) = usuario, embutido, pasta; senao (reaplicar) = pasta, usuario, embutido.
Public Function StyleStore_Get(name As String, Optional wb As Workbook, Optional preferUser As Boolean = False) As String
    Dim t As String
    If preferUser Then
        t = StyleStore_UserText(name)
        If Len(t) = 0 Then t = StyleDef_Builtin(name)
        If Len(t) = 0 And Not wb Is Nothing Then t = StyleStore_BookText(wb, name)
    Else
        If Not wb Is Nothing Then t = StyleStore_BookText(wb, name)
        If Len(t) = 0 Then t = StyleStore_UserText(name)
        If Len(t) = 0 Then t = StyleDef_Builtin(name)
    End If
    StyleStore_Get = t
End Function

' Nomes: embutidos, do usuario (ordem do disco) e os so da pasta; sem repeticao; "|". Nunca propaga.
Public Function StyleStore_Names(Optional wb As Workbook) As String
    Dim seen As Object, s As String, n As Variant
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1
    On Error Resume Next
    For Each n In StyleDef_BuiltinNames()
        If Not seen.Exists(CStr(n)) Then seen.Add CStr(n), True: s = s & IIf(Len(s) > 0, "|", "") & CStr(n)
    Next
    For Each n In StyleStore_UserNames()
        If Not seen.Exists(CStr(n)) Then seen.Add CStr(n), True: s = s & IIf(Len(s) > 0, "|", "") & CStr(n)
    Next
    If Not wb Is Nothing Then
        For Each n In StyleStore_BookAll(wb).Keys
            If Not seen.Exists(CStr(n)) Then seen.Add CStr(n), True: s = s & IIf(Len(s) > 0, "|", "") & CStr(n)
        Next
    End If
    StyleStore_Names = s
    Err.Clear
End Function

' Importa um arquivo de estilo (validado): grava no deposito do usuario com o nome do arquivo quando o "nome=" nao serve. Devolve o texto
' canonico gravado ou "!" + mensagem; name = nome gravado; ignored = linhas ignoradas.
Public Function StyleStore_Import(path As String, ByRef name As String, ByRef ignored As Long) As String
    Dim t As String, msg As String, d As Object, base As String
    t = StyleStore_ReadFile(path, msg)
    If Len(msg) > 0 Then StyleStore_Import = "!" & msg: Exit Function
    Set d = StyleDef_Parse(t, ignored)
    If d Is Nothing Then StyleStore_Import = "!O arquivo não é um estilo do ThinQQ válido (falta a linha " & STYLEDEF_HEADER & ", ou é grande demais).": Exit Function
    name = StyleDef_Get(d, "nome")
    If Len(name) = 0 Then
        base = Mid$(path, InStrRev(path, "\") + 1)
        If LCase$(Right$(base, Len(STYLEDEF_EXT))) = STYLEDEF_EXT Then base = Left$(base, Len(base) - Len(STYLEDEF_EXT))
        name = StyleDef_SafeName(base)
    End If
    If StyleDef_IsBuiltin(name) Then name = name & " (importado)"
    d("nome") = name
    t = StyleDef_Serialize(d)
    msg = StyleStore_SaveUser(name, t)
    If Len(msg) > 0 Then StyleStore_Import = msg: Exit Function
    StyleStore_Import = t
End Function

Public Function StyleStore_Export(name As String, path As String, Optional wb As Workbook) As String
    Dim t As String, d As Object, ign As Long
    t = StyleStore_Get(name, wb, True)
    Set d = StyleDef_Parse(t, ign)
    If d Is Nothing Then StyleStore_Export = "!Estilo não encontrado: " & name: Exit Function
    StyleStore_Export = StyleStore_WriteFile(path, StyleDef_Serialize(d))
End Function

Public Function StyleStore_DefaultName() As String
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then StyleStore_DefaultName = mTestDefault: Exit Function
    StyleStore_DefaultName = GetSetting(APP_NAME, STYLESTORE_SECTION, "Padrao", "")
End Function

Public Sub StyleStore_SetDefaultName(n As String)
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then mTestDefault = n: Exit Sub
    SaveSetting APP_NAME, STYLESTORE_SECTION, "Padrao", n
End Sub
```
  Nota do teste de importação: "nome=Recebido;x" é inválido (`;`) e "comando=..." é chave desconhecida ⇒ `ignored = 2` e o nome vem do arquivo.

- [ ] **Passo 4: API** — em `modCoreApi.bas` (seção nova "Estilos de grafico (1.12.0)"):

```vb
' Nomes dos estilos para a faixa (embutidos, do usuario e os da pasta ativa), "|". Nunca propaga erro.
Public Function Core_StyleNames() As String
    On Error Resume Next
    Core_StyleNames = StyleStore_Names(ActiveWorkbook)
End Function

Public Function Core_StyleDefault() As String
    On Error Resume Next
    Core_StyleDefault = StyleStore_DefaultName()
End Function

Public Sub Core_StyleSetDefault(name As String)
    On Error Resume Next
    StyleStore_SetDefaultName StyleDef_SafeName(name)
End Sub

' Importar estilo (arquivo .tqqestilo escolhido pelo usuario; build Test sai antes do dialogo). Nunca propaga erro.
Public Sub Core_StyleImport()
    Dim f As Variant, r As String, nm As String, ign As Long
    On Error GoTo Fail
    If BUILD_CONFIG = "Test" Then Exit Sub
    f = Application.GetOpenFilename("Estilo do ThinQQ (*" & STYLEDEF_EXT & "),*" & STYLEDEF_EXT, , "Importar estilo do ThinQQ")
    If VarType(f) = vbBoolean Then Exit Sub
    r = StyleStore_Import(CStr(f), nm, ign)
    If Left$(r, 1) = "!" Then Core_Notify Mid$(r, 2), vbExclamation: Exit Sub
    Core_Notify "Estilo importado: " & nm & "." & IIf(ign > 0, " Linhas ignoradas (inválidas ou desconhecidas): " & ign & ".", ""), vbInformation
    Exit Sub
Fail:
    Errors_Handle "Core_StyleImport"
End Sub

Public Sub Core_StyleExport(name As String)
    Dim f As Variant, r As String
    On Error GoTo Fail
    If BUILD_CONFIG = "Test" Then Exit Sub
    f = Application.GetSaveAsFilename(StyleDef_SafeName(name) & STYLEDEF_EXT, "Estilo do ThinQQ (*" & STYLEDEF_EXT & "),*" & STYLEDEF_EXT, , "Exportar estilo do ThinQQ")
    If VarType(f) = vbBoolean Then Exit Sub
    r = StyleStore_Export(name, CStr(f), ActiveWorkbook)
    If Len(r) > 0 Then Core_Notify Mid$(r, 2), vbExclamation Else Core_Notify "Estilo exportado para " & CStr(f) & ".", vbInformation
    Exit Sub
Fail:
    Errors_Handle "Core_StyleExport"
End Sub
```
- [ ] **Passo 5:** verde + suíte `-Target all -LoaderSwap` → 0; commit `feat: depositos de estilos (usuario, pasta, embutidos), importar e exportar`.

---

### Tarefa S3: Aplicar por papel, hachura, reaplicação (sobrevivência) e ler do gráfico — modelo: most-capable

**Files:** Create `src/core/modules/modStyleApply.bas`; Modify `src/core/modules/modCoreApi.bas` (ganchos de reaplicação, `Opt_ApplyPalette`,
`Core_StyleApply`, `Core_GetStyleName`, `Core_StyleFromChart`), `src/core/modules/modChartKinds.bas` (recurso `style`); Test `tests/vba/core/TestStyles.bas`,
`tests/vba/core/TestKinds.bas` (matriz).

**Interfaces (Produces):** `Style_ApplyTo(c As Chart, d As Object)`; `Style_ApplyNamed(c As Chart, name As String, Optional preferUser As Boolean = True)` (erro
`vbObjectError + 7140` com a mensagem se não acha); `Style_Reapply(c As Chart)`; `Style_ReapplyIfNeeded(c As Chart)`; `Style_Stamp(c As Chart) As String`;
`Style_FromChart(c As Chart, name As String) As String` (texto canônico); `Opt_ApplyPalette(c As Chart, pal As String)`; `Core_StyleApply(name As String,
scope As String)` (`chart`|`sheet`|`book`); `Core_GetStyleName() As String`; `Core_StyleFromChart(name As String) As String` ("" ok ou mensagem);
`Kinds_Supports(kind, "style")` = todos os tipos.

- [ ] **Passo 1: Testes (RED)** — acrescente a `TestStyles.bas` (fixtures: a mesma tabela de cascata dos testes de `TestWaterfall.bas`; copie o helper de lá
  para um `Private Function WfTable(ws As Worksheet) As Range` que escreve `A1:B6` = cabeçalho "Etapa","Valor"; "2026",20; "Ganho A",0,6; "Ganho B",0,3;
  "Base","e"; "Perda",-0,9 e devolve `A1:B6`):

```vb
Private Function NewWfChart(wb As Workbook) As Chart
    Dim rng As Range
    Set rng = WfTable(wb.Worksheets(1))
    rng.Select
    Core_Insert "waterfall"
    Set NewWfChart = Core_ActiveChart()
End Function

Private Function SeriesFill(c As Chart, nm As String) As Long
    SeriesFill = c.SeriesCollection(nm).Format.Fill.ForeColor.RGB
End Function

' Papeis da cascata: aumento, reducao, total; a 1a barra (aumento a partir de zero) hachurada com borda; zero de outra serie no ponto 1 nao
' ganha borda (linha fantasma na base).
Public Sub Test_StyleApply_Waterfall_Roles_AndHatch()
    Dim wb As Workbook, c As Chart, p As Point
    Set wb = NewTempWorkbook()
    Set c = NewWfChart(wb)
    Style_ApplyNamed c, "Cascata MC"
    AssertEqual "Cascata MC", Registry_GetOption(c, "style", "")
    AssertEqual 8, Len(Registry_GetOption(c, "stylever", ""))
    AssertEqual RGB(112, 173, 71), SeriesFill(c, "tqq_uppos")
    AssertEqual RGB(232, 23, 60), SeriesFill(c, "tqq_downneg")
    AssertEqual RGB(217, 217, 217), SeriesFill(c, "tqq_total")
    Set p = c.SeriesCollection("tqq_uppos").Points(1)
    AssertEqual msoFillPatterned, p.Format.Fill.Type, "1a barra hachurada"
    AssertEqual msoPatternWideUpwardDiagonal, p.Format.Fill.Pattern
    AssertEqual RGB(0, 0, 0), p.Format.Fill.ForeColor.RGB
    AssertEqual RGB(255, 255, 255), p.Format.Fill.BackColor.RGB
    AssertEqual msoTrue, p.Format.Line.Visible, "borda da 1a barra"
    AssertEqual msoFalse, c.SeriesCollection("tqq_total").Points(1).Format.Line.Visible, "total zero no ponto 1: sem borda"
    AssertTrue InStr(StyleStore_BookText(wb, "Cascata MC"), "aumento=#70AD47") > 0, "a pasta passa a levar o estilo"
    CloseTemp wb
End Sub

' Sobrevivencia (Review Focus 6): ligar/desligar conectores, mudar formato, Alterar dados com uma linha a mais, refresh com categoria nova.
Public Sub Test_StyleApply_SurvivesTogglesRebuildAndRefresh()
    Dim wb As Workbook, c As Chart, ws As Worksheet, i As Long, conn As Shape
    Set wb = NewTempWorkbook()
    Set c = NewWfChart(wb)
    Set ws = wb.Worksheets(1)
    Style_ApplyNamed c, "Cascata MC"
    Core_SetToggle "connectors", False
    Core_SetToggle "connectors", True
    For i = 1 To c.Shapes.Count
        If Left$(c.Shapes(i).Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then Set conn = c.Shapes(i): Exit For
    Next
    AssertTrue Not conn Is Nothing, "conectores recriados"
    AssertEqual RGB(127, 127, 127), conn.Line.ForeColor.RGB
    AssertEqual msoLineRoundDot, conn.Line.DashStyle, "conector pontilhado reaplicado"
    Core_SetOption "numfmt", "dec1"
    AssertEqual RGB(112, 173, 71), SeriesFill(c, "tqq_uppos"), "trocar o formato nao perde as cores"
    ws.Range("A7:B7").Value = Array("Final", "e")
    AssertTrue Core_RebuildFromRange(c, ws.Range("A1:B7"))
    Set c = Core_ActiveChart()
    AssertEqual "Cascata MC", Registry_GetOption(c, "style", "")
    AssertEqual RGB(112, 173, 71), SeriesFill(c, "tqq_uppos"), "Alterar dados reaplica"
    AssertEqual msoFillPatterned, c.SeriesCollection("tqq_uppos").Points(1).Format.Fill.Type
    AssertEqual Style_Stamp(c), Registry_GetOption(c, "stylestamp", "")
    Core_SetOption "palette", "cinza"
    AssertEqual "", Registry_GetOption(c, "style", ""), "escolher uma paleta embutida tira o estilo"
    CloseTemp wb
End Sub

' Todos os 21 tipos aceitam um estilo de series sem erro; nos tipos de series, a 1a serie de dados recebe serie1.
Public Sub Test_StyleApply_AllKinds_NoError()
    Dim wb As Workbook, k As Variant, c As Chart, rng As Range, d As Object, ign As Long, s As Series, got As Boolean
    Set d = StyleDef_Parse(STYLEDEF_HEADER & vbCrLf & "serie1=#123456" & vbCrLf & "serie2=#654321" & vbCrLf & "fonte.tamanho=10", ign)
    For Each k In Array("stacked", "clustered", "stacked100", "line", "waterfall", "barstacked", "barclustered", "barstacked100", "barwaterfall", "area", "area100", "combo", "mekko", "mekkounits", "gantt", "pie", "doughnut", "butterfly", "scatter", "bubble", "waterfallpair")
        Set wb = NewTempWorkbook()
        Set rng = Samples_TableFor(wb.Worksheets(1), CStr(k))
        rng.Select
        Core_Insert CStr(k)
        Set c = Core_ActiveChart()
        AssertTrue Not c Is Nothing, CStr(k)
        Style_ApplyTo c, d
        AssertEqual 10, CLng(c.ChartArea.Font.Size), "fonte: " & CStr(k)
        got = False
        For Each s In c.SeriesCollection
            If s.Format.Fill.ForeColor.RGB = RGB(&H12, &H34, &H56) Or s.Format.Line.ForeColor.RGB = RGB(&H12, &H34, &H56) Then got = True
        Next
        If InStr(",stacked,clustered,stacked100,line,barstacked,barclustered,barstacked100,area,area100,combo,", "," & CStr(k) & ",") > 0 Then AssertTrue got, "serie1: " & CStr(k)
        CloseTemp wb
    Next
End Sub

Public Sub Test_StyleFromChart_CapturesManualColors()
    Dim wb As Workbook, c As Chart, t As String, d As Object, ign As Long
    Set wb = NewTempWorkbook()
    Set c = NewWfChart(wb)
    c.SeriesCollection("tqq_uppos").Format.Fill.ForeColor.RGB = RGB(1, 2, 3)        ' pintado a mao
    c.SeriesCollection("tqq_uppos").Points(1).Format.Fill.Patterned msoPatternDarkUpwardDiagonal
    t = Style_FromChart(c, "Do Grafico")
    Set d = StyleDef_Parse(t, ign)
    AssertEqual "#010203", StyleDef_Get(d, "aumento")
    AssertEqual "diagonal_escura_para_cima", StyleDef_Get(d, "primeira.padrao")
    AssertEqual "Do Grafico", StyleDef_Get(d, "nome")
    CloseTemp wb
End Sub

Public Sub Test_Core_StyleApply_Refusals_NeverRaise()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Set c = NewWfChart(wb)
    Core_ClearLastNotice
    Core_StyleApply "Nao Existe", "chart"
    AssertTrue InStr(Core_LastNotice(), "Estilo não encontrado") > 0, Core_LastNotice()
    Core_StyleApply "Cascata MC", "chart"
    AssertEqual "Cascata MC", Core_GetStyleName()
    Core_StyleApply "Tons de cinza", "chart"
    AssertEqual "Tons de cinza", Core_GetStyleName(), "paleta embutida vira o nome do estilo"
    CloseTemp wb
End Sub
```
  E em `TestKinds.bas`, no teste da matriz, acrescente `AssertTrue Kinds_Supports(k, "style")` para todos os tipos (no laço que já percorre os 21).
  `Samples_TableFor(ws, kind)` existe desde a G1 (dados de exemplo dos 21 tipos); se a assinatura for outra, use a dela.

- [ ] **Passo 2:** RED.
- [ ] **Passo 3: `src/core/modules/modStyleApply.bas`**:

```vb
Attribute VB_Name = "modStyleApply"
Option Explicit
' Estilos de grafico (1.12.0; spec 2026-10-23 secao 8.5): aplica um estilo validado por PAPEL. Camada reaplicada pelo core depois de tudo que
' recolore (Insert, toggles, opcoes, Alterar dados; refresh quando o carimbo muda). Opcoes do grafico: style (nome), stylever (8 hex do texto),
' stylestamp (carimbo do ultimo aplicar: versao, series, categorias, series que carregam a 1a e a ultima barra da cascata). Nunca mexe na
' formatacao que o estilo nao define. Hachura e cor de ponto sobrevivem a mudar valores (PX1); recriar o grafico perde: por isso a reaplicacao.
Private Const STYLE_ERR As Long = 7140

' Preenchimento por papel: hachura (padrao valido) com a cor do padrao na frente e a do papel atras; senao cor solida. Nunca propaga (registrado).
Private Sub Style_Fill(ff As FillFormat, colorHex As String, patternName As String, patternHex As String)
    Dim fg As Long, bg As Long, pt As Long
    On Error GoTo Fail
    bg = StyleDef_ColorOf(colorHex)
    pt = StyleDef_PatternOf(patternName)
    If pt <> 0 Then
        fg = StyleDef_ColorOf(patternHex)
        If fg < 0 Then fg = RGB(0, 0, 0)
        If bg < 0 Then bg = RGB(255, 255, 255)
        ff.Visible = msoTrue
        ff.Patterned pt
        ff.ForeColor.RGB = fg
        ff.BackColor.RGB = bg
    ElseIf bg >= 0 Then
        ff.Visible = msoTrue
        ff.Solid
        ff.ForeColor.RGB = bg
    End If
    Exit Sub
Fail:
    Errors_Log "Style_Fill" & vbTab & Err.Number & vbTab & Err.Description
End Sub

Private Function Style_PointHasValue(s As Series, i As Long) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = s.Values
    If IsArray(v) Then
        If IsNumeric(v(i)) Then Style_PointHasValue = (CDbl(v(i)) <> 0)
    End If
    Err.Clear
End Function

' Ponto i de uma serie de papel: com valor = o papel especial (primeira/ultima: cor, hachura, borda); sem valor = volta a cor da serie e sem
' borda (um zero com borda vira linha na base). So age se o estilo define o papel especial. Nunca propaga.
Private Sub Style_EdgePoint(s As Series, i As Long, d As Object, role As String, seriesHex As String)
    Dim p As Point, b As String
    If Len(StyleDef_Get(d, role)) = 0 And Len(StyleDef_Get(d, role & ".padrao")) = 0 And Len(StyleDef_Get(d, role & ".borda")) = 0 Then Exit Sub
    On Error GoTo Fail
    Set p = s.Points(i)
    If Style_PointHasValue(s, i) Then
        Style_Fill p.Format.Fill, IIf(Len(StyleDef_Get(d, role)) > 0, StyleDef_Get(d, role), seriesHex), StyleDef_Get(d, role & ".padrao"), StyleDef_Get(d, role & ".padrao.cor")
        b = StyleDef_Get(d, role & ".borda")
        If b = "nenhuma" Then
            p.Format.Line.Visible = msoFalse
        ElseIf Len(b) > 0 Then
            p.Format.Line.Visible = msoTrue
            p.Format.Line.ForeColor.RGB = StyleDef_ColorOf(b)
            p.Format.Line.Weight = 0.75
        End If
    Else
        If Len(seriesHex) > 0 Then Style_Fill p.Format.Fill, seriesHex, "", ""
        p.Format.Line.Visible = msoFalse
    End If
    Exit Sub
Fail:
    Errors_Log "Style_EdgePoint" & vbTab & role & vbTab & Err.Number & vbTab & Err.Description
End Sub

Private Sub Style_WaterfallRoles(c As Chart, d As Object)
    Dim s As Series, role As String, n As Long, hx As String
    For Each s In c.SeriesCollection
        role = ""
        Select Case s.Name
            Case "tqq_uppos", "tqq_upneg": role = "aumento"
            Case "tqq_downpos", "tqq_downneg": role = "reducao"
            Case "tqq_total": role = "total"
        End Select
        If Len(role) > 0 Then
            hx = StyleDef_Get(d, role)
            If Len(hx) > 0 Then
                Style_Fill s.Format.Fill, hx, "", ""
                s.Format.Line.Visible = msoFalse
            End If
            n = s.Points.Count
            If n >= 1 Then Style_EdgePoint s, 1, d, "primeira", hx
            If n >= 2 Then Style_EdgePoint s, n, d, "ultima", hx
        End If
    Next
End Sub

' Cor de serie k (ciclo nas cores serieN definidas). Linhas e dispersao: linha e marcador; area: preenchimento (a borda branca fica); demais:
' preenchimento sem borda. Nunca propaga.
Private Sub Style_SeriesColor(s As Series, d As Object, k As Long)
    Dim m As Long, hx As String, rgbv As Long
    m = StyleDef_SeriesCount(d)
    If m = 0 Then Exit Sub
    hx = StyleDef_Get(d, "serie" & (((k - 1) Mod m) + 1))
    rgbv = StyleDef_ColorOf(hx)
    On Error GoTo Fail
    Select Case s.ChartType
        Case xlLine, xlLineMarkers, xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
            s.Format.Line.ForeColor.RGB = rgbv
            s.Format.Fill.ForeColor.RGB = rgbv
        Case xlArea, xlAreaStacked, xlAreaStacked100
            Style_Fill s.Format.Fill, hx, "", ""
        Case Else
            Style_Fill s.Format.Fill, hx, "", ""
            s.Format.Line.Visible = msoFalse
    End Select
    Exit Sub
Fail:
    Errors_Log "Style_SeriesColor" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Series de dados pela ordem (auxiliares tqq_* nunca); cascata de varias series: as partes tqq_dn_j/tqq_upn_j/tqq_dnn_j com a cor da serie j e
' tqq_total com "total" (a regra de Waterfall_ApplyColors).
Private Sub Style_SeriesRoles(c As Chart, d As Object, waterfallMulti As Boolean)
    Dim s As Series, k As Long, j As Long, p As Long, nm As String
    For Each s In c.SeriesCollection
        nm = s.Name
        If waterfallMulti Then
            j = 0
            If nm = "tqq_total" Then
                If Len(StyleDef_Get(d, "total")) > 0 Then Style_Fill s.Format.Fill, StyleDef_Get(d, "total"), "", ""
            ElseIf Left$(nm, 7) = "tqq_dn_" Or Left$(nm, 8) = "tqq_upn_" Or Left$(nm, 8) = "tqq_dnn_" Then
                p = InStrRev(nm, "_")
                j = CLng(Val(Mid$(nm, p + 1)))
            ElseIf Not Style_IsHelperSeries(s) Then
                k = k + 1: j = k
            End If
            If j > 0 Then Style_SeriesColor s, d, j
        ElseIf Not Style_IsHelperSeries(s) Then
            k = k + 1
            Style_SeriesColor s, d, k
        End If
    Next
End Sub

' Pizza e rosca: cor por fatia (serieN em ciclo). Nunca propaga.
Private Sub Style_PointRoles(c As Chart, d As Object)
    Dim s As Series, i As Long, m As Long
    m = StyleDef_SeriesCount(d)
    If m = 0 Or c.SeriesCollection.Count = 0 Then Exit Sub
    On Error Resume Next
    Set s = c.SeriesCollection(1)
    For i = 1 To s.Points.Count
        Style_Fill s.Points(i).Format.Fill, StyleDef_Get(d, "serie" & (((i - 1) Mod m) + 1)), "", ""
    Next
    Err.Clear
End Sub

Private Sub Style_Labels(c As Chart, d As Object)
    Dim s As Series, i As Long, p As Point, col As String, pos As Long
    If Len(StyleDef_Get(d, "rotulo.cor")) = 0 And Len(StyleDef_Get(d, "rotulo.negrito")) = 0 And Len(StyleDef_Get(d, "rotulo.tamanho")) = 0 Then Exit Sub
    col = StyleDef_Get(d, "rotulo.cor")
    On Error Resume Next
    For Each s In c.SeriesCollection
        For i = 1 To s.Points.Count
            Set p = s.Points(i)
            If p.HasDataLabel Then
                If Len(StyleDef_Get(d, "rotulo.negrito")) > 0 Then p.DataLabel.Font.Bold = (StyleDef_Get(d, "rotulo.negrito") = "1")
                If Len(StyleDef_Get(d, "rotulo.tamanho")) > 0 Then p.DataLabel.Font.Size = CLng(StyleDef_Get(d, "rotulo.tamanho"))
                If col = "auto" Then
                    pos = 0: pos = p.DataLabel.Position
                    If pos = xlLabelPositionCenter Or pos = xlLabelPositionInsideEnd Or pos = xlLabelPositionInsideBase Then p.DataLabel.Font.Color = Style_ContrastFont(CLng(p.Format.Fill.ForeColor.RGB))
                ElseIf Len(col) > 0 Then
                    p.DataLabel.Font.Color = StyleDef_ColorOf(col)
                End If
            End If
        Next
    Next
    Err.Clear
End Sub

Private Sub Style_Line(lf As Object, d As Object, role As String)
    On Error Resume Next
    If Len(StyleDef_Get(d, role & ".cor")) > 0 Then lf.Visible = msoTrue: lf.ForeColor.RGB = StyleDef_ColorOf(StyleDef_Get(d, role & ".cor"))
    If Len(StyleDef_Get(d, role & ".espessura")) > 0 Then lf.Weight = Val(StyleDef_Get(d, role & ".espessura"))
    Select Case StyleDef_Get(d, role & ".traco")
        Case "pontos": lf.DashStyle = msoLineRoundDot
        Case "tracos": lf.DashStyle = msoLineDash
        Case "solido": lf.DashStyle = msoLineSolid
    End Select
    Err.Clear
End Sub

Private Sub Style_Connectors(c As Chart, d As Object)
    Dim i As Long, g As ChartGroup
    If Len(StyleDef_Get(d, "conector.cor")) = 0 And Len(StyleDef_Get(d, "conector.traco")) = 0 And Len(StyleDef_Get(d, "conector.espessura")) = 0 Then Exit Sub
    On Error Resume Next
    For i = 1 To c.Shapes.Count
        If Left$(c.Shapes(i).Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then Style_Line c.Shapes(i).Line, d, "conector"
    Next
    For Each g In c.ChartGroups
        If g.HasSeriesLines Then Style_Line g.SeriesLines.Format.Line, d, "conector"
    Next
    Err.Clear
End Sub

' Setas (CAGR, diferenca, delta, variacao da comparativa): linhas com seta.cor; caixas de texto com fundo e borda do rotulo. Nunca propaga.
Private Sub Style_Arrows(c As Chart, d As Object)
    Dim i As Long, s As Shape, nm As String, b As String, hasTxt As Boolean
    If Len(StyleDef_Get(d, "seta.cor")) = 0 And Len(StyleDef_Get(d, "seta.rotulo.fundo")) = 0 And Len(StyleDef_Get(d, "seta.rotulo.borda")) = 0 Then Exit Sub
    On Error Resume Next
    For i = 1 To c.Shapes.Count
        Set s = c.Shapes(i): nm = Mid$(s.Name, Len(ANN_PREFIX) + 1)
        If Left$(s.Name, Len(ANN_PREFIX)) = ANN_PREFIX And Left$(nm, 5) <> "conn_" And Left$(nm, 7) <> "ind100" And Left$(nm, 6) <> "break_" Then
            hasTxt = False
            hasTxt = (s.TextFrame2.HasText = msoTrue)
            If hasTxt Then
                If Len(StyleDef_Get(d, "seta.rotulo.fundo")) > 0 Then
                    s.Fill.Visible = msoTrue: s.Fill.Solid: s.Fill.ForeColor.RGB = StyleDef_ColorOf(StyleDef_Get(d, "seta.rotulo.fundo"))
                End If
                b = StyleDef_Get(d, "seta.rotulo.borda")
                If b = "nenhuma" Then
                    s.Line.Visible = msoFalse
                ElseIf Len(b) > 0 Then
                    s.Line.Visible = msoTrue: s.Line.ForeColor.RGB = StyleDef_ColorOf(b)
                End If
            ElseIf Len(StyleDef_Get(d, "seta.cor")) > 0 Then
                s.Line.ForeColor.RGB = StyleDef_ColorOf(StyleDef_Get(d, "seta.cor"))
            End If
        End If
    Next
    Err.Clear
End Sub

Private Sub Style_Gap(c As Chart, gap As Long)
    Dim g As ChartGroup, n As Long
    On Error Resume Next
    For Each g In c.ChartGroups
        Select Case g.SeriesCollection(1).ChartType
            Case xlColumnClustered, xlColumnStacked, xlColumnStacked100, xlBarClustered, xlBarStacked, xlBarStacked100
                g.GapWidth = gap: n = n + 1
        End Select
    Next
    If n = 0 Then c.ChartGroups(1).GapWidth = gap
    Registry_SetOption c, "gap", CStr(gap)
    If Kinds_Family(Registry_ChartType(c)) = "waterfallpair" Then Wfp_SyncGap c
    Err.Clear
End Sub

' Aplica o estilo d (validado) ao grafico. Chaves sem papel no tipo sao ignoradas. Gantt (1.12.0): so fonte, rotulos e espaco (as cores do Gantt
' seguem a paleta; backlog). Propaga so erro estrutural (quem chama trata).
Public Sub Style_ApplyTo(c As Chart, d As Object)
    Dim fam As String
    fam = Kinds_Family(Registry_ChartType(c))
    On Error Resume Next
    If Len(StyleDef_Get(d, "fonte")) > 0 Then c.ChartArea.Font.Name = StyleDef_Get(d, "fonte")
    If Len(StyleDef_Get(d, "fonte.tamanho")) > 0 Then c.ChartArea.Font.Size = CLng(StyleDef_Get(d, "fonte.tamanho"))
    Err.Clear
    On Error GoTo 0
    Select Case fam
        Case "waterfall"
            If Waterfall_SeriesCount(c) > 1 Then Style_SeriesRoles c, d, True Else Style_WaterfallRoles c, d
        Case "pie": Style_PointRoles c, d
        Case "gantt"
        Case Else: Style_SeriesRoles c, d, False
    End Select
    Style_Labels c, d
    Style_Connectors c, d
    Style_Arrows c, d
    If Len(StyleDef_Get(d, "eixo.categoria.linha")) > 0 And fam <> "pie" Then
        On Error Resume Next
        c.Axes(xlCategory).Format.Line.Visible = IIf(StyleDef_Get(d, "eixo.categoria.linha") = "1", msoTrue, msoFalse)
        Err.Clear
        On Error GoTo 0
    End If
    If Len(StyleDef_Get(d, "espaco")) > 0 Then
        If Kinds_SupportsChart(c, "gap") Then Style_Gap c, CLng(StyleDef_Get(d, "espaco"))
    End If
End Sub

' Series (nomes) que carregam a 1a e a ultima barra da cascata de uma serie ("" nos outros tipos): entra no carimbo (o dado mudando de sinal
' muda a serie da barra). Nunca propaga.
Private Function Style_CarrierKey(c As Chart) As String
    Dim s As Series, n As Long, a As String, b As String
    On Error Resume Next
    If Kinds_Family(Registry_ChartType(c)) <> "waterfall" Then Exit Function
    If Waterfall_SeriesCount(c) > 1 Then Exit Function
    For Each s In c.SeriesCollection
        Select Case s.Name
            Case "tqq_uppos", "tqq_upneg", "tqq_downpos", "tqq_downneg", "tqq_total"
                n = s.Points.Count
                If Style_PointHasValue(s, 1) Then a = s.Name
                If n > 1 Then
                    If Style_PointHasValue(s, n) Then b = s.Name
                End If
        End Select
    Next
    Style_CarrierKey = a & "/" & b
    Err.Clear
End Function

Public Function Style_Stamp(c As Chart) As String
    On Error Resume Next
    Style_Stamp = Registry_GetOption(c, "stylever", "") & "." & c.SeriesCollection.Count & "." & Geom_ChartCategoryCount(c) & "." & Style_CarrierKey(c)
End Function

' Estilo pelo nome aplicado e gravado nas opcoes do grafico. Embutido so-paleta (ou "") = a paleta antiga (ApplyOption) e sem style. A pasta
' passa a levar o estilo. Erro STYLE_ERR (vbObjectError + 7140) com a mensagem quando o nome nao existe ou o texto nao vale.
Public Sub Style_ApplyNamed(c As Chart, name As String, Optional preferUser As Boolean = True)
    Dim txt As String, d As Object, ign As Long, pal As String
    pal = StyleDef_BuiltinPalette(name)
    If Len(pal) > 0 Or Len(name) = 0 Then
        If Len(pal) = 0 Then pal = "padrao"
        Opt_ApplyPalette c, pal                                     ' grava palette e limpa style (Style_Clear dentro do ApplyOption da paleta)
        Exit Sub
    End If
    txt = StyleStore_Get(name, Registry_Workbook(c), preferUser)
    If Len(txt) = 0 Then Err.Raise vbObjectError + STYLE_ERR, "Style", "Estilo não encontrado: " & name & "."
    Set d = StyleDef_Parse(txt, ign)
    If d Is Nothing Then Err.Raise vbObjectError + STYLE_ERR, "Style", "Estilo inválido: " & name & "."
    StyleStore_EnsureInBook Registry_Workbook(c), StyleDef_SafeName(name), txt
    Registry_SetOption c, "style", StyleDef_SafeName(name)
    Registry_SetOption c, "stylever", StyleDef_Version(txt)
    Style_ApplyTo c, d
    Registry_SetOption c, "stylestamp", Style_Stamp(c)
End Sub

' Tira o estilo do grafico (as cores ficam as da paleta reaplicada por quem chamou). Nunca propaga.
Public Sub Style_Clear(c As Chart)
    On Error Resume Next
    If Len(Registry_GetOption(c, "style", "")) = 0 Then Exit Sub
    Registry_SetOption c, "style", ""
    Registry_SetOption c, "stylever", ""
    Registry_SetOption c, "stylestamp", ""
    Err.Clear
End Sub

' Reaplica o estilo gravado no grafico, pela copia da PASTA (quem recebe a pasta ve o mesmo visual). Nunca propaga erro (registrado).
Public Sub Style_Reapply(c As Chart)
    Dim name As String
    On Error GoTo Fail
    name = Registry_GetOption(c, "style", "")
    If Len(name) = 0 Then Exit Sub
    Style_ApplyNamed c, name, False
    Exit Sub
Fail:
    Errors_Log "Style_Reapply" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Core_SafeRefresh: so reaplica quando o carimbo mudou (linha nova, serie nova, sinal da 1a/ultima barra). Nunca propaga erro.
Public Sub Style_ReapplyIfNeeded(c As Chart)
    On Error GoTo Fail
    If Len(Registry_GetOption(c, "style", "")) = 0 Then Exit Sub
    If Registry_GetOption(c, "stylestamp", "") = Style_Stamp(c) Then Exit Sub
    Style_Reapply c
    Exit Sub
Fail:
    Errors_Log "Style_ReapplyIfNeeded" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Le o visual atual do grafico (inclusive cores pintadas a mao) como estilo: papeis da cascata ou cores das series de dados, hachura/borda da 1a
' barra, fonte, conectores. Devolve o texto canonico. Propaga erro estrutural.
Public Function Style_FromChart(c As Chart, name As String) As String
    Dim d As Object, s As Series, k As Long, fam As String, i As Long, p As Point, nm As String
    Set d = CreateObject("Scripting.Dictionary")
    d("nome") = StyleDef_SafeName(name)
    fam = Kinds_Family(Registry_ChartType(c))
    On Error Resume Next
    d("fonte") = StyleDef_Validate("fonte", CStr(c.ChartArea.Font.Name))
    d("fonte.tamanho") = StyleDef_Validate("fonte.tamanho", CStr(CLng(c.ChartArea.Font.Size)))
    If fam = "waterfall" And Waterfall_SeriesCount(c) = 1 Then
        d("aumento") = StyleDef_HexOf(c.SeriesCollection("tqq_uppos").Format.Fill.ForeColor.RGB)
        d("reducao") = StyleDef_HexOf(c.SeriesCollection("tqq_downpos").Format.Fill.ForeColor.RGB)
        d("total") = StyleDef_HexOf(c.SeriesCollection("tqq_total").Format.Fill.ForeColor.RGB)
        For Each s In c.SeriesCollection
            nm = s.Name
            If nm = "tqq_uppos" Or nm = "tqq_upneg" Or nm = "tqq_downpos" Or nm = "tqq_downneg" Or nm = "tqq_total" Then
                If Style_PointHasValue(s, 1) Then
                    Set p = s.Points(1)
                    If p.Format.Fill.Type = msoFillPatterned Then
                        d("primeira") = StyleDef_HexOf(p.Format.Fill.BackColor.RGB)
                        d("primeira.padrao") = StyleDef_PatternName(CLng(p.Format.Fill.Pattern))
                        d("primeira.padrao.cor") = StyleDef_HexOf(p.Format.Fill.ForeColor.RGB)
                    Else
                        d("primeira") = StyleDef_HexOf(p.Format.Fill.ForeColor.RGB)
                    End If
                    If p.Format.Line.Visible = msoTrue Then d("primeira.borda") = StyleDef_HexOf(p.Format.Line.ForeColor.RGB)
                End If
            End If
        Next
        ' a barra pintada a mao no ponto 1 de uma serie sem valor tambem conta (ex.: o teste pinta tqq_uppos)
        Set p = c.SeriesCollection("tqq_uppos").Points(1)
        If p.Format.Fill.Type = msoFillPatterned Then d("primeira.padrao") = StyleDef_PatternName(CLng(p.Format.Fill.Pattern))
    Else
        For Each s In c.SeriesCollection
            If Not Style_IsHelperSeries(s) And k < 8 Then
                k = k + 1
                If s.ChartType = xlLine Or s.ChartType = xlLineMarkers Then
                    d("serie" & k) = StyleDef_HexOf(s.Format.Line.ForeColor.RGB)
                Else
                    d("serie" & k) = StyleDef_HexOf(s.Format.Fill.ForeColor.RGB)
                End If
            End If
        Next
    End If
    For i = 1 To c.Shapes.Count
        If Left$(c.Shapes(i).Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then
            d("conector.cor") = StyleDef_HexOf(c.Shapes(i).Line.ForeColor.RGB)
            Exit For
        End If
    Next
    Err.Clear
    On Error GoTo 0
    Style_FromChart = StyleDef_Serialize(d)
End Function
```
  Notas para o implementador: (1) confira os nomes das formas de anotação em `modAnnotations.bas` (`grep -n 'ANN_PREFIX & "' src/core/modules/modAnnotations.bas`)
  e ajuste a exclusão de `Style_Arrows` (o indicador 100% e as marcas do corte não são setas); (2) se `Test_StyleApply_AllKinds_NoError` mostrar
  que uma família usa só séries `tqq_*` para os dados (Mekko, borboleta, comparativa), trate-a num `Case` próprio com a mesma regra da sua
  `*_ApplyColors` (relate no relatório); (3) `Wfp_SyncGap` existe em `modChartWaterfallPair.bas`.

- [ ] **Passo 4: reaplicação no core** — `modCoreApi.bas`:
  - `Opt_ApplyPalette`:

```vb
' Paleta embutida (estilos "Padrão ThinQQ", "Tons de cinza", "Cores do tema"): o mesmo ApplyOption da faixa, que tira o estilo. Propaga erros.
Public Sub Opt_ApplyPalette(c As Chart, pal As String)
    ApplyOption c, "palette", pal
End Sub
```
  - em `ApplyOption`, no `Case "palette"` (depois do `Select Case fam` das cores): `Style_Clear c`.
  - `Core_SetToggle`: depois de `ApplyToggle c, key, value`: `Style_Reapply c`; `Core_SetOption`: depois de `ApplyOption c, key, value`: `If key <> "palette" Then Style_Reapply c`.
  - `Core_SafeRefresh`: no fim, `Style_ReapplyIfNeeded cht` e `If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Style" ...: Err.Clear`.
  - `Core_RebuildFromRange`: depois de `ReapplyOptions nc`: `Style_Reapply nc`.
  - `PptCmd_One`: depois do `End Select` e antes de `Core_SafeRefresh c`: `Style_Reapply c`.
  - `Core_Insert`: depois de `ApplySourceNumFmt co.Chart, kind, data`: `StyleDefault_ApplyNew co.Chart` com

```vb
' Estilo padrao para novos graficos (StyleStore_DefaultName). Falha (estilo apagado): registrada, o grafico fica com a paleta. Nunca propaga.
Private Sub StyleDefault_ApplyNew(c As Chart)
    Dim n As String
    On Error GoTo Fail
    n = StyleStore_DefaultName()
    If Len(n) = 0 Then Exit Sub
    Style_ApplyNamed c, n, True
    Exit Sub
Fail:
    Errors_Log "StyleDefault_ApplyNew" & vbTab & Err.Number & vbTab & Err.Description
End Sub
```
  - API:

```vb
' Estilo do grafico ativo para a faixa: o nome gravado ou, sem estilo, o embutido da paleta. "" sem grafico. Nunca propaga erro.
Public Function Core_GetStyleName() As String
    Dim c As Chart, n As String
    On Error Resume Next
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Function
    n = Registry_GetOption(c, "style", "")
    If Len(n) = 0 Then
        Select Case Registry_GetOption(c, "palette", "padrao")
            Case "cinza": n = "Tons de cinza"
            Case "tema": n = "Cores do tema"
            Case Else: n = "Padrão ThinQQ"
        End Select
    End If
    Core_GetStyleName = n
End Function

' Aplica o estilo: scope "chart" (os selecionados ou o ativo), "sheet" (todos os do ThinQQ da planilha ativa) ou "book" (da pasta ativa).
' Recusa (estilo inexistente) vira aviso. Nunca propaga erro.
Public Sub Core_StyleApply(name As String, scope As String)
    Dim targets As New Collection, c As Variant, ws As Object, co As ChartObject, n As Long
    On Error GoTo Fail
    Select Case scope
        Case "sheet", "book"
            If ActiveWorkbook Is Nothing Then Exit Sub
            For Each ws In ActiveWorkbook.Worksheets
                If scope = "book" Or ws Is ActiveSheet Then
                    For Each co In ws.ChartObjects
                        If Left$(co.Name, Len(ID_PREFIX)) = ID_PREFIX Then targets.Add co.Chart
                    Next
                End If
            Next
        Case Else
            For Each c In Core_SelectedCharts()
                targets.Add c
            Next
            If targets.Count = 0 Then
                If Not Core_ActiveChart() Is Nothing Then targets.Add Core_ActiveChart()
            End If
    End Select
    For Each c In targets
        Style_ApplyNamed c, name, True
        Core_SafeRefresh c
        n = n + 1
    Next
    If n = 0 Then Core_Notify "Selecione um gráfico do ThinQQ (ou escolha planilha/pasta).", vbInformation
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7140 Then Core_Notify Err.Description, vbExclamation Else Errors_Handle "Core_StyleApply"
End Sub

' Salva o visual do grafico ativo como estilo do usuario. "" ok ou a mensagem. Nunca propaga erro.
Public Function Core_StyleFromChart(name As String) As String
    Dim c As Chart, r As String
    On Error GoTo Fail
    Set c = Core_ActiveChart()
    If c Is Nothing Then Core_StyleFromChart = "Selecione um gráfico do ThinQQ.": Exit Function
    r = StyleStore_SaveUser(name, Style_FromChart(c, name))
    If Left$(r, 1) = "!" Then Core_StyleFromChart = Mid$(r, 2)
    Exit Function
Fail:
    Core_StyleFromChart = "Não foi possível salvar o estilo: " & Err.Description
End Function
```
  (`Core_SelectedCharts()` existe: devolve os gráficos do ThinQQ selecionados.) `modChartKinds.bas`: `Case "style": list = allKinds` em `Kinds_Supports`
  e `Case "style": Kinds_FeatureLabel = "Estilo"`.

- [ ] **Passo 5:** `-Target core -Filter "TestStyles|TestKinds|TestWaterfall|TestCoreApi"` → verde; suíte `-Target all -LoaderSwap` → 0 (o refresh
  ganhou um passo: compare a duração com a da U6 e relate). Commit `feat: estilos aplicados por papel (hachura, cascata) e reaplicados pelo core`.

---

### Tarefa S4: Janela "Estilos" e faixa do Excel — modelo: standard

**Files:** Create `src/core/forms/frmStyles.json`, `src/core/forms/frmStyles.vba`, `src/core/modules/modStyleUi.bas`; Modify `src/core/modules/modCoreApi.bas`
(`Core_StyleManage`), `src/loader/ribbon/customUI14.xml`, `src/loader/modules/modRibbon.bas`; Test `tests/vba/core/TestStyles.bas`, `tests/vba/loader/TestRibbon.bas`,
`tests/ps/Test-Ribbon.ps1`, `tests/vba/core/TestForms.bas` (conta 2 formulários).

**Interfaces (Produces):** `StyleUi_RoleLabel(key) As String`; `StyleUi_PickItems(key) As Variant` (rótulos); `StyleUi_PickValue(key, index) As String`;
`StyleUi_IsColorKey(key) As Boolean`; `Core_StyleManage()`; loader `Rx_StyleCount`, `Rx_StyleLabel`, `Rx_StyleId`, `Rx_GetStyleSel`, `Rx_OnStyle`,
`Rx_OnStyles`, `Rx_StyleNamesFor(raw As String) As Variant` (testável).

- [ ] **Passo 1 (RED):** testes puros

```vb
Public Sub Test_StyleUi_RolesAndPicks()
    Dim k As Variant, items As Variant
    For Each k In StyleDef_Keys()
        AssertTrue Len(StyleUi_RoleLabel(CStr(k))) > 0, "rotulo em pt-BR: " & CStr(k)
        items = StyleUi_PickItems(CStr(k))
        If IsArray(items) Then
            AssertEqual StyleDef_Validate(CStr(k), StyleUi_PickValue(CStr(k), 0)), StyleUi_PickValue(CStr(k), 0), "1o item valido: " & CStr(k)
        End If
    Next
    AssertEqual "Aumentos (cascata)", StyleUi_RoleLabel("aumento")
    AssertEqual "#70AD47", StyleUi_PickValue("aumento", 9)
    AssertEqual "diagonal_larga_para_cima", StyleUi_PickValue("primeira.padrao", 3)
    AssertTrue StyleUi_IsColorKey("serie1")
    AssertFalse StyleUi_IsColorKey("espaco")
End Sub
```
  e em `TestRibbon.bas`: `Rx_StyleNamesFor("A|B|C")` devolve 3 itens; `Rx_StyleNamesFor("")` devolve 1 item "Padrão ThinQQ" (sem core: lista mínima);
  o XML tem `dropDown[@id='ddStyle'][@tag='style'][@getItemCount='Rx_StyleCount'][@getItemLabel='Rx_StyleLabel'][@getItemID='Rx_StyleId'][@getSelectedItemID='Rx_GetStyleSel'][@onAction='Rx_OnStyle'][@getEnabled='Rx_GetEnabled']`,
  não tem `ddPalette`, e `button[@id='btnStyles'][@onAction='Rx_OnStyles']` vem logo depois de `ddLegend`. Acrescente as assinaturas novas ao
  `ExpectedSig` do `TestRibbon` (`getItemCount` = `(control, ByRef count)`, `getItemLabel`/`getItemID` = `(control, index As Integer, ByRef ...)`).

- [ ] **Passo 2: `modStyleUi.bas`**:

```vb
Attribute VB_Name = "modStyleUi"
Option Explicit
' Textos e listas da janela Estilos (1.12.0). Puras.
Private Const STYLEUI_COLORS As String = "Preto|#000000;Branco|#FFFFFF;Cinza-escuro|#404040;Cinza|#7F7F7F;Cinza-médio|#A6A6A6;Cinza-claro|#D9D9D9;Azul-marinho|#1F3864;Azul|#4472C4;Azul-claro|#9DC3E6;Verde|#70AD47;Verde-escuro|#548235;Verde-claro|#A9D18E;Laranja|#ED7D31;Amarelo|#FFC000;Vermelho|#E8173C;Vinho|#C00000;Roxo|#7030A0;Turquesa|#009688;Âmbar|#EDA224;Salmão|#F4B183"

Public Function StyleUi_IsColorKey(key As String) As Boolean
    Select Case key
        Case "nome", "fonte", "fonte.tamanho", "rotulo.tamanho", "espaco", "rotulo.negrito", "eixo.categoria.linha", "primeira.padrao", "ultima.padrao", "conector.traco", "conector.espessura"
        Case Else: StyleUi_IsColorKey = (Len(StyleDef_Validate(key, "#000000")) > 0)
    End Select
End Function

Public Function StyleUi_RoleLabel(key As String) As String
    Select Case key
        Case "nome": StyleUi_RoleLabel = "Nome"
        Case "fonte": StyleUi_RoleLabel = "Fonte"
        Case "fonte.tamanho": StyleUi_RoleLabel = "Tamanho da fonte"
        Case "total": StyleUi_RoleLabel = "Totais"
        Case "aumento": StyleUi_RoleLabel = "Aumentos (cascata)"
        Case "reducao": StyleUi_RoleLabel = "Reduções (cascata)"
        Case "primeira": StyleUi_RoleLabel = "Primeira barra"
        Case "primeira.padrao": StyleUi_RoleLabel = "Primeira barra: hachura"
        Case "primeira.padrao.cor": StyleUi_RoleLabel = "Primeira barra: cor da hachura"
        Case "primeira.borda": StyleUi_RoleLabel = "Primeira barra: borda"
        Case "ultima": StyleUi_RoleLabel = "Última barra"
        Case "ultima.padrao": StyleUi_RoleLabel = "Última barra: hachura"
        Case "ultima.padrao.cor": StyleUi_RoleLabel = "Última barra: cor da hachura"
        Case "ultima.borda": StyleUi_RoleLabel = "Última barra: borda"
        Case "rotulo.cor": StyleUi_RoleLabel = "Rótulos: cor (auto = contraste)"
        Case "rotulo.negrito": StyleUi_RoleLabel = "Rótulos: negrito"
        Case "rotulo.tamanho": StyleUi_RoleLabel = "Rótulos: tamanho"
        Case "conector.cor": StyleUi_RoleLabel = "Conectores: cor"
        Case "conector.traco": StyleUi_RoleLabel = "Conectores: traço"
        Case "conector.espessura": StyleUi_RoleLabel = "Conectores: espessura (pt)"
        Case "seta.cor": StyleUi_RoleLabel = "Setas: cor"
        Case "seta.rotulo.fundo": StyleUi_RoleLabel = "Setas: fundo do rótulo"
        Case "seta.rotulo.borda": StyleUi_RoleLabel = "Setas: borda do rótulo"
        Case "eixo.categoria.linha": StyleUi_RoleLabel = "Linha do eixo de categorias"
        Case "espaco": StyleUi_RoleLabel = "Espaço entre colunas (%)"
        Case Else
            If Left$(key, 5) = "serie" Then StyleUi_RoleLabel = "Série " & Mid$(key, 6)
    End Select
End Function

' Rotulos da lista de escolha do papel (Empty = so digitar). Pura.
Public Function StyleUi_PickItems(key As String) As Variant
    Dim a() As String, p() As String, i As Long
    Select Case key
        Case "primeira.padrao", "ultima.padrao"
            StyleUi_PickItems = Array("nenhum", "diagonal_clara_para_cima", "diagonal_escura_para_cima", "diagonal_larga_para_cima", "diagonal_clara_para_baixo", "diagonal_escura_para_baixo", "diagonal_larga_para_baixo", "horizontal_escura", "vertical_escura", "xadrez_pequeno", "pontilhado_50", "grade_pequena")
        Case "conector.traco": StyleUi_PickItems = Array("solido", "pontos", "tracos")
        Case "rotulo.negrito", "eixo.categoria.linha": StyleUi_PickItems = Array("0 (não)", "1 (sim)")
        Case Else
            If StyleUi_IsColorKey(key) Then
                p = Split(STYLEUI_COLORS, ";")
                ReDim a(0 To UBound(p))
                For i = 0 To UBound(p)
                    a(i) = Replace(p(i), "|", " ")
                Next
                StyleUi_PickItems = a
            End If
    End Select
End Function

' Valor (o que vai para o estilo) do item index da lista. Pura.
Public Function StyleUi_PickValue(key As String, index As Long) As String
    Dim items As Variant, s As String
    items = StyleUi_PickItems(key)
    If Not IsArray(items) Then Exit Function
    If index < LBound(items) Or index > UBound(items) Then Exit Function
    s = CStr(items(index))
    If StyleUi_IsColorKey(key) Then
        StyleUi_PickValue = Mid$(s, InStrRev(s, " ") + 1)
    ElseIf key = "rotulo.negrito" Or key = "eixo.categoria.linha" Then
        StyleUi_PickValue = Left$(s, 1)
    Else
        StyleUi_PickValue = s
    End If
End Function
```
  (os índices do teste: cor 9 = "Verde #70AD47"; padrão 3 = `diagonal_larga_para_cima`.)

- [ ] **Passo 3: formulário** — `frmStyles.json` (Excel, modal **não**: `Core_StyleManage` usa `Show vbModeless` para o usuário poder escolher outro gráfico
  com a janela aberta):

```json
{
  "name": "frmStyles",
  "caption": "Estilos de gráfico do ThinQQ",
  "width": 620,
  "height": 400,
  "controls": [
    { "type": "Label", "name": "lblStyles", "Left": 12, "Top": 8, "Width": 170, "Height": 14, "Caption": "Estilos", "fontName": "Segoe UI", "fontSize": 9, "fontBold": true },
    { "type": "ListBox", "name": "lstStyles", "Left": 12, "Top": 24, "Width": 170, "Height": 250, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "Label", "name": "lblRoles", "Left": 192, "Top": 8, "Width": 300, "Height": 14, "Caption": "Papel e valor", "fontName": "Segoe UI", "fontSize": 9, "fontBold": true },
    { "type": "ListBox", "name": "lstRoles", "Left": 192, "Top": 24, "Width": 300, "Height": 250, "ColumnCount": 2, "ColumnWidths": "170 pt;120 pt", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "Label", "name": "lblName", "Left": 12, "Top": 284, "Width": 170, "Height": 12, "Caption": "Nome do estilo:", "fontName": "Segoe UI", "fontSize": 8 },
    { "type": "TextBox", "name": "txtName", "Left": 12, "Top": 298, "Width": 170, "Height": 20, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "Label", "name": "lblValue", "Left": 192, "Top": 284, "Width": 120, "Height": 12, "Caption": "Valor:", "fontName": "Segoe UI", "fontSize": 8 },
    { "type": "TextBox", "name": "txtValue", "Left": 192, "Top": 298, "Width": 100, "Height": 20, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "ComboBox", "name": "cboPick", "Left": 296, "Top": 298, "Width": 150, "Height": 20, "Style": 2, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "Label", "name": "lblSwatch", "Left": 450, "Top": 298, "Width": 20, "Height": 20, "Caption": "", "BorderStyle": 1 },
    { "type": "CommandButton", "name": "btnSetValue", "Left": 474, "Top": 296, "Width": 72, "Height": 24, "Caption": "Usar valor", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnNewFromChart", "Left": 504, "Top": 24, "Width": 104, "Height": 24, "Caption": "Novo do gráfico", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnDuplicate", "Left": 504, "Top": 52, "Width": 104, "Height": 24, "Caption": "Duplicar", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnSave", "Left": 504, "Top": 80, "Width": 104, "Height": 24, "Caption": "Salvar", "fontName": "Segoe UI", "fontSize": 9, "fontBold": true },
    { "type": "CommandButton", "name": "btnDelete", "Left": 504, "Top": 108, "Width": 104, "Height": 24, "Caption": "Excluir", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnImport", "Left": 504, "Top": 144, "Width": 104, "Height": 24, "Caption": "Importar...", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnExport", "Left": 504, "Top": 172, "Width": 104, "Height": 24, "Caption": "Exportar...", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "ComboBox", "name": "cboScope", "Left": 504, "Top": 210, "Width": 104, "Height": 20, "Style": 2, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnApply", "Left": 504, "Top": 234, "Width": 104, "Height": 24, "Caption": "Aplicar", "fontName": "Segoe UI", "fontSize": 9, "fontBold": true },
    { "type": "CheckBox", "name": "chkDefault", "Left": 12, "Top": 330, "Width": 300, "Height": 18, "Caption": "Usar este estilo nos novos gráficos", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "Label", "name": "lblStatus", "Left": 12, "Top": 356, "Width": 480, "Height": 28, "Caption": "", "fontName": "Segoe UI", "fontSize": 9, "WordWrap": true },
    { "type": "CommandButton", "name": "btnClose", "Left": 504, "Top": 352, "Width": 104, "Height": 24, "Caption": "Fechar", "Cancel": true, "fontName": "Segoe UI", "fontSize": 9 }
  ]
}
```
  `frmStyles.vba`:

```vb
Option Explicit
' Estilos de grafico (1.12.0): lista, edicao por papel (valor digitado ou escolhido; validado por StyleDef_Validate), salvar no deposito do
' usuario, importar/exportar, aplicar e padrao para novos. Logica pura em modStyleDef/modStyleStore/modStyleUi; aqui so a tela.
Private mEdit As Object                 ' Dictionary do estilo em edicao
Private mKeys As Variant
Private mLoading As Boolean

Public Sub Prepare()
    mLoading = True
    mKeys = StyleDef_Keys()
    cboScope.Clear
    cboScope.AddItem "Gráfico selecionado"
    cboScope.AddItem "Planilha"
    cboScope.AddItem "Pasta"
    cboScope.ListIndex = 0
    LoadNames Core_GetStyleName()
    mLoading = False
End Sub

Private Sub LoadNames(selectName As String)
    Dim n As Variant, i As Long
    lstStyles.Clear
    For Each n In Split(StyleStore_Names(ActiveWorkbook), "|")
        lstStyles.AddItem CStr(n)
        If StrComp(CStr(n), selectName, vbTextCompare) = 0 Then i = lstStyles.ListCount - 1
    Next
    If lstStyles.ListCount > 0 Then lstStyles.ListIndex = i
End Sub

Private Sub ShowStyle(name As String)
    Dim ign As Long, k As Variant, i As Long
    Set mEdit = StyleDef_Parse(StyleStore_Get(name, ActiveWorkbook, True), ign)
    If mEdit Is Nothing Then Set mEdit = StyleDef_Parse(STYLEDEF_HEADER, ign)
    txtName.Text = name
    chkDefault.Value = (StrComp(StyleStore_DefaultName(), name, vbTextCompare) = 0)
    lstRoles.Clear
    For Each k In mKeys
        If CStr(k) <> "nome" Then
            lstRoles.AddItem StyleUi_RoleLabel(CStr(k))
            i = lstRoles.ListCount - 1
            lstRoles.List(i, 1) = StyleDef_Get(mEdit, CStr(k))
        End If
    Next
    btnDelete.Enabled = Not StyleDef_IsBuiltin(name)
End Sub

Private Function CurrentKey() As String
    If lstRoles.ListIndex < 0 Then Exit Function
    CurrentKey = CStr(mKeys(lstRoles.ListIndex + 1))          ' mKeys(0) = "nome" (fora da lista)
End Function

Private Sub lstStyles_Click()
    If lstStyles.ListIndex >= 0 Then ShowStyle CStr(lstStyles.List(lstStyles.ListIndex))
End Sub

Private Sub lstRoles_Click()
    Dim k As String, items As Variant, i As Long
    k = CurrentKey()
    If Len(k) = 0 Then Exit Sub
    txtValue.Text = StyleDef_Get(mEdit, k)
    cboPick.Clear
    items = StyleUi_PickItems(k)
    If IsArray(items) Then
        For i = LBound(items) To UBound(items)
            cboPick.AddItem CStr(items(i))
        Next
    End If
    Swatch
End Sub

Private Sub cboPick_Change()
    If cboPick.ListIndex >= 0 Then txtValue.Text = StyleUi_PickValue(CurrentKey(), cboPick.ListIndex)
    Swatch
End Sub

Private Sub Swatch()
    Dim c As Long
    c = StyleDef_ColorOf(txtValue.Text)
    If c >= 0 Then lblSwatch.BackColor = c Else lblSwatch.BackColor = RGB(240, 240, 240)
End Sub

Private Sub txtValue_Change()
    Swatch
End Sub

Private Sub btnSetValue_Click()
    Dim k As String, v As String
    k = CurrentKey()
    If Len(k) = 0 Then lblStatus.Caption = "Escolha um papel na lista.": Exit Sub
    If Len(Trim$(txtValue.Text)) = 0 Then
        If mEdit.Exists(k) Then mEdit.Remove k
        lstRoles.List(lstRoles.ListIndex, 1) = ""
        lblStatus.Caption = "Valor removido (o ThinQQ usa o padrão)."
        Exit Sub
    End If
    v = StyleDef_Validate(k, txtValue.Text)
    If Len(v) = 0 Then lblStatus.Caption = "Valor inválido para '" & StyleUi_RoleLabel(k) & "'.": Exit Sub
    mEdit(k) = v
    lstRoles.List(lstRoles.ListIndex, 1) = v
    lblStatus.Caption = ""
End Sub

Private Sub btnSave_Click()
    Dim r As String
    r = StyleStore_SaveUser(txtName.Text, StyleDef_Serialize(mEdit))
    If Left$(r, 1) = "!" Then lblStatus.Caption = Mid$(r, 2): Exit Sub
    If chkDefault.Value Then StyleStore_SetDefaultName StyleDef_SafeName(txtName.Text)
    lblStatus.Caption = "Estilo salvo: " & StyleDef_SafeName(txtName.Text) & "."
    LoadNames StyleDef_SafeName(txtName.Text)
End Sub

Private Sub btnDuplicate_Click()
    txtName.Text = txtName.Text & " (cópia)"
    btnDelete.Enabled = True
    lblStatus.Caption = "Mude o nome se quiser e clique em Salvar."
End Sub

Private Sub btnNewFromChart_Click()
    Dim r As String, nm As String
    nm = StyleDef_SafeName(txtName.Text)
    If Len(nm) = 0 Or StyleDef_IsBuiltin(nm) Then nm = "Meu estilo"
    r = Core_StyleFromChart(nm)
    If Len(r) > 0 Then lblStatus.Caption = r: Exit Sub
    lblStatus.Caption = "Estilo criado a partir do gráfico selecionado: " & nm & "."
    LoadNames nm
End Sub

Private Sub btnDelete_Click()
    Dim r As String
    r = StyleStore_DeleteUser(txtName.Text)
    If Left$(r, 1) = "!" Then lblStatus.Caption = Mid$(r, 2): Exit Sub
    lblStatus.Caption = "Estilo excluído."
    LoadNames ""
End Sub

Private Sub btnImport_Click()
    Core_StyleImport
    LoadNames txtName.Text
End Sub

Private Sub btnExport_Click()
    Core_StyleExport txtName.Text
End Sub

Private Sub btnApply_Click()
    Dim scope As String
    scope = Choose(cboScope.ListIndex + 1, "chart", "sheet", "book")
    Core_StyleApply StyleDef_SafeName(txtName.Text), scope
    If chkDefault.Value Then StyleStore_SetDefaultName StyleDef_SafeName(txtName.Text)
    lblStatus.Caption = "Aplicado."
End Sub

Private Sub chkDefault_Click()
    If mLoading Then Exit Sub
    If chkDefault.Value Then
        StyleStore_SetDefaultName StyleDef_SafeName(txtName.Text)
    ElseIf StrComp(StyleStore_DefaultName(), StyleDef_SafeName(txtName.Text), vbTextCompare) = 0 Then
        StyleStore_SetDefaultName ""
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
```
  `modCoreApi.bas`:

```vb
' Janela Estilos (1.12.0), nao modal. Build Test: nada (formulario). Nunca propaga erro.
Public Sub Core_StyleManage()
    On Error GoTo Fail
    If BUILD_CONFIG = "Test" Then Exit Sub
    frmStyles.Prepare
    frmStyles.Show vbModeless
    Exit Sub
Fail:
    Errors_Handle "Core_StyleManage"
End Sub
```
  `TestForms.bas` do core: `AssertEqual 2, CheckFormCaptions("core")`.

- [ ] **Passo 4: faixa do Excel** — no `grpChart`, troque o `dropDown` `ddPalette` por

```xml
          <dropDown id="ddStyle" label="Estilo" tag="style" getItemCount="Rx_StyleCount" getItemLabel="Rx_StyleLabel" getItemID="Rx_StyleId"
                    getSelectedItemID="Rx_GetStyleSel" onAction="Rx_OnStyle" getEnabled="Rx_GetEnabled" screentip="Estilo do gráfico"
                    supertip="Cores por papel (séries, aumentos, reduções, totais, primeira e última barra com hachura), conectores, setas, rótulos, fonte e espaço. Os embutidos Padrão ThinQQ, Tons de cinza e Cores do tema são as paletas de antes; Cascata MC e Cascata DDE são exemplos; os seus aparecem aqui depois de salvos em Estilos..."/>
```
  e, logo depois de `ddLegend`, `<button id="btnStyles" label="Estilos..." image="tqq_options16" onAction="Rx_OnStyles" screentip="Estilos de gráfico" supertip="Cria, edita, importa, exporta e aplica estilos (ao gráfico, à planilha ou à pasta) e escolhe o estilo dos gráficos novos. 'Novo do gráfico' guarda as cores do gráfico selecionado, inclusive as pintadas à mão."/>`.
  Em `modRibbon.bas` (cache por ciclo de Invalidate, como `mSupports`):

```vb
Private mStyleNames As Variant                 ' nomes dos estilos do ciclo de Invalidate (Rx_StyleNamesFor)

' Lista de nomes a partir do texto do core ("a|b|c"); vazio (core anterior ou ausente) = so "Padrão ThinQQ". Pura.
Public Function Rx_StyleNamesFor(raw As String) As Variant
    If Len(raw) = 0 Then Rx_StyleNamesFor = Array("Padrão ThinQQ") Else Rx_StyleNamesFor = Split(raw, "|")
End Function

Private Function StyleNames() As Variant
    Dim v As Variant
    If IsEmpty(mStyleNames) Then
        If Not Loader_CoreWorkbook() Is Nothing Then v = Loader_RunCore("Core_StyleNames")
        mStyleNames = Rx_StyleNamesFor(CStr(v))
    End If
    StyleNames = mStyleNames
End Function

Public Sub Rx_StyleCount(control As IRibbonControl, ByRef count)
    On Error Resume Next
    count = 0
    count = UBound(StyleNames()) + 1
End Sub

Public Sub Rx_StyleLabel(control As IRibbonControl, index As Integer, ByRef label)
    On Error Resume Next
    label = CStr(StyleNames()(index))
End Sub

Public Sub Rx_StyleId(control As IRibbonControl, index As Integer, ByRef id)
    On Error Resume Next
    id = "st" & index
End Sub

Public Sub Rx_GetStyleSel(control As IRibbonControl, ByRef itemId)
    Dim n As String, a As Variant, i As Long
    On Error Resume Next
    itemId = "st0"
    If Loader_CoreWorkbook() Is Nothing Then Exit Sub
    n = CStr(Loader_RunCore("Core_GetStyleName"))
    a = StyleNames()
    For i = 0 To UBound(a)
        If StrComp(CStr(a(i)), n, vbTextCompare) = 0 Then itemId = "st" & i: Exit Sub
    Next
End Sub

Public Sub Rx_OnStyle(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_StyleApply", Array(CStr(StyleNames()(selectedIndex)), "chart")
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_OnStyles(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_StyleManage"
    Loader_InvalidateRibbon
End Sub
```
  e em `ResetRibbonCaches`: `mStyleNames = Empty`. `Rx_OptionValueFor`/`Rx_SelectedIdFor` deixam de tratar `palette` (sem `ddPalette`); o core continua
  aceitando `Core_SetOption "palette", ...` (compatibilidade). Test-Ribbon: o conjunto de callbacks lido do XML ganha os novos; acrescente
  `Check ($null -eq $xx.SelectSingleNode("//u:dropDown[@id='ddPalette']", $xns)) 'Excel: Paleta virou Estilo'`.
- [ ] **Passo 5:** suíte `-Target all -LoaderSwap` → 0; commits `feat: janela Estilos e lista Estilo na faixa do Excel`, `test: estilos na faixa`.

---

### Tarefa S5: Estilo no PowerPoint — formatação acompanha a sincronização e menu "Estilo" na aba contextual (API 7) — modelo: most-capable

**Files:** Modify `src/core/modules/modPptSync.bas` (`PptSync_ApplyFormat`, chamada em `PptSync_SyncShape`, `CORE_PPT_API = 7`), `src/core/modules/modCoreApi.bas`
(`PptCmd_One` "style", lista de comandos), `src/ppt/modules/modPptTqq.bas` (`PPT_API_STYLE`, `PPT_MSG_OLD_CORE_STYLE`, `PptTqq_StyleMenuXml`), `src/ppt/modules/modPptRibbon.bas`
(`PptRx_GetStyleMenu`, `PptRx_OnStyle`), `src/ppt/ribbon/customUI14.xml`; Test `tests/vba/core/TestPptSync.bas`, `tests/vba/ppt/TestPptAddin.bas`, `TestXlFake.cls`, `Test-Ribbon.ps1`.

- [ ] **Passo 1 (RED):** core

```vb
' 1.12.0 (S5): a formatacao do estilo (cor das series e hachura da 1a barra) chega ao grafico do slide na atualizacao; comando "style" pelo PowerPoint.
Public Sub Test_PptSync_StyleFormat_ReachesSlide_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, own As New Collection, en As Long, ed As String, msg As String
    Dim ws As Worksheet, c As Chart
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\pptstyle.xlsx")
    msg = Core_PptCommand(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), "style", "Tons de cinza")
    AssertEqual "", msg
    msg = Core_PptCommand(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), "style", "Cascata MC")
    AssertEqual "", msg
    Set c = Core_ActiveChart()
    If c Is Nothing Then Set c = rng.Worksheet.ChartObjects(1).Chart
    AssertEqual "Cascata MC", Registry_GetOption(c, "style", "")
    AssertEqual RGB(112, 173, 71), CLng(s1.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB), "serie1 do estilo no slide"
    AssertTrue InStr(Core_PptCommand(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), "style", "Nao Existe"), "Estilo não encontrado") > 0
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set s1 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub
```
  `Test_Core_PptApiVersion` → 7. `.ppam`: `TestXlFake.Run` ganha `Case "Core_StyleNames": Run = StyleNamesResult` (`Public StyleNamesResult As String`); teste
  `Test_PptTqq_StyleMenuXml`: `PptTqq_StyleMenuXml("A|B <&>")` contém `<button id="pst0" label="A" tag="A" onAction="PptRx_OnStyle"/>` e o 2º com `&lt;&amp;&gt;`
  escapados; `PptTqq_StyleMenuXml("")` contém um botão desabilitado com `PPT_MSG_OLD_CORE_STYLE`; e `PptTqq_CommandShapes` com core API 6 e `cmd = "style"`
  mostra `PPT_MSG_OLD_CORE_STYLE` sem chamar `Core_PptCommand` (`InStr(f.Calls, "Core_PptCommand") = 0`).

- [ ] **Passo 2: core** — `modPptSync.bas`: `CORE_PPT_API = 7` (`; 7 = Core_StyleNames e comando "style" (1.12.0)`) e

```vb
' Estilo (1.12.0; spec 2026-10-23 secao 8.5): LinkFormat.Update traz so os dados; com a opcao style no grafico do Excel, copia para o grafico do
' slide (pc) a formatacao das series (preenchimento solido/hachura e borda), dos pontos 1 e n (primeira/ultima barra) e o negrito/cor dos
' rotulos no nivel da serie, e a fonte do grafico. Mesma ordem de series dos dois lados (o slide e uma copia do grafico do Excel). Nunca
' propaga erro (registrado): o resto da sincronizacao segue.
Public Sub PptSync_ApplyFormat(pc As Object, cht As Chart)
    Dim i As Long, xs As Series, ps As Object, n As Long
    On Error GoTo Fail
    If Len(Registry_GetOption(cht, "style", "")) = 0 Then Exit Sub
    On Error Resume Next
    pc.ChartArea.Format.TextFrame2.TextRange.Font.Name = cht.ChartArea.Font.Name
    pc.ChartArea.Format.TextFrame2.TextRange.Font.Size = cht.ChartArea.Font.Size
    Err.Clear
    For i = 1 To cht.SeriesCollection.Count
        Set xs = cht.SeriesCollection(i)
        Set ps = Nothing
        Set ps = pc.SeriesCollection(i)
        If Not ps Is Nothing Then
            PptSync_CopyFill xs.Format.Fill, ps.Format.Fill
            ps.Format.Line.Visible = xs.Format.Line.Visible
            n = xs.Points.Count
            If n >= 1 Then PptSync_CopyPoint xs.Points(1), ps.Points(1)
            If n >= 2 Then PptSync_CopyPoint xs.Points(n), ps.Points(n)
            If xs.HasDataLabels Then
                ps.DataLabels.Font.Bold = xs.DataLabels.Font.Bold
                ps.DataLabels.Font.Color = xs.DataLabels.Font.Color
            End If
        End If
        Err.Clear
    Next
    Exit Sub
Fail:
    Errors_Log "PptSync_ApplyFormat" & vbTab & Err.Number & vbTab & Err.Description
End Sub

Private Sub PptSync_CopyFill(src As FillFormat, dst As Object)
    On Error Resume Next
    If src.Visible = msoFalse Then dst.Visible = msoFalse: Exit Sub
    dst.Visible = msoTrue
    If src.Type = msoFillPatterned Then
        dst.Patterned src.Pattern
        dst.ForeColor.RGB = src.ForeColor.RGB
        dst.BackColor.RGB = src.BackColor.RGB
    Else
        dst.Solid
        dst.ForeColor.RGB = src.ForeColor.RGB
    End If
    Err.Clear
End Sub

Private Sub PptSync_CopyPoint(xp As Point, pp As Object)
    On Error Resume Next
    PptSync_CopyFill xp.Format.Fill, pp.Format.Fill
    pp.Format.Line.Visible = xp.Format.Line.Visible
    If xp.Format.Line.Visible = msoTrue Then
        pp.Format.Line.ForeColor.RGB = xp.Format.Line.ForeColor.RGB
        pp.Format.Line.Weight = xp.Format.Line.Weight
    End If
    Err.Clear
End Sub
```
  e em `PptSync_SyncShape`, depois de `PptSync_Apply pc, snaps, cht, dx, dy`: `PptSync_ApplyFormat pc, cht`. `modCoreApi.bas`: em `PptCommandImpl` o `Select Case cmd`
  aceita `"style"`; em `PptCmd_One` acrescente `Case "style": Style_ApplyNamed c, arg, True` e, no `Fail:`, `vbObjectError + 7140` na lista das recusas com texto;
  `Core_StyleNames` já existe (S2).

- [ ] **Passo 3: `.ppam`** — `modPptTqq.bas`: `Public Const PPT_API_STYLE As Long = 7` e `Public Const PPT_MSG_OLD_CORE_STYLE As String = "Atualize o ThinQQ no Excel
  (1.12.0 ou mais novo) para usar os estilos pelo PowerPoint."`; em `PptTqq_CommandShapes`, no ramo `ElseIf PptTqq_CoreApi(xl, core) < PPT_API_CMD Then`
  acrescente antes `ElseIf cmd = "style" And PptTqq_CoreApi(xl, core) < PPT_API_STYLE Then msg = PPT_MSG_OLD_CORE_STYLE` (a API é lida uma vez numa variável
  `api` para não chamar o Excel duas vezes); e

```vb
Private Function PptTqq_XmlEsc(s As String) As String
    PptTqq_XmlEsc = Replace(Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;")
End Function

' Conteudo do menu dinamico "Estilo" (aba contextual): um botao por nome (tag = nome). Sem nomes (Excel sem o core 1.12.0): um botao
' desabilitado com o aviso. Pura.
Public Function PptTqq_StyleMenuXml(names As String) As String
    Dim s As String, a() As String, i As Long
    s = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"
    If Len(names) = 0 Then
        s = s & "<button id=""pst_none"" label=""" & PptTqq_XmlEsc(PPT_MSG_OLD_CORE_STYLE) & """ enabled=""false""/>"
    Else
        a = Split(names, "|")
        For i = 0 To UBound(a)
            If i >= 60 Then Exit For
            s = s & "<button id=""pst" & i & """ label=""" & PptTqq_XmlEsc(a(i)) & """ tag=""" & PptTqq_XmlEsc(a(i)) & """ onAction=""PptRx_OnStyle""/>"
        Next
    End If
    PptTqq_StyleMenuXml = s & "</menu>"
End Function

' Nomes dos estilos do core (API 7) no Excel JA aberto; "" sem Excel/core/API. Nunca propaga.
Public Function PptTqq_StyleNames() As String
    Dim xl As Object, created As Boolean, core As String
    On Error GoTo Done
    Set xl = PptTqq_GetExcel(created, True)
    If xl Is Nothing Then Exit Function
    core = PptTqq_EnsureCore(xl, created)
    If Len(core) > 0 Then
        If PptTqq_CoreApi(xl, core) >= PPT_API_STYLE Then PptTqq_StyleNames = CStr(xl.Run("'" & core & "'!Core_StyleNames"))
    End If
Done:
    On Error Resume Next
    PptTqq_Release xl, created
End Function
```
  `modPptRibbon.bas`:

```vb
Public Sub PptRx_GetStyleMenu(control As IRibbonControl, ByRef content)
    On Error Resume Next
    content = PptTqq_StyleMenuXml(PptTqq_StyleNames())
End Sub

Public Sub PptRx_OnStyle(control As IRibbonControl)
    On Error Resume Next
    PptTqq_CommandShapes Nothing, PptTqq_SelectedTagged(), "style", PptRx_TagOf(control)
End Sub
```
  XML: na aba contextual, grupo `cgrpChart`, depois de `cddLegend`:
  `<dynamicMenu id="cdmStyle" label="Estilo" imageMso="ChartColorsGallery" getContent="PptRx_GetStyleMenu" getEnabled="PptRx_GetElemEnabled" screentip="Estilo do gráfico" supertip="Aplica um estilo do ThinQQ ao gráfico selecionado (a lista vem do Excel aberto com o ThinQQ 1.12.0)."/>`
  e o mesmo (`id="dmPptStyle"`) no grupo `grpPptChart`. **`imageMso="ChartColorsGallery"`**: se o Test-Ribbon da lista permitida não o tiver, use
  `imageMso="FormatPainter"` (já usado no Office há décadas) e registre. Test-Ribbon: `dynamicMenu` com `getContent='PptRx_GetStyleMenu'` nos dois grupos.
- [ ] **Passo 4:** suíte `-Ppt` → 0; commit `feat: estilos chegam ao slide e menu Estilo no PowerPoint (API 7)`.

---

### Tarefa S6: Documentação da 1.12.0 — modelo: standard

- [ ] Tutorial: seção nova `#estilos` ("Estilos de gráfico", com uma figura em HTML/CSS da cascata do exemplo — barras como `div` com `background` e hachura
  `repeating-linear-gradient`, sem PNG novo), passo a passo "reproduzir o gráfico do exemplo" (inserir cascata, Estilo > Cascata MC, setas de diferença, balão
  feito à mão até a 1.14), papéis por tipo (tabela), formato do arquivo `.tqqestilo` com o exemplo completo, importar/exportar, padrão para novos, o que viaja
  para o PowerPoint, limites (Gantt só fonte/rótulos/espaço; rótulo dentro da barra da cascata fica no backlog; formatação feita no PowerPoint num gráfico com
  estilo é sobrescrita na atualização). Galeria (`#galeria`): menção aos estilos. Referência da faixa: "Paleta" → "Estilo" + "Estilos...". Teto do tutorial.
- [ ] README (recursos, arquitetura `modStyleDef/Store/Apply/Ui`, `frmStyles`, `CustomXMLParts`, API 7, verificações manuais **119–125**, limitações), CHANGELOG
  `[Não lançado]`, `release/notes.txt`, Test-Docs (rótulos "Estilo", "Estilos...", legendas do `frmStyles`, âncora `#estilos`, mensagens novas: "Estilo não
  encontrado: ...", "Estilo inválido: ...", "O arquivo não é um estilo do ThinQQ válido (...)", "Estilo importado: ...", "Estilos embutidos não podem ser
  excluídos.", "'X' é um estilo embutido do ThinQQ: use outro nome (Duplicar).", `PPT_MSG_OLD_CORE_STYLE`). pwsh 7 e 5.1 → 0. Commit `docs: 1.12.0 (estilos de grafico)`.

### Tarefa S7: Suíte final, revisão e verificação manual — modelo: standard (controlador + humano)

- [ ] Suíte `-Ppt` com o PowerPoint do usuário fechado → 0 e 0 pulados; revisão final do ramo (Review Focus 5–7; segurança da importação; desempenho do
  refresh com estilo); onda única; verificação manual 119–125 (o usuário reproduz o gráfico do exemplo); push `feat/estilos`.

# PARTE X — Pontos de entrada como o think-cell (1.13.0)

Spec §8.4. Depende das sondas manuais da E0; a X2 é ela mesma a sonda P-U6 automatizada (teste `_WhenAvailable`).

### Tarefa E0: Sondas manuais P-U3/P-U4/P-U5 (controlador + usuário) — modelo: standard

- [ ] Combinar com o usuário uma janela de ~30 min com o PowerPoint/Excel dele **fechados** (sessão visível, com a permissão dele; nada de gravação de tela
  em andamento). Montar na cópia `.xlam`/`.ppam` de sonda (padrão `px2.ps1` de `.superpowers/sdd/2026-10-23-thinkcell-ux/design.md` §6.1: um suplemento por
  `idMso` candidato, cada um com a aba "P" e `onLoad` que grava um nome/etiqueta), abrir o Excel/PowerPoint **visível** pelo script, anotar se a aba "P"
  aparece e se `onLoad` foi chamado. Candidatos: PowerPoint `ContextMenuChartArea`, `ContextMenuChartPlotArea`, `ContextMenuShape`, `ContextMenuObjectsGroup`;
  Excel `ContextMenuChartArea`, `ContextMenuCell`; e `TabSetChartTools` nos dois (P-U3, se a U6 não confirmou).
- [ ] P-U5 (Excel): pasta de teste com um gráfico do ThinQQ e um módulo de sonda que liga `BeforeRightClick` (`Cancel = True` + `CommandBars("tqq_probe").ShowPopup`):
  o usuário clica com o botão direito no gráfico (selecionado e não selecionado); anotar se o menu aparece e se o menu do Excel some; testar
  `Application.CommandBars.ExecuteMso "ChartFormatSelection"` (abre o painel Formatar?).
- [ ] Registrar em `.superpowers/sdd/2026-10-23-thinkcell-ux/probes-X-report.md` e no ledger; encerrar tudo pelo PID; apagar as pastas de sonda. A X3 só
  inclui o que foi confirmado.

### Tarefa X1: "Vincular ao PowerPoint" — uma galeria na aba ThinQQ e na guia Inserir do Excel — modelo: standard

**Files:** Modify `src/loader/ribbon/customUI14.xml`, `src/loader/modules/modRibbon.bas`; Test `tests/vba/loader/TestRibbon.bas`, `tests/ps/Test-Ribbon.ps1`.

**Interfaces:** `Rx_OnLinkToPpt(control As IRibbonControl)`; `Rx_LinkProcFor(tag As String) As String` (testável: tipo ⇒ `Core_InsertToPowerPoint`;
`table` ⇒ `Core_PptInsertTable`; `text` ⇒ `Core_PptLinkTextFromPpt`; outro ⇒ "").

- [ ] **Passo 1 (RED):** `TestRibbon.bas`:

```vb
Public Sub Test_Rx_LinkProcFor()
    AssertEqual "Core_InsertToPowerPoint", Rx_LinkProcFor("waterfall")
    AssertEqual "Core_PptInsertTable", Rx_LinkProcFor("table")
    AssertEqual "Core_PptLinkTextFromPpt", Rx_LinkProcFor("text")
    AssertEqual "", Rx_LinkProcFor("x,y")
    AssertEqual "", Rx_LinkProcFor("")
End Sub
```
  e as checagens de XML: `group[@id='grpLinkPpt']` é o 1º grupo de `tqqTab`; `menu[@id='lmnuLinkPpt'][@size='large']` com 21 botões `onAction='Rx_OnLinkToPpt'` cujos
  `tag` são, na ordem, os do `grpInsert`, mais `lbtnLinkTable` (`tag='table'`) e `lbtnLinkText` (`tag='text'`); `tab[@idMso='TabInsert']/group[@id='grpTqqInsertTab']`
  com `menu[@id='imnuLinkPpt']` igual (ids com prefixo `i`); `chkPptInsert` não existe mais; `mnuPpt` não existe mais em `grpChart` e seus itens restantes
  (`btnPptImage`, `btnPptChart`, `btnPptSync`, `btnPptFrames`) estão no `menu[@id='lmnuPptMore']` do `grpLinkPpt`.
- [ ] **Passo 2:** XML — antes de `grpInsert` em `tqqTab`:

```xml
        <!-- 1.13.0: o "Link to PowerPoint" do think-cell: um caminho so (selecione a tabela, escolha o tipo: o grafico nasce no Excel, vai
             vinculado ao slide ativo e o intervalo ganha a moldura). Ids com prefixo "l" (aba ThinQQ) e "i" (guia Inserir). -->
        <group id="grpLinkPpt" label="PowerPoint">
          <menu id="lmnuLinkPpt" label="Vincular ao PowerPoint" size="large" image="tqq_ppt" screentip="Vincular ao PowerPoint"
                supertip="Selecione a tabela e escolha o tipo: o gráfico é criado no Excel e colado vinculado no slide ativo do PowerPoint (ou numa apresentação nova), com a moldura verde no intervalo. Também: tabela vinculada e texto vinculado à célula. A pasta precisa estar salva.">
            <!-- 21 botoes: copie os <button> dos cinco menus de grpInsert, na mesma ordem, trocando o prefixo do id por "l" (btnStacked => lbtnStacked)
                 e onAction por Rx_OnLinkToPpt; nenhum leva size -->
            <menuSeparator id="lmsepLink"/>
            <button id="lbtnLinkTable" label="Tabela vinculada" tag="table" onAction="Rx_OnLinkToPpt" image="tqq_ppt16"/>
            <button id="lbtnLinkText" label="Texto: célula ao texto selecionado" tag="text" onAction="Rx_OnLinkToPpt" image="tqq_ppt16"/>
          </menu>
          <button id="btnPptSync" label="Atualizar no PowerPoint" imageMso="Refresh" onAction="Rx_OnPptSync" screentip="Atualizar no PowerPoint"/>
          <menu id="lmnuPptMore" label="Mais" image="tqq_ppt16" screentip="Mais opções do PowerPoint">
            <button id="btnPptImage" label="Copiar como imagem" onAction="Rx_OnPptImage" getEnabled="Rx_GetEnabled"/>
            <button id="btnPptChart" label="Copiar como gráfico vinculado" onAction="Rx_OnPptChart" getEnabled="Rx_GetEnabled"/>
            <button id="btnPptFrames" label="Mostrar/ocultar molduras" image="tqq_ppt16" onAction="Rx_OnPptFrames"/>
          </menu>
        </group>
```
  (os `screentip`/`supertip` dos itens movidos ficam os de hoje); remova `mnuPpt` e `sep3`/`btnPptLinkText`/`btnPptTable` antigos (viraram os dois itens da
  galeria). Depois de `</tab>` do `tqqTab`, ainda dentro de `<tabs>`:

```xml
      <tab idMso="TabInsert">
        <group id="grpTqqInsertTab" label="ThinQQ">
          <menu id="imnuLinkPpt" label="Vincular ao PowerPoint" size="large" image="tqq_ppt" screentip="Vincular ao PowerPoint (ThinQQ)">
            <!-- os mesmos 23 itens de lmnuLinkPpt com o prefixo "i" -->
          </menu>
        </group>
      </tab>
```
  `modRibbon.bas`:

```vb
' Galeria "Vincular ao PowerPoint" (1.13.0): o tipo vem do tag. Parte testavel em Rx_LinkProcFor.
Public Function Rx_LinkProcFor(tag As String) As String
    If tag = "table" Then
        Rx_LinkProcFor = "Core_PptInsertTable"
    ElseIf tag = "text" Then
        Rx_LinkProcFor = "Core_PptLinkTextFromPpt"
    ElseIf Len(tag) > 0 And InStr(tag, ",") = 0 And tag Like "[a-z]*" And Not tag Like "*[!a-z0-9]*" Then
        Rx_LinkProcFor = "Core_InsertToPowerPoint"
    End If
End Function

Public Sub Rx_OnLinkToPpt(control As IRibbonControl)
    Dim p As String
    On Error Resume Next
    p = Rx_LinkProcFor(CStr(control.Tag))
    If Len(p) = 0 Then Exit Sub
    If CoreOrWarn() Then
        If p = "Core_InsertToPowerPoint" Then RunCoreFeature p, CStr(control.Tag) Else RunCoreFeature p
    End If
    Loader_InvalidateRibbon
End Sub
```
  `Rx_OnInsert` passa a chamar sempre `Core_Insert` (a caixa de sessão saiu); mantenha `Rx_InsertProcFor`/`Rx_SetPptInsertForTest` só se testes
  antigos os usam (senão, apague-os com os testes). Test-Ribbon: `idMso` permitido `TabInsert`; ids únicos; 23 itens em cada galeria.
- [ ] **Passo 3:** suíte `-Target all -LoaderSwap` → 0; commit `feat: Vincular ao PowerPoint (galeria na aba ThinQQ e na guia Inserir)`.

### Tarefa X2: Moldura com estado e "Alterar dados" sem perder o slide (mesmo id) — modelo: most-capable

**Files:** Modify `src/core/modules/modCoreApi.bas` (`Core_RebuildFromRange` mantém o id, `RebuildKeepId`, `Core_RefreshFrames`), `src/core/modules/modLinkFrame.bas`
(`LinkFrame_SetState`, `LINKFRAME_GREEN`, `LINKFRAME_RED`); Test `tests/vba/core/TestLinkFrame.bas`, `tests/vba/core/TestPptSync.bas`.

**Interfaces:** `LinkFrame_SetState(cht As Chart, used As Boolean)`; `Core_RefreshFrames(wb As Workbook) As Long` (quantas molduras verdes; nunca propaga);
`Core_RebuildFromRange(c As Chart, rng As Range, Optional newKind As String = "") As Boolean` (o parâmetro novo é da X4; aqui só o id).

- [ ] **Passo 1 (RED):**

```vb
' TestLinkFrame.bas
Public Sub Test_Rebuild_KeepsChartId_AndFrame()
    Dim wb As Workbook, ws As Worksheet, c As Chart, id0 As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5): ws.Range("A3:C3").Value = Array("T2", 20, 6)
    ws.Range("A4:C4").Value = Array("T3", 30, 7)
    ws.Range("A1:C3").Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    id0 = Registry_IdFromChart(c)
    LinkFrame_Ensure c, ws.Range("A1:C3")
    Core_ClearLastNotice
    AssertTrue Core_RebuildFromRange(c, ws.Range("A1:C4"))
    Set c = Core_ActiveChart()
    AssertEqual id0, Registry_IdFromChart(c), "Alterar dados mantem o id (o slide continua achando o grafico)"
    AssertEqual "3", Registry_GetOption(c, "n", ""), "bloco do grafico novo, com o id antigo"
    AssertTrue Not LinkFrame_Find(c) Is Nothing, "moldura redesenhada no intervalo novo"
    AssertEqual ws.Range("A1:C4").Address, LinkFrame_Find(c).TopLeftCell.Resize(4, 3).Address
    AssertEqual 0, InStr(Core_LastNotice(), LINKFRAME_NOTE_REBUILT), "sem o aviso de grafico recriado"
    CloseTemp wb
End Sub

Public Sub Test_LinkFrame_SetState_Colors()
    Dim wb As Workbook, ws As Worksheet, c As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Ano", "V"): ws.Range("A2:B2").Value = Array("T1", 1): ws.Range("A3:B3").Value = Array("T2", 2)
    ws.Range("A1:B3").Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    LinkFrame_Ensure c, ws.Range("A1:B3")
    LinkFrame_SetState c, False
    AssertEqual LINKFRAME_RED, LinkFrame_Find(c).Line.ForeColor.RGB
    LinkFrame_SetState c, True
    AssertEqual LINKFRAME_GREEN, LinkFrame_Find(c).Line.ForeColor.RGB
    AssertEqual 0, Core_RefreshFrames(wb), "sem PowerPoint: nenhuma apresentacao usa -> vermelha"
    AssertEqual LINKFRAME_RED, LinkFrame_Find(c).Line.ForeColor.RGB
    CloseTemp wb
End Sub
```
  E em `TestPptSync.bas` (é a sonda **P-U6** automatizada):

```vb
' P-U6 (1.13.0): colado vinculado, "Alterar dados" com uma linha a mais (mesmo id e mesmo nome do ChartObject) e Atualizar: o slide acompanha.
Public Sub Test_Rebuild_KeepId_SlideStillFollows_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, own As New Collection, en As Long, ed As String, msg As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\keepid.xlsx")
    rng.Worksheet.Range("A4:C4").Value = Array("T3", 99, 8)
    AssertTrue Core_RebuildFromRange(rng.Worksheet.ChartObjects(1).Chart, rng.Worksheet.Range("A1:C4"))
    msg = Core_PptSyncShapes(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), True)
    AssertTrue InStr(msg, "Gráficos do ThinQQ atualizados: 1.") = 1, "P-U6: " & msg
    AssertAlmost 99, PptSeriesValue(s1.Chart, "Sul", 3), 0.0001, "P-U6: o slide acompanha o grafico recriado com o mesmo id"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set s1 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub
```
- [ ] **Passo 2: manter o id** — em `Core_RebuildFromRange`: guarde `oldId = Registry_IdFromChart(c)` e `oldSrc` = o intervalo da moldura antes (`hadFrame`); depois de
  `CalcSheet_Cleanup wb` (o bloco velho, órfão, já saiu) chame `RebuildKeepId co, oldId, wb` e troque o bloco `If hadFrame Then ... LINKFRAME_NOTE_REBUILT ... End If` por
  `If hadFrame Then LinkFrame_Ensure nc, data.Source`. Acrescente:

```vb
' Alterar dados/tipo (1.13.0; spec 2026-10-23 secao 8.4): o grafico NOVO assume o id do antigo (nome do ChartObject, nome definido do bloco e a
' celula A da meta), entao o slide (TQQ_CHART) e o vinculo (que acha o grafico pelo nome) continuam valendo (P-U6). O bloco antigo ja saiu pela
' limpeza (CalcSheet_Cleanup: orfao). As formulas das series apontam para as celulas do bloco, nao para o nome. Falha: registrada; o grafico fica
' com o id novo (como antes da 1.13.0) e o aviso LINKFRAME_NOTE_REBUILT volta a valer.
Private Function RebuildKeepId(co As ChartObject, oldId As String, wb As Workbook) As Boolean
    Dim newId As String, blk As Range
    On Error GoTo Fail
    newId = Mid$(co.Name, Len(ID_PREFIX) + 1)
    If Len(oldId) = 0 Or oldId = newId Then RebuildKeepId = True: Exit Function
    If CalcSheet_BlockExists(wb, oldId) Then CalcSheet_DeleteBlock wb, oldId
    If Not Registry_FindChartObject(wb, oldId) Is Nothing Then Err.Raise vbObjectError + 7150, "RebuildKeepId", "o id antigo ainda esta em uso"
    Set blk = CalcSheet_FindBlock(wb, newId)
    wb.Names(ID_PREFIX & newId).Name = ID_PREFIX & oldId
    blk.Cells(1, 1).Value = oldId
    co.Name = ID_PREFIX & oldId
    RebuildKeepId = True
    Exit Function
Fail:
    Errors_Log "RebuildKeepId" & vbTab & Err.Number & vbTab & Err.Description
End Function
```
  (se `RebuildKeepId` devolver False e havia moldura: o comportamento antigo — limpar órfãs e `Core_Notify LINKFRAME_NOTE_REBUILT`.) Testes antigos que
  afirmavam o id novo e precisam virar o comportamento novo (o gráfico continua achado; sem o aviso): `TestPptSync.Test_PptInsertChart_FromPpt_WhenAvailable`
  (o trecho "M9": `Core_PptEditData` agora acha o gráfico e `Core_LastNotice` não tem `LINKFRAME_NOTE_REBUILT`) e `TestLinkFrame.Test_LinkFrame_FollowsChangeSource`
  (passa a esperar 1 moldura no intervalo novo); procure outros com `grep -n "LINKFRAME_NOTE_REBUILT\|IdFromChart" tests/vba` e liste no relatório. Confira em `modCalcSheet.bas`
  que `CalcSheet_FindBlock` usa o nome definido `tqq_<id>` e a coluna A da meta: se usar outra chave, ajuste aqui e relate.
- [ ] **Passo 3: moldura com estado** — `modLinkFrame.bas`: `Public Const LINKFRAME_GREEN As Long = 5287936` (`RGB(0, 176, 80)`) e `Public Const LINKFRAME_RED As Long = 192`
  (`RGB(192, 0, 0)`), `LinkFrame_Draw` usa `LINKFRAME_GREEN`, e

```vb
' Cor da moldura pelo estado (como o think-cell): verde = alguma apresentacao aberta usa o grafico; vermelha = nenhuma. Nunca propaga.
Public Sub LinkFrame_SetState(cht As Chart, used As Boolean)
    Dim s As Shape
    On Error Resume Next
    Set s = LinkFrame_Find(cht)
    If s Is Nothing Then Exit Sub
    If used Then
        If s.Line.ForeColor.RGB <> LINKFRAME_GREEN Then s.Line.ForeColor.RGB = LINKFRAME_GREEN
    Else
        If s.Line.ForeColor.RGB <> LINKFRAME_RED Then s.Line.ForeColor.RGB = LINKFRAME_RED
    End If
    Err.Clear
End Sub
```
  `modCoreApi.bas`:

```vb
' Molduras da pasta pelo estado: ids dos graficos usados nas apresentacoes ABERTAS (PowerPoint ja aberto; TQQ_LINKIDX das entradas desta pasta,
' etiqueta TQQ_CHART de cada forma) = verde; o resto = vermelha. Chamado por Atualizar (Core_Refresh), Mostrar/ocultar molduras e depois de
' inserir no slide (nunca no ciclo automatico: mudar a cor altera a pasta e limparia o Desfazer). Devolve quantas verdes. Nunca propaga.
Public Function Core_RefreshFrames(wb As Workbook) As Long
    Dim used As Object, app As Object, pres As Object, e As Variant, sl As Object, shp As Object, ws As Worksheet, s As Shape, id As String, co As ChartObject
    On Error Resume Next
    Set used = CreateObject("Scripting.Dictionary")
    Set app = Ppt_GetApp(False)
    If Not app Is Nothing Then
        For Each pres In app.Presentations
            For Each e In PptIdx_Parse(PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX))
                If StrComp(CStr(e(2)), wb.FullName, vbTextCompare) = 0 Then
                    Set sl = Nothing: Set shp = Nothing
                    Set sl = pres.Slides.FindBySlideID(CLng(e(0)))
                    If Not sl Is Nothing Then Set shp = sl.Shapes(CStr(e(1)))
                    If Not shp Is Nothing Then used(PptSync_Tag(shp, PPTSYNC_TAG_CHART)) = True
                    Err.Clear
                End If
            Next
        Next
    End If
    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            If Left$(s.Name, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then
                id = Mid$(s.Name, Len(LINKFRAME_PREFIX) + 1)
                Set co = Registry_FindChartObject(wb, id)
                If Not co Is Nothing Then
                    LinkFrame_SetState co.Chart, used.Exists(id)
                    If used.Exists(id) Then Core_RefreshFrames = Core_RefreshFrames + 1
                End If
            End If
        Next
    Next
    Err.Clear
End Function
```
  chamadas: `Core_Refresh` (depois de `LinkFrame_CleanOrphans`), `Core_ToggleLinkFrames` (antes do toggle) e `Core_InsertToPowerPoint` (depois da colagem que deu certo).
- [ ] **Passo 4:** suíte `-Ppt` (sem POWERPNT do usuário: o teste P-U6 roda; se falhar, **não** force: mantenha o id novo, restaure o aviso e registre "P-U6
  negativa" no relatório — a decisão passa ao controlador). Commit `feat: Alterar dados mantem o id (o slide acompanha) e moldura com estado`.

### Tarefa X3: Botão direito e duplo clique no gráfico do ThinQQ no Excel — modelo: most-capable

**Files:** Create `src/core/modules/modChartMenu.bas`; Modify `src/core/classes/clsChartWatcher.cls`; Test `tests/vba/core/TestCoreApi.bas` (ou `TestChartMenu.bas` novo).

**Interfaces:** `Menu_Items(c As Chart) As Collection` (itens `Array(caption, parameter, isToggle, pressed, enabled)`; separador = `Array("-", "", False, False, True)`);
`Menu_DoParam(param As String)`; `Menu_Do()` (alvo do `OnAction`); `Menu_OnChartRightClick(c As Chart) As Boolean`; `Menu_OnChartDoubleClick(c As Chart) As Boolean`;
`Menu_Enabled() As Boolean` / `Menu_SetEnabledForTest(v)`; `Menu_LastBuiltForTest() As String`.

- [ ] **Passo 1 (RED):**

```vb
Public Sub Test_Menu_Items_ByKind()
    Dim wb As Workbook, c As Chart, it As Variant, caps As String
    Set wb = NewTempWorkbook()
    wb.Worksheets(1).Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    wb.Worksheets(1).Range("A2:C2").Value = Array("T1", 10, 5): wb.Worksheets(1).Range("A3:C3").Value = Array("T2", 20, 6)
    wb.Worksheets(1).Range("A1:C3").Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    For Each it In Menu_Items(c)
        caps = caps & it(0) & "=" & it(1) & IIf(it(4), "", "(off)") & ";"
    Next
    AssertTrue InStr(caps, "Totais=t:totals;") > 0, caps
    AssertTrue InStr(caps, "Seta CAGR...=a:cagr;") > 0
    AssertTrue InStr(caps, "Corte de eixo...=break;") > 0
    AssertTrue InStr(caps, "Vincular ao PowerPoint=ppt;") > 0
    AssertTrue InStr(caps, "Selecionar a tabela de origem=selsrc;") > 0
    Menu_DoParam "t:totals"
    AssertEqual "0", Registry_GetOption(c, "totals", "1"), "inverte o alternador (como a faixa)"
    Menu_DoParam "selsrc"
    AssertEqual "$A$1:$C$3", Selection.Address, "duplo clique / menu: seleciona a tabela de origem"
    Menu_DoParam "x:lixo"                                 ' parametro desconhecido: nada, sem erro
    CloseTemp wb
End Sub
```
- [ ] **Passo 2:** `modChartMenu.bas`:

```vb
Attribute VB_Name = "modChartMenu"
Option Explicit
' Menu do botao direito no grafico do ThinQQ no Excel (1.13.0; o menu de contexto do think-cell; sonda P-U5). clsChartWatcher chama
' Menu_OnChartRightClick: com a opcao ligada (registro ThinQQ\Excel\MenuDireito, padrao "1"; build Test: memoria) monta uma barra popup
' temporaria (CommandBars, medido PX1) com os comandos da faixa e cancela o menu do Excel; o ultimo item devolve o painel Formatar do Excel.
' Duplo clique seleciona a tabela de origem (o "datasheet"). Os itens chamam Menu_Do pelo OnAction com o Parameter. Nunca propaga erro.
Private Const MENU_NAME As String = "ThinQQ_Grafico"
Private mTestEnabled As Variant
Private mTestLast As String

Public Sub Menu_SetEnabledForTest(v As Variant)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestEnabled = v
End Sub

Public Function Menu_LastBuiltForTest() As String
    Menu_LastBuiltForTest = mTestLast
End Function

Public Function Menu_Enabled() As Boolean
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then
        If IsEmpty(mTestEnabled) Then Menu_Enabled = True Else Menu_Enabled = CBool(mTestEnabled)
        Exit Function
    End If
    Menu_Enabled = (GetSetting(APP_NAME, "Excel", "MenuDireito", "1") <> "0")
End Function

Private Function Menu_Item(cap As String, param As String, isToggle As Boolean, pressed As Boolean, enabled As Boolean) As Variant
    Menu_Item = Array(cap, param, isToggle, pressed, enabled)
End Function

' Itens do menu para o grafico c (capacidades pela matriz). Separador: caption "-". Nunca propaga erro.
Public Function Menu_Items(c As Chart) As Collection
    Dim out As New Collection, kind As String, k As Variant, lbl As String
    Set Menu_Items = out
    On Error Resume Next
    kind = Registry_ChartType(c)
    For Each k In Array("labels", "totals", "connectors", "vline", "ind100")
        lbl = Kinds_FeatureLabel(CStr(k))
        out.Add Menu_Item(lbl, "t:" & CStr(k), True, Registry_GetOption(c, CStr(k), "0") = "1", Kinds_SupportsChart(c, CStr(k)))
    Next
    out.Add Menu_Item("-", "", False, False, True)
    out.Add Menu_Item("Seta CAGR...", "a:cagr", False, False, Kinds_Supports(kind, "cagr"))
    out.Add Menu_Item("Seta diferença...", "a:diff", False, False, Kinds_Supports(kind, "diff"))
    out.Add Menu_Item("Delta clicando nas barras", "delta", False, False, Kinds_Supports(kind, "delta"))
    out.Add Menu_Item("Corte de eixo...", "break", False, False, Kinds_Supports(kind, "break"))
    out.Add Menu_Item("Remover setas", "clear", False, False, True)
    out.Add Menu_Item("-", "", False, False, True)
    out.Add Menu_Item("Estilos...", "styles", False, False, True)
    out.Add Menu_Item("Alterar dados...", "source", False, False, True)
    out.Add Menu_Item("Selecionar a tabela de origem", "selsrc", False, False, True)
    out.Add Menu_Item("Vincular ao PowerPoint", "ppt", False, False, True)
    out.Add Menu_Item("-", "", False, False, True)
    out.Add Menu_Item("Formatar (painel do Excel)", "format", False, False, True)
    Err.Clear
End Function

' Acao de um item. Nunca propaga erro.
Public Sub Menu_DoParam(param As String)
    Dim key As String
    On Error GoTo Fail
    If Left$(param, 2) = "t:" Then
        key = Mid$(param, 3)
        Core_SetToggle key, Not Core_GetToggle(key)
    Else
        Select Case param
            Case "a:cagr": Core_AddAnnotation "cagr"
            Case "a:diff": Core_AddAnnotation "diff"
            Case "delta": Core_DeltaPick
            Case "break": Core_SetAxisBreak
            Case "clear": Core_RemoveAnnotations
            Case "styles": Core_StyleManage
            Case "source": Core_ChangeSource
            Case "selsrc": Menu_SelectSource
            Case "ppt": Core_CopyToPowerPoint "chart"
            Case "format"
                If BUILD_CONFIG <> "Test" Then Application.CommandBars.ExecuteMso "ChartFormatSelection"
        End Select
    End If
    Watchers_NotifySelection
    Exit Sub
Fail:
    Errors_Handle "Menu_DoParam(" & param & ")"
End Sub

' Alvo do OnAction dos itens: o Parameter do controle clicado. Nunca propaga erro.
Public Sub Menu_Do()
    Dim p As String
    On Error Resume Next
    p = CStr(Application.CommandBars.ActionControl.Parameter)
    Err.Clear
    If Len(p) > 0 Then Menu_DoParam p
End Sub

' Seleciona a tabela de origem do grafico ativo (endereco da meta do bloco, coluna C). Nunca propaga.
Private Sub Menu_SelectSource()
    Dim c As Chart, addr As String, rng As Range
    On Error Resume Next
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    addr = CStr(Registry_Block(c).Cells(1, 3).Value)
    If Left$(addr, 1) = "=" Then addr = Mid$(addr, 2)
    Set rng = Application.Range(addr)
    If rng Is Nothing Then Exit Sub
    rng.Worksheet.Activate
    rng.Select
    Err.Clear
End Sub

' Monta a barra popup temporaria (recriada a cada clique). Devolve-a (Nothing se falhou). Nunca propaga.
Private Function Menu_Build(c As Chart) As Object
    Dim cb As Object, it As Variant, ctl As Object, beginGroup As Boolean, s As String
    On Error Resume Next
    Application.CommandBars(MENU_NAME).Delete
    Err.Clear
    Set cb = Application.CommandBars.Add(MENU_NAME, 5, False, True)        ' msoBarPopup, temporaria
    For Each it In Menu_Items(c)
        If it(0) = "-" Then
            beginGroup = True
        Else
            Set ctl = cb.Controls.Add(1)                                  ' msoControlButton
            ctl.Caption = CStr(it(0))
            ctl.Parameter = CStr(it(1))
            ctl.OnAction = "'" & ThisWorkbook.Name & "'!Menu_Do"
            ctl.Enabled = CBool(it(4))
            If CBool(it(2)) Then ctl.State = IIf(CBool(it(3)), -1, 0)     ' msoButtonDown / msoButtonUp
            If beginGroup Then ctl.BeginGroup = True: beginGroup = False
            s = s & CStr(it(0)) & "|"
        End If
    Next
    mTestLast = s
    Set Menu_Build = cb
    Err.Clear
End Function

' Botao direito num grafico do ThinQQ: True = o menu do ThinQQ foi mostrado (quem chama cancela o do Excel). Nunca propaga.
Public Function Menu_OnChartRightClick(c As Chart) As Boolean
    Dim cb As Object
    On Error Resume Next
    If Not Menu_Enabled() Then Exit Function
    If Not Registry_IsThinQQChart(c) Then Exit Function
    Set cb = Menu_Build(c)
    If cb Is Nothing Then Exit Function
    If BUILD_CONFIG <> "Test" Then cb.ShowPopup
    Menu_OnChartRightClick = (Err.Number = 0)
    Err.Clear
End Function

' Duplo clique num grafico do ThinQQ: seleciona a tabela de origem (como o datasheet do think-cell). True = tratado. Nunca propaga.
Public Function Menu_OnChartDoubleClick(c As Chart) As Boolean
    On Error Resume Next
    If Not Menu_Enabled() Then Exit Function
    If Not Registry_IsThinQQChart(c) Then Exit Function
    c.Parent.Activate
    Menu_SelectSource
    Menu_OnChartDoubleClick = (Err.Number = 0)
    Err.Clear
End Function
```
  `clsChartWatcher.cls`:

```vb
' Botao direito (1.13.0; sonda P-U5): menu do ThinQQ no lugar do menu do Excel (opcao ThinQQ\Excel\MenuDireito).
Private Sub Cht_BeforeRightClick(Cancel As Boolean)
    On Error Resume Next
    If Menu_OnChartRightClick(Cht) Then Cancel = True
End Sub

' Duplo clique (1.13.0): seleciona a tabela de origem.
Private Sub Cht_BeforeDoubleClick(ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long, Cancel As Boolean)
    On Error Resume Next
    If Menu_OnChartDoubleClick(Cht) Then Cancel = True
End Sub
```
  Se a E0 mostrou que `ExecuteMso "ChartFormatSelection"` não existe, troque o item por "Formatar: use o menu do Excel com a opção desligada" e documente. Se a
  E0 mostrou que o evento não chega com o gráfico não selecionado, documente "clique no gráfico antes".
  **PowerPoint**: só se a E0 confirmou um `idMso` de menu de contexto do gráfico, acrescente ao `customUI14.xml` do `.ppam`
  `<contextMenus><contextMenu idMso="<confirmado>"><menu id="xmnuTqq" label="ThinQQ" imageMso="Refresh"> ... os botões da aba contextual com ids "x" ... </menu></contextMenu></contextMenus>`
  (callbacks já existentes) e o `idMso` à lista permitida do Test-Ribbon.
- [ ] **Passo 3:** suíte → 0; commit `feat: menu do botao direito e duplo clique no grafico do ThinQQ (Excel)`.

### Tarefa X4: Alterar tipo do gráfico (mesmo id) — modelo: standard

**Files:** Modify `src/core/modules/modCoreApi.bas` (`Core_RebuildFromRange(..., Optional newKind)`, `Core_ChangeKind`), `src/loader/ribbon/customUI14.xml` (menu `mnuKind`),
`src/loader/modules/modRibbon.bas` (`Rx_OnChangeKind`), `src/core/modules/modChartMenu.bas` (submenu "Tipo"); Test `tests/vba/core/TestCoreApi.bas`, `TestRibbon.bas`.

- [ ] **Passo 1 (RED):**

```vb
Public Sub Test_Core_ChangeKind_KeepsIdOptionsAndRefuses()
    Dim wb As Workbook, ws As Worksheet, c As Chart, id0 As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5): ws.Range("A3:C3").Value = Array("T2", 20, 6)
    ws.Range("A1:C3").Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    id0 = Registry_IdFromChart(c)
    Core_SetOption "numfmt", "dec1"
    Core_ChangeKind "clustered"
    Set c = Core_ActiveChart()
    AssertEqual "clustered", Registry_ChartType(c)
    AssertEqual id0, Registry_IdFromChart(c), "mesmo id"
    AssertEqual "dec1", Registry_GetOption(c, "numfmt", ""), "opcoes mantidas"
    Core_ClearLastNotice
    Core_ChangeKind "waterfallpair"
    AssertTrue Len(Core_LastNotice()) > 0, "tabela incompativel: aviso e o grafico antigo fica"
    AssertEqual "clustered", Registry_ChartType(Core_ActiveChart())
    Core_ChangeKind "naoexiste"
    AssertEqual "clustered", Registry_ChartType(Core_ActiveChart())
    CloseTemp wb
End Sub
```
- [ ] **Passo 2:** em `Core_RebuildFromRange` acrescente `Optional newKind As String = ""` e troque `kind = Registry_ChartType(c)` por
  `If Len(newKind) > 0 Then kind = newKind Else kind = Registry_ChartType(c)`; e

```vb
' Alterar tipo (1.13.0): recria o grafico ativo com o mesmo intervalo de origem (meta do bloco), o mesmo id e as mesmas opcoes que o tipo novo
' aceita (ReapplyOptions pela matriz). Tabela que nao serve ao tipo novo: o aviso da validacao e o grafico antigo fica. O grafico ja colado num
' slide NAO muda de tipo (o slide e uma copia): aviso para colar de novo. Nunca propaga erro.
Public Sub Core_ChangeKind(newKind As String)
    Dim c As Chart, addr As String, rng As Range
    On Error GoTo Fail
    If Not Kinds_IsValid(newKind) Then Core_Notify "Tipo de gráfico desconhecido: " & newKind, vbExclamation: Exit Sub
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    If Registry_ChartType(c) = newKind Then Exit Sub
    addr = CStr(Registry_Block(c).Cells(1, 3).Value)
    If Left$(addr, 1) = "=" Then addr = Mid$(addr, 2)
    Set rng = Application.Range(addr)
    If Core_RebuildFromRange(c, rng, newKind) Then
        If Not LinkFrame_Find(Core_ActiveChart()) Is Nothing Then Core_Notify "O tipo mudou no Excel. O gráfico que já está no slide não muda de tipo: cole-o de novo (Vincular ao PowerPoint).", vbInformation
    End If
    Exit Sub
Fail:
    Errors_Handle "Core_ChangeKind"
End Sub
```
  Faixa do Excel: `menu id="mnuKind" label="Tipo" image="tqq_more" getEnabled="Rx_GetEnabled"` no `grpChart` (antes de `ddFormat`) com os 21 tipos (ids `k` + id do
  `grpInsert`, `tag` = tipo, `onAction="Rx_OnChangeKind"`, sem `size`); `Rx_OnChangeKind` = `If CoreOrWarn() Then RunCoreFeature "Core_ChangeKind", control.Tag` +
  invalidar. No `modChartMenu`, item "Tipo..." que chama `Core_ChangeKind` via um `msoControlPopup` com os 21 tipos (`Parameter = "kind:<tipo>"`, tratado em
  `Menu_DoParam`). PowerPoint: fora (o slide é uma cópia; documentado).
- [ ] **Passo 3:** suíte → 0; commit `feat: alterar o tipo do grafico mantendo o id e as opcoes`.

### Tarefa X5: Documentação da 1.13.0 — modelo: standard

- [ ] Tutorial: `#ppt-inserir` reescrita como "Vincular ao PowerPoint" (galeria; guia Inserir), moldura verde/vermelha, "Alterar dados" que mantém o slide,
  botão direito/duplo clique, Alterar tipo; matriz `#ppt-paridade`; referência da faixa (grupo "PowerPoint" novo, "Tipo", guia Inserir; sem "Inserir também
  no slide"). README (recursos, verificações manuais **126–131**, limitações: o tipo no slide não muda; menu do botão direito substitui o do Excel — opção
  `MenuDireito`), CHANGELOG, notas, Test-Docs (rótulos novos; `LINKFRAME_NOTE_REBUILT` só no caminho de falha). Commit `docs: 1.13.0`.

### Tarefa X6: Suíte final, revisão e verificação manual — modelo: standard (controlador + humano)

- [ ] Suíte `-Ppt` (PowerPoint do usuário fechado) → 0, 0 pulados; revisão final do ramo; onda; verificação manual 126–131; push `feat/ux-entrada`.

---

## Resumo da estimativa (detalhe, suposições e confiança: spec §13)

| Versão | Tarefas | Agente (min, otim./prov./pess.) | Relógio | Humano |
|---|---|---|---|---|
| 1.11.0 | U0–U6 | 410 / 680 / 1125 | 1,5–2,5 dias | 45–75 min |
| 1.12.0 | S1–S7 | 470 / 800 / 1330 | 2–3 dias | 40–60 min |
| 1.13.0 | E0, X1–X6 | 400 / 700 / 1160 | 2–3 dias | 70–105 min |
| **Total** | 21 tarefas | **1280 / 2180 / 3615** (≈ 21 / 36 / 60 h) | **6–9 dias** | **≈ 2,5–4 h** |

As maiores incertezas: P-U3/P-U4 (faixa descartada em silêncio), P-U6 (alterar sem perder o slide), PowerPoint do usuário aberto (suítes com `-Ppt`
reagendadas), memória da máquina (suítes de 15 a 77 min) e o EDR.
