# ThinQQ — Pesquisa e Design: paridade com o think-cell na integração Excel ↔ PowerPoint (1.6.0 a 1.10.0)

Data: 2026-10-15
Autor: Israel Dias (com apoio do Claude Code)
Status: rascunho para decisão do usuário (nenhum desenvolvimento iniciado). Adendo às specs da Fase 3 (§4.4 "Copiar para o PowerPoint") e
de 2026-10-14 (`docs/superpowers/specs/2026-10-14-thinqq-ppt-delta-design.md`: motor de sincronização `modPptSync`, suplemento `ThinQQ.ppam`,
atualizador). Plano: `.superpowers/sdd/2026-10-15-thinkcell-parity/plan.md` (a Tarefa C0 copia os dois para `docs/superpowers/`).
**Emenda 1 (2026-10-15, pedido do usuário):** "ajuste o plano para permitir a criação de gráficos diretamente pelo PowerPoint assim como o
think-cell" ⇒ **Parte G** (§14): menu "Inserir gráfico" na aba ThinQQ do PowerPoint com os 21 tipos, pasta de dados do Excel como datasheet,
"Editar dados" e (numa segunda etapa) comandos do gráfico pelo PowerPoint. Versões renumeradas: 1.8.0 = G-I (manchete), 1.9.0 = G-II + E,
1.10.0 = F (§8). As seções 1, 3.1, 6, 8, 9, 10, 11 e 12 foram atualizadas; o detalhe da Parte G está na §14.

## 1. Objetivo e escopo

Pedido do usuário: "pode estimar e planejar os ajustes para deixarmos as funcionalidades fiéis ao think-cell" na integração Excel ↔
PowerPoint, sem iniciar o desenvolvimento. O ThinQQ 1.5.0 já tem: gráfico nativo ligado ao bloco `_ThinQQ`; "Copiar para o PowerPoint"
(imagem ou gráfico vinculado); motor de sincronização (`LinkFormat.Update` + recálculo das setas no tamanho do slide + clonagem das
`tqq_ann_*`); suplemento `ThinQQ.ppam` (ao abrir: perguntar/sempre/nunca); botão "Atualizar gráficos no PowerPoint" no Excel; atualização do
suplemento pelo atualizador do Excel; delta por cliques. Lacunas listadas pelo usuário:

| # | Lacuna | Decisão | Versão |
|---|---|---|---|
| 1 | Atualização automática com Excel e PowerPoint abertos ("Update Automatically") | **Entra** — motor no Excel (eventos + `OnTime` com espera de 2 s) sobre um índice das formas automáticas gravado na apresentação; liga/desliga por elemento no `.ppam` e padrão para vínculos novos no Excel | **1.6.0** |
| 2 | Criar a partir do intervalo no Excel e inserir direto no slide, com o intervalo destacado | **Entra** — "Inserir também no slide" (caixa de seleção no menu PowerPoint: qualquer tipo inserido vai ao slide ativo, vinculado) + moldura colorida em volta do intervalo de origem | **1.7.0** |
| G | **Criar o gráfico pelo PowerPoint** ("Insert > Elements" do think-cell) — emenda 1 | **Entra** — aba ThinQQ do PowerPoint > "Inserir gráfico" (os 21 tipos, nos 5 menus do Excel): o core no Excel cria o gráfico numa **pasta de dados** ao lado da apresentação (`<apresentação> - dados ThinQQ.xlsx`, uma planilha por gráfico, dados de exemplo) e o cola **vinculado e automático** no slide; "Editar dados" traz a planilha à frente num Excel visível (o Excel é o datasheet); comandos do gráfico pelo PowerPoint numa 2.ª etapa (§14) | **1.8.0** (G-I) / **1.9.0** (G-II) |
| 4 | Texto vinculado a célula (caixa de texto no slide; título do gráfico) | **Entra** — campo único por forma (texto antes + valor da célula + texto depois), atualizado pelo mesmo motor (manual, ao abrir e automático); título do gráfico ligado a uma célula (nativo, `ChartTitle.Formula`, medido) | **1.9.0** (com G-II) |
| 3 | Tabelas do PowerPoint vinculadas a intervalos | **Entra** — tabela nativa do PowerPoint criada a partir do intervalo (formatação inicial do Excel), atualizada só nos textos que mudaram, linhas/colunas acrescentadas/removidas | **1.10.0** |
| 5 | Conexão com Tableau | **Fora do escopo** (§5.5) | — |
| 6 | Datasheet interno (editar dados no próprio PowerPoint) | **Fora do escopo** como datasheet DENTRO do .pptx (§5.6; `BreakLink` por automação derruba o PowerPoint). A experiência equivalente (criar no slide, digitar os dados, ver o slide mudar) vem da Parte G com o Excel como datasheet numa pasta ao lado da apresentação (§14.2) | — |

## 2. Pesquisa (2026-10-15)

### 2.1 think-cell (fontes oficiais)
- **Atualização manual por padrão; automática por elemento.** "Excel-linked elements are set to update manually by default"; com dados
  novos o elemento mostra moldura verde e o botão Update; "Update Once" aplica uma vez; "Update Automatically" (minibarra) faz o elemento
  atualizar "whenever their linked Excel range changes", desde que o think-cell "can connect to the linked Excel range"; "Disable Automatic
  Updates" volta ao manual. Elementos ligados ao Tableau não atualizam sozinhos (Update Automatically = uma atualização).
  Fonte: https://www.thinkcell.com/en/resources/manual/manage-data-links
- **Sessão do Office.** As fontes são recarregadas no início de cada sessão ("Not yet loaded", interrogação laranja); fonte que não abre:
  exclamação vermelha "File cannot be opened"; o elemento guarda a última cópia dos dados no datasheet. Mesma fonte.
- **Diálogo Data Links / Manage Data Sources.** Miniaturas dos slides com marcas de atualização pendente, fontes por pasta/planilha,
  filtros (Updates Available, Updated, Automatically Updating, Tableau), Update Once/Revert/Update Automatically; religar a fontes
  alternativas (cópia renomeada/movida). Mesma fonte.
- **Criar a partir do Excel.** "select the cell range ... Insert > think-cell > Link To PowerPoint" e o tipo; o elemento é posicionado no
  slide; "Linked Excel ranges have color-coded data link frames": **verde** = uma apresentação aberta usa o intervalo, **vermelho** = nenhuma,
  **azul** = moldura selecionada; acrescentar série/linha = redimensionar a moldura; "Transpose Link", "Edit Data Layout".
  Fonte: https://www.thinkcell.com/en/resources/manual/exceldatalinks
- **Tabelas.** Três modos: "Create tables", "Create tables with Excel formatting" (formatação inicial igual à do Excel), "Create images of
  tables". Atualizar a tabela "works exactly the same as updating the chart" (inclusive automática). Fontes: mesma página;
  https://www.thinkcell.com/en/blog/think-cell-integrations (think-cell 11, 08/12/2020).
- **Texto vinculado.** Selecionar a célula no Excel; no PowerPoint selecionar o objeto inteiro (substitui o texto), parte do texto
  (substitui a parte) ou pôr o cursor (insere); "Insert > think-cell > Elements > Text Linked to Excel"; "You can link text in any object in
  your presentation to Excel, except for PowerPoint tables and SmartArt"; no think-cell 11: campo em qualquer rótulo do think-cell ou caixa de
  texto, "always updates automatically when the cell contents in Excel changes". Fontes: as duas anteriores.
- **Tableau.** Exige servidor Tableau (local ou nuvem), extensão do think-cell no navegador (Chrome, Edge, Firefox); "can only import Tableau
  data from views displayed in a browser, not directly from Tableau Desktop"; atualização pelo botão Update na aba do navegador ou pelo diálogo
  Data Links; sem atualização automática. Fonte: https://www.thinkcell.com/en/resources/manual/tableaudata
- **Datasheet do elemento.** Gráficos, caixas de texto, tabelas, bolas de Harvey e caixas de seleção têm "element datasheets", abertos por
  duplo clique; "a customized Microsoft Excel sheet for data input"; slide workbooks por slide. Fontes:
  https://www.thinkcell.com/en/resources/manual/element-datasheets , https://www.think-cell.com/en/resources/manual/slide-workbooks
- **Não confirmado** (as páginas não dizem): se a atualização automática exige as duas aplicações abertas o tempo todo e com que latência;
  o que acontece com a tabela vinculada quando o intervalo muda de tamanho (além de "redimensionar a moldura"); se o datasheet fica dentro do
  .pptx; se o texto vinculado funciona em rótulos de gráfico nativos do PowerPoint (só "rótulos do think-cell").

### 2.2 Microsoft (VBA)
- Excel: `Application.SheetChange`, `SheetCalculate`, `WorkbookAfterSave` (eventos do aplicativo, já usados pelo core via `clsAppEvents`) e
  `Application.OnTime` (agendar/cancelar pelo mesmo horário; já usado pelo delta): https://learn.microsoft.com/en-us/office/vba/api/excel.application.ontime
- PowerPoint **não tem `OnTime`** nem evento de "dados do vínculo mudaram"; tem `WindowActivate`, `WindowSelectionChange`,
  `AfterPresentationOpen`, `PresentationBeforeSave`: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application (lista de
  eventos). Timer no PowerPoint só com `Declare` (SetTimer): proibido no projeto.
- `ChartTitle.Formula` / `DataLabel.Formula` (texto ligado a célula num gráfico do Excel): https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.formula
- `Shapes.AddTable`, `Table.Cell(r, c).Shape.TextFrame`, `Rows.Add`/`Columns.Add`: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addtable
- `Slides.FindBySlideID` (o `SlideID` não muda quando o slide é movido): https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slides.findbyslideid
- `ChartData.BreakLink`, `ChartData.Activate`, `ChartData.Workbook`: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.chartdata.breaklink

## 3. Sondas (2026-10-15; Excel 2019 32 bits e PowerPoint 16.0 32 bits; instâncias NOVAS e invisíveis criadas por scripts PowerShell externos, PIDs registrados e encerrados; nenhum POWERPNT do usuário aberto (conferido antes); arquivos em `C:\temp\tqq-parity-probe`, scripts `p1.ps1`, `p2.ps1`, `p3.ps1`)

| # | Medida | Resultado |
|---|---|---|
| S1 | `ChartTitle.Formula = "=Dados!$E$1"` e `Points(1).DataLabel.Formula = "=Dados!$E$2"` no Excel | Aceitos; textos "Receita 2026" e "rotulo A". |
| S2 | Colar vinculado, mudar E1/E2/B2 **sem salvar**, `LinkFormat.Update` (pasta aberta) | 58,8 ms; **título e rótulo ligados a célula acompanham** ("Receita NOVA", "rotulo B"); a caixa `tqq_ann_txt_1` dentro do gráfico **não** (como já medido em 2026-10-14: o motor do ThinQQ copia o texto). |
| S3 | Varrer 99 formas (15 gráficos marcados) lendo `Tags.Item` de fora do processo | 554 ms (com try/catch do PowerShell). Sem exceção: 391 formas = 2 520 ms (**~6,4 ms por forma**); só `Tags.Count`: 1 830 ms. `Tags.Item` de chave inexistente devolve "" sem erro. |
| S4 | Índice na etiqueta da **apresentação** + `Slides.FindBySlideID` + `Shapes(nome)` | 20 buscas = 96,5 ms (**~5 ms por forma**, independe do tamanho da apresentação). |
| S5 | `LinkFormat.Update` em 15 gráficos (pasta aberta) | 465 ms (~31 ms cada). |
| S6 | Tabela nativa 11 × 5: `AddTable` + 55 `Cell().Shape.TextFrame.TextRange.Text` a partir de `Range.Text` | 1 451 ms; ler o intervalo em bloco (`Value2`) 7,4 ms; reescrever os 55 textos 569 ms (~10 ms por célula); reescrever 1 célula 14 ms; **negrito da célula mantido** ao trocar o texto; `Rows.Add` copia preenchimento e itálico da última linha (altura da forma cresce); `Columns.Add` mantém a largura da forma (colunas ficam mais estreitas). |
| S7 | Moldura no Excel: retângulo sem preenchimento, `Placement = xlMoveAndSize`, sobre A1:C4; nome definido `TQQ_T_1` | Linha inserida acima: a moldura desce (top 0 → 15) e o nome vira `$A$2:$C$5`; linha inserida dentro: altura 60 → 75 e o nome `$A$2:$C$6`; apagar volta tudo. |
| S8 | Campo dentro do texto marcado por **hiperlink** (subendereço `#tqq1`) | Criado (sem sublinhado/cor visíveis pela leitura); inserir texto antes mantém o campo; mas depois de `Hyperlink.TextToDisplay = "4.567"` o objeto hiperlink passa a apontar para **outro trecho** ("Nova "), e a troca seguinte escreveu no começo do texto ⇒ **hiperlink não serve de marcador** (rejeitado). Hiperlink sem endereço é recusado ("This kind of object cannot have a hyperlink"). |
| S9 | Etiqueta numa caixa de texto do slide (`TQQ_TEXT`) | Gravada e lida. |
| S10 | `ChartData.BreakLink` num gráfico colado vinculado (origem em planilha visível e em `_ThinQQ` muito oculta), apresentação **sem janela** | Nos dois casos: depois do `BreakLink` `IsLinked` continua `True` e a fórmula da série fica vazia; o `SaveAs` seguinte **derruba o PowerPoint** (RPC 0x800706BA; sem evento do SentinelOne nem do Windows). Repetido 3 vezes. ⇒ nunca chamar `BreakLink` por automação; caminho "datasheet" fora do escopo. |

Ambiente: memória livre ~2,5 GB no início; nenhum processo deixado; nenhum evento do SentinelOne nas sondas.

### 3.1 Sondas pendentes (a agendar depois da gravação de tela; NÃO rodadas: a máquina passou a gravar Excel/PowerPoint reais)
Todas em instâncias próprias e invisíveis por script PowerShell externo (padrão de `p1.ps1`: conferir que não há `POWERPNT`, registrar PIDs,
encerrar no `finally`, matar só `-Embedding` criados durante a sonda), pasta `C:\temp\tqq-parity-probe`.

| # | Pergunta | Passos (script `p4.ps1` .. `p8.ps1`) | Esperado / decisão que depende |
|---|---|---|---|
| P1 | Custo da impressão digital (`PptSync_Fingerprint`, SHA-256 em VBA) num bloco grande | Excel próprio com o core de teste (`dist\ThinQQ.core.test.xlam` montado numa cópia `C:\temp\tqq-pd`), pasta com gráfico de 12 categorias × 8 séries e outro de 60 × 3; `xl.Run "PptSync_Fingerprint"` 20× com um dublê `TestPptFake` (ou módulo injetado que cronometra) | < 30 ms por gráfico ⇒ C2 como está; > 100 ms ⇒ C2 troca a impressão pelo hash só dos valores (sem SHA) antes da comparação (risco R3) |
| P2 | `WindowActivate` do PowerPoint dispara quando o usuário volta de OUTRO aplicativo? | Não automatizável (precisa de foco real): fica como **verificação manual** do item C6-7 | Só importa para a alternativa B (§5.1); a escolhida não depende |
| P3 | Excel chamando um PowerPoint ocupado (texto em edição) mostra a caixa "aguardando ação OLE"? | PowerPoint próprio com uma caixa `MsgBox` modal aberta por macro injetada no `.pptm` de teste via `Application.Run` assíncrono não é possível sem processo filho ⇒ **verificação manual** (C6-5): editar uma caixa de texto no PowerPoint enquanto muda um número no Excel | Se a caixa aparecer: C2 reduz o tempo de espera (`Application.OLE...` não existe no Excel) e passa a pular o ciclo quando `PowerPoint.ActiveWindow.Selection.Type = ppSelectionText` (texto em edição) — já previsto como guarda no código da C2 |
| P4 | Moldura sem preenchimento bloqueia cliques nas células? | Manual (D4-2) | Se bloquear: trocar o retângulo por 4 linhas (`AddLine`) — `LinkFrame_Draw` já isola o desenho |
| P5 | Seleção de texto do PowerPoint legível pelo Excel enquanto o Excel está ativo | PowerPoint próprio, apresentação com janela (`Presentations.Add(-1)`, aplicativo invisível não permite janela: a sonda abre com janela **minimizada**), selecionar caracteres por `TextRange.Characters(5,3).Select`, depois ler `ActiveWindow.Selection.Type`/`TextRange.Start/Length` de fora | Type = 3 e Start/Length corretos ⇒ o botão do Excel "Vincular célula ao texto selecionado" funciona; senão só o botão do `.ppam` (E3) |
| P6 | `ChartTitle.Formula` sobrevive a `Core_SafeRefresh` e a "Alterar dados" de um gráfico ThinQQ | Excel próprio + core de teste: inserir colunas empilhadas, ligar o título, `Core_SafeRefresh`, `Core_RebuildFromRange` | Título mantido no refresh; no rebuild o `Style_ApplyBase` o apaga ⇒ E2 reaplica pela opção `titleref` (já no plano) |
| P7 | `Range.Text` com coluna estreita ("####") e datas | Excel próprio: coluna de largura 2 com 12345,6; data | Confirma o recurso de `PptText_CellText` (formatar pelo `NumberFormat` quando o texto é só "#") |
| PG1 | (G) Excel criado pelo VBA do PowerPoint, tornado visível e entregue ao usuário: sobrevive ao fim da referência do `.ppam`? Abrir o loader instalado por `Workbooks.Open` mostra a aba ThinQQ e carrega o core? | Sonda `pg1.ps1` NÃO reproduz (precisa ser VBA do PowerPoint): rodar no `.pptm` de teste uma macro de teste que cria o Excel, `Visible = True`, `UserControl = True`, abre um `.xlam` de teste (o loader de teste da cópia `C:\temp\tqq-pd\dist`) e solta a referência; de fora, conferir que o `EXCEL.EXE` continua e que `Workbooks` tem o loader e o core; encerrar pelo PID | Fica ⇒ G4 como está; some ⇒ G4 guarda a referência em variável de módulo do `.ppam` até o PowerPoint fechar |
| PG2 | (G) `AppActivate` do VBA do PowerPoint traz o Excel à frente? | Manual (precisa de foco real): item G6-3 | Não traz ⇒ aviso "O Excel abriu a planilha de dados (veja a barra de tarefas)." |
| PG3 | (G) Colagem vinculada feita pelo core DURANTE um `xl.Run` vindo do PowerPoint (área de transferência + reentrada COM) | `pg3.ps1`: PowerPoint próprio + `.pptm` de teste; macro de teste faz `xl.Run "'core'!Core_PptInsertChart", slide, "stacked"` com o Excel próprio do teste (core de teste); conferir forma colada, `IsLinked`, etiquetas | Funciona ⇒ G2 como está; falha ⇒ o core devolve o caminho/planilha e o `.ppam` faz a colagem (`ChartObject.Copy` pelo `xl` + `Shapes.Paste` no próprio PowerPoint) |
| PG4 | (G) `LinkFormat.SourceFullName` é gravável num gráfico vinculado? O PowerPoint resolve caminho relativo ao mover a pasta inteira? | `pg4.ps1`: colar vinculado, salvar .pptx e .xlsx, copiar os dois para outra pasta, abrir a cópia, ler `SourceFullName`; gravar o caminho novo; `LinkFormat.Update` | Gravável ⇒ reapontar no lugar; não ⇒ recolar (perde animação) |
| PG5 | (G) O evento `PresentationSave` do PowerPoint chega antes da gravação e permite o `xl.Run` que salva a pasta de dados? | `.pptm` de teste com a classe de eventos; salvar por automação; registrar a ordem | Chega ⇒ salvar junto; senão só o automático salva |
| PG6 | (G) `Presentation.FullName`/`Path` de apresentação no OneDrive sincronizado (https ou caminho local?) | Manual (conta do usuário) | https ⇒ recusa com a mensagem (já no código) |
| PG7 | (G) Seleção de um ponto do gráfico no slide é visível ao `.ppam` (`WindowSelectionChange` + `Selection`)? | `.pptm` de teste com janela: selecionar `Chart.SeriesCollection(1).Points(2)` por código e ler `Selection.Type`/`TypeName` | Só para um delta futuro no PowerPoint; hoje fora do escopo |
| PG8 | (G) Tempo de ida e volta "Inserir gráfico" com o Excel já aberto e com Excel frio | Manual com cronômetro (G6-1) | Meta < 3 s com Excel aberto; frio (abrir Excel + loader + core) < 12 s com barra de status |
| PG10 | (G) Copiar a pasta de dados ABERTA no Excel ("Empacotar para enviar") | Manual (G11) | Falha ⇒ o `.ppam` pede para fechar a pasta no Excel antes de empacotar |
| PG9 | (G) A faixa do `.ppam` com `image=` próprios (ícones PNG do build) carrega no PowerPoint real | Manual na instalação (G6-2); o build já monta ícones PNG para o `.ppam` (`build.ps1` alvo ppt usa `-Icons Png`) | Não carrega ⇒ `-Icons Mso` (conversão já existe no build) |
| P8 | Tabela criada DENTRO do PowerPoint (VBA do `.ppam`) vs de fora | Rodar o preenchimento 11 × 5 no `.pptm` de teste (módulo de teste) | Se ~10× mais rápido: F2 passa a montar a tabela pelo `.ppam` (`PptTqq_FillTable`) quando ele existir; senão fica no core (como no plano) |

## 4. Critérios e restrições (herdados)
(1) gráfico/tabela/texto nativos e corretos **sem** o ThinQQ; (2) sem `Declare`, sem referências extras, ligação tardia; (3) Office 2016+
32/64 bits; (4) testável em instâncias próprias invisíveis; (5) nunca tocar Excel/PowerPoint/arquivos do usuário sem ação dele; (6) **nunca
iniciar processo filho a partir do VBA** (o EDR SentinelOne mata o Excel e põe arquivos em quarentena; `Test-VbaLint` já barra); (7) pastas de
origem abertas pelo motor sempre com macros desligadas (`AutomationSecurity = 3`, já em `PptSync_OpenBook`) e somente leitura;
(8) PowerPoint de instância única: o core só usa o PowerPoint que já está aberto (`Ppt_GetApp(False)`) para sincronizar; só a colagem/inserção
(ação explícita do usuário) abre um.

## 5. Decisões por lacuna

### 5.1 Lacuna 1 — Atualização automática (1.6.0)
**(a) think-cell:** por elemento; nasce manual; "Update Automatically" liga; atualiza quando o intervalo muda e o think-cell alcança a fonte;
recarrega a cada sessão. **(b) Fiel no ThinQQ =** cada gráfico vinculado (e, nas versões seguintes, texto e tabela) tem "Atualizar
automaticamente" (liga/desliga) e "Atualizar uma vez"; com o Excel e o PowerPoint abertos, mudar um valor da pasta (digitado, colado ou
recalculado) atualiza **no slide** dados **e setas** dos elementos automáticos daquela pasta em ~2–3 s, sem caixa de mensagem; o padrão de
vínculo novo é manual (como o think-cell) e há uma opção para nascer automático; uma chave global pausa tudo. **Fica de fora:** indicador
de estado por elemento (interrogação laranja/exclamação vermelha), diálogo Data Links, "Revert", automático sem o Excel aberto (sem Excel
não há o que mudar), religar fontes.

**(c) Arquiteturas avaliadas**
| Candidata | Como | Medido/avaliado | Veredito |
|---|---|---|---|
| **A. Motor no Excel** | Eventos `SheetChange`/`SheetCalculate`/`WorkbookAfterSave` de pastas com `_ThinQQ` marcam a pasta "suja" e agendam `OnTime` +2 s (cada evento reagenda: espera o usuário parar). No disparo: `Ppt_GetApp(False)` (só um PowerPoint já aberto); para cada apresentação editável, lê a etiqueta **da apresentação** `TQQ_AUTOIDX` (índice slideID/forma/pasta) e acha só as formas automáticas daquela pasta por `FindBySlideID` (~5 ms cada, S4); sincroniza com o motor existente sem forçar (impressão digital pula o que não mudou) | Peças medidas: Update com pasta aberta lê não salvo (58 ms, S2); 15 gráficos 465 ms (S5); busca indexada 5 ms (S4) contra varredura 6,4 ms **por forma** (S3: 400 formas = 2,5 s por ciclo, inviável) | **Escolhida** |
| B. Suplemento do PowerPoint | `WindowActivate`/`WindowSelectionChange` ao voltar ao PowerPoint pedem ao core uma sincronização não forçada | Sem timer no PowerPoint (sem `Declare`); não é "ao vivo" com as duas janelas lado a lado; `WindowActivate` vindo de outro aplicativo não medido (P2) | Alternativa (complemento possível numa 1.6.x) |
| C. Vínculo nativo `AutoUpdate = 2` | O PowerPoint atualiza vínculos automáticos | Medido (2026-10-14): abrir por automação não atualiza; `UpdateLinks` só reabre; setas não recalculam | Rejeitada |
| D. Varredura completa a cada mudança | Sem índice | 6,4 ms × formas (S3) | Rejeitada (índice obrigatório) |

**Modelo**
- Etiqueta da forma `TQQ_AUTO = "1"`; etiqueta da apresentação `TQQ_AUTOIDX` = entradas separadas por quebra de linha (`vbLf`), campos por
  `|`: `SlideID`, nome da forma, pasta de origem (`PptSync_SourcePath`). **Medido na revisão final da C** (PowerPoint real, salvar `.pptx`,
  fechar, reabrir): o formato original com `Chr(30)`/`Chr(31)` perdia os separadores (caracteres proibidos no XML: o índice voltava de 43 para
  38 caracteres); numa etiqueta, Tab vira espaço e CR vira LF; `vbLf` e `|` voltam iguais. Os campos são saneados (`|`, LF, CR, Tab,
  `Chr(30)`, `Chr(31)` viram `_`). Vale para qualquer etiqueta nova das versões seguintes (ex.: `TQQ_TBLAST` e a referência de reserva do
  texto vinculado no plano usam `Chr(30)`/`Chr(31)` e precisam trocar de separador). Mantido por: "Atualizar automaticamente" do `.ppam` (C3), a colagem com a
  opção "Novos vínculos atualizam automaticamente" (C1) e **toda** sincronização completa (`PptSync_Presentation` reconstrói o índice das formas
  marcadas que tem — conserta slides copiados de outra apresentação na próxima "Atualizar gráficos"/abertura). Entrada velha (slide ou forma
  apagados, `TQQ_AUTO` ausente) sai do índice quando o motor a encontra. Nome repetido no mesmo slide: o ThinQQ renomeia a forma para
  `<nome>_2`, `_3`… antes de indexar (o nome da forma do slide não é usado pelo Excel).
- Chave global `HKCU\Software\VB and VBA Program Settings\ThinQQ\PowerPoint\AutoUpdate` (padrão "1": o motor age só nas formas marcadas) e
  `AutoNew` (padrão "0": vínculo nasce manual). No build Test só memória.
- `CORE_PPT_API = 2`: novas `Core_PptSyncShapes(pres, entries, force)` ("Atualizar uma vez" das formas selecionadas), `Core_PptGetFlag`/
  `Core_PptSetFlag`. A API 1 continua igual (contrato de 2026-10-14).
- **Ciclo** (`PptAuto_Tick`): reentrada barrada; eventos do próprio ciclo ignorados; PowerPoint ocupado (RPC 0x80010001/0x8001010A) ⇒ nova
  tentativa em 4, 6, 8, 10, 12 s e depois desiste (registrado); apresentação somente leitura pulada; texto em edição no PowerPoint
  (`ActiveWindow.Selection.Type = 3`) ⇒ adia (guarda contra P3); erro de um gráfico conta e segue (`PptSync_Items`); a barra de status mostra
  "ThinQQ: atualizando gráfico i de N do PowerPoint..." só enquanto trabalha. Ocupado é **propagado** por `PptSync_SyncOne` (antes virava
  "falha" e a nova tentativa nunca aconteceria).
- **Custo esperado:** gráfico sem mudança = busca 5 ms + impressão digital (P1); com mudança ≈ Update 30–60 ms + redimensionar/recalcular/
  restaurar 0,3–1 s (medido em 2026-10-14) ⇒ 1–3 gráficos por pasta: o Excel fica 1–3 s ocupado ~2 s depois que o usuário para de digitar.
  Mitigação: só elementos automáticos; só da pasta que mudou; impressão digital; chave global de pausa.
- **Excel/PowerPoint do usuário:** o motor usa o PowerPoint em execução (como o botão "Atualizar gráficos no PowerPoint" da 1.4.0); nunca abre
  um. A pasta de origem é a própria pasta aberta (não abre outras: elementos de outra pasta não estão "sujos").

**(d) UI.** `.ppam`: grupo novo "Gráfico selecionado" com `tglPptAuto` "Atualizar automaticamente" (pressionado quando todas as formas ThinQQ
selecionadas são automáticas; ligar já atualiza uma vez) e `btnPptUpdateSel` "Atualizar uma vez"; invalidado em `WindowSelectionChange`.
Ligar/desligar **não** abre o Excel (o `.ppam` grava etiqueta e índice com as mesmas funções puras do core, testadas com as mesmas cadeias).
Excel, menu PowerPoint: `chkPptAutoOn` "Atualização automática ligada" e `chkPptAutoNew` "Novos vínculos atualizam automaticamente"
(checkBox, getPressed/onAction pelo core).

**Valor alto · esforço médio-alto · risco técnico médio** (eventos entre aplicativos, P3) · depende de nada · base de E e F (roteamento por
tipo de elemento e índice). **Manual:** cadência real (digitar e ver o slide mudar), PowerPoint com texto em edição, apresentação somente
leitura, Excel com outra pasta mudando, pausa global, 20 gráficos automáticos, salvar a apresentação depois (o índice persiste).

### 5.2 Lacuna 2 — Inserir a partir do intervalo, direto no slide (1.7.0)
**(a) think-cell:** seleciona o intervalo → "Link To PowerPoint" → tipo → o elemento vai ao slide; o intervalo ganha moldura colorida
(verde/vermelho/azul) que se redimensiona para mudar o vínculo. **(b) Fiel =** selecionar a tabela, ligar "Inserir também no slide" e
escolher qualquer um dos 21 tipos: o gráfico nasce no Excel (motor do ThinQQ), é colado **vinculado** no slide ativo (ou numa apresentação
nova), já com as setas no tamanho do slide e a opção automática conforme `AutoNew`; o intervalo de origem ganha a **moldura verde** do
ThinQQ. **Diferente:** o gráfico também existe no Excel (é ele que calcula setas e rótulos: o ThinQQ não tem motor de gráfico no PowerPoint);
a moldura tem uma cor (verde = intervalo de um gráfico ThinQQ que foi ao PowerPoint) e não muda o vínculo ao ser arrastada (mudar o intervalo
continua sendo "Alterar dados", que recria o gráfico com id novo: o slide deixa de acompanhá-lo, a moldura sai e o ThinQQ avisa para inserir ou colar de novo; revisão final da Parte D). **Fora:** vermelho/azul, transpor, "Edit Data Layout", posicionar com o mouse no slide.

**(c) Arquitetura:** `Core_InsertToPowerPoint(kind)` = conferência da pasta salva **antes** de criar qualquer coisa + `Core_Insert` + moldura
(`modLinkFrame`: retângulo `tqqlink_<id>` sem preenchimento, linha verde 2 pt, `Placement = xlMoveAndSize` (S7: acompanha linhas
inseridas acima e dentro), não impresso, bloqueado, texto alternativo) + `Core_CopyToPowerPoint "chart"`. Faixa: `chkPptInsert`
"Inserir também no slide" no menu PowerPoint (estado da sessão, no loader); `Rx_OnInsert` escolhe `Core_InsertToPowerPoint` quando marcado.
`btnPptFrames` "Mostrar/ocultar molduras" (todas da pasta ativa). Molduras órfãs (gráfico apagado) somem no `Watchers_Rehook`.
Alternativas: 21 botões duplicados num submenu (rejeitada: faixa dobra, testes de contagem); galeria dinâmica (rejeitada: callbacks e ícones
novos); moldura por borda de célula (rejeitada: altera a formatação do usuário); formatação condicional (rejeitada: some com "Limpar
formatos" e polui o gerenciador de regras).

**Valor médio-alto · esforço baixo-médio · risco baixo** (P4) · depende da A (colagem) e de C1 (`AutoNew`, opcional). **Manual:** clicar em
células dentro da moldura (P4), imprimir, inserir/apagar linhas, copiar a planilha, pasta não salva (mensagem e nada criado), PowerPoint
fechado (abre um com apresentação nova, como hoje).

### 5.3 Lacuna 4 — Texto vinculado a célula (1.9.0, com G-II)
**(a) think-cell:** campo em qualquer objeto (inteiro, parte ou cursor), atualizado como os gráficos; em rótulos do think-cell. **(b) Fiel =**
(1) caixa de texto/forma do slide com **um** campo ligado a uma célula: o texto é `prefixo + valor exibido da célula + sufixo`; vincular a
forma inteira, a parte selecionada ou o ponto do cursor; atualiza manual/ao abrir/automático como os gráficos; formatação dos caracteres
mantida (só os caracteres do campo são trocados); (2) **título do gráfico** ThinQQ ligado a uma célula (nativo do Excel, `ChartTitle.Formula`,
segue pelo vínculo do gráfico: S2). **Diferente/fora:** um campo por forma (think-cell: vários); texto em tabelas e SmartArt (o think-cell
também não); rótulos de dados ligados a célula (os rótulos do ThinQQ vêm do bloco `_ThinQQ`; `DataLabel.Formula` brigaria com o
`Ann_Refresh`; o tutorial ensina a usar o título ou uma caixa de texto vinculada).

**(c) Arquitetura:** etiquetas na forma `TQQ_TEXT="1"`, `TQQ_BOOK`, `TQQ_TNAME` (nome definido oculto `TQQ_X_<id>` na pasta de origem:
acompanha linhas inseridas, S7), `TQQ_TREF` (endereço de reserva), `TQQ_TPRE`, `TQQ_TSUF`, `TQQ_TLAST` (último valor escrito). Atualizar =
abrir/usar a pasta (o mesmo `PptSync_OpenBook`, macros desligadas), ler o texto exibido da célula (`Range.Text`; "####" ⇒ `Format` pelo
`NumberFormat`, P7), e: texto atual = prefixo+último+sufixo ⇒ troca só `Characters(Len(prefixo)+1, Len(último))`; senão, se ainda começa pelo
prefixo e termina pelo sufixo ⇒ troca o meio (o usuário editou só o campo); senão ⇒ não toca e conta "Texto editado no PowerPoint (vincule de
novo): K". Roteamento no motor por tipo (`PptSync_ItemKind`: chart/text/table). Vincular: no `.ppam` (`btnPptLinkText` "Texto vinculado ao
Excel": lê a seleção do PowerPoint e chama `Core_PptLinkText(shp, início, tamanho)` no Excel **já aberto** com o core, que usa a célula ativa)
e no Excel (`btnPptLinkText` no menu PowerPoint: lê a seleção da janela ativa do PowerPoint, P5). Pasta não salva: recusado (o nome definido
precisa estar salvo). Rejeitado: marcar o campo por hiperlink (S8: o objeto hiperlink se desloca e escreve no lugar errado); campos do
PowerPoint (só data/número do slide); marcar por propriedade de fonte (frágil, invisível ao usuário).

**Valor médio · esforço médio · risco médio** (P5, P7; edição do usuário). Depende de C (índice, roteamento, automático). **Manual:**
vincular forma inteira/parte/cursor pelas duas entradas; editar o texto em volta e atualizar; editar o próprio campo; célula com data, %,
moeda; coluna estreita; título ligado e "Alterar dados".

### 5.4 Lacuna 3 — Tabelas vinculadas (1.10.0)
**(a) think-cell:** tabela criada do intervalo (com ou sem a formatação do Excel, ou imagem), atualizada como os gráficos, redimensionando a
moldura para linhas/colunas. **(b) Fiel =** "Inserir tabela vinculada no slide" a partir do intervalo selecionado: tabela **nativa** do
PowerPoint com a formatação inicial do Excel (negrito, itálico, cor da fonte, preenchimento, alinhamento horizontal, tamanho relativo),
textos como exibidos no Excel; atualização (manual/ao abrir/automática) troca **só os textos que mudaram** e mantém a formatação feita no
PowerPoint; linhas/colunas a mais na origem (intervalo pelo nome definido `TQQ_T_<id>`, que cresce com inserções, S7) ⇒ `Rows.Add`/
`Columns.Add` (herdam a formatação da última, S6), a menos ⇒ apagadas do fim. **Fora:** "imagem de tabela" (o "Copiar como imagem" de
intervalo do Excel resolve sem vínculo), células mescladas (recusadas na criação), bordas do Excel, largura proporcional das colunas na
atualização, transpor.

**(c) Arquitetura:** `modPptTable` (core): `PptTable_Insert(rng)` (pasta salva; nome definido; `AddTable` no slide de destino de
`Ppt_TargetSlide`; tamanho 80% do slide como os gráficos; etiquetas `TQQ_TABLE`, `TQQ_BOOK`, `TQQ_TNAME`, `TQQ_TREF`, `TQQ_TBLAST` = matriz
de textos serializada); `PptTable_Sync(shp, wb)` (diferença pura célula a célula contra `TQQ_TBLAST`: escreve só o que mudou, ~10 ms por
célula, S6). Custo: 11 × 5 criada em 1,5 s (S6) — aceitável numa ação explícita; atualização típica (1–5 células) < 100 ms. Alternativa:
objeto OLE de planilha vinculado (`PasteSpecial` link) — rejeitada (não é tabela do PowerPoint, fonte escala, duplo clique abre o Excel).
P8 pode mover o preenchimento para o `.ppam`.

**Valor médio · esforço médio · risco baixo-médio.** Depende de C e E (roteamento, nome definido, leitura de texto exibido). **Manual:**
formatação inicial fiel, formatar no PowerPoint e atualizar, linhas/colunas inseridas, intervalo encolhido, 30 × 10.

### 5.5 Lacuna 5 — Tableau: **fora do escopo** (decisão fundamentada)
- O think-cell não lê o Tableau Desktop: exige **servidor** Tableau e uma **extensão de navegador** própria que extrai os dados da view
  (§2.1). O ThinQQ é VBA: não tem extensão de navegador, não pode iniciar processo (EDR) e não tem onde guardar credenciais com segurança.
- Uma via só-VBA seria a API REST do servidor Tableau (não pesquisada aqui: a pesquisa web liberada cobre só think-cell e Microsoft Learn;
  **não confirmado**), com token pessoal guardado no registro do usuário (risco de segurança), sem servidor de teste e sem uso confirmado de
  Tableau na empresa.
- **Recomendação:** fora do escopo; o tutorial ensina o caminho que já funciona: exportar a view do Tableau (crosstab/CSV) para uma planilha,
  inserir o gráfico ThinQQ a partir dela e ligar "Atualizar automaticamente". Reavaliar se houver demanda e um servidor de testes.

### 5.6 Lacuna 6 — Datasheet interno: **fora do escopo** como recurso
- O datasheet do think-cell é uma planilha do Excel própria do elemento, que **o próprio think-cell** recalcula. No ThinQQ quem desenha setas,
  rótulos e conectores é o core no Excel com o bloco `_ThinQQ`: um gráfico com dados incorporados no .pptx não teria quem recalcule as setas
  sem uma cópia completa do motor rodando dentro do Excel incorporado (sem garantia de que o ThinQQ está carregado nesse Excel).
- Sonda S10: `ChartData.BreakLink` por automação deixa o gráfico inconsistente e **derruba o PowerPoint** ao salvar ⇒ o ThinQQ nunca o chama.
- **Recomendação:** sem recurso novo; o tutorial explica que "Quebrar vínculo" (pela interface do PowerPoint, Arquivo > Informações > Editar
  Links) congela setas e rótulos do ThinQQ e deixa os dados editáveis em "Editar Dados" (o motor já conta "Sem vínculo com o Excel"), e que o
  caminho recomendado para editar é a pasta de origem. Verificação manual (F5-6) confirma o comportamento pela interface.

## 6. Matriz de paridade

| Funcionalidade do think-cell (integração PowerPoint) | ThinQQ 1.5.0 (hoje) | ThinQQ depois | Fora do escopo |
|---|---|---|---|
| Gráfico vinculado a intervalo do Excel | Sim ("Copiar como gráfico vinculado") | Sim | |
| Criar a partir do intervalo direto no slide | Não (cria no Excel e copia) | **1.7.0** "Inserir também no slide" | posicionar com o mouse |
| Moldura colorida no intervalo vinculado | Não | **1.7.0** moldura verde | vermelho/azul, arrastar para mudar o vínculo, transpor, layout |
| Atualização manual (Update Once) | Sim (botões no Excel e no PowerPoint, apresentação inteira) | + **1.6.0** "Atualizar uma vez" só nos selecionados | |
| Atualização automática por elemento (Update Automatically) | Não | **1.6.0** | |
| Recarregar ao abrir / nova sessão | Sim (perguntar/sempre/nunca) | Sim | estado "Not yet loaded" |
| Fonte ausente: elemento fica com os últimos dados | Sim (aviso com o arquivo) | Sim | ícones de estado por elemento |
| Setas/decorações recalculadas com os dados | Sim (motor) | Sim (também no automático) | |
| Diálogo Data Links (lista, filtros, Revert) | Não (PowerPoint: Editar Links) | Não | sim |
| Religar a fontes alternativas (Manage Data Sources) | Não (Alterar Origem do PowerPoint) | Não | sim |
| **Inserir gráfico pelo PowerPoint (Insert > Elements), todos os tipos** | Não | **1.8.0** (G-I) "Inserir gráfico" na aba ThinQQ do PowerPoint, 21 tipos, dados de exemplo | posicionar arrastando um retângulo (nasce centralizado, 80% do slide); tipos do think-cell que o ThinQQ não tem |
| Datasheet que abre ao inserir / duplo clique, slide muda ao digitar | Não | **1.8.0** (G-I) "Editar dados": a planilha do gráfico num Excel visível; o slide acompanha pelo automático (~2 s) | duplo clique no gráfico (abre o "Editar Dados" nativo do PowerPoint); atualização a cada tecla |
| Formatar/decorar o gráfico pelo PowerPoint (rótulos, totais, setas) | Não | **1.9.0** (G-II) rótulos, totais, conectores, formato, legenda, CAGR, diferença, remover setas | delta por dois cliques no slide (continua no Excel, que "Editar dados" abre) |
| Tabela vinculada | Não | **1.10.0** (formatação inicial do Excel) | imagem de tabela, mescladas, bordas |
| Texto vinculado a célula em caixa de texto | Não | **1.9.0** (um campo por forma) | vários campos por forma |
| Texto vinculado em rótulos do gráfico | Não | **1.9.0** título do gráfico | rótulos de dados ligados a célula |
| Imagem por URL, bolas de Harvey, caixas de seleção ligadas | Não | Não | sim |
| Tableau | Não | Não | sim (§5.5) |
| Datasheet interno / slide workbook (dados DENTRO do .pptx) | Não | Equivalente pela Parte G com a pasta de dados AO LADO da apresentação + "Empacotar para enviar" (G-II) | dados dentro do .pptx (§5.6, §14.2) |

## 7. Faixas de opções

### 7.1 Excel (loader, menu `mnuPpt`)
1.6.0: depois de `btnPptSync`, `menuSeparator id="msepPptAuto"`, `checkBox id="chkPptAutoOn"` "Atualização automática ligada"
(`tag="autoon"`, `getPressed="Rx_GetPptFlag"`, `onAction="Rx_OnPptFlag"`, supertip "Com o Excel e o PowerPoint abertos, atualiza
sozinho os gráficos marcados como 'Atualizar automaticamente' cerca de 2 segundos depois que um valor da pasta muda. Desmarque para pausar.")
e `checkBox id="chkPptAutoNew"` "Novos vínculos atualizam automaticamente" (`tag="autonew"`, mesmas chamadas, supertip "Gráficos colados ou
inseridos no PowerPoint a partir de agora já nascem com 'Atualizar automaticamente' ligado."). 1.7.0: antes de `btnPptImage`,
`checkBox id="chkPptInsert"` "Inserir também no slide" (`getPressed="Rx_GetPptInsert"`, `onAction="Rx_OnPptInsert"`) e, depois de `chkPptAutoNew`,
`button id="btnPptFrames"` "Mostrar/ocultar molduras" (`onAction="Rx_OnPptFrames"`, `image="tqq_ppt16"`: ícone próprio já existente,
sem arriscar um `imageMso` não verificado). 1.9.0 (E): `button id="btnPptLinkText"` "Vincular célula ao texto selecionado"
(`onAction="Rx_OnPptLinkText"`) e, no grupo do gráfico, nada (título: `button id="btnLinkTitle"` "Título ligado a uma célula..." em `mnuMore`,
`onAction="Rx_OnLinkTitle"`, `getEnabled="Rx_GetEnabled"`). 1.10.0 (F): `button id="btnPptTable"` "Inserir tabela vinculada no slide"
(`onAction="Rx_OnPptTable"`, sempre ativo: age sobre as células selecionadas).

### 7.2 PowerPoint (`.ppam`)
1.6.0: grupo `grpPptElem` "Gráfico selecionado": `toggleButton id="tglPptAuto"` "Atualizar automaticamente" (`imageMso="Refresh"`,
`getPressed="PptRx_GetAuto"`, `onAction="PptRx_OnAuto"`, `getEnabled="PptRx_GetElemEnabled"`) e `button id="btnPptUpdateSel"` "Atualizar uma
vez" (`imageMso="Refresh"`, `onAction="PptRx_OnUpdateSel"`, `getEnabled="PptRx_GetElemEnabled"`). 1.8.0: `button id="btnPptLinkText"` "Texto
vinculado ao Excel" (`imageMso="TextBoxInsert"`, `onAction="PptRx_OnLinkText"`) — agora na 1.9.0. Grupo renomeado para "Elemento
selecionado" na 1.9.0. **Parte G** (§14.6): 1.8.0 grupo `grpPptInsert` "Inserir gráfico" (primeiro grupo da aba) com os menus `pmnuColumns`,
`pmnuBars`, `pmnuWaterfall`, `pmnuLines`, `pmnuOthers` (mesmos rótulos, ícones `image=` e supertips dos menus do Excel; 21 botões
`pbtn<Tipo>` com `tag` = tipo e `onAction="PptRx_OnInsertChart"`) e, no `grpPptElem`, `btnPptEditData` "Editar dados" (`imageMso="TableInsert"`,
`onAction="PptRx_OnEditData"`, `getEnabled="PptRx_GetElemEnabled"`); 1.9.0 grupo `grpPptChart` "Gráfico do ThinQQ" (§14.5).

## 8. Ordem, versões e por quê
1. **1.6.0 — Atualização automática (lacuna 1).** Maior valor (o pedido central do usuário: "fiel ao think-cell") e infraestrutura das demais:
   índice da apresentação, `PptSync_Items` com roteamento por tipo, motor de eventos/`OnTime`, `CORE_PPT_API = 2`. Maior risco de descoberta
   (eventos entre aplicativos): começar por ele dá tempo para a verificação manual acontecer enquanto as próximas versões são feitas.
2. **1.7.0 — Inserir a partir do intervalo + moldura (lacuna 2).** Pequena e independente no código (pode ser feita em paralelo com a
   verificação manual da 1.6.0); usa `AutoNew` se existir.
3. **1.8.0 — Criar gráficos pelo PowerPoint (Parte G-I, manchete; emenda 1).** Precisa da C (o automático é o que faz o slide acompanhar a
   digitação na pasta de dados; índice; API 2) e da D1 (moldura verde que destaca os dados ao "Editar dados"). Não precisa de E nem F.
4. **1.9.0 — Comandos do gráfico pelo PowerPoint (G-II) + texto vinculado (E).** G-II reaproveita a pasta de dados e o despacho da G-I; E introduz
   o modelo "nome definido + texto exibido + elemento não-gráfico", que a F reaproveita.
5. **1.10.0 — Tabelas vinculadas (F).** Reusa a E; maior volume de formatação.
Alternativas: (a) com pressa pela manchete, **1.6.0 = C + D** (independentes; D é pequena) e **1.7.0 = G-I**, adiantando a criação pelo
PowerPoint em uma versão (o mesmo esforço total, um release a menos); (b) menos releases: 1.9.0 = G-II + E + F.
API do core para o `.ppam` (`CORE_PPT_API`): 2 = C1 (`Core_PptSyncShapes`, flags); 3 = G-I (`Core_PptInsertChart`, `Core_PptEditData`,
`Core_PptDataSave`; o reapontamento é só do `.ppam`, `PptTqq_RelinkData`); 4 = G-II (`Core_PptCommand`, `Core_PptCategories`; "Empacotar" é só do `.ppam`); 5 = E1 (`Core_PptLinkText`).

`version.json` não é tocado pelas tarefas (o `publicar-release.bat` sobe as versões). Compatibilidade: `.ppam` 1.6.0 com core 1.5.x
(`Core_PptApiVersion = 1`): "Atualizar automaticamente" grava a etiqueta (sem Excel) mas o motor não existe ⇒ aviso "Atualize o ThinQQ no Excel
para a atualização automática funcionar."; "Atualizar uma vez" cai para a apresentação inteira. Loader 1.6.0 com core 1.5.x: as caixas
ficam desmarcadas e o clique avisa "Atualize o ThinQQ para usar este recurso." (`RunCoreFeature`).

## 9. Riscos e mitigações

| # | Risco | Prob. | Mitigação |
|---|---|---|---|
| R1 | EDR (SentinelOne) mata o Excel se o VBA iniciar processo filho | certa se acontecer | Nenhum fluxo novo inicia processo; `Test-VbaLint` já barra `Shell`/`.Exec`/`powershell`; investigar morte do Excel pelo log do SentinelOne primeiro |
| R2 | PowerPoint de instância única (testes e sondas anexariam ao do usuário) | certa | Testes reais só sem `POWERPNT` (padrão `StartScratchPowerPoint`); o motor automático em produção usa só o PowerPoint já aberto e nunca cria |
| R3 | Ciclo automático pesado (impressão digital SHA-256 em VBA, gráficos grandes) travando o Excel | média | Índice (S4), só a pasta que mudou, só elementos automáticos, espera de 2 s, chave global; P1 decide trocar a impressão por comparação barata |
| R4 | PowerPoint ocupado / caixa "aguardando ação OLE" no Excel | média | Erros RPC ⇒ novas tentativas espaçadas; guarda de texto em edição; P3 manual |
| R5 | Laço de eventos (o próprio refresh do ThinQQ recalcula `_ThinQQ`) | média | `mInTick` ignora eventos do ciclo; impressão digital pula sem mudança; teste dedicado |
| R6 | Índice desatualizado (slides copiados de outra apresentação, formas renomeadas) | média | Reconstruído em toda sincronização completa; entradas velhas descartadas; documentado |
| R7 | Modo de Exibição Protegido | baixa | Apresentação protegida não está em `Presentations` editável (`ReadOnly`/janela protegida): o motor não a toca; ao habilitar edição vale o fluxo de 2026-10-14 |
| R8 | Segurança: pastas de origem escolhidas pela apresentação | média | Texto e tabela usam o mesmo `PptSync_OpenBook` (somente leitura, `AutomationSecurity = 3`, eventos desligados ao abrir/fechar); o automático nunca abre pasta (só a que mudou, já aberta); "Atualizar sempre" com origem remota continua perguntando |
| R9 | Pasta não salva (nome definido e vínculo precisam do arquivo) | certa | Inserir/vincular recusa antes de criar qualquer coisa, com a mensagem; o automático lê a pasta aberta (inclusive não salva, como o vínculo nativo) |
| R10 | Moldura atrapalhar cliques nas células | baixa-média | P4 manual; desenho isolado (`LinkFrame_Draw`) troca para 4 linhas; botão para ocultar |
| R11 | Edição do texto vinculado no PowerPoint | média | Só os caracteres do campo mudam; edição incompatível não é sobrescrita (contada no resumo) |
| R12 | `BreakLink` por automação derruba o PowerPoint (S10) | certa se chamado | Nunca chamado; lacuna 6 fora do escopo |
| R13 | Memória baixa (Excel + PowerPoint + `-Embedding`) | média | Uma instância de cada nos testes; o automático só usa a pasta aberta (sem `-Embedding`) |
| R14 | Tempo da suíte crescer além de 1200 s | média | Módulos novos com `-Filter` na iteração; testes de PowerPoint reais agrupados |
| R15 | Desalinhamento core × `.ppam` (índice gravado pelos dois) | baixa | Funções puras idênticas com os mesmos testes de cadeia literal nos dois projetos (`Test_PptIdx_Literal`) |
| R16 | (G) Ida e volta PowerPoint → Excel → PowerPoint (`xl.Run` do `.ppam`; o core cola de volta pela área de transferência) | média | Mesmo padrão já medido no e2e da 1.4.0 (o core chama a apresentação durante o `xl.Run`); PG3 confirma a colagem nesse contexto; falha desfaz a planilha criada |
| R17 | (G) Janela do Excel: trazer à frente a partir do PowerPoint (bloqueio de primeiro plano do Windows) | média | `AppActivate` pelo `.ppam` (o PowerPoint é o processo em primeiro plano no clique); se falhar, a barra de tarefas pisca e o aviso diz "o Excel abriu a planilha de dados"; PG2 |
| R18 | (G) Excel criado pelo suplemento sem o ThinQQ carregado (automação não carrega suplementos) | média | Abrir o loader `%LOCALAPPDATA%\ThinQQ\ThinQQ.xlam` nessa instância (carrega o core e a faixa) e entregá-la ao usuário (`Visible`, `UserControl`); sem loader, o core pelo caminho de hoje; PG1 |
| R19 | (G) Pasta de dados não salva / apresentação movida, renomeada ou enviada sem a pasta | alta | O automático salva a pasta de dados no fim de cada ciclo; "Salvar" no PowerPoint salva a pasta (evento `PresentationSave`); ao abrir, vínculo quebrado com a pasta de dados ao lado ⇒ reapontado sem perguntar; "Empacotar para enviar" (G2); limitação documentada (think-cell guarda no .pptx) |
| R20 | (G) Pasta de dados aberta por outra pessoa (rede) ou somente leitura | baixa | Inserir/editar recusa com a mensagem; atualizar continua lendo |
| R21 | (G) Segurança: o core abre a pasta de dados PARA EDIÇÃO (não só leitura) | média | `AutomationSecurity = 3` também na abertura para edição; a pasta é `.xlsx` (sem macros) marcada `TQQ_DATABOOK`; caminho só derivado da própria apresentação (mesma pasta), os comandos do PowerPoint (G2) só alteram pasta do usuário (não marcada) se ela já estiver aberta no Excel; "Editar dados" a abre para edição por ação explícita do usuário, também com macros desligadas |
| R22 | (G) Apresentação no OneDrive/SharePoint aberta pelo endereço https | média | Inserir recusa ("salve numa pasta local ou sincronizada"); PG6 mede o `FullName` real |
| R23 | (G) Desfazer do Excel some a cada ciclo do ThinQQ (macro limpa a pilha) | certa | Já acontece hoje com o refresh dos gráficos; documentado |

## 10. Fora do escopo (e por quê)
Tableau (§5.5); datasheet interno (§5.6); diálogo Data Links, Revert, religar fontes alternativas (o PowerPoint tem "Editar Links para
Arquivos > Alterar Origem"); ícones de estado por elemento; imagem por URL, bolas de Harvey, caixas de seleção; vários campos por forma de
texto; rótulos de dados ligados a célula; imagem de tabela; células mescladas e bordas em tabelas; transpor/editar layout do vínculo; mover a
moldura para mudar o vínculo; atualização automática sem o Excel aberto ou para pastas fechadas (o automático reage a mudanças na pasta aberta;
pastas fechadas continuam pelo botão e ao abrir); Mac; 64 bits só verificado manualmente se houver máquina.
**Parte G (§14.8):** dados dentro do .pptx (o datasheet interno continua fora: a Parte G entrega a MESMA experiência — inserir no slide,
digitar, ver o slide mudar — com o Excel como datasheet numa pasta ao lado da apresentação); posicionar o gráfico arrastando um retângulo;
duplo clique no gráfico abrindo a pasta de dados (o PowerPoint abre o "Editar Dados" nativo do vínculo; o ThinQQ não intercepta);
reimplementar o motor de gráficos no PowerPoint (gráfico nativo com pasta incorporada, §14.1); delta por dois cliques no slide; colar dados
direto no PowerPoint; tipos do think-cell que o ThinQQ não tem.

## 11. Pré-flight (interfaces entre tarefas)

| Par | Produz | Consome | Resultado |
|---|---|---|---|
| C0 → todas | spec/plano em `docs/superpowers/` | Test-Docs/Privacy | limpo |
| C1 → C2 | `PptIdx_*` (puras), `PptSync_SetAuto`, `PptSync_Items(items, force)`, `PptSync_ItemBook`, `PPTSYNC_TAG_AUTO/INDEX`, `Ppt_IsBusyError` | `PptAuto_ItemsFor`, `PptAuto_Tick` | limpo |
| C1 → C3 | formato do índice (cadeia literal do teste `Test_PptIdx_Literal`) e `Core_PptSyncShapes(pres, entries, force)` (API 2) | `PptTqq_IdxSet`, `PptTqq_UpdateShapes` | limpo: os dois projetos testam a mesma cadeia |
| C1 → C4 | `Core_PptGetFlag/Core_PptSetFlag("autoon"/"autonew")` | `Rx_GetPptFlag/Rx_OnPptFlag` | limpo |
| C2 → C5 | mensagens/limites do automático | tutorial, README, Test-Docs | limpo |
| C1 → D2 | `PptAuto_NewLinksAuto` lido na colagem | `Core_InsertToPowerPoint` (via `Core_CopyToPowerPoint`) | limpo |
| D1 → D2 | `LinkFrame_Ensure(c, rng)`, `LinkFrame_Remove`, `LinkFrame_ToggleAll`, `LinkFrame_CleanOrphans` | `Core_InsertToPowerPoint`, `Core_RebuildFromRange`, `Watchers_Rehook` | limpo |
| C1 → E1 | `PptSync_ItemKind`, rota "text" em `PptSync_Items` | `PptText_Sync` | limpo |
| E1 → F2 | `PptLink_DefineName`, `PptLink_ResolveRange`, `PptText_CellText` | `PptTable_Insert/Sync` | limpo |
| C2 → E1/F2 | `PptAuto_IsCandidate` (amplia: nomes `TQQ_X_`/`TQQ_T_`) | automático de textos/tabelas | limpo |
| E1 → E3 | `Core_PptLinkText(shp, start, length)` (API 5) | `.ppam` `PptTqq_LinkText`, `Rx_OnPptLinkText` | limpo |
| G1 → G2 | `Samples_TableFor(kind)` | `PptData_NewChart` | limpo: teste cobre os 21 tipos |
| C1/C2 → G2 | `PptSync_SetAuto`, motor automático, `PptSync_Items` | gráfico criado no PowerPoint nasce automático; "Editar dados" depende do ciclo | limpo: sem C a G funciona com "Atualizar uma vez" |
| D1 → G3 | `LinkFrame_Ensure/Find` | `PptData_ShowData` (destaca os dados) | limpo: sem D1 só seleciona o intervalo |
| modCoreApi → G2 | `Ppt_PasteChart` (público), `PptSyncAfterPaste` (privado do mesmo módulo) | `Core_PptInsertChart` (em `modCoreApi`) | limpo |
| G2/G3 → G4 | `Core_PptInsertChart(slide, kind)`, `Core_PptEditData(pres, entries)`, `Core_PptDataSave(pres)` (API 3); `PptTqq_RelinkData(pres)` no `.ppam` (sem Excel) | `.ppam` `PptTqq_InsertChart`, `PptTqq_EditData`, eventos | limpo |
| G4 → G8 | `PptTqq_GetEditExcel`, grupo `grpPptElem` | grupo `grpPptChart` | limpo |
| G7 → G8 | `Core_PptCommand(pres, entries, cmd, arg)`, `Core_PptCategories`, etiquetas `TQQ_KIND`/`TQQ_OPTS` (API 4) | botões do PowerPoint | limpo |

| Tarefa | Autoconsistência | Resultado |
|---|---|---|
| C1 | refatoração `PptSync_Presentation` → `PptSync_Items` sem mudar textos; testes antigos de `TestPptSync` continuam | limpo |
| C2 | `PptAuto_Tick` chamável pelo teste sem `OnTime` (gancho) | limpo |
| C3 | `.ppam` grava etiqueta/índice sem Excel; API 2 só para "Atualizar uma vez" | limpo |
| D2 | conferência de pasta salva antes de `Core_Insert` | limpo |
| E1 | campo único; edição incompatível nunca sobrescrita | limpo |
| F2 | diferença pura testada sem PowerPoint | limpo |

## 12. Estimativa

**Calibração (medido neste projeto):** planejamento + sondas das Partes A+B ≈ 2,5 h; Parte A (10 tarefas, 34 commits, ~9,5 mil linhas) e
Parte B (7 tarefas, 19 commits, ~3,5 mil linhas) ≈ 1,5–2 dias corridos de ponta a ponta com esperas; implementador 10–45 min (integração até
~60), revisão 3–8 min (revisor mais capaz até ~20), rodada de correção 6–30 min, 1–3 rodadas por tarefa; suíte completa 9–14 min estável;
diagnóstico do EDR ~1,2 h; revisões finais por parte acharam 1 crítico de segurança e 1–3 importantes. Este planejamento: ~2 h (com
interrupção para a gravação; as sondas P1–P8 pendentes somam ~0,5–1 h de agente quando agendadas).

**Suposições:** controlador com subagentes (implementador + revisor + 1–2 rodadas; 2–3 nas tarefas de eventos/PowerPoint); suíte por tarefa
(9–14 min; +50% de reexecução no pessimista); nenhum PowerPoint do usuário aberto nas rodadas que testam o PowerPoint (senão os testes reais
pulam e a verificação passa ao humano); VBOM do PowerPoint já habilitado (foi na 1.4.0); revisões em paralelo com a tarefa seguinte quando não
mexem no mesmo módulo (corrido ~10% abaixo da soma); uma onda de correção da revisão final por versão.

### 12.1 Por tarefa (minutos; otimista / provável / pessimista)
| Tarefa | Agêntico | Humano | Observação |
|---|---|---|---|
| C0 Promover spec/plano | 10 / 15 / 25 | — | |
| C1 Índice, `TQQ_AUTO`, `PptSync_Items`, API 2, flags | 60 / 95 / 140 | — | refatoração do motor; testes reais de PowerPoint |
| C2 Motor automático (eventos + `OnTime` + ciclo) | 100 / 160 / 250 | — | maior risco de descoberta (eventos, ocupado) |
| C3 `.ppam`: grupo "Gráfico selecionado" | 90 / 150 / 240 | — | build do `.ppam`, corredor próprio |
| C4 Faixa do Excel (2 caixas) | 35 / 50 / 75 | — | |
| C5 Documentação 1.6.0 | 45 / 70 / 100 | — | |
| C6 Suíte final + revisão do ramo + onda | 45 / 80 / 120 | 120 / 180 / 270 | humano: §5.1 manual, PR, release |
| **1.6.0** | **385 / 620 / 950 (6,4 / 10,3 / 15,8 h; corrido ~5,8 / 9,3 / 14,3 h)** | **2 / 3 / 4,5 h** | confiança média-baixa |
| D1 `modLinkFrame` | 50 / 80 / 120 | — | |
| D2 `Core_InsertToPowerPoint` + faixa | 70 / 110 / 170 | — | |
| D3 Documentação 1.7.0 | 35 / 55 / 80 | — | |
| D4 Final + revisão + onda | 40 / 70 / 100 | 60 / 100 / 160 | humano: P4, inserir nos 21 tipos (amostra), PR, release |
| **1.7.0** | **195 / 315 / 470 (3,3 / 5,3 / 7,8 h; corrido ~3 / 4,8 / 7 h)** | **1 / 1,7 / 2,7 h** | confiança média-alta |
| G1 Dados de exemplo dos 21 tipos (`Samples_TableFor`) | 40 / 60 / 90 | — | emenda 1 |
| G2 Pasta de dados + `Core_PptInsertChart` (API 3) | 110 / 170 / 260 | — | ida e volta PowerPoint/Excel; PG3, PG6 |
| G3 Editar dados e salvar (`Core_PptEditData/DataSave`; salvar no ciclo automático) | 80 / 125 / 190 | — | PG4, PG5 |
| G4 `.ppam`: menu "Inserir gráfico" (21) + "Editar dados" + Excel visível + eventos | 120 / 190 / 300 | — | PG1, PG2; ícones próprios no `.ppam` |
| G5 Documentação 1.8.0 | 45 / 70 / 100 | — | |
| G6 Final + revisão + onda | 45 / 80 / 120 | 120 / 180 / 270 | humano: inserir os 21 tipos pelo PowerPoint, editar, salvar, mover, enviar |
| **1.8.0 (G-I)** | **440 / 695 / 1 060 (7,3 / 11,6 / 17,7 h; corrido ~6,6 / 10,4 / 15,9 h)** | **2 / 3 / 4,5 h** | confiança média-baixa (UI entre aplicativos) |
| G7 `Core_PptCommand`, categorias, `TQQ_KIND/OPTS` (API 4) | 80 / 125 / 190 | — | |
| G8 `.ppam`: grupo "Gráfico do ThinQQ" | 90 / 140 / 220 | — | |
| G9 Empacotar para enviar | 40 / 60 / 90 | — | |
| G10 Documentação G2 | 35 / 55 / 80 | — | |
| G11 Final G2 | 30 / 50 / 80 | 60 / 100 / 150 | |
| E1 Texto vinculado: modelo + motor | 90 / 140 / 220 | — | |
| E2 Título ligado à célula | 60 / 95 / 150 | — | P6 |
| E3 UI (`.ppam` + Excel), API 5 | 60 / 90 / 140 | — | P5 |
| E4 Documentação | 40 / 60 / 90 | — | inclui orientação Tableau/datasheet |
| E5 Final + revisão + onda | 40 / 70 / 100 | 60 / 100 / 150 | |
| **1.9.0 (G-II + E)** | **565 / 885 / 1 360 (9,4 / 14,8 / 22,7 h; corrido ~8,5 / 13,3 / 20,4 h)** | **2 / 3,4 / 5 h** | confiança média |
| F1 Tabela: modelo puro (diferença, serialização) | 70 / 110 / 170 | — | |
| F2 Criar/atualizar no slide + motor | 110 / 170 / 260 | — | formatação inicial; P8 |
| F3 Faixa | 35 / 50 / 75 | — | |
| F4 Documentação | 40 / 60 / 90 | — | |
| F5 Final + revisão + onda | 40 / 70 / 100 | 60 / 100 / 150 | |
| **1.10.0 (F)** | **295 / 460 / 695 (4,9 / 7,7 / 11,6 h; corrido ~4,4 / 6,9 / 10,4 h)** | **1 / 1,7 / 2,5 h** | confiança média |
| PR + release por versão (humano) | — | 0,5 / 0,75 / 1 h × 5 = 2,5 / 3,75 / 5 h | |
| **Total** | **1 880 / 2 975 / 4 535 min ⇒ corrido ~28 / 44,5 / 68 h de agente (≈ 4 / 6,5 / 10 dias corridos com esperas, pela cadência medida das Partes A+B)** | **10,5 / 16,5 / 24 h** | |
| *Antes da emenda 1* | *1 165 / 1 850 / 2 815 min (corrido ~17,5 / 27,8 / 42 h); humano 7 / 11 / 16 h* | | |
| *A Parte G acrescenta* | *+715 / +1 125 / +1 720 min (corrido ~+10,7 / +16,9 / +25,8 h; ≈ +1,5 / +2,5 / +4 dias corridos)* | *+3,5 / +5,5 / +8 h (inclui 1 release a mais)* | |

### 12.2 Incertezas que mais pesam
- Eventos entre aplicativos e o PowerPoint ocupado (P3): pode exigir mais 1–2 rodadas na C2 (+1–2 h) ou um modo mais conservador (só ao
  salvar a pasta).
- Custo da impressão digital (P1): pode acrescentar uma tarefa curta na C2 (+30–60 min).
- Leitura da seleção do PowerPoint a partir do Excel (P5): se não funcionar, a entrada do Excel da E3 (1.9.0) sai (−30 min) e só o `.ppam` vincula.
- Formatação "com a do Excel" da tabela tem muitas pontas (fontes do tema, cores automáticas); o plano limita a 6 propriedades.
- Gravação de vídeo e a verificação humana competem pelo mesmo Excel/PowerPoint reais: as janelas de teste com PowerPoint real precisam de
  horários sem o usuário usando o PowerPoint.
- (G) Janela do Excel e foco entre aplicativos (PG1, PG2): se o Excel criado pelo suplemento não ficar com o usuário ou não vier à frente,
  +1 rodada na G4 (+1–2 h). Reapontar vínculo (PG4): se `LinkFormat.SourceFullName` não for gravável em gráfico vinculado, o reapontamento
  passa a recolar o gráfico (perde animação/ordem Z, documentado) (+1 h na G3). A verificação humana da G-I (tarefa G6) é a mais longa do plano (UI real).

## 13. Fontes
- think-cell — Manage data links: https://www.thinkcell.com/en/resources/manual/manage-data-links
- think-cell — Excel data links: https://www.thinkcell.com/en/resources/manual/exceldatalinks
- think-cell — External data links (visão geral): https://www.thinkcell.com/en/resources/manual/external-data-links
- think-cell — Tableau data links: https://www.thinkcell.com/en/resources/manual/tableaudata
- think-cell — Element datasheets: https://www.thinkcell.com/en/resources/manual/element-datasheets
- think-cell — Slide workbooks: https://www.think-cell.com/en/resources/manual/slide-workbooks
- think-cell — think-cell 11 integrations (08/12/2020): https://www.thinkcell.com/en/blog/think-cell-integrations
- Microsoft — Application.OnTime (Excel): https://learn.microsoft.com/en-us/office/vba/api/excel.application.ontime
- Microsoft — PowerPoint Application object (eventos): https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application
- Microsoft — ChartTitle.Formula: https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.formula
- Microsoft — Shapes.AddTable (PowerPoint): https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addtable
- Microsoft — Slides.FindBySlideID: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slides.findbyslideid
- Microsoft — ChartData.BreakLink: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.chartdata.breaklink
- Spec anterior (sondas de vínculo, motor, segurança): `docs/superpowers/specs/2026-10-14-thinqq-ppt-delta-design.md`
- think-cell — Introduction to charting (inserir pelo PowerPoint, datasheet): https://www.thinkcell.com/en/resources/manual/introductioncharting
- Microsoft — Shapes.AddChart2 (gráfico nativo com pasta incorporada): https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addchart2
- Microsoft — ChartData.Activate / Workbook: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.chartdata.activate
- Microsoft — AppActivate (VBA): https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/appactivate-statement
- Microsoft — Application.UserControl (Excel): https://learn.microsoft.com/en-us/office/vba/api/excel.application.usercontrol
- Microsoft — PowerPoint Application.PresentationSave: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.presentationsave

## 14. Emenda 1 — Criar gráficos pelo PowerPoint (Parte G: 1.8.0 G-I e 1.9.0 G-II)

Pedido: "ajuste o plano para permitir a criação de gráficos diretamente pelo PowerPoint assim como o think-cell". **O que o think-cell faz**
(fonte: introductioncharting): no PowerPoint, grupo think-cell > **Elements** > tipo; um retângulo segue o ponteiro para posicionar; o
**datasheet** (planilha do Excel própria do elemento) abre sozinho; digita-se só os números (o think-cell calcula totais e arredondamento) e
"the chart on the slide instantly updates"; o datasheet reabre por duplo clique ou "Open Datasheet". **Não confirmado** pela página: se o
datasheet fica no .pptx (§2.1 Element datasheets também não diz; o comportamento conhecido do produto é que fica no arquivo).

**Fiel no ThinQQ =** na aba ThinQQ do PowerPoint, "Inserir gráfico" com os **mesmos 21 tipos e os mesmos 5 menus** da faixa do Excel; um
clique põe o gráfico no slide ativo com dados de exemplo e abre a planilha dos dados num Excel visível (o "datasheet"); o que se digita lá
aparece no slide em ~2 s (motor automático da 1.6.0), com setas, rótulos, totais e conectores do ThinQQ; "Editar dados" reabre a planilha;
numa 2.ª etapa, rótulos/totais/conectores/formato/legenda/CAGR/diferença pelo próprio PowerPoint. **Diferente:** os dados ficam numa pasta do
Excel **ao lado** da apresentação (não dentro do .pptx), a atualização é a cada ~2 s depois da última mudança (não a cada tecla), o gráfico
nasce centralizado (não se arrasta um retângulo) e existe uma cópia do gráfico na planilha de dados (é ela que o motor calcula).

### 14.1 Arquitetura (avaliada)
| Candidata | Como | Avaliação | Veredito |
|---|---|---|---|
| **G-A. Core no Excel cria numa pasta de dados e cola vinculado** | O `.ppam` obtém um Excel com o core (o do usuário se já tem o core; senão uma instância nova **visível**, com o loader aberto, entregue ao usuário) e chama `Core_PptInsertChart(slide, tipo)`: o core abre/cria a pasta de dados, cria a planilha "Gráfico N" com a tabela de exemplo, insere o gráfico pelo `Core_Insert` de sempre, salva, cola **vinculado** no slide (`Ppt_PasteChart` + `PptSyncAfterPaste`, o mesmo caminho do "Copiar como gráfico vinculado"), marca `TQQ_DATA`, liga "Atualizar automaticamente" e destaca os dados (moldura da D1) | Reusa 100% do motor (21 tipos, setas, rótulos, sincronização, automático); o PowerPoint chamar o Excel que chama o PowerPoint já é o padrão medido do e2e da 1.4.0 (`Core_PptSyncPresentation` mexe na apresentação durante o `xl.Run`); custo novo só na pasta de dados, no menu e no gerenciamento da janela do Excel | **Escolhida** |
| G-B. Gráfico nativo do PowerPoint com pasta incorporada (`Shapes.AddChart2`) formatado pelo ThinQQ | Os dados ficam no .pptx (como o think-cell) | O motor do ThinQQ é escrito sobre objetos do **Excel** (`Chart`/`ChartObject`/`Range`/`Worksheet`, planilha `_ThinQQ`, ~20 módulos de tipo + anotações): seria reescrever o motor para o modelo do PowerPoint (ou abrir a pasta incorporada por `ChartData.Activate`, que abre um Excel de incorporação sem o ThinQQ carregado e janelas próprias); S10 mostrou instabilidade de `ChartData` por automação. Estimativa grosseira: 5–10× a Parte G inteira, com risco alto | Rejeitada |
| G-C. Objeto OLE "Planilha do Excel" incorporado contendo o gráfico | Dados no .pptx, duplo clique edita no lugar | No slide fica uma **imagem** do gráfico (não um gráfico nativo do PowerPoint); a edição no lugar roda num Excel de incorporação sem o ThinQQ; não dá para o core de outro processo mexer nessa pasta | Rejeitada |
| G-D. Só um atalho: o `.ppam` abre o Excel numa pasta nova e o usuário usa o "Inserir também no slide" (D) | Barato | Não é "criar pelo PowerPoint" (o usuário escolhe o tipo no Excel) | Rejeitada como solução; vira o caminho "Editar dados" |

### 14.2 Pasta de dados (modelo)
- **Onde:** uma pasta por apresentação, **na mesma pasta do .pptx**: `<nome da apresentação sem extensão> - dados ThinQQ.xlsx` (`.xlsx`: sem
  macros). Marcada pelo nome oculto `TQQ_DATABOOK`. Primeira planilha "Leia-me" (o que é o arquivo, que ele acompanha a apresentação).
- **Uma planilha por gráfico:** "Gráfico 1", "Gráfico 2"… (o primeiro número livre); tabela em A1 (dados de exemplo do tipo, §14.7), o gráfico
  ThinQQ ao lado (onde o `Core_Insert` o põe), moldura verde nos dados (D1).
- **Etiquetas** (as do motor, mais uma): `TQQ_CHART` (id), `TQQ_BOOK` (caminho da pasta de dados), `TQQ_SHEET` ("Gráfico N"), `TQQ_VER`,
  `TQQ_FP`, `TQQ_AUTO = "1"` (nasce automático: é o "datasheet ao vivo") e **`TQQ_DATA = "1"`** (dados gerenciados pelo ThinQQ: podem ser
  abertos para edição, reapontados e empacotados). Na apresentação: `TQQ_DATABOOK` = caminho da pasta de dados, e o índice da 1.6.0.
- **Apresentação nunca salva:** "Inserir gráfico" pede para salvar antes (mensagem com o nome do arquivo que será criado) — limitação frente ao
  think-cell (que guarda no .pptx). **Apresentação pelo endereço https (OneDrive/SharePoint):** recusado com "Salve a apresentação numa
  pasta do computador (ou na pasta sincronizada do OneDrive)" (PG6).
- **Salvar:** o vínculo nativo e o motor leem a pasta **aberta** (inclusive alterações não salvas, medido); para o arquivo nunca ficar para
  trás: (1) o ciclo automático **salva** as pastas marcadas `TQQ_DATABOOK` que sincronizou; (2) o `.ppam` no evento `PresentationSave`
  chama `Core_PptDataSave(pres)` (salva a pasta de dados se estiver aberta e alterada; PG5); (3) fechar o Excel com a pasta alterada mostra o
  "Salvar?" normal do Excel.
- **Apresentação movida/renomeada/enviada:** o vínculo aponta para o caminho absoluto antigo. Ao abrir (fluxo da 1.4.0) e em "Editar dados",
  o `.ppam` roda `PptTqq_RelinkData(pres)` (só PowerPoint e `FileSystemObject`, sem Excel): para formas `TQQ_DATA` cuja pasta não existe, se existir `<nome atual> - dados ThinQQ.xlsx` ou o
  arquivo com o **nome antigo** na pasta atual da apresentação, reaponta (`LinkFormat.SourceFullName`, PG4; senão recola) e atualiza
  `TQQ_BOOK`/`TQQ_DATABOOK` — sem perguntar: o arquivo está na mesma pasta da apresentação e é marcado. Enviar por e-mail: **"Empacotar para
  enviar"** (G-II) copia o .pptx e a pasta de dados para uma pasta nova `<nome> (ThinQQ)` ao lado (o usuário compacta/anexa a pasta); sem a
  pasta de dados, quem recebe vê o gráfico com os últimos dados (como hoje) e o aviso de pasta não encontrada.
- **Concorrência:** pasta de dados já aberta no Excel usado ⇒ usada como está; aberta por outra pessoa (somente leitura) ⇒ inserir/editar
  recusam ("A pasta de dados está aberta por outra pessoa ou é somente leitura"); pastas do usuário nunca são tocadas (a de dados é outra).
- **Segurança:** aberta para edição com `AutomationSecurity = 3` (macros desligadas) e eventos desligados durante o `Open`; só o caminho
  derivado da própria apresentação (mesma pasta) ou a etiqueta de uma forma `TQQ_DATA` cuja pasta **tem** a marca `TQQ_DATABOOK` é aberto
  para edição; no "Editar dados" (§14.3) a pasta de origem do usuário também é aberta para edição (macros desligadas; já aberta, mesmo
  somente leitura, é usada como está); a sincronização segue o `PptSync_OpenBook` (somente leitura). Só extensões de pasta do Excel
  (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`) são abertas, e uma forma `TQQ_DATA` só com um arquivo `<nome> - dados ThinQQ.xlsx` (revisão final, I2).
  Sem processo filho (o Excel é obtido por `GetObject`/`CreateObject`, como hoje).

### 14.3 Editar dados (UX)
- Botão **"Editar dados"** (`btnPptEditData`, grupo "Gráfico selecionado"), também depois de inserir (automático), e a dica no supertip do
  grupo ("Selecione o gráfico e clique em Editar dados"). Duplo clique no gráfico abre o "Editar Dados" nativo do vínculo (fora do controle do
  ThinQQ; documentado).
- `Core_PptEditData(pres, entries)`: abre a pasta da forma **para edição com as macros desligadas** (a de dados e também a do usuário, para
  gráficos colados do Excel — o caminho vem da apresentação, que pode ser de outra pessoa; se for `.xlsm` o aviso diz "as macros desta pasta
  não rodam quando ela é aberta pelo ThinQQ; abra-a pelo Excel se precisar delas"); já aberta no Excel ⇒ usada como está;
  `Application.Visible = True`, ativa a pasta e a planilha do gráfico (`TQQ_SHEET`, ou a planilha do ChartObject), seleciona o intervalo dos
  dados (o da moldura ou `CurrentRegion` de A1) e devolve o `Caption` do Excel; o `.ppam` faz `AppActivate` nessa legenda (PG2).
- **Qual Excel:** `PptTqq_GetEditExcel`: (1) o Excel aberto do usuário **com** o core; (2) senão uma instância nova, **visível**, com
  `UserControl = True`, onde o `.ppam` abre o loader instalado (`%LOCALAPPDATA%\ThinQQ\ThinQQ.xlam`, que carrega o core e a faixa: o usuário
  ganha a faixa ThinQQ para usar todos os recursos, inclusive o delta por cliques) — nunca encerrada pelo suplemento (o usuário fecha); (3)
  Excel do usuário aberto **sem** o ThinQQ ⇒ também (2), em janela própria (nunca carrega o ThinQQ no Excel de quem não o tem). PG1.
- **Primeira vez:** depois de inserir, aviso único por sessão do PowerPoint: "Digite os dados na planilha 'Gráfico 1' do Excel: o slide
  acompanha em cerca de 2 segundos. Os dados ficam em '<arquivo>' ao lado da apresentação: envie os dois arquivos juntos (ThinQQ > Empacotar
  para enviar)."

### 14.4 Desempenho (estimado com medidas anteriores)
Excel já aberto com o core: criar planilha + `Core_Insert` (~0,3–1 s) + salvar (~0,2 s) + colar e ajustar ao slide (`PptSyncAfterPaste`,
~0,5–1 s) ⇒ ~1–3 s. Excel frio: iniciar Excel + loader + core ~5–10 s (não medido: PG8), com a barra de status do PowerPoint não
disponível ⇒ o `.ppam` não mostra andamento (aviso no fim).

### 14.5 Comandos do gráfico pelo PowerPoint (G-II, 1.9.0)
- API 4: `Core_PptCommand(pres, entries, cmd, arg) As String` — resolve as formas (`PptSync_ResolveEntries`), abre a pasta (de dados: edição;
  do usuário: só se já aberta no Excel, senão "Abra a pasta de origem no Excel para mudar o gráfico (o ThinQQ não altera pastas fechadas)"),
  aplica pelo mesmo código da faixa do Excel (`ApplyToggle`, `ApplyOption`, `Ann_AddCagr`, `Ann_AddDifference`, `Ann_RemoveAll`), `Core_SafeRefresh`,
  salva se for pasta de dados e sincroniza **forçado** só aquelas formas (`PptSync_Items`). Comandos v1: `toggle:labels|totals|connectors`,
  `numfmt:<id>`, `legend:right|none`, `cagr:<de>,<até>`, `diff:<de>,<até>`, `clear`. Recusa da matriz (7001/7002) volta como texto.
- Estado na faixa **sem chamar o Excel**: a cada sincronização/colagem o core grava na forma `TQQ_KIND` (tipo) e `TQQ_OPTS`
  (`labels=1;totals=0;connectors=1;numfmt=int;legend=right;caps=labels,totals,...`): o `.ppam` lê para `getPressed`/`getEnabled`.
- `Core_PptCategories(pres, entry) As String` (nomes das categorias separados por `Chr(31)`) para o `.ppam` montar a escolha de/até da CAGR e
  da diferença (`InputBox` com a lista numerada).
- **Delta por dois cliques:** fora no PowerPoint (precisa dos eventos `Select`/`MouseUp` de um gráfico do Excel); o caminho é "Editar dados" e
  o delta no Excel (o slide acompanha pelo automático). PG7 avalia uma seleção pelo PowerPoint para o futuro.
- "Empacotar para enviar" (só no `.ppam`, `PptTqq_Package`, sem o Excel: cópia por `FileSystemObject` — a apresentação precisa estar
  salva; copia o .pptx e a pasta de dados para `<pasta>\<nome> (ThinQQ)\`; sem `Shell`, sem ZIP).

### 14.6 Faixa do PowerPoint (G-I)
Grupo `grpPptInsert` "Inserir gráfico", primeiro da aba: os 5 menus com os mesmos ids prefixados por `p` (`pmnuColumns` ...), rótulos,
`image=` (`tqq_columns`, `tqq_bars`, `tqq_waterfall`, `tqq_lines`, `tqq_more`; o build do `.ppam` já usa `-Icons Png`), supertips e os 21
botões (`pbtnStacked` ... com `tag` = tipo, `onAction="PptRx_OnInsertChart"`; os 4 botões com ícone próprio no Excel repetem o `image`). Um
teste (Test-Ribbon.ps1) confere que os 21 `tag` do PowerPoint são exatamente os 21 do Excel e na mesma ordem por menu.

### 14.7 Tipos e dados de exemplo
**Todos os 21 tipos já na 1.8.0**: o custo por tipo é só a tabela de exemplo (o motor é o mesmo); fasear economizaria pouco (≈ 15 min) e
deixaria menus incompletos em relação ao Excel. Nova função pura `Samples_TableFor(kind)` (tabela por família, no formato que o `ChartData`
de cada tipo exige — os mesmos dados dos exemplos da pasta "Exemplos" quando há): colunas/barras/áreas/linha = Ano × 3 regiões; cascatas
com "e" de total; combinação com uma linha de %; Mekko (%) e (unidades, com Largura); Gantt com datas; pizza/rosca uma coluna; borboleta 2
colunas; dispersão Nome/X/Y; bolhas + Tamanho; cascata comparativa 2 colunas. Teste: para cada um dos 21, a tabela passa em
`ChartData_FromRange(...).Validate()` e o `Core_Insert` cria o gráfico.

### 14.8 Limitações honestas frente ao think-cell
Dados ao lado do .pptx, não dentro (enviar os dois arquivos; reapontamento automático na mesma pasta; "Empacotar"); precisa do Excel instalado
e do ThinQQ no Excel (o think-cell também exige o próprio suplemento); a apresentação precisa estar salva antes do primeiro gráfico; o slide
muda ~2 s depois da última mudança, não a cada tecla; uma cópia do gráfico mora na planilha de dados; duplo clique no gráfico abre o "Editar
Dados" do PowerPoint (pasta de dados num Excel sem o ThinQQ se o vínculo for aberto por ali), não o "Editar dados" do ThinQQ; posicionamento
por arrastar não existe; delta por cliques só no Excel; sem colar dados direto no PowerPoint; desfazer do Excel some após cada ciclo.

### 14.9 Verificação manual (G6/G11)
(1) inserir os 21 tipos pelo PowerPoint com o Excel fechado (frio, PG8) e aberto; (2) a aba ThinQQ do PowerPoint com os ícones (PG9);
(3) "Editar dados" traz o Excel à frente (PG2) e o Excel continua aberto depois (PG1); (4) digitar e ver o slide mudar; (5) salvar no
PowerPoint salva a pasta de dados (PG5); (6) mover a pasta com os dois arquivos, renomear a apresentação, abrir: reapontado (PG4); (7) abrir
a apresentação em outro computador sem a pasta de dados: últimos dados + aviso; (8) Excel do usuário aberto sem o ThinQQ (janela própria);
(9) OneDrive (PG6); (10) comandos G2 pelo PowerPoint em 5 tipos; (11) "Empacotar para enviar".
