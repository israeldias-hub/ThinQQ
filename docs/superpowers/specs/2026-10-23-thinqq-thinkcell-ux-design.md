# ThinQQ — Pesquisa e Design: UX e funcionalidades como o think-cell, rodada 2 (1.11.0 a 1.13.0)

Data: 2026-10-23. Base: `main` em 3a30f38 (ThinQQ 1.10.1 publicado). Autor: planejador (opus). Plano: `plan.md` nesta pasta; briefs: `task-<ID>-brief.md`.
Regras de ambiente: `.superpowers/sdd/2026-10-15-thinkcell-parity/implementer-rules.md` (valem sem exceção, inclusive EDR, PowerPoint de instância única,
`Get-CimInstance Win32_Process` em vez de `tasklist` com vários `/FI`, sessão bloqueada quebra a área de transferência, teto do tutorial 655.360 B (720.896 B desde a revisão final da 1.12.0),
disciplina Test-Docs/CHANGELOG).

## 1. Objetivo e escopo

O usuário (2026-10-23): "A funcionalidade ainda não está boa... UX ruim, a funcionalidade de atualizar a cada alteração também é ruim e bem diferente
do think-cell original. [...] A ideia é que a aplicação seja a mais semelhante possível com o think-cell em UX e funcionalidades." Pedido adicional
(mesmo dia, via coordenador): "permita também que o usuário crie padrões de gráficos (paleta de cores e salve)", com um gráfico de cascata de
exemplo (seção 8.5).

Três problemas a resolver, nesta ordem (a ordem é a das queixas):
1. **Modelo de atualização**: hoje (1.6.0+) o ThinQQ atualiza o slide ~2 s depois de cada mudança da pasta (motor `OnTime`, ligado por padrão e
   com os gráficos criados pelo PowerPoint nascendo automáticos). O think-cell é **manual por padrão**, mostra que há dados novos (bandeira verde /
   moldura verde e botão Update ao selecionar o elemento) e só atualiza sozinho quem o usuário marcou "Update Automatically"; a central é o diálogo
   **Data Links**.
2. **UX**: comandos espalhados, estados escondidos, caixas de mensagem para tudo, nada aparece "no elemento" (nem barra de ferramentas, nem botão
   direito). Proposta: uma aba contextual do ThinQQ que aparece ao selecionar um gráfico, o diálogo Vínculos (não modal), silêncio no sucesso,
   um único "Vincular ao PowerPoint" no Excel, menus de botão direito, moldura com estado.
3. **Estilos de gráfico** (o "Style file / Color Scheme" do think-cell): padrões nomeados (paleta por papel, hachura, conectores, rótulos, fonte,
   espaço), salvos por usuário e na pasta, importáveis/exportáveis, aplicáveis e padrão para novos gráficos, que sobrevivem a atualizar/Alterar
   dados e viajam para o PowerPoint.

Fora desta rodada (com decisão na matriz §9): datasheet interno, Slide Workbook, Tableau, barra de ferramentas flutuante/radial real, alças de
arrastar, Revert, ferramentas de slide (alinhar, limpar, enviar), Harvey balls/caixas de seleção, Gantt com calendário. Ficam no backlog 1.14.0+.

## 2. Fontes (o que foi tirado de cada uma)

| Fonte | O que foi usado |
|---|---|
| Guia do usuário "think-cell 14 Step by Step Guide for Beginners v1.4, Nov 2025" (PDF do usuário, resumido no pedido) | Clique esquerdo = barra de formatação do elemento; clique direito = menu radial; alças; datasheet por duplo clique; Slide Workbook; "Link to PowerPoint" na guia Inserir do Excel com galeria; moldura verde com alças brancas; **bandeira verde** no canto do gráfico quando o Excel mudou, clique atualiza; botão "Auto Update" por gráfico; diálogo Data Links (lista por slide, filtro por estado, Select All, Update/Revert/Auto Update); Transpose Link; Break Excel Link; Gantt; ferramentas; Style files. |
| https://www.thinkcell.com/en/resources/manual/overview | Índice completo do manual (capítulos e URLs). |
| https://www.thinkcell.com/en/resources/manual/manage-data-links | **Modelo de atualização**: "By default, elements that are linked to Excel update manually"; com dados novos "a green selection frame and an Update button"; mini toolbar "Update Once"/"Update Automatically" por elemento; "thinkcell can automatically update linked elements whenever their linked Excel range changes" ("as long as thinkcell can connect to the linked Excel range"); Revert só das atualizações desde a abertura. **Data Links**: abre por Insert > thinkcell Data > Data Links, Tools > Data Links ou pela mini toolbar; miniaturas dos slides à esquerda, fontes agrupadas por pasta/planilha à direita; ícones de estado: verde = atualização pendente, laranja "?" = "Not yet loaded" (fonte não carregada nesta sessão), vermelho "!" = "File cannot be opened"; filtros "Updates Available", "Updated", "Automatically Updating", "Tableau Data Links"; botões Update Once, Revert, Update Automatically, Manage Data Sources, Break Excel Link; Ctrl/Shift/Ctrl+A; duplo clique abre o intervalo no Excel; fontes alternativas (pasta renomeada/copiada) e "Manage Data Sources" para reapontar; link quebrado mantém os dados no datasheet. |
| https://www.thinkcell.com/en/resources/manual/exceldatalinks | Criar vínculo: Insert > thinkcell > Link To PowerPoint (gráficos, 3 tipos de tabela, Harvey ball, checkbox, imagem) e "To Existing Element"; texto por "Text Linked to Excel". Moldura: **verde** = alguma apresentação aberta usa o intervalo; **vermelha** = nenhuma apresentação aberta o usa; **azul** = selecionada; arrastar a moldura/alças muda o intervalo; Transpose Link no botão direito da moldura; cópias da pasta compartilham o id do vínculo. |
| https://www.thinkcell.com/en/resources/kb/0163 | Detecção de mudança: "Excel sends a notification to other programs when data in a cell range has changed"; copiar/colar na mesma pasta não notifica (F9 resolve). Base para inferir que o automático do think-cell reage a notificações do Excel com os dois arquivos abertos. |
| https://www.thinkcell.com/en/resources/manual/mini-toolbar, /context-menu, /think-cell-user-interface | Mini toolbar ao selecionar elemento/recurso (fonte, preenchimento, contorno, formato de número, marcadores); menu de contexto que "adapts to the elements or features that you select" para adicionar/remover/ajustar recursos. |
| https://www.thinkcell.com/en/resources/manual/think-cell-ribbon | Grupos da aba think-cell no PowerPoint (Structure, Elements, Data, Text, Layout, Color, Quality, Collaboration, More) e no Excel (Elements = Excel data links, Formulas = rounding/CAGR...). |
| https://www.thinkcell.com/en/resources/manual/chartdecorations | Anotações: Level Difference Arrow, Total Difference Arrow, "Add Compound Growth Arrow", series CAGR labels, value line (encaixa na média), 100% indicator, series connectors, universal connectors; tudo por botão direito + alças. |
| https://www.thinkcell.com/en/resources/manual/textlabels | Tipos de rótulo (categoria, série, segmento, ponto, total, grand total...), **Label Content** na mini toolbar, formato de número por exemplo digitado, prefixo/sufixo, **magnitude ×10⁶**, posicionamento automático, texto livre com campos. |
| https://www.thinkcell.com/en/resources/manual/axes | Add Tick Labels/Tick Marks/Gridlines, alças de escala, "Fit Scale To Data", "Add Break" (eixo de valores), category gap, "Set Logarithmic Scale", "Set Same Scale"/"Reset to Independent Scale". |
| https://www.thinkcell.com/en/resources/manual/waterfall | Convenção "e" (equals), build-down, conector por par de colunas, "Add Waterfall Connector", Segment Order, seta com "% of datasheet 100%". |
| https://www.thinkcell.com/en/resources/manual/style-files | **Style files** XML: `fillLst` (`solidFill`, **`pattFill`**, `noFill`, `separator`), `lnLst` (peso e `prstDash`), `fillSchemeLst`/`fillScheme` com esquemas por tipo (stacked, waterfall, Mekko...), `defRPr` (tamanho, negrito, cor do texto e `solidfillRefAlt` = cor para fundo escuro, fontes); "Tools > Change Style" (vale para elementos inseridos depois), "Current Style"; distribuição pelo modelo do PowerPoint. |
| https://www.thinkcell.com/en/resources/manual/format-and-style-elements, /format-and-style-charts | Menu Fill com "Use Datasheet Fill"; esquema de cores do gráfico; "Use Datasheet Fill on Top" (cor da célula do Excel vira cor do segmento); "Use Datasheet Number Format"; cores por hex/RGB e conta-gotas. |
| https://www.thinkcell.com/en/resources/manual/element-datasheets, /slide-workbooks, /introductioncharting | Datasheet por duplo clique ou Open Datasheet, abre ao inserir; Transpose Rows and Columns; Slide Workbook embutido no slide; inserir = Elements > tipo > clicar no slide. |
| https://www.thinkcell.com/en/resources/manual/gantt, /clean-up, /keyboardshortcuts | Gantt (escalas, milestones, brackets, shades, colunas Responsible/Remark, Fit Scale to Data); Clean Up (inclui "Sanitize Contents: Numbers"); atalhos (Ctrl+clique, duplo clique cicla setas/Harvey balls). |

**Não confirmado** (documentação atual não diz; registrado como suposição): (a) o atraso do "Update Automatically" e se ele exige o Excel e o
PowerPoint abertos ao mesmo tempo — a página diz "whenever their linked Excel range changes ... as long as thinkcell can connect to the linked
Excel range" e o KB0163 fala de notificação do Excel: assumimos **com os dois abertos, na notificação de mudança** (o que o ThinQQ já faz, com 2 s
de espera); (b) se ao abrir uma apresentação o think-cell abre sozinho as pastas fechadas — a página diz que as fontes ficam "Not yet loaded" até
o usuário as abrir/atualizar: assumimos **não abre** e não pergunta; (c) os nomes exatos das entradas do menu radial (só ícones na página) — usamos
os do guia do usuário; (d) o formato exato da bandeira (o guia mostra o ícone do Excel no canto superior esquerdo; o manual fala em moldura verde e
botão Update): tratamos os dois como "indicador de dados novos ao selecionar".

## 3. O think-cell de referência (o que copiar)

Fluxos do usuário final, resumidos (fonte: §2):
- **Inserir gráfico (PowerPoint)**: Inserir > Elements > tipo > clicar no slide; o gráfico nasce selecionado e o datasheet abre; digitar atualiza na hora.
- **Vincular do Excel**: selecionar o intervalo > Inserir > Link to PowerPoint > tipo > (no PowerPoint) clicar onde colocar; o intervalo ganha moldura verde
  (alças para mudar o intervalo; botão direito: Transpose Link); "To Existing Element" religa um elemento existente ao intervalo.
- **Dados mudaram**: nada muda sozinho (padrão). Ao selecionar o gráfico aparece o indicador verde e o botão Update; clicar atualiza. O diálogo Data
  Links lista tudo com estados (pendente, atualizado, automático, não carregado, quebrado) e atualiza em lote. "Update Automatically" é por
  elemento, ligado pelo usuário.
- **Mudar o gráfico**: clique esquerdo = mini toolbar (cores, fonte, formato de número, conteúdo do rótulo, tipo); clique direito = menu de contexto
  com os recursos (totais, rótulos de série/categoria, CAGR, diferença, linha de valor, corte, conectores, mesma escala, log...).
- **Estilo**: esquema de cores do gráfico (dropdown) e Style file da organização (Change Style; vale para os próximos).
- **Quebrar vínculo**: botão direito > Break Excel Link (os dados ficam no datasheet).

## 4. Estado atual e por que a UX é ruim (evidências)

| # | Problema | Evidência | Think-cell |
|---|---|---|---|
| G1 | **Atualização automática a cada mudança, ligada por padrão**: o motor marca a pasta em `SheetChange`/`SheetCalculate`/`AfterSave` e roda 2 s depois da última mudança; a opção global nasce ligada | `src/core/modules/modPptAuto.bas:11` (`PPTAUTO_DEBOUNCE_S = 2`), `:58` (padrão `"1"`), `:189-206` (`PptAuto_OnChange`) | Manual por padrão; automático só por elemento |
| G2 | **Gráficos criados pelo PowerPoint nascem automáticos** (a experiência mais visível do usuário: digitar no Excel faz o slide "piscar") | `modCoreApi.bas:1703` (`PptSync_SetAuto shp, True`); CHANGELOG 1.8.0 "o gráfico já nasce automático" | Datasheet atualiza ao editar, mas o elemento vinculado ao Excel é manual |
| G3 | O ciclo automático **limpa o Desfazer do Excel** ~2 s depois de digitar (o refresh redimensiona o gráfico do Excel) | CHANGELOG [1.10.0] "Alterado": "O Desfazer do Excel pode ser limpo cerca de 2 segundos depois de digitar"; `modPptSync.bas:372` (`co.Width = shp.Width`) | Não mexe no Excel |
| G4 | **Nenhum indicador de "dados novos"**: o usuário não sabe se o slide está atrasado; só há "Atualizar uma vez" | `src/ppt/ribbon/customUI14.xml:54-61` (grupo sem estado; o único estado é o toggle automático) | Bandeira verde/moldura verde + Update |
| G5 | **Pergunta modal ao abrir** a apresentação ("Atualizar?") com lista de pastas | `modPptTqq.bas:731-757` (`PptTqq_HandleOpen` + `PptTqq_Ask`), padrão `"ask"` em `:199` | Não pergunta; mostra estado |
| G6 | **Caixa de mensagem em todo sucesso** ("Gráficos do ThinQQ atualizados: 1.") | `modPptTqq.bas:957` (`PptTqq_Msg msg` sempre), `:675` | Silêncio; o indicador some |
| G7 | **Sem visão geral dos vínculos**: não há lista de elementos, estados, origens, nem ação em lote | (ausente) | Diálogo Data Links |
| G8 | **Comandos longe do elemento**: a aba ThinQQ do PowerPoint tem 4 grupos fixos; selecionar o gráfico abre a aba "Design do Gráfico" do PowerPoint, sem nada do ThinQQ; nenhum botão direito | `src/ppt/ribbon/customUI14.xml:7-94` (só `<tabs>`, sem `contextualTabs` nem `contextMenus`); `src/loader/ribbon/customUI14.xml` idem | Mini toolbar e menu de contexto no próprio elemento |
| G9 | **Excel: o caminho "vincular ao PowerPoint" está espalhado e com estado escondido**: caixa de sessão "Inserir também no slide" dentro de um menu (muda o que os 21 botões de inserir fazem), "Copiar como imagem", "Copiar como gráfico vinculado", "Vincular célula ao texto selecionado", "Inserir tabela vinculada no slide", "Atualizar gráficos no PowerPoint", duas caixas de opção do automático e "Mostrar/ocultar molduras", tudo no menu "PowerPoint" (12 itens) | `src/loader/ribbon/customUI14.xml:187-209`; estado de sessão em `modRibbon.bas:9,147-151` | Um botão "Link to PowerPoint" com galeria de tipos |
| G10 | **Escolher categorias por número em InputBox** (CAGR/diferença pelo PowerPoint); Alterar dados, Corte de eixo e Título ligado por InputBox | `modPptTqq.bas:1693,1697`; `modCoreApi.bas:1069,846,1104` | Seleção direta + alças |
| G11 | **Alterar dados recria o gráfico com id novo: o slide deixa de acompanhar** | `modCoreApi.bas:1178-1180` + CHANGELOG 1.7.0 ("Alterar dados recria o gráfico: a moldura sai, o slide deixa de acompanhá-lo") | Arrastar a moldura muda o intervalo sem perder o vínculo |
| G12 | **Moldura sem estado** (sempre verde, mesmo sem apresentação aberta usando o intervalo) | `modLinkFrame.bas:24-54` (cor fixa) | Verde/vermelha/azul |
| G13 | **Sem trocar o tipo do gráfico** depois de inserido | (ausente; só recriar) | Ícones de tipo na mini toolbar |
| G14 | **Paleta fixa de 3 opções; cor manual se perde** em Alterar dados, troca de paleta e ligar conectores; sem hachura; sem estilo salvo | `src/loader/ribbon/customUI14.xml:95-100`; `modCoreApi.bas:408-412` (só `padrao/cinza/tema`); `modChartWaterfall.bas:306-342` (cores reescritas); `modAnnotations.bas:778-791` (conectores recriados em cinza) | Esquema de cores + Style files (com `pattFill`) |
| G15 | **Grupo "Gráfico do ThinQQ" do PowerPoint pobre** (8 controles; legenda só Direita/Nenhuma) e só para gráficos; texto/tabela vinculados fora de "Atualizar uma vez"/"Atualizar automaticamente" | `src/ppt/ribbon/customUI14.xml:67-93`; `modPptTqq.bas:855-878` (`PptTqq_SelectedTagged` só `TQQ_CHART`) | Todos os elementos vinculados no mesmo fluxo |
| G16 | **Duplo clique no gráfico do slide abre o "Editar Dados" do PowerPoint** (planilha embutida errada para o ThinQQ) | CHANGELOG 1.8.0 (limites) | Duplo clique abre o datasheet |

## 5. O que o VBA + RibbonX permitem (viabilidade)

| Recurso | No Excel 2019 | No PowerPoint 2019 | Uso proposto |
|---|---|---|---|
| Aba contextual (`<contextualTabs><tabSet idMso="TabSetChartTools">` com aba própria) | Sim (RibbonX 2010+) | Sim | 1.11.0: aba "ThinQQ" ao selecionar gráfico no PowerPoint; 1.13.0 no Excel |
| `getVisible`/`getEnabled`/`getLabel`/`getImage` + `InvalidateControl` dirigidos pela seleção | Sim (`Watchers_NotifySelection` já invalida) | Sim (`App_WindowSelectionChange` já invalida, `clsPptEvents.cls`) | Bandeira por `getVisible` (dois botões, um com ícone de bandeira) |
| `IRibbonUI.ActivateTab` | Sim | Sim | Não usado por padrão (trocar de aba sozinho incomoda); opção futura |
| `contextMenus` (`ContextMenuCell`, `ContextMenuShape`, menus de gráfico) | Sim para célula; para gráfico **não confirmado** (sonda PX2 não valida sem janela) | **Não confirmado** quais `idMso` existem para o gráfico selecionado | 1.13.0, condicionado à sonda manual P-U5 |
| Eventos de gráfico `BeforeRightClick`/`BeforeDoubleClick` (`WithEvents Chart`) | **Sim** (o ThinQQ já usa `WithEvents Chart` em `clsChartWatcher.cls`) | Não existem para gráfico no slide | Excel 1.13.0: botão direito = menu ThinQQ (`CommandBars` popup, medido PX1) e duplo clique = selecionar a origem |
| `CommandBars.Add(msoBarPopup)` + `ShowPopup` | Sim (medido PX1) | Sim (objeto existe) | Menu do botão direito do Excel |
| UserForm **não modal** | Sim (frmAnnotation é modal, mas o recurso existe) | Sim; **o build do .ppam ainda não importa formulários** (`build.ps1:27-83`) | Diálogo Vínculos no PowerPoint (1.11.0); editor de estilos no Excel (1.12.0) |
| Painel de tarefas (task pane) | **Não** sem suplemento COM/VSTO (`CustomTaskPanes` só via `ICustomTaskPaneConsumer` de um COM add-in) | **Não** | Substituto: formulário não modal |
| Barra de ferramentas flutuante sobre o elemento | Não (sem coordenadas da seleção na tela de forma confiável; sem API de janela) | Não | Substituto: aba contextual |
| Barra de status | Sim (`Application.StatusBar`) | **Não existe** no modelo do PowerPoint | Mensagens curtas só no Excel |
| Interceptar duplo clique no gráfico do slide | — | Não há evento | Substituto: "Editar dados" na aba contextual |
| Backstage | Sim | Sim | Não usado |
| Cor por ponto / hachura (`Fill.Patterned`) | Sim, sobrevive a mudar valor, `SetSourceData` maior, `Series.Values` e salvar/reabrir (PX1) | Sim (mesmo objeto de gráfico) | Estilos 1.12.0 |
| Depósito na pasta (`CustomXMLParts`) | Sim: 48 ms, sobrevive a salvar/reabrir (PX1) | — | Estilos da pasta |
| Diálogo de cor `xlDialogEditColor` | Existe (PX1), mas é **modal e muda a paleta da pasta ativa** | — | Rejeitado; seletor próprio com amostras + `#RRGGBB` |

## 6. Sondas

### 6.1 Medidas nesta rodada (Excel 2019 32 bits, instância própria e invisível criada por PowerShell externo; PIDs registrados e encerrados; arquivos em `C:\temp\tqq-ux-probe`, apagados ao fim)

| # | Medida | Resultado |
|---|---|---|
| PX1a | `Points(1).Format.Fill.Patterned(16)` (diagonal escura), `ForeColor` preto, `BackColor` branco | 204 ms a 1ª vez; `Type = 2` (patterned), `Pattern = 16` lidos de volta. |
| PX1b | Cor sólida num ponto + mudar valor + `SetSourceData` com uma linha a mais + `Series.Values` reatribuído | Hachura e cor do ponto **sobrevivem** a tudo (o Excel guarda formatação por índice de ponto). Recriar o gráfico (Alterar dados) perde: a camada de estilo precisa reaplicar. |
| PX1c | `CustomXMLParts.Add` com namespace `urn:thinqq:styles:1`, salvar `.xlsx`, fechar, reabrir | 48 ms; 1 parte no namespace depois de reabrir; hachura e cor dos pontos também reabriram. |
| PX1d | `Dialogs(223)` (`xlDialogEditColor`) | Acessível (não exibido: é modal). |
| PX1e | `CommandBars.Add(..., msoBarPopup, , True)` + controle | Criado e apagado; barras de contexto existentes no Excel: `Cell`, `Series`, `Plot Area`, `Chart`, `Format Data Series`, `Shapes`, `Inactive Chart`... |
| PX2 | Suplemento `.xlam` com `customUI14.xml` e `onLoad` (7 variantes: base, `ContextMenuCell`, `TabSetChartTools`, `ContextMenuChartArea`, `idMso` inexistente, atributo inválido, `gallery`/`dynamicMenu`/`splitButton`) aberto por `Workbooks.Open` no Excel invisível | **`onLoad` nunca chamado em nenhuma variante (0 vezes, inclusive a base válida)**: o Excel invisível de automação não carrega a faixa. Consequência: não há validação automática do `customUI` sem janela; `idMso` duvidosos só se confirmam numa sessão visível (P-U5). |

### 6.2 Bloqueadas nesta rodada
O PowerPoint do usuário estava aberto (POWERPNT, com uma apresentação do usuário, e uma instância `/Embedding` criada 3 s depois, do
mesmo usuário) e o Excel do usuário e um Excel `/Embedding` também: nenhuma instância do PowerPoint foi criada (seria a do
usuário). Memória livre ~580 MB. Sem LogonUI (sessão desbloqueada). As sondas de PowerPoint ficam para a tarefa de sonda P-U (plano, U1/U2 passos
de sonda e tarefa E0), a rodar pelo controlador com o PowerPoint do usuário fechado.

### 6.3 Sondas pendentes (rodar antes da tarefa indicada)

| # | Pergunta | Como (instância própria; padrão `p1.ps1`) | Decisão que depende |
|---|---|---|---|
| P-U1 (antes da U1) | `Shape.Tags.Add` marca a apresentação como alterada (`Saved = False`)? `pres.Saved = msoTrue` depois volta? Entra no Desfazer? | `pu1.ps1`: apresentação salva em temp, colar uma forma, `Saved`; `Tags.Add`; ler `Saved`; `Saved = -1`; ler | Já tratado no código: a U1 sempre preserva `Saved`. Se `Saved = -1` não "colar", a U1 registra e o README lista "o ThinQQ marca a apresentação como alterada ao detectar dados novos" (risco R2). O Desfazer só se vê à mão (item manual). |
| P-U2 (antes da U2) | Custo de ler `Tags` de 1 forma em processo (no `.ppam`) em `WindowSelectionChange` com 200 formas no slide | `.pptm` de teste com módulo cronometrado (padrão de `TestPptAddin`) | < 5 ms: a faixa lê a etiqueta a cada seleção (decidido). |
| P-U3 (antes da U2) | `TabSetChartTools` existe no PowerPoint 2019 e a aba própria dentro dele aparece ao selecionar um gráfico vinculado | **Manual** (sessão visível): instalar o `.ppam` de teste montado da branch na máquina de desenvolvimento com o PowerPoint do usuário fechado e com a permissão dele; ver se a aba ThinQQ continua (se o XML for recusado a aba ThinQQ inteira some) | Some ⇒ tirar o bloco `contextualTabs` (a aba fixa continua com o grupo "Elemento selecionado"). |
| P-U4 (antes da E0/X3) | `idMso` dos menus de contexto do gráfico no PowerPoint (`ContextMenuChartArea`, `ContextMenuChartPlotArea`, `ContextMenuShape`, `ContextMenuObjectsGroup`) e no Excel (`ContextMenuChartArea`...) | **Manual**: um `.ppam`/`.xlam` por `idMso` candidato, cada um com a aba "P" (se o `idMso` for inválido a aba some; PX2 mostrou que só a sessão visível dirá) | Só os confirmados entram na X3. |
| P-U5 (antes da X3) | Excel: `BeforeRightClick` com `Cancel = True` + `CommandBars("tqq_menu").ShowPopup` num gráfico do ThinQQ | Manual (precisa de mouse) | Funciona ⇒ menu ThinQQ no botão direito do Excel; senão só a aba contextual. |
| P-U6 (antes da X2) | O PowerPoint acha a origem pelo nome do `ChartObject` (`...xlsx!Planilha![tqq_x Gráfico 1]`)? Recriar o gráfico com **o mesmo nome e no mesmo lugar** mantém o vínculo do slide? | `pu6.ps1`: colar vinculado, apagar o `ChartObject`, criar outro com o mesmo `Name`, `LinkFormat.Update` com a pasta aberta | Mantém ⇒ X2 "Alterar intervalo sem perder o slide" reaproveita o id; não ⇒ X2 só muda o intervalo no bloco (sem recriar) quando o número de séries/categorias permitir, senão avisa como hoje. |
| P-U7 (manual, U6) | `WindowActivate` do PowerPoint ao voltar de outro aplicativo (P2 antigo) | Manual | Só para a alternativa "atualizar ao voltar ao PowerPoint" (§8.1.6), não escolhida. |

## 7. Arquiteturas de UX avaliadas

**A (recomendada) — Faixa contextual + Vínculos não modal + silêncio no sucesso (+ botão direito na 1.13.0).** Ao selecionar um gráfico do ThinQQ no
PowerPoint aparece, ao lado das abas "Design do Gráfico"/"Formatar" do PowerPoint, a aba contextual **ThinQQ** com o que o think-cell põe na mini
toolbar e no menu de contexto: Atualizar (com a bandeira quando há dados novos), Automático, Editar dados, Vínculos, rótulos/totais/conectores,
formato, estilo, setas, remover setas. O diálogo **Vínculos** (formulário não modal) é o Data Links. Sucesso não abre caixa. No Excel, a 1.13.0 traz
"Vincular ao PowerPoint" (uma galeria), a moldura com estado e o menu do botão direito no gráfico (evento `BeforeRightClick`, que o Excel oferece).
- Prós: só recursos medidos ou documentados do RibbonX/VBA; nenhum objeto estranho no slide; funciona com o modelo de seleção que o `.ppam` já tem.
- Contras: a "bandeira" aparece na faixa, não sobre o gráfico; a aba contextual exige um clique na aba (o PowerPoint abre "Design do Gráfico").

**B — Barra flutuante (UserForm não modal posicionado perto do gráfico) + forma "bandeira" no slide.** Um formulário pequeno seguiria a seleção e uma
forma verde "Excel ↻" seria posta no canto do gráfico com dados novos (clicar nela = atualizar).
- Prós: visualmente mais próximo do think-cell.
- Contras: sem API de janela não há como posicionar o formulário sobre o gráfico com segurança (zoom, rolagem, vários monitores); a forma bandeira é
  conteúdo do slide (entra no Desfazer, marca a apresentação como alterada, pode ser salva/impressa, aparece na coautoria); um clique nela muda a
  seleção e exige lógica frágil. **Rejeitada como padrão**; a forma bandeira fica como experimento opcional no backlog (risco alto, valor médio).

**C — Só a aba fixa (status quo melhorado).** Mais barata, mas não ataca G8 (comandos longe do elemento). Rejeitada.

## 8. Decisões

### 8.1 Modelo de atualização (1.11.0) — manual com bandeira, automático por elemento, como o think-cell

1. **Padrões novos**: todo vínculo novo nasce **manual** — colado, inserido no slide, criado pelo PowerPoint (G2: deixa de nascer automático), texto e
   tabela vinculados. A opção "Novos vínculos atualizam automaticamente" continua existindo (padrão desligada), mas sai da faixa do Excel e vai para
   Vínculos > Opções. "Atualizar automaticamente" continua **por elemento**, ligado pelo usuário (aba contextual ou Vínculos).
2. **Bandeira (dados novos)**: o motor do Excel (o mesmo `modPptAuto`, eventos + 2 s depois da última mudança) passa a fazer duas coisas por pasta
   que mudou: (a) atualiza os elementos **automáticos** (como hoje); (b) para os elementos **manuais** das apresentações abertas, **compara sem
   alterar nada** o que o slide mostra com o que a pasta daria agora (`TQQ_FP` do gráfico; `TQQ_TLAST` do texto; `TQQ_TBLAST` da tabela) e grava na
   forma a etiqueta `TQQ_STATE = "new"` (ou a apaga, se os dados voltaram). Essa comparação **não mexe no Excel**: não redimensiona o gráfico, não
   roda `Core_SafeRefresh` nem salva — o Desfazer do Excel fica (G3 resolvido para o caso manual, que passa a ser o padrão).
3. **Onde a bandeira aparece**: no `.ppam`, ao selecionar o elemento (evento `WindowSelectionChange`, leitura de etiqueta em processo, barata): na aba
   contextual e no grupo "Elemento selecionado" o botão **Atualizar** dá lugar ao botão **Dados novos: atualizar** com o ícone da bandeira verde
   (`tqq_flag`), por `getVisible`; no diálogo Vínculos a linha mostra "Dados novos". Sem bandeira: "Atualizar" normal (forçar continua possível).
4. **Índice de todos os vínculos**: etiqueta da apresentação `TQQ_LINKIDX` (mesmo formato do `TQQ_AUTOIDX`: `SlideID|nome|pasta`, `vbLf` entre
   entradas) com **todos** os elementos vinculados; mantida pelo core ao criar/colar/vincular e reconstruída na sincronização completa; o `.ppam` a
   reconstrói ao abrir a apresentação (varredura em processo). O `TQQ_AUTOIDX` continua (compatibilidade com o `.ppam`/core anteriores).
5. **Preservar `Saved`**: gravar `TQQ_STATE`/`TQQ_LINKIDX` não deve deixar a apresentação "alterada" (o estado é da sessão): quem grava lê
   `pres.Saved` antes e o devolve depois (P-U1 confirma). Ao abrir, o `.ppam` apaga `TQQ_STATE` de todas as formas (o arquivo pode ter sido salvo com
   uma bandeira velha) e, se o Excel já estiver aberto com o core, pede os estados (`Core_PptLinkStates`, API 6) **sem abrir pasta fechada**.
6. **Ao abrir**: o padrão passa de "Perguntar" para **"Não perguntar"** (como o think-cell). A escolha continua em Vínculos > Opções (Perguntar /
   Atualizar sempre / Não perguntar); quem já gravou uma escolha no registro a mantém. Alternativa avaliada e não escolhida: "atualizar ao voltar ao
   PowerPoint" (depende de P-U7 e repete a sensação de mudança sozinha).
7. **Silêncio no sucesso**: Atualizar (seleção), Automático e Atualizar tudo **não abrem caixa** quando tudo deu certo; problemas (pasta não
   encontrada, gráfico apagado, Excel ocupado, core antigo) continuam em caixa, com o texto de hoje. No diálogo Vínculos o resultado vai para a linha
   de status do próprio diálogo.
8. **Chave mestra**: "Atualização automática ligada" (registro `AutoUpdate`, padrão ligado) passa a se chamar **"Acompanhar mudanças do Excel
   (bandeiras e automáticos)"** e fica em Vínculos > Opções; desligada, o motor não marca nem atualiza nada.
9. **Migração**: decks da 1.6.0–1.10.x mantêm a marca `TQQ_AUTO` de cada forma (inclusive os gráficos criados pelo PowerPoint, que nasciam
   automáticos): respeitamos a escolha gravada no elemento, como o think-cell. Para voltar tudo ao manual: Vínculos > Selecionar tudo > Manual
   (documentado nas notas da versão e no tutorial). Nada é alterado em silêncio nos arquivos do usuário.
10. **Segurança mantida**: estados de pastas fechadas só com "Verificar" do diálogo (abre somente leitura, macros desligadas, `PptSync_OpenBook`);
    nunca `LinkFormat.Update` sem conferir o arquivo; nunca sondar caminho remoto ao abrir; o ThinQQ nunca salva pasta do usuário.

### 8.2 Vínculos (Data Links) — 1.11.0

Formulário **não modal** `frmTqqLinks` no `.ppam` (o build do `.ppam` passa a importar `src/ppt/forms`). Abre por: aba ThinQQ > grupo "Dados do
Excel" > **Vínculos** (botão grande) e aba contextual > Vínculos.
- Lista (ListBox de 5 colunas, seleção múltipla estendida): **Slide** | **Elemento** (Gráfico: tipo / Texto / Tabela + nome) | **Modo** (Manual /
  Automático) | **Estado** (Dados novos / Atualizado / Não verificado / Pasta não encontrada / Não encontrado na pasta / Sem vínculo) | **Origem**
  (arquivo).
- Filtro: Todos / Dados novos / Atualizados / Automáticos / Com problema. Botões: **Atualizar** (selecionados; forçado), **Atualizar com dados
  novos** (todos os "Dados novos"), **Automático**, **Manual**, **Ir para** (seleciona a forma no slide), **Abrir no Excel** (Editar dados),
  **Verificar** (estados de todos, abrindo pastas fechadas somente leitura), **Desvincular** (o ThinQQ deixa de atualizar: apaga as etiquetas `TQQ_*`
  da forma; o vínculo nativo do PowerPoint não é quebrado — S10), **Selecionar tudo**, **Fechar**. Opções: Acompanhar mudanças do Excel; Novos
  vínculos nascem automáticos; Ao abrir (3 opções).
- Linha de status com o resumo da última ação (sem caixa).
- Fora (com substituto): Revert (use Ctrl+Z do PowerPoint), miniaturas de slides, fontes alternativas automáticas, "Manage Data Sources" (1.13.0
  avalia "Trocar pasta de origem..." pela gravação de `SourceFullName`, medida na PG4).

### 8.3 PowerPoint — faixa (1.11.0)

- Aba **ThinQQ** (fixa): grupos "Inserir gráfico" (como hoje), **"Dados do Excel"** (era "Gráficos do Excel": **Vínculos** grande, **Atualizar tudo**,
  Empacotar para enviar, Sobre; sai o dropdown "Ao abrir"), **"Elemento selecionado"** (Atualizar ⇄ Dados novos: atualizar, Atualizar
  automaticamente, Editar dados, Texto vinculado ao Excel), "Gráfico do ThinQQ" (como hoje).
- **Aba contextual** "ThinQQ" dentro de `TabSetChartTools` (aparece com qualquer gráfico selecionado; os controles ficam desabilitados se não for do
  ThinQQ): grupos "Vínculo" (Atualizar/Dados novos, Automático, Editar dados, Vínculos) e "Gráfico" (Rótulos, Totais, Conectores, Formato, Legenda,
  Seta CAGR, Seta diferença, Remover setas). Mesmos callbacks, ids próprios (`c` + id).
- "Atualizar" e "Automático" passam a valer para **texto e tabela vinculados** também (G15).

### 8.4 Excel — pontos de entrada (1.13.0)

- **Vincular ao PowerPoint**: botão grande (galeria em menu: os 21 tipos por família, "Tabela vinculada", "Texto: célula ao texto selecionado") no
  1º grupo da aba ThinQQ **e** num grupo "ThinQQ" da guia **Inserir** do Excel (como o think-cell). Substitui a caixa de sessão "Inserir também no
  slide" (estado escondido) e os itens "Copiar como gráfico vinculado/imagem" vão para "Mais" (continuam existindo).
- **Moldura com estado**: verde = alguma apresentação aberta usa o gráfico do intervalo; vermelha = nenhuma (conferida pelo motor quando a pasta muda
  e por Atualizar); a moldura continua não impressa e bloqueada.
- **Alterar intervalo sem perder o slide** (G11), condicionada à P-U6.
- **Botão direito** no gráfico do ThinQQ no Excel = menu ThinQQ (P-U5); **duplo clique** = seleciona a tabela de origem (como o datasheet).
- **Alterar tipo** do gráfico (G13): menu "Tipo" na faixa e no botão direito do Excel, recriando com o mesmo id (mesma mecânica da X2). **No
  PowerPoint o tipo não muda** (o gráfico do slide é uma cópia: `LinkFormat.Update` traz dados, não o tipo): o ThinQQ avisa para colar de novo.
- PowerPoint: menus de contexto com os `idMso` confirmados na P-U4.

### 8.5 Estilos de gráfico (1.12.0) — pedido do usuário

**O que é**: um **estilo** é um conjunto nomeado de valores por **papel** (não por série física), validado e salvo, como o Style file do think-cell
mas editável numa janela. Formato de texto `chave=valor` (UTF-8), uma linha por chave, arquivo `.tqqestilo`:

```
ThinQQ-Estilo=1
nome=Cascata MC
fonte=Calibri
fonte.tamanho=9
serie1=#1F3864
...
aumento=#70AD47
reducao=#E8173C
total=#D9D9D9
primeira=#FFFFFF
primeira.padrao=diagonal_larga_para_cima
primeira.padrao.cor=#000000
primeira.borda=#000000
rotulo.cor=auto
rotulo.negrito=1
conector.cor=#7F7F7F
conector.traco=pontos
conector.espessura=0.75
seta.cor=#404040
seta.rotulo.fundo=#FFFFFF
seta.rotulo.borda=#1F3864
eixo.categoria.linha=1
espaco=50
```

**Chaves** (todas opcionais; a ausente = comportamento atual do ThinQQ): `nome` (1–40 caracteres, sem `; = | " \ / : * ? < >` nem controles);
`fonte` (nome de fonte, `[A-Za-z0-9 -]`, até 40); `fonte.tamanho` (6–24); `serie1`..`serie8` (`#RRGGBB`); `total`, `aumento`, `reducao`,
`primeira`, `ultima` (`#RRGGBB`); `primeira.padrao`/`ultima.padrao` (lista fechada de 12 hachuras mapeadas para `msoPattern*`: `nenhum`,
`diagonal_clara_para_cima`, `diagonal_escura_para_cima`, `diagonal_larga_para_cima`, `diagonal_clara_para_baixo`, `diagonal_escura_para_baixo`,
`diagonal_larga_para_baixo`, `horizontal_escura`, `vertical_escura`, `xadrez_pequeno`, `pontilhado_50`, `grade_pequena`);
`primeira.padrao.cor`/`ultima.padrao.cor` (`#RRGGBB`); `primeira.borda`/`ultima.borda` (`#RRGGBB` ou `nenhuma`); `rotulo.cor` (`auto` = contraste
branco/cinza-escuro como hoje, ou `#RRGGBB`); `rotulo.negrito` (0/1); `rotulo.tamanho` (6–24); `conector.cor`; `conector.traco` (`solido`,
`pontos`, `tracos`); `conector.espessura` (0,25–3, ponto ou vírgula); `seta.cor`; `seta.rotulo.fundo`; `seta.rotulo.borda` (`#RRGGBB` ou
`nenhuma`); `eixo.categoria.linha` (0/1); `espaco` (0–500). Linhas desconhecidas são ignoradas e contadas; valor inválido = a chave é ignorada e
contada (o resto do estilo vale). Arquivo maior que 16 KB ou com mais de 200 linhas é recusado inteiro. Nada é executado (só `Split`/`Like`).

**Papéis por família** (`Style_ApplyTo`):
- Cascata de 1 série: `aumento` → `tqq_uppos`/`tqq_upneg`; `reducao` → `tqq_downpos`/`tqq_downneg`; `total` → `tqq_total`; `primeira`/`ultima` →
  o ponto 1/n de todas essas séries (a barra inicial do usuário, "2026", é um aumento a partir de zero no ThinQQ, `modChartWaterfall.bas:125-145`;
  por isso o papel é "primeira barra", não "primeiro total").
- Cascata de várias séries, colunas/barras/áreas empilhadas, agrupadas, 100%, combinação, Mekko, borboleta, cascata comparativa: `serieK` pela ordem
  das séries de dados (auxiliares `tqq_*` nunca); cascata de várias séries: `total` nos totais.
- Pizza/rosca: `serieK` por fatia. Dispersão/bolhas: `serieK` por grupo. **Gantt (1.12.0): só fonte, rótulos e espaço** (as séries do Gantt
  são todas auxiliares `tqq_*`; cores por papel do Gantt ficam no backlog).
- **Rótulo dentro da barra da cascata** (os "0,6%" brancos em negrito dentro das barras verdes do exemplo): fora da 1.12.0 — os rótulos da cascata
  ficam onde o ThinQQ os põe (acima/abaixo, série auxiliar `tqq_lbl`); o estilo muda cor/negrito/tamanho. Backlog: chave `rotulo.posicao`.
- Comuns: `fonte`/`fonte.tamanho` (área do gráfico), `rotulo.*` (rótulos das séries com rótulo, inclusive `tqq_lbl*`), `conector.*` (formas
  `tqq_ann_conn_*` da cascata e `SeriesLines` das empilhadas), `seta.*` (formas `tqq_ann_cagr_*`, `tqq_ann_diff_*`, `tqq_ann_delta*`: linha e caixa do
  rótulo), `eixo.categoria.linha`, `espaco` (o mesmo `ApplyOption "gap"`).
- `Kinds_Supports(kind, "style")` = todos os tipos; dentro de cada família, chave sem papel é ignorada sem erro.

**Onde fica**:
- **Na pasta** (para quem recebe a pasta ver o mesmo visual): `CustomXMLParts` no namespace `urn:thinqq:styles:1`, um elemento `<style>` por estilo
  com o texto `chave=valor` (medido PX1c). O gráfico guarda só o nome (`style=<nome>`) e uma versão curta (`stylever=<8 hex do SHA-256 do texto>`) nas
  opções do bloco: a versão muda a impressão digital, então o PowerPoint percebe a troca de estilo (bandeira/automático).
- **Do usuário**: `%LOCALAPPDATA%\ThinQQ\styles\<nome>.tqqestilo` (o build Test nunca usa essa pasta: gancho `Style_SetStoreDirForTest`).
- **Embutidos** (só leitura, duplicáveis): "Padrão ThinQQ", "Tons de cinza", "Cores do tema" (as três paletas de hoje, que viram estilos), **"Cascata
  MC"** (o exemplo do usuário: primeira barra branca com hachura diagonal preta e borda preta, aumentos verdes ~RGB 112,173,71 com rótulo branco
  em negrito, totais cinza ~217,217,217, redução vermelha ~232,23,60, conectores pontilhados, seta com rótulo em elipse branca de borda azul-escura,
  sem eixo de valores nem grade) e **"Cascata DDE"** (aumentos laranja, redução verde, totais azuis, primeira barra hachurada — o visual anterior do
  usuário).
- **Padrão para novos gráficos**: registro `HKCU\...\ThinQQ\Estilos\Padrao` (build Test: memória); `Core_Insert` aplica depois de criar.
- **Importar/Exportar**: `GetOpenFilename`/`GetSaveAsFilename` com o filtro `Estilo do ThinQQ (*.tqqestilo)`; importar valida tudo e grava no depósito
  do usuário (nome repetido pergunta se substitui).

**Sobrevivência** (G14): o estilo é uma camada **reaplicada pelo core** depois de tudo que recolore: `Core_Insert`, `Core_SetToggle`,
`Core_SetOption` (paleta escolhida = estilo embutido; outras opções reaplicam o estilo), `Core_RebuildFromRange` (Alterar dados) e
`Core_SafeRefresh` quando o "carimbo" (`stylestamp`: nome + versão + nº de séries + nº de categorias) mudou (linha nova na tabela muda a posição da
"última barra"). Conectores recriados (`Ann_SetWaterfallConnectors`) e setas novas recebem o estilo na reaplicação.

**PowerPoint**: o gráfico vinculado é uma cópia com dados vinculados; `LinkFormat.Update` traz só os dados (a formatação do Excel não acompanha). A
1.12.0 acrescenta à sincronização (`PptSync_SyncShape`) a cópia da formatação das séries de dados e dos pontos 1 e n (preenchimento sólido/hachura,
borda, cor/negrito/tamanho dos rótulos, fonte do gráfico) do gráfico do Excel para o do slide, só quando a opção `style` existe. Na aba contextual do
PowerPoint entra o menu dinâmico **Estilo** (lista do core, API 7) que aplica pelo mesmo `Core_PptCommand`.

**Interface no Excel**: o dropdown "Paleta" vira **Estilo** (itens dinâmicos: os embutidos, os do usuário e os da pasta) + botão **Estilos...**
(janela `frmStyles`: lista, **Novo a partir do gráfico selecionado** (lê as cores atuais das séries, inclusive as pintadas à mão), Editar,
Duplicar, Excluir, Importar, Exportar, Aplicar ao gráfico / à planilha / à pasta, Definir como padrão). Editor com uma linha por papel: amostra
(Label com `BackColor`), caixa `#RRGGBB` e um seletor próprio de 40 amostras (as cores do Office), hachura num dropdown. `xlDialogEditColor`
rejeitado (modal e altera a paleta da pasta ativa). **Entregue (1.12.0, S4):** o seletor é uma lista de 20 cores nomeadas mais a caixa `#RRGGBB` (não as 40 amostras).

### 8.6 Balão de comentário

O exemplo do usuário tem um balão arredondado ("Redução Promoshare e Melhoria nas Negociações") preso a uma barra. **Decisão**: recurso próprio no
backlog 1.14.0 ("Adicionar comentário (balão)": forma `msoShapeRoundedRectangularCallout` dentro do gráfico, ponta presa à barra/categoria
escolhida, reposicionada pelo `Ann_Refresh` como as setas, estilo pelo papel `seta.rotulo.*`). Até lá: inserir a forma do PowerPoint/Excel à mão.

## 9. Matriz de paridade (think-cell × ThinQQ)

Legenda: ✔ tem; ◐ parcial/substituto; ✘ não tem; **versão** = quando entra; "fora" = fora do escopo com motivo.

| Recurso do think-cell | ThinQQ 1.10.1 | Depois | Observação |
|---|---|---|---|
| Atualização manual por padrão + indicador de dados novos | ✘ (automático ligado, sem indicador) | ✔ **1.11.0** | Bandeira na faixa (não sobre o gráfico) |
| Update Automatically por elemento | ✔ (1.6.0) | ✔ | Semântica mantida; padrão manual |
| Data Links (lista, filtros, lote, auto, quebrar) | ✘ | ◐ **1.11.0** | Sem miniaturas e sem Revert (Ctrl+Z) |
| Não perguntar ao abrir | ✘ (pergunta) | ✔ **1.11.0** | Opção mantida |
| Mini toolbar ao selecionar | ✘ | ◐ **1.11.0** | Aba contextual |
| Menu de contexto (botão direito) | ✘ | ◐ **1.13.0** | Excel por evento; PowerPoint conforme P-U4 |
| Link to PowerPoint (galeria) no Excel | ◐ (caixa de sessão + menus) | ✔ **1.13.0** | Guia Inserir + aba ThinQQ |
| Moldura verde/vermelha/azul | ◐ (só verde) | ◐ **1.13.0** | Verde/vermelha; sem azul (seleção) |
| Arrastar a moldura para mudar o intervalo | ✘ | ◐ **1.13.0** | "Alterar intervalo..." sem perder o slide (P-U6) |
| Transpose Link | ✘ | backlog | Exige transpor a leitura da tabela |
| To Existing Element | ✘ | backlog | Depende da X2 |
| Break Excel Link | ◐ (pela interface do PowerPoint) | ◐ **1.11.0** | "Desvincular": o ThinQQ para de atualizar; o vínculo nativo fica (S10) |
| Datasheet interno / Slide Workbook | ◐ (pasta de dados ao lado) | fora | Sem COM; a pasta de dados é o substituto |
| Duplo clique abre o datasheet | ✘ | ◐ **1.13.0** (Excel) | No slide não há evento: "Editar dados" |
| Trocar o tipo do gráfico | ✘ | ✔ **1.13.0** | Recria com o mesmo id |
| Esquema de cores do gráfico | ◐ (3 paletas) | ✔ **1.12.0** | Estilos |
| Style files (cores, hachuras, linhas, fontes) | ✘ | ✔ **1.12.0** | `.tqqestilo` + janela de edição |
| Use Datasheet Fill (cor da célula) | ✘ | backlog | Papel "cor da célula" no estilo |
| Totais, rótulos, conectores, CAGR, diferença, delta | ✔ | ✔ | — |
| Linha de valor | ✔ (média) | ✔ | Sem arrastar |
| 100% indicator | ✔ | ✔ | — |
| Corte de eixo | ✔ (valor) | ✔ | Sem arrastar |
| Mesma escala | ✔ | ✔ | — |
| Escala log, rótulos/linhas de grade por botão direito | ✘ | backlog 1.14.0 | — |
| Formato de rótulo por exemplo, magnitude ×10ⁿ, prefixo/sufixo, Label Content | ◐ (7 formatos fixos) | backlog 1.14.0 | — |
| Barras de erro, linha "net", série "Outros", ordenar segmentos | ✘ | backlog | Baixo valor para o usuário |
| Balão de comentário | ✘ | backlog 1.14.0 | §8.6 |
| Mekko %/unidades, pizza/rosca, dispersão/bolhas, área, combinação, borboleta, cascata comparativa | ✔ | ✔ | — |
| Gantt: calendário por duplo clique, colunas Responsible/Remark, sombras, linha de marco | ◐ (barras, marcos, colchete) | backlog | — |
| Harvey balls, caixas de seleção, imagem por URL | ✘ | backlog | — |
| Tabela vinculada / texto vinculado | ✔ (1.9/1.10) | ✔ (com bandeira na 1.11.0) | — |
| Alinhar/igualar tamanhos, Clean Up (sanitizar números), Enviar/Salvar slides, trocar fontes, separador decimal, agenda, biblioteca | ✘ | backlog 1.15.0 | Ferramentas do PowerPoint |
| Tableau | ✘ | fora | Exportar a view para planilha |
| Tarefa de painel (task pane) | — | fora | Sem COM |
| Alças de arrastar (largura das barras, escala) | ✘ | fora | Sem eventos de arrasto em gráfico |

## 10. Versões, ordem e branches

| Versão | Tema | Tarefas | Branch (criada pelo controlador) |
|---|---|---|---|
| **1.11.0** | Modelo de atualização + Vínculos + aba contextual (queixas principais) | U0–U6 | `feat/ux-vinculos` (de `main`) |
| **1.12.0** | Estilos de gráfico (+ hachura e papéis da cascata) | S1–S7 | `feat/estilos` (da ponta da 1.11.0) |
| **1.13.0** | Pontos de entrada como o think-cell (galeria, moldura com estado, alterar intervalo/tipo sem perder o slide, botão direito) | E0, X1–X6 | `feat/ux-entrada` (da ponta da 1.12.0) |
| 1.14.0+ | Backlog (balão, formatos de rótulo, eixo log, ferramentas) | nova rodada de planejamento | — |

Por quê: 1.11.0 ataca as duas queixas explícitas e cria a infraestrutura (estado, índice, formulários no `.ppam`, aba contextual) que 1.12.0 e
1.13.0 reaproveitam. Estilos vem em seguida por ser pedido explícito e independente do PowerPoint (só a S5 toca o `.ppam`). 1.13.0 depende de
sondas manuais (P-U4..P-U6) que podem ser feitas enquanto 1.12.0 é desenvolvida. Alternativa com pressa: publicar 1.11.0 e 1.12.0 juntas.

**API do core para o `.ppam`**: 6 = `Core_PptLinkStates` (U1); 7 = `Core_StyleNames` + comando `style` em `Core_PptCommand` (S5).

## 11. Riscos e mitigações

| # | Risco | Prob. | Impacto | Mitigação |
|---|---|---|---|---|
| R1 | `contextualTabs`/`idMso` inválido faz o PowerPoint/Excel **descartar a faixa inteira em silêncio** (PX2: não há como testar sem janela) | média | alto | Só `TabSetChartTools` na 1.11.0 (documentado pela Microsoft); Test-Ribbon com lista fechada de `idMso` permitidos; verificação manual P-U3 **antes do release**; o bloco é isolado (remover = 1 commit). |
| R2 | Gravar `TQQ_STATE` marca a apresentação como alterada / entra no Desfazer | média | médio | Preservar `Saved` (U1); P-U1; gravar só quando o estado muda; se o Desfazer for poluído, o README registra (verificação manual). |
| R3 | Usuários acostumados ao automático estranham o manual | média | médio | Decks antigos mantêm `TQQ_AUTO`; notas da versão e tutorial com "Como voltar ao automático"; bandeira visível. |
| R4 | PowerPoint ocupado (texto em edição) durante a marcação | alta | baixo | Mesmo caminho de ocupado do motor (novas tentativas 4–12 s). |
| R5 | Custo: comparar estados de muitos manuais a cada mudança | baixa | médio | Só elementos da pasta que mudou e só apresentações abertas; impressão digital medida 0,4–1,3 ms (P1); leitura por índice ~5 ms por forma (S4). |
| R6 | Formulário no `.ppam` (primeiro do projeto): build e teste novos | média | médio | U3 acrescenta `Add-FormFromJson` ao alvo ppt (mesma função do Excel) e um teste de legenda. |
| R7 | Estilo do usuário com valores hostis (arquivo recebido) | baixa | médio | Validação por chave com lista fechada; teto de tamanho/linhas; nenhum `Evaluate`/execução; nome sanitizado; testes de lixo. |
| R8 | Reaplicar estilo em todo `Core_SafeRefresh` deixa o Excel lento | média | médio | Carimbo `stylestamp`: só reaplica quando muda nome/versão/séries/categorias. |
| R9 | Formatação do slide sobrescrita na sincronização (usuário formatou no PowerPoint) | média | médio | Só gráficos com `style=` recebem a cópia de formatação; documentado ("o estilo manda"). |
| R10 | Suíte cresce (1.10: ~15–25 min; memória baixa ~0,6–1 GB) e o PowerPoint do usuário fica aberto (testes reais pulam) | alta | médio | Controlador agenda as suítes `-Ppt` com o PowerPoint fechado; nunca suíte concorrente; pedir ao usuário fechar Chrome/PowerPoint nas suítes finais. |
| R11 | Sessão bloqueada (LogonUI) quebra os testes de área de transferência | alta | baixo | Regra herdada: conferir LogonUI antes; reexecutar. |
| R12 | Tutorial passa do teto 655.360 B (hoje 607.118 B) | média | baixo | Cada tarefa de docs mede; condensar seções 1.6.0 (substituídas) antes de acrescentar. |
| R13 | Alterar intervalo/tipo "sem perder o slide" não funciona (P-U6 negativa) | média | médio | Caminho B descrito na X2 (reescrever o bloco quando a forma da tabela é compatível); senão mantém o aviso de hoje. |

## 12. Pré-flight (interfaces entre tarefas)

| Produz (tarefa) | Assinatura | Consome |
|---|---|---|
| U1 | `PPTSTATE_TAG = "TQQ_STATE"`, `PPTSTATE_TAG_LINKIDX = "TQQ_LINKIDX"`, `PPTSTATE_NEW/OK/UNKNOWN/MISSING/NOSOURCE/UNLINKED` | U2 (`.ppam` copia as constantes), U3 |
| U1 | `PptState_Of(shp As Object, wb As Workbook) As String` (só lê) | U1 motor, `Core_PptLinkStates` |
| U1 | `PptState_SetTag(shp As Object, isNew As Boolean) As Boolean`; `PptState_SavedOf(pres) As Boolean`; `PptState_RestoreSaved(pres, wasSaved As Boolean)` | U1, S5 |
| U1 | `PptLinkIdx_Add(shp As Object)`; `PptLinkIdx_Rebuild(pres As Object, items As Collection)` | caminhos de criação do core; `PptSync_Presentation` |
| U1 | `PptState_ManualFor(pres As Object, dirty As Object) As Collection`; `PptState_MarkItems(items As Collection) As Long` | `PptAuto_Tick` |
| U1 | `Core_PptLinkStates(pres As Object, entries As String, verifyClosed As Boolean) As String` → linhas `slideId|nome|pasta|tipo|modo|estado` (`vbLf`) | U2 (abertura), U3 (Verificar) |
| U1 | `CORE_PPT_API = 6` | U2 `PPT_API_STATES = 6` |
| U1 | `PptSync_OpenBook`, `PptSync_FindOpenBook`, `PptSync_CloseOpened` passam a `Public` (mesma assinatura) | U1 `modPptState` |
| U2 | `PptTqq_SelectedLinked() As Collection` (gráfico, texto e tabela); `PptTqq_StateOfSel() As String` ("new" se alguma forma selecionada tem `TQQ_STATE = "new"`); `PptTqq_SyncEntriesMsg(pres, entries, force, Optional fromAuto) As String` (sem caixa); `PptTqq_IsPlainSuccess(msg) As Boolean`; `PptTqq_PrepareOnOpen(pres)`; `PptTqq_LinkStates(pres, entries, verifyClosed) As String`; `PPT_DEFAULT_OPENMODE = "never"` | U3 |
| U2 | Callbacks `PptRx_GetUpdVisible`, `PptRx_GetNewVisible`, `PptRx_GetLinkedEnabled`, `PptRx_OnLinks`, `PptRx_OnUpdateAll`, `PptRx_GetCtxEnabled` | Test-Ribbon |
| U3 | `frmTqqLinks` (não modal); `PptLinks_Collect(pres) As Collection` (itens `Array(slideIndex, slideId, nome, tipo, modo, estado, pasta)`); `PptLinks_Filter(rows, filtro) As Collection`; `PptLinks_StateLabel(code) As String`; `PptLinks_KindLabel(kind, chartKind) As String`; `PptLinks_Show()` | U5 docs |
| S1 | `StyleDef_Parse(text As String, ByRef ignored As Long) As Object` (Dictionary validado); `StyleDef_Serialize(d) As String`; `StyleDef_Validate(key, value) As String` (valor normalizado ou "" se inválido); `StyleDef_Builtin(name) As String`; `StyleDef_BuiltinNames() As Variant`; `StyleDef_Version(text) As String`; `StyleDef_PatternOf(nome) As Long`; `StyleDef_ColorOf(hex) As Long`; `StyleDef_SafeName(s) As String` | S2–S5 |
| S2 | `StyleStore_Get(name, Optional wb, Optional preferUser) As String`; `StyleStore_Names(Optional wb) As String` (`|`); `StyleStore_SaveUser(name, text) As String`; `StyleStore_SaveBook(wb, name, text)`; `StyleStore_EnsureInBook`; `StyleStore_DeleteUser(name)`; `StyleStore_Import(path, ByRef name, ByRef ignored) As String`; `StyleStore_Export(name, path, Optional wb) As String`; `StyleStore_SetStoreDirForTest(dir)`; `StyleStore_DefaultName()`/`StyleStore_SetDefaultName(n)`; `Core_StyleNames`, `Core_StyleImport`, `Core_StyleExport`, `Core_StyleSetDefault`, `Core_StyleDefault` | S3, S4, S5 |
| S3 | `Style_ApplyTo(c As Chart, d As Object)`; `Style_ApplyNamed(c As Chart, name As String, Optional preferUser)`; `Style_Clear`; `Style_Reapply(c As Chart)`; `Style_ReapplyIfNeeded(c As Chart)`; `Style_Stamp`; `Style_FromChart(c, name) As String`; `Opt_ApplyPalette(c, pal)`; `Core_StyleApply(name, scope)`; `Core_GetStyleName() As String`; `Core_StyleFromChart(name) As String`; `Kinds_Supports(kind, "style")` | S4, S5 |
| S4 | `Core_StyleManage()` (`frmStyles` não modal); `StyleUi_*`; callbacks `Rx_StyleCount/Label/Id`, `Rx_GetStyleSel`, `Rx_OnStyle`, `Rx_OnStyles` | — |
| S5 | `CORE_PPT_API = 7`; `Core_PptCommand(..., "style", nome)`; `PptSync_ApplyFormat(pc, cht)`; `.ppam`: `PptTqq_StyleMenuXml`, `PptTqq_StyleNames`, `PptRx_GetStyleMenu`, `PptRx_OnStyle` | — |
| X1 | `Rx_OnLinkToPpt(control)`; `Rx_LinkProcFor(tag) As String` (tipo ⇒ `Core_InsertToPowerPoint`; `table`/`text`) | — |
| X2 | `RebuildKeepId(co, oldId, wb) As Boolean` (privada de `Core_RebuildFromRange`); `LinkFrame_SetState(c, used As Boolean)`; `Core_RefreshFrames(wb) As Long` | X4 |
| X3 | `clsChartWatcher.Cht_BeforeRightClick/BeforeDoubleClick` → `Menu_OnChartRightClick(c)`/`Menu_OnChartDoubleClick(c)`; `Menu_Items(c)`, `Menu_DoParam(param)` | X4 (submenu Tipo) |
| X4 | `Core_RebuildFromRange(c, rng, Optional newKind)`; `Core_ChangeKind(newKind)`; `Rx_OnChangeKind` | — |

## 13. Estimativa

Calibração (projeto 2026-10-15, `progress.md`): planejamento + sondas de uma parte ≈ 2,5–3 h; tarefa típica 10–45 min de implementador, revisão 3–20
min, 1–3 rodadas de correção; suíte completa 15–25 min (até 77 min sob pressão de memória); cada uma das 5 versões anteriores levou ~1,5 dia de
relógio (esperas de usuário, sessão bloqueada, suítes). Tarefas com PowerPoint real dependem do PowerPoint do usuário fechado.

### 13.1 Por tarefa (minutos de agente: otimista / provável / pessimista; inclui revisão e correções)

| Tarefa | Otim. | Prov. | Pess. | Notas |
|---|---|---|---|---|
| U0 promover | 5 | 10 | 15 | |
| U1 estado + motor + API 6 (+ P-U1) | 100 | 160 | 260 | most-capable; testes com PowerPoint real |
| U2 `.ppam`: seleção, bandeira, silêncio, abertura, aba contextual | 90 | 150 | 240 | most-capable |
| U3 Vínculos (form + build de forms no `.ppam`) | 90 | 140 | 230 | most-capable |
| U4 faixa do Excel | 25 | 40 | 70 | standard |
| U5 docs 1.11.0 | 40 | 70 | 110 | standard; Test-Docs |
| U6 suíte + revisão final + onda | 60 | 110 | 200 | |
| **1.11.0** | **410** | **680** | **1125** | ≈ 7–19 h de agente; 1,5–2,5 dias de relógio |
| S1 modelo puro | 50 | 80 | 130 | |
| S2 depósitos + importar/exportar + API de nomes | 60 | 100 | 170 | |
| S3 aplicar por papel + sobrevivência + ler do gráfico | 110 | 180 | 290 | most-capable |
| S4 janela + faixa Excel | 80 | 140 | 230 | |
| S5 PowerPoint (formato + comando, API 7) | 70 | 120 | 200 | |
| S6 docs | 40 | 70 | 110 | galeria de estilos no tutorial |
| S7 suíte + revisão | 60 | 110 | 200 | |
| **1.12.0** | **470** | **800** | **1330** | ≈ 8–22 h; 2–3 dias |
| E0 sondas manuais (controlador + usuário) | 20 | 40 | 60 | + 30–45 min do usuário |
| X1 galeria Vincular ao PowerPoint | 60 | 100 | 160 | |
| X2 moldura com estado + alterar intervalo sem perder o slide | 100 | 170 | 280 | depende de P-U6 |
| X3 botão direito Excel (+ PowerPoint se P-U4) | 70 | 120 | 200 | |
| X4 alterar tipo | 50 | 90 | 150 | |
| X5 docs | 40 | 70 | 110 | |
| X6 suíte + revisão | 60 | 110 | 200 | |
| **1.13.0** | **400** | **700** | **1160** | ≈ 7–19 h; 2–3 dias |
| **Total** | **1280** | **2180** | **3615** | ≈ 21 / 36 / 60 h de agente; **6–9 dias de relógio** |

### 13.2 Tempo humano (verificação)
1.11.0: 45–75 min (bandeira, Vínculos, aba contextual P-U3, migração de deck antigo, Saved/Desfazer); 1.12.0: 40–60 min (criar o estilo do exemplo,
importar/exportar, Alterar dados, PowerPoint); 1.13.0: 30–45 min de sondas E0 + 40–60 min de verificação. Total ≈ 2,5–4 h do usuário.

### 13.3 Incertezas que mais pesam
P-U3/P-U4 (faixa recusada em silêncio), P-U6 (alterar sem perder o slide), PowerPoint do usuário aberto (testes reais pulados ⇒ suítes
reagendadas), memória da máquina (suítes lentas) e o EDR.

## 14. Verificação manual (lista por versão; o README continua a numeração do item 110)

**1.11.0**: (110) deck novo: colar vinculado, mudar B2, selecionar o gráfico no PowerPoint: aparece "Dados novos: atualizar" com a bandeira;
clicar: atualiza, sem caixa; (111) Desfazer do Excel intacto depois de 5 s com só gráficos manuais; (112) apresentação salva continua "salva" depois
da bandeira (título sem asterisco/sem pergunta ao fechar); (113) Vínculos: filtros, Atualizar com dados novos, Automático/Manual, Ir para, Abrir no
Excel, Verificar com a pasta fechada, Desvincular; (114) aba contextual ThinQQ aparece ao clicar num gráfico do ThinQQ (P-U3) e a aba fixa continua;
(115) deck da 1.10 com gráficos automáticos continua automático; Vínculos > Selecionar tudo > Manual; (116) abrir apresentação com o Excel fechado:
nenhuma pergunta; com o Excel aberto: bandeiras aparecem; (117) texto e tabela vinculados com bandeira; (118) core 1.10 + `.ppam` 1.11: Vínculos abre e
"Verificar" avisa para atualizar o ThinQQ.
**1.12.0**: (119) reproduzir o gráfico do exemplo só com o ThinQQ (estilo "Cascata MC"); (120) salvar estilo a partir de um gráfico pintado à mão,
aplicar em outro; (121) exportar, mandar para outro usuário, importar; (122) Alterar dados, ligar conectores e trocar de estilo mantêm o visual;
(123) pasta enviada a quem não tem o estilo mostra o visual (depósito da pasta); (124) PowerPoint: atualizar traz a cor/hachura; menu Estilo da aba
contextual; (125) arquivo `.tqqestilo` corrompido/hostil recusado com contagem.
**1.13.0**: (126) Vincular ao PowerPoint pela guia Inserir; (127) moldura verde/vermelha; (128) alterar intervalo sem perder o slide; (129) botão
direito no Excel; (130) alterar tipo; (131) menus de contexto do PowerPoint (se entraram).

## 15. Backlog (1.14.0+; nova rodada de planejamento)
Balão de comentário; formato de rótulo por exemplo + magnitude ×10ⁿ + prefixo/sufixo + Label Content; eixo log; rótulos/grade/título de eixo por
botão direito; Use Datasheet Fill; Transpose Link; To Existing Element; Trocar pasta de origem; série "Outros"/ordenar; barras de erro; Harvey
balls e caixas de seleção; Gantt (calendário, colunas, sombras); ferramentas (alinhar/igualar, limpar/sanitizar números, enviar/salvar slides,
trocar fontes, separador decimal, agenda); forma "bandeira" no slide (experimental, §7 B).
