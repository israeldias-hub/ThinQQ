# ThinQQ — Pesquisa e Design: PowerPoint atualizado ao reabrir (1.4.0) e delta entre barras clicando nas barras (1.5.0)

Data: 2026-10-14
Autor: Israel Dias (com apoio do Claude Code)
Status: adendo às especificações da Fase 1 (`docs/superpowers/specs/2026-10-10-thinqq-design.md`: arquitetura, distribuição, atualização),
Fase 2 (`...-fase2-design.md`: setas de CAGR/diferença, matriz), Fase 3 (`...-fase3-design.md`: §3.4 sondas de PowerPoint, §4.4 "Copiar para o
PowerPoint") e da cascata comparativa (`docs/superpowers/specs/2026-10-13-thinqq-cascata-comparativa-design.md`: matriz 21 × 23).
Rascunho em `.superpowers/sdd/2026-10-14-ppt-delta/design.md`; a Tarefa A0 do plano o copia (sem mudar nada) para
`docs/superpowers/specs/2026-10-14-thinqq-ppt-delta-design.md`, que passa a ser a fonte que o teste da matriz lê (§6, Parte B).
Plano: `.superpowers/sdd/2026-10-14-ppt-delta/plan.md` (a Tarefa A0 o copia para `docs/superpowers/plans/2026-10-14-thinqq-ppt-delta.md`).
**Emenda 1 (2026-10-14, pedido do usuário):** "Adicione nesse plano a atualização automática do plugin via Excel": o suplemento do PowerPoint é
atualizado automaticamente pelo atualizador do Excel (§3.4 sondas, §4.9 decisões, Tarefa A7 nova do plano; A2, A4, A5, A6, A8 e A9 ampliadas).

## 1. Objetivo e escopo

O usuário pediu duas melhorias, com estimativa de tempo e **sem iniciar o desenvolvimento**:

1. **Parte A — "A integração do gráfico com o PowerPoint: quando o dado do Excel é atualizado e o PowerPoint é reaberto, ele atualiza o
   gráfico"** (o comportamento do think-cell). Hoje o ThinQQ cola um gráfico do PowerPoint vinculado à pasta (`ChartObject.Copy` +
   `Shapes.Paste`, Fase 3 §4.4); o vínculo é **manual** e as anotações do ThinQQ (setas, rótulos de seta, colchetes, conectores da cascata,
   marcas de corte) são **formas** dentro do gráfico, que o vínculo nativo não recalcula.
2. **Parte B — "Inserir a observação do delta das barras em qualquer barra via front gráfico: o usuário insere o delta selecionando as
   barras"**. Hoje a seta de diferença é inserida por um diálogo "Da categoria / Até a categoria", sempre entre totais de categorias.

Entregas propostas: **Parte A no ThinQQ 1.4.0** (core, loader e um terceiro arquivo, o suplemento do PowerPoint `ThinQQ.ppam`) e
**Parte B no 1.5.0** (core e loader). As partes são independentes no código; a ordem recomendada (A primeiro) está em §8.1.

| Item do brief | Decisão | Resumo |
|---|---|---|
| A.1 Pesquisa think-cell | §2.1 | Link por elemento com "Update Automatically"; recarrega a fonte a cada sessão do Office; fonte ausente = elemento fica com os últimos dados, marcado; só com o think-cell instalado. |
| A.2 Sondas PowerPoint | §3.1 | Abrir por automação **não** atualiza nem com `AutoUpdate` automático; `LinkFormat.Update` atualiza séries e rótulos ligados (inclusive via `_ThinQQ` muito oculta), **não** as formas; com a pasta aberta lê os valores **não salvos** em 60–280 ms (fechada: abre um Excel `-Embedding`, ~3,7 s); **fonte inexistente trava o PowerPoint**; redimensionar o gráfico do slide **estica** as formas internas (defeito medido da colagem atual). |
| A.3 Arquitetura | **(b) suplemento companheiro `ThinQQ.ppam`** + motor de sincronização no core do Excel | Ao abrir a apresentação, o `.ppam` encontra os gráficos marcados (`Shape.Tags`), pergunta (padrão) e chama `Core_PptSyncPresentation` no Excel com o ThinQQ: dados pelo vínculo nativo e as formas recalculadas pelo ThinQQ no tamanho do slide, aplicadas **no lugar** (animação, ordem Z, nome e texto alternativo preservados). Alternativa: (a) só o botão do Excel "Atualizar gráficos no PowerPoint" (o motor sem o `.ppam`). |
| A.4 Distribuição | §4.6 | 6 arquivos no release (+ `ThinQQ.ppam`), instalador registra o suplemento do PowerPoint (`HKCU\...\PowerPoint\AddIns\ThinQQ`), local confiável do PowerPoint, assinatura do `.ppam`, manifesto com `pptAddinUrl`/`pptAddinSha256`. |
| A.4b Atualização automática do suplemento (emenda) | §4.9 | O atualizador do Excel (o mesmo da 1.3.0: verificação na abertura, `Transport` com proxy/TLS, SHA-256) atualiza o `ThinQQ.ppam` **só de quem já o tem**, junto com o core e na mesma versão (`pptAddinVersion`); baixa e confere antes de trocar o core; troca atômica por renomeação; arquivo preso pelo PowerPoint ⇒ `ThinQQ.ppam.new` pendente, aplicado pelo próprio Excel na próxima abertura e a cada 10 min enquanto ele estiver aberto; contrato de API `Core_PptApiVersion`; opção "Atualizar também o suplemento do PowerPoint" (padrão ligada). |
| A.5 Testes | §4.7 | Testes VBA com PowerPoint real **só** sem `POWERPNT` aberto (padrão existente), projeto de teste do PowerPoint `.pptm` com corredor próprio, teste PowerShell de ponta a ponta; abrir pela interface, aviso de segurança e Modo de Exibição Protegido ficam manuais. |
| B.1 Pesquisa think-cell | §2.2 | "Level difference arrow": selecionar dois segmentos (Ctrl+clique) e "Add Level Difference Arrow"; modos relativo/absoluto; liga a segmentos, pontos e linha de valor; acompanha os dados. |
| B.2 Sondas Excel | §3.2 | `Points(i).Select` dispara `Chart_Select(3, série, ponto)`; selecionar a série dá `Arg2 = -1`; `GetChartElement` funciona no Excel invisível, em pixels (pt × 4/3 a 100%); forma dentro do gráfico dá `ElementID = 14` e `TypeName(Selection) = "Line"`; 200 seleções com o manipulador em 0,62 s. |
| B.3 Modelo do delta | §5.2 | Referências "k:c" por família (série de dados × categoria; barra inteira nas cascatas; lado × etapa na comparativa); `AlternativeText` versionado `kind=delta;v=1;...`; texto valor/percentual/ambos; cotovelo acima das barras com faixas sem colisão. |
| B.4 UI | §7 | Menu Setas: "Delta clicando nas barras" (arma o modo) e "Delta: valor / % / ambos"; Seta diferença (diálogo) continua como alternativa de teclado. |
| B.5 Testes | §5.6 | `Points(i).Select` (eventos reais) e o manipulador de `MouseUp` chamado com coordenadas calculadas; cliques reais de mouse ficam manuais. |
| Integração | §8 | O delta é forma `tqq_ann_*`: vai ao PowerPoint pelo mesmo motor da Parte A (clonado se não existia no slide). |

## 2. Pesquisa (2026-10-14)

### 2.1 think-cell — links de dados do Excel (Parte A)
- **Atualização automática por elemento.** "Automatic updates are turned on per element from the mini toolbar, using **Update
  Automatically**. Once enabled, any changes in the linked Excel range will automatically appear in the linked element" — desde que o
  think-cell consiga se conectar à faixa. Fonte: https://www.think-cell.com/en/resources/manual/manage-data-links
- **Sessão do Office.** O think-cell "must reload data sources at the start of each new Office session"; até lá o elemento mostra "Not yet
  loaded" (interrogação laranja). Para recarregar: "Update Once"/"Update Automatically" no elemento, ou abrir a pasta (duplo clique no
  elemento ou pelo Explorer). Mesma fonte.
- **Fonte ausente/movida/renomeada.** "If thinkcell can't open the source, the link is broken and the element shows a red exclamation mark
  with the tooltip 'File cannot be opened'"; o elemento continua editável e "stores a copy of the most recently used data from the linked
  data source in the element's datasheet". Religar: "Manage Data Sources". Mesma fonte.
- **Sem mudança silenciosa de vínculo.** "thinkcell won't set or change any Excel links without your explicit instruction to do so" (o
  comportamento antigo de religar sozinho a outra pasta foi abandonado). Fonte:
  https://www.think-cell.com/en/blog/stay-in-control-when-managing-excel-links-with-think-cell
- **O que liga.** Gráficos, tabelas, bolas de Harvey, caixas de seleção, textos e imagens; não liga texto em tabelas do PowerPoint e
  SmartArt. Fonte: https://www.think-cell.com/en/resources/manual/exceldatalinks
- **Abrir o gráfico para editar.** "If the linked workbook is not open, think-cell will attempt to open it from the location it had when the
  link was last established ... If this fails, the internal datasheet will be opened instead" (resumo da busca no manual).
- **ThinQQ fará igual:** atualização ao abrir a apresentação (com opção perguntar/sempre/nunca, em vez de por elemento), fonte ausente =
  gráfico fica com os últimos dados e o aviso diz qual pasta faltou, nunca religa sozinho a outra pasta, precisa do ThinQQ instalado para
  recalcular as anotações. **Fica de fora:** o diálogo "Data Links" (lista de vínculos por slide com estado colorido), fontes alternativas,
  vínculo de tabelas/textos/imagens, atualização "ao vivo" com as duas aplicações abertas sem reabrir (o botão do Excel cobre esse caso).

### 2.2 think-cell — setas de diferença (Parte B)
- **Inserir clicando.** "Select the start and end segments together (Ctrl-click for multiple selections), right-click, and choose 'Add
  Level Difference Arrow.' In line or area charts, select data points instead." Também pelo botão do menu de contexto e depois arrastando as
  alças; as pontas ligam a "column segments, data points or to a value line". Fonte: https://www.think-cell.com/en/resources/manual/chartdecorations
- **Modos.** O botão da seta alterna: uma ponta = diferença **relativa**, duas pontas = **absoluta**, uma ponta invertida = relativa no outro
  sentido; "the difference ... is automatically calculated and updated whenever the underlying data changes"; seta e rótulo podem ser
  arrastados; uma lacuna de categoria ajuda a pôr a seta entre colunas. Mesma fonte.
- **CAGR.** Vários: "If you select segments in more than two categories, you will get a CAGR for each pair of successive selected
  categories." Mesma fonte.
- **ThinQQ fará igual:** origem e destino escolhidos clicando em qualquer par de barras/segmentos/pontos (inclusive de séries diferentes na
  mesma categoria), vários deltas no mesmo gráfico, texto valor/percentual/ambos alternável, recalculado com os dados. **Diferente:** o
  Excel não tem seleção múltipla de elementos de gráfico (Ctrl+clique não existe em gráfico; medido §3.2), então o fluxo é de **dois
  cliques** com o modo armado por um botão; a seta não é arrastável (o refresh a recoloca: limite documentado); o desenho é um **cotovelo**
  acima das barras (como as setas de variação da 1.3.0), não uma seta vertical ao lado.

### 2.3 Microsoft (PowerPoint, Excel)
- `LinkFormat.AutoUpdate` (`ppUpdateOptionManual` 1, `ppUpdateOptionAutomatic` 2), `LinkFormat.Update`, `Presentation.UpdateLinks`
  ("Updates linked OLE objects" — atualiza os vínculos automáticos): https://learn.microsoft.com/en-us/office/vba/api/powerpoint.linkformat.autoupdate ,
  https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.updatelinks
- Eventos do PowerPoint: `Application.AfterPresentationOpen` ("occurs after an existing presentation is opened") e
  `ProtectedViewWindowOpen` (janela do Modo de Exibição Protegido; os eventos de apresentação não valem para ela):
  https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Application.AfterPresentationOpen ,
  https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.protectedviewwindowopen
- `Shape.PickupAnimation`/`ApplyAnimation` (copiar animação): https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.applyanimation
- Instalar suplemento do PowerPoint pelo registro: `HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\<nome>` com `Path` (REG_SZ) e
  `AutoLoad` (DWORD) — fonte comunitária consolidada: https://pptfaq.com/FAQ00469_Add_-_install_an_addin_via_the_registry.htm ; `Auto_Open`
  roda quando o `.ppam` carrega.
- Excel: `Chart.Select` (ElementID, Arg1, Arg2), `Chart.MouseUp` (x, y em coordenadas do cliente), `Chart.GetChartElement` (as mesmas
  coordenadas): https://learn.microsoft.com/en-us/office/vba/api/excel.chart.select(even) ,
  https://learn.microsoft.com/en-us/office/vba/api/excel.chart.mouseup , https://learn.microsoft.com/en-us/office/vba/api/excel.chart.getchartelement
- Vínculos de gráfico do Excel no PowerPoint (comunidade): o PowerPoint pergunta "Atualizar vínculos" ao abrir uma apresentação com vínculos
  automáticos; há relatos de a opção "Automático" voltar a manual em arquivos do OneDrive/SharePoint:
  https://learn.microsoft.com/en-us/answers/questions/5104591/powerpoint-doesnt-update-linked-excel-charts

## 3. Sondas (Excel 2019 32 bits pt-BR e PowerPoint 16.0 32 bits; instâncias NOVAS e invisíveis criadas pela sonda, PIDs conferidos; nenhum Excel/PowerPoint do usuário aberto nem tocado; arquivos em `C:\temp\tqq-ppt-probe`, apagados ao final; scripts `a1.ps1`..`a5.ps1`, `b1.ps1`, `b2.ps1`)

### 3.1 PowerPoint (Parte A)
Gráfico de colunas com a série ligada à planilha **muito oculta** `_ThinQQ` (fórmulas para `Dados`), rótulo do ponto 1 ligado por
`DataLabel.Formula` a `_ThinQQ!C2` (`="v="&Dados!B2`), uma linha `tqq_ann_diff_1` (texto alternativo `kind=diff;from=1;to=3`) e uma
caixa de texto `tqq_ann_diff_1_lbl` ("+20") dentro do gráfico; pasta salva; colado 3 vezes (`ChartObject.Copy` + `Slide.Shapes.Paste`).

| Medida | Resultado |
|---|---|
| Forma colada | `Type = 3` (msoChart), `HasChart`, `ChartData.IsLinked = True`, `LinkFormat.SourceFullName = C:\temp\tqq-ppt-probe\src.xlsx` (o caminho puro), **`LinkFormat.AutoUpdate = 1` (manual) por padrão**; `Shape.Name` = nome do ChartObject (`tqq_probe1`); `Tags.Add` aceito e **persistido** no .pptx (relido depois de reabrir). |
| `AutoUpdate = 2` (automático) gravado e o .pptx reaberto | Persistiu (lido 2 depois de reabrir) — num arquivo local. |
| Dados mudados no Excel (B2 = 99), pasta salva e fechada; o .pptx reaberto **por automação** (`Presentations.Open`, sem janela, `DisplayAlerts` = nenhum) | **Nenhum** gráfico mudou ao abrir, nem o de `AutoUpdate = 2` (valores 10,20,30). O aviso "Atualizar vínculos" da interface não aparece por automação: comportamento ao abrir pela interface = verificação manual (§4.7). |
| `LinkFormat.Update` (pasta fechada) | 3,76 s; abre um `EXCEL.EXE -Embedding` (que não carrega suplementos); **séries e rótulo ligado atualizam** (99 e "v=99", inclusive passando pela `_ThinQQ` muito oculta); as **formas internas não mudam** (texto "+20" e posição antigos, embora no Excel o texto fosse "NEW" e a linha tivesse andado 30 pt). |
| `Presentation.UpdateLinks` | 3,5 s; atualiza **só** o gráfico com `AutoUpdate = 2`; o manual fica. |
| `LinkFormat.Update` com a pasta **aberta** num Excel (alteração **não salva** B3 = 55) | 279 ms (2.ª vez: 61 ms; pasta aberta somente leitura: 145–189 ms); lê os valores **não salvos** (99,55,30). |
| `LinkFormat.Update` com a pasta de origem **renomeada** (inexistente) | **Travou o PowerPoint** (> 5 min; diálogo modal invisível); o processo da sonda foi encerrado pelo PID. ⇒ conferir a existência do arquivo **antes** de qualquer `Update`. |
| Redimensionar o gráfico do slide (460 → 709,7 pt, `LockAspectRatio`, o que `Ppt_PlaceShape` faz hoje ao colar) | As formas internas **escalam** junto (54 → 83,3 pt; caixa 34 → 52,5 pt de largura) enquanto a área interna **não** escala na mesma proporção (InsideLeft fica 21,9; altura 212,9 → 364,9, razão 1,71 contra 1,54): **a colagem atual do 1.3.0 deixa setas e rótulos fora do lugar** sempre que o PowerPoint ajusta o tamanho. |
| Mesmo tamanho nos dois lados (709,7 × 432; 500 × 300; 900 × 500) depois do `Update` | Área interna **idêntica** (L 26,99, T 38,08, W 614,68/404,98/804,98, H 364,86/232,87/432,87); `ChartArea.Width` do Excel = tamanho do ChartObject − 5 e `Point.Left/Top` do PowerPoint = Excel + 4 (deslocamento constante); na colagem a 460 × 280 as formas internas ficaram nas mesmas coordenadas do Excel. ⇒ ChartObject do Excel com o tamanho da forma do slide reproduz o desenho; as formas são copiadas com o deslocamento da área interna (medido 0). |
| Editar as formas internas do gráfico do slide | Texto, `Left`, `Chart.Shapes.AddLine` e `Delete` funcionam. |
| Animação (Esmaecer) e `LinkFormat.Update` | Mantida (1 efeito, mesma forma). |
| Substituir o gráfico (colar de novo, `PickupAnimation`/`ApplyAnimation`, apagar o antigo) | A animação vai para o **fim** da sequência (antes,grafico,depois → antes,depois,grafico) e a ordem Z muda (1 → 3). ⇒ atualizar **no lugar**, nunca substituir. |

### 3.2 Excel — eventos de seleção do gráfico (Parte B)
Projeto VBA injetado numa pasta nova da instância da sonda; classe com `WithEvents Cht As Chart` num ChartObject (como `clsChartWatcher`).

| Medida | Resultado |
|---|---|
| `ChartObject.Activate` | `Activate`, depois `Select(2, 0, 0)` (área do gráfico). |
| `SeriesCollection(1).Select` | `Deactivate`, `Activate`, `Select(3, 1, -1)`: **série inteira**, `Arg2 = -1` (é o 1.º clique do Excel numa série). |
| `SeriesCollection(1).Points(2).Select` / `SeriesCollection(2).Points(3).Select` | `Select(3, 1, 2)` / `Select(3, 2, 3)`: série e ponto; `TypeName(Selection) = "Point"`. |
| `PlotArea.Select` | `Select(19, 0, 0)`. |
| `Range("A1").Select` | `Deactivate`; `ActiveChart` vira `Nothing`. |
| Forma dentro do gráfico `.Select` | `Select(14, 1, 0)` (xlShape, índice da forma); `TypeName(Selection) = "Line"`, `Selection.Name = "tqq_ann_delta_1"`. |
| `GetChartElement(x, y, ...)` sem ativar o gráfico (Excel invisível, zoom 100%) | Funciona; a primeira coluna (Left 65,3 pt, largura 40,2) é atingida de x = 90 a ~140 px: **px = pt × 4/3** (96 ppp); devolve `3/série/ponto` em barras, `19` área de plotagem, `21` eixo, `24` legenda, `12` entrada da legenda, `28` nada. Com o gráfico ativado: o mesmo. |
| Custo | 200 `Points(i).Select` com o manipulador registrando: 0,62 s (~3 ms por evento). |
| `Application.StatusBar` com o gráfico ativo | Gravado e lido. |
| Não medido (sem mouse/teclado no Excel invisível) | Se `MouseUp` chega no **1.º** clique que ativa um gráfico embutido; se `Application.OnKey "{ESC}"` dispara com o gráfico ativo; escala com 125%/150% de DPI. ⇒ fluxo robusto aos dois caminhos (§5.4) e verificação manual (§5.6). |

### 3.3 Ambiente (para o plano)
- `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1` (já usado pelo build); **a chave equivalente do PowerPoint não
  existe**: montar o `.ppam` por automação exige o usuário habilitar uma vez "Confiar no acesso ao modelo de objeto do projeto VBA" **no
  PowerPoint** (pré-requisito da Tarefa A4; o implementador nunca grava o registro).
- Memória livre ~1 GB: uma instância de cada aplicativo por vez; a sonda A5 deixou um `POWERPNT /AUTOMATION -Embedding` travado (fonte
  inexistente), encerrado pelo PID que ela criou.

### 3.4 Trava de arquivo e marca da Internet (emenda; sonda `p1.ps1`, instância própria, PowerPoint sem nenhum outro aberto; pasta `C:\temp\tqq-ppt-probe2`, apagada)
| Medida | Resultado |
|---|---|
| Apresentação (.pptm) aberta **para edição** (`Presentations.Open(f, ReadOnly:=0)`) | Ler/copiar o arquivo: OK. Sobrescrever, **renomear** e apagar: **falham** ("being used by another process"). Fechada: renomear OK. |
| A mesma aberta **somente leitura** (`ReadOnly:=-1`) | Nenhuma trava: sobrescrever, renomear e apagar funcionam (o PowerPoint leu tudo para a memória). |
| Suplemento `.ppam` **carregado** | **Não medido**: carregar um suplemento exige `AddIns.Add` (grava `HKCU\...\PowerPoint\AddIns`, proibido nas sondas). ⇒ o mecanismo não presume nada: a própria renomeação diz se o arquivo está preso (§4.9). Verificação manual na A9. |
| Download como o atualizador faz (`MSXML2.XMLHTTP` + `ADODB.Stream.SaveToFile`; o `Transport` grava os bytes ele mesmo) | Status 200; fluxos do arquivo: só `:$DATA` — **sem `Zone.Identifier`** (a marca da Internet só vem de navegador/anexo). O instalador já faz `Unblock-File` em tudo que copia do ZIP baixado. |
| Versão de um arquivo do Office sem abri-lo | O instalador já lê `docProps/core.xml` do `.xlam` (`Get-LoaderFileVersion`, comentário "ThinQQ loader x.y.z"); o mesmo vale para o `.ppam` (OOXML). No VBA do loader ler um zip não é simples ⇒ a versão instalada fica no registro do ThinQQ (`PpamVersion`), gravada pelo instalador e pelo atualizador; o carimbo `Comments = "ThinQQ ppt x.y.z (...)"` no `.ppam` é a prova independente (instalador e testes PowerShell). |

## 4. Decisões — Parte A (1.4.0)

Critérios das fases anteriores: (1) gráfico nativo e correto sem o ThinQQ; (2) nada de `Declare` nem referências extras; (3) Excel/PowerPoint
2016+ 32/64 bits; (4) testável com instâncias próprias e invisíveis; (5) nunca tocar Excel/PowerPoint/arquivos do usuário sem ação dele.

### 4.1 Arquiteturas avaliadas
| Candidata | O que faz | Medido/avaliado | Veredito |
|---|---|---|---|
| (a) Vínculo nativo automático + botão no Excel | Colar com `AutoUpdate = 2`; ao reabrir o PowerPoint pergunta "Atualizar vínculos" e atualiza séries e rótulos ligados; botão "Atualizar gráficos no PowerPoint" no Excel refaz as formas | Dados e rótulos ligados: sim (§3.1). Formas: **não** (ficam com os números antigos: o "+20" de uma seta de diferença continua "+20" com os dados novos). Comportamento do aviso na interface e relatos de "Automático" voltando a manual no OneDrive (§2.3). | Alternativa (custo baixo, cumpre só metade: os rótulos das setas mentiriam) |
| **(b) Suplemento companheiro `ThinQQ.ppam` + motor no core** | `AfterPresentationOpen`: acha os gráficos marcados (`Tags`), pergunta, chama `Core_PptSyncPresentation` num Excel com o ThinQQ; o core atualiza os dados (`LinkFormat.Update`, com a pasta já aberta: 60–280 ms), põe o gráfico do Excel no tamanho do slide, recalcula as anotações e as copia **no lugar** | Todas as peças medidas (§3.1): mesma área interna com o mesmo tamanho, formas internas editáveis, animação mantida no `Update`, leitura de pasta aberta e não salva. | **Escolhida** |
| (b') Re-colar o gráfico inteiro | Mais simples de codificar | Animação vai para o fim, ordem Z muda (§3.1); perde a formatação feita no PowerPoint | Rejeitada |
| (b'') "Manifesto de geometria" gravado no Excel e aplicado pelo `.ppam` sem o Excel | O ThinQQ grava as posições das formas | As posições dependem do tamanho do gráfico **no slide** (que o usuário muda) e das fontes; o manifesto ficaria errado a cada redimensionamento (§3.1: formas esticam) | Rejeitada |
| (c) Apresentação `.pptm`/modelo com macros | Macros na própria apresentação | Aviso de macros em toda abertura; cada apresentação precisaria do código; não cobre apresentações existentes | Rejeitada |
| (d) Suplemento COM/VSTO ou Office Add-in (JavaScript) | — | Sem .NET 3.5; o Office Add-in não acessa gráficos vinculados nem o Excel local; instalação e assinatura muito maiores | Rejeitada |

### 4.2 Motor de sincronização (core, `modPptSync.bas`, prefixo `PptSync_`)
- **Marcação na colagem.** `Core_CopyToPowerPoint "chart"` grava na forma colada as etiquetas `TQQ_CHART` (id do gráfico, 6 caracteres),
  `TQQ_BOOK` (caminho completo da pasta), `TQQ_SHEET` (planilha), `TQQ_VER` ("1") e, depois da primeira sincronização, `TQQ_FP` (impressão
  digital). A imagem não é marcada (não tem vínculo).
- **Sincronizar UM gráfico (`PptSync_SyncShape shp, co`).** (1) dados: `shp.LinkFormat.Update` se `Ppt_LinkState = 1`; (2) o ChartObject vai
  para `shp.Width × shp.Height` (pt), `PptSync_ForceLayout` (exporta para um PNG temporário e apaga: o mesmo truque do `ForceLayout` dos
  testes, porque o Excel invisível só refaz a geometria ao desenhar), `Core_SafeRefresh`, outro `ForceLayout`; (3) **foto** das formas
  `tqq_ann_*` (nome, caixa, giro horizontal/vertical, visível, texto); (4) aplicação no gráfico do slide: formas que existem pelo nome
  recebem caixa/giro/texto/visível (deslocamento = diferença das áreas internas, medido 0); as que faltam são **clonadas** (linha, caixa de
  texto ou autoforma, com cor, espessura, tracejado, pontas de seta, fonte, preenchimento e borda); `tqq_ann_*` do slide que não existem
  mais no Excel são apagadas; (5) o ChartObject volta ao tamanho original e a pasta volta a "salva" se estava salva. Erro em qualquer passo:
  o tamanho é restaurado e o erro sobe ao chamador (que conta e registra).
- **Colagem corrige o defeito medido.** Depois de colar e ajustar ao slide, a colagem chama `PptSync_SyncShape` (a falha só é registrada:
  a colagem continua valendo; "Atualizar" refaz).
- **Sincronizar a apresentação (`PptSync_Presentation pres, force`).** Junta as formas marcadas de todos os slides (formas soltas; gráfico
  dentro de grupo fica de fora, documentado); agrupa pelo caminho de origem (`LinkFormat.SourceFullName` até o primeiro "!", ou `TQQ_BOOK`);
  **arquivo inexistente** ⇒ o gráfico fica como está e entra na lista do aviso (nunca `Update`: trava, §3.1); pasta já aberta no Excel é
  usada como está (com as alterações não salvas, como o próprio vínculo); senão é aberta **somente leitura**, sem atualizar vínculos, sem
  MRU, com eventos desligados durante o `Open`, e fechada sem salvar no fim; gráfico ausente da pasta ⇒ contado; impressão digital igual à
  `TQQ_FP` e `force = False` ⇒ pulado ("sem mudança"). Devolve o resumo:
  `Gráficos do ThinQQ atualizados: N.` [` Sem mudança: M.`] e, em linhas separadas, `Pasta de origem não encontrada ou sem acesso (o gráfico ficou como estava): <arquivos>. Se ela foi movida ou renomeada, use no PowerPoint Arquivo > Informações > Editar Links para Arquivos > Alterar Origem.`,
  `Gráfico não encontrado na pasta de origem (apagado ou recriado): K.`, `Sem vínculo com o Excel (Quebrar Vínculo no PowerPoint): U.`,
  `Falha ao atualizar: F (veja o log do ThinQQ).` Pasta que existe mas não abre (senha, corrompida) entra na lista "sem acesso". Forma marcada
  cujo gráfico teve o vínculo quebrado no PowerPoint (`ChartData.IsLinked = False`) é contada em "Sem vínculo" e não é tocada.
- **Impressão digital (`PptSync_Fingerprint`).** SHA-256 (o `Sha256_Bytes` do core) do bloco inteiro do gráfico na `_ThinQQ` (valores, inclusive
  a linha de metadados com as opções), dos nomes e textos alternativos das formas `tqq_ann_*` e do tamanho da forma do slide.
- **API pública (nunca propaga erro).** `Core_PptSyncPresentation(pres As Object, Optional force As Boolean = False) As String` (o `.ppam` chama por
  `Application.Run` com o objeto da apresentação) e `Core_PptSyncOpen()` (botão do Excel: todas as apresentações abertas, inclusive as somente
  leitura (atualizadas só na memória; o PowerPoint pergunta se quer salvar ao fechar; só a atualização automática ao abrir pula as somente leitura), `force = True`, aviso por `Core_Notify` com o resumo de cada uma; sem PowerPoint aberto: "Abra a apresentação no PowerPoint e use
  de novo 'Atualizar gráficos no PowerPoint'.").

### 4.3 Suplemento do PowerPoint (`ThinQQ.ppam`, fonte em `src/ppt/`, prefixo `PptTqq_`)
- `Auto_Open` liga `clsPptEvents` (`WithEvents App As Application`: o projeto do PowerPoint já referencia a biblioteca do PowerPoint; nenhuma
  referência extra). Eventos: `AfterPresentationOpen` e `WindowActivate` (este cobre a apresentação que sai do **Modo de Exibição Protegido**
  ao clicar "Habilitar Edição", que não dispara a abertura de novo; cada apresentação é tratada **uma vez enquanto a apresentação está aberta; esquecida ao fechar** (`PresentationCloseFinal`), pelo caminho: reabrir pergunta de novo, e "Salvar como" com outro nome ou pasta pode perguntar mais uma vez).
- **Decisão ao abrir (`PptTqq_OnOpenAction`, pura):** nada sem gráficos marcados, já vista, somente leitura ou modo "nunca"; "sync" no modo
  "sempre"; "ask" no modo "perguntar" (padrão): "Esta apresentação tem N gráfico(s) do ThinQQ vinculado(s) ao Excel. Atualizar agora com os
  dados atuais?" (Sim/Não), seguida de "Pastas de origem:" (as pastas distintas, caminho completo; as remotas primeiro; no máximo 5 e "...";
  caracteres de controle viram espaço; cada caminho com até 150 caracteres, cortado no meio; a pergunta inteira com até 900 caracteres, para a
  dica "Para não perguntar mais" sempre aparecer). **Revisão de segurança:** no modo "sempre", uma pasta de origem remota (`\\servidor\...`,
  `//servidor/...`, `\\?\UNC\`, WebDAV, `http(s)` ou `file:`) que ainda não está aberta no Excel faz o suplemento perguntar mesmo assim
  (`hasRemoteSource`); letras de unidade mapeadas não são detectáveis. A sincronização abre as pastas de origem somente leitura e com as macros
  desligadas (`AutomationSecurity` = 3 durante o `Open`) e as fecha com os eventos desligados. O modo fica em `HKCU\Software\VB and VBA Program Settings\ThinQQ\PowerPoint\OpenMode` (`ask|always|never`),
  apagado pelo desinstalador junto com o resto de `ThinQQ`.
- **Qual Excel.** (1) o Excel já aberto (`GetObject`) **se** ele tem o core do ThinQQ (`Core_Ping` responde "pong") — usa as pastas abertas
  do usuário, com as alterações não salvas, que é o que o vínculo nativo também lê; (2) senão uma instância **nova e invisível**
  (`CreateObject`; o Excel é de várias instâncias: nunca reaproveita a do usuário), onde o `.ppam` abre o core mais novo de
  `%LOCALAPPDATA%\ThinQQ\versions\<x.y.z>\ThinQQ.core.xlam` (somente leitura) e, no fim, fecha essa instância (`Quit`). Nunca carrega o core
  no Excel do usuário que não o tem.
- **Modos degradados.** Sem o core instalado: só os dados (`LinkFormat.Update` nos gráficos cuja pasta existe) e o aviso "O ThinQQ não está
  instalado no Excel deste computador: só os dados foram atualizados; setas e outras anotações do ThinQQ ficaram como estavam." Sem Excel
  (429): "O Excel não está instalado (ou não pôde ser iniciado): os gráficos do ThinQQ não foram atualizados." Excel ocupado (RPC
  0x80010001/0x8001010A): "O Excel está ocupado (uma célula em edição ou uma caixa de diálogo aberta). Termine o que está fazendo no Excel e
  use ThinQQ > Atualizar gráficos." Core antigo sem `Core_PptSyncPresentation` (erro 1004 citando o nome): degrada para só os dados e acrescenta
  "Atualize o ThinQQ no Excel para atualizar também as setas.".
- **Faixa do PowerPoint** (§7.2): aba ThinQQ, grupo "Gráficos do Excel": "Atualizar gráficos" (força, apresentação ativa), "Ao abrir"
  (Perguntar / Atualizar sempre / Não atualizar), "Sobre".
- **Sem o suplemento** (colega sem ThinQQ): os gráficos continuam vinculados normais (vínculo manual): "Design do Gráfico > Atualizar Dados"
  ou "Editar Links" atualiza séries e rótulos ligados; as formas ficam como estavam (documentado).

### 4.4 Casos de borda (decididos)
| Caso | Comportamento |
|---|---|
| Pasta de origem movida/renomeada/sem acesso (rede fora) | Não chama `Update` (trava, §3.1); o gráfico fica; o aviso lista os arquivos e ensina "Alterar Origem". `FileSystemObject.FileExists` num compartilhamento inacessível pode demorar alguns segundos (risco aceito). |
| Apresentação somente leitura | Ao abrir: nada (o usuário não poderia salvar). O botão "Atualizar gráficos" atualiza em memória mesmo assim. |
| Modo de Exibição Protegido (anexo/Internet) | Nenhum evento de apresentação; ao "Habilitar Edição", o `WindowActivate` faz a pergunta. Sem habilitar, nada. Verificação manual. |
| Excel do usuário com a pasta aberta e alterada | Lida com as alterações não salvas (igual ao vínculo nativo, medido); o ChartObject é redimensionado e restaurado; o estado "salva" volta ao que era. |
| Excel do usuário sem o ThinQQ | Nova instância invisível própria (não toca a do usuário). |
| Gráfico dentro de grupo no slide; gráfico de outra pasta que não é ThinQQ | Ignorado (só formas soltas com `TQQ_CHART`). |
| Gráfico editado no PowerPoint (legenda movida, cores) | Cores e formatação ficam (atualização no lugar); se a área interna do slide difere da do Excel em mais de 2 pt, as formas são aplicadas com o deslocamento e o log registra "layout do slide diferente". |
| 30 gráficos | Com a pasta aberta: `Update` 60–280 ms + redimensionar/recalcular/restaurar (~0,3–1 s): ~10–40 s; a impressão digital pula os que não mudaram; a barra de status do Excel mostra "ThinQQ: atualizando gráfico i de N do PowerPoint". |
| 32/64 bits | Só ligação tardia e `Application.Run`; sem `Declare`. Não medido em 64 bits (sem máquina): verificação manual quando houver. |

### 4.5 Excel (faixa do loader)
Menu **PowerPoint** passa a ficar **sempre ativo** (o item novo não exige gráfico selecionado); "Copiar como imagem" e "Copiar como gráfico
vinculado" continuam exigindo um gráfico (`getEnabled` próprio); separador e o item novo **"Atualizar gráficos no PowerPoint"**
(`btnPptSync`, `Rx_OnPptSync` → `Core_PptSyncOpen`; core antigo: "Atualize o ThinQQ para usar este recurso.").

### 4.6 Distribuição (custo de cada decisão)
| Peça | Mudança | Custo (testes afetados) |
|---|---|---|
| Arquivo novo `ThinQQ.ppam` | Fonte em `src/ppt/` (módulos, classe, `ribbon/customUI14.xml`); `build.ps1 -Target ppt` monta com uma instância **própria** do PowerPoint (recusa se já houver `POWERPNT`), `AccessVBOM` do PowerPoint conferido (`Assert-PptVbomTrusted`), salva `.ppam` (Release, formato 30) ou `ThinQQ.ppt.test.pptm` (Test, formato 25, com os testes); faixa inserida por `Add-RibbonToPackage` (genérica, só `imageMso`) | `Build-Common.ps1`; novo `tests/ps/Test-PptAddin.ps1`; `Test-Ribbon.ps1` valida o XML do PowerPoint |
| `release.ps1` | Monta e confere `dist\ThinQQ.ppam` (marcadores de teste em `ppt/vbaProject.bin`), assina junto com os `.xlam` (o SIP do Office cobre `.ppam`), SHA-256 no manifesto | `Test-ReleaseGuard.ps1` |
| Manifesto | Campos novos `pptAddinVersion` (= `coreVersion`), `pptAddinUrl`, `pptAddinSha256` (o `schemaVersion` fica 1: campos extras são ignorados pelos loaders antigos) | `New-ReleaseManifest`, `Test-ReleaseGuard`, `TestUpdater` |
| ZIP | `Get-ReleaseTopFiles` + `ThinQQ.ppam` (soltos 8 → 9 arquivos) | `Test-ReleaseZipContent` (lista esperada) |
| `Publish-Release` | 5 → **6 assets** (`Get-ReleaseAssetNames`) | `Test-PublishRelease.ps1` |
| Instalador | Copia `ThinQQ.ppam` para `%LOCALAPPDATA%\ThinQQ\`, cria `HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\ThinQQ` (`Path` = caminho, `AutoLoad` = 1) e o local confiável `...\PowerPoint\Security\Trusted Locations\LocationThinQQ`; pacote **sem** `.ppam` (ZIP antigo) continua instalando (o suplemento do PowerPoint é opcional); "Feche o Excel" vira "Feche o Excel e o PowerPoint" | `Test-Installer.ps1` (registro de teste `ThinQQ_Test_*`) |
| Desinstalador | Remove a chave `AddIns\ThinQQ` e o local confiável do PowerPoint | `Test-Installer.ps1` |
| Atualizador (loader) | Atualização automática do suplemento: §4.9 (Tarefa A7) | `TestUpdater.bas` (servidor HTTP de teste com `.ppam` falso) |
| Instalador (emenda) | Grava `PpamVersion` (do `pptAddinVersion`) e descarta uma pendência antiga do atualizador | `Test-Installer.ps1` |
| Assinatura | `Sign-ThinQQ.ps1 -Files` recebe também o `.ppam` | `Test-Signing.ps1` (inalterado: lista de arquivos é parâmetro) |
| Documentação | Tutorial (seção `#powerpoint` reescrita: atualizar ao abrir, botão, modos, limites), README (Recursos, Instalação, Verificações manuais, Limitações), CHANGELOG, `release/notes.txt`, `installer/LEIA-ME.txt` | `Test-Docs.ps1` (frases e rótulos) |

### 4.7 Testes (Parte A)
- **Automáticos no core (`tests/vba/core/TestPptSync.bas`)**: funções puras (`PptSync_SourcePath`, `PptSync_SummaryText`, impressão digital muda
  com dado/tamanho/forma e não muda sem mudança), falha injetada (`PptSync_SetFailForTest`) restaura o tamanho do gráfico do Excel (com dublê
  `TestPptFake` acrescido de `Width`/`Height`), e — **só quando nenhum `POWERPNT` existe** e o PowerPoint está instalado (o padrão
  `StartScratchPowerPoint` de `TestPowerPoint.bas`) — integração real: colagem que hoje distorce fica alinhada (regra de três própria com a
  área interna e o eixo **do gráfico do slide**), dados novos + seta nova clonada + seta apagada + animação mantida, pasta inexistente sem
  travar (< 10 s), pasta fechada aberta e fechada, impressão digital pulando, `Core_PptSyncOpen` com instância injetada.
- **Projeto do PowerPoint (`tests/vba/ppt/`)**: corredor próprio `modPptTestRunner` (o comum depende de `ThisWorkbook`), asserts próprios,
  testes das funções puras do `.ppam` (decisão ao abrir, versão mais nova do core, textos, contagem de gráficos marcados numa apresentação
  criada pelo teste).
- **Ponta a ponta (`tests/ps/Test-PptAddin.ps1`, rodado pelo `run-tests` só sem `POWERPNT` aberto e com o VBOM do PowerPoint habilitado; senão
  "SKIP" com o motivo)**: Excel próprio com o core de teste + pasta salva + gráfico com seta; PowerPoint próprio com o `.pptm` de teste;
  colagem; muda o dado; `PptTqq_Sync` com o Excel injetado (`PptTqq_SetExcelForTest`): resumo "atualizados: 1", valores e rótulo da seta novos.
- **Manuais (humano, Tarefa A9)**: abrir pela interface com o modo "perguntar"/"sempre"/"nunca"; anexo em Modo de Exibição Protegido +
  Habilitar Edição; pasta movida; pasta em rede; 30 gráficos; Excel do usuário aberto com a pasta alterada; PowerPoint sem o ThinQQ no Excel;
  instalação pelo ZIP (suplemento aparece na aba ThinQQ do PowerPoint); 64 bits se houver máquina.

### 4.9 Atualização automática do suplemento pelo Excel (emenda)
**Ponto de partida (já entregue na 1.3.0):** o loader verifica o manifesto na abertura do Excel (`Updater_StartupCheck`, a cada
`CheckIntervalHours`, padrão 24 h; recuo de 1 h depois de falha), lê o manifesto pelo endereço estável (`release/manifest.json` da `main`), baixa
pelo `Transport` (XMLHTTP com recurso ao ServerXMLHTTP, proxy do usuário, TLS), confere SHA-256 (manifesto sem SHA é recusado, salvo
`AllowUnverified`), troca o core a quente e agenda a troca do loader. O suplemento passa a ser o terceiro arquivo desse mesmo fluxo.

**Fluxo de ponta a ponta**
1. *Precisa?* (`Updater_PpamNeeded`) — opção `PpamAutoUpdate` ligada (padrão) **e** suplemento instalado (`%LOCALAPPDATA%\ThinQQ\ThinQQ.ppam`
   existe: o atualizador **nunca instala** o suplemento para quem não o instalou pelo ZIP) **e** manifesto com `pptAddinVersion` válida, URL e
   SHA-256 de 64 caracteres **e** `pptAddinVersion` > `PpamVersion` (registro do ThinQQ, gravado pelo instalador e pelo atualizador; ausente =
   0.0.0) **e** essa versão não está já pendente.
2. *Junto com o core* (`Updater_Apply` com core): o `.ppam` é **baixado e conferido antes** de trocar o core; depois que o core foi trocado e
   carregado (`committed`), o `.ppam` é aplicado. Falha no download/SHA do `.ppam`: o core segue (os dois são compatíveis pelo contrato abaixo), o
   suplemento fica como estava e o aviso diz "O suplemento do PowerPoint não foi atualizado agora; nova tentativa na próxima verificação." (a
   comparação de versões repete na próxima verificação). Falha do core: o `.ppam` baixado é apagado e nada é trocado (rollback do core como hoje).
   Nunca fica "suplemento novo + core velho"; "core novo + suplemento velho" só acontece com aviso e é seguro.
3. *Só o suplemento* (core em dia, por exemplo depois de uma falha anterior ou de religar a opção): `Updater_Check` aplica sem abrir o
   formulário (a mesma versão do core já foi aceita) e registra; na verificação manual mostra a frase do resultado.
4. *Troca atômica* (`Updater_PpamSwap`, mesma pasta): `ThinQQ.ppam` → `ThinQQ.ppam.old`, baixado → `ThinQQ.ppam`, apaga `.old`; falha no 2.º passo
   devolve o `.old`. A **renomeação é o teste de trava** (§3.4: o PowerPoint prende o arquivo aberto para edição; um suplemento carregado pode ou
   não prender — não medido): se renomear falha, nada muda.
5. *Pendente* (arquivo preso): o baixado vira `ThinQQ.ppam.new` com `PendingPpam`/`PendingPpamSha256`/`PendingPpamVersion` no registro do ThinQQ e
   é aplicado **automaticamente pelo Excel**: (a) na próxima abertura do Excel (`Loader_Startup`), (b) por nova tentativa a cada 10 min enquanto o
   Excel estiver aberto (`Application.OnTime`, reagendada só enquanto houver pendência). Cada tentativa **reconfere o SHA-256** do `.new` (arquivo
   em pasta do usuário); adulterado ou ausente = descartado. Trocado com o PowerPoint em execução mas sem trava = "feche e abra o PowerPoint".
6. *Avisos* (`Updater_PpamStatusText`): "Suplemento do PowerPoint atualizado."; "Suplemento do PowerPoint atualizado: feche e abra o PowerPoint
   para usar a versão nova."; "O suplemento do PowerPoint será atualizado automaticamente quando o PowerPoint for fechado (o Excel tenta de novo a
   cada 10 minutos e na próxima abertura)."; "O suplemento do PowerPoint não foi atualizado agora; nova tentativa na próxima verificação." — a
   frase entra no aviso de sucesso do atualizador; a linha de versão do formulário de atualização acrescenta "Suplemento do PowerPoint: a > b".

**Mecanismos avaliados para a troca com o PowerPoint aberto**
| Opção | Avaliação | Veredito |
|---|---|---|
| Excel na próxima abertura + nova tentativa a cada 10 min (`OnTime`) com o teste de trava pela renomeação | Funciona sem nada no PowerPoint; não usa `GetObject` (nunca anexa ao PowerPoint do usuário); a trava real é medida pela própria operação; o WMI só decide a frase "feche e abra o PowerPoint" | **Escolhida** |
| O próprio `.ppam` se trocar (no `Auto_Open`/`PresentationOpen`) | O arquivo carregado é o que precisaria ser trocado; a troca só valeria no próximo início do PowerPoint, e um `.ppam` "copiador" separado seria um segundo suplemento registrado (mais registro, mais assinatura) | Rejeitada |
| Lançador no início do PowerPoint | O PowerPoint não tem um ponto de início fora dos suplementos (seria um suplemento a mais) | Rejeitada |
| Só na próxima abertura do Excel | Simples, mas quem deixa o Excel aberto o dia todo ficaria com o suplemento velho até o dia seguinte | Incluída como (a), não sozinha |
| Esperar o fim do PowerPoint por `GetObject`/eventos do PowerPoint | Anexaria ao PowerPoint do usuário (proibido) e exigiria referência | Rejeitada |

**Contrato de versão (ppam × core × loader)**
- Core: `CORE_PPT_API = 1` e `Core_PptApiVersion()` (A2). Regra: a assinatura de `Core_PptSyncPresentation(pres, force)` nunca muda; mudança
  incompatível = função nova e `CORE_PPT_API` maior, mantendo a antiga ⇒ core mais novo atende suplemento mais velho.
- Suplemento: `PPT_REQUIRED_API = 1`; antes de chamar, lê `Core_PptApiVersion` (core sem ela = 0, anterior a 1.4.0) e, se menor, cai no modo "só
  os dados" com "Atualize o ThinQQ no Excel para atualizar também as setas." ⇒ suplemento mais novo nunca chama API inexistente.
- Loader: só baixa e troca arquivos; `pptAddinVersion` = `coreVersion` no release (os dois saem juntos); loader antigo (sem a emenda) ignora os
  campos novos (o suplemento fica na versão do ZIP até a próxima instalação, e o contrato garante o funcionamento).

**Controle do usuário:** `PpamAutoUpdate` (padrão "1"): caixa "Atualizar também o suplemento do PowerPoint (feche o PowerPoint para aplicar na
hora)" no formulário de atualização (só para quem tem o suplemento) e pergunta em ThinQQ > Configurações. Desligada: nada é baixado.

**Confiança e segurança:** o `.ppam` trocado fica no **mesmo caminho**, dentro do local confiável do PowerPoint que o instalador cria
(`...\PowerPoint\Security\Trusted Locations\LocationThinQQ` = `%LOCALAPPDATA%\ThinQQ\`) e registrado no mesmo `AddIns\ThinQQ\Path` ⇒ continua
carregando sem aviso de macro, assinado ou não (a release assina o `.ppam` como os `.xlam`). Sem marca da Internet (§3.4): o `Transport` grava os
bytes; o ZIP baixado pelo navegador é desbloqueado pelo instalador (`Unblock-File`, já existente). SHA-256 obrigatório e fixado pelo manifesto (que
vem do endereço estável por HTTPS; `AllowUnverified` **não** vale para o suplemento); o atualizador nunca abre nem executa o `.ppam`; troca por
renomeação na mesma pasta; temporário e `.old` apagados em todo caminho; `.new` reconferido antes de entrar; mesmo `Transport` (proxy/TLS).

**Testes (A7):** `TestUpdater.bas` com o servidor `tests/http/server.js` servindo um `.ppam` falso (gerado pelo `run-tests` com o manifesto de teste):
não instalado nunca instala; troca com o arquivo livre (HTTP); arquivo preso (aberto com trava exclusiva pelo próprio teste: mesma trava que o
PowerPoint aplica a arquivo aberto para edição) ⇒ pendente, nova tentativa agendada, aplicada ao liberar; SHA divergente; `.new` adulterado; opção
desligada; PowerPoint "em execução" (gancho de teste) ⇒ "feche e abra"; core ok + `.ppam` 404 ⇒ core trocado e aviso; só o suplemento sem
formulário; linha de versão. **Manual (A9):** PowerPoint real com o suplemento carregado (trava ou não), fechar e ver a troca em ≤ 10 min ou na
reabertura do Excel; "Sobre" do PowerPoint mostra a versão nova.

## 5. Decisões — Parte B (1.5.0)

### 5.1 Fluxo escolhido: modo armado + dois cliques
1. Com um gráfico ThinQQ selecionado, **Setas › Delta clicando nas barras** arma o modo naquele gráfico (o botão de novo cancela). Barra de
   status: "ThinQQ - Delta: clique na barra de ORIGEM (Esc ou um clique numa célula cancela)."
2. 1.º clique numa barra/segmento/ponto = **origem**; uma moldura laranja temporária (`tqq_pickmark`, fora do prefixo `tqq_ann_`) marca a
   origem; status "ThinQQ - Delta: agora clique na barra de DESTINO."
3. 2.º clique noutra barra = **destino**; a seta é inserida, a moldura some, o modo desarma.
4. Cancela: Esc (`Application.OnKey`, enquanto armado), clique numa célula (`SheetSelectionChange`), clique noutro gráfico ThinQQ (evento
   `Select` de outro gráfico), o botão de novo, 60 s sem clique (`OnTime`). Clicar em área vazia, eixo, legenda ou numa seta: continua armado.
- **Primeiro clique do Excel seleciona a série inteira** (`Arg2 = -1`, §3.2): o `MouseUp` + `GetChartElement` acha o ponto **no mesmo
  clique** (caminho principal); se só o `Select` chegar com `Arg2 = -1`, o status pede "clique de novo na mesma barra para escolher só ela" (o
  2.º clique do Excel seleciona o ponto). O mesmo clique chegando pelos dois caminhos é contado uma vez (mesma série/ponto vindo da outra
  fonte em menos de 0,6 s).
- **Alternativas avaliadas.** Ctrl+clique: não existe seleção múltipla em gráfico (rejeitada). Menu de contexto do gráfico por `CommandBars`
  ("Chart"/"Series"): funciona só no clique direito sobre a série e não diz qual ponto (o `Select` já diria), além de mexer em barras de
  comando globais do usuário (rejeitada). Minibarra/formulário não modal com dois combos: é o diálogo de hoje com outra cara (a Seta diferença
  continua sendo a alternativa por teclado). Realimentação só pela barra de status: fraca para quem não olha o rodapé (moldura escolhida, só
  enquanto armado).
- **Sem interferência quando desarmado:** os manipuladores saem na primeira linha (`If Not mArmed Then Exit Sub`); custo medido ~3 ms por
  evento já com registro.
- Gráfico em **folha de gráfico** não é do ThinQQ (só ChartObjects `tqq_*`); planilha protegida (`ProtectDrawingObjects`): aviso "A planilha
  está protegida. Desproteja a planilha para inserir o delta."; gráfico em grupo de formas: os eventos do ChartObject agrupado não chegam ao
  observador (Watchers liga só os `ChartObjects` soltos) ⇒ o modo arma mas nada acontece: limite documentado.

### 5.2 Modelo de dados
- **Referência** `"k:c"` (texto) resolvida por família (`Delta_ResolvePick`, a partir de `Arg1` = índice em `SeriesCollection` e `Arg2` = ponto):
  | Família/tipos | Clique aceito | Referência | Valor (texto) | Extensão da barra (onde a seta encosta) |
  |---|---|---|---|---|
  | `clustered`, `barclustered`, `line` | ponto de série de dados | `k:c`, k = ordinal da série de dados (auxiliares `tqq_*` não contam) | valor do ponto | coluna: de 0 ao valor; linha: o próprio ponto |
  | `stacked`, `barstacked`, `combo` | segmento de série de dados (na combinação, também a linha) | `k:c` | valor do segmento | borda externa do segmento: positivos sobem e negativos descem na ordem das séries |
  | `waterfall`, `barwaterfall` | qualquer série da coluna (todas são `tqq_*`) | `0:c` (a barra da categoria) | `End` (valor acumulado, como a Seta diferença) | `Lo`/`Hi` do bloco (colunas 7 e 8) |
  | `waterfallpair` | séries de colunas nas fatias A ou B (o vão e as XY/linhas auxiliares são recusados) | `j:i` (j = 1 lado A, 2 lado B; i = etapa) | `Val` do lado (o mesmo da seta de variação) | `Lo`/`Hi` do lado |
- **Texto alternativo da linha principal:** `kind=delta;v=1;from=<ref>;to=<ref>;fmt=both|abs|pct`. As pernas `_t1`/`_t2` levam
  `kind=deltaleg;owner=<nome>` e o rótulo `_lbl` o `kind=label;owner=<nome>` de sempre. Versão `v=1` para formatos futuros; as setas antigas
  (`kind=diff;from=i;to=j`, categorias) continuam com o código delas, sem migração.
- **Texto** (`Delta_Text`): `both` = `"+5 (+50%)"` (igual à Seta diferença), `abs` = `"+5"`, `pct` = `"+50%"`; percentual `(B − A) / |A|`; `A = 0`:
  `both` mostra só o valor e `pct` mostra "n/d"; formato do valor pelo `numfmt` do gráfico (`Style_NumberFormat`, primeira seção, `Format$`
  com sinal), o que segue a localidade. Diferença conhecida (documentada): a seta de variação da cascata comparativa usa `(B − A) / A` (−4 → −7
  = +75%); o delta, como a Seta diferença, dá −75%.
- **Recálculo:** `Ann_Refresh` (eventos `Calculate`/`Resize`, Atualizar) relê valores e extensões pelas referências. Referência que não existe
  mais (menos séries/categorias, dado novo) ⇒ seta oculta (como as demais). "Alterar dados" recria o gráfico e, como hoje, **não** preserva
  setas (documentado); copiar o gráfico copia as formas com o texto alternativo (continua valendo).
- **Várias setas**: sem limite; faixas sem colisão (§5.3). **Edição**: "Delta: valor / % / ambos" alterna o texto da seta selecionada (ou de
  todas do gráfico se nenhuma estiver selecionada); excluir = "Remover setas" (todas) ou apagar as formas à mão (as órfãs são limpas no
  próximo refresh). **Fora:** arrastar a seta (o refresh a recoloca), alterar origem/destino de uma seta existente (apague e insira de novo).

### 5.3 Desenho e faixas
- **Cotovelo** (colunas, acima): perna `_t1` da origem (3 pt acima da borda da barra) até a linha horizontal, linha principal até sobre o
  destino, perna `_t2` descendo com **ponta de seta** no destino; rótulo `_lbl` (caixa branca com borda cinza, a de sempre) centrado na linha
  horizontal. Barras horizontais: o mesmo girado (linha vertical à direita das barras). Ambas as barras sem parte positiva (as duas
  inteiramente ≤ 0): abaixo (à esquerda nas barras).
- **Emenda (revisão da B2):** com Rótulos (ou Totais) ligados, a perna começa **depois do rótulo de valor de fora** da barra (agrupadas,
  linha, a linha da combinação; com Totais, o segmento de cima da pilha): 3 + 18 pt nas colunas, 3 + máx(18, largura estimada do rótulo +
  6) nas barras (a mesma folga entra na altura base das barras); rótulos de dentro (segmentos das empilhadas, cascatas), âncora no zero e
  Rótulos desligados ficam nos 3 pt.
- **Altura base**: a borda mais externa de **todas** as barras entre origem e destino (inclusive; nas agrupadas, o maior valor positivo de
  qualquer série; nas empilhadas, a soma dos positivos; nas cascatas, `Hi`; na comparativa, a coluna `Top` das fatias) + 8 pt + meia altura do
  rótulo + 18 pt quando Rótulos ou Totais estão ligados.
- **Faixas** (`Ann_LayoutDeltas`): as setas visíveis são ordenadas pela largura do trecho (menor primeiro); cada uma começa na altura base e
  sobe (desce) para ficar a `meia altura própria + meia altura da outra + 6 pt` de toda seta já posta cujo trecho se sobrepõe (com 2 pt de
  margem) e que fica do mesmo lado. Determinístico e sem colisão de rótulos entre setas sobrepostas; setas de trechos disjuntos compartilham a
  altura.
- **Espaço** (`Ann_EnsureDeltaRoom`): depois do layout, se o rótulo ou a linha mais externa passa da borda livre (2 pt do topo; da borda de
  baixo ou da legenda embaixo; da borda direita ou da legenda à direita nas barras), a área de plotagem encolhe pelo déficit e o layout é
  refeito; até 3 passadas (o eixo automático pode mudar), como nos colchetes do Gantt. Também em `Ann_EnsureRoomAfterChange`.
- **Coordenadas**: sentido das categorias = centro da categoria (`Geom_ChartCatCoord`; nas agrupadas, o centro do **ponto** da série,
  `Points(c).Left + Width/2` ou `Top + Height/2`, porque cada série tem a sua barra); sentido dos valores = regra de três do eixo
  (`Geom_ChartValCoord`).

### 5.4 Componentes
- `modDelta.bas` (puro + leitura do gráfico): `Delta_Text`, `Delta_ParseRef`, `Delta_SeriesIndex`, `Delta_DataOrdinal`, `Delta_ResolvePick`,
  `Delta_Extent` (valor, `lo`, `hi`), `Delta_SpanEdge`, `Delta_AxisIndex`, `Delta_CatCenter`.
- `modAnnotations.bas`: `Ann_AddDelta`, `Ann_LayoutDeltas` (chamado no fim de `Ann_Refresh`), `Ann_EnsureDeltaRoom`, `Ann_CycleDeltaFormat`,
  `Ann_SelectedDeltaName`, `Ann_RefusalText` ("delta"), limpeza de órfãs e da moldura.
- `modDeltaPick.bas` (estado do modo): `DeltaPick_Arm`, `DeltaPick_Disarm`, `DeltaPick_OnSelect`, `DeltaPick_OnMouseUp`, `DeltaPick_OnSheetClick`,
  `DeltaPick_EscKey`, `DeltaPick_Timeout`, `DeltaPick_IsArmedFor`, `DeltaPick_FromRef`.
- `clsChartWatcher.cls`: `Cht_Select`, `Cht_MouseUp` (chamam o modo); `clsAppEvents.App_SheetSelectionChange` chama `DeltaPick_OnSheetClick`.
- `modCoreApi.bas`: `Core_DeltaPick` (arma/cancela), `Core_AddDelta(fromRef, toRef, Optional fmt, Optional cht)` (API e gerador da
  demonstração; avisos 7001/7003), `Core_DeltaCycleFormat`.
- Erros: **7001** recusa da matriz (`"O delta entre barras não está disponível para este tipo de gráfico."`), **7003** referência inválida
  (`"Escolha duas barras diferentes do gráfico."`), mostrados como aviso.

### 5.5 Matriz
Recurso novo `delta` ("Delta clicando nas barras"): `stacked, clustered, line, waterfall, barstacked, barclustered, barwaterfall, combo,
waterfallpair` (§6). Recusados (7001): 100% e Mekko (o "nível" é percentual; a Seta diferença também recusa), área (não há barra), pizza,
rosca, Gantt (o colchete cobre), borboleta (dois lados espelhados; ficaria para depois), dispersão e bolhas (não há categoria).

### 5.6 Testes (Parte B)
- `tests/vba/core/TestDelta.bas`: textos (localidade por `Format$`), referências, resolução por família (inclusive auxiliares e o vão da
  comparativa), extensões (segmentos positivos e negativos), geometria **independente** (regra de três com `PlotArea.Inside*` e escala reais;
  centro do ponto pelas `Points`) para colunas, barras, negativos abaixo, faixas sobrepostas, sala no topo com escala fixa, recálculo com
  dados novos, referência inválida oculta, recusas, órfãs.
- `tests/vba/core/TestDeltaPick.bas`: eventos **reais** por `Points(i).Select` (origem, destino, série inteira, mesma barra, auxiliar),
  `DeltaPick_OnMouseUp` com pixels calculados (pt × 4/3) depois de conferir `GetChartElement` (se o ambiente não der `xlSeries`, registra e
  pula só esse trecho), cancelamentos (célula, outro gráfico, botão, Esc e tempo chamados diretamente), desarmado sem efeito.
- Manuais (Tarefa B6): cliques reais com o mouse nos 9 tipos; o 1.º clique num gráfico não ativado; Esc com o gráfico ativo; zoom 75%/150%
  e monitor a 125%; delta no PowerPoint (com a Parte A).

## 6. Matriz de capacidades 1.5.0 (`Kinds_Supports`)

Tipos: os 21 da 1.3.0. Recurso novo: `delta`. Esta tabela é lida pelo teste `Test_Kinds_Delta_Matrix_ExhaustiveFromSpec` (cópia em
`docs/superpowers/specs/`). A Parte A não muda a matriz.

| feature | tipos |
|---|---|
| totals | stacked, stacked100, barstacked, barstacked100, area, area100, combo, mekko, mekkounits, butterfly, waterfallpair |
| labels | todos |
| connectors | stacked, stacked100, barstacked, barstacked100, waterfall, barwaterfall, combo, waterfallpair |
| numfmt | stacked, clustered, stacked100, line, waterfall, barstacked, barclustered, barstacked100, barwaterfall, area, area100, combo, mekko, mekkounits, pie, doughnut, butterfly, waterfallpair |
| palette | todos |
| gap | stacked, clustered, stacked100, barstacked, barclustered, barstacked100, waterfall, barwaterfall, combo, gantt, butterfly, waterfallpair |
| legend | todos exceto gantt (cascata só com várias séries) |
| lineend | line, combo |
| ind100 | stacked100, barstacked100, area100, mekko |
| vline | stacked, clustered, line, barstacked, barclustered, area, combo, mekkounits |
| secondary | combo |
| pieout | pie |
| cagr | stacked, clustered, stacked100, line, waterfall, barstacked, barclustered, barstacked100, barwaterfall, area, area100, combo |
| diff | stacked, clustered, line, waterfall, barstacked, barclustered, barwaterfall, area, combo |
| bracket | gantt |
| samescale | stacked, clustered, line, waterfall, barstacked, barclustered, barwaterfall, area, combo, mekkounits, butterfly, waterfallpair |
| roundsum | stacked, stacked100, barstacked, barstacked100, area, area100 |
| break | stacked, clustered, barstacked, barclustered |
| trend | scatter, bubble |
| quad | scatter, bubble |
| pairdiff | waterfallpair |
| pairabs | waterfallpair |
| inbar | waterfallpair |
| delta | stacked, clustered, line, waterfall, barstacked, barclustered, barwaterfall, combo, waterfallpair |

## 7. Faixas de opções

### 7.1 Excel (loader)
- **1.4.0 (Parte A)**: menu `mnuPpt` sem `getEnabled` (sempre ativo), `screentip="PowerPoint"`, supertip "Cola o gráfico selecionado no slide
  ativo do PowerPoint (como imagem ou gráfico vinculado) e atualiza os gráficos vinculados das apresentações abertas."; `btnPptImage` e
  `btnPptChart` como hoje (com `getEnabled="Rx_GetEnabled"`); `menuSeparator id="msepPpt"`; `btnPptSync` "Atualizar gráficos no PowerPoint",
  `imageMso="Refresh"`, `onAction="Rx_OnPptSync"`, sem `getEnabled`, supertip "Atualiza dados e setas dos gráficos do ThinQQ vinculados nas
  apresentações abertas no PowerPoint (também acontece ao abrir a apresentação, com o suplemento do PowerPoint)." Contagens: 8 menus, 3 botões
  no `mnuPpt`.
- **1.5.0 (Parte B)**: menu **Setas**, logo depois de `btnDiff`: `btnDelta` "Delta clicando nas barras" (`image="tqq_arrows16"`, `tag="delta"`,
  `onAction="Rx_OnDelta"`, `getEnabled="Rx_GetEnabled"`, supertip "Clique aqui e depois em duas barras do gráfico (origem e destino): a seta
  mostra a diferença em valor e em % entre elas. Esc ou um clique numa célula cancela.") e `btnDeltaFmt` "Delta: valor / % / ambos"
  (`image="tqq_arrows16"`, `tag="delta"`, `onAction="Rx_OnDeltaFormat"`, `getEnabled`). Os 3 botões `Rx_OnAnnotation` (cagr, diff,
  bracket) ficam; o supertip do menu Setas acrescenta "delta entre duas barras escolhidas com o mouse".

### 7.2 PowerPoint (`src/ppt/ribbon/customUI14.xml`)
Aba `tabThinQQPpt` "ThinQQ", grupo `grpPptSync` "Gráficos do Excel": `btnPptSyncNow` "Atualizar gráficos" (grande, `imageMso="Refresh"`,
`PptRx_OnSync`), `ddOpenMode` "Ao abrir" (itens `modeask` "Perguntar", `modealways` "Atualizar sempre", `modenever` "Não atualizar";
`PptRx_GetOpenMode`/`PptRx_OnOpenMode`), `btnPptAbout` "Sobre" (`imageMso="Info"`, `PptRx_OnAbout`). Só `imageMso` (sem ícones próprios).

## 8. Integração, ordem e versões

### 8.1 Ordem: Parte A primeiro (1.4.0), Parte B depois (1.5.0)
- A colagem vinculada **atual** (1.3.0) já entrega anotações fora do lugar sempre que o PowerPoint ajusta o tamanho (§3.1): a Parte A corrige
  um defeito em produção, e as setas da Parte B, se viessem antes, nasceriam com o mesmo defeito no PowerPoint.
- O motor da Parte A sincroniza **qualquer** forma `tqq_ann_*` (clonando as que o slide não tem): os deltas da Parte B passam a funcionar no
  PowerPoint sem código novo na Parte A (o teste de integração entra na Tarefa B2, condicionado ao PowerPoint).
- A Parte A tem o pré-requisito do usuário (VBOM do PowerPoint) e as incertezas maiores (abertura pela interface, Modo de Exibição
  Protegido): começar por ela dá tempo de a verificação manual acontecer enquanto a Parte B é feita. Se o `.ppam` (A4) travar, as Tarefas
  A1–A3 já entregam o botão do Excel (alternativa (a) completa) e podem sair numa 1.4.0 sem o `.ppam` (o plano diz como).

### 8.2 Como as duas partes se combinam
O delta é forma estática dentro do gráfico. No PowerPoint ele só fica certo depois de passar pelo core: ao reabrir a apresentação, o `.ppam`
chama o core, que redimensiona o gráfico do Excel para o tamanho do slide, recalcula as setas (inclusive os deltas, com as faixas para aquele
tamanho) e copia a foto delas para dentro do gráfico do slide. Sem o ThinQQ, o delta no slide fica com o último texto e a última posição.

### 8.3 Versões e documentação
`version.json` não é tocado pelas tarefas. Loader 1.4.0 com core 1.3.x: "Atualizar gráficos no PowerPoint" avisa "Atualize o ThinQQ para usar
este recurso."; `.ppam` com core 1.3.x (sem `Core_PptApiVersion`): modo "só os dados" com o aviso de atualizar; loader 1.3.x ignora os campos
`pptAddin*` (o suplemento fica na versão do ZIP; contrato §4.9). Loader 1.5.0 com core 1.4.x: os botões do delta avisam
"Atualize" (`RunCoreFeature`). CHANGELOG `[Não lançado]` ganha "Adicionado"/"Corrigido"; `release/notes.txt` uma linha por parte.

| Recurso | CHANGELOG | Tutorial | README | Exemplos/demonstração | Testes de documento |
|---|---|---|---|---|---|
| Atualizar no PowerPoint (A) | Adicionado (suplemento do PowerPoint, botão), Corrigido (setas fora do lugar na colagem vinculada) | `#powerpoint` reescrita | Recursos, Instalação (3 arquivos), Verificações manuais, Limitações | — | Test-Docs: rótulos `btnPptSync`, `btnPptSyncNow`, `ddOpenMode`, frases dos limites, mensagens |
| Delta (B) | Adicionado | `#setas` (subseção `#delta`), matriz (coluna Delta), mensagens 7001/7003 | Recursos, Verificações manuais, Limitações | E02 ganha um delta; Exemplos: o de colunas agrupadas ganha um delta (contagens iguais) | Test-Docs: 24 recursos, rótulos, mensagens; DemoCoverage `delta` |

## 9. Riscos e mitigações

| Risco | Prob. | Mitigação |
|---|---|---|
| VBOM do PowerPoint desabilitado (build do `.ppam` impossível) | certa até o usuário habilitar | pré-requisito explícito antes da A4; `Assert-PptVbomTrusted` com a instrução; A1–A3 entregáveis sem o `.ppam` |
| `AfterPresentationOpen`/`WindowActivate` não dispararem como esperado (Modo de Exibição Protegido, abrir pelo Explorer com o PowerPoint fechado) | média | dois eventos + tratado uma vez por caminho; botão manual sempre disponível; verificação manual na A9 |
| `LinkFormat.Update` travar (fonte inexistente/inacessível) | certa se chamado | existência conferida antes; nenhuma chamada sem o arquivo; teste com tempo máximo |
| Geometria do Excel invisível desatualizada depois de redimensionar | média | `PptSync_ForceLayout` (exportar) antes e depois do refresh; teste de alinhamento independente |
| Áreas internas diferentes (gráfico editado no PowerPoint) | baixa | deslocamento pela área interna; log quando a diferença passa de 2 pt |
| Clonagem perder formatação de forma nova | baixa | só formas criadas depois da colagem; copia linha/texto/preenchimento; teste da seta nova |
| PowerPoint de instância única nos testes | certa | testes reais só sem `POWERPNT`; destino fixo; `UserWindowAppeared` (padrão existente) |
| Hospedeiro com pouca memória (0x800706BE) | média | um aplicativo de cada por vez; suíte reexecutada; PowerPoint fechado entre testes |
| `MouseUp` não chegar no 1.º clique / DPI diferente | média | o `Select` sozinho completa o fluxo (2.º clique na mesma barra); `GetChartElement` usa as coordenadas do próprio evento (sem conversão); manual com 125% |
| `OnKey "{ESC}"` não disparar com o gráfico ativo | média | outros 4 cancelamentos; status diz "Esc ou um clique numa célula"; manual |
| Moldura temporária ficar na pasta (Excel fechado com o modo armado) | baixa | `Ann_Refresh` apaga `tqq_pickmark` de gráfico não armado; fora do prefixo `tqq_ann_` (não vai ao PowerPoint) |
| Faixas: rótulos de setas não sobrepostas se tocando | baixa | só setas com trechos sobrepostos competem; teste de interseção de caixas |
| Tempo da suíte | média | módulos novos com `-Filter`; testes de PowerPoint só quando disponíveis |
| Suplemento carregado não prender o arquivo e a troca acontecer com o PowerPoint aberto (emenda) | média | a versão nova vale no próximo início do PowerPoint; aviso "feche e abra o PowerPoint" (WMI); contrato de API garante o par misto |
| Suplemento carregado prender o arquivo indefinidamente (PowerPoint nunca fechado) | média | pendência reconferida a cada 10 min e a cada abertura do Excel; aviso "quando o PowerPoint for fechado"; core novo atende o suplemento velho |
| `.new` ou temporário adulterado na pasta do usuário | baixa | SHA-256 reconferido antes de qualquer troca; adulterado é apagado |
| Par core/suplemento desencontrado | baixa | `.ppam` só trocado depois do core; `Core_PptApiVersion`/`PPT_REQUIRED_API`; falha do `.ppam` sempre avisada e repetida |
| `OnTime` da nova tentativa disparar com o loader descarregado | baixa | `OnTime` aponta para o loader (que é o dono do atualizador); ao fechar o Excel nada fica agendado; a próxima abertura aplica |

## 10. Fora do escopo (e por quê)
- **Parte A:** diálogo "Data Links" com estado por elemento e religar fontes (o PowerPoint já tem "Editar Links para Arquivos"); vínculo de
  tabelas, textos e imagens; atualização ao vivo sem reabrir quando as duas aplicações estão abertas (há o botão no Excel e no PowerPoint);
  gráficos dentro de grupos no slide; imagens coladas (sem vínculo); Mac; atualização sem o Excel instalado. Emenda: instalar o suplemento
  pela primeira vez pelo atualizador do Excel (instalação nova continua sendo pelo ZIP, que registra o suplemento e o local confiável);
  atualizar um `.ppam` instalado à mão fora de `%LOCALAPPDATA%\ThinQQ`; o próprio PowerPoint verificar atualizações; atualizar o suplemento sem o
  Excel aberto; rebaixar versão (o atualizador só sobe).
- **Parte B:** arrastar a seta e as pontas; alterar origem/destino de uma seta existente; delta em 100%, Mekko, área, pizza, rosca, Gantt,
  borboleta, dispersão e bolhas; delta entre gráficos; preservar setas em "Alterar dados" (vale para todas as setas, como hoje); seleção por
  teclado das barras (a Seta diferença por categorias continua).

## 11. Pré-flight do plano (varredura de coerência)

| Par | Produz | Consome | Resultado |
|---|---|---|---|
| A0 → todas | spec e plano promovidos em `docs/superpowers/` | B0 lê a §6 da cópia | limpo: B0 confere que a cópia existe (se a Parte B rodar sem a A, B0 faz a cópia) |
| A1 → A2 | `PptSync_SyncShape`, `PptSync_Tag`, `PptSync_SourcePath`, `PptSync_Fingerprint`, constantes `PPTSYNC_TAG_*`, `PptSync_SetFailForTest` | `PptSync_Presentation` | limpo |
| A1 → A2/A4 | `TestPptFake` com `Width`/`Height`; helpers privados de PowerPoint copiados em `TestPptSync.bas` | testes de A2 no mesmo módulo | limpo: A2 acrescenta ao módulo de A1 |
| A2 → A3 | `Core_PptSyncOpen` | `Rx_OnPptSync` (`RunCoreFeature`) | limpo |
| A2 → A4 | `Core_PptSyncPresentation(pres, force)` | `PptTqq_Sync` por `xl.Run` | limpo: dois argumentos sempre passados |
| A4 → A5/A6 | `dist\ThinQQ.ppam` (Release) e `dist\ThinQQ.ppt.test.pptm` (Test); `Write-TestDispatchModule` com `AllTests()` | instalador, release, publish, updater | limpo |
| A2 → A4 | `Core_PptApiVersion` (`CORE_PPT_API = 1`) | `PptTqq_ApiCompatible` antes de `Core_PptSyncPresentation` | limpo |
| A4 → A6 | carimbo `Comments = "ThinQQ ppt x.y.z"` no `.ppam` | release confere marcadores; manifesto `pptAddinVersion` | limpo |
| A6 → A5/A7 | manifesto `pptAddinVersion`/`pptAddinUrl`/`pptAddinSha256` | instalador grava `PpamVersion`; atualizador compara e baixa | limpo |
| A5 → A7 | `PpamVersion`, `<raiz>\ThinQQ.ppam` (estado "instalado"), pendência descartada ao reinstalar | `Updater_PpamNeeded`, `Updater_ApplyPendingPpam` | limpo |
| A3/A4/A7 → A8 | rótulos exatos das faixas; frases de `Updater_PpamStatusText`; caixa do formulário | tutorial, README, Test-Docs | limpo |
| A5/A6/A7 → A9 | ZIP com 9 soltos, 6 assets, atualização automática | suíte completa e verificação manual | limpo |
| B0 → B1..B4 | `Kinds_Supports(..., "delta")`, `Kinds_FeatureLabel("delta")`, `Ann_RefusalText(..., "delta")`, `DELTA_REFUSAL` | `Delta_ResolvePick`, `Ann_AddDelta`, `Core_DeltaPick`, faixa | limpo |
| B1 → B2 | `Delta_*` (texto, refs, extensões, centros, borda do trecho) e as constantes `WF_COL_LO`/`WF_COL_HI` | `Ann_AddDelta`, `Ann_LayoutDeltas` | limpo |
| B2 → B3 | `Core_AddDelta(fromRef, toRef, fmt, cht)` | `DeltaPick_Take` | limpo |
| B3 → B4 | `Core_DeltaPick`, `DeltaPick_*` | `Rx_OnDelta` | limpo |
| B2 → B4 | `Ann_CycleDeltaFormat`, `Ann_SelectedDeltaName` | `Core_DeltaCycleFormat` | limpo |
| B4 → B5 | rótulos "Delta clicando nas barras", "Delta: valor / % / ambos" | tutorial, DemoCoverage, Test-Docs | limpo |
| A1 ↔ B2 | motor clona `tqq_ann_*` (linhas e caixas de texto) | delta = 3 linhas + caixa de texto | limpo: teste B2 condicional ao PowerPoint |

| Tarefa | Autoconsistência | Resultado |
|---|---|---|
| A1 | testes usam só `PptSync_*` públicas, `Core_*`, `Ann_*`, `Ppt_*`; regra de três com a escala do gráfico do slide | limpo |
| A2 | resumo com textos exatos de §4.2; arquivo inexistente nunca chega ao `Update` | limpo |
| A3 | contagens 8 menus/3 itens no `mnuPpt` batem com o XML descrito | limpo |
| A4 | `.ppam` sem referências além das padrão do PowerPoint; corredor e asserts próprios; `BUILD_SUFFIX` no nome do core | limpo |
| A5 | instalador aceita pacote sem `.ppam`; desinstalador remove as 2 chaves novas | limpo |
| A6 | 6 assets em `Get-ReleaseAssetNames`, lista do ZIP, manifesto com os 3 campos do suplemento | limpo |
| A7 | testes usam só funções públicas do updater e os ganchos `Updater_TestSet*`; fixture `.ppam` gerada pelo `run-tests` com o SHA no manifesto de teste; trava simulada pelo próprio teste | limpo |
| B1 | resolução e extensões puras sobre o gráfico; testes comparam com valores da tabela | limpo |
| B2 | geometria por regra de três própria e `Points` reais | limpo |
| B3 | eventos reais; manipuladores chamáveis diretamente | limpo |
| B4 | 3 botões `Rx_OnAnnotation`, 2 botões novos com tag `delta` e `getEnabled` | limpo |
| B5 | contagens de exemplos inalteradas; DemoCoverage com `delta` | limpo |

## 12. Estimativa

**Calibração (medido neste projeto):** a cascata comparativa (10 tarefas, 28 commits, 6040 linhas, com pesquisa e sondas) levou ~1 dia corrido;
implementador 10–45 min por tarefa (integração com várias rodadas de teste ~45 min), revisão 2–6 min, rodada de correção 6–15 min, suíte completa
7–15 min (às vezes cai por memória e é reexecutada), planejamento + sondas 1–2 h; quase toda tarefa teve 1 rodada de correção; houve quedas de
agente (retomadas). Este planejamento com sondas levou ~2,5 h (fora das contas abaixo).

**Suposições:** subagentes como na cascata (implementador, revisor, 1 rodada de correção típica, 2 nas tarefas de integração com o PowerPoint e de
eventos); suíte completa por tarefa (7–15 min; +50% de reexecução no pessimista); o usuário habilita o VBOM do PowerPoint antes da A4; nenhum
PowerPoint do usuário aberto durante as rodadas que testam o PowerPoint (senão os testes reais pulam e a verificação passa para o humano); revisões
em paralelo com a tarefa seguinte quando não mexem no mesmo módulo (o "corrido" fica ~10% abaixo da soma).

### 12.1 Por tarefa (minutos; otimista / provável / pessimista)
| Tarefa | Agêntico (impl. + revisão + correção + suíte) | Humano | Observação |
|---|---|---|---|
| A0 Promover spec/plano | 10 / 15 / 25 | — | só cópia + Test-Privacy |
| A1 Motor + colagem corrigida | 75 / 120 / 180 | — | integração com PowerPoint real; 2 rodadas prováveis |
| A2 Apresentação inteira + botão do core (+ `Core_PptApiVersion`) | 60 / 95 / 140 | — | arquivo inexistente, pasta fechada, impressão digital |
| A3 Faixa do Excel | 35 / 45 / 60 | — | XML + testes de faixa |
| A4 `.ppam` + build + corredor + ponta a ponta (+ contrato de API, carimbo de versão) | 110 / 185 / 300 | 5 (habilitar VBOM) | maior risco de descoberta (build pelo PowerPoint, eventos) |
| A5 Instalador/desinstalador (+ `PpamVersion`) | 50 / 75 / 105 | — | PS 7 e 5.1 |
| A6 Release, manifesto, publish | 50 / 75 / 110 | — | 6 assets, 3 campos do suplemento |
| A7 Atualização automática do suplemento pelo Excel (emenda) | 70 / 110 / 160 | — | atualizador + formulário + 8 testes; 1–2 rodadas |
| A8 Documentação | 50 / 75 / 105 | — | Test-Docs |
| A9 Suíte final + revisão do ramo + 1 onda de correção | 45 / 80 / 120 | 105 / 180 / 285 | humano: verificação manual §4.7 e §4.9, PR, `publicar-release` |
| **Parte A** | **555 / 875 / 1305 min (~9,25 / 14,6 / 21,75 h; corrido ~8,5 / 13 / 19,5 h)** | **110 / 185 / 290 min + revisão do PR e release ~0,5–1,5 h ⇒ 2,25 / 4 / 6,25 h** | confiança **média-baixa** |
| *Emenda (atualização automática do suplemento)* | *+75 / +120 / +180 min (corrido +1,5 / +2 / +2,5 h): A7 nova +70/+110/+160; A2 +5/+5/+10; A4 +10/+15/+20; A5 +5/+10/+15; A8 +5/+10/+15; A6 −20/−30/−40 (o atualizador saiu dela)* | *+15 / +30 / +45 min* | |
| B0 Matriz `delta` | 35 / 45 / 60 | — | |
| B1 `modDelta` (resolução, extensões, textos) | 55 / 85 / 120 | — | |
| B2 Desenho, faixas, espaço, refresh | 95 / 150 / 230 | — | geometria; 2 rodadas prováveis |
| B3 Modo de dois cliques (eventos) | 80 / 130 / 200 | — | eventos reais; risco médio |
| B4 Alternar texto + faixa | 45 / 65 / 85 | — | |
| B5 Exemplos, demonstração, galeria, docs | 75 / 105 / 150 | — | regera xlsx e galeria |
| B6 Suíte final + revisão do ramo + 1 onda | 45 / 80 / 120 | 60 / 100 / 160 | humano: cliques reais nos 9 tipos, Esc, zoom/DPI, PR, release |
| **Parte B** | **430 / 660 / 965 min (~7 / 11 / 16 h; corrido ~6 / 9,5 / 14 h)** | **1,5 / 2,5 / 4 h** (inclui PR e release) | confiança **média-alta** |
| **Total** | **corrido ~14,5 / 22,5 / 33,5 h ⇒ 2 / 3 / 4–5 dias de trabalho agêntico** (com a cadência da cascata: ~1 dia por parte no otimista) | **3,75 / 6,5 / 10,25 h** | |

### 12.2 Incertezas que mais pesam
- **Parte A:** montar o `.ppam` por automação do PowerPoint (nunca feito neste projeto; o PowerPoint é de instância única e não esconde a
  janela do aplicativo); comportamento dos eventos ao abrir pela interface e no Modo de Exibição Protegido (só manual); o aviso de segurança de
  vínculos na interface; host com pouca memória com Excel + PowerPoint + Excel `-Embedding` ao mesmo tempo; se o suplemento carregado prende o
  próprio arquivo (não medido: decide se a troca acontece na hora ou fica pendente — os dois caminhos estão cobertos).
- **Parte B:** se o `MouseUp` chega no clique que ativa o gráfico (só manual) e o `OnKey` com o gráfico ativo; DPI acima de 100% (a conversão
  pt → px só aparece nos testes; a produção usa as coordenadas do próprio evento).

## 13. Fontes
- think-cell — Manage data links: https://www.think-cell.com/en/resources/manual/manage-data-links
- think-cell — Excel data links: https://www.think-cell.com/en/resources/manual/exceldatalinks
- think-cell — Stay in control when managing Excel links: https://www.think-cell.com/en/blog/stay-in-control-when-managing-excel-links-with-think-cell
- think-cell — Chart decorations (difference arrows, CAGR): https://www.think-cell.com/en/resources/manual/chartdecorations
- Microsoft — `LinkFormat.AutoUpdate`: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.linkformat.autoupdate
- Microsoft — `Presentation.UpdateLinks`: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.updatelinks
- Microsoft — `Application.AfterPresentationOpen`: https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Application.AfterPresentationOpen
- Microsoft — `Application.ProtectedViewWindowOpen`: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.protectedviewwindowopen
- Microsoft — `Shape.ApplyAnimation`: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.applyanimation
- Microsoft — `Chart.Select` (evento), `Chart.MouseUp`, `Chart.GetChartElement`: https://learn.microsoft.com/en-us/office/vba/api/excel.chart.select(even) ,
  https://learn.microsoft.com/en-us/office/vba/api/excel.chart.mouseup , https://learn.microsoft.com/en-us/office/vba/api/excel.chart.getchartelement
- Microsoft Q&A — PowerPoint não atualiza gráficos vinculados (Automático volta a manual no OneDrive/SharePoint): https://learn.microsoft.com/en-us/answers/questions/5104591/powerpoint-doesnt-update-linked-excel-charts
- PPT FAQ — instalar suplemento pelo registro: https://pptfaq.com/FAQ00469_Add_-_install_an_addin_via_the_registry.htm
