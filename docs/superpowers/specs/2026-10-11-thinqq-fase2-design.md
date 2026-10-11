# ThinQQ — Pesquisa e Design (Fase 2)

Data: 2026-10-11
Autor: Israel Dias (com apoio do Claude Code)
Status: adendo à especificação da Fase 1 (`2026-10-10-thinqq-design.md`, que continua sendo a autoridade para
arquitetura, distribuição e atualização). Este documento cobre só o que a Fase 2 acrescenta ou muda.
Plano de execução: `docs/superpowers/plans/2026-10-11-thinqq-fase2.md`.

## 1. Objetivo e escopo

Entregar o ThinQQ **1.1.0** (core e loader) com os tipos e recursos listados na tabela de fases da spec (§2.1 e §8)
mais os carry-overs da Fase 1, mantendo o princípio central: **todo gráfico é um gráfico nativo do Excel ligado por
fórmulas a um bloco da planilha `_ThinQQ`**, de modo que quem recebe o arquivo sem o ThinQQ vê o gráfico correto e vivo.

| Item | Fase 2 | Observação |
|---|---|---|
| Tarefa 0 (carry-over) | Correção residual de `CalcSheet_Cleanup` | rótulos ligados (`DataLabel.Formula`) contam como referência; leitura falha => nada é apagado |
| Barras horizontais | empilhadas, agrupadas, 100% | mesma tabela das colunas |
| Cascata horizontal | sim | mesmo bloco da cascata vertical |
| Cascata com várias séries | sim (até 6 séries) | segmentos por série, coloridos por série, com legenda |
| Cascata "de cima para baixo" | sim | `e` na primeira linha + número na última linha (cálculo para trás, como o think-cell) |
| Mekko | variante % (eixo percentual) e variante unidades (larguras explícitas) | área empilhada sobre eixo de datas |
| Gantt | atividades, barras, marcos, escala de datas (semana/mês/trimestre/ano), colchetes, rótulos de data | barras empilhadas + dispersão |
| Pizza e rosca | rótulos dentro/fora, valor ou participação | rótulos nativos |
| Área | empilhada e 100% | rótulos por série auxiliar |
| Combinação | colunas empilhadas + linha (última série), eixo secundário opcional | |
| Linhas de valor | linha horizontal (ou vertical, em barras) na média, atualiza com os dados | série auxiliar ligada ao bloco |
| Indicador 100% | em gráficos 100% e Mekko (%) | forma reposicionada pelo ThinQQ |
| Mesma escala | entre gráficos selecionados (ou todos da planilha) | escala fixa persistida; "Escala automática" desfaz |
| Rótulos de série no fim da linha | alternativa à legenda em linha e combinação | opção "Legenda" |
| Faixa de opções | menus em vez de 17 botões; controles habilitados por capacidade do gráfico | `menu` do customUI14 |
| Ícones próprios na faixa | **adiado para a Fase 3** | `imageMso` validado por `GetImageMso` continua sendo suficiente |

Fora do escopo (Fase 3): borboleta, dispersão/bolhas, `TQQ.ROUND`, copiar para PowerPoint, assinatura digital, quebra de
eixo, e os itens listados em §10.

## 2. O que o think-cell faz (pesquisa, 2026-10-11)

Fontes: manual oficial (URLs em §11). Resumo do que importa para o desenho:

### 2.1 Mekko (manual "Mekko chart")
- Duas variantes: **eixo percentual** (alturas relativas a 100%; larguras = totais das colunas, então a área de cada
  segmento é proporcional ao valor absoluto; linha opcional `100%=` para informar o que vale 100%) e **unidades**
  (eixo e dados absolutos; larguras vêm de uma linha "X extent" no topo da planilha de dados).
- **Sem espaço entre colunas** (colunas encostadas), separadas visualmente por bordas; "ridge" opcional.
- Rótulos podem mostrar absoluto, percentual ou ambos; série "Outros"; ordem de segmentos configurável.
- Eixo da base é um eixo de valores (marcas e títulos opcionais).

### 2.2 Cascata (manual "Waterfall chart")
- **Várias séries por coluna**: positivos sobem a partir do nível acumulado, negativos descem; o conector segue a soma
  (5 e -2 => conector sobe 3). Cada segmento é desenhado com a sua extensão correta.
- `e` em **no máximo um segmento** da coluna; `e` na **primeira coluna** é permitido: o think-cell parte da última
  coluna e calcula para trás.
- **Build-up × build-down**: "de resto, são idênticos"; o ícone invertido só preenche a planilha com valores que
  descem a partir de um total. Isto é, build-down = começar pelo total e subtrair.
- Rótulos de segmento mostram a extensão (sempre positiva); sinal via formato numérico. Conectores são parte do cálculo.

### 2.3 Decorações (manual "Chart annotations")
- **Linhas de valor**: criadas na **média aritmética dos totais** (colunas empilhadas/agrupadas) ou na **média
  ponderada** (Mekko); se "encaixadas" na média, atualizam sozinhas; rótulo à esquerda ou à direita, com formato.
- **Indicador 100%**: em colunas/barras/áreas 100% e Mekko com eixo percentual; aparece por padrão nos 100%;
  à esquerda ou à direita.
- Setas de diferença (nível e total) e CAGR por totais; CAGR por série como rótulo de série; conectores de série.

### 2.4 Eixos (manual "Axes")
- **Mesma escala**: seleciona-se um elemento em cada gráfico e "Set Same Scale"; as escalas menores crescem até a
  maior; "Reset to Independent Scale" desfaz; se um gráfico muda de dados, todos na mesma escala acompanham.
- Em barras, "inverter" vira o gráfico todo; eixo de datas em categorias estritamente crescentes; quebras de eixo.

### 2.5 Rótulos (manual "Labels")
- Tipos por gráfico: categoria, série, segmento, ponto, total, total geral (coluna, barra, Mekko %, combinação),
  **dentro/fora** (pizza), atividade/responsável/observação/item/escala (Gantt), **indicador percentual** (100%).
- Pizza: rótulos internos o mais perto possível da borda; rótulos externos movíveis ao longo da borda.
- Rótulos de série colocados automaticamente; a legenda é um dos locais alternativos.

### 2.6 Gantt (manual "Gantt chart (timeline)")
- Escalas de **ano, trimestre, mês, semana e dia**, automáticas conforme o intervalo e o tamanho; separadores
  verticais (sólidos nas unidades maiores, tracejados nas menores); fins de semana sombreados (opcional).
- Linhas = atividades, com até três colunas de rótulo (Atividade, Responsável, Observação); linha especial de marcos.
- Itens: **barras**, setas de processo, **marcos** (losango, um dia), **colchetes** (rotulam um período, ancorados
  nas datas de início e fim), sombras. Rótulos de data (intervalo, início ou fim) atualizam sozinhos.
- Via Excel: selecionar as células (com uma linha extra para marcos) e "Gantt/Timeline"; dados: nome + datas por
  linha; fim sem hora = meia-noite (exclui o dia).

### 2.7 Pizza e rosca (manual "Pie and doughnut chart")
- Pizza = coluna 100% com uma categoria; rosca = pizza com furo (0% vira pizza). Rotação encaixa em 0/90/180/270;
  fatias destacadas; ordem das fatias (planilha, crescente, decrescente); conteúdo: absoluto, % ou ambos.

## 3. Técnicas nativas do Excel: alternativas e decisões

Critérios usados em todas as decisões: (1) o gráfico continua correto sem o ThinQQ (dados em fórmulas, não em
formas); (2) cabe na arquitetura atual (bloco em `_ThinQQ`, séries auxiliares `tqq_*`, rótulos ligados por
`DataLabel.Formula`, formas `tqq_ann_*` só para o que não é dado); (3) Excel 2016+ 32/64 bits, sem `Declare`;
(4) testável em Excel invisível (`PlotArea.Inside*`, `DataLabel.Left/Top`, `DataLabel.Text` funcionam lá — lição da
Fase 1); (5) pt-BR: formatos de rótulo via `NumberFormatLocal` (Style_LocalizeCode), formatos de **célula** com
`Range.NumberFormat` em inglês (inclusive datas: `dd/mm`, `mmm/yy`), fórmulas via `.Formula` em inglês.

### 3.1 Séries auxiliares em gráficos horizontais (barras, cascata horizontal, Gantt)
- Problema: em gráficos de **barras** o Excel não permite combinar com **linha** (as séries `tqq_total`/`tqq_lbl`
  da Fase 1 são linhas invisíveis). A combinação permitida é **barras + dispersão (XY)**. Peltier confirma que, em
  barras, a série XY precisa do **grupo secundário** (no grupo primário só funciona em colunas/linhas/áreas).
- Decisão: `Chart_AddHelperSeries(cht, name, valueCells, posCells)`: em gráficos verticais cria a linha invisível
  (como hoje); em horizontais cria uma série **XY no grupo secundário** com `X = valor` (posição no eixo de valores
  das barras) e `Y = posição da categoria` = `n - i + 0,5` (coluna `Pos` do bloco; eixo Y secundário fixo em
  `0..n`, oculto, não invertido). O eixo de categorias primário fica **invertido** (`ReversePlotOrder`, primeira
  categoria no topo, como o think-cell).
- O eixo X secundário normalmente não é exibido e acompanha o eixo de valores primário; se o Excel exigir alinhamento
  explícito, `Chart_SyncSecondaryX` copia min/max do primário (chamado no refresh). A tarefa das barras começa por um
  **teste-sonda** que mede o alinhamento real (rótulo do ponto i dentro da faixa i e no fim da barra, por regra de
  três no `PlotArea`/eixo) e decide entre os dois caminhos.
- Rótulos nesses auxiliares: `xlLabelPositionRight/Left` (sinal) em vez de Above/Below.

### 3.2 Mekko
Alternativas: (a) colunas empilhadas — largura por categoria não existe no Excel; (b) **área empilhada em eixo de
datas** (Peltier, "Marimekko charts in Excel"): pontos nas bordas de cada coluna, repetidos, formam blocos
retangulares; (c) área em eixo de texto com uma categoria por unidade de largura (centenas de linhas por série);
(d) formas desenhadas — viola o critério 1.
- Decisão: **(b)**. Eixo de categorias `xlTimeScale` (base dias) com X inteiro de 1 a 1001 (resolução 0,1 %; o
  eixo de datas **trunca frações**, por isso inteiros), escala fixa `1..1001`, sem rótulos. Cinco linhas por coluna no
  bloco: `L0 (x=L, y=0)`, `L (x=L, y=v)`, `C (x=centro, y=v)`, `R (x=R, y=v)`, `R0 (x=R, y=0)`: a descida a zero com
  largura nula, somada à borda branca das áreas, desenha o separador vertical entre colunas (sem espaço, como o
  think-cell).
- Rótulos: séries auxiliares de **linha** (invisíveis) no mesmo eixo de datas, com valor só nas linhas `C` (`NA()` nas
  demais): rótulo de segmento no meio da faixa, total/largura no topo (`tqq_total`), nome da categoria em y=0
  (`tqq_cat`, posição Abaixo). Tudo ligado a células.
- Variantes: **%**: largura = total da linha / total geral; altura = participação (eixo 0..1). **Unidades**: a
  **segunda coluna da tabela é a largura** (cabeçalho livre, ex.: "Largura"), demais colunas são as séries; altura
  = valor absoluto.
- Linha de valor (unidades): série de linha constante na média ponderada atravessa todo o eixo (de 1 a 1001).

### 3.3 Gantt
Alternativas: (a) **barras empilhadas** com série inicial invisível (`Start`) e duração visível + marcos em XY
(técnica padrão); (b) formas desenhadas; (c) dispersão pura com barras de erro.
- Decisão: **(a)** com os auxiliares de §3.1. `tqq_start` (invisível) + `tqq_dur` (duração inclusiva: `Fim - Início + 1`;
  **fim vazio = marco**). Marcos: XY `tqq_ms` (losango). Rótulos: XY `tqq_lbl` à direita (data do fim/marco em célula
  formatada `dd/mm`, ou o texto da 4.ª coluna da tabela, se existir).
- **Escala de datas** sem depender de eixo de datas (o eixo de valores de barras é numérico): o bloco calcula por
  fórmula a unidade (`w` semanas se o intervalo ≤ 98 dias; `m` meses ≤ 1100; `q` trimestres ≤ 3700; `y` anos), o
  `AxisMin`/`AxisMax` alinhados à unidade e até **40 marcas** (`EDATE`/`+7`); uma série XY `tqq_scale` no topo mostra
  os rótulos (células com `NumberFormat` `dd/mm` ou `mmm/yy`; trimestre/ano como texto `T1/2026`, `2026`) e desenha os
  separadores verticais com **barras de erro** (`ErrorBar Direction:=xlY, Include:=xlErrorBarIncludeMinusValues,
  Type:=xlErrorBarTypeFixedValue, Amount:=n`). O min/max do eixo de valores é aplicado por VBA a partir das células
  (`Gantt_RefreshScale`, no insert e em cada refresh); sem o ThinQQ, barras/marcos/rótulos continuam vivos e só o
  intervalo do eixo fica como estava.
- **Colchetes**: anotação (`kind=bracket;from=i;to=j`) com linha + dois tiques + rótulo `dd/mm–dd/mm`, posicionada
  na área da escala (acima da primeira linha), por `Geom_ChartValueToX` das datas mínima/máxima das atividades `i..j`.
- Fora da Fase 2: sombreamento de fins de semana, colunas Responsável/Observação com colchetes automáticos,
  separadores de linha, âncoras, setas de processo, linha especial de marcos.

### 3.4 Cascata com várias séries (algoritmo de empilhamento)
O Excel empilha positivos para cima **na ordem das séries** e negativos para baixo **na mesma ordem**, a partir de
zero. Numa coluna com início `s`, a série `j` ocupa `[A_j, B_j]` (acima de `s` se `D_j ≥ 0`, abaixo se `< 0`); de
baixo para cima a ordem é: negativos em ordem inversa de série, depois positivos em ordem de série. Como cada parte
pode cruzar o zero, cada segmento vira **duas** séries de gráfico: parte acima de zero (`UP_j` se `D_j ≥ 0`, `DN_j`
se `< 0`, valores ≥ 0) e parte abaixo de zero (`UPN_j`/`DNN_j`, valores ≤ 0). Ordem das séries no gráfico:
`tqq_base`, `DN_S..DN_1`, `UP_1..UP_S` (pilha positiva, de baixo para cima), `UPN_S..UPN_1`, `DNN_1..DNN_S` (pilha
negativa, de cima para baixo), `tqq_total`, auxiliares de rótulo. `UP_j` leva o nome da série (legenda); o resto é
`tqq_*` (fora da legenda). Fórmulas em §4.3. Com uma série o bloco e o gráfico são **idênticos aos da Fase 1**.

### 3.5 Linhas de valor
Alternativas: forma (`Chart.Shapes.AddLine`) — não vive sem o ThinQQ; série XY no grupo primário — alinha só em
colunas; **série de linha** com valor constante — simples, viva, atravessa da primeira à última categoria (em eixo
"entre categorias" fica meia fatia aquém das bordas: aceito e documentado).
- Decisão: coluna `VLine` no bloco (`=AVERAGE(Total)`; Mekko unidades: média ponderada pela largura) + série
  `tqq_vline` (linha cinza tracejada, sem marcadores) com rótulo ligado no último ponto (célula com formato
  `"Média "#,##0`, localizado por `Style_LocalizeCode`). Em barras: XY secundária com 2 pontos (`Y = 0` e `Y = n`).

### 3.6 Combinação e eixo secundário
`Series.ChartType = xlLineMarkers` na última série de um gráfico de colunas empilhadas; `Series.AxisGroup =
xlSecondary` quando a opção `secondary` está ligada (padrão: ligada quando as magnitudes diferem por fator > 5).
Eixo secundário oculto como o primário. Totais só das colunas.

### 3.7 Pizza e rosca
Rótulos **nativos** (`ShowValue`/`ShowPercentage`, `NumberFormatLocal`), `xlLabelPositionInsideEnd` (dentro, perto da
borda, como o think-cell) ou `xlLabelPositionOutsideEnd` + nome da categoria + linhas de chamada (fora);
`VaryByCategories` com cores da paleta por ponto; `FirstSliceAngle = 0`; `DoughnutHoleSize = 50`. Uma série
(primeira coluna de valores). Sem anotações.

### 3.8 Faixa de opções
customUI14 permite `menu` com `button`/`toggleButton`/`menuSeparator` filhos; as assinaturas dos callbacks são as
mesmas dos controles soltos (`button`: `Sub(control)`; `toggleButton`: `Sub(control, pressed As Boolean)`;
`getPressed`: `Sub(control, ByRef returnValue)`; `gallery.onAction`: `Sub(control, selectedId As String,
selectedIndex As Integer)`, igual ao `dropDown`). Decisão: `menu` (não `gallery`, que precisaria de imagens por item)
para os tipos e para "Mais opções"/"Setas"; `getEnabled` por **capacidade** (`Core_Supports(tag)`), não só por
"há gráfico".

## 4. Decisões por tipo

Identificadores de tipo (`Registry_ChartType`): Fase 1 `stacked, clustered, stacked100, line, waterfall`; Fase 2
`barstacked, barclustered, barstacked100, barwaterfall, area, area100, combo, mekko, mekkounits, gantt, pie, doughnut`.
Famílias (`Kinds_Family`): `stacked` (colunas/barras/linha), `waterfall`, `area`, `combo`, `mekko`, `gantt`, `pie`.

### 4.1 Tabela-resumo

| Tipo | Tabela do usuário | Gráfico nativo | Rótulos | Totais | Conectores | Anotações | Decorações |
|---|---|---|---|---|---|---|---|
| barstacked / barclustered / barstacked100 | igual às colunas | `xlBarStacked/Clustered/Stacked100`, categorias invertidas | nativos nos segmentos; totais via XY secundária | empilhadas e 100% | linhas de série (empilhadas) | CAGR, diferença (não em 100%) | vline (não em 100%), ind100 (100%), mesma escala |
| waterfall (1..6 séries) / barwaterfall | Etapa + 1..6 colunas; `e` marca total | colunas/barras empilhadas (§3.4) | segmento (delta assinado) e total | n/a | conectores da cascata (formas) | CAGR, diferença | mesma escala |
| area / area100 | igual às colunas | `xlAreaStacked/Stacked100` | auxiliares de linha no meio da faixa | sim | n/a | CAGR, diferença (não em 100%) | vline (area), ind100 (area100), mesma escala |
| combo | igual às colunas; **última coluna = linha** | colunas empilhadas + `xlLineMarkers` | segmentos; pontos da linha | das colunas | linhas de série | CAGR, diferença (totais das colunas) | vline, eixo secundário, mesma escala, fim da linha |
| mekko | categorias nas linhas, séries nas colunas | área empilhada em eixo de datas (§3.2) | participação no segmento; largura (%) no topo; categorias em y=0 | sim (largura %) | n/a | nenhuma (recusa 7001) | ind100 |
| mekkounits | 2.ª coluna = largura | idem; eixo em unidades | valor no segmento; total no topo | sim (total) | n/a | nenhuma | vline (média ponderada) |
| gantt | Atividade, Início, Fim [, Rótulo] | barras empilhadas + XY (§3.3) | data do fim/marco ou 4.ª coluna | n/a | n/a | colchete | gap (espaço entre linhas) |
| pie / doughnut | Categoria + Valor | `xlPie/xlDoughnut` | nativos, dentro/fora; valor ou % | n/a | n/a | nenhuma | rótulos fora |

Opções persistidas (`Registry_SetOption`, chave=valor): as da Fase 1 (`numfmt, palette, gap, totals, labels,
connectors`) mais `legend` (`right|none|lineend`), `ind100`, `vline`, `secondary`, `pieout` (`0|1`), `vmin`/`vmax`
(escala fixa, texto neutro com ponto decimal; vazio = automática) e `n` (número de categorias, usado por
`Core_ChartSourceMismatch` nos tipos cujo bloco não tem uma linha por categoria).

### 4.2 Bloco da família `stacked` (colunas, barras, linha, combo)
Corpo: cabeçalho + 1 linha por categoria. Colunas: `1 Categoria | 2 Total | 3 TotalPos | [4..3+S Part_j, só 100%] |
P = Pos | P+1 = VLine`, onde `Pos = n - i + 0,5` e `VLine = AVERAGE(Total)` (`NA()` em 100%). Em `combo` o total
soma só as séries de coluna (todas menos a última). As colunas 1..3 (e `Part_j`) são exatamente as da Fase 1.

### 4.3 Bloco da cascata (uma ou várias séries, vertical ou horizontal)
Colunas 1..18 como na Fase 1 (`Cat, Src, IsTotal, Delta, Start, End, Lo, Hi, Base, PiecePos, PieceNeg, UpPos, UpNeg,
DownPos, DownNeg, Total, LabelVal, LabelPos`), mais `19 Pos`, `20 BD` (build-down: `=AND(marcador(linha 1),
ISNUMBER(valor da última linha))`, igual em todas as linhas) e, **só com S > 1**, 10 colunas por série `j` a partir da
coluna `21 + 10(j-1)`: `D_j, CP_j, CN_j, A_j, B_j, M_j, UP_j, DN_j, UPN_j, DNN_j`:

```
D_j  = IF(IsTotal, 0, IF(ISNUMBER(src_j), src_j, 0))
CP_j = CP_{j-1} + MAX(0, D_j)           CN_j = CN_{j-1} + MAX(0, -D_j)      (CP_0 = CN_0 = 0)
A_j  = IF(D_j >= 0, Start + CP_{j-1}, Start - CN_j)
B_j  = IF(D_j >= 0, Start + CP_j,     Start - CN_{j-1})
M_j  = (A_j + B_j) / 2
UP_j = IF(D_j >= 0, MAX(0,B_j) - MAX(0,A_j), 0)      DN_j  = IF(D_j < 0, MAX(0,B_j) - MAX(0,A_j), 0)
UPN_j = IF(D_j >= 0, MIN(0,A_j) - MIN(0,B_j), 0)     DNN_j = IF(D_j < 0, MIN(0,A_j) - MIN(0,B_j), 0)
Delta (col 4) = SUM(D_1..D_S)     Lo = IF(IsTotal, MIN(0,End), Start - CN_S)     Hi = IF(IsTotal, MAX(0,End), Start + CP_S)
IsTotal (S>1) = SUMPRODUCT(--ISNUMBER(MATCH(LOWER(TRIM(IFERROR(src_1:src_S,"")&"")), {"e","t","=","total","subtotal"}, 0))) > 0
```
Com S = 1, `Lo/Hi` ficam `MIN/MAX(Start, End)` (equivalente). **Build-down** (`BD` verdadeiro): `End_1 = src_n -
SUM(Delta_2..Delta_{n-1})`, a última linha vira total com `Total_n = src_n` (`IsTotal_n` forçado), `Start_1 = 0`.
Com várias séries, `LabelVal/LabelPos` valem só nas linhas de total (`NA()` nas demais) e cada série ganha um
auxiliar de rótulo `tqq_lbl_j` em `M_j` ligado a `D_j`.

### 4.4 Bloco da área
`1 Categoria | 2 Total | 3 TotalPos (Total; 1 em area100) | 4..3+S Val_j | 4+S..3+2S Mid_j | Pos | VLine`;
`Val_j = IF(ISNUMBER(src), src, 0)` (area100: participação) e `Mid_j = SUM(Val_1..Val_j) - Val_j/2`.

### 4.5 Bloco do Mekko
Corpo com `5n` linhas (`L0, L, C, R, R0` por categoria): `1 X | 2 Categoria (só em C) | 3 W (largura, só em C) |
4 Total (só em C) | 5..4+S Y_j (área) | 5+S..4+2S Lbl_j (NA fora de C; em C = meio da faixa) | 5+2S..4+3S LblVal_j
(valor ou participação em C) | 5+3S TotPos (1 ou total em C) | 6+3S TotLbl | 7+3S CatPos (0 em C) | 8+3S VLine (unidades)`.
`X(L0)=X(L)=1+ROUND(1000*cum_{i-1})`, `X(C)=1+ROUND(1000*(cum_{i-1}+cum_i)/2)`, `X(R)=X(R0)=1+ROUND(1000*cum_i)`;
`cum_i = SUM(W_1..W_i)/SUM(W)`; `W_i` = total da linha (%) ou coluna de largura (unidades). Linha com total zero: `W = 0`
(coluna de largura nula, sem erro). Opção `n` = número de categorias; `Core_ChartSourceMismatch` espera `5n` células.

### 4.6 Bloco do Gantt
Corpo com `MAX(n, 40)` linhas e 19 colunas. Atividades (linhas 2..n+1): `1 Atividade | 2 Inicio | 3 Fim | 4 IsMs |
5 BarStart | 6 Dur | 7 MsX | 8 Pos | 9 LblX | 10 LblTxt` (11 vazia). Escala (linhas 2..41, colunas 12..17): `12 Unit |
13 AxisMin | 14 AxisMax` (só na linha 2) `| 15 Tick | 16 TickY | 17 TickLbl`; `18 DMin | 19 DMax` (só na linha 2):
```
Inicio = IF(ISNUMBER(src2), src2, NA())      Fim = IF(ISNUMBER(src3), src3, NA())
IsMs = AND(ISNUMBER(src2), NOT(ISNUMBER(src3)))
BarStart = IF(IsMs, NA(), Inicio)            Dur = IF(IsMs, NA(), MAX(0, Fim - Inicio + 1))
MsX = IF(IsMs, Inicio, NA())                 Pos = n - i + 0.5        LblX = IF(IsMs, Inicio, Fim)
dmin = AGGREGATE(5,6,Inicio)   dmax = AGGREGATE(4,6,Inicio:Fim)   span = dmax - dmin
Unit = IF(span<=98,"w",IF(span<=1100,"m",IF(span<=3700,"q","y")))
AxisMin = w: dmin-WEEKDAY(dmin,3) | m: DATE(YEAR(dmin),MONTH(dmin),1) | q: DATE(YEAR(dmin),3*INT((MONTH(dmin)-1)/3)+1,1) | y: DATE(YEAR(dmin),1,1)
AxisMax = w: dmax-WEEKDAY(dmax,3)+7 | m: EDATE(primeiro dia do mês de dmax,1) | q: EDATE(primeiro dia do trimestre,3) | y: DATE(YEAR(dmax)+1,1,1)
Tick_k (k=0..39) = IF(Unit="w", AxisMin+7k, EDATE(AxisMin, k*{m:1,q:3,y:12}));  = IF(Tick<AxisMax, Tick, NA())
TickY = IF(ISNUMBER(Tick), n, NA())          TickLbl = IF(Unit="q","T"&(INT((MONTH(Tick)-1)/3)+1)&"/"&YEAR(Tick), IF(Unit="y", YEAR(Tick), Tick))
```
`Gantt_RefreshScale` aplica `AxisMin/AxisMax` ao eixo de valores, o formato da coluna `TickLbl` (`dd/mm` para `w`,
`mmm/yy` para `m`; `General` para `q`/`y`) e o eixo Y secundário `0..n`.

### 4.7 Bloco da pizza
`1 Categoria | 2 Valor | 3 Share` (informativo; rótulos nativos). Primeira coluna de valores; valores não numéricos
ou negativos contam como 0 (`IF(AND(ISNUMBER(src),src>0),src,0)` alimenta a série, não a tabela do usuário).

## 5. Geometria generalizada e anotações

- `modGeometry` ganha funções puras espelhadas para o eixo horizontal (`Geom_CategoryCenterY`, `Geom_ValueToX`) e
  despachantes por orientação: `Geom_ChartIsHorizontal(cht)`, `Geom_ChartCatCoord(cht, idx)` (x nas colunas, y nas
  barras, categoria 1 no topo), `Geom_ChartValCoord(cht, v)` (y nas colunas, x nas barras), `Geom_ChartBarHalfWidth`
  usa `InsideHeight` nas barras. As funções da Fase 1 continuam (testes antigos intactos).
- `Ann_LayoutCagr/Diff/Conn` passam a trabalhar em coordenadas `(cat, val)` e trocam os eixos em horizontais:
  CAGR em barras = seta ao longo das categorias à direita dos fins das barras, rótulo à direita; diferença em barras =
  seta dupla horizontal abaixo da última barra, linha de nível vertical a partir do fim da barra inicial; conectores
  da cascata horizontal = linhas verticais em `x = End_i`.
- `Ann_GetCategoryData`: `area/area100` (soma; topo 1 em 100%), `combo` (soma das colunas, sem a linha), barras
  (como as colunas). Mekko, Gantt e pizza não têm setas: `Ann_AddCagr/AddDifference` recusam com erro 7001 e mensagem
  por tipo; `Core_AddAnnotation` mostra a mensagem sem registrar falha (como hoje).
- Novas anotações: `ind100` (tique + "100%" no nível 1; esquerda nas colunas, topo nas barras) e `bracket` (Gantt).
  `Ann_RemoveAll` continua removendo só setas e colchetes (conectores e `ind100` são alternáveis pelos botões).

## 6. Decorações

| Recurso | Opção | Implementação | Tipos |
|---|---|---|---|
| Linha de valor | `vline=1` (toggle "Linha de valor (média)") | série `tqq_vline` (§3.5) | stacked, clustered, line, barstacked, barclustered, area, combo, mekkounits |
| Indicador 100% | `ind100=1`, **ligado por padrão** nos 100% | forma `tqq_ann_ind100` + `_lbl`, posicionada em `Ann_Refresh` | stacked100, barstacked100, area100, mekko |
| Mesma escala | `vmin`/`vmax` | `Deco_SameScale(charts)`: lê a escala automática de cada gráfico, aplica a união a todos, persiste; `Deco_ApplyFixedScale` reaplica no refresh; "Escala automática" limpa | todos com eixo de valores livre (não 100%, mekko, pie, gantt) |
| Legenda | `legend=right|none|lineend` | `Deco_SetLegend`: `lineend` oculta a legenda e mostra o nome da série no rótulo do último ponto (`ShowSeriesName`) | lineend só em line e combo (linha); right/none em todos com legenda |
| Eixo secundário | `secondary=1` | `Combo_SetSecondary` | combo |
| Rótulos fora | `pieout=1` | `Pie_SetOutside` | pie, doughnut |

"Mesma escala" age sobre os gráficos ThinQQ **selecionados** (`Selection.ShapeRange`, 2 ou mais) ou, sem seleção
múltipla, sobre **todos os gráficos ThinQQ da planilha ativa**; o botão fica sempre habilitado (a seleção múltipla
de gráficos não dispara evento que invalide a faixa) e informa o que fez. A escala é fixa (não dinâmica como no
think-cell): quem mudar muito os dados clica de novo. Documentado no tutorial.

## 7. Faixa de opções (loader 1.1.0)

```
[ThinQQ]
┌ Inserir gráfico ─────────────────────────────────────────────┐ ┌ Gráfico selecionado ───────────────────────────────────────────────┐ ┌ ThinQQ ─────────┐
│ Colunas ▾   Barras ▾   Cascata ▾   Linha e área ▾   Mais tipos ▾ │ │ [x] Totais [x] Rótulos [ ] Conectores                                  │ │ Exemplos Tutorial│
│  empilhadas  empilhadas  (colunas)   Linha           Mekko (%)    │ │ Formato ▾ Paleta ▾ Espaço ▾ Legenda ▾                                   │ │ Verificar atual. │
│  agrupadas   agrupadas   (barras)    Área empilhada  Mekko (unid.)│ │ Mais opções ▾ (Indicador 100%, Linha de valor, Eixo secundário,         │ │ Configurações    │
│  100%        100%                    Área 100%       Gantt        │ │   Rótulos fora da pizza | Mesma escala, Escala automática)               │ │ Sobre            │
│                                      Combinação      Pizza, Rosca │ │ Setas ▾ (Seta CAGR, Seta diferença, Colchete (Gantt) | Remover setas)   │ └──────────────────┘
└──────────────────────────────────────────────────────────────┘ │ Atualizar   Alterar dados                                               │
                                                                  └─────────────────────────────────────────────────────────────────────────┘
```
- `Rx_GetEnabled`: controles com `tag` consultam `Core_Supports(tag)` (cache por ciclo de `Invalidate`); sem `tag`
  (Remover setas, Alterar dados) usam `Core_HasChart`; `Atualizar`, `Mesma escala` e `Escala automática` ficam sempre
  habilitados. Core antigo (sem `Core_Supports`): `Loader_RunCore` devolve `Empty` → cai em `Core_HasChart`.
- Matriz de capacidades (`Kinds_Supports(kind, feature)`), fonte única no core, usada pela faixa, por
  `Core_SetToggle/SetOption` (mensagem quando não se aplica) e por `Core_RebuildFromRange`:

| feature | tipos |
|---|---|
| totals | stacked, stacked100, barstacked, barstacked100, area, area100, combo, mekko, mekkounits |
| labels | todos |
| connectors | stacked, stacked100, barstacked, barstacked100, waterfall, barwaterfall, combo |
| numfmt | todos exceto gantt |
| palette | todos |
| gap | stacked, clustered, stacked100, barstacked, barclustered, barstacked100, waterfall, barwaterfall, combo, gantt |
| legend | todos exceto gantt (em waterfall só com várias séries) |
| ind100 | stacked100, barstacked100, area100, mekko |
| vline | stacked, clustered, line, barstacked, barclustered, area, combo, mekkounits |
| secondary | combo |
| pieout | pie, doughnut |
| cagr | stacked, clustered, stacked100, line, waterfall, barstacked, barclustered, barstacked100, barwaterfall, area, area100, combo |
| diff | stacked, clustered, line, waterfall, barstacked, barclustered, barwaterfall, area, combo |
| bracket | gantt |
| samescale | stacked, clustered, line, waterfall, barstacked, barclustered, barwaterfall, area, combo, mekkounits |

Os testes da faixa (`TestRibbon`) passam a validar `menu` (callbacks só nos filhos; `getEnabled` opcional no menu),
17 botões de inserção com `tag` em `Kinds_IsValid`, 7 toggles, 4 dropdowns, e `imageMso` válidos por
`GetImageMso`. `Test-Docs` passa a exigir no tutorial os rótulos de `menu` além de `button/toggleButton/dropDown`.

## 8. Versões, compatibilidade e documentação

- `version.json`: `{"loader":"1.1.0","core":"1.1.0"}`. A Fase 1 (1.0.0) **não foi publicada**; o CHANGELOG ganha a
  seção `[1.1.0]` com tudo da seção "Não lançado" (correções pós-1.0.0) + o conteúdo da Fase 2; `release/notes.txt`
  resume as novidades (vai para `notes` do manifesto). Loader e core sobem juntos porque a faixa mudou; loader 1.1.0
  com core 1.0.0 funciona (botões novos avisam "Tipo de gráfico desconhecido"/caem em `Core_HasChart`); core 1.1.0
  com loader 1.0.0 funciona com os botões antigos.
- Atualizações por recurso:

| Recurso | CHANGELOG | Tutorial (seção) | README | Exemplos | Test-Docs/Test-Ribbon |
|---|---|---|---|---|---|
| Tarefa 0 | Corrigido | "Copiei ou movi um gráfico" (FAQ) | Limitações (remover a ressalva) | — | — |
| Barras + cascata horizontal | Adicionado | "Barras horizontais" | Recursos | 7, 9 | rótulos do menu Barras/Cascata |
| Cascata várias séries / de cima para baixo | Adicionado | "Cascata" (ampliada) | Recursos, Limitações | 8, 2 (texto) | — |
| Área, combinação, legenda | Adicionado | "Linha, área e combinação" | Recursos | 15, 16 | menu Linha e área, dropdown Legenda |
| Linha de valor, indicador 100%, mesma escala | Adicionado | "Decorações" | Recursos | 17, 18, 5 | menu Mais opções |
| Pizza e rosca | Adicionado | "Pizza e rosca" | Recursos | 13, 14 | menu Mais tipos |
| Mekko | Adicionado | "Mekko" | Recursos | 10, 11 | menu Mais tipos |
| Gantt | Adicionado | "Gantt" | Recursos, Limitações | 12 | menu Mais tipos, Setas > Colchete |
| Faixa nova | Alterado | maquete da faixa + "Neste tutorial" | Verificações manuais | — | contagens |

## 9. Riscos e fallbacks

| Risco | Probabilidade | Mitigação / fallback |
|---|---|---|
| XY secundária em barras não alinha com o eixo de valores primário (rótulos/marcos deslocados) | média | teste-sonda na tarefa das barras; `Chart_SyncSecondaryX` no insert e no refresh (escala copiada do primário); em Gantt o primário já é fixo |
| Eixo de datas do Mekko com `MinimumScale = 1` recusado ou com arredondamento visível | baixa | X já é inteiro; se o Excel recusar 1, usar 2..1002; teste de geometria mede o centro real do rótulo |
| Rótulo de ponto `NA()` em série de linha (Mekko/cascata multi) gerar erro ao ligar | média | ligação por ponto só onde a célula não é erro, sob `On Error Resume Next` por ponto (`Mekko_LinkLabels`) |
| Muitas séries na cascata multi (4S+3+S) deixam o insert lento | baixa | limite de 6 séries (mensagem); medição no teste de integração |
| Barras de erro do `tqq_scale` (Gantt) com direção invertida no eixo Y secundário | média | eixo Y secundário não invertido (Y cresce para cima; `Minus` desce); checagem visual no roteiro manual; fallback: linhas de grade do eixo de valores (`MajorUnit = 7`) |
| `DataLabel.Formula` de rótulo não ligado levantar erro na leitura (Tarefa 0) | média | leitura por ponto: erro em `.Formula` E em `.Text` = falha de leitura (nada apagado); só `.Formula` = rótulo de texto |
| `Style_ApplyBase` em pizza (sem eixos) | certa | bloco de eixos protegido para tipos sem eixo |
| Faixa: `imageMso` novos inválidos | baixa | `Test_Ribbon_ImageMso_AreValid` (GetImageMso) decide; alternativas listadas no plano |
| Crescimento do tempo da suíte (> 4 min) | média | testes de tipo usam fixtures pequenas; `-Filter` na iteração; integração completa só na tarefa final |

## 10. Adiado para a Fase 3

Borboleta; dispersão/bolhas; `TQQ.ROUND`; copiar para PowerPoint (imagem/vinculado); assinatura digital; quebra de
eixo; ícones próprios na faixa (`getImage`); Gantt: sombreamento de fins de semana, colunas Responsável/Observação,
separadores de linha, âncoras, setas de processo, linha de marcos, escala de dias; rosca com vários anéis; série
"Outros"; linha de valor em valor arbitrário (só média); mesma escala dinâmica; setas em Mekko; rótulos no fim da
linha em áreas; ordem de segmentos/fatias; rótulos de CAGR por série; rótulo de intervalo `dd/mm–dd/mm` nas barras do
Gantt (só fim/marco); "100%=" do Mekko.

## 11. Fontes

- think-cell — Mekko: https://www.think-cell.com/en/resources/manual/mekko
- think-cell — cascata: https://www.think-cell.com/en/resources/manual/waterfall
- think-cell — decorações (linhas de valor, indicador 100%, setas): https://www.think-cell.com/en/resources/manual/chartdecorations
- think-cell — eixos (mesma escala, datas, quebras): https://www.think-cell.com/en/resources/manual/axes
- think-cell — rótulos: https://www.think-cell.com/en/resources/manual/textlabels
- think-cell — Gantt (timeline): https://www.think-cell.com/en/resources/manual/gantt
- think-cell — pizza e rosca: https://www.think-cell.com/en/resources/manual/pie
- think-cell — introdução a gráficos: https://www.think-cell.com/en/resources/manual/introductioncharting
- Peltier — Marimekko em Excel (área empilhada em eixo de datas; inteiros no eixo de datas): https://peltiertech.com/marimekko-charts/
- Peltier — barras + XY (XY precisa do grupo secundário em barras; Y = centro da barra): https://peltiertech.com/bar-line-xy-combo-chart-in-excel/
- Peltier — posicionamento de XY no eixo primário de colunas (categoria i em x = i): https://peltiertech.com/precision-positioning-of-xy-data-points/
- Peltier — linha horizontal em gráfico de colunas (XY no primário, 0,5..n+0,5): https://peltiertech.com/add-horizontal-line-to-excel-chart/
- Microsoft — `Series.AxisGroup`: https://learn.microsoft.com/en-us/office/vba/api/excel.series.axisgroup
- Microsoft — `Series.ErrorBar`: https://learn.microsoft.com/en-us/office/vba/api/excel.series.errorbar
- Microsoft — `XlDataLabelPosition`: https://learn.microsoft.com/en-us/office/vba/api/excel.xldatalabelposition
- Microsoft — `Axis.CategoryType` (xlTimeScale, BaseUnit): https://learn.microsoft.com/en-us/office/vba/api/excel.axis.categorytype
- Microsoft — customUI (Fluent Ribbon) parte 3, assinaturas de callbacks por controle (menu, gallery, toggleButton): https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)
