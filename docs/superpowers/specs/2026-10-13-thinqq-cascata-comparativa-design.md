# ThinQQ — Pesquisa e Design: Cascata comparativa (versão 1.3.0)

Data: 2026-10-13
Autor: Israel Dias (com apoio do Claude Code)
Status: adendo às especificações da Fase 1 (`docs/superpowers/specs/2026-10-10-thinqq-design.md`, arquitetura, distribuição e atualização),
da Fase 2 (`docs/superpowers/specs/2026-10-11-thinqq-fase2-design.md`, matriz de capacidades, blocos por família, faixa em menus) e da Fase 3
(`docs/superpowers/specs/2026-10-11-thinqq-fase3-design.md`, matriz 20 × 20, ícones, demonstração). Este documento cobre só o tipo novo.
Rascunho em `.superpowers/sdd/2026-10-13-cascata-comparativa/design.md`; a Tarefa 0 do plano o copia (sem mudar nada) para
`docs/superpowers/specs/2026-10-13-thinqq-cascata-comparativa-design.md`, que passa a ser a fonte que o teste da matriz lê (§6).
Plano: `.superpowers/sdd/2026-10-13-cascata-comparativa/plan.md` (a Tarefa 0 o copia para `docs/superpowers/plans/2026-10-13-thinqq-cascata-comparativa.md`).

## 1. Objetivo e escopo

O usuário pediu "o modelo de gráfico comparativo que existe no think-cell": uma **cascata (ponte) com duas séries lado a lado em cada etapa**
(ex.: Segmento 1 | Segmento 2: Bruto 27 | 29, Componente 1 −4 | −7, Componente 2 −6 | −2, Componente 3 −6 | −10, Líquido 11 | 10), em que cada
série tem a sua própria ponte (barra inicial, degraus flutuantes, total), conectores tracejados entre barras da mesma série, uma **seta em
cotovelo** entre as duas barras de cada etapa com a variação percentual numa elipse ("+7%", "+75%", "−67%", "+67%", "−9%"), totais com o
percentual do primeiro total da própria série ("11 (41%)", "10 (34%)") e os nomes das séries escritos na vertical dentro das barras da primeira
etapa (alternativa: legenda).

Entrega: ThinQQ **1.3.0** (core e loader; a faixa muda) com o tipo `waterfallpair` ("Cascata comparativa", menu Cascata), mantendo o princípio
central: **gráfico nativo do Excel ligado por fórmulas a um bloco da planilha muito oculta `_ThinQQ`**. Diferente das setas de CAGR/diferença
(formas estáticas sem o ThinQQ), **tudo neste tipo é série do gráfico ligada a células**: barras, conectores, setas de variação com o texto da
elipse, nomes das séries e rótulos acompanham os dados mesmo sem o ThinQQ (§3). Só o afastamento das setas em pontos, o "rótulo fino demais"
e a escolha automática "nome dentro da barra ou legenda" são recalibrados pelo ThinQQ (refresh); sem ele ficam como na última calibração.

| Item do brief | Decisão | Resumo |
|---|---|---|
| 1. Pesquisa think-cell | §2 | O think-cell monta este gráfico como uma cascata padrão com **lacunas de categoria** (category gaps) e cores por série, mais **setas de diferença** de nível; o próprio think-cell diz que a técnica só funciona bem com exatamente duas séries. |
| 2. Tipo novo | `waterfallpair`, família própria `waterfallpair` | Tabela `Etapa \| Série 1 \| Série 2`; recusa clara com 1 coluna (6016) e com 3 ou mais (6017); marcadores `e`/`t`/`=`/`total`/`subtotal` como na cascata; primeira linha numérica = barra inicial. |
| 3. Construção nativa | **(a) colunas empilhadas com categorias expandidas** (A, B, vão por etapa) + **eixo de categorias de 2 níveis** | Medido: o rótulo da etapa fica centrado sob o par (nível externo agrupa A e B); o vão é uma categoria vazia; cruza o zero com o empilhamento de base invisível da cascata atual. |
| 4. Conectores por série | **auxiliares XY com barra de erro X** (`ErrorBar` custom ligado ao bloco) | Vivos sem o ThinQQ; tracejados; uma série por lado. Atravessam a barra da outra série (como no think-cell; documentado). |
| 5. Setas de variação | **série XY com linhas** (5 pontos por par, linha vazia entre pares) com **ponta de seta** e **rótulo em elipse** (`DataLabel.Format.AutoShapeType = msoShapeOval`) ligado a uma célula com o texto | Vivas sem o ThinQQ (posição e texto); percentual `(B − A) / A`, "n/d" com A = 0 ou sinais opostos; alternador "Variação em valor"; espaço acima das barras reservado por uma auxiliar XY invisível (entra na escala automática); afastamento em pontos convertido por `PtV` (valor por ponto) recalibrado pelo refresh com teto de 3 passadas e tolerância de 2%. |
| 6. Rótulos | Degraus no centro da barra (auxiliar de linha por série, fonte por contraste); barras grandes (início e totais) acima em negrito (abaixo quando negativas) com "valor (percentual do primeiro total)" | Texto dos totais montado por `FIXED` no bloco (usa os separadores do Excel que calculou o bloco por último; outra localidade: Ctrl+Alt+F9); percentual liga e desliga pelo alternador **Totais**; degraus finos demais (altura < 10 pt ou texto mais largo que a fatia) ficam com texto vazio pela coluna `LblOn` do bloco (o rótulo apagado pelo usuário continua apagado). |
| 7. Integração | Matriz 21 tipos × 23 recursos; faixa (Cascata › Cascata comparativa com ícone próprio; Setas › Setas de variação e Variação em valor; Legenda › Nome dentro das barras); Alterar dados, Atualizar, cópia, limpeza; Exemplos (27); demonstração E27 e E28; galeria; tutorial; README; CHANGELOG | §6, §7, §8 |
| 8. Fora do escopo | §10 | N > 2 séries, barras horizontais, cálculo "de cima para baixo" (`e` na primeira linha), setas de variação entre etapas diferentes, Mekko. |

## 2. Pesquisa (think-cell, 2026-10-13)

- **Como o think-cell faz este gráfico.** O guia "Expert guide to multiple series waterfall charts" descreve exatamente o caso do pedido:
  "If you have two series and you are more interested in highlighting the comparison between them, rather than their sum total, then you can
  use a standard waterfall chart. Simple category gaps let you cluster the columns and color makes it easy for your audience to see the
  comparison." e adverte: "This alternative to a multiple series waterfall chart only works well when you have exactly two series to compare."
  Ou seja: não é um tipo próprio do think-cell; é uma cascata padrão com colunas alternadas (A, B), **lacunas de categoria** entre os pares e
  cores por série. Fonte: https://www.think-cell.com/en/resources/content-hub/expert-guide-to-multiple-series-waterfall-charts
- **Cascata.** "In a basic waterfall chart, every two columns are connected by precisely one horizontal connector"; os conectores podem ser
  arrastados para ligar outras colunas (é como o think-cell liga A₁ a A₂ passando por cima de B₁); "the segment labels in waterfall charts show
  the extent of the segment which is always a positive value" (sinal só pelo formato); `e` ("equals") cria um subtotal até a linha de base.
  Fonte: https://www.think-cell.com/en/resources/manual/waterfall
- **Setas de diferença.** "Level difference arrows ... designed to show differences between pairs of segments or points in a chart" (ligam
  segmentos de colunas e de cascatas); a seta de uma ponta calcula a diferença **relativa**, a de duas pontas a **absoluta**; o rótulo é uma
  "bubble" posta junto da seta e acompanha os dados; o manual sugere uma lacuna de categoria quando a seta fica entre duas colunas.
  Fonte: https://www.think-cell.com/en/resources/manual/chartdecorations
- **Percentual da imagem do usuário.** Os cinco percentuais da imagem (+7%, +75%, −67%, +67%, −9%) só fecham com `(B − A) / A` sobre o valor
  assinado da barra: Componente 1 de −4 para −7 dá (−7 − (−4)) / (−4) = **+75%** (o custo cresceu 75%); com `(B − A) / |A|` (sugerido no brief)
  daria −75%, contrariando a imagem. Decisão: `(B − A) / A`, "n/d" quando A = 0 **ou** quando A e B têm sinais opostos (ex.: −7 para +13 daria
  −286%, sem leitura útil; medido na sonda 4), e o alternador "Variação em valor" mostra `B − A` no formato do gráfico.
- **O que o ThinQQ faz igual:** duas pontes lado a lado com lacuna entre as etapas, cores por série, conectores por série (passando sobre a outra
  barra), `e`/`t`/`=`/`total`/`subtotal`, seta de uma ponta com a diferença relativa numa "bolha" (elipse), valor e percentual nos totais,
  nomes das séries na primeira barra. **Diferente:** os rótulos dos degraus mostram o valor **assinado** (−4), como a cascata do ThinQQ (o
  think-cell mostra a extensão positiva); a seta é sempre do par da mesma etapa (série 1 → série 2), não entre quaisquer segmentos; o percentual
  dos totais é sempre em relação à primeira barra da série.

## 3. Sondas em Excel real (Excel 2019 32 bits pt-BR; instâncias novas e invisíveis criadas pela sonda, PIDs conferidos como novos; arquivos em `C:\temp\tqq-cc-probe`, apagados ao final; nenhum Excel do usuário tocado)

### 3.1 Construção (sondas 1 e 2)
| Medida | Resultado |
|---|---|
| Colunas empilhadas com 3n − 1 categorias (A₁, B₁, vão₁, ..., Aₙ, Bₙ), séries `base` (invisível), `A+`, `A−`, `B+`, `B−` (gap 30%) | Barras corretas; centro da fatia k = `InsideLeft + (k − 0,5) × InsideWidth / (3n − 1)` (fatia 1: L = 12,41, W = 20,53; largura da fatia 26,69 = 373,66 / 14; 26,69 / 20,53 = 1,30 = 1 + 30%). |
| `XValues` de 2 colunas com a coluna interna **vazia** | O Excel junta tudo num nível só: um rótulo por fatia, girado 45°, preso à fatia A (não centrado no par) — PNG `p1_multilevel.png`. |
| `XValues` de 2 colunas, externa = nome da etapa na fatia A, vazio na B, " " no vão; interna = " " em todas | `TickLabels.MultiLevel = True`; o nome da etapa fica **centrado sob o par A–B** (grupo externo = {A, B}; o " " do vão forma um grupo invisível). Com a linha do eixo visível o Excel desenha os separadores verticais do eixo de 2 níveis (feio); com `Axes(xlCategory).Format.Line.Visible = msoFalse` eles somem (PNG `p2_a`/`p2_b`). A linha interna " " ocupa uma faixa curta abaixo das barras. |
| Fórmula que devolve `""` na externa da fatia A (categoria vazia na origem) | Igual a vazio: o par fica sem nome, sem desarrumar os vizinhos. O bloco devolve " " para categoria em branco. |
| Categoria data na origem | O rótulo segue o `NumberFormat` da célula externa do bloco (o bloco copia o formato da origem, como a borboleta). |
| Auxiliar XY no grupo **primário** do gráfico de colunas (`ChartType = xlXYScatter`, `AxisGroup` lido 1) | X é o **número da fatia**: x = 1, 2 e 14 caem em 22,64 / 53,93 / 429,36 pt = os centros das fatias 1, 2 e 14 exatamente; y(10) = 168,00 pt = regra de três no eixo. O gráfico passa a `ChartType = xlCombination` (−4111). |
| Pontos XY no grupo primário e a escala automática | Entram nela: com as setas o máximo do eixo foi de 35 para 40 (a sala acima das barras é reservada pelo próprio Excel, viva). |
| `GapWidth` 30 → 80 | Posições XY não mudam (fatias); a barra estreita de 20,53 para 17,38 pt: a meia-largura em unidades de fatia (`HW = 0,5 / (1 + gap/100)`) precisa acompanhar o espaço (célula de meta atualizada pelo ThinQQ). |

### 3.2 Conectores e setas (sondas 2 e 3)
| Medida | Resultado |
|---|---|
| Auxiliar XY sem marcador em (xA_i + HW, End_i) com `ErrorBar Direction:=xlX, Include:=xlErrorBarIncludePlusValues, Type:=xlErrorBarTypeCustom, Amount:=faixa` (3 − 2·HW) e `EndStyle = xlNoCap`, linha tracejada 0,75 pt | Um segmento horizontal por conector, do lado direito de Aᵢ ao lado esquerdo de Aᵢ₊₁, na altura do fim de Aᵢ; atravessa Bᵢ (desenhado por cima das colunas). Vivo (amount e posição em células). |
| Série XY com linhas (`xlXYScatterLinesNoMarkers`) com 5 pontos por par e uma **linha vazia** entre pares, `DisplayBlanksAs = xlNotPlotted`, `Format.Line.EndArrowheadStyle = msoArrowheadTriangle` | Cada par vira um cotovelo separado **com a sua própria ponta de seta** (29 pontos para 5 pares). |
| Rótulo do ponto do meio ligado a célula + `DataLabel.Format.AutoShapeType = 9` (oval), preenchimento branco, borda cinza, `Position = Center` | Aceito; elipse centrada na linha horizontal; H = 19,9 pt, W = 24 a 35 pt ("+7%" a "+200%"). Texto gravado como **constante** "+7%" vira o número 7% (o rótulo mostra "7%"): o texto vem de **fórmula** no bloco. |
| Nomes das séries: auxiliar XY com 2 pontos no meio da 1.ª barra de cada série, rótulo ligado ao cabeçalho, `Orientation = xlUpward` | Vertical, dentro da barra; "Segmento 1" a 9 pt: W = 14,0, H = 49,6 pt. Numa barra baixa (valor 5 de 25) o nome transborda (PNG `p4_zero.png`): precisa da regra de caber (§4.6). |

### 3.3 Bloco com fórmulas, rótulos e vida sem o ThinQQ (sondas 3 e 4)
| Medida | Resultado |
|---|---|
| Bloco inteiro em fórmulas (como a produção), dados da imagem | PNG `p3_base.png`: as duas pontes, nomes verticais, rótulos "27"/"29" acima das barras iniciais, degraus no centro, "11 (41%)" e "10 (34%)", elipses "+7%", "+75%", "−67%", "+67%", "−9%" — os números da imagem do usuário. |
| `PtV` (valor por ponto) inicial 0,15; calibração `(máx − mín) / InsideHeight` | 1.ª passada: eixo 0..40, H = 229,1, PtV = 0,1746 (gravado); 2.ª: igual (para). Dados que cruzam o zero: eixo −20..25, PtV = 0,1964, também em 1 gravação. |
| Mudar a origem (C4 = −12) sem VBA, só `Calculate` | Barras, conectores, elipse ("+200%") e total ("5 (17%)") acompanham (PNG `p3_base_after_change.png`). |
| `FIXED` (texto com os separadores do Excel que calcula) | `FIXED(ROUND(−0,4;0);0)` = "0" (sem "−0"); `FIXED(1234,5;1)` = "1.234,5"; `FIXED(−1234567;0)` = "−1.234.567". |
| Total negativo com rótulo "Acima" no ponto Lo | O rótulo cai dentro da barra (PNG `p4_zero.png`): posição "Abaixo" quando o total é negativo (refresh). |
| Sinais opostos (−7 → +13) | `(B − A)/A` = −286% (sem leitura útil) → "n/d". |
| Legenda | 14 entradas para 14 séries em 3 grupos (colunas empilhadas 5, linha 3, XY 6); com a legenda à direita as 5 entradas das colunas vêm em ordem **invertida** (cores lidas nas `LegendKey`: B−, B+, A−, A+, base) e as dos grupos de linha e XY na ordem das séries — exatamente o mapeamento de `LegendGroupReversed`, então `Chart_HideHelperLegendEntries` deixa só as duas séries de dados. |
| Escala lida logo depois de gravar uma célula do bloco (sem `Chart.Refresh` nem exportar) | `PtV` 0,6 → eixo 0..60; de volta a 0,1746 → 0..40: o máximo automático lido já é o novo (a calibração pode ler o eixo na passada seguinte). `TickLabels.Orientation` lido = −4128 (`xlHorizontal`). |
| Rótulos | Rótulo de degrau: H = 14,0 pt; rótulo de total: 5,5 pt acima da barra, H = 14,0 pt; elipse H = 19,9 pt. |
| Redimensionar para 300 × 220 sem VBA | PtV fica 0,1746 com o real 0,2366: o afastamento encolhe 26% (a seta encosta no rótulo do total) e rótulos "−12,0" ficam mais largos que a barra (PNG `p4_base_small.png`). Com o ThinQQ o evento Resize recalibra (§4.5, §4.6); sem ele fica legível e vivo, só mais apertado. |

## 4. Decisões

Critérios das fases anteriores: (1) correto sem o ThinQQ; (2) bloco em `_ThinQQ`, auxiliares `tqq_*`, rótulos por `DataLabel.Formula`; (3) Excel
2016+ 32/64 bits, sem `Declare`; (4) testável em Excel invisível; (5) pt-BR: formato de rótulo por `NumberFormatLocal`, fórmulas em inglês.

### 4.1 Tipo, família, tabela e recusas
- Chave `waterfallpair`, rótulo "Cascata comparativa", **família própria `waterfallpair`** (não `waterfall`): o bloco, os rótulos, a legenda (sempre
  disponível, sem a regra "só com várias séries") e os despachos são outros, e vários pontos do código tratam `Kinds_Family = "waterfall"` como
  "bloco da cascata atual" (conectores de forma em `Ann_Refresh`, `Labels_RefreshPositions`, `Labels_ApplyCellFormat`, `Waterfall_SeriesCount`).
  Vertical (`Kinds_IsHorizontal` falso), não 100%. Módulo `modChartWaterfallPair.bas`, prefixo `Wfp_`.
- Tabela: `Etapa | Série 1 | Série 2`, **exatamente 2 colunas de valores**. Recusas antes de alocar (faixa 6007..6099, `Core_Insert` mostra como
  aviso, nada criado): **6016** "A cascata comparativa precisa de exatamente 2 colunas de valores (uma por série); a tabela tem só 1." e **6017**
  "A cascata comparativa compara exatamente 2 séries; a tabela tem N colunas de valores. Selecione a coluna das etapas e só as duas colunas a
  comparar." (N = número de colunas de valores).
- Cada série tem a sua ponte, independente: marcadores `e`, `t`, `=`, `total`, `subtotal` (a mesma lista `WF_MARKERS`) por célula; texto,
  vazio e erro valem 0; a primeira linha é a barra inicial (de 0 ao valor); marcador na primeira linha = total de nada (0). **Sem** cálculo de
  trás para frente (`e` na primeira linha com número na última): fora do escopo, documentado.

### 4.2 Bloco (`Wfp_BuildBlock`)
Corpo com `6n` linhas de dados + cabeçalho e **58 colunas**; cada seção usa a sua indexação de linhas (linha 2 do corpo = 1.ª etapa / 1.ª fatia /
1.º ponto):

| Seção | Colunas (corpo) | Linhas | Conteúdo |
|---|---|---|---|
| Etapas | 1 `Cat`; 2..13 lado A; 14..25 lado B: `Src, Tot, D, Start, End, Lo, Hi, Base, PP, PN, Val, Big`; 26 `ArrowTxt`; 27 `TotTxtA`; 28 `TotTxtB` | n | ponte de cada série e textos |
| Fatias | 29 `Outer`, 30 `Inner`, 31 `SBase`, 32 `APos`, 33 `ANeg`, 34 `BPos`, 35 `BNeg`, 36 `LblAPos`, 37 `LblAVal`, 38 `LblBPos`, 39 `LblBVal`, 40 `TotPos`, 41 `TotTxt`, 42 `LblOn`, 43 `Top` | 3n − 1 | o que o gráfico plota, fatia a fatia (A = 3i − 2, B = 3i − 1, vão = 3i) |
| Conectores | 44 `CxA`, 45 `CyA`, 46 `CLen`, 47 `CxB`, 48 `CyB` | n − 1 | origem e comprimento de cada conector |
| Setas | 49 `PX`, 50 `PY`, 51 `PLbl` | 6n − 1 | 5 pontos por par + 1 linha vazia (separador) |
| Sala | 52 `RX`, 53 `RY` | n | ponto invisível acima de cada elipse |
| Nomes | 54 `NX`, 55 `NY`, 56 `NTxt` | 2 | meio da 1.ª barra de cada série e o texto |
| Zero | 57 `ZX`, 58 `ZY` | 2 | linha do zero (0,5 a 3n − 0,5) |

Meta (linha 1 do bloco, colunas E..O): `E PtV` (valor por ponto), `F HW` (meia-largura da barra em fatias), `G Div`, `H Dec`, `I Suf`, `J SufD`
(formato dos textos `FIXED`), `K Abs` (variação em valor), `L Lbl` (rótulos ligados), `M TotPct` (percentual nos totais), `N NameOn`
(nome dentro das barras), `O Fit` (os dois nomes cabem).

Fórmulas (lado j, etapa i; `s` = célula `Src`; tudo relativo dentro do bloco, metas absolutas):
```
Src   = 'Planilha'!$X$r                     Tot = IF(ISERROR(s),FALSE,IF(ISNUMBER(s),FALSE,OR(LOWER(TRIM(s&""))={"e","t","=","total","subtotal"})))
D     = IF(Tot,0,IF(ISNUMBER(s),s,0))       Start = 0 (i = 1) | End da linha anterior
End   = IF(Tot,Start,Start+D)               Lo = IF(Tot,MIN(0,End),MIN(Start,End))     Hi = IF(Tot,MAX(0,End),MAX(Start,End))
Base  = MAX(0,Lo)+MIN(0,Hi)                 PP = MAX(0,Hi)-MAX(0,Lo)                    PN = MIN(0,Lo)-MIN(0,Hi)
Big   = TRUE (i = 1) | Tot                  Val = IF(Big,End,D)
ArrowTxt = IF(Abs=1, sinal & FIXED(ROUND((B-A)/Div,Dec),Dec) & SufD,
              IF(OR(A=0,A*B<0),"n/d", sinal & FIXED(ROUND(100*(B-A)/A,0),0) & "%"))          (A = Val do lado A, B = Val do lado B)
TotTxt_j = IF(Big, FIXED(ROUND(End/Div,Dec),Dec) & Suf & IF(AND(TotPct=1, i>1, End_1<>0)," (" & FIXED(ROUND(100*End/End_1,0),0) & "%)",""), "")
Fatia A_i: Outer = Cat_i (formato da origem); Inner = " "; SBase = Base_A; APos = PP_A; ANeg = PN_A; BPos = BNeg = 0;
           LblAPos = IF(Big_A,NA(),(Lo_A+Hi_A)/2); LblAVal = IF(OR(Big_A,LblOn=0),"",D_A); LblBPos = NA(); LblBVal = "";
           TotPos = IF(Big_A,IF(End_A>=0,Hi_A,Lo_A),NA()); TotTxt = TotTxtA_i; LblOn = 1; Top = Hi_A        (B_i idem com o lado B)
Vão_i:     Outer = Inner = " "; SBase..BNeg = 0; posições NA(); textos ""; Top = 0
Conector i (i < n): CxA = (3i-2)+HW; CyA = End_A,i; CLen = 3-2*HW; CxB = (3i-1)+HW; CyB = End_B,i
Seta i:    sA = Hi_A + PtV*IF(AND(Big_A,Lbl=1,End_A>=0),22,3); sB idem; yH = MAX(sA,sB) + 12*PtV
           P1 (3i-2, sA) | P2 (3i-2, yH) | P3 ((6i-3)/2, yH, rótulo = ArrowTxt_i) | P4 (3i-1, yH) | P5 (3i-1, sB) | linha vazia
Sala i:    RX = (6i-3)/2; RY = yH + 12*PtV          Nomes: (1, (Lo_A1+Hi_A1)/2), (2, (Lo_B1+Hi_B1)/2); NTxt = IF(AND(NameOn=1,Fit=1), cabeçalho&"", "")
```
Constantes em pontos (VBA, entram nas fórmulas como número): afastamento da barra 3; sala do rótulo do total 22 (5,5 pt de afastamento + 14 pt do rótulo, medidos, + 2,5);
subida até a linha horizontal 12; meia elipse + respiro 12.

### 4.3 Gráfico (`Wfp_Insert`)
- Séries de colunas empilhadas na ordem `tqq_base` (invisível), `<cabeçalho 1>` (A+), `tqq_aneg` (A−), `<cabeçalho 2>` (B+), `tqq_bneg` (B−),
  categorias = `Outer:Inner` (2 colunas). `xlColumnStacked`, `DisplayBlanksAs = xlNotPlotted` (necessário às linhas vazias das setas; as
  colunas são fórmulas, nunca vazias), `GapWidth = 30` (padrão do tipo: barras do par quase encostadas), `Style_ApplyBase`, eixo de categorias
  com `TickLabels.MultiLevel = True`, `Orientation = xlHorizontal`, posição Baixo e **linha invisível** (sem os separadores do 2 níveis).
- **Linha do zero**: auxiliar XY `tqq_zero` com linha cinza de 0,75 pt de x = 0,5 a 3n − 0,5 em y = 0 (o eixo sem linha não marca a base).
- Cores: série 1 = `pal(0)`, série 2 = `pal(2)` (navy e azul-claro na paleta padrão, o contraste da imagem; cinza escuro e claro na "cinza";
  destaques 1 e 3 na "tema"); partes negativas com a cor da série; base invisível.
- Opções do Insert: `numfmt=int, palette=padrao, gap=30, n, totals=1, labels=1, connectors=1, pairdiff=1, pairabs=0, legend=inbar`.
- Rollback: `Util_TestHookAfterBlock` depois do bloco, `Util_TestHookAfterChart` depois do gráfico; falha apaga gráfico e bloco.

### 4.4 Conectores (`Wfp_SetConnectors`, alternador Conectores)
Duas auxiliares XY sem marcador (`tqq_conn_a`, `tqq_conn_b`), n − 1 pontos cada, com barra de erro X custom (`CLen`, faixa passada em texto R1C1: a forma que o Excel aceita para valores personalizados), sem ponta, cinza
RGB(127,127,127), tracejada, 0,75 pt. `HW` acompanha o **Espaço** (`Wfp_SyncGap`: em `ApplyOption "gap"` e em todo refresh, lendo o
`GapWidth` real, inclusive mudado à mão no Excel). Atravessam a barra da outra série (decisão: como o think-cell; quem não quiser desliga os
Conectores). Tabela de uma etapa: sem conectores.

### 4.5 Setas de variação (`Wfp_SetArrows`, alternadores Setas de variação e Variação em valor)
- `tqq_pd`: XY com linhas, 6n − 1 pontos (P1..P5 por par e a linha vazia), cinza RGB(89,89,89) 1 pt, ponta triangular curta no fim de cada par;
  rótulo só no P3 de cada par, ligado a `PLbl`, centro, **elipse** branca com borda cinza RGB(166,166,166) 0,75 pt, 9 pt cinza-escuro.
- `tqq_pdroom`: XY invisível com um ponto 12 pt acima de cada elipse: a escala automática deixa a sala da elipse (§3.1).
- **Calibração `PtV`** (`Wfp_RefreshRoom`, em todo refresh e no fim de `Chart_EnsureRoomAfterChange`): `PtV = (máximo − mínimo do eixo) /
  InsideHeight`; grava só se mudou mais de 2%; no máximo 3 passadas; cálculo manual recalcula o bloco (`Range.Calculate`); guarda de reentrada
  (gravar dispara o Calculate do gráfico, que volta ao refresh). Medido: converge em 1 gravação (§3.3). Teto: as 3 passadas; com o eixo fixo
  (Mesma escala) a primeira já é exata. Sem o ThinQQ, `PtV` fica: as setas continuam vivas, só o afastamento em pontos varia com o tamanho.
- Texto: percentual `(B − A) / A` arredondado a inteiro com sinal ("+75%"), "n/d" se A = 0 ou A·B < 0; com **Variação em valor** (`pairabs=1`,
  meta `Abs`): `B − A` com sinal no formato do gráfico (`Div`/`Dec`; nos formatos percentuais o sufixo é " p.p.").
- Desligar Setas de variação apaga `tqq_pd` e `tqq_pdroom` (as colunas do bloco ficam); religar recria.

### 4.6 Rótulos, totais e nomes
- `tqq_lbl_a` / `tqq_lbl_b`: auxiliares de linha (`Chart_AddHelperSeries` vertical) em `LblAPos`/`LblBPos`, rótulos ligados a `LblAVal`/`LblBVal`
  por `Labels_LinkPoints` (centro, formato `numfmt` por `NumberFormatLocal`, fonte por contraste com a cor da série). Pontos `NA()` não aparecem
  (`xlNotPlotted`) e o texto ligado deles é "" (nunca "#N/D": `CheckNoErrorsAnywhere` dos Exemplos).
- `tqq_lbltot`: auxiliar de linha em `TotPos` com rótulos ligados a `TotTxt`, negrito; posição Acima, ou **Abaixo** quando `TotPos < Top` (total
  negativo), aplicada pelo refresh (só escreve se mudou; pula ponto sem rótulo).
- **Totais** (alternador existente, matriz `totals`): neste tipo liga/desliga o "(41%)" das barras de total (meta `TotPct`), não uma série.
- **Rótulos finos** (`Wfp_SyncLabels`): degrau com altura < 10 pt (texto de 9 pt; a caixa de 14 pt tem margem) ou texto (estimado por `Labels_EstimateWidth`) mais largo que a fatia fica
  com `LblOn = 0` (texto vazio); volta quando cabe. Como o rótulo continua existindo (só o texto muda), um rótulo apagado pelo usuário continua
  apagado: o ThinQQ nunca recria rótulo no refresh (contabilidade transacional sem lista).
- **Formato**: `Wfp_ApplyNumberFormat` grava `NumberFormatLocal` nas células e rótulos dos degraus e `Div/Dec/Suf/SufD` na meta (int 1/0, dec1
  1/1, dec2 1/2, pct0 0,01/0/"%"/" p.p.", pct1 0,01/1/"%"/" p.p.", mil 1000/0, mi 1000000/1).
- **Nomes dentro das barras** (`tqq_names`, item novo da Legenda "Nome dentro das barras", padrão do tipo): texto ligado a `NTxt`, vertical,
  9 pt, fonte por contraste. Cabe se a 1.ª barra da série tem altura ≥ largura estimada do nome + 8 pt e largura ≥ 14 pt (medido: texto vertical
  de 9 pt tem 14 pt de largura); se **algum** não cabe, os dois somem (`Fit = 0`) e a legenda à direita aparece (troca automática pelo refresh);
  Direita = legenda sem nomes; Nenhuma = nada.

### 4.7 Integração
- Matriz (§6), `Kinds_FeatureLabel`: `pairdiff` "Setas de variação", `pairabs` "Variação em valor", `inbar` "Nome dentro das barras".
- `modCoreApi`: `ApplyToggle` (totals, labels, connectors, pairdiff, pairabs), `PALETTE_FAMILIES` + despacho, `gap` → `Wfp_SyncGap`, legenda
  "inbar" em `ValidateOption`, `Core_SafeRefresh` → `Wfp_RefreshAfterDataChange`, `Core_ChartSourceMismatch` (`Wfp_SourceRowsSpanned`, como o Mekko:
  as séries apontam para o bloco), `ReapplyOptions` (pairdiff padrão 1), `Kinds_ExpectedCategoryRefs = 2 × (3n − 1)` (categorias de 2 colunas).
- `Deco_SetLegend`: modo `inbar` (recusa 7002 fora do tipo); família própria despacha a `Wfp_SetLegend`.
- `Ann_GetCategoryData`: `Top` de cada fatia (para `Deco_ScaleClipped`/Mesma escala). Setas CAGR/diferença, colchete, linha de valor, corte etc.:
  recusa pela matriz (7001/7002) com os textos existentes.
- `Labels_RefreshPositions` sai cedo na família (o refresh do tipo cuida); `Labels_ApplyCellFormat` e `Style_ApplyNumberFormat` despacham a
  `Wfp_ApplyCellFormat`/`Wfp_ApplyNumberFormat`.
- Limpeza (`CalcSheet_Cleanup`): todas as séries do tipo citam o bloco em `Series.Formula` (as faixas das barras de erro não aparecem na fórmula,
  mas X/Y das mesmas séries sim): cópia de gráfico e gráfico movido continuam mantendo o bloco vivo, sem mudança na limpeza.
- Faixa (§7), Exemplos ("27) Cascata comparativa", planilha Exemplos 3), demonstração E27 ("Da receita ao lucro, 2024 x 2025") e E28 ("Aurora
  Centro x Porto Sereno", variação em valor, legenda à direita), galeria regerada, tutorial (seção `#cascata-comparativa`), README, CHANGELOG
  `[Não lançado]`, `release/notes.txt`.

### 4.8 Tabela de decisões
| Decisão | Escolha | Alternativa rejeitada | Por quê |
|---|---|---|---|
| Construção | colunas empilhadas com fatias A, B, vão | (b) duas cascatas no eixo principal e no secundário | pilhas não se agrupam; o secundário sobrepõe no mesmo centro; (a) é a técnica do próprio think-cell (lacunas de categoria) |
| Nome da etapa sob o par | eixo de categorias de 2 níveis (interna " ") | auxiliar XY com rótulo abaixo | medido: centrado e vivo, sem calcular o mínimo do eixo |
| Conectores | XY + barra de erro X custom | formas (como a cascata atual) | vivos sem o ThinQQ; nenhuma forma a reposicionar |
| Setas | série XY com linhas, gaps por linha vazia, ponta de seta, elipse no rótulo | formas `tqq_ann_*` | vivas (posição e texto) sem o ThinQQ; uma série para todos os pares |
| Sala acima | ponto XY invisível + `PtV` | folga fixa da área de plotagem | a escala automática já reserva; a folga fixa não acompanha os dados |
| Percentual | `(B − A) / A`, "n/d" com A = 0 ou sinais opostos | `(B − A) / |A|` | só `(B − A)/A` reproduz a imagem (+75% para −4 → −7) |
| Texto dos totais | `FIXED` no bloco | rótulo numérico com formato | o formato não junta dois números ("11 (41%)"); `FIXED` usa os separadores do Excel que calculou o bloco por último (outra localidade: Ctrl+Alt+F9) |
| Rótulo fino | texto vazio pela coluna `LblOn` | apagar/religar o rótulo com lista `thinlbl` | o rótulo apagado pelo usuário continua apagado sem lista; nada é recriado no refresh |
| Nomes | dentro da 1.ª barra quando cabem, senão legenda (automático) | sempre a legenda | é o que a imagem mostra; a troca fica no refresh |
| Família | própria | dentro de `waterfall` | o bloco e os despachos são outros (§4.1) |
| `pairabs`/`inbar` | recursos da matriz | casos especiais em `Core_Supports` | matriz é dado: faixa, recusas, tutorial e demonstração conferem tudo pela mesma fonte |

## 5. Contrato com as tarefas

Mensagens exatas: 6016 e 6017 (§4.1); recusa de legenda "inbar" em outro tipo = `Kinds_NotApplicableText("inbar")` ("'Nome dentro das barras'
não se aplica a este tipo de gráfico."). Opções persistidas novas: `pairdiff`, `pairabs` (0|1), `legend=inbar`. Séries auxiliares: `tqq_base`,
`tqq_aneg`, `tqq_bneg`, `tqq_zero`, `tqq_names`, `tqq_conn_a`, `tqq_conn_b`, `tqq_pd`, `tqq_pdroom`, `tqq_lbl_a`, `tqq_lbl_b`, `tqq_lbltot`.

## 6. Matriz de capacidades 1.3.0 (`Kinds_Supports`)

Tipos: os 20 da 1.2.0 + `waterfallpair` (família `waterfallpair`, vertical). Recursos novos: `pairdiff`, `pairabs`, `inbar`. Esta tabela é lida
pelo teste `Test_Kinds_PairWaterfall_Matrix_ExhaustiveFromSpec` (cópia em `docs/superpowers/specs/`).

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

## 7. Faixa de opções (loader 1.3.0)

- Menu **Cascata**: terceiro item `btnWaterfallPair` "Cascata comparativa", `tag="waterfallpair"`, `image="tqq_waterfallpair16"`, supertip com a
  tabela (etapa e exatamente 2 colunas de valores; `e` para total).
- Menu **Setas**: dois alternadores antes do separador: `tglPairDiff` "Setas de variação" (`tag="pairdiff"`) e `tglPairAbs` "Variação em valor"
  (`tag="pairabs"`), `Rx_OnToggle`/`Rx_GetPressed`/`Rx_GetEnabled`; `tglPairDiff` com `image="tqq_arrows16"`, `tglPairAbs` com `image="tqq_waterfallpair16"`, nenhum com `imageMso`.
- Dropdown **Legenda**: item `inbar` "Nome dentro das barras".
- Supertip de **Totais** acrescenta: "Na cascata comparativa, acrescenta o percentual do primeiro total ao rótulo das barras de total."
- Ícone `tqq_waterfallpair` (32 e 16 px) na tabela `Get-TqqIconSpecs`: três pares de barras navy/azul-claro descendo (início, degrau, total),
  reserva `imageMso="ChartTypeOtherInsertGallery"` (a mesma da cascata). 14 ícones na tabela.
- Contagens novas que os testes conferem: 21 botões de inserção, 12 alternadores, 4 dropdowns (Legenda com 4 itens), 8 menus, 3 botões de anotação.

## 8. Versões e documentação

`version.json` não é tocado pelas tarefas (o `publicar-release.bat 1.3.0` sobe loader e core juntos). Loader 1.3.0 com core 1.2.x: o botão novo
avisa "Tipo de gráfico desconhecido: waterfallpair" e os alternadores novos ficam cinza (`Core_Supports` não conhece os recursos); core 1.3.0 com
loader 1.2.x: o tipo só não aparece na faixa (Exemplos mostra). CHANGELOG `[Não lançado]` ganha "Adicionado"/"Alterado"; `release/notes.txt`
ganha a linha da cascata comparativa e, só se a tag `v1.2.3` ainda não existir (`git ls-remote --tags origin`), mantém as 2 linhas da correção 1.2.3
(download e reinstalação única pelo ZIP); limite 1 a 8 linhas. Se a 1.2.3 for publicada antes do merge, o controlador rebaseia a branch na
`main` e o `### Corrigido` da 1.2.3 sai do `[Não lançado]` (fica na `[1.2.3]`).

| Recurso | CHANGELOG | Tutorial | README | Exemplos | Testes de documento |
|---|---|---|---|---|---|
| Cascata comparativa | Adicionado | `#cascata-comparativa` (passo a passo, cálculo do percentual, limites), linha em `#tipos`, matriz (4.ª tabela) | Recursos, Verificações manuais, Limitações | Exemplos 27; demonstração E27, E28 | Test-Docs: 21 tipos, 23 recursos, seção, mensagens 6016/6017, frases dos limites; Test-Ribbon/TestRibbon: contagens |

## 9. Riscos e mitigações

| Risco | Prob. | Mitigação |
|---|---|---|
| `PtV` oscilar entre dois máximos "redondos" da escala automática | baixa | medido estável em 1 gravação (dois conjuntos de dados); grava só com mudança > 2%; 3 passadas no máximo; teste de estabilidade em dois refresh e num gráfico pequeno |
| Reentrada: gravar meta (`PtV`, `HW`, `Fit`, `LblOn`) dispara o Calculate do gráfico | certa | guarda `mBusy` em `Wfp_RefreshAfterDataChange`; cada sub grava só quando muda; cálculo manual recalcula o bloco |
| Legenda com 3 grupos de séries (colunas, linha, XY) | média | medido: 14 entradas, grupos na ordem colunas/linha/XY; `Chart_HideHelperLegendEntries` cobre grupos mistos (teste: 2 entradas) |
| Conectores sobre a barra da outra série | certa | comportamento do think-cell; documentado; alternador Conectores |
| Nome da série que não cabe | média | regra de caber (altura e largura) no refresh; troca para a legenda; teste com barra inicial baixa |
| Rótulos de degraus em gráfico estreito | média | `LblOn` por altura e por largura da fatia; teste com valor minúsculo |
| Sem o ThinQQ, redimensionar aperta as setas | certa | documentado ("afastamento das setas"); continua vivo e legível |
| Categoria longa | média | o Excel quebra o rótulo do eixo de 2 níveis em linhas; documentado |
| Faixa com controle novo e core antigo | baixa | `Core_Supports` desconhecido = cinza; tipo desconhecido = aviso existente |
| Tempo da suíte | média | um módulo de teste novo (~28 testes, 1 gráfico por teste, fixtures de 5 etapas); `-Filter` na iteração |
| Limpeza de blocos órfãos | baixa | todas as séries citam o bloco; teste de cópia de gráfico + `CalcSheet_Cleanup` (`Test_Wfp_CopiedChart_KeepsBlockAlive`, T2) |

## 10. Fora do escopo (e por quê)

- **Mais de 2 séries comparadas lado a lado**: a construção generaliza (fatias por série + vão), mas as setas (qual par?), os nomes e as cores
  pediriam outro desenho de faixa e de rótulos; o próprio think-cell recomenda a técnica só para duas séries (§2). Fica para uma versão futura.
- **Barras horizontais** (cascata comparativa deitada): exigiria o eixo de 2 níveis vertical e auxiliares XY no grupo secundário (como as barras
  da Fase 2): o dobro das sondas e dos testes de geometria.
- **Cálculo "de cima para baixo"** (`e` na primeira linha com número na última), **setas entre etapas diferentes** (ex.: A₁ → A₅), **rótulos de
  extensão sem sinal** (think-cell), **Mekko** e **cascata comparativa com quebra de eixo**.
- **Seta CAGR/diferença** neste tipo (as setas de variação cobrem a comparação).

## 11. Pré-flight do plano (varredura de coerência)

Pares de tarefas que compartilham arquivo ou interface (o que uma produz, o que a outra consome, resultado) e a autoconsistência de cada tarefa.

| Par | Produz | Consome | Resultado |
|---|---|---|---|
| T0 → T1 | `Kinds_IsValid("waterfallpair")`, família, matriz 21 × 23, `Kinds_FeatureLabel` dos 3 recursos, `Kinds_ExpectedCategoryRefs`, `Kinds_Insert` caindo no 7200, saídas cedo da família em `Labels_*`/`Style_ApplyNumberFormat`, teste temporário de tipo não pronto | T1 troca o `Case "waterfallpair"` de `Kinds_Insert` e remove o teste temporário | limpo |
| T0 → T5 | matriz do tutorial com `<span class="k">` na linha e nas 3 colunas novas | T5 troca para `ui` (rótulos passam a existir na faixa) | limpo: o Test-Docs só confere o rótulo da linha contra a faixa quando o botão existe |
| T1 → T2/T3/T4 | bloco completo (todas as seções e metas), `Wfp_SideCol`, `Wfp_SlotOf`, `Wfp_StepCount`, `AddXY`, `Wfp_RefreshAfterDataChange` com `SyncGap`/`SyncNames` | T2 acrescenta `Wfp_SetConnectors`; T3 `Wfp_SetArrows`/`Wfp_SetArrowMode`/`Wfp_RefreshRoom`; T4 `Wfp_SetLabels`/`Wfp_SetTotalPct`/`Wfp_SyncLabels`/`Wfp_ApplyNumberFormat` e cada um acrescenta uma linha a `Wfp_Insert` e a `Wfp_RefreshAfterDataChange` | limpo: as linhas a acrescentar são dadas com a âncora exata |
| T1 → T6 | `Wfp_Insert`, `Wfp_TotalPctOn` (T4) | `FixtureFor`, `AssertRealState`, Exemplos 27 | limpo |
| T3 → T4 | meta `Lbl` lida pela seta (sala do rótulo do total) | `Wfp_SetLabels` grava `Lbl` | limpo: o teste de T4 confere que desligar os rótulos baixa o início da seta |
| T2..T4 → T5 | recursos `connectors`, `pairdiff`, `pairabs`, `inbar` despachados | alternadores e item de legenda na faixa | limpo |
| T5 → T7/T8 | rótulos exatos da faixa | DemoCoverage (`Menu`), tutorial, README | limpo |
| T6 → T8 | 27 exemplos (`Samples_Create`) | Test-Docs `$nEx` | limpo: T6 troca "26 exemplos" de ThinQQ › Exemplos para 27 |
| T7 → T8 | 28 exemplos da demonstração, galeria regerada | Test-Docs `$nDemo`, "E01 a E28" | limpo: T7 troca as citações da demonstração |
| T8 → T9 | documentação | suíte completa, conferência visual | limpo |

| Tarefa | Autoconsistência | Resultado |
|---|---|---|
| T0 | testes usam só funções da matriz; Test-Docs 21/23 com a 4.ª tabela; a cópia do spec existe antes do teste | limpo |
| T1 | `TestWaterfallPair.bas` só usa `Wfp_*` públicas, `Kinds_*`, `Geom_*`, `Registry_*`; geometria por regra de três própria (`XOfX`, `YOf`) | limpo |
| T2 | testes de conectores leem X/Y/CLen do bloco e comparam com as bordas das colunas (`Points.Left/Width`) | limpo |
| T3 | textos esperados derivados dos valores da tabela; `PtV` comparado com o eixo real | limpo |
| T4 | textos com `Application.International`/`Format$`; rótulo apagado continua apagado | limpo |
| T5 | contagens 21/12/4/8/3 e 14 ícones batem com o XML descrito | limpo |
| T6 | `TestSamples` 9 seções em Exemplos 3, 28 gráficos no total | limpo |
| T7 | `Get-TqqDemoFeatures` + mapa do Test-DemoWorkbook com os 3 recursos; galeria com 29 imagens | limpo |
| T8 | frases do `mustDocument` presentes no texto entregue pela própria tarefa | limpo |

## 12. Fontes

- think-cell — guia de cascatas com várias séries (duas séries lado a lado com lacunas de categoria): https://www.think-cell.com/en/resources/content-hub/expert-guide-to-multiple-series-waterfall-charts
- think-cell — cascata (conectores, `e`, rótulos de extensão): https://www.think-cell.com/en/resources/manual/waterfall
- think-cell — anotações (setas de diferença de nível, relativa/absoluta, "bubble"): https://www.think-cell.com/en/resources/manual/chartdecorations
- Peltier — colunas agrupadas e empilhadas por categorias expandidas: https://peltiertech.com/clustered-stacked-column-bar-charts/
- Peltier — XY no eixo principal de colunas (categoria k em x = k): https://peltiertech.com/precision-positioning-of-xy-data-points/
- Microsoft — `Series.ErrorBar`: https://learn.microsoft.com/en-us/office/vba/api/excel.series.errorbar
- Microsoft — `TickLabels.MultiLevel`: https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.multilevel
- Microsoft — `ChartFormat.AutoShapeType` (forma do rótulo de dados): https://learn.microsoft.com/en-us/office/vba/api/excel.chartformat.autoshapetype
- Microsoft — `FIXED`: https://support.microsoft.com/office/fixed-function-ffd5723c-324c-45e9-8b96-e41be2a8274a
