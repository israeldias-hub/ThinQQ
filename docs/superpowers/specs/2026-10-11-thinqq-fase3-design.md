# ThinQQ — Pesquisa e Design (Fase 3, versão 1.2.0)

Data: 2026-10-11
Autor: Israel Dias (com apoio do Claude Code)
Status: adendo às especificações da Fase 1 (`docs/superpowers/specs/2026-10-10-thinqq-design.md`, autoridade de arquitetura,
distribuição e atualização) e da Fase 2 (`docs/superpowers/specs/2026-10-11-thinqq-fase2-design.md`, matriz de capacidades,
blocos por família, faixa em menus). Este documento cobre só o que a Fase 3 acrescenta ou muda. Plano de execução: `plan.md`
(mesma pasta; rascunho a ser promovido a `docs/superpowers/plans/2026-10-11-thinqq-fase3.md` pelo controlador).

## 1. Objetivo e escopo

Entregar o ThinQQ **1.2.0** (core e loader) com os itens adiados para a Fase 3, mantendo o princípio central: **todo gráfico é um
gráfico nativo do Excel ligado por fórmulas a um bloco da planilha muito oculta `_ThinQQ`**; quem recebe o arquivo sem o ThinQQ vê
o gráfico correto e vivo. Cada item abaixo foi decidido por pesquisa (§2) e por sondas em Excel e PowerPoint reais (§3).

| Item do brief | Decisão | Resumo |
|---|---|---|
| 1. Borboleta (`butterfly`) | **Entregue** | Barras empilhadas com dois lados espelhados (valores negativos à esquerda com formato que mostra o valor absoluto), nomes das categorias no centro, totais, legenda, eixo simétrico **vivo** (auxiliar XY nos dois extremos) e mesma escala. |
| 2. Dispersão (`scatter`) e bolhas (`bubble`) | **Entregue** | Dois eixos de valor, rótulo por ponto com o nome, uma série por grupo (coluna de grupo opcional), títulos de eixo ligados aos cabeçalhos, linha de tendência (nativa) e quadrantes nas médias (eixos cruzando nas médias). |
| 3. `TQQ.ROUND` (arredondamento consistente) | **Entregue como fórmula nativa** (opção (a)); **UDF recusada** | Alternador "Rótulos com soma consistente" (maior resto por fórmulas no bloco; os rótulos passam a apontar para as células arredondadas) nas colunas/barras/áreas empilhadas e 100%; botão "Arredondar intervalo..." grava a mesma fórmula numa faixa da planilha do usuário. Sem UDF: o Excel reescreve referências a funções de suplementos como vínculos externos ao caminho do `.xlam`, que muda a cada versão do core (§4.3). |
| 4. Copiar para o PowerPoint | **Entregue** | Menu "PowerPoint": como imagem (`CopyPicture` + `Paste`) e como gráfico vinculado (`ChartObject.Copy` + `Paste`: gráfico do PowerPoint vinculado à pasta salva, formas das anotações preservadas). Automação tardia, sem referência e sem `Declare`. Mensagens para PowerPoint ausente, sem apresentação, pasta não salva. |
| 5. Assinatura digital e política corporativa | **Parcialmente entregue (o que é possível sem admin e sem certificado corporativo)** | Scripts em `tools/`: criação do certificado autoassinado + editor confiável por usuário (testado), roteiro de assinatura por `signtool` + Office SIPs (suportado pela Microsoft, exige SDK e registro das DLLs com administrador: **não testável nesta máquina**, o script detecta e explica), gancho opcional `-SignPfx` no release, e o roteiro manual (VBE > Ferramentas > Assinatura Digital, assinatura V3) no README/tutorial. A assinatura pelo modelo de objetos do VBE **não existe**. |
| 6. Quebra de eixo | **Entregue como "corte de eixo"** (variante viável) | Máximo fixo do eixo de valores (o Excel corta as colunas/barras na borda da área de plotagem), rótulo de total numa coluna `LblPos = MIN(TotalPos, BreakAt)` do bloco (vivo), marca de corte (formas) reposicionada por `Ann_Refresh`, rótulos de segmentos cortados ocultos/na base. Colunas e barras empilhadas e agrupadas. Sem compressão de escala (não há como "encolher" um trecho do eixo num gráfico nativo). |
| 7. Ícones próprios na faixa | **Entregue** | PNGs 32×32 e 16×16 gerados no build por um gravador PNG puro em PowerShell (zlib "stored", bytes determinísticos, sem `System.Drawing`), empacotados em `customUI/images/` com `customUI/_rels/customUI14.xml.rels` e `Default png` no `[Content_Types].xml`; `getImage` não é necessário. Fallback: `build.ps1 -RibbonIcons Mso` troca `image=` por `imageMso=` (tabela no build). Teste automatizado do pacote. |
| 8. Itens baratos da §10 da Fase 2 | **Nenhum entregue** | Avaliados: linha de valor em valor arbitrário (exige regerar a fórmula da coluna VLine por família ao voltar para a média: não é barato), ordem de segmentos (mudaria a construção de todas as pilhas). Lista completa do que fica de fora em §11. |
| 9. Polimento e integração | **Entregue** | Faixa (menu "Mais tipos" com 3 itens novos, "Mais opções" com 3 alternadores + corte de eixo + arredondar intervalo, menu "PowerPoint"), matriz (20 tipos × 20 recursos), "Exemplos 3", tutorial/README/CHANGELOG, `version.json` 1.2.0/1.2.0, manifesto/release/instalador, pré-flight (§12). |
| Extra A (pedido do controlador): pasta de demonstração | **Entregue** | `exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx` gerada por `tools/New-DemoWorkbook.ps1` (dados fictícios determinísticos de uma rede de materiais de construção, um exemplo por funcionalidade, índice com hiperlinks internos) e `tests/ps/Test-DemoWorkbook.ps1`. |
| Extra B (pedido do controlador): documentação de instalação | **Entregue** | Seção "Instalação" completa e autossuficiente no README (usuário não técnico) e subseção "Ambientes corporativos (TI)" com o resultado da assinatura/editor confiável. |

## 2. Pesquisa (2026-10-11)

### 2.1 think-cell
- **Dispersão e bolhas** (manual "Scatter and bubble charts"): cada linha da planilha é um ponto; colunas X, Y, (Tamanho, só bolhas) e **Grupo** (pontos do mesmo grupo compartilham marcador e podem ser selecionados juntos); até dois rótulos por ponto com o texto do rótulo e/ou os valores x, y e tamanho; esquema de marcadores por grupo; bolhas com tamanho proporcional à **área** (opção "diâmetro"); legenda de tamanho de bolha; **linhas de tendência** (regressão linear, ponderada pelo tamanho nas bolhas; só com escala linear) e **partições** (linhas que dividem o gráfico; pode-se preencher um dos lados). Fonte: https://www.think-cell.com/en/resources/manual/scatterbubble
- **Arredondamento** (manual "Excel data rounding" e KB0028): o problema é que "o total dos valores arredondados não é necessariamente igual ao total arredondado dos valores precisos" (33% + 33% + 33% = 99%); o think-cell round (TCROUND) arredonda "o conjunto" de modo que as somas fechem, alterando o menor número possível de valores e com o menor desvio; só existe solução garantida com `+`, `-` e `SUM` (senão `#NUM!`). Fontes: https://www.think-cell.com/en/resources/manual/exceldatarounding e https://www.think-cell.com/en/resources/kb/0028
- **Quebra de eixo** (manual "Axes"): "Add Axis Break"/"Add Column Break" encolhe um trecho grande do eixo de valores para dar leitura aos segmentos pequenos; precisa de espaço para as duas linhas da marca; estilos "reto" e "ondulado"; a quebra vale para todos os gráficos na mesma escala. É uma **compressão de escala**, que o gráfico nativo do Excel não tem (o eixo de valores é linear ou logarítmico). Fonte: https://www.think-cell.com/en/resources/manual/axes
- **Borboleta**: o manual não tem página própria; na galeria de tipos o think-cell lista "butterfly" entre os gráficos (duas barras espelhadas com o eixo de categorias no meio). A técnica em Excel é consagrada: uma série negativa (lado esquerdo) com formato numérico que oculta o sinal e uma positiva, mesma escala nos dois lados. Fontes (técnica geral): https://www.exceldemy.com/create-butterfly-chart-in-excel/ e https://www.vizwiz.com/2020/12/butterfly-chart.html
- **Links Excel → PowerPoint** (manual "Excel data links"): no think-cell o gráfico do PowerPoint fica ligado a uma faixa do Excel e é atualizado sob demanda ou automaticamente; as desvantagens que o próprio think-cell aponta no "Colar especial" nativo (renomear/mover o arquivo quebra o vínculo) valem para o nosso gráfico vinculado e estão documentadas no tutorial. Fontes: https://www.think-cell.com/en/resources/manual/exceldatalinks e https://www.think-cell.com/en/resources/manual/manage-data-links

### 2.2 Microsoft e prática (assinatura de VBA)
- **Office Subject Interface Packages for Digitally Signing VBA Projects** (Microsoft, publicado em 21/10/2025; pacotes x86 e x64): duas bibliotecas SIP (uma para OOXML, outra para os demais formatos), `offsign.bat` e `offclearsig.exe`; registram-se com `regsvr32` (exige administrador na prática) e assinam-se os arquivos com o **SignTool do Windows SDK (versão x86)**. Fonte: https://www.microsoft.com/en-us/download/details.aspx?id=56617 e https://support.microsoft.com/topic/2b8b3cae-ad64-4b4b-aa85-c4a98ca6da87
- **Assinatura V3** (Microsoft 365 Dev Blog e KB5000676): esquema mais seguro após a CVE-2020-0760; o VBE (Ferramentas > Assinatura Digital) grava V3; o SignTool com os SIPs atuais serve para reassinar em lote; política de grupo "Only trust VBA macros that use V3 signatures". Fontes: https://devblogs.microsoft.com/microsoft365dev/upgrade-signed-office-vba-macro-projects-to-v3-signature/ e https://support.microsoft.com/en-us/servicing/office/hotfix/upgrade-signed-office-vba-macro-projects-to-v3-signature-kb5000676
- **SelfCert.exe** e assinatura pelo VBE (Microsoft "Digitally sign your macro project"): certificado autoassinado só é confiável na máquina em que o certificado está em "Autoridades de Certificação Raiz Confiáveis"; o VBE é a única interface documentada de assinatura (nenhuma API no modelo de objetos `VBProject`). Fontes: https://support.microsoft.com/office/digitally-sign-your-macro-project-956e9cc8-bbf6-4365-8bfa-98505ecd1c01 e https://msdn.microsoft.com/en-us/library/aa141471.aspx
- **PowerShell**: `New-SelfSignedCertificate -Type CodeSigningCert`, `Export-Certificate`, `Export-PfxCertificate`, `Import-Certificate -CertStoreLocation Cert:\CurrentUser\TrustedPublisher` (sem diálogo; em `Cert:\CurrentUser\Root` o Windows pede confirmação). Medido em §3.5.

### 2.3 Microsoft (faixa, PowerPoint, gráficos)
- **`image` no customUI**: imagens próprias ficam em `customUI/images/*.png`, relacionadas por `customUI/_rels/customUI14.xml.rels` (tipo `.../officeDocument/2006/relationships/image`); o `id` da relação é o valor de `image=`; PNG com canal alfa, 16×16 e 32×32. Fontes: https://bettersolutions.com/vba/ribbon/images-editor-custom.htm e https://learn.microsoft.com/en-us/answers/questions/917971/custom-image-for-group-in-custom-ribbon-xml
- **PowerPoint**: `Shapes.PasteSpecial(DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link)` com `ppPasteOLEObject` (10), `ppPasteEnhancedMetafile` (2); `ChartData.IsLinked` (somente leitura) e `ChartData.BreakLink`. Fontes: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.chartdata.islinked e https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa212207(v=office.11)
- **Excel**: `DataLabel.Formula` (Fase 1), `AxisTitle.Formula`/`Text`, `Axis.Crosses`/`CrossesAt`, `Series.Trendlines.Add`, `ChartGroup.BubbleScale`/`ShowNegativeBubbles`/`SizeRepresents`. Comportamento medido em §3.

## 3. Sondas em Excel/PowerPoint reais (Excel 2019 32-bit pt-BR; PowerPoint 16.0; instâncias novas e invisíveis criadas pela sonda; pasta própria em `C:\temp`, apagada ao final)

### 3.1 Borboleta
| Medida | Resultado |
|---|---|
| Barras empilhadas com 4 séries na ordem `tqq_gapL (−g)`, `A (−a)`, `tqq_gapR (+g)`, `B (+b)` | Negativos empilham para a esquerda na ordem das séries e positivos para a direita: gap à esquerda de 0 a −g e A de −g a −(g+a); gap à direita de 0 a +g e B de +g a +(g+b) (ponto A1 `Left=130,5 Width=45,7`, gapL1 `Left=176,2 Width=21,3`, B1 `Left=218,8`; x(0) = 197,5). |
| Rótulo nativo de A (valor negativo) com `NumberFormatLocal = "#.##0;#.##0;"` | Texto `120` (sem sinal). |
| Auxiliar XY `tqq_cat` no grupo secundário em (x = 0, y = Pos) com `DataLabel.Formula` ligado ao nome da categoria, posição `Center` | Centro do rótulo em x = 197,499 = x(0) (diferença < 0,001 pt). O Excel **não** cria eixo X secundário (como medido na Fase 2 para barras). |
| Eixo de valores automático sem auxiliar de simetria | Assimétrico: −500..400 (valores até 310 à esquerda e 290 à direita). |
| Auxiliar XY `tqq_sym` com 2 pontos em x = ±(g + max) | Eixo automático **simétrico**: −500..500; após mudar um valor para 600 (sem nenhuma ação do ThinQQ): −800..800. Ou seja, os pontos XY do grupo secundário participam da escala automática do eixo de valores primário: a simetria é **viva**, sem VBA. |
| Nome de categoria longo ("Trinta a cinquenta e nove") com gap fixo de 18% do máximo | Largura do rótulo 79 pt contra um vão de 42,5 pt: o nome invade as barras. Logo o vão deve derivar da largura do maior nome (estimada por `Labels_EstimateWidth`) e da largura interna do gráfico (§4.1). |
| Legenda antes de esconder auxiliares | 7 entradas (4 séries + 3 auxiliares): `Chart_HideHelperLegendEntries` (que já trata barras com XY secundária) deixa só A e B. |

### 3.2 Dispersão e bolhas
| Medida | Resultado |
|---|---|
| `xlXYScatter` com uma série por grupo (X/Y do bloco com `NA()` fora do grupo) e `DisplayBlanksAs = xlNotPlotted` | Pontos de outro grupo não são desenhados; `Points.Count` continua n. |
| `DataLabel.Formula = =Dados!$A$k` por ponto (inclusive pontos NA) | Texto do nome (`P1`); num ponto NA o rótulo existe (`HasDataLabel = True`, texto `P2`) mas com posição indefinida (`Left = NaN`) e **não aparece no PNG** (mesmo comportamento do Gantt na Fase 2). Trocar o grupo de um ponto ao vivo: o rótulo some de uma série e aparece na outra. |
| `Axes(xlCategory).HasTitle = True; AxisTitle.Formula = "=Dados!$B$10"` | Título "Preço"; mudar a célula para "Preço (R$)" muda o título: ligado e vivo. Idem no eixo Y e nas bolhas. |
| `Series.Trendlines.Add(xlLinear)` | Aceito em dispersão e bolhas (tracejada por `Format.Line.DashStyle`); gera uma entrada de legenda "Linear (A)". |
| `Axes(xlCategory).CrossesAt = 25` | `Crosses = xlAxisCrossesCustom`, lido de volta 25: a linha **vertical** (eixo de valores) passa a cruzar em x = 25. `Axes(xlValue).CrossesAt = 0,3`: a linha **horizontal** (eixo X) cruza em y = 0,3. `Crosses = xlAxisCrossesAutomatic` desfaz. Valor fora do intervalo: o eixo cola na borda. Com `TickLabelPosition = xlTickLabelPositionLow` os rótulos dos eixos ficam na borda, não junto da linha. => quadrantes nas médias sem série auxiliar nem forma. |
| `xlBubble` com `BubbleSizes` por grupo; `BubbleScale = 60` | OK. `ShowNegativeBubbles = False`; tamanho negativo: bolha não desenhada (`Width = 0`) mas o rótulo continua: o bloco converte tamanho não positivo em `NA()`. |
| Série XY de linha dentro de um gráfico de bolhas | O Excel converte o gráfico inteiro para dispersão (tipo 75): bolhas **não** combinam com outros tipos; por isso os quadrantes usam `CrossesAt` e não séries auxiliares. |

### 3.3 Corte de eixo
| Medida | Resultado |
|---|---|
| Colunas empilhadas (10, 20, 230) com `MaximumScale = 50` | A coluna de 230 é desenhada até a borda superior da área de plotagem (ponto `Top = −724`, `Height = 975`: cortada pelo plot); PNG confirma. Nas barras o ponto cortado tem `Width = 1561` para `InsideWidth = 390`. |
| Rótulos nativos dos segmentos cortados (`Center`) | Empurrados para a borda de cima do gráfico (`Top = −4`): texto "200" e "30" amontoados sobre a coluna. Rótulo `OutsideEnd` de uma coluna agrupada cortada: idem (`Top = −4`). => os rótulos de segmentos acima do corte precisam ser ocultados e o do segmento atravessado posto na base (`InsideBase`). |
| Auxiliar `tqq_total` com valor `MIN(total, BreakAt)` no bloco e rótulo ligado ao total real | Ponto no topo do plot (`Top = InsideTop`), rótulo "230" acima da coluna cortada (`Top = −4`: precisa de folga acima da área interna, como o Mekko faz com 24 pt). Um auxiliar **acima** do máximo também é desenhado no topo (`Top = −870`) com o rótulo em −4: o `MIN` é o que garante a posição. |
| Marca de corte: duas linhas diagonais (branca de 3 pt por baixo, cinzas de 1 pt por cima) atravessando a coluna 8 pt abaixo do topo | Visível e legível no PNG. |

### 3.4 PowerPoint (instância criada pela sonda; nenhum PowerPoint do usuário aberto)
| Medida | Resultado |
|---|---|
| `Application.Visible = msoFalse` | Recusado ("Hiding the application window is not allowed"); `Presentations.Add(WithWindow:=msoFalse)` cria uma apresentação sem janela (`Windows.Count = 0`, `ActiveWindow`/`ActivePresentation` inexistentes). Com `WithWindow:=msoTrue`, `ActiveWindow.View.Slide` e `View.Paste` funcionam. Slide padrão 960×540 pt; 11 layouts (o 1.º é "Slide de Título"; usamos `ppLayoutBlank` = 12 por índice de `PpSlideLayout`). |
| `Chart.CopyPicture(xlScreen, xlPicture)` no Excel **invisível** + `Slide.Shapes.Paste` | Imagem (`Type = msoPicture` 13) de 460,7×280,6 pt. `Chart.Export` PNG + `Shapes.AddPicture`: idem (6 KB). |
| `ChartObject.Copy` + `Slide.Shapes.Paste` (pasta salva) | Gráfico do PowerPoint (`Type = msoChart` 3, `HasChart`), `ChartData.IsLinked = True`, **4 formas dentro do gráfico preservadas** (as `tqq_ann_*`), 4 séries (inclusive as auxiliares `tqq_*`). `PasteSpecial(ppPasteDefault)`: igual. |
| `PasteSpecial(ppPasteOLEObject, Link:=msoTrue)` | Objeto OLE vinculado (`Type = msoLinkedOLEObject` 10) com `LinkFormat.SourceFullName = <arquivo>!Dados![probe.xlsx]Dados brk`; sem `Link`: OLE incorporado (7); `ppPasteEnhancedMetafile`: imagem. |
| Pasta **não salva** | `Shapes.Paste` cria um gráfico com `IsLinked = True` apontando para uma pasta sem arquivo (vínculo que não sobrevive); `PasteSpecial(ppPasteOLEObject, Link)` falha ("The specified data type is unavailable"). => o modo vinculado exige pasta salva. |
| `ChartData.BreakLink` logo após colar | `IsLinked` continua `True` na leitura seguinte (comportamento não confiável para um modo "incorporado": não entregue; o usuário usa "Quebrar vínculo" no PowerPoint). |
| `GetObject(, "PowerPoint.Application")` sem PowerPoint aberto | Erro 429 (padrão COM); `CreateObject` abre uma instância nova. Como o PowerPoint é aplicativo de instância única, qualquer `CreateObject` anexa ao PowerPoint já aberto do usuário: os testes só criam a instância quando **nenhum** `POWERPNT` está em execução. |
| Redimensionar o gráfico colado a 80% da largura do slide com `LockAspectRatio` | 768×467 pt, centralizado (`Left = 96`). Com o limite de altura da §4.4 (80% de 540 = 432 pt), o gráfico padrão de 460×280 fica limitado pela altura: 709,7×432 pt. |

### 3.5 Certificado por usuário (sem administrador)
| Medida | Resultado |
|---|---|
| `New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=..." -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -CertStoreLocation Cert:\CurrentUser\My -NotAfter ...` | Criado em 1,4 s; EKU "Assinatura do Código". `Export-Certificate` (.cer, 794 bytes) e `Export-PfxCertificate` (.pfx, 2,7 KB) OK. |
| `Import-Certificate -FilePath x.cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher` | Importado em 0,03 s **sem diálogo**; presente na loja. Cadeia sem a raiz: `UntrustedRoot` (esperado: o certificado autoassinado precisa ir também para `CurrentUser\Root`, importação que o Windows confirma por diálogo e por isso é passo manual/`-TrustRoot`). |
| Remoção (`Remove-Item Cert:\...\<thumbprint> -DeleteKey`) | OK em `My` e `TrustedPublisher`. Mas `New-SelfSignedCertificate` grava também uma cópia em `Cert:\CurrentUser\CA` (Autoridades de Certificação Intermediárias), que a sonda não removeu (sobra a apagar à parte, fora do plano): são **três** lojas a limpar; `New-ThinQQCert.ps1` remove a cópia de `CA` logo depois de criar e o `Test-Signing.ps1` limpa as três pelo assunto. |
| `SELFCERT.EXE` | Existe em `<Office>\root\Office16\SELFCERT.EXE` (Click-to-Run 32 bits). |
| `signtool.exe` / SIPs do Office (`msosip.dll`, `msosipx.dll` registrados em `CryptSIPDllGetSignedDataMsg`) / administrador | **Ausentes** nesta máquina; usuário sem administrador. Logo a assinatura por linha de comando não é testável aqui; o script `tools/Sign-ThinQQ.ps1` detecta as três condições e explica como obtê-las. |

### 3.6 `image=` na faixa (xlsx mínimo com customUI14, instância própria visível, aba lida por UI Automation e captura de tela)
| Variante | Resultado |
|---|---|
| `image="tqq_test"` em `button size="large"`, `button` pequeno e `menu size="large"` + `customUI/images/tqq_test.png` + `customUI/_rels/customUI14.xml.rels` (Id = nome, Type image) + `Default Extension="png"` | A aba aparece e **os ícones são desenhados** (captura: barras azuis no botão grande, no pequeno e no menu). |
| A mesma sem a relação `.rels` | A aba aparece, os controles ficam **sem ícone** (mesma degradação de um `imageMso` inválido medida na Fase 2): o pacote errado não derruba a aba. |
| PNG gerado pelo gravador puro (zlib "stored", CRC-32 e Adler-32 próprios) | 32×32 = 4.196 bytes, 16×16 = 1.108 bytes; bytes idênticos em duas gravações (determinístico); `System.Drawing` abre e lê os pixels esperados; o Excel renderiza. |

### 3.7 Fórmula de arredondamento por maior resto (só `ROUND`, `SUMPRODUCT`, `POWER`, `IF`; sem `COUNTIF`, sem matriz CSE)
| Caso (valores; d = casas) | Resultado |
|---|---|
| 0,335 / 0,335 / 0,33; d = 2 (nearest daria 1,01) | 0,33 / 0,34 / 0,33 (soma 1,00): o empate de restos é desfeito pela ordem das colunas. |
| 1,5 / 1,5 / 1,5; d = 0 (total 4,5 → 5; nearest 6) | 1 / 2 / 2. |
| 33,3333 / 33,3333 / 33,3334; d = 0 (nearest 99) | 33 / 33 / 34. |
| 1234 / 2345 / 3456; d = −3 (milhares; total 7035 → 7000; nearest 6000) | 1000 / 2000 / 4000. |
| −1,5 / 2,5 / 2,5; d = 0 | −2 / 3 / 3 (sem ajuste: já fecha em 4). |
| 0,333 / 0,333 / 0,334; d = 2 (participações com `12%`) | 0,33 / 0,33 / 0,34. |
| 1/3 / 1/3 / 1/3 (três participações exatamente iguais); d = 2 e d = 3 | 0,34 / 0,33 / 0,33 e 0,334 / 0,333 / 0,333: no empate total a **primeira** coluna recebe a unidade (dedução da fórmula: `RD = j`; não medido na sonda, coerente com o caso 0,335 / 0,335 / 0,33, em que a primeira perde a unidade). |
| 10 / 20 / 30; d = 0 | inalterados. |
| texto na origem | `#VALOR!` => as famílias saneiam antes (`V_j = IF(ISNUMBER(src), src, 0)`). |
Fórmula por célula (843 caracteres; ver §4.3): `=ROUND(v,D)+POWER(10,-D)*(IF(K>0,--(RD<=K),0)-IF(K<0,--(RA<=-K),0))`.

## 4. Decisões por item

Critérios (os mesmos da Fase 2): (1) o gráfico continua correto sem o ThinQQ; (2) cabe na arquitetura (bloco em `_ThinQQ`,
auxiliares `tqq_*`, rótulos ligados por `DataLabel.Formula`, formas `tqq_ann_*` só para o que não é dado); (3) Excel 2016+
32/64 bits, sem `Declare`; (4) testável em Excel invisível; (5) pt-BR: formatos de rótulo via `NumberFormatLocal` (`Style_LocalizeCode`),
fórmulas em inglês via `.Formula`.

### 4.1 Borboleta (`butterfly`, família `butterfly`, horizontal)
- **Tabela**: categoria + **exatamente 2** colunas de valores (lado esquerdo, lado direito); outra quantidade => erro de validação 6011
  "A borboleta precisa de exatamente 2 colunas de valores (lado esquerdo e lado direito)." (aviso, nada criado). Valores negativos,
  texto e erro valem 0 (`MAX(0, ...)`), como na pizza.
- **Gráfico**: `xlBarStacked` (categorias invertidas por `Style_ApplyBase`), 4 séries na ordem `tqq_gapL`, `<cabeçalho 1>`, `tqq_gapR`,
  `<cabeçalho 2>`; as duas auxiliares sem preenchimento criam o vão central; rótulos do eixo de categorias ocultos; vão 50%.
- **Bloco (corpo)**: `1 Categoria | 2 Total (A+B) | 3 TotalPos (Gap+B) | 4 A | 5 B | 6 Gap | 7 NegGap | 8 NegA | 9 CatPos (0) | 10 Pos`; linha
  de meta: `E = GapK` (número gravado por VBA), `F = SymL` (`=-(Gap+MAX(A,B))`), `G = SymR` (`=+...`), `H:I = n/2` (Y dos dois pontos
  de simetria). `Gap = GapK * MAX(A_1..B_n)` em todas as linhas.
- **Vão central**: `GapK` vem de `Bfly_RefreshGap`: `Lw` = maior largura estimada dos nomes (`Labels_EstimateWidth`) + 8; o vão em pontos
  é medido na escala real do eixo (`Geom_ChartValueToX`) e comparado com `Lw`: `GapK` cresce se o vão não cabe e encolhe se sobra mais de 60%
  (até 3 passadas; mínimo 0,15) e só é gravado quando muda mais de 2% (gravar dispara o Calculate do gráfico, que reentra: guarda `mGapBusy`). Chamado no Insert
  e em `Core_SafeRefresh` (redimensionar/mudar nomes). Sem o ThinQQ: `k` fica fixo e o vão continua proporcional ao máximo (vivo).
- **Auxiliares XY** (grupo secundário, `Chart_AddHelperSeries`): `tqq_sym` (2 pontos em `SymL:SymR` × `H:I`, primeiro: a escala automática
  passa a ser simétrica e viva, §3.1), `tqq_cat` (x = CatPos, y = Pos; rótulos ligados à coluna 1, `Center`, 9 pt cinza), `tqq_total`
  (x = TotalPos, y = Pos; rótulo ligado ao Total, negrito, `Right`; alternador Totais). `Chart_SetupSecondaryAxes cht, n` ao final
  (o `tqq_sym` de 1 linha o teria fixado em 0..1).
- **Rótulos**: nativos nas séries A e B (`InsideEnd`, fonte por contraste); A com o formato "absoluto" `sec0;sec0;` (a seção negativa igual
  à positiva: §3.1); `numfmt` passa por `Bfly_ApplyNumberFormat`.
- **Matriz**: totals, labels, numfmt, palette, gap, legend (right/none), samescale; não: connectors, cagr, diff, vline, ind100, secondary,
  pieout, lineend, bracket, roundsum, break, trend, quad. `Ann_GetCategoryData` ganha o caso `butterfly` (values = Total, tops =
  TotalPos) para `Deco_ScaleClipped`.
- **Sem o ThinQQ**: tudo é fórmula + séries nativas; só `GapK` deixa de se adaptar a nomes/tamanho novos.
- **Mesma escala com outros tipos** (limitação documentada): `Deco_SameScale` aplica a união [mínimo, máximo] a todos os gráficos; num grupo
  misto a borboleta perde o centro e os outros tipos ganham mínimo negativo. Use `Mesma escala` só entre borboletas.

### 4.2 Dispersão (`scatter`) e bolhas (`bubble`, família `scatter`)
- **Tabela**: `Nome | X | Y [| Grupo]` (dispersão: ≥ 3 colunas) e `Nome | X | Y | Tamanho [| Grupo]` (bolhas: ≥ 4 colunas); menos colunas
  => 6012 "Dispersão: a tabela precisa de nome, X e Y (e, nas bolhas, tamanho)." Grupos = valores distintos da coluna de grupo na ordem
  de aparição (valor da célula como texto, igual ao `=src&""` do bloco), no máximo 6 (6013 "No máximo 6 grupos (coluna de grupo)."); sem coluna de grupo, 1 série com o nome do
  cabeçalho de Y.
- **Bloco (corpo)**: `1 Nome | 2 X | 3 Y | 4 Tam | 5 Grupo | 6..5+G Xg_k | 6+G..5+2G Yg_k | 6+2G..5+3G Sg_k`; meta: `E = AvgX`, `F = AvgY`
  (`AVERAGE` das colunas X/Y saneadas; `AGGREGATE(1,6,...)` ignora `#N/A`). `X/Y = IF(ISNUMBER(src), src, NA())`; `Tam =
  IF(ISNUMBER(src), IF(src>0, src, NA()), NA())`; `Xg_k = IF(AND(Grupo="k", ISNUMBER(X), ISNUMBER(Y)[, ISNUMBER(Tam)]), X, NA())` (e `Yg`, `Sg`).
  O nome do grupo entra como literal na fórmula (aspas dobradas); um grupo novo criado depois na tabela não tem série (documentado:
  "Alterar dados").
- **Gráfico**: `xlXYScatter`/`xlBubble` (`BubbleScale = 60`, `SizeRepresents = xlSizeIsArea`, `ShowNegativeBubbles = False`),
  `DisplayBlanksAs = xlNotPlotted`, marcadores círculo 7 (dispersão), cores por `Style_PaletteRGB` em `MarkerBackground/ForegroundColor`
  (o `Format.Fill` do marcador XY não pega a cor: medido na Fase 2); eixos **visíveis** (linha cinza, rótulos 9 pt cinza, sem grade), títulos
  ligados por `AxisTitle.Formula` aos cabeçalhos X e Y; rótulos por ponto ligados à coluna Nome (`Right` na dispersão, `Center` com fonte
  por contraste nas bolhas); legenda à direita com os grupos (`Deco_SetLegend` right/none).
- **Linha de tendência** (`trend=1`): `Trendlines.Add(xlLinear)` por série de dados (tracejada, cor da série); entradas de legenda das
  tendências apagadas (`Scatter_HideTrendEntries`: as últimas T entradas). **Quadrantes** (`quad=1`): `Axes(xlCategory).CrossesAt = AvgX`
  (linha vertical) e `Axes(xlValue).CrossesAt = AvgY` (linha horizontal) (§3.2), eixos tracejados; `Scatter_RefreshAfterDataChange` reaplica as médias em cada
  `Core_SafeRefresh`; desligar devolve `Crosses = xlAxisCrossesAutomatic`. Sem o ThinQQ: as médias ficam no último valor aplicado
  (como o intervalo do Gantt).
- **Matriz**: labels, palette, legend, trend, quad; não: totals, connectors, numfmt, gap, cagr, diff, vline, ind100, secondary, pieout,
  lineend, bracket, samescale, roundsum, break. `Labels_RefreshPositions`, `Labels_ApplyCellFormat` e `Style_ApplyNumberFormat` saem cedo
  na família `scatter`.

### 4.3 Arredondamento consistente (`roundsum`) — decisão (a), sem UDF
- **Por que não UDF**: uma função definida num suplemento aberto é referenciada nas células como `=TQQ_ROUND(...)` enquanto o suplemento
  está aberto; quando a pasta é aberta sem ele, o Excel **reescreve** a referência como vínculo externo `'C:\...\versions\1.2.0\ThinQQ.core.xlam'!TQQ_ROUND(...)`
  (ou mostra `#NOME?`), e o caminho muda a cada versão do core (`versions\<x.y.z>`), o que quebraria mesmo quem tem o ThinQQ após uma
  atualização. Também o nome `TQQ.ROUND` com ponto não é um nome VBA válido. O princípio "sem o ThinQQ não pode quebrar" exclui a UDF.
- **Fórmula** (maior resto, por célula; `D` = casas decimais em célula do bloco, `u = POWER(10,-D)`, `VR` = linha de valores saneados,
  `VP` = faixa parcial da primeira coluna até a própria, `T` = alvo: `SUM(VR)` nos 100% (a soma das participações: 1, ou 0 numa categoria vazia, que assim não ganha um "1%" fantasma), a célula Total nos demais; no empate total a **primeira** coluna recebe a unidade, §3.7):
  ```
  K   = ROUND((ROUND(T,D) - SUMPRODUCT(ROUND(VR,D))) / u, 0)
  rem = (v - ROUND(v,D)) / u
  RD  = SUMPRODUCT(--((VR-ROUND(VR,D))/u > rem)) + SUMPRODUCT(--((VP-ROUND(VP,D))/u = rem))     (posição em ordem decrescente de resto; empate pela ordem)
  RA  = SUMPRODUCT(--((VR-ROUND(VR,D))/u < rem)) + SUMPRODUCT(--((VP-ROUND(VP,D))/u = rem))
  Rnd = ROUND(v,D) + u * (IF(K>0, --(RD<=K), 0) - IF(K<0, --(RA<=-K), 0))
  ```
  Verificada em §3.7 (inclusive d negativo para milhares/milhões e negativos). Casas por formato: `int` 0, `dec1` 1, `dec2` 2,
  `pct0` 2, `pct1` 3, `mil` −3, `mi` −5.
- **Onde fica**: colunas acrescentadas **à direita do bloco** na primeira vez que o alternador é ligado (`Block_Extend`: o nome
  `tqq_<id>` passa a cobrir as colunas novas; a área à direita de um bloco é sempre livre porque os blocos se empilham a partir da
  coluna A): `RndD | V_1..V_S | Rnd_1..Rnd_S`. `V_j` = `IF(ISNUMBER(src),src,0)` (empilhadas), `=Part_j` (100%), `=Val_j` (área).
  Por isso `Stacked_PosCol/VLineCol`, `Area_PosCol/VLineCol` e `Area_SeriesCount` passam a localizar as colunas **pelo cabeçalho**
  (`Block_FindCol`), com o cálculo posicional de hoje como reserva (blocos antigos). Nada muda nos blocos de quem não liga o recurso.
- **Rótulos**: com `roundsum=1`, `Stacked_SetLabels` liga os rótulos de segmento às células `Rnd_j` (`Labels_LinkPoints`, `Center`,
  fonte por contraste) em vez do rótulo nativo (empilhadas) ou de `Part_j` (100%); `Area_SetLabels` liga `tqq_lbl_j` a `Rnd_j`. O
  rótulo do total continua ligado ao Total: pelo formato ele mostra `ROUND(T,D)`, que é a soma dos `Rnd_j`. Desligar religa como antes
  (as colunas ficam). Mudar o formato: `Labels_ApplyCellFormat` atualiza `RndD` e o formato das células `Rnd_j`.
- **Tipos**: stacked, stacked100, barstacked, barstacked100, area, area100. Fora: clustered/linha (sem total a fechar), combo (a linha
  segue o formato da célula), cascata (os totais são barras), Mekko (bloco de 5n linhas e `Mekko_SeriesCount` posicional), pizza
  (rótulos nativos de percentual: §11), borboleta.
- **"Arredondar intervalo..."** (`Core_RoundRange`): com uma faixa de uma coluna ou uma linha (≥ 2 células) selecionada, pede as casas
  decimais (`InputBox` tipo 1) e a célula de destino (`InputBox` tipo 8) e grava, por `Round_WriteFormulas(src, dest, decimals)`, uma
  fórmula por célula referenciando só a faixa de origem (as casas entram como literal). Em build `Test` o comando sai antes do `InputBox`;
  `Round_WriteFormulas` é testada diretamente. Mensagens exatas: validação `"Selecione uma única coluna ou linha só com números (pelo menos 2)."`;
  sucesso `"Fórmulas de arredondamento com soma consistente gravadas em <endereço>."`; alternador fora da matriz (7002) `"'Rótulos com soma consistente'
  não se aplica a este tipo de gráfico."`.

### 4.4 Copiar para o PowerPoint (`modPowerPoint`, `Ppt_*`)
- `Ppt_GetApp(create)`: `GetObject(, "PowerPoint.Application")` → `CreateObject` (só se `create`); gancho de teste `Ppt_SetAppForTest(obj)`; em build `Test`, `Ppt_GetApp` só devolve a instância injetada (sem ela, `Nothing`: um teste nunca anexa ao PowerPoint do usuário).
  Erro 429 em ambos => "O PowerPoint não está instalado neste computador (ou não pôde ser iniciado)."
- `Ppt_TargetSlide(app)`: com janela (`Windows.Count > 0`) → `ActiveWindow.View.Slide`; sem janela mas com apresentação →
  último slide da primeira apresentação; sem apresentação → `Presentations.Add(msoTrue)` + `Slides.Add(1, ppLayoutBlank)`. Depois do aviso de
  sucesso (para a caixa do Release não ficar atrás do PowerPoint), `app.Visible = msoTrue` e `app.Activate` (o usuário vê o resultado); no build `Test`, não.
- **Imagem**: `cht.CopyPicture xlScreen, xlPicture` + `slide.Shapes.Paste` → `Ppt_Fit` (80% da largura do slide, proporção travada, centralizado;
  se a altura passar de 80% do slide, limita pela altura).
- **Gráfico vinculado**: exige `wb.Path <> ""` (senão "Salve a pasta de trabalho antes de copiar o gráfico vinculado: o vínculo aponta
  para o arquivo salvo."); `co.Copy` + `slide.Shapes.Paste` (gráfico do PowerPoint com `ChartData.IsLinked`, anotações preservadas: §3.4);
  se `Not wb.Saved`, o aviso de sucesso acrescenta "A pasta tem alterações não salvas: o PowerPoint lê o arquivo salvo nas atualizações do vínculo."
- Mensagem de sucesso: "Gráfico copiado para o slide N de '<apresentação>' como imagem/gráfico vinculado." Erros de COM durante a colagem:
  `Errors_HandleErr` com o contexto. Sem `Declare`, sem referência ao PowerPoint, constantes numéricas (`ppLayoutBlank = 12`, `msoTrue = -1`).
- **Testes**: funções puras (`Ppt_FitSize`, `Ppt_LinkedPrecheck(wb)`, `Ppt_ErrorText`) e um teste de integração que **só** cria o PowerPoint
  quando `Ppt_IsRunning() = False`, não há processo `POWERPNT` (a ROT não vê um PowerPoint aberto para incorporação) e `CreateObject` funciona (senão sai verde com registro no log): apresentação sem janela, colagem dos dois
  modos, contagem de formas/`HasChart`, `Quit` (só se ninguém mais abriu uma apresentação nessa instância). A colagem usa a área de transferência do Windows (limitação documentada).

### 4.5 Assinatura digital e política corporativa
- **O que é automatizável**: a assinatura do projeto VBA por linha de comando existe e é suportada pela Microsoft (SignTool x86 + Office SIPs,
  §2.2), mas exige o Windows SDK e o registro das DLLs SIP com administrador; a verificação da assinatura também passa pelo SIP. Nenhuma das
  três condições existe na máquina de desenvolvimento (§3.5) e a conta não tem administrador; o **modelo de objetos do VBE não assina** (só a
  interface Ferramentas > Assinatura Digital). Decisão: entregar o roteiro automatizado como script que detecta os pré-requisitos e, sem eles,
  sai com código 2 e instruções (não é falha de build); a release ganha o gancho opcional `-SignPfx` (assina `dist\*.xlam` **antes** de
  calcular os SHA-256 do manifesto) e o roteiro manual pelo VBE (assinatura V3) fica documentado.
- **Scripts (`tools/`)**: `New-ThinQQCert.ps1` (cria o certificado autoassinado de assinatura de código em `Cert:\CurrentUser\My`, exporta
  `.cer`/`.pfx`, importa o `.cer` em `Cert:\CurrentUser\TrustedPublisher`; `-TrustRoot` importa também em `Cert:\CurrentUser\Root` avisando do
  diálogo do Windows), `Trust-ThinQQPublisher.ps1` (para a TI/usuário: importa um `.cer` dado em TrustedPublisher e, opcionalmente, Root),
  `Sign-ThinQQ.ps1` (procura `signtool.exe` x86 nos Windows Kits e a SIP registrada, nas visões de 64 bits e `WOW6432Node` do registro; assina com `/fd SHA256 /f <pfx> [/p senha] [/tr <timestamp> /td SHA256]`;
  `-WhatIf` só relata). `tests/ps/Test-Signing.ps1`: ciclo completo do certificado com assunto único e remoção pelo assunto nas três lojas `My`, `CA` e `TrustedPublisher` (§3.5), num processo próprio com `exit` explícito, `Sign-ThinQQ.ps1 -WhatIf`
  reporta corretamente a ausência dos pré-requisitos, e `dist\ThinQQ.xlam` não muda (hash conferido; a assinatura real, quando as ferramentas existem, é feita numa cópia).
- **Política corporativa (README/tutorial, seção para a TI)**: (1) certificado interno de assinatura de código emitido pela AC da empresa e
  assinatura V3 pelo VBE (ou SignTool+SIP num servidor de build); (2) distribuição do certificado em "Editores Confiáveis" por GPO; (3) ou
  local confiável corporativo (`%LOCALAPPDATA%\ThinQQ` já é registrado pelo instalador); (4) "Confiar em todos os editores" não existe
  como opção: o equivalente é a política "Trust access/Trusted Publishers" + bloqueio de macros não assinadas; (5) após cada release o
  `.xlam` precisa ser **reassinado** (o build gera arquivos novos) e os SHA-256 do manifesto são os dos arquivos assinados.

### 4.6 Corte de eixo (`break`)
- **Alternativas**: (a) compressão real da escala (think-cell): impossível num eixo nativo; (b) valores recortados no bloco com séries
  apontando para o bloco em vez da tabela: muda a construção de todas as famílias e a leitura de `Series.Values` em `Ann_GetCategoryData`;
  (c) **máximo fixo do eixo** (o Excel corta as colunas na borda: §3.3) + rótulo de total em `MIN(TotalPos, BreakAt)` + marca de corte
  (formas) + rótulos de segmento sincronizados. Decisão: (c).
- **Bloco**: coluna `LblPos = MIN(TotalPos, $BreakAt)` acrescentada à direita (`Block_Extend`) na primeira vez; `BreakAt` na linha de meta
  da própria coluna `LblPos` (dentro do bloco: a limpeza a apaga junto e `Block_Extend` a vê como parte do bloco, então a soma consistente
  pode ser ligada depois do corte; gravado por VBA; vazio = sem corte, e `MIN(x, vazio) = x`). `tqq_total` passa a usar `LblPos` quando ela existe (`Stacked_SetTotals`
  e `Stacked_EnsureRoomAfterChange` leem `Block_FindCol(blk, "LblPos")`).
- **`Break_Set(cht, limit)`**: valida `0 < limit < maior topo` (senão 7003 "Corte de eixo: informe um valor maior que zero e menor que o
  maior valor do gráfico."); grava `vbreak` (texto neutro) e `BreakAt`; limpa `vmin/vmax` (Mesma escala e corte são excludentes:
  `Deco_SameScale` ignora gráficos com `vbreak`); `MaximumScale = limit`; folga de 24 pt acima da área interna para o rótulo do total
  (`Break_EnsureRoom`, em `Chart_EnsureRoomAfterChange`); `Break_SyncLabels` (segmento com base ≥ limite: sem rótulo; segmento atravessado:
  `InsideBase`; demais: `Center`; agrupadas: ponto cortado em `InsideBase`); `Break_SyncMarks` (uma marca `tqq_ann_break_<i>` por
  categoria com topo > limite: linha branca de 3 pt e duas cinzas de 1 pt, `kind=break;cat=i`, posicionadas por `Ann_LayoutBreak`;
  barras: marca vertical). `Break_Set(cht, 0)` remove tudo e devolve o eixo ao automático. `Core_SetAxisBreak` (InputBox com sugestão
  = 1,2 × segundo maior topo; em build `Test` sai antes do InputBox), `Core_SetAxisBreakAt(limit)` e `Core_RemoveAxisBreak`.
- **Refresh**: `Break_RefreshAfterDataChange` em `Core_SafeRefresh` (reaplica o máximo se o eixo voltou ao automático, sincroniza rótulos
  e marcas). **Tipos**: stacked, barstacked, clustered, barclustered. **Sem o ThinQQ**: o máximo fixo e `LblPos` são vivos; rótulos de
  segmento e marcas ficam como na última sincronização (documentado). Setas CAGR/diferença: `Ann_GetCategoryData` limita o topo
  ao corte (as setas ancoram na borda do corte, não acima da área de plotagem).

### 4.7 Ícones próprios na faixa
- `build/lib/Ribbon-Icons.ps1`: `Write-TqqPng` (RGBA 8 bits, zlib com blocos "stored", CRC-32/Adler-32 próprios: bytes determinísticos em
  qualquer runtime; validado em §3.6) e a tabela `Get-TqqIconSpecs` (nome → retângulos numa grade de 16 × 16 e `imageMso` de reserva).
  Ícones (32 e 16 px): `tqq_columns`, `tqq_bars`, `tqq_waterfall`, `tqq_lines`, `tqq_more`, `tqq_butterfly`, `tqq_scatter`, `tqq_bubble`,
  `tqq_options`, `tqq_arrows`, `tqq_ppt`, `tqq_round`, `tqq_break`. O XML 1.2.0 tem 16 atributos `image=` com 14 nomes distintos (`tqq_round16` e
  `tqq_scatter16` repetem-se; `tqq_lines` em 32 px no menu e `tqq_lines16` no alternador): 14 PNGs e 14 relações no pacote.
- `Add-RibbonToPackage` ganha o parâmetro `-Icons Png|Mso`: `Png` (padrão) gera os PNGs, grava `customUI/images/<nome>.png`,
  `customUI/_rels/customUI14.xml.rels` e `Default Extension="png"`; `Mso` reescreve a cópia do XML trocando cada `image="x"` por
  `imageMso="<reserva>"` (o fonte não muda). `build.ps1 -RibbonIcons Mso` expõe a escolha (fallback documentado).
- `Test-Ribbon.ps1`: todo `image=` do XML tem PNG no pacote, relação com `Id` igual ao nome e `Target` `images/<nome>.png`, e o content type
  `png`; variante `Mso`: XML do pacote igual ao fonte convertido por `Convert-TqqRibbonToMso`, sem imagens. `TestRibbon.bas`: nomes `image` em `^tqq_[a-z0-9_]+$`, nenhum controle com `image` e `imageMso`
  ao mesmo tempo, `imageMso` restantes válidos por `GetImageMso` (teste existente) e as 13 reservas da tabela também (`Test_Ribbon_IconFallbackMso_AreValid`). `getImage` dispensado (imagens estáticas).

### 4.8 Pasta de demonstração (`exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx`)
- Gerada por `tools/New-DemoWorkbook.ps1` (PowerShell + COM, Excel novo e invisível por `New-ExcelApp`, core de **Teste** aberto de uma cópia em
  `C:\temp` — em build Test `Core_Notify` só registra, o que impede um `MsgBox` de travar a automação — e chamadas à API real via
  `Application.Run` (`Core_Insert` com a seleção feita pelo script, `Core_SetToggle`, `Core_SetOption`, `Core_AddAnnotation`, `Core_SameScale`,
  `Core_SetAxisBreakAt`, `Round_WriteFormulas`); salva como `.xlsx` (sem macros). Dados fictícios determinísticos (gerador congruencial com
  semente fixa): rede "Rede Demo ThinQQ - Materiais de Construção (fictícia)" (nome claramente fictício) com 8 lojas em cidades/regiões diferentes, 8 departamentos, 24 meses de vendas,
  custo, margem, unidades, ticket médio, metas por loja, mix B2C/B2B e cronograma de obras/expansão (Gantt). Nenhuma pessoa, empresa real ou
  dado pessoal.
- Um exemplo por funcionalidade visível (20 tipos + recursos: totais, rótulos, legenda nenhuma e nome no fim da linha, formato, espaço,
  paleta, CAGR, diferença, colchete, linha de valor, indicador 100%, mesma escala, eixo secundário, rótulos fora da pizza, Mekko % e unidades,
  escala do Gantt em semanas e em meses, soma consistente, corte de eixo, tendência, quadrantes), cada um na sua planilha com legenda
  (o que mostra e o menu que o cria) e "Voltar ao índice" em A1; planilha 1 "Indice" com uma tabela (nº, funcionalidade, tipo, descrição)
  e `Hyperlinks.Add` com `SubAddress` por linha. A lista de cobertura (`tools/DemoCoverage.ps1`) é a mesma que o gerador e o teste usam.
- `tests/ps/Test-DemoWorkbook.ps1`: Excel próprio; primeira planilha "Indice"; todo hiperlink tem destino; toda planilha-destino tem ≥ 1 gráfico;
  nº de exemplos = tamanho da lista de cobertura, que por sua vez cobre todos os tipos e todos os recursos da matriz de `modChartKinds` (lidos do código, não da própria lista) e a lista
  fixa de recursos visíveis; nenhuma célula com erro (fora da `_ThinQQ`); `_ThinQQ` muito oculta; arquivo < 3 MB; `docProps` sem autor (o gerador grava
  com `RemovePersonalInformation`; o `Test-Privacy` não lê binários). Roda num processo próprio, com `exit` explícito.

### 4.9 Documentação de instalação
- README: seção "Instalação" (requisitos; baixar o ZIP das Releases; descompactar a pasta inteira; fechar o Excel; `Instalar-ThinQQ.cmd`;
  abrir o Excel e procurar a aba ThinQQ; como as atualizações chegam; desinstalar; instalação manual; os 5 problemas mais comuns: macros
  bloqueadas / arquivo bloqueado pelo Windows (Propriedades > Desbloquear), Central de Confiabilidade, local confiável, antivírus/proxy,
  "aba não aparece"), e "Ambientes corporativos (TI)" (§4.5). `Test-Docs` confere títulos e frases-chave.

## 5. Tabela de decisões

| Decisão | Escolha | Alternativa rejeitada | Por quê |
|---|---|---|---|
| Eixo simétrico da borboleta | auxiliar XY `tqq_sym` nos dois extremos (vivo) | VBA fixando `MinimumScale/MaximumScale` a cada refresh | medido: os pontos XY entram na escala automática; sem o ThinQQ continua simétrico |
| Vão central da borboleta | `GapK` (VBA) × máximo (fórmula) | 18% fixo | nomes longos invadem as barras (§3.1) |
| Séries por grupo (dispersão) | colunas `Xg/Yg/Sg` por grupo com `NA()` fora do grupo | uma série e cores por ponto | legenda por grupo e marcadores por grupo (think-cell); `xlNotPlotted` esconde os rótulos dos pontos NA |
| Quadrantes | `CrossesAt` nas médias | séries auxiliares de linha / formas | bolhas não combinam com XY (§3.2); formas não vivem sem o ThinQQ |
| Arredondamento | fórmula nativa por maior resto no bloco + botão para faixas | UDF `TQQ_ROUND` | vínculo externo ao `.xlam` quando aberto sem o ThinQQ e a cada troca de versão; `#NOME?` |
| Onde ficam as colunas de arredondamento | à direita do bloco, sob demanda (`Block_Extend`) | bloco separado `tqq_<id>_r` | o scan da limpeza para no primeiro rótulo que cita a `_ThinQQ`: um 100% sem totais ficaria só com o bloco auxiliar e perderia o principal |
| PowerPoint "gráfico" | `Copy` + `Paste` (gráfico do PowerPoint vinculado) | OLE vinculado (`ppPasteOLEObject`) / "incorporado" via `BreakLink` | o gráfico do PowerPoint mantém as anotações e é editável; OLE abre o Excel dentro do slide; `BreakLink` não refletiu em `IsLinked` (§3.4) |
| Assinatura | scripts de certificado + roteiro SignTool/SIP + gancho `-SignPfx` + manual VBE | assinar pelo VBE por automação de interface | não há API; UI Automation do VBE é frágil e não roda sem sessão interativa |
| Quebra de eixo | corte por máximo fixo + marca + `LblPos` | séries recortadas no bloco | muda a construção de todas as famílias; o corte por eixo é vivo e reversível |
| Ícones | PNG gerado no build (gravador puro) | PNGs versionados / `System.Drawing` | fonte em texto; bytes determinísticos; `System.Drawing` não garante bytes iguais entre runtimes |
| Demonstração | gerador PowerShell + core de Teste + API real | montar à mão | reprodutível; cobre a API como teste de integração |

## 6. Matriz de capacidades 1.2.0 (`Kinds_Supports`)

Tipos: os 17 da Fase 2 + `butterfly` (família `butterfly`, horizontal), `scatter` e `bubble` (família `scatter`). Recursos novos: `roundsum`,
`break`, `trend`, `quad`.

| feature | tipos |
|---|---|
| totals | stacked, stacked100, barstacked, barstacked100, area, area100, combo, mekko, mekkounits, **butterfly** |
| labels | todos |
| connectors | stacked, stacked100, barstacked, barstacked100, waterfall, barwaterfall, combo |
| numfmt | todos exceto gantt, scatter, bubble |
| palette | todos |
| gap | stacked, clustered, stacked100, barstacked, barclustered, barstacked100, waterfall, barwaterfall, combo, gantt, **butterfly** |
| legend | todos exceto gantt (cascata só com várias séries) |
| lineend | line, combo |
| ind100 | stacked100, barstacked100, area100, mekko |
| vline | stacked, clustered, line, barstacked, barclustered, area, combo, mekkounits |
| secondary | combo |
| pieout | pie |
| cagr | stacked, clustered, stacked100, line, waterfall, barstacked, barclustered, barstacked100, barwaterfall, area, area100, combo |
| diff | stacked, clustered, line, waterfall, barstacked, barclustered, barwaterfall, area, combo |
| bracket | gantt |
| samescale | stacked, clustered, line, waterfall, barstacked, barclustered, barwaterfall, area, combo, mekkounits, **butterfly** |
| **roundsum** | stacked, stacked100, barstacked, barstacked100, area, area100 |
| **break** | stacked, clustered, barstacked, barclustered |
| **trend** | scatter, bubble |
| **quad** | scatter, bubble |

`Kinds_FeatureLabel`: `roundsum` → "Rótulos com soma consistente", `break` → "Corte de eixo", `trend` → "Linha de tendência", `quad` →
"Quadrantes (médias)". Opções persistidas novas: `roundsum`, `trend`, `quad` (`0|1`), `vbreak` (texto neutro; vazio = sem corte).

## 7. Faixa de opções (loader 1.2.0)

```
[ThinQQ]
┌ Inserir gráfico ──────────────────────────────────────────────────┐ ┌ Gráfico selecionado ──────────────────────────────────────────────────────┐ ┌ ThinQQ ────────┐
│ Colunas ▾  Barras ▾  Cascata ▾  Linha e área ▾  Mais tipos ▾       │ │ [x] Totais [x] Rótulos [ ] Conectores | Formato ▾ Paleta ▾ Espaço ▾ Legenda ▾ │ │ Exemplos       │
│ (ícones próprios tqq_*)             Mekko (%), Mekko (unid.), Gantt,│ │ Mais opções ▾ (Indicador 100%, Linha de valor (média), Eixo secundário (linha),│ │ Tutorial       │
│                                     Pizza, Rosca, Borboleta,        │ │   Rótulos fora da pizza, Rótulos com soma consistente, Linha de tendência,    │ │ Verificar at.  │
│                                     Dispersão, Bolhas               │ │   Quadrantes (médias) | Corte de eixo..., Remover corte de eixo |             │ │ Configurações  │
│                                                                     │ │   Arredondar intervalo... | Mesma escala, Escala automática)                  │ │ Sobre          │
│                                                                     │ │ Setas ▾ (CAGR, diferença, Colchete (Gantt) | Remover setas)                   │ └────────────────┘
│                                                                     │ │ PowerPoint ▾ (Copiar como imagem, Copiar como gráfico vinculado)               │
│                                                                     │ │ Atualizar   Alterar dados                                                     │
└─────────────────────────────────────────────────────────────────────┘ └───────────────────────────────────────────────────────────────────────────────┘
```
Contagens que os testes verificam: 8 menus, 20 botões de inserção (`Rx_OnInsert`), 10 toggles, 4 dropdowns, 3 botões de anotação,
callbacks novos `Rx_OnAxisBreak`, `Rx_OnAxisBreakOff`, `Rx_OnRoundRange`, `Rx_OnPptImage`, `Rx_OnPptChart`. Regras herdadas: ids válidos, sem
`size` em filho de menu, `getEnabled` por `tag` (matriz) ou sem tag (há gráfico); "Arredondar intervalo..." fica sempre ativo (age sobre
células), como Mesma escala.

## 8. Blocos novos (resumo; detalhes nas tarefas)

| Família | Corpo | Meta (colunas E..) |
|---|---|---|
| butterfly | `Categoria, Total, TotalPos, A, B, Gap, NegGap, NegA, CatPos, Pos` (n + 1 linhas) | `GapK, SymL, SymR, SymY1, SymY2` |
| scatter | `Nome, X, Y, Tam, Grupo, Xg_1..Xg_G, Yg_1..Yg_G, Sg_1..Sg_G` | `AvgX, AvgY` |
| stacked/area + roundsum | colunas da família + `RndD, V_1..V_S, Rnd_1..Rnd_S` (à direita, sob demanda) | — |
| stacked (incl. clustered) + break | colunas da família + `LblPos` (à direita, sob demanda) | `BreakAt` na linha de meta da coluna `LblPos` |

`Block_FindCol(blk, header)`, `Block_Extend(wb, id, extraCols)` e `Block_SourceRange(cht)` entram em `modCalcSheet`.

## 9. Versões, compatibilidade e documentação

- `version.json` → `{"loader":"1.2.0","core":"1.2.0"}` (a faixa mudou). Loader 1.2.0 com core 1.1.0: botões novos avisam "Tipo de gráfico
  desconhecido" ou caem em `Core_HasChart` (procedimentos ausentes devolvem `Empty`); core 1.2.0 com loader 1.1.0 funciona sem os controles
  novos. CHANGELOG `[1.2.0]`; `release/notes.txt` (6 a 8 linhas); `LEIA-ME.txt` com `versions\1.2.0\`.
- Tutorial: seções novas `#borboleta`, `#dispersao`, `#arredondamento`, `#powerpoint`, `#corte`, `#instalacao-ti`; matriz `#matriz` com 20 tipos ×
  20 recursos (uma tabela a mais, `class="mx"`); limite de tamanho 160 KB. README: "Instalação" completa, "Ambientes corporativos (TI)",
  "Recursos 1.2.0", verificações manuais 27+, limitações. `Test-Docs`: 20 tipos, 20 recursos, seções e mensagens novas, título "Instalação".

## 10. Riscos e mitigações

| Risco | Prob. | Mitigação |
|---|---|---|
| Vão da borboleta estreito demais com fonte de tema mais larga | média | `Labels_EstimateWidth` superestima de propósito; `GapK` refeito a cada refresh; mínimo 0,15 |
| Escala automática arredonda e os dois lados ficam ligeiramente diferentes | baixa | os dois pontos de simetria têm o mesmo módulo: o Excel escolhe o mesmo "número bonito" dos dois lados (medido −500/500, −800/800) |
| Grupo novo na coluna de grupo depois de inserir (dispersão) | certa | documentado: "Alterar dados" recria as séries; `Core_ChartSourceMismatch` não detecta (mesmo n) |
| Rótulos de dispersão sobrepostos em nuvens densas | média | limite do Excel (sem fuga de sobreposição); documentado |
| Fórmula de arredondamento longa (≈ 850 caracteres por célula) deixa o recálculo lento em tabelas grandes | baixa | S ≤ 6 séries × n categorias; medido em §3.7 só para n pequeno; teste da Tarefa 3 mede 30×6 |
| Extensão do bloco à direita quebra funções posicionais | média | `Block_FindCol` por cabeçalho nas famílias afetadas; testes antigos de contagem de colunas continuam válidos (a extensão só ocorre ao ligar o recurso) |
| `MIN(x, célula vazia)` | baixa | `MIN` ignora células vazias em referências (padrão Excel); teste da Tarefa 4 confere `LblPos = TotalPos` sem corte |
| PowerPoint do usuário aberto durante os testes | certa (às vezes) | o teste de integração só roda quando nenhum `POWERPNT` existe; produção usa `GetObject` primeiro (anexa ao do usuário de propósito) |
| Assinatura automatizada não verificada | certa | `Sign-ThinQQ.ps1 -WhatIf` e `Test-Signing.ps1` cobrem a detecção; a assinatura real fica para o ambiente da TI com SDK + SIP + certificado |
| `image=` quebra a aba em alguma versão do Office | baixa | medido em 2019: sem a relação só o ícone some; fallback `-RibbonIcons Mso` |
| Pasta de demonstração gerada com o core de Teste traz ganchos? | nenhuma | o `.xlsx` não tem código; só gráficos e fórmulas |
| Tempo da suíte (~9–10 min ao fim da fase, estimado no pré-flight; limite prático ~8 min) | alta | testes de tipo com fixtures pequenas; `-Filter` na iteração; `-TimeoutSec 900` basta; se passar de 10 min: juntar `Insert` e `Geometry` da borboleta e da dispersão (o mesmo gráfico) e conferir a demonstração direto do zip, sem Excel (~1 s em vez de ~20 s) |

## 11. O que ficou de fora (e por quê)

- Compressão real da escala (quebra de eixo do think-cell): não existe em gráfico nativo; o corte entregue é a alternativa documentada.
- Modo "gráfico incorporado" no PowerPoint (`BreakLink` não confiável, §3.4): o usuário usa "Quebrar vínculo" no PowerPoint.
- Legenda de tamanho de bolha, partições arrastáveis, tendências não lineares, mais de 6 grupos, log scale.
- Arredondamento consistente na pizza/rosca (rótulos nativos de percentual) e no Mekko/cascata/combo.
- UDF `TQQ.ROUND`.
- Linha de valor em valor arbitrário; ordem de segmentos/fatias; rótulos de CAGR por série; rótulos no fim da linha em áreas; Gantt (fins de
  semana, Responsável/Observação, âncoras, setas de processo, escala de dias); rosca com vários anéis; série "Outros"; mesma escala dinâmica;
  setas em Mekko; "100%=" do Mekko; rótulo de intervalo nas barras do Gantt.
- Assinatura do projeto por automação **nesta máquina** (sem SDK/SIP/administrador): roteiro e scripts entregues; execução fica para a TI.

## 12. Pré-flight do plano (varredura de coerência)

Pares de tarefas que compartilham arquivo/interface (o que uma produz, o que a outra consome, resultado) e autoconsistência de cada tarefa.
Corrigido até todas as linhas ficarem limpas; um pré-flight independente posterior achou e corrigiu o que as linhas marcadas "corrigido no pré-flight" descrevem.

| Par | Produz | Consome | Resultado |
|---|---|---|---|
| T0 → T1 | `Kinds_IsValid("butterfly")`, `Kinds_Family = "butterfly"`, `Kinds_IsHorizontal` inclui, matriz, `Kinds_Insert` (7200 provisório), `Block_FindCol`, `Style_PaletteRGB`, `Labels_*`/`Style_*` cientes da família | T1 troca o `Case "butterfly"` de 7200 por `Bfly_Insert` e remove o teste temporário `Test_Core_Insert_NewKindsNotReady` só na T2 (último tipo novo) | limpo: T1 mantém o teste temporário para `scatter`/`bubble`; T2 o remove |
| T0 → T2 | idem + família `scatter` nas exclusões de `Labels_RefreshPositions`, `Labels_ApplyCellFormat`, `Style_ApplyNumberFormat` | `Scatter_Insert(ws, data, kind)` no `Case "scatter"` | limpo |
| T0 → T3 | `Block_Extend`, `Block_FindCol`, `Block_SourceRange`; `Stacked_PosCol/VLineCol`, `Area_PosCol/VLineCol/SeriesCount` por cabeçalho; recurso `roundsum` na matriz e `Kinds_FeatureLabel` | `Round_EnsureColumns` estende o bloco; `Stacked_SetLabels`/`Area_SetLabels` leem `roundsum` | limpo |
| T0 → T4 | `Block_Extend`, `Block_FindCol`; recurso `break` | `Break_EnsureColumn` cria `LblPos`; `Stacked_SetTotals` usa `Block_FindCol(blk,"LblPos")` (alterado na T4, não na T3) | limpo: T3 não toca `Stacked_SetTotals`; corrigido no pré-flight: `BreakAt` na meta da coluna `LblPos` (a célula `J` ficava fora do bloco de 5 ou 6 colunas e a soma consistente ligada depois do corte dava 6014) |
| T1 → T8 | tipo `butterfly` válido | botão `btnButterfly tag=butterfly` | limpo |
| T2 → T8 | tipos `scatter`, `bubble`; toggles `trend`, `quad` | botões/toggles correspondentes; `Test_Ribbon_Tags_MatchCoreKeys` consulta o core real | limpo |
| T3 → T8 | `Core_RoundRange`, toggle `roundsum` | `btnRoundRange`, `tglRoundSum` | limpo |
| T4 → T8 | `Core_SetAxisBreak`, `Core_RemoveAxisBreak` | `btnBreak`, `btnBreakOff` (`Rx_OnAxisBreak`, `Rx_OnAxisBreakOff`) | corrigido no pré-flight: o cabeçalho `break` da matriz continua `class="k"` ("Corte de eixo" não é rótulo da faixa) |
| T5 → T8 | `Core_CopyToPowerPoint(mode)` | `mnuPpt` com `Rx_OnPptImage`/`Rx_OnPptChart` | limpo |
| T6 → T8 | `Add-RibbonToPackage -Icons`, nomes de ícones `tqq_*`, `Test-Ribbon.ps1` exige PNG para cada `image=` (nomes distintos: `Sort-Object -Unique` no build e no teste, por isso um nome repetido não duplica entradas no pacote) | o XML da T8 usa só nomes da tabela `Get-TqqIconSpecs` (16 atributos, 14 nomes distintos, 13 ícones) | limpo: a tabela da T6 lista os 13 nomes que a T8 usa |
| T6 → T7 | `build.ps1 -RibbonIcons` | `release.ps1 -SignPfx` chama o build e depois `Sign-ThinQQ.ps1` | limpo (parâmetros independentes) |
| T3/T4 → T9 | `Round_SetLabels` via `Core_SetToggle "roundsum"`; `Core_SetAxisBreakAt` | exemplos 22 a 25 em "Exemplos 3" | corrigido no pré-flight: no empate total (1/3 cada) a **primeira** coluna recebe a unidade (34/33/33); testes, título e Review Focus acertados |
| T1/T2/T3/T4/T5/T8 → T10 | API completa 1.2.0 (`Core_Insert` dos 20 tipos, toggles, `Core_SetAxisBreakAt`, `Round_WriteFormulas`, `Core_SameScale`) | gerador da demonstração via `Application.Run` | limpo: o gerador não usa `Core_CopyToPowerPoint` nem `InputBox` |
| T8 → T11 | rótulos exatos do XML | tutorial (`<span class="ui">`), README (`ThinQQ > Rótulo`), `Test-Docs` | limpo |
| T10 → T11 | `exemplos/*.xlsx` e `Test-DemoWorkbook.ps1` | README cita a pasta e o item manual; `run-tests.ps1` chama o teste (T10) | limpo |
| T7 → T11 | scripts `tools/*.ps1` e o resultado da assinatura | README "Ambientes corporativos (TI)" | limpo |
| T11 → T12 | `version.json` 1.2.0, docs | suíte completa, release de ensaio | limpo |

| Tarefa | Autoconsistência (testes × código, arquivos criados × tocados) | Resultado |
|---|---|---|
| T0 | testes em `TestKinds`/`TestCalcSheet`/`TestCoreApi` só usam funções listadas em Interfaces; `Test-Docs` atualizado para 20/20 | corrigido no pré-flight: cabeçalhos `Mid<k>` da área gravados depois dos `Val<k>`; limite do tutorial (160 KB) já na T0 |
| T1 | `TestButterfly.bas` só usa `Bfly_*`, `Kinds_*`, `Chart_FindSeries`, `Geom_*`; `FixtureFor`/`Test_Core_AllKinds` ganham `butterfly` | limpo |
| T2 | `TestScatter.bas` idem com `Scatter_*`; `Style_PaletteRGB` da T0 | limpo |
| T3 | `TestRounding.bas` usa `Round_*`, `Block_FindCol`; `Labels_ApplyCellFormat` atualiza `RndD` | corrigido no pré-flight: expectativas do empate (primeira coluna), série procurada pelo índice, alvo `SUM` nos 100% |
| T4 | `TestAxisBreak.bas` usa `Break_*`, `Core_SetAxisBreakAt`; `Ann_RemoveAll` não remove marcas | corrigido no pré-flight: `ElseIf` num `If` de uma linha (não compilava); `BreakAt` dentro do bloco |
| T5 | `TestPowerPoint.bas` usa `Ppt_*`, `Core_CopyToPowerPoint`; integração condicional | corrigido no pré-flight: o gráfico padrão é limitado pela altura (709,7×432); guardas de teste do PowerPoint |
| T6 | `Test-Ribbon.ps1` e `TestRibbon.bas` referem-se ao XML **atual** (sem `image=` ainda): os testes de imagem são tolerantes a zero `image=` (`>= 0` imagens, todas consistentes) e a T8 os aperta para `>= 13` | corrigido no pré-flight: a função auxiliar `R` (sombreada pelo alias embutido `r`) virou `New-TqqRect` |
| T7 | `Test-Signing.ps1` cria e remove o próprio certificado (assunto único); nada em `HKLM`; `release.ps1` só chama o script com `-SignPfx` | corrigido no pré-flight: cópia em `CA` removida, saída capturada com `*>&1`, processo próprio com `exit` explícito, SIP também na `WOW6432Node` |
| T8 | contagens do XML (8/20/10/4/3) batem com o XML descrito; `Test-Docs` passa com o tutorial parcial (rótulos novos acrescentados na mesma tarefa) | corrigido no pré-flight: limite do tutorial movido para a T0; `k` → `ui` sem o cabeçalho `break` |
| T9 | `TestSamples` conta 6 + 13 + 8 gráficos (planilhas 1, 2 e 3); `Test-Docs` exige que tutorial e README citem o maior número de exemplos de `Samples_Create` (`$nEx`): T9 troca "18 exemplos" por "26 exemplos" nos dois (o conteúdo novo fica na T11) | limpo |
| T10 | gerador e teste leem `tools/DemoCoverage.ps1`; o teste não regenera | corrigido no pré-flight: literais de linhas sem vírgula inicial, `return ,$rng`, recursos lidos da matriz, `docProps` sem autor |
| T11 | `Test-Docs`: seções, mensagens, versões, instalação; ensaio de release na cópia | limpo |
| T12 | nenhum arquivo de produção; auditoria e relatório | limpo |

## 13. Fontes

- think-cell — dispersão e bolhas: https://www.think-cell.com/en/resources/manual/scatterbubble
- think-cell — arredondamento: https://www.think-cell.com/en/resources/manual/exceldatarounding e https://www.think-cell.com/en/resources/kb/0028
- think-cell — eixos (quebra de eixo, mesma escala): https://www.think-cell.com/en/resources/manual/axes
- think-cell — links Excel/PowerPoint: https://www.think-cell.com/en/resources/manual/exceldatalinks e https://www.think-cell.com/en/resources/manual/manage-data-links
- think-cell — rótulos: https://www.think-cell.com/en/resources/manual/textlabels
- Borboleta em Excel (técnica): https://www.exceldemy.com/create-butterfly-chart-in-excel/ ; https://www.vizwiz.com/2020/12/butterfly-chart.html
- Microsoft — Office SIPs para assinar projetos VBA: https://www.microsoft.com/en-us/download/details.aspx?id=56617 ; https://support.microsoft.com/topic/2b8b3cae-ad64-4b4b-aa85-c4a98ca6da87
- Microsoft — assinatura V3: https://devblogs.microsoft.com/microsoft365dev/upgrade-signed-office-vba-macro-projects-to-v3-signature/ ; https://support.microsoft.com/en-us/servicing/office/hotfix/upgrade-signed-office-vba-macro-projects-to-v3-signature-kb5000676
- Microsoft — assinar o projeto de macro (VBE, SelfCert): https://support.microsoft.com/office/digitally-sign-your-macro-project-956e9cc8-bbf6-4365-8bfa-98505ecd1c01 ; https://msdn.microsoft.com/en-us/library/aa141471.aspx
- Keyfactor — SignTool x86 com os SIPs do Office: https://docs.keyfactor.com/signum/latest/using-signum-with-office-macros
- customUI — imagens próprias (`image`, `customUI/images`, `.rels`): https://bettersolutions.com/vba/ribbon/images-editor-custom.htm ; https://learn.microsoft.com/en-us/answers/questions/917971/custom-image-for-group-in-custom-ribbon-xml
- Microsoft — `ChartData.IsLinked`: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.chartdata.islinked ; `PasteSpecial`: https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa212207(v=office.11)
- PPT FAQ — vínculos e objetos OLE do Excel: https://pptfaq.com/FAQ00935_Convert_linked_Excel_charts_to_embedded_Excel_charts.htm
- Maior resto (método): https://pypi.org/project/largest-remainder/ ; https://www.mrexcel.com/board/goto/post?id=4381665
