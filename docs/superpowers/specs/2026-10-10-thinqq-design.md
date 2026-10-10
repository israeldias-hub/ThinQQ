# ThinQQ — Pesquisa e Design (Fase 1)

Data: 2026-10-10
Autor: Israel Dias (com apoio do Claude Code)
Status: proposta para revisão antes da implementação

## 1. Objetivo

Criar o **ThinQQ**, um suplemento para Excel inspirado no think-cell, para montar gráficos de
apresentação (cascata, colunas empilhadas com totais, setas de CAGR e diferença, conectores,
rótulos limpos) em poucos cliques, com:

- compartilhamento simples com colegas (sem instalador administrativo, sem servidor);
- design limpo e intuitivo (aba própria na faixa de opções, opções contextuais ao gráfico selecionado);
- tutorial de uso embutido e acessível pela faixa;
- aviso de novas versões dentro do Excel e atualização por download sem sair do Excel.

## 2. O que o think-cell faz (pesquisa)

Fontes consultadas em 2026-10-10: manual oficial do think-cell (páginas de anotações, cascata,
Mekko, rótulos e eixos), página de produto, e material de terceiros para técnicas Excel.

### 2.1 Tipos de gráfico

| think-cell | Descrição | ThinQQ |
|---|---|---|
| Colunas/barras empilhadas | Séries empilhadas, totais automáticos, conectores entre segmentos | Fase 1 (colunas), Fase 2 (barras) |
| Colunas/barras agrupadas | Séries lado a lado | Fase 1 (colunas), Fase 2 (barras) |
| 100% empilhado | Eixo percentual, indicador de 100% | Fase 1 (colunas) |
| Linha / área / combinação | Linha com rótulos de ponto | Fase 1 (linha); área e combinação Fase 2 |
| Cascata (waterfall/bridge) | Marcador `e` para totais calculados, conectores, cruza o zero, build-up/build-down | Fase 1 (vertical, 1 série); horizontal e multi-segmento Fase 2 |
| Mekko (% e unidades) | Largura das colunas proporcional ao total; variante com eixo em unidades | Fase 2 |
| Gantt | Atividades, marcos, escala de datas | Fase 2 |
| Pizza / rosca | Rótulos internos/externos | Fase 2 |
| Borboleta (butterfly) | Barras espelhadas | Fase 3 |
| Dispersão / bolhas | Dois eixos de valor | Fase 3 |
| Fluxo (chevrons), agenda, tabelas | Layout de slide, não gráfico | Fora do escopo (é PowerPoint) |

### 2.2 Anotações e decorações (manual "Chart annotations")

- **Setas de diferença**: *level difference* (entre segmentos/pontos) e *total difference* (entre
  totais de colunas); modos: seta simples com diferença relativa, seta dupla com diferença absoluta.
  Atualizam sozinhas quando os dados mudam.
- **Setas de CAGR**: da categoria inicial até a final, taxa composta anual; usa datas das categorias
  quando existem (convenção 30/360 para frações de ano).
- **Rótulos de CAGR por série**.
- **Linhas de valor** (linha horizontal em um valor; encaixa na média).
- **Indicador 100%** em gráficos 100%.
- **Conectores de série** (equivalente às "linhas de série" do Excel) e **conectores universais**.
- **Lacunas de categoria** para abrir espaço para setas.

### 2.3 Rótulos (manual "Labels")

- Tipos: categoria, série, segmento, ponto, total, total geral, interno/externo (pizza).
- Posicionamento automático com fuga de sobreposição; cor de fonte automática contrastante.
- Formato numérico por exemplo (`1.234,5`), separadores configuráveis, prefixo/sufixo, sinais
  forçados (+), negativos entre parênteses, magnitude (×10⁶).
- Rótulos ligados aos dados: atualizam ao mudar a planilha.

### 2.4 Eixos (manual "Axes")

- Escala automática com ajuste manual; **mesma escala** entre vários gráficos; quebra de eixo;
  escala logarítmica; lacunas de categoria; eixo de datas.

### 2.5 Cascata (manual "Waterfall")

- Positivos sobem, negativos descem; `e` ("equals") marca uma coluna de total/subtotal calculada;
  `e` na primeira coluna calcula de trás para frente; várias segmentos por coluna; conectores
  entre colunas adjacentes; segmentos exibem sempre a extensão positiva (sinal via formato).

### 2.6 think-cell no Excel

- Links de dados Excel → PowerPoint; **think-cell round** (arredondamento com soma consistente).
- Importante para o ThinQQ: no Excel, o diferencial do think-cell é "gráfico ligado aos dados e
  que se mantém formatado" — exatamente o que o ThinQQ entrega nativamente, pois os gráficos
  gerados são gráficos Excel comuns ligados por fórmulas.

## 3. Decisão de tecnologia

### 3.1 Ambiente verificado nesta máquina (2026-10-10)

| Item | Valor |
|---|---|
| Excel | 16.0.19127 **x86 (32-bit)**, SKU ProPlus2019Retail |
| Windows | 11 Pro 10.0.26200 |
| PowerShell | 7.6.6 (pwsh) e 5.1 (powershell) |
| Git | presente |
| Node.js | presente (usado só para servir fixtures de teste HTTP) |
| .NET SDK | **ausente** |
| Trust Center "Acesso ao modelo de objeto VBA" | já habilitado (`AccessVBOM=1`) |
| Locale | pt-BR |

### 3.2 Alternativas consideradas

| Opção | Prós | Contras | Veredito |
|---|---|---|---|
| **VBA add-in (.xlam)** | Roda em qualquer Excel Windows 2016+; compartilha por cópia de arquivo; sem hospedagem; gráficos resultantes são nativos (quem recebe o arquivo não precisa do ThinQQ); atualização por download de arquivo; acesso total ao modelo de objetos (eventos de gráfico, formas dentro do gráfico, posição da área de plotagem) | UI limitada a faixa de opções + UserForms; testes exigem automação COM; sem Mac | **Escolhido** |
| Office.js (web add-in) | UI HTML moderna, atualização automática, Mac/web | Precisa hospedagem HTTPS; sideload por usuário ou deploy central por TI; Excel 2019 perpétuo não está mais na lista de plataformas suportadas e a API de formas (ExcelApi 1.9) pode não existir nele; proxies corporativos | Rejeitado para esta base |
| VSTO / Excel-DNA (C#) | Ribbon + WPF, ClickOnce tem auto-update | Exige .NET SDK/Visual Studio (ausente), certificado para ClickOnce, runtime VSTO | Rejeitado |

### 3.3 Restrições do Windows/Office que moldam o design

- **Mark of the Web / macros bloqueadas**: arquivos baixados por navegador/e-mail têm macros
  bloqueadas por padrão. Solução: o instalador copia os arquivos para `%LOCALAPPDATA%\ThinQQ`,
  remove o MOTW (`Unblock-File`) e registra a pasta como **Local Confiável** no Trust Center
  (chave HKCU, sem admin). Downloads feitos pelo próprio atualizador (XMLHTTP + ADODB.Stream) não
  recebem MOTW.
- **Registro do suplemento sem diálogo**: chave `HKCU\Software\Microsoft\Office\16.0\Excel\Options`,
  valor `OPEN`/`OPENn` = `/R "caminho\ThinQQ.xlam"`. Excel precisa estar fechado ao gravar.
- **VBScript está em descontinuação** (Windows 11 24H2 como recurso opcional; desligado por padrão
  a partir de 2027). O ThinQQ **não usa .vbs**; scripts auxiliares são PowerShell ou `.cmd`.
- **Um add-in não consegue sobrescrever a si mesmo** enquanto carregado. Daí a arquitetura em dois
  arquivos (seção 4).
- **Eventos de gráfico embutido** (`Calculate`, `Resize`, `Activate`) só existem via classe com
  `WithEvents` ligada a cada gráfico; `Calculate` só dispara se o gráfico estiver visível.
- `Chart.Shapes` permite desenhar linhas e caixas **dentro** do gráfico (movem junto com ele).
- `DataLabel.Formula` liga um rótulo a uma célula (`='_ThinQQ'!$Q$5`), mantendo o rótulo vivo.
- Fórmulas escritas por VBA usam sempre sintaxe em inglês (`Range.Formula`), independente do
  Excel pt-BR; `NumberFormat` também usa códigos em inglês (`#,##0`).

## 4. Arquitetura

```
┌──────────────────────────────┐    Application.Run     ┌──────────────────────────────┐
│ ThinQQ.xlam (loader)         │ ─────────────────────▶ │ ThinQQ.core.xlam (core)      │
│ • faixa de opções (ribbon)   │                        │ • motor de gráficos           │
│ • boot: abre o core mais novo│ ◀───────────────────── │ • anotações + eventos         │
│ • verificador/instalador de  │  Loader_InvalidateRibbon│ • API pública Core_*          │
│   atualizações (manifest)    │                        │ • exemplos                    │
│ • configurações (registro)   │                        └──────────────────────────────┘
└──────────────────────────────┘                                  ▲ substitui
        ▲ instalado via chave OPEN                                 │
        │ %LOCALAPPDATA%\ThinQQ\ThinQQ.xlam        %LOCALAPPDATA%\ThinQQ\versions\<x.y.z>\ThinQQ.core.xlam
```

- **Loader** muda raramente. Na abertura, localiza a versão mais nova em `versions\`, abre o core
  (`Workbooks.Open` de um .xlam fica invisível) e agenda a verificação de atualização
  (`Application.OnTime`, no máximo 1×/24h).
- **Core** contém toda a lógica. Pode ser trocado "a quente": o loader fecha o core antigo e abre
  o novo. O loader, por sua vez, é substituído pelo core (via `AddIn.Installed = False/True`),
  agendado por `OnTime` para não haver código do loader na pilha.
- **Manifesto de versão** (JSON plano) hospedado em URL HTTP(S) (GitHub raw / Pages) **ou** em
  pasta de rede/OneDrive sincronizado (caminho UNC ou local com `%VARIÁVEIS%`), configurável
  por registro (`HKCU\Software\VB and VBA Program Settings\ThinQQ\Updater\UpdateSource`).

### 4.1 Manifesto (`manifest.json`)

```json
{
  "schemaVersion": 1,
  "coreVersion": "1.1.0",
  "coreUrl": "https://github.com/israeldias-hub/ThinQQ/releases/download/v1.1.0/ThinQQ.core.xlam",
  "coreSha256": "…64 hex…",
  "loaderVersion": "1.0.0",
  "loaderUrl": "https://github.com/israeldias-hub/ThinQQ/releases/download/v1.1.0/ThinQQ.xlam",
  "loaderSha256": "…64 hex…",
  "tutorialUrl": "https://github.com/israeldias-hub/ThinQQ/releases/download/v1.1.0/tutorial.html",
  "publishedAt": "2026-10-10",
  "notes": "• Novo: gráfico Mekko\n• Correção: rótulos negativos na cascata"
}
```

JSON **plano** (sem objetos aninhados) para que um parser de ~60 linhas em VBA, testável, baste —
sem dependência externa.

### 4.2 Dados e cálculo

- Convenção de entrada (igual ao Excel nativo): selecionar a tabela **com** linha de cabeçalho
  (nomes das séries) e coluna de categorias. Uma única célula selecionada expande para
  `CurrentRegion`.
- Planilha **muito oculta** `_ThinQQ` na pasta do usuário guarda, por gráfico, um **bloco de
  cálculo** com fórmulas que apontam para a fonte (totais, séries auxiliares da cascata, valores de
  rótulo). O gráfico referencia o bloco; assim tudo recalcula ao mudar os dados, inclusive sem o
  ThinQQ instalado.
- Cada gráfico tem id (`tqq_xxxxxx`), usado no nome do `ChartObject`, no nome definido
  (`tqq_xxxxxx` → bloco) e nos nomes das formas de anotação (`tqq_ann_cagr_1`).
- Opções do gráfico (totais, rótulos, conectores, formato, paleta, espaço) ficam na linha meta do
  bloco (`chave=valor;…`).

### 4.3 Cascata — modelo que cruza o zero

Para cada categoria com início `s` e fim `e = s + delta` (`delta = 0` e barra de 0 a `e` quando
a célula contém marcador de total `e`, `t`, `total`, `subtotal`, `=`):

```
Lo = IF(total, MIN(0,e), MIN(s,e))      Hi = IF(total, MAX(0,e), MAX(s,e))
Base      = MAX(0,Lo) + MIN(0,Hi)          (só um dos termos é ≠ 0)
PiecePos  = MAX(0,Hi) - MAX(0,Lo)          (parte acima de zero)
PieceNeg  = MIN(0,Lo) - MIN(0,Hi)          (parte abaixo de zero, valor ≤ 0)
```

Séries empilhadas: `Base` (invisível), `UpPos`, `UpNeg`, `DownPos`, `DownNeg`, `Total`, e uma
série de linha invisível `Lbl` (posição = Hi se o valor exibido ≥ 0, senão Lo) cujos rótulos
são ligados por fórmula ao valor exibido (delta ou total). Conectores são linhas desenhadas
dentro do gráfico na altura de `End` de cada barra e reposicionadas no evento `Calculate`.

### 4.4 Anotações

- Formas dentro do gráfico (`Chart.Shapes`), parâmetros no `AlternativeText`
  (`kind=cagr;from=1;to=4`). `Ann_Refresh` recalcula posições a partir de
  `PlotArea.InsideLeft/Top/Width/Height` e da escala do eixo de valores, e reescreve o texto.
- CAGR = `(fim/início)^(1/n) − 1`, com `n` = diferença em anos quando as categorias são anos ou
  datas, senão número de períodos.
- Seta de diferença (total): seta vertical dupla à direita da coluna final, com linha de nível
  tracejada a partir da coluna inicial; rótulo com diferença absoluta e relativa.

### 4.5 Faixa de opções (loader, `customUI14.xml`)

```
[ThinQQ]
┌ Inserir gráfico ───────────────────────┐ ┌ Gráfico selecionado ───────────────────────────────────────┐ ┌ ThinQQ ─────────────────┐
│ Colunas   Colunas   Colunas   Cascata  │ │ [x] Totais  [x] Rótulos  [ ] Conectores                     │ │ Exemplos  Tutorial       │
│ empilhadas agrupadas 100%      Linha   │ │ Formato ▾  Paleta ▾  Espaço ▾   Seta CAGR  Seta diferença    │ │ Verificar atualizações   │
│                                        │ │ Remover anotações   Atualizar   Alterar dados                │ │ Configurações  Sobre     │
└────────────────────────────────────────┘ └─────────────────────────────────────────────────────────────┘ └──────────────────────────┘
```

Controles do grupo "Gráfico selecionado" ficam desabilitados sem um gráfico ThinQQ ativo
(`getEnabled`), e os toggles refletem o estado real (`getPressed`).

### 4.6 Estilo padrão (aparência think-cell)

- Sem linhas de grade; eixo de valores oculto; eixo de categorias sem marcas; área do gráfico sem
  borda; espaço entre colunas 50%; fonte 9 pt; legenda à direita (empilhadas) ou nenhuma (cascata).
- Formato numérico padrão `#,##0;-#,##0;` (oculta zeros). Paletas: *Padrão ThinQQ* (azul-marinho,
  azul, azul-claro, cinza, verde-água, âmbar), *Cinza*, *Cores do tema*.
- Totais em negrito acima da coluna (série de linha invisível com rótulos).

## 5. Distribuição, instalação e atualização

1. `release.ps1` gera `dist/ThinQQ-<versão>.zip` com `ThinQQ.xlam`, `ThinQQ.core.xlam`,
   `tutorial.html`, `manifest.json`, `Instalar-ThinQQ.cmd`, `install.ps1`, `uninstall.ps1`,
   `LEIA-ME.txt`, e atualiza `release/manifest.json` no repositório.
2. Publicação: GitHub Releases (URLs estáveis) **ou** pasta compartilhada (Teams/SharePoint
   sincronizado, rede). O `-BaseUrl` do release define as URLs do manifesto.
3. Usuário final: descompacta, fecha o Excel, dá duplo clique em `Instalar-ThinQQ.cmd`. O script
   copia os arquivos, desbloqueia, registra local confiável, registra o add-in e grava a fonte de
   atualização. Alternativa manual descrita no tutorial.
4. Atualização: ao abrir o Excel (1×/24h) o loader lê o manifesto; se houver versão nova, mostra
   o formulário "Nova versão disponível" com as novidades e os botões *Atualizar agora*, *Depois*,
   *Pular esta versão*. *Atualizar agora* baixa, confere SHA-256, instala em
   `versions\<nova>\` e recarrega o core sem reiniciar o Excel. Botão "Verificar atualizações"
   força a verificação.

## 6. Qualidade e testes

- Fonte em texto (`.bas`, `.cls`, `.vba`, `.json`, `.xml`) no Git; `build.ps1` monta os `.xlam`
  via COM (importa módulos, cria UserForms a partir de JSON, injeta o ribbon no pacote zip).
- `run-tests.ps1` compila a configuração *Test* (código + módulos `Test*.bas`), abre o Excel
  invisível, executa `RunAllTests` (mini-framework TAP: `ok`/`not ok`) e devolve código de saída.
  Testes reais criam pastas temporárias e gráficos de verdade.
- Lógica pura (versões, JSON, geometria, fórmulas da cascata) tem testes unitários; integração
  cobre builders, registro, bloco de cálculo, anotações, transporte por arquivo e HTTP local
  (servidor Node de fixtures).
- Checklist manual de QA por tarefa para o que só o olho vê (faixa, formulários, aparência).

## 7. Riscos e mitigação

| Risco | Mitigação |
|---|---|
| Política corporativa bloqueia macros/local confiável do usuário | Tutorial com passos para TI (assinatura digital com certificado interno; Fase 3 inclui `SelfCert`/assinatura) |
| `VBComponents.Add(MSForm).Designer` falhar em automação | Tarefa 1 testa a criação de formulário cedo; fallback: exportar `.frm/.frx` uma vez e versionar |
| `DataLabel.Formula` não aceitar vínculo em alguma versão | Fallback `TextFrame2.TextRange.InsertChartField msoChartFieldRange` (Excel 2013+) |
| Evento `Calculate` não disparar com gráfico fora da tela | Botão "Atualizar" e re-layout ao ativar o gráfico/planilha |
| Proxy corporativo bloqueia GitHub | Fonte de atualização configurável para pasta de rede/OneDrive sincronizado |
| Perda do ponteiro `IRibbonUI` após erro não tratado | Todos os callbacks com tratamento de erro; mensagem orienta reiniciar o Excel |

## 8. Fora do escopo da Fase 1 (planos futuros)

- Fase 2: barras horizontais (incl. cascata horizontal), Mekko (% e unidades), Gantt, pizza/rosca,
  rótulos de série no fim da linha, linhas de valor, indicador 100%, "mesma escala", cascata com
  várias séries, ícones próprios na faixa.
- Fase 3: borboleta, dispersão/bolhas, `TQQ.ROUND` (arredondamento consistente), copiar para
  PowerPoint como imagem/vinculado, assinatura digital, quebra de eixo (se viável).

## 9. Fontes

- think-cell — anotações: https://www.think-cell.com/en/resources/manual/chartdecorations
- think-cell — cascata: https://www.think-cell.com/en/resources/manual/waterfall
- think-cell — Mekko: https://www.think-cell.com/en/resources/manual/mekko
- think-cell — rótulos: https://www.think-cell.com/en/resources/manual/textlabels
- think-cell — eixos: https://www.think-cell.com/en/resources/manual/axes
- think-cell — links de dados Excel: https://www.think-cell.com/en/resources/manual/exceldatalinks
- think-cell — produto: https://www.think-cell.com/en/product
- Microsoft — eventos em gráficos embutidos: https://learn.microsoft.com/en-us/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-embedded-charts
- Microsoft — `DataLabel.Formula`: https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.formula
- Microsoft — `Chart.Shapes`: https://learn.microsoft.com/en-us/office/vba/api/excel.chart.shapes
- Microsoft — `AddIn` (carregar add-in instalado): https://learn.microsoft.com/en-us/office/vba/api/Excel.AddIn
- Microsoft — descontinuação do VBScript: https://www.bleepingcomputer.com/news/security/microsoft-to-kill-off-vbscript-in-windows-to-block-malware-delivery/
- Microsoft — conjuntos de requisitos Office.js (plataformas suportadas): https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets
- Chaves de registro de add-ins do Excel: https://bettersolutions.com/excel/add-ins/registry-keys.htm
- Macros bloqueadas / local confiável para .xlam: https://exceloffthegrid.com/unblock-macros-downloaded-from-the-internet/
- Ribbon customUI em pacote zip: https://excelguru.ca/using-customui-editor/ e https://jkp-ads.com/rdb/win/s2/win001.htm
- Mekko em Excel (Peltier): https://peltiertech.com/marimekko-charts/
- SHA-256 via COM em VBA: https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/String_Hashing_in_VBA
- UserForm por automação (`VBComponents.Add` + `Designer`): https://www.bettersolutions.com/vba/userforms/creating-dynamically.htm
