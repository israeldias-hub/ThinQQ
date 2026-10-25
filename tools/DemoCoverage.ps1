# Lista de cobertura da pasta de demonstracao: a MESMA fonte para o gerador (tools/New-DemoWorkbook.ps1) e o teste (tests/ps/Test-DemoWorkbook.ps1).
# Cada item vira a planilha E<Num> com um exemplo; Kind = tipo do ThinQQ (modChartKinds), Features = recursos visiveis no exemplo,
# Menu = o caminho na faixa ThinQQ que cria o exemplo, Desc = o que observar no grafico. Os textos aparecem na pasta (portugues).

# Recursos visiveis que a demonstracao tem de mostrar (alem dos tipos): o teste confere que cada um aparece em Features de algum item.
function Get-TqqDemoFeatures {
    return @('totals', 'labels', 'legend-none', 'lineend', 'numfmt', 'gap', 'palette', 'cagr', 'diff', 'delta', 'bracket', 'vline', 'ind100', 'samescale',
             'secondary', 'pieout', 'gantt-weeks', 'gantt-months', 'roundsum', 'roundrange', 'break', 'trend', 'quad', 'connectors',
             'pairdiff', 'pairabs', 'inbar', 'style')
}

# Nome do tipo como aparece no menu "Inserir gráfico" da faixa.
function Get-TqqDemoKindName([string]$Kind) {
    $names = @{
        'stacked' = 'Colunas empilhadas'; 'clustered' = 'Colunas agrupadas'; 'stacked100' = 'Colunas 100%'; 'line' = 'Linha'
        'waterfall' = 'Cascata (colunas)'; 'barstacked' = 'Barras empilhadas'; 'barclustered' = 'Barras agrupadas'; 'barstacked100' = 'Barras 100%'
        'barwaterfall' = 'Cascata (barras)'; 'area' = 'Área empilhada'; 'area100' = 'Área 100%'; 'combo' = 'Combinação (colunas + linha)'
        'mekko' = 'Mekko (%)'; 'mekkounits' = 'Mekko (unidades)'; 'gantt' = 'Gantt'; 'pie' = 'Pizza'; 'doughnut' = 'Rosca'
        'butterfly' = 'Borboleta'; 'scatter' = 'Dispersão'; 'bubble' = 'Bolhas'
        'waterfallpair' = 'Cascata comparativa'
    }
    if ($names.ContainsKey($Kind)) { return $names[$Kind] }
    return $Kind
}

function Get-TqqDemoCoverage {
    $items = @(
        @{ Num = 1;  Kind = 'stacked';       Features = @('totals', 'labels')
           Menu = 'Inserir gráfico > Colunas > Colunas empilhadas (Totais e Rótulos já vêm ligados)'
           Title = 'Vendas 2025 por departamento e canal (colunas empilhadas, totais e rótulos)'
           Desc = 'O total de cada coluna fica acima dela e os rótulos dos segmentos acompanham as células. A tabela soma a planilha Dados (SOMASES): mude um valor em Dados e o gráfico muda junto.' }
        @{ Num = 2;  Kind = 'clustered';     Features = @('gap', 'delta')
           Menu = 'Inserir gráfico > Colunas > Colunas agrupadas; Gráfico selecionado > Espaço > Muito largo (120%); Setas > Delta clicando nas barras (clique na 1.ª e na última barra)'
           Title = 'Vendas por loja, 2024 x 2025 (colunas agrupadas, espaço muito largo e delta entre barras)'
           Desc = 'O espaço entre os grupos de colunas foi alargado para 120% (as agrupadas nascem com 80%). A seta no alto é um delta: ele compara duas barras quaisquer, aqui a de 2024 da primeira loja com a de 2025 da última (séries e lojas diferentes), e mostra a diferença em valor e em %. Os dois cliques nas barras escolhem as duas pontas; o delta é desenhado pelo ThinQQ e, sem ele instalado, fica como está se os números mudarem.' }
        @{ Num = 3;  Kind = 'stacked100';    Features = @('ind100')
           Menu = 'Inserir gráfico > Colunas > Colunas 100% (o Indicador 100% já vem ligado; Mais opções > Indicador 100% liga e desliga)'
           Title = 'Mix de canais B2C/B2B por região (colunas 100% e indicador 100%)'
           Desc = 'Cada coluna é a participação de cada canal nas vendas da região; o indicador "100%" marca o topo comum das colunas.' }
        @{ Num = 4;  Kind = 'line';          Features = @('lineend')
           Menu = 'Inserir gráfico > Linha e área > Linha; Legenda > Nome no fim da linha'
           Title = 'Vendas e custo mensais da loja Aurora Centro (nome no fim da linha)'
           Desc = 'Em vez da legenda, o nome de cada série aparece ao lado do último ponto da linha.' }
        @{ Num = 5;  Kind = 'waterfall';     Features = @('connectors', 'diff')
           Menu = 'Inserir gráfico > Cascata > Cascata (colunas) (Conectores já vêm ligados); Setas > Seta diferença'
           Title = 'Da receita ao lucro em 2025 (cascata com conectores e seta de diferença)'
           Desc = 'As linhas finas ligam o topo de cada degrau ao próximo. Na tabela, a letra "e" pede uma coluna de total calculada (EBITDA e Lucro). A seta à direita mostra a diferença entre o total da primeira barra (Receita) e o da última (Lucro), em valor e em %. A seta é desenhada pelo ThinQQ: sem ele instalado, fica como está se os números mudarem.' }
        @{ Num = 6;  Kind = 'barstacked';    Features = @('diff')
           Menu = 'Inserir gráfico > Barras > Barras empilhadas; Setas > Seta diferença'
           Title = 'Vendas 2025 por região e canal (barras empilhadas e seta de diferença)'
           Desc = 'A seta mostra a diferença entre o total da primeira e o da última região da tabela.' }
        @{ Num = 7;  Kind = 'barclustered';  Features = @('numfmt')
           Menu = 'Inserir gráfico > Barras > Barras agrupadas; Formato > Milhares (1.234 → 1)'
           Title = 'Vendas por loja em R$ (barras agrupadas, rótulos em milhares)'
           Desc = 'A tabela está em reais; o formato "Milhares" mostra os rótulos em R$ mil sem mudar os números da tabela.' }
        @{ Num = 8;  Kind = 'barstacked100'; Features = @()
           Menu = 'Inserir gráfico > Barras > Barras 100%'
           Title = 'Mix de canais B2C/B2B por loja (barras 100%)'
           Desc = 'Participação de cada canal nas vendas de 2025 de cada loja.' }
        @{ Num = 9;  Kind = 'barwaterfall';  Features = @()
           Menu = 'Inserir gráfico > Cascata > Cascata (barras)'
           Title = 'Da receita ao lucro em 2025 (cascata horizontal)'
           Desc = 'A mesma ponte do exemplo 5, na horizontal: os nomes das etapas ficam legíveis à esquerda.' }
        @{ Num = 10; Kind = 'area';          Features = @('vline')
           Menu = 'Inserir gráfico > Linha e área > Área empilhada; Mais opções > Linha de valor (média)'
           Title = 'Vendas trimestrais dos três maiores departamentos (área empilhada e linha de valor)'
           Desc = 'A linha tracejada marca a média dos totais trimestrais, com o valor no rótulo à direita.' }
        @{ Num = 11; Kind = 'area100';       Features = @()
           Menu = 'Inserir gráfico > Linha e área > Área 100%'
           Title = 'Participação trimestral dos três maiores departamentos (área 100%)'
           Desc = 'A mesma tabela do exemplo 10, em participação de cada departamento no total do trimestre: pisos ganham espaço.' }
        @{ Num = 12; Kind = 'combo';         Features = @('secondary')
           Menu = 'Inserir gráfico > Linha e área > Combinação (colunas + linha); Mais opções > Eixo secundário (linha)'
           Title = 'Vendas por canal e margem % da loja Aurora Centro (combinação com eixo secundário)'
           Desc = 'A linha da margem % (última coluna da tabela) tem a sua própria escala, à direita, num eixo secundário que liga sozinho quando as grandezas diferem muito; o estilo ThinQQ esconde os eixos de valores, o secundário também. Com os totais das colunas ligados, a linha fica numa faixa acima deles.' }
        @{ Num = 13; Kind = 'mekko';         Features = @()
           Menu = 'Inserir gráfico > Mais tipos > Mekko (%)'
           Title = 'Regiões x grupos de produto em 2025 (Mekko %)'
           Desc = 'A largura de cada coluna é o peso da região nas vendas; a altura de cada faixa é a participação do grupo de produto na região.' }
        @{ Num = 14; Kind = 'mekkounits';    Features = @()
           Menu = 'Inserir gráfico > Mais tipos > Mekko (unidades)'
           Title = 'Regiões x grupos, largura = nº de lojas (Mekko em unidades)'
           Desc = 'A largura é o número de lojas da região (coluna Lojas) e a altura são as vendas médias por loja em 2025: a área de cada bloco é a venda total do grupo na região.' }
        @{ Num = 15; Kind = 'gantt';         Features = @('bracket', 'gantt-weeks')
           Menu = 'Inserir gráfico > Mais tipos > Gantt; Setas > Colchete (Gantt)'
           Title = 'Reformas de março a maio (Gantt em semanas, com colchete)'
           Desc = 'Cronograma curto: a escala do tempo fica em semanas. O colchete agrupa as duas primeiras obras; a linha sem data de fim vira um marco.' }
        @{ Num = 16; Kind = 'gantt';         Features = @('gantt-months')
           Menu = 'Inserir gráfico > Mais tipos > Gantt'
           Title = 'Plano de abertura de uma nova loja (Gantt em meses)'
           Desc = 'Cronograma de 18 meses: a escala do tempo passa a meses sozinha. A inauguração (sem data de fim) é um marco.' }
        @{ Num = 17; Kind = 'pie';           Features = @('pieout')
           Menu = 'Inserir gráfico > Mais tipos > Pizza; Mais opções > Rótulos fora da pizza'
           Title = 'Vendas 2025 por região (pizza, rótulos fora)'
           Desc = 'O nome e o valor de cada fatia ficam fora da pizza, legíveis mesmo na fatia pequena (Norte).' }
        @{ Num = 18; Kind = 'doughnut';      Features = @('palette')
           Menu = 'Inserir gráfico > Mais tipos > Rosca; Estilo > Tons de cinza'
           Title = 'Vendas 2025 por região (rosca, estilo em tons de cinza)'
           Desc = 'O estilo "Tons de cinza" troca as cores de todas as séries de uma vez (útil para impressão).' }
        @{ Num = 19; Kind = 'butterfly';     Features = @()
           Menu = 'Inserir gráfico > Mais tipos > Borboleta'
           Title = 'Vendas B2C x B2B por loja (borboleta)'
           Desc = 'B2C à esquerda e B2B à direita, com os nomes das lojas no centro e o total de cada loja à direita.' }
        @{ Num = 20; Kind = 'scatter';       Features = @('trend', 'quad')
           Menu = 'Inserir gráfico > Mais tipos > Dispersão; Mais opções > Linha de tendência e Quadrantes (médias)'
           Title = 'Ticket médio x margem por loja (dispersão, tendência e quadrantes)'
           Desc = 'Lojas com mais vendas para empresas (B2B) têm ticket maior e margem menor: a tendência tracejada mostra a relação e os eixos cruzam nas médias, formando quadrantes.' }
        @{ Num = 21; Kind = 'bubble';        Features = @()
           Menu = 'Inserir gráfico > Mais tipos > Bolhas'
           Title = 'Ticket x margem x vendas por departamento (bolhas)'
           Desc = 'Cada bolha é um departamento: posição = ticket médio e margem, área = vendas de 2025, cor = grupo de produto.' }
        @{ Num = 22; Kind = 'stacked';       Features = @('cagr')
           Menu = 'Inserir gráfico > Colunas > Colunas empilhadas; Setas > Seta CAGR'
           Title = 'Crescimento 2024-2025 de três departamentos (seta CAGR)'
           Desc = 'A seta mostra a taxa de crescimento anual composta entre os totais das duas colunas. A seta é desenhada pelo ThinQQ: sem ele, continua onde está se os números mudarem.' }
        @{ Num = 23; Kind = 'stacked100';    Features = @('roundsum', 'roundrange')
           Menu = 'Inserir gráfico > Colunas > Colunas 100%; Mais opções > Rótulos com soma consistente; Mais opções > Arredondar intervalo...'
           Title = 'Participações que somam 100% (soma consistente no gráfico e na tabela)'
           Desc = 'Os rótulos de cada coluna somam exatamente 100%; compare com o exemplo 13, a mesma tabela sem a soma consistente, onde Nordeste e Centro-Oeste somam 101% e Norte 99%. Abaixo, "Arredondar intervalo" grava fórmulas que fazem o mesmo numa tabela comum (compare com o arredondamento simples).' }
        @{ Num = 24; Kind = 'stacked';       Features = @('break')
           Menu = 'Inserir gráfico > Colunas > Colunas empilhadas; Mais opções > Corte de eixo...'
           Title = 'Cimento e argamassa dominam as vendas: corte de eixo'
           Desc = 'A mesma tabela do exemplo 1: a coluna muito maior é cortada (marca de corte) e o total real continua no rótulo; as demais colunas ganham altura para comparação.' }
        @{ Num = 25; Kind = 'clustered';     Features = @('samescale')
           Menu = 'Inserir gráfico > Colunas > Colunas agrupadas (duas vezes); Mais opções > Mesma escala'
           Title = 'A maior e a menor loja lado a lado (mesma escala)'
           Desc = 'Os dois gráficos usam o mesmo eixo de valores: a diferença de tamanho entre as lojas fica visível.' }
        @{ Num = 26; Kind = 'line';          Features = @('legend-none')
           Menu = 'Inserir gráfico > Linha e área > Linha; Legenda > Nenhuma'
           Title = 'Vendas mensais da loja Aurora Centro (sem legenda)'
           Desc = 'Com uma série só, a legenda é dispensável: o título da planilha já diz o que é.' }
        @{ Num = 27; Kind = 'waterfallpair'; Features = @('pairdiff', 'inbar')
           Menu = 'Inserir gráfico > Cascata > Cascata comparativa (Setas de variação e Nome dentro das barras já vêm ligados)'
           Title = 'Da receita ao lucro, 2024 x 2025 (cascata comparativa)'
           Desc = 'As duas pontes lado a lado, 2024 à esquerda e 2025 à direita, com os anos escritos dentro das barras da receita. A seta de cada etapa mostra quanto 2025 variou sobre 2024, (2025 - 2024) / 2024 (num custo, "+" quer dizer que o custo cresceu), e o lucro mostra o percentual da receita do próprio ano. Tudo acompanha os números, com ou sem o ThinQQ.' }
        @{ Num = 28; Kind = 'waterfallpair'; Features = @('pairabs')
           Menu = 'Inserir gráfico > Cascata > Cascata comparativa; Setas > Variação em valor; Legenda > Direita'
           Title = 'As duas maiores lojas: da receita ao lucro em 2025 (variação em valor)'
           Desc = 'A mesma ponte para as duas lojas que mais venderam em 2025; a seta de cada etapa mostra a diferença em R$ milhões entre a segunda loja e a primeira, e os nomes das lojas ficam na legenda.' }
        @{ Num = 29; Kind = 'waterfall';     Features = @('style', 'diff')
           Menu = 'Inserir gráfico > Cascata > Cascata (colunas); Estilo > Cascata MC; Setas > Seta diferença'
           Title = 'Da receita ao lucro (cascata no estilo Cascata MC)'
           Desc = 'A ponte do exemplo 5 no estilo embutido Cascata MC: primeira barra hachurada, totais cinza, reduções vermelhas (os aumentos ficam verdes), conectores pontilhados e seta com rótulo de fundo branco. O estilo fica no gráfico e na pasta e sobrevive a Alterar dados.' }
    )
    foreach ($i in $items) { $i.Sheet = ('E{0:00}' -f $i.Num); $i.KindName = Get-TqqDemoKindName $i.Kind }
    return $items
}
