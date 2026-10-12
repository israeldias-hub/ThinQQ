# Lista de cobertura da pasta de demonstracao: a MESMA fonte para o gerador (tools/New-DemoWorkbook.ps1) e o teste (tests/ps/Test-DemoWorkbook.ps1).
# Cada item vira a planilha E<Num> com um exemplo; Kind = tipo do ThinQQ (modChartKinds), Features = recursos visiveis no exemplo,
# Menu = o caminho na faixa ThinQQ que cria o exemplo, Desc = o que observar no grafico. Os textos aparecem na pasta (portugues).

# Recursos visiveis que a demonstracao tem de mostrar (alem dos tipos): o teste confere que cada um aparece em Features de algum item.
function Get-TqqDemoFeatures {
    return @('totals', 'labels', 'legend-none', 'lineend', 'numfmt', 'gap', 'palette', 'cagr', 'diff', 'bracket', 'vline', 'ind100', 'samescale',
             'secondary', 'pieout', 'gantt-weeks', 'gantt-months', 'roundsum', 'roundrange', 'break', 'trend', 'quad', 'connectors')
}

# Nome do tipo como aparece no menu "Inserir gráfico" da faixa.
function Get-TqqDemoKindName([string]$Kind) {
    $names = @{
        'stacked' = 'Colunas empilhadas'; 'clustered' = 'Colunas agrupadas'; 'stacked100' = 'Colunas 100%'; 'line' = 'Linha'
        'waterfall' = 'Cascata (colunas)'; 'barstacked' = 'Barras empilhadas'; 'barclustered' = 'Barras agrupadas'; 'barstacked100' = 'Barras 100%'
        'barwaterfall' = 'Cascata (barras)'; 'area' = 'Área empilhada'; 'area100' = 'Área 100%'; 'combo' = 'Combinação (colunas + linha)'
        'mekko' = 'Mekko (%)'; 'mekkounits' = 'Mekko (unidades)'; 'gantt' = 'Gantt'; 'pie' = 'Pizza'; 'doughnut' = 'Rosca'
        'butterfly' = 'Borboleta'; 'scatter' = 'Dispersão'; 'bubble' = 'Bolhas'
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
        @{ Num = 2;  Kind = 'clustered';     Features = @('gap')
           Menu = 'Inserir gráfico > Colunas > Colunas agrupadas; Gráfico selecionado > Espaço > Muito largo (120%)'
           Title = 'Vendas por loja, 2024 x 2025 (colunas agrupadas, espaço muito largo)'
           Desc = 'O espaço entre os grupos de colunas foi alargado para 120% (as agrupadas nascem com 80%).' }
        @{ Num = 3;  Kind = 'stacked100';    Features = @('ind100')
           Menu = 'Inserir gráfico > Colunas > Colunas 100% (o Indicador 100% já vem ligado; Mais opções > Indicador 100% liga e desliga)'
           Title = 'Mix de canais B2C/B2B por região (colunas 100% e indicador 100%)'
           Desc = 'Cada coluna é a participação de cada canal nas vendas da região; o indicador "100%" marca o topo comum das colunas.' }
        @{ Num = 4;  Kind = 'line';          Features = @('lineend')
           Menu = 'Inserir gráfico > Linha e área > Linha; Legenda > Nome no fim da linha'
           Title = 'Vendas e custo mensais da loja Aurora Centro (nome no fim da linha)'
           Desc = 'Em vez da legenda, o nome de cada série aparece ao lado do último ponto da linha.' }
        @{ Num = 5;  Kind = 'waterfall';     Features = @('connectors')
           Menu = 'Inserir gráfico > Cascata > Cascata (colunas) (Conectores já vêm ligados)'
           Title = 'Da receita ao lucro em 2025 (cascata com conectores)'
           Desc = 'As linhas finas ligam o topo de cada degrau ao próximo. Na tabela, a letra "e" pede uma coluna de total calculada (EBITDA e Lucro).' }
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
           Desc = 'Participação de cada canal nas vendas de cada loja (24 meses).' }
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
           Desc = 'A mesma tabela do exemplo 10, em participação de cada departamento no total do trimestre: pisos e revestimentos ganham espaço.' }
        @{ Num = 12; Kind = 'combo';         Features = @('secondary')
           Menu = 'Inserir gráfico > Linha e área > Combinação (colunas + linha); Mais opções > Eixo secundário (linha)'
           Title = 'Vendas por canal e margem % da loja Aurora Centro (combinação com eixo secundário)'
           Desc = 'Colunas B2C + B2B e a linha da margem % (última coluna da tabela) em escalas próprias: o eixo secundário liga sozinho quando as grandezas diferem muito. Os totais foram desligados para não cruzar com a linha.' }
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
           Title = 'Reformas do trimestre (Gantt em semanas, com colchete)'
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
           Menu = 'Inserir gráfico > Mais tipos > Rosca; Paleta > Tons de cinza'
           Title = 'Vendas 2025 por região (rosca, paleta em tons de cinza)'
           Desc = 'A paleta "Tons de cinza" troca as cores de todas as séries de uma vez (útil para impressão).' }
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
           Desc = 'A seta mostra a taxa de crescimento anual composta entre os totais das duas colunas. A tabela soma a planilha Dados (SOMASES).' }
        @{ Num = 23; Kind = 'stacked100';    Features = @('roundsum', 'roundrange')
           Menu = 'Inserir gráfico > Colunas > Colunas 100%; Mais opções > Rótulos com soma consistente; Mais opções > Arredondar intervalo...'
           Title = 'Participações que somam 100% (soma consistente no gráfico e na tabela)'
           Desc = 'Os rótulos de cada coluna somam exatamente 100%. Abaixo, "Arredondar intervalo" grava fórmulas que fazem o mesmo numa tabela comum: compare com o arredondamento simples.' }
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
    )
    foreach ($i in $items) { $i.Sheet = ('E{0:00}' -f $i.Num); $i.KindName = Get-TqqDemoKindName $i.Kind }
    return $items
}
