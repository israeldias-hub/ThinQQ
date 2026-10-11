# Changelog

Formato baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/); versões no formato x.y.z.

## [Não lançado]

## [1.1.0] - 2026-10-11

### Adicionado
- Barras horizontais empilhadas, agrupadas e 100%, com a primeira categoria no topo, totais no fim das barras, rótulos, conectores e setas (menu Barras).
- Cascata horizontal (Cascata (barras)) e cascata com várias séries: uma coluna de valores por série (até 6, uma cor por série), segmentos com a extensão certa mesmo quando cruzam o zero, legenda e conectores no total líquido.
- Cascata "de cima para baixo": `e` na primeira linha e um número na última linha; o primeiro total é calculado para trás.
- Área empilhada e Área 100%, com rótulos no meio de cada faixa, e Combinação (colunas empilhadas + última coluna como linha), com eixo secundário opcional (ligado sozinho quando as grandezas diferem mais de 5 vezes).
- Legenda com três opções: Direita, Nenhuma e Nome no fim da linha (só em linha e combinação).
- Mekko (%) e Mekko (unidades): colunas de largura proporcional, sem espaço entre elas, com participação nos segmentos, largura ou total no topo e nomes das categorias embaixo.
- Gantt: barras de atividades, marcos (fim vazio), escala de datas automática (semanas, meses, trimestres ou anos), separadores, rótulo de data ou da 4ª coluna e colchetes sobre atividades (Setas > Colchete (Gantt)).
- Pizza e Rosca, com rótulos de valor ou participação, cores por fatia e, só na pizza, rótulos fora das fatias.
- Linha de valor na média dos valores (média ponderada pela largura no Mekko (unidades)), Indicador 100% (ligado por padrão nos gráficos 100% e no Mekko (%)), Mesma escala e Escala automática.
- Faixa de opções em menus por família (Colunas, Barras, Cascata, Linha e área, Mais tipos; Mais opções e Setas no grupo do gráfico selecionado).
- Exemplos: a planilha "Exemplos 2" traz 12 exemplos novos (18 no total, em duas planilhas).
- Tutorial reescrito: uma seção por tipo, com a tabela de dados de cada um, a tabela de recursos por tipo, as mensagens do ThinQQ e os limites conhecidos.

### Alterado
- Faixa de opções reorganizada em menus por família de gráfico; os controles do grupo Gráfico selecionado só ficam ativos nos tipos que os aceitam, e o que não se aplica avisa em vez de falhar. Mais opções fica ativo mesmo sem gráfico selecionado, por causa de Mesma escala e Escala automática.
- Cascata: uma tabela com várias colunas de valores passa a gerar várias séries (antes só a primeira coluna era usada), e `e` na primeira linha com um número na última passa a calcular o primeiro total para trás (antes o primeiro total ficava em zero).
- Totais, antes só nas colunas empilhadas e 100%, também valem em barras, áreas, combinação e Mekko; Remover setas também apaga os colchetes do Gantt.
- Desenvolvimento: o log dos builds de teste (core e loader) vai para `%TEMP%\ThinQQ-test-logs`, e não para a pasta do usuário, e os builds de teste do loader não criam mais `%LOCALAPPDATA%\ThinQQ` (só as pastas de `%TEMP%`).
- Desenvolvimento: `tests/ps/Test-Privacy.ps1` (rodado por `build/run-tests.ps1`) falha se um arquivo versionado trouxer caminho de perfil do Windows, id de máquina ou e-mail corporativo.

### Corrigido
- Copiar um gráfico ou movê-lo para uma planilha de gráfico não apaga mais os dados: a limpeza de blocos decide pelas referências das séries e dos rótulos ligados (inclusive em gráficos 100% sem totais e em gráficos dentro de grupos de formas) e, se algum gráfico não puder ser lido, não apaga nada.
- Categorias de data usam eixo de texto (uma coluna por data), alinhado à geometria das setas.
- Alterar dados recusa uma tabela de outra pasta de trabalho.
- Aviso único quando linhas são inseridas ou removidas na tabela de origem (use Alterar dados).
- Gráficos 100%: rótulos dos segmentos mostram a participação; o formato escolhido vale para os totais.
- Rótulo do total em colunas empilhadas de sinais misturados no topo da parte positiva.
- Escala automática só desfaz a escala fixa de Mesma escala: um eixo ou uma área de plotagem ajustados à mão ficam como estão, e um mínimo fixado à mão no eixo secundário da combinação não é mais trocado por zero.
- Faixa de opções: os itens da lista Espaço usavam ids que começavam por dígito, o que o Excel não aceita no `customUI14` (ele descarta a aba inteira, sem aviso); agora são `gap30` a `gap120`. Só afetava compilações internas anteriores a esta versão.

### Segurança
- SHA-256 em VBA puro: o ThinQQ não inicia mais o `certutil`.
- Ganchos de teste inertes em Release; downloads em `%TEMP%\ThinQQ`, fora do local confiável.

## [1.0.0] - 2026-10-10 (não publicada)

### Adicionado
- Gráficos: colunas empilhadas, agrupadas, 100% e linha; cascata com marcadores `e`, `t`, `=`, `total` e
  `subtotal` para barras de total calculadas, incluindo barras que cruzam o zero.
- Totais, rótulos ligados às células, conectores (linhas de série e conectores da cascata).
- Setas de CAGR e de diferença, reposicionadas automaticamente ao mudar os dados ou o tamanho do gráfico.
  A seta de diferença não está disponível em gráficos 100%.
- Paletas (Padrão ThinQQ, Tons de cinza, Cores do tema), formatos numéricos e espaço entre colunas.
- Aba ThinQQ na faixa de opções, com controles que seguem o gráfico selecionado; Atualizar e Alterar dados.
- Pasta de exemplos com seis gráficos prontos e tutorial em HTML.
- Instalador e desinstalador por usuário, sem administrador, com local confiável e remoção do bloqueio dos arquivos.
- Atualização por manifesto (HTTP ou pasta de rede/OneDrive) com conferência SHA-256, troca a quente do
  core sem reiniciar o Excel e substituição segura do loader com backup.

### Observações
- Células vazias, com texto ou com erro contam como zero.
- Rótulos de categoria ficam na base do gráfico, mesmo com valores negativos.
