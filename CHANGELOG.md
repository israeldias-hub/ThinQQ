# Changelog

Formato baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/); versões no formato x.y.z.

## [Não lançado]

### Corrigido
- Copiar um gráfico ou movê-lo para uma planilha de gráfico não apaga mais os dados: a limpeza de blocos decide pela referência das séries.
- Categorias de data usam eixo de texto (uma coluna por data), alinhado à geometria das setas.
- Alterar dados recusa uma tabela de outra pasta de trabalho.
- Aviso único quando linhas são inseridas ou removidas na tabela de origem (use Alterar dados).
- Gráficos 100%: rótulos dos segmentos mostram a participação; o formato escolhido vale para os totais.
- Rótulo do total em colunas empilhadas de sinais misturados no topo da parte positiva.

### Segurança
- SHA-256 em VBA puro: o ThinQQ não inicia mais o `certutil`.
- Ganchos de teste inertes em Release; downloads em `%TEMP%\ThinQQ`, fora do local confiável.

## [1.0.0] - 2026-10-10

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
