# Changelog

Formato baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/); versões no formato x.y.z.

## [Não lançado]

### Adicionado
- Atualização automática dos gráficos do ThinQQ no PowerPoint (por gráfico, como o "Update Automatically" do think-cell): no PowerPoint, aba ThinQQ > Gráfico selecionado > Atualizar automaticamente (ligar já atualiza uma vez; clicar de novo volta ao manual) e Atualizar uma vez, só nos gráficos selecionados. Com o Excel e o PowerPoint abertos, mudar um valor da pasta (digitar, recalcular ou salvar) atualiza dados e setas dos gráficos automáticos cerca de 2 segundos depois da última mudança; uma pasta que nunca para de recalcular é atualizada no máximo uns 10 a 11 s depois da primeira mudança. Só apresentações já abertas e editáveis, só pastas salvas e abertas com gráficos do ThinQQ; o ciclo nunca abre pasta, PowerPoint nem apresentação, e sem gráfico automático não usa a janela do PowerPoint. Espera uma colagem ou atualização manual; com texto em edição no PowerPoint tenta de novo depois de 4, 6, 8, 10 e 12 s e depois desiste daquele ciclo (log) até a próxima mudança. Nada vira automático sem uma ação do usuário; a marca fica na forma (`TQQ_AUTO`) e a apresentação guarda o índice (`TQQ_AUTOIDX`).
- Excel: ThinQQ > PowerPoint ganha as caixas Atualização automática ligada (padrão: ligada; desmarcada pausa o automático) e Novos vínculos atualizam automaticamente (padrão: desligada; os gráficos colados ou inseridos a partir daí já nascem automáticos).
- Inserir no slide (1.7.0): Inserir também no slide (ThinQQ > PowerPoint, o "Link To PowerPoint" do think-cell): com a caixa marcada, escolher um tipo nos menus Colunas, Barras, Cascata, Linha e área e Mais tipos insere o gráfico no Excel e o cola vinculado no slide da janela ativa do PowerPoint (ou numa apresentação nova, com o PowerPoint fechado), com as setas no tamanho do slide e, se Novos vínculos atualizam automaticamente estiver marcada, já automático. A caixa é um estado da sessão (volta desmarcada ao abrir o Excel) e a pasta precisa estar salva: pasta não salva avisa e não cria nada.
- Inserir no slide (1.7.0): moldura verde no intervalo de origem do gráfico inserido (retângulo de 2 pt que acompanha as linhas e colunas inseridas, não impresso e bloqueado) e ThinQQ > PowerPoint > Mostrar/ocultar molduras (todas as molduras da pasta ativa). A moldura é só um indicador: não muda o vínculo (mudar o intervalo é Alterar dados, que a redesenha), só os gráficos inseridos por este caminho a ganham, não há molduras vermelha e azul e ela não é impressa; colagem que falhou ou colada sem vínculo deixa o gráfico no Excel, sem moldura.
- Inserir no slide (1.7.0): com um núcleo anterior à 1.7.0 no Excel e a caixa marcada, escolher um tipo só mostra "Atualize o ThinQQ para usar este recurso." e não insere nada.- Contrato com o suplemento do PowerPoint na API 2 do core (Core_PptSyncShapes, Core_PptGetFlag, Core_PptSetFlag), conferida por recurso: com o ThinQQ 1.5.x no Excel, ligar o automático só grava a marca ("Atualize o ThinQQ no Excel para a atualização automática funcionar (a marca foi gravada).") e Atualizar uma vez atualiza a apresentação inteira, com um aviso.

### Alterado
- Os nomes dos gráficos do ThinQQ no slide ficam únicos sem diferenciar maiúsculas de minúsculas ao ligar o automático e em Atualizar uma vez (o Shapes(nome) do PowerPoint ignora a caixa).
- Documentação: tutorial (subseção Atualização automática, mensagens e limites), README (recursos, arquitetura, verificações manuais 61 a 68 e limitações) e Test-Docs (rótulos das faixas do Excel e do PowerPoint, âncora #ppt-auto, mensagens e limites por lugar). Não verificado em PowerPoint real: o tempo de verdade entre digitar e o slide mudar, a caixa "aguardando ação OLE" com texto em edição, o Ctrl+Z, o descarregamento com um ciclo pendente e a apresentação somente leitura.
- Documentação (1.7.0): tutorial (subseção Inserir também no slide, referência da faixa e limites), README (recursos, arquitetura, verificações manuais 69 a 74 e limitações) e Test-Docs (rótulos Inserir também no slide e Mostrar/ocultar molduras, âncora #ppt-inserir, mensagens e itens 69 a 74). Não verificado em Excel e PowerPoint reais: clicar dentro da moldura (sonda P4), a impressão sem a moldura, linhas inseridas ou apagadas, a planilha copiada e o fluxo com o PowerPoint fechado.- O Desfazer do Excel pode ser limpo cerca de 2 segundos depois de digitar, quando um gráfico realmente atualiza.

## [1.5.0] - 2026-10-15
- Delta (1.5.0): Delta entre barras: Setas > Delta clicando nas barras insere uma seta com a diferença (valor e %) entre quaisquer duas barras, escolhidas com dois cliques; Delta: valor / % / ambos alterna o texto.
- - Delta (1.5.0): faixa do Excel: o menu Setas ganha Delta clicando nas barras e Delta: valor / % / ambos (ativos nos tipos que aceitam o delta; coluna Delta na tabela do tutorial).
- Delta (1.5.0): exemplos: o exemplo 3 de ThinQQ > Exemplos (colunas agrupadas) e o E02 da pasta de demonstração (com a imagem na galeria do tutorial) ganham um delta; as contagens de exemplos não mudam.
- Delta (1.5.0): documentação: tutorial (seção Delta entre barras, mensagens e limitações), README (recursos, verificações manuais 55 a 60, limitações) e a especificação (regra do lado de baixo na seção 5.3); o Test-Docs confere as mensagens e os limites novos.


## [1.4.0] - 2026-10-15

### Adicionado
- Suplemento do PowerPoint (ThinQQ.ppam): ao abrir uma apresentação, atualiza dados e setas dos gráficos do ThinQQ colados como gráfico vinculado (perguntar, sempre ou nunca). Aba ThinQQ no PowerPoint, grupo Gráficos do Excel: Atualizar gráficos, Ao abrir (Perguntar, Atualizar sempre, Não atualizar) e Sobre. Usa o Excel aberto com o ThinQQ ou, senão, um Excel invisível só para isso; sem o ThinQQ no Excel atualiza só os dados. As pastas de origem são abertas somente leitura e com as macros desligadas; a pergunta lista as pastas, e "Atualizar sempre" pergunta assim mesmo com uma pasta de rede ou da web ainda não aberta no Excel. Apresentações somente leitura não são atualizadas ao abrir (só pelos botões, na memória); Modo de Exibição Protegido: a pergunta aparece depois de Habilitar Edição (não verificado em PowerPoint real).
- Excel: PowerPoint > Atualizar gráficos no PowerPoint (atualiza dados e setas dos gráficos vinculados das apresentações abertas; funciona sem o suplemento do PowerPoint).
- Atualização: o Excel também atualiza o suplemento do PowerPoint (com o PowerPoint aberto, quando ele fechar), só para quem o instalou pelo ZIP; opção "Atualizar também o suplemento do PowerPoint" na janela de atualização e em ThinQQ > Configurações.
- Instalação: o instalador do ZIP instala o `ThinQQ.ppam` (registro `PowerPoint\AddIns\ThinQQ` e Local Confiável `LocationThinQQ` do PowerPoint; o desinstalador os remove); um ZIP anterior à 1.4.0, sem o `.ppam`, continua instalando só o Excel. O release passa a publicar 6 assets (inclui `ThinQQ.ppam`) e a assinar o `.ppam` com os `.xlam`.

### Alterado
- Faixa do Excel: o menu PowerPoint fica sempre ativo (só os dois itens de copiar exigem um gráfico selecionado).
- Documentação: tutorial (seção Copiar para o PowerPoint reescrita, atualizações, instalação, problemas e limitações), README (instalação, verificações manuais 43 a 54, limitações) e `installer/LEIA-ME.txt` (instalação manual do suplemento e requisito do PowerPoint 16.0); `Test-Docs` confere os rótulos da faixa do PowerPoint, as mensagens e os arquivos da release.


### Corrigido
- Copiar como gráfico vinculado: as setas e rótulos de seta não ficam mais esticados e fora do lugar quando o PowerPoint ajusta o gráfico ao slide.

## [1.3.0] - 2026-10-14

### Adicionado
- Cascata comparativa (menu Cascata): duas pontes lado a lado, uma por série, com conectores tracejados por série, uma seta de variação em cotovelo entre as duas barras de cada etapa (`(B - A) / A` numa elipse; `n/d` com o primeiro valor zero, sinais opostos ou total contra degrau na mesma linha; ou a diferença em valor pelo alternador Setas > Variação em valor), totais com o percentual do primeiro total da série ("11 (41%)") e os nomes das séries dentro das barras da primeira etapa (Legenda > Nome dentro das barras; quando não cabem, a legenda). Exatamente 2 colunas de valores; marcadores `e`, `t`, `=`, `total`, `subtotal`. Tudo em séries nativas ligadas ao bloco da `_ThinQQ`: barras, conectores, setas e textos acompanham os números mesmo sem o ThinQQ.
- Exemplo 27 em ThinQQ > Exemplos (planilha Exemplos 3) e exemplos E27 e E28 na pasta de demonstração, com as imagens na galeria do tutorial.

### Alterado
- Faixa de opções: Cascata ganha "Cascata comparativa" (ícone próprio); Setas ganha "Setas de variação" e "Variação em valor"; Legenda ganha "Nome dentro das barras"; o alternador Totais, na cascata comparativa, liga o percentual do primeiro total. A matriz de capacidades passa a 21 tipos x 23 recursos.

### Corrigido
- Atualização pelo Excel: o download dos arquivos da release do GitHub parava com "Falha ao atualizar: O servidor de atualização respondeu com erro. HTTP 0 ao acessar https://github.com/.../ThinQQ.core.xlam". O GitHub redireciona (302) o download para outro domínio (`release-assets.githubusercontent.com`) e o `MSXML2.XMLHTTP` do carregador bloqueia esse redirecionamento. O carregador continua tentando primeiro o `MSXML2.XMLHTTP` (respeita o proxy e o script de configuração automática do Windows) e, numa falha de transporte (exceção, HTTP 0 ou erro de rede 12xxx), repete a mesma requisição com `MSXML2.ServerXMLHTTP`, que segue o redirecionamento, usando o proxy fixo do usuário (`ProxyServer`/`ProxyOverride` de Internet Settings); script de configuração automática (PAC) não é usado nessa segunda tentativa. Respostas HTTP de erro (404, 500...) não disparam a segunda tentativa; se as duas falharem, a mensagem traz as duas causas. O log do carregador registra qual caminho baixou. Quem tem a 1.2.0, 1.2.1 ou 1.2.2 precisa reinstalar uma vez pelo ZIP da 1.2.3 (o carregador antigo não consegue baixar a atualização).
- Testes: o servidor HTTP de teste (`tests/http/server.js`) ganhou as rotas `/redir-host/` (302 entre `127.0.0.1` e `localhost`, como o GitHub) e `/redir-same/`; testes novos do transporte com a reserva, o 404 sem reserva, as duas falhas e a escolha do proxy a partir de valores.

## [1.2.2] - 2026-10-13

### Adicionado
- Exemplo de cascata com seta de diferença nos Exemplos (exemplo 2) e na pasta de demonstração (E05), com o passo a passo no tutorial.
- Galeria com a imagem dos 26 exemplos da demonstração no tutorial (imagens embutidas no próprio `tutorial.html`, que continua sendo um arquivo único), com legenda, caminho de menu e links cruzados das seções; `tools/Export-DemoChartImages.ps1` e `tools/Build-TutorialGallery.ps1` refazem a galeria. O limite de tamanho do tutorial no teste subiu de 192 KB para 640 KB.
- Publicação automatizada de uma versão nova, em duas etapas locais: `publicar-release.bat x.y.z` (prepara a branch `release/vx.y.z` com `version.json`, CHANGELOG, tutorial, LEIA-ME e `release/manifest.json`, constrói numa cópia em `C:\temp`, guarda os arquivos e um `state.json` em `dist\release-x.y.z\`, envia a branch e abre o pull request) e, depois de mesclar o pull request, `publicar-release.bat x.y.z publicar` (cria a release em rascunho pela API do GitHub, envia os 5 assets, confere tamanho e SHA-256 de cada um e publica). Opções `-DryRun`, `-Draft`, `-NoBrowser`, `-RunTests` e `-AllowMainChanges`; recomenda o merge "Create a merge commit" (um squash ou rebase com o mesmo conteúdo também é aceito); o token vem do login do Git (`git credential fill`), fica só na memória e nunca é impresso nem gravado. Script `tools/Publish-Release.ps1` (Windows PowerShell 5.1 e pwsh 7) e teste `tests/ps/Test-PublishRelease.ps1` (sem rede e sem Excel: backend HTTP falso e repositório git descartável).

### Alterado
- `publicar-release.bat` aberto com dois cliques (sem argumentos) agora pergunta a versão e a etapa em vez de só mostrar o uso.
- `Test-Docs` deixa de fixar a versão 1.2.0 no `version.json` (aceita as versões seguintes) e confere a documentação do `publicar-release.bat` no README e em `release/README.md`; `build/run-tests.ps1` roda também o `Test-PublishRelease` (pwsh e Windows PowerShell 5.1).

## [1.2.0] - 2026-10-12

### Adicionado
- Borboleta (menu Mais tipos): duas séries espelhadas, com os nomes das categorias no centro, o total no fim da barra direita, os dois lados na mesma escala (eixo simétrico) e um vão central calculado pelo maior nome. Exige exatamente 2 colunas de valores; os rótulos de barras muito finas ficam fora da barra.
- Dispersão e Bolhas (menu Mais tipos): tabela Nome, X, Y (e Tamanho nas bolhas) com coluna de Grupo opcional (uma série por grupo, até 6), nome ao lado de cada ponto, títulos dos eixos ligados aos cabeçalhos, eixos no formato da coluna de origem e apertados em volta dos dados, Linha de tendência (regressão linear por série) e Quadrantes (médias).
- Rótulos com soma consistente (Mais opções): os rótulos dos segmentos de colunas e barras empilhadas e 100% e de áreas somam exatamente o total mostrado (método do maior resto, em fórmulas da planilha que continuam certas sem o ThinQQ).
- Arredondar intervalo... (Mais opções): grava, a partir da célula escolhida, fórmulas de arredondamento com soma consistente para uma coluna ou linha de até 1.000 células; elas só referenciam a faixa de origem e não precisam de gráfico nem do ThinQQ.
- Corte de eixo... e Remover corte de eixo (Mais opções), em colunas e barras empilhadas e agrupadas: as colunas maiores são cortadas num valor que você informa, com marca de corte e o total real no rótulo; nas empilhadas, ao definir o corte, o ThinQQ recusa um valor que esconderia o segmento de cima.
- Menu PowerPoint (Copiar como imagem e Copiar como gráfico vinculado): cola o gráfico no slide da janela ativa do PowerPoint (abre o PowerPoint ou cria uma apresentação, se preciso) sem nunca fechar o que você já tinha aberto.
- Ícones próprios na faixa de opções (PNG gerados no build; `build.ps1 -RibbonIcons Mso` volta aos ícones do Office).
- Certificado e assinatura digital: `tools/New-ThinQQCert.ps1` (certificado autoassinado de teste, por usuário), `tools/Trust-ThinQQPublisher.ps1` (Editores Confiáveis) e `tools/Sign-ThinQQ.ps1` (signtool + Office SIPs; sai com código 2 e instruções sem os pré-requisitos), além de `release.ps1 -SignPfx`.
- Pasta de trabalho de demonstração `exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx` (rede fictícia de materiais de construção, 26 exemplos E01 a E26 com índice de links, sem macros), distribuída dentro do ZIP da release, na pasta `exemplos`, com um LEIA-ME curto.
- Exemplos 3: terceira planilha de ThinQQ > Exemplos, com os 8 exemplos novos (26 exemplos no total).
- Documentação: seção "Instalação" completa no README (requisitos, passo a passo, atualizações, desinstalação, instalação manual, problemas comuns e o que fazer com o repositório privado) e subseção "Ambientes corporativos (TI)"; no tutorial, seções Borboleta, Dispersão e bolhas, Arredondamento consistente, Corte de eixo, Copiar para o PowerPoint, Pasta de demonstração e Instalação em empresas (TI).
- Com o loader 1.2.0 e um core anterior, os botões Arredondar intervalo... e os do menu PowerPoint (com um gráfico selecionado) avisam "Atualize o ThinQQ para usar este recurso." em vez de não fazer nada.

### Alterado
- Faixa de opções: passa a oferecer 20 tipos (Mais tipos ganha Borboleta, Dispersão e Bolhas); Mais opções ganha Rótulos com soma consistente, Linha de tendência, Quadrantes (médias), Corte de eixo..., Remover corte de eixo e Arredondar intervalo...; novo menu PowerPoint no grupo Gráfico selecionado. A matriz de capacidades (20 tipos x 20 recursos) é conferida por teste contra a especificação.
- Mesma escala: a borboleta entra (use só entre borboletas); rosca, dispersão, bolhas e gráficos com corte de eixo ficam fora do grupo, e a mensagem conta os de corte de eixo à parte.
- Combinação com Totais e eixo secundário: a linha passa a ocupar uma faixa acima dos totais das colunas (o máximo do eixo de valores das colunas, que fica oculto, sobe; com Mesma escala volta o comportamento anterior).
- Formato inicial dos rótulos de um gráfico novo: vem da primeira célula de valor da tabela (uma, duas ou mais casas decimais, percentual, milhares ou milhões), em vez de sempre inteiro.
- Pizza e rosca: paletas de 10 cores (as 6 primeiras como antes); até 10 fatias nenhuma cor se repete.
- O bloco da planilha oculta `_ThinQQ` localiza as colunas Pos e VLine (e as novas) pelo cabeçalho, para aceitar colunas acrescentadas à direita (soma consistente e corte de eixo); blocos antigos continuam funcionando.
- `build/release.ps1`: o ZIP passa a levar a pasta `exemplos` (o Excel de demonstração e um LEIA-ME), com nomes de entrada com `/`; o release é recusado se o ZIP tiver outro `.xlsx` ou `.xlsm`, se o xlsx diferir (SHA-256) do repositório ou se tiver macros; `-SignPfx` assina os `.xlam` antes dos SHA-256 do manifesto.
- `build/run-tests.ps1` passa a rodar também Test-Signing e Test-DemoWorkbook; `Test-Docs` confere as seções novas, a lista dos 26 exemplos da demonstração, as mensagens e os limites documentados.

### Corrigido
- Dispersão e bolhas: o nome do ponto passa para a esquerda do ponto quando passaria da borda direita da área do gráfico (não invade mais a legenda) e volta para a direita quando cabe.
- Corte de eixo: em colunas e barras empilhadas, ao definir o corte, o ThinQQ recusa um valor que esconderia por inteiro o segmento de cima de uma coluna cortada e diz o mínimo (se os dados subirem depois, um segmento pode ficar acima do corte, sem rótulo); posições de rótulo que você ajustou são mantidas; o aviso aparece quando a linha de valor (média) fica acima do corte, quando o gráfico sai de um grupo de Mesma escala e quando Alterar dados retira o corte ou o deixa escondendo um segmento; o rótulo de base que ficaria sob a marca é ocultado; um mínimo de eixo fixado à mão acima do corte volta ao automático.
- Dispersão e bolhas: os números dos eixos seguem o formato da coluna de origem (um percentual aparecia como 0,3) e os eixos não começam mais em zero quando os dados estão longe dele; a escala ou o formato que você fixar são respeitados.
- Pizza e rosca com mais de 6 fatias repetiam cores (a 7.ª repetia a 1.ª).
- Combinação com totais: a linha do eixo secundário cruzava os rótulos de total das colunas.
- Rótulos iniciais inteiros em tabelas com casas decimais, em percentual ou em milhares.
- Área: os rótulos da primeira e da última categoria colidiam com os vizinhos (o que encosta agora fica oculto e volta quando cabe).
- Borboleta: o rótulo de uma barra fina desaparecia (agora aparece fora da barra, e o total vem depois dele).
- Arredondar intervalo: teto de 1.000 células (acima disso o recálculo ficava lento); recusa, antes de gravar, uma planilha protegida sem a permissão Formatar células; uma origem com "%" só no formato (como em `0,0"%"`) dá número simples, e não 3400%; ao desfazer por uma falha, devolve as constantes digitadas (como "00123") do jeito que estavam.
- Copiar para o PowerPoint: o destino é a apresentação da janela ativa (nunca a de outra janela); só fecha o que a própria chamada criou e remove a forma colada se algo falhar depois; avisa quando o PowerPoint está ocupado, quando a área de transferência está indisponível e quando o PowerPoint cola o gráfico sem manter o vínculo.
- Testes: os de PowerPoint nunca miram nem fecham o PowerPoint do usuário; os ganchos de teste do Arredondar intervalo, do corte de eixo e das bolhas são zerados em toda saída.

### Segurança
- Guarda do release: o ZIP só pode conter os arquivos esperados, e o Excel de demonstração vai sem macros e idêntico ao do repositório.
- `tools/Sign-ThinQQ.ps1` e `release.ps1 -SignPfx` permitem à TI distribuir os dois `.xlam` assinados (veja "Ambientes corporativos (TI)" no README). O ramo `signtool` + Office SIPs não foi executado ponta a ponta na máquina de desenvolvimento: confira o resultado no Excel.

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
- Espaço vale para todos os grupos de colunas ou barras do gráfico: numa combinação com uma série de colunas levada ao eixo secundário, os dois grupos de colunas recebem o espaço escolhido (antes só o primeiro).
- Escala automática só desfaz a escala fixa de Mesma escala: um eixo ou uma área de plotagem ajustados à mão ficam como estão, e um mínimo fixado à mão no eixo secundário da combinação não é mais trocado por zero.
- Pizza: trocar a Paleta não devolve mais à posição padrão um rótulo arrastado nem recria um rótulo apagado (só a cor da fonte acompanha as fatias).
- Gantt: muitos colchetes não esmagam mais a área do gráfico: o espaço deles vai até 40% da altura do gráfico e os colchetes ficam empilhados dentro desse limite.
- Cascata (barras): um rótulo à esquerda perto do mínimo do eixo não invade mais os nomes das categorias (passa para dentro da barra, com fundo branco), e a seta CAGR passa à direita dos rótulos das pontas em vez de cruzá-los.
- Mekko: um nome de categoria em branco aparece em branco (antes "0"); uma coluna estreita demais para a resolução do eixo não mostra mais nome nem rótulos soltos sobre as vizinhas; no Mekko (%) os valores negativos entram como zero (com aviso); inserir e religar os rótulos de um Mekko grande ficou cerca de 3 vezes mais rápido.
- Cascata com várias séries: o rótulo de um segmento mais fino que o texto fica oculto (antes encostava nas bordas) e volta quando o segmento passa a caber.
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
