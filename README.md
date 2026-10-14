# ThinQQ

Suplemento do Excel (VBA, Windows) para criar gráficos de apresentação no estilo think-cell: colunas e barras
(empilhadas, agrupadas e 100%), cascata (ponte, com várias séries), linha, área, combinação, Mekko, Gantt, pizza,
rosca, borboleta, dispersão e bolhas, com totais, rótulos vivos, conectores, setas de CAGR e de diferença, linha de
valor, indicador 100%, mesma escala, soma consistente, corte de eixo e cópia para o PowerPoint. Os gráficos gerados são
gráficos nativos do Excel, ligados aos dados por fórmulas: quem recebe o arquivo não precisa ter o ThinQQ instalado.

Repositório: <https://github.com/israeldias-hub/ThinQQ>

**Para instalar, vá direto à seção [Instalação](#instalação)** (baixar o ZIP em Releases, extrair a pasta inteira e dar dois cliques em `Instalar-ThinQQ.cmd`). Para conhecer os recursos sem montar nada, abra o Excel de demonstração
[`exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx`](exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx), que também vai dentro do ZIP.

## Instalação

Para quem vai **usar** o ThinQQ: não precisa saber programar nem ser administrador, e leva uns 5 minutos.

### Requisitos

- Windows com o **Excel do Office 16.0 para desktop**: Excel 2016, 2019, 2021 ou Microsoft 365 (o aplicativo instalado). Não funciona no Mac nem no Excel na web.
- Não precisa de administrador: tudo é instalado só para o seu usuário do Windows.
- Os testes automatizados rodaram só no Excel 2019 de 32 bits; o 2016, o 2021 e o Microsoft 365 são esperados, não testados. O código não faz chamadas à API do Windows (nenhuma instrução `Declare`), então deve servir em 32 e 64 bits por construção.
- Acesso à página de [Releases](https://github.com/israeldias-hub/ThinQQ/releases) do repositório, que hoje é público (veja "Visibilidade do repositório", abaixo).

### Passo a passo

1. **Baixe o ZIP.** Na página de [Releases](https://github.com/israeldias-hub/ThinQQ/releases), abra a versão mais nova e, em "Assets", baixe `ThinQQ-x.y.z.zip` (por exemplo `ThinQQ-1.2.0.zip`).
2. **Descompacte a pasta inteira:** botão direito no ZIP > **Extrair tudo...**. Não execute o instalador de dentro do ZIP e não separe os arquivos: o instalador precisa de todos eles na mesma pasta. Se o ZIP veio da internet, antes de extrair clique nele com o botão direito > Propriedades e marque **Desbloquear** (se a opção aparecer) > OK: isso evita avisos de segurança.
3. **Feche o Excel** (todas as janelas).
4. **Dê dois cliques em `Instalar-ThinQQ.cmd`.** Se o Windows perguntar se quer executar, clique em **Executar**; se aparecer "O Windows protegeu o computador", clique em **Mais informações** > **Executar assim mesmo**. Uma janela mostra o andamento e termina com "INSTALACAO CONCLUIDA".
5. **Abra o Excel e procure a aba ThinQQ**, ao lado de Inserir. Para ver o ThinQQ funcionando, abra `exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx` (vem no ZIP, na pasta `exemplos`; no GitHub: [exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx](exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx)): a primeira aba é um índice; clique no nome do exemplo para ir ao gráfico. Os gráficos também funcionam sem o ThinQQ instalado, mas setas, colchetes e marcas de corte só se recalculam com ele. ThinQQ > Exemplos e ThinQQ > Tutorial também estão na aba.

### Visibilidade do repositório: quem consegue baixar

Hoje o repositório é **público**: qualquer pessoa abre a página de Releases e baixa o ZIP, e o atualizador do ThinQQ alcança o manifesto (`release/manifest.json` no GitHub) sem login. O tutorial vai dentro do ZIP e ThinQQ > Tutorial abre o instalado; o tutorial online (GitHub Pages) só existe depois de habilitado (veja "Tutorial online (GitHub Pages)").

Se o repositório voltar a ser **privado**, quem não tem acesso a ele **deixa de abrir a página de Releases e de baixar o ZIP**, e o atualizador não alcança o manifesto: a consulta falha em silêncio, sem aviso de versão nova (e, no plano gratuito, o GitHub Pages exige repositório público). Opções para quem mantém o repositório nesse caso:

1. **Convidar as pessoas como colaboradores** (Settings > Collaborators; acesso de leitura basta): cada uma precisa de uma conta no GitHub.
2. **Manter o repositório público** (a situação atual; a limpeza de privacidade descrita em "Antes de distribuir" vale antes de abrir qualquer repositório). As releases têm a mesma visibilidade do repositório; para ter releases públicas sem abrir o código, publique o ZIP e o manifesto num repositório público à parte e use o endereço dele em `-BaseUrl` e `-ManifestUrl`.
3. **Publicar o ZIP e o manifesto numa pasta compartilhada** (rede, OneDrive ou SharePoint sincronizado), sem passar pelo GitHub: `pwsh build/release.ps1 -Version x.y.z -BaseUrl '\\servidor\ThinQQ' -PublishTo '\\servidor\ThinQQ'`. Quem instala por esse ZIP passa a consultar `<pasta>\manifest.json` (o instalador grava esse caminho como fonte de atualização); quem já tem o ThinQQ pode apontar ThinQQ > Configurações para o `manifest.json` da pasta.

Quem recebeu o ZIP por outro caminho (e-mail, Teams, pasta de rede) instala normalmente pelo passo a passo acima; só as atualizações dependem da fonte configurada em ThinQQ > Configurações.

### Como as atualizações chegam

- Ao abrir o Excel, o ThinQQ consulta a fonte de atualização no máximo **uma vez por dia**. Se houver versão nova, abre uma janela com as novidades e os botões **Atualizar agora** e **Depois** (a verificação automática traz também **Pular esta versão**).
- ThinQQ > Verificar atualizações consulta na hora e diz também quando você já está na versão mais recente.
- O download tenta primeiro o mesmo caminho do Windows (proxy e script de configuração automática); se ele falhar na conexão, como no redirecionamento do GitHub para outro domínio, tenta de novo por um segundo caminho que usa só o proxy fixo do Windows (sem script PAC). O log (`%LOCALAPPDATA%\ThinQQ\logs\thinqq.log`) registra qual caminho baixou.
- Os arquivos baixados são conferidos pelo SHA-256 do manifesto antes de serem instalados. O núcleo é trocado na hora, sem reiniciar o Excel; se a versão traz um loader novo, ele é trocado logo em seguida, também sem reiniciar.
- A fonte (um endereço `https://...` ou uma pasta ou arquivo de rede) pode ser trocada em ThinQQ > Configurações. Se o repositório voltar a ser privado, quem não tem acesso deixa de receber as atualizações (veja "Visibilidade do repositório", acima).

### Desinstalar

Feche o Excel e abra o PowerShell na pasta onde você descompactou o ZIP (**guarde essa pasta**: o desinstalador só existe no ZIP): no Explorador de Arquivos, clique na barra de endereços, digite `powershell` e tecle Enter. Na janela que abrir, execute:

```powershell
powershell -ExecutionPolicy Bypass -File uninstall.ps1
```

Isso remove o registro do suplemento, o Local Confiável, **toda** a chave `HKCU\Software\VB and VBA Program Settings\ThinQQ` (as configurações do atualizador, inclusive a fonte de atualização e a versão pulada) e a pasta `%LOCALAPPDATA%\ThinQQ` (o parâmetro `-KeepFiles` mantém os arquivos). Os gráficos que você já criou continuam nas suas pastas de trabalho, como gráficos comuns do Excel.

### Instalação manual

Use se a política da empresa bloquear o script. Os arquivos estão no ZIP (`ThinQQ.xlam`, `ThinQQ.core.xlam`, `tutorial.html` e `manifest.json`, que traz a versão):

1. Crie a pasta `%LOCALAPPDATA%\ThinQQ` (cole esse caminho na barra de endereços do Explorador de Arquivos).
2. Copie `ThinQQ.xlam` para essa pasta, e `ThinQQ.core.xlam` e `tutorial.html` para `%LOCALAPPDATA%\ThinQQ\versions\<versão>\` (a versão está no `manifest.json`, por exemplo `versions\1.2.0\`).
3. Em cada `.xlam`: botão direito > Propriedades > marque **Desbloquear** > OK.
4. Excel > Arquivo > Opções > Central de Confiabilidade > Configurações da Central de Confiabilidade > Locais Confiáveis > Adicionar novo local: `%LOCALAPPDATA%\ThinQQ`, marcando "Subpastas também são confiáveis".
5. Excel > Arquivo > Opções > Suplementos > Gerenciar: Suplementos do Excel > Ir... > Procurar, e escolha `%LOCALAPPDATA%\ThinQQ\ThinQQ.xlam` > OK.

A mesma instalação manual, com a recuperação de uma atualização do loader que falhou, está em `installer/LEIA-ME.txt` (incluído no ZIP).

### Problemas mais comuns

1. **Macros bloqueadas, ou arquivo bloqueado pelo Windows.** *O que você vê:* uma faixa de aviso de segurança sobre macros, ou a aba ThinQQ não aparece depois de instalar. *O que fazer:* feche o Excel e rode `Instalar-ThinQQ.cmd` de novo (ele desbloqueia os arquivos e cria o local confiável). Se instalou à mão: botão direito em `ThinQQ.xlam` e em `ThinQQ.core.xlam` > Propriedades > marque **Desbloquear** > OK. Um arquivo baixado da internet fica "bloqueado" até isso.
2. **Central de Confiabilidade.** *O que você vê:* "As macros foram desabilitadas", ou o ThinQQ não carrega. *O que fazer:* Arquivo > Opções > Central de Confiabilidade > Configurações da Central de Confiabilidade > Configurações de Macro. Com "Desabilitar todas as macros, exceto as assinadas digitalmente", o ThinQQ só funciona a partir de um Local Confiável (item seguinte) ou com os `.xlam` assinados pela TI (veja "Ambientes corporativos (TI)"). Só mude essa configuração se a política da empresa permitir.
3. **Local Confiável.** *O que você vê:* aviso de macros a cada abertura, ou o instalador não conseguiu criar o local. *O que fazer:* Arquivo > Opções > Central de Confiabilidade > Configurações da Central de Confiabilidade > Locais Confiáveis > Adicionar novo local: `%LOCALAPPDATA%\ThinQQ`, marcando "Subpastas também são confiáveis" (o core fica em `versions\<versão>`). O instalador já cria esse local; refaça-o se uma política da empresa o removeu.
4. **Antivírus ou proxy.** *O que você vê:* o download do ZIP bloqueado ou incompleto, a mensagem "Pacote corrompido ou adulterado" (o SHA-256 não confere), o `ThinQQ.xlam` que sumiu da pasta (quarentena do antivírus) ou "Não foi possível verificar atualizações." / "Falha ao atualizar: ..." atrás de um proxy. *O que fazer:* baixe o ZIP de novo pelo navegador, restaure o arquivo da quarentena do antivírus e peça à TI para liberar `%LOCALAPPDATA%\ThinQQ`; atrás de um proxy, reinstale pelo ZIP em vez de usar o atualizador.
5. **"Falha ao atualizar: O servidor de atualização respondeu com erro. HTTP 0 ao acessar https://github.com/..."** *O que você vê:* essa mensagem em ThinQQ > Verificar atualizações > Atualizar agora, com o ThinQQ 1.2.0, 1.2.1 ou 1.2.2. *Por quê:* o GitHub entrega os arquivos da release por um redirecionamento para outro domínio, que o carregador dessas versões não consegue seguir. *O que fazer:* reinstale **uma vez** pelo ZIP da release mais nova (passo a passo acima); da 1.2.3 em diante a atualização pelo Excel funciona. Atrás de um proxy configurado por **script de configuração automática** (PAC), a segunda tentativa de download do carregador (a que segue o redirecionamento) não usa o script, só o proxy fixo do Windows: se a atualização falhar nessa rede, reinstale pelo ZIP ou peça à TI um proxy fixo ou a liberação de `github.com` e `release-assets.githubusercontent.com`.
6. **A aba ThinQQ não aparece.** *O que você vê:* o Excel abre sem a aba depois de instalar. *O que fazer:* feche **todas** as janelas do Excel (confira o Gerenciador de Tarefas), rode `Instalar-ThinQQ.cmd` de novo e abra o Excel. Em Arquivo > Opções > Suplementos > Gerenciar: Suplementos do Excel > Ir..., o ThinQQ precisa estar marcado (se estiver em "Itens Desabilitados", reabilite). Com a aba visível, ThinQQ > Sobre mostra as versões do carregador e do núcleo instalados. Se o instalador disser "Feche o Excel antes de instalar", encerre todas as janelas e tente de novo.

### Ambientes corporativos (TI)

Para quem administra os computadores. O ThinQQ chega **sem assinatura digital**, e o instalador cria um Local Confiável **por usuário**, que uma política da empresa pode ignorar ou proibir. Em ambientes que bloqueiam macros não assinadas, o roteiro é assinar os dois arquivos com um certificado da própria empresa:

1. **Certificado interno de assinatura de código**, emitido pela autoridade certificadora da empresa.
2. **Assinar os dois `.xlam`** (`ThinQQ.xlam` e `ThinQQ.core.xlam`):
   - *Manual (sempre disponível):* abra o arquivo no Excel, Alt+F11 (Editor do VBA) > Ferramentas > **Assinatura Digital** > escolha o certificado > salve (assinatura V3 do projeto VBA).
   - *Automático, num servidor de build:* `signtool` do Windows SDK mais os **Office Subject Interface Packages** (DLLs de assinatura de projetos VBA, registradas com administrador), pelo `tools/Sign-ThinQQ.ps1` ou por `pwsh build/release.ps1 ... -SignPfx <arquivo.pfx>` (assina os `.xlam` antes de calcular os SHA-256 do manifesto).
3. **Distribuir o certificado em Editores Confiáveis** (Trusted Publishers) por GPO: Configuração do Computador > Políticas > Configurações do Windows > Configurações de Segurança > Políticas de Chave Pública > Editores Confiáveis. Para um único usuário, `tools/Trust-ThinQQPublisher.ps1 -CerPath arquivo.cer` (`-Scope LocalMachine` exige administrador).
4. **Ou um local confiável corporativo:** o instalador já registra `%LOCALAPPDATA%\ThinQQ` (com subpastas) por usuário; uma GPO pode distribuir um local confiável.
5. **Política do Office:** com "Only trust VBA macros that use V3 signatures" os projetos precisam ter a assinatura V3 que o Editor do VBA grava.
6. **Reassine a cada release** (os arquivos do build são novos a cada versão) e publique os SHA-256 dos arquivos assinados. **Atenção às atualizações:** depois que a TI assina os dois arquivos, a atualização automática pela fonte padrão instalaria o próximo `.xlam` sem assinatura e, com uma política de só macros assinadas, o ThinQQ deixaria de carregar. Aponte a fonte de atualização (ThinQQ > Configurações, ou `install.ps1 -UpdateSource`) para um `manifest.json` interno com os arquivos assinados (gerado por `release.ps1 -SignPfx` ou com os SHA-256 recalculados depois de assinar pelo Editor do VBA).

O que o repositório entrega:

- `tools/New-ThinQQCert.ps1`: cria um certificado autoassinado de **teste** (por usuário, sem administrador), exporta o `.cer` e o `.pfx` e instala o `.cer` em Editores Confiáveis; com `-TrustRoot` também o instala nas autoridades raiz do usuário (o Windows pede confirmação), sem o que o Office mostra o editor como "não verificado". Em produção use o certificado da empresa.
- `tools/Trust-ThinQQPublisher.ps1` e `tools/Sign-ThinQQ.ps1` (acima) e o parâmetro `-SignPfx` do `build/release.ps1`.

Limites: sem o Windows SDK e as SIPs a assinatura automática não roda (o script sai com código 2 e instruções, e o `release.ps1 -SignPfx` não gera o release). O ramo `signtool` mais SIPs **não foi executado ponta a ponta** na máquina de desenvolvimento (sem SDK nem SIPs): abra o `.xlam` assinado no Excel e confira antes de distribuir. A verificação do `signtool` com raiz não confiável (certificado autoassinado) vira aviso, e `-SkipVerify` a pula. O caminho manual pelo Editor do VBA é o padrão do Office e não depende do SDK.

## Primeiros passos

1. Digite uma tabela com cabeçalho e categorias à esquerda e clique numa célula dela.
2. Na aba **ThinQQ**, grupo **Inserir gráfico**, abra **Colunas** e clique em **Colunas empilhadas**: o gráfico aparece à direita, com totais e rótulos.
3. Troque um número da tabela: colunas, totais e rótulos acompanham.
4. Para ver o que mais o ThinQQ faz, abra a pasta de demonstração `exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx` (vem no ZIP; [no GitHub](exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx)): a primeira aba é um índice com 28 exemplos, de E01 a E28; clique no nome do exemplo para ir ao gráfico. Os gráficos também funcionam sem o ThinQQ instalado, mas setas, colchetes e marcas de corte só se recalculam com ele. ThinQQ > Exemplos gera 27 exemplos menores numa pasta nova e ThinQQ > Tutorial abre o tutorial (`docs/tutorial/tutorial.html`), que também explica a pasta de demonstração. O tutorial traz uma galeria com a imagem de cada um dos 28 exemplos (embutidas no próprio arquivo, que continua único e funciona sem rede).

## Recursos

**Fase 1 (1.0.0)**

- Aba própria na faixa de opções, com opções que mudam conforme o gráfico selecionado.
- Colunas empilhadas (com totais), agrupadas, 100% e linha; cascata com marcador `e` para totais e
  subtotais calculados e barras que cruzam o zero.
- Rótulos ligados às células, conectores entre colunas ou barras, paletas (Padrão ThinQQ, Tons de cinza,
  Cores do tema), formatos numéricos (inteiro, decimais, percentual, milhares, milhões) e espaço entre colunas.
- Setas de CAGR e de diferença que se reposicionam quando os dados ou o tamanho do gráfico mudam.
- Pasta de exemplos (ThinQQ > Exemplos) e tutorial (ThinQQ > Tutorial).
- Instalação por usuário, sem administrador. Atualização dentro do Excel, por manifesto (HTTP ou pasta),
  com conferência SHA-256 e troca do núcleo sem reiniciar o Excel.

**Fase 2 (1.1.0)**

- Faixa de opções em menus por família (Colunas, Barras, Cascata, Linha e área, Mais tipos) e, no grupo do gráfico
  selecionado, os menus Mais opções e Setas. Os controles só ficam ativos nos tipos que os aceitam.
- Barras horizontais (empilhadas, agrupadas e 100%) com totais no fim das barras, e Cascata (barras).
- Cascata com várias séries (até 6, com legenda) e "de cima para baixo" (`e` na primeira linha e um número na última).
- Área empilhada e 100%; Combinação (colunas empilhadas + última coluna como linha, com eixo secundário opcional);
  Legenda (Direita, Nenhuma, Nome no fim da linha).
- Mekko (%) e Mekko (unidades), Gantt (escala automática de semanas a anos, marcos, colchetes), Pizza e Rosca.
- Linha de valor (média), Indicador 100% (ligado por padrão nos 100% e no Mekko (%)), Mesma escala e Escala automática.
- 18 exemplos em duas planilhas (ThinQQ > Exemplos; na 1.2.0 são 26, em três) e tutorial com uma seção por tipo.

**Fase 3 (1.2.0)**

- Três tipos novos no menu Mais tipos (a faixa passou a oferecer 20 tipos, e 21 com a cascata comparativa da 1.3.0): **Borboleta** (duas séries espelhadas, com os nomes no centro e os totais no fim da barra direita), **Dispersão** e **Bolhas** (uma série por grupo, nomes por ponto, Linha de tendência e Quadrantes (médias)).
- **Rótulos com soma consistente** (os rótulos dos segmentos somam exatamente o total mostrado: 34% + 33% + 33% = 100%) e **Arredondar intervalo...** (a mesma conta numa tabela comum, em fórmulas da planilha que não dependem do ThinQQ, em até 1.000 células).
- **Corte de eixo...** em colunas e barras empilhadas e agrupadas: a coluna gigante é cortada, com marca de corte, e o total real continua no rótulo.
- Menu **PowerPoint**: Copiar como imagem e Copiar como gráfico vinculado (a esta pasta de trabalho) no slide da janela ativa do PowerPoint.
- **Pasta de demonstração** `exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx`, com 28 exemplos (E01 a E28) e um índice de links, distribuída dentro do ZIP; as imagens dos gráficos estão embutidas na galeria do tutorial (`tools/Export-DemoChartImages.ps1` e `tools/Build-TutorialGallery.ps1`); a planilha Exemplos 3 completa os 27 exemplos de ThinQQ > Exemplos.
- Ícones próprios na faixa de opções; scripts de certificado e de assinatura digital para a TI (`tools/`) e `release.ps1 -SignPfx`.
- Acertos de aparência: eixos da dispersão no formato da origem e apertados em volta dos dados, pizza e rosca com até 10 cores, linha da combinação acima dos totais, formato inicial dos rótulos pela primeira célula de valor, rótulos das pontas da área e das barras finas da borboleta.
- Seção **Instalação** completa (passo a passo, atualizações, desinstalação, problemas comuns, repositório privado) e roteiro para a TI.

**Cascata comparativa (1.3.0)**

- **Cascata comparativa** (ThinQQ > Cascata > Cascata comparativa): duas pontes lado a lado, uma por série (ex.: 2024 x 2025), com conectores por série, uma seta de variação em cotovelo entre as duas barras de cada etapa com `(B - A) / A` numa elipse (ou `n/d` com o primeiro valor zero ou sinais opostos; ThinQQ > Setas > Variação em valor mostra a diferença), totais com o percentual do primeiro total ("11 (41%)") e os nomes das séries dentro das barras (ThinQQ > Legenda > Nome dentro das barras) ou na legenda. Tudo em séries nativas ligadas a fórmulas: acompanha os números mesmo sem o ThinQQ; só o afastamento das setas e os rótulos finos são recalculados por ele.

### Galeria de gráficos (1.3.0)

| Menu | Tipos | Tabela de origem |
|---|---|---|
| Colunas | Colunas empilhadas, Colunas agrupadas, Colunas 100% | cabeçalho, categorias na 1ª coluna e uma coluna de valores por série |
| Barras | Barras empilhadas, Barras agrupadas, Barras 100% | a mesma das colunas |
| Cascata | Cascata (colunas), Cascata (barras), Cascata comparativa | etapa e 1 a 6 colunas de valores (a comparativa: exatamente 2); `e` marca um total |
| Linha e área | Linha, Área empilhada, Área 100%, Combinação (colunas + linha) | a mesma das colunas; a combinação precisa de 2 ou mais colunas de valores |
| Mais tipos | Mekko (%), Mekko (unidades), Gantt, Pizza, Rosca, Borboleta, Dispersão, Bolhas | Mekko (unidades): 2ª coluna = largura; Gantt: atividade, início, fim; pizza e rosca: categorias e uma coluna de valores; borboleta: categorias e exatamente 2 colunas de valores; dispersão: nome, X, Y e grupo opcional; bolhas: nome, X, Y, tamanho e grupo opcional |

O tutorial (`docs/tutorial/tutorial.html`, aberto por ThinQQ > Tutorial) traz a tabela de dados de cada tipo, a tabela
do que cada tipo aceita, as mensagens do ThinQQ e os limites conhecidos.

## Antes de distribuir (checklist do responsável pela publicação)

**Esta lista é de ações suas, fora do código** (as branches já estão no GitHub, em `israeldias-hub/ThinQQ`). A primeira versão a ser publicada é a **1.2.0**: a 1.0.0 (Fase 1) e a 1.1.0 (Fase 2) nunca foram publicadas (não há tags nem releases).

- [x] Decidir a visibilidade do repositório: o repositório é **público** desde 13/10/2026. Se ele voltar a ser **privado**, ninguém sem acesso ao repositório consegue baixar a release (ZIP e arquivos) nem o `release/manifest.json`: o atualizador dos colegas falha em silêncio (sem aviso de versão nova) e o GitHub Pages do tutorial não fica público (no plano gratuito o Pages exige repositório público). Para distribuir, torne o repositório público (depois da limpeza de privacidade abaixo) ou use uma fonte de atualização interna (pasta de rede) e o tutorial que vai no ZIP. As opções estão em "Visibilidade do repositório", na seção Instalação.
- [ ] Gerar os arquivos com `build/release.ps1` e publicar a release **v1.2.0** com os assets (ZIP, `ThinQQ.xlam`, `ThinQQ.core.xlam`, `manifest.json`, `tutorial.html`). O ZIP traz a pasta `exemplos` (o Excel de demonstração e um LEIA-ME) ao lado de `Instalar-ThinQQ.cmd`: confira o conteúdo, e que a pasta de demonstração foi regenerada com o core final (`tools/New-DemoWorkbook.ps1`).
- [ ] Fazer commit de `release/manifest.json` (com URLs reais) e de `version.json` (já em `{ "loader": "1.2.0", "core": "1.2.0" }`) na `main`.
- [ ] **Privacidade antes de publicar** (o repositório ou o Pages): os arquivos versionados já foram limpos (caminhos de perfil do Windows, ids de máquina, PIDs do Excel do autor e e-mail corporativo) e `tests/ps/Test-Privacy.ps1` (rodado por `build/run-tests.ps1`) falha se algum voltar. O que sobra é o **histórico do Git**: o autor e o e-mail dos commits e o texto antigo dos arquivos continuam nos commits antigos. Para publicar só o estado atual, crie um repositório novo a partir da árvore de trabalho (`git archive HEAD`); para manter o histórico, reescreva-o antes de publicar (`git filter-repo --mailmap <arquivo>` troca nome e e-mail dos autores; `git filter-repo --replace-text <arquivo>` troca o texto antigo; instale o `git-filter-repo`, rode numa cópia e confira com `git log --format='%an <%ae>'` e `git grep` no histórico).
- [ ] Habilitar o GitHub Pages para o tutorial online (veja "Tutorial online (GitHub Pages)": pela pasta `/docs`, tudo o que está em `docs/` fica público, inclusive a especificação e o plano em `docs/superpowers`; em repositório privado o Pages só existe em plano pago).
- [ ] Verificar a verificação de atualização a partir de um perfil limpo (usuário do Windows novo ou máquina sem o ThinQQ): instalar uma versão anterior (como a 1.2.0 é a primeira publicada, use um ensaio local com uma versão menor) e conferir que o aviso de atualização aparece e atualiza.
- [ ] Ativar autenticação em dois fatores na conta do GitHub (veja Segurança abaixo).
- [ ] Rodar o roteiro manual no Excel real (seção "Verificações manuais no Excel real"), a começar pela conferência de que a aba ThinQQ aparece.

## Segurança e confiança

- **Os arquivos não vêm assinados digitalmente.** A integridade das atualizações vem da conferência SHA-256 contra o manifesto; a TI pode assinar os dois `.xlam` com um certificado da empresa (veja "Ambientes corporativos (TI)", na seção Instalação).
- **Quem controla a branch `main` do repositório e os assets das releases controla o código que roda no Excel dos colegas** (o manifesto padrão vem de lá). Proteja a conta com autenticação em dois fatores e limite quem pode publicar; em empresas, use uma fonte de atualização interna.
- O SHA-256 é calculado em VBA puro (módulo `modSha256`), sem iniciar processos externos (nada de `certutil`).
- Downloads ficam em `%TEMP%\ThinQQ`, fora do local confiável; só depois de conferidos vão para `%LOCALAPPDATA%\ThinQQ`.

### O que o instalador altera

- Arquivos em `%LOCALAPPDATA%\ThinQQ` (loader na raiz, core e tutorial em `versions\<versão>`).
- `HKCU\Software\Microsoft\Office\16.0\Excel\Options`: valor `OPEN` (ou `OPEN1`, `OPEN2`...) que registra o suplemento.
- `HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ`: local confiável para `%LOCALAPPDATA%\ThinQQ\` **com subpastas**.
- `HKCU\Software\VB and VBA Program Settings\ThinQQ`: configurações do atualizador.
- O ZIP também traz a pasta `exemplos` (o Excel de demonstração e um LEIA-ME): o instalador não a instala nem a abre.
- Nada em `HKLM`. O `uninstall.ps1` desfaz tudo.

### Frequência das verificações

O valor `CheckIntervalHours` (horas; padrão 24) em `HKCU\Software\VB and VBA Program Settings\ThinQQ\Updater` controla o intervalo da verificação automática. Para praticamente desligá-la, grave um número muito grande (por exemplo `876000`); o botão de verificação manual continua funcionando sob demanda.

## Desenvolvimento

Pré-requisitos: Windows, Excel 2016+ (os `.xlam` são montados por automação COM), PowerShell 7 (`pwsh`),
Git e Node.js (só para o servidor HTTP de fixtures dos testes do atualizador).

No Excel, habilite uma vez: Arquivo > Opções > Central de Confiabilidade > Configurações da Central de
Confiabilidade > Configurações de Macro > "Confiar no acesso ao modelo de objeto do projeto VBA"
(`AccessVBOM=1`). O build confere isso e avisa se faltar.

```powershell
pwsh build/build.ps1 -Target all -Configuration Release          # monta dist/ThinQQ.xlam e dist/ThinQQ.core.xlam
pwsh build/build.ps1 -Target all -Configuration Release -Install # + instala em %LOCALAPPDATA%\ThinQQ (Excel fechado)
pwsh build/run-tests.ps1                                          # lint, compilação, testes VBA (core e loader), faixa, instalador, documentação
pwsh build/run-tests.ps1 -Target core -NoBuild                    # só o core, reaproveitando o build de teste
pwsh build/run-tests.ps1 -LoaderSwap                              # + teste de troca do loader em Excel real
pwsh build/run-tests.ps1 -Target core -Filter "TestBars|TestGantt" -TimeoutSec 240   # só os testes cujo Modulo.Teste contém o texto (pula as suítes PowerShell)
pwsh tests/ps/Test-Docs.ps1                                       # só a consistência da documentação (tutorial, README, CHANGELOG, faixa)
pwsh tests/ps/Test-Privacy.ps1                                    # só a privacidade: nenhum arquivo versionado com caminho de perfil, id de máquina ou e-mail corporativo
pwsh tests/ps/Test-ReleaseGuard.ps1                                  # só a guarda do release: módulos de teste nos artefatos, manifesto e conteúdo do ZIP (exemplos\, sem macros)
pwsh tests/ps/Test-DemoWorkbook.ps1                                  # só a pasta de demonstração (Excel próprio, somente leitura)
pwsh tests/ps/Test-Signing.ps1                                       # só o ciclo do certificado de teste e o diagnóstico da assinatura (remove o que criou)
pwsh tests/ps/Test-PublishRelease.ps1                                # só o publicar-release (sem rede e sem Excel: backend HTTP falso e repositório git descartável)
pwsh tools/New-DemoWorkbook.ps1 -CoreXlam <core de teste>            # regenera exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx (Excel novo e invisível; use o core de TESTE de uma cópia em C:\temp)
pwsh tools/Export-DemoChartImages.ps1 -OutDir C:\temp\tqq-imagens       # exporta os 27 gráficos da pasta de demonstração em PNG (Excel próprio, invisível, numa cópia do .xlsx) e reduz cada um
pwsh tools/Build-TutorialGallery.ps1 -ImageDir C:\temp\tqq-imagens       # embute as imagens na galeria do tutorial (idempotente; refaça depois de mudar a demonstração)
```

Os testes VBA ficam em `tests/vba/core` e `tests/vba/loader` (`Test*.bas`, importados pelo build sem registro manual).
No core, um arquivo por área: `TestKinds` (matriz de capacidades), `TestStacked`, `TestBars`, `TestWaterfall` e
`TestWaterfall2` (cascata, várias séries, horizontal), `TestArea` (área, combinação, legenda), `TestPie`, `TestMekko`,
`TestGantt`, `TestDecorations` (linha de valor, indicador 100%, mesma escala), `TestButterfly`, `TestScatter`, `TestRounding` (soma consistente e Arredondar intervalo), `TestAxisBreak` (corte de eixo) e `TestPowerPoint` (com dublês do PowerPoint), além de `TestAnnotations`,
`TestGeometry`, `TestCalcSheet`, `TestCoreApi`, `TestSamples` e outros. `tests/ps/Test-Docs.ps1` confere que os rótulos do
tutorial existem na faixa, que a tabela "O que cada tipo aceita" bate com `Kinds_Supports`, que as mensagens citadas
existem no código, que as seções novas, os 28 exemplos da pasta de demonstração e os limites documentados estão no tutorial e no README.

Formatos numéricos: rótulos e células de números do bloco usam `NumberFormatLocal` com `Style_LocalizeCode` (ou
`Style_AssignNumberFormat`). `Range.NumberFormat =` direto só com código invariante (lido em inglês pelo Excel) ou copiado da origem;
casos aceitos na auditoria da 1.1.0, por categoria: texto `"@"` (`modCalcSheet`, `TestWaterfall`, `TestWaterfall2`); `"General"`
(`TestArea`, `TestWaterfall2`); percentual `"0%"`/`"0.0%"` (`modSamples`, `TestArea`); datas `dd/mm`, `mmm/yy`, `dd/mm/yyyy`
(`modChartGantt`, `modSamples`, `TestAnnotations`, `TestArea`, `TestChartData`, `TestGantt`, `TestMekko`); `"#,##0.00000"`
(`TestChartData`); o formato copiado da categoria de origem (`modChartMekko`).

Os testes usam um Excel invisível próprio e pastas temporárias. Feche instâncias do Excel que tenham
`dist\ThinQQ.core.test.xlam` aberto (o arquivo fica travado). O build de teste (`*.test.xlam`) nunca é
publicado: o `release.ps1` recusa artefatos com módulos de teste.

## Publicar uma versão

A publicação é um comando local em duas etapas: `publicar-release.bat` (na raiz do repositório), que chama `tools/Publish-Release.ps1`.
Ele não roda no GitHub Actions porque a construção dos `.xlam` usa o Excel instalado na máquina.

**Antes de começar**, na `main` (atualizada e sem alterações pendentes): escreva as notas da versão nova em `release/notes.txt` (de 6 a 8
linhas com o marcador •; vão para o campo `notes` do manifesto e para o texto da release) e as entradas do `CHANGELOG.md` na seção
`[Não lançado]`, e faça o commit delas pela via normal (pull request).

1. **Preparar:** `publicar-release.bat x.y.z` (por exemplo `publicar-release.bat 1.2.1`). Com dois cliques no arquivo, sem informar a versão, o `.bat` pergunta a versão (x.y.z) e a etapa (1 ou 2); em um terminal, passe a versão como argumento. O script confere as pré-condições e para com uma
   mensagem clara se alguma falhar: árvore limpa; branch atual `main`, igual à `origin/main` depois de um `git fetch`; versão x.y.z maior que
   a do `version.json`; `release/notes.txt` alterado desde a última tag `v*`; entradas em `[Não lançado]` (se estiver vazia, avisa e pergunta);
   `Test-Docs.ps1` e `Test-Privacy.ps1` passando. Depois cria a branch `release/vx.y.z`, atualiza o `version.json` (loader e core), o
   `CHANGELOG.md` (`[Não lançado]` vira `[x.y.z] - AAAA-MM-DD` e uma `[Não lançado]` vazia fica no topo), o rodapé do tutorial e o exemplo de
   pasta do `installer/LEIA-ME.txt`; constrói o release numa cópia em `C:\temp\tqq-release-x.y.z` com `build/release.ps1` (Excel próprio e
   invisível; a cópia é apagada no fim) e guarda o ZIP, `ThinQQ.xlam`, `ThinQQ.core.xlam`, `manifest.json`, `tutorial.html` e um `state.json`
   (SHA-256 e tamanho de cada arquivo, o commit de preparação e a `main` de origem) em `dist\release-x.y.z\`; copia o `manifest.json` para
   `release/manifest.json`, faz o commit `release: ThinQQ x.y.z`, envia a branch (nunca a `main`) e abre a página do pull request. Se falhar
   antes do commit, tudo é desfeito (arquivos, branch e pasta).
2. **Mesclar o pull request** (manual; a página já vem com título e descrição): confira o diff e mescle com "Create a merge commit".
   Recomendado: em Settings > General > Pull Requests do repositório, desmarque "Allow squash merging" e "Allow rebase merging". Se mesmo
   assim o pull request for mesclado por squash ou rebase, a etapa 3 segue quando os arquivos da versão na `main` são idênticos aos do commit
   de preparação; se não forem, ela recusa e aponta a alternativa manual com os arquivos de `dist\release-x.y.z\`. Rode a etapa 3
   **imediatamente** depois de mesclar: a partir do merge, o `release/manifest.json` da `main` já anuncia a versão nova, e quem clicar em
   atualizar antes de os assets existirem recebe um erro de download. Depois da etapa 1 você continua na branch `release/vx.y.z`: no fim,
   volte com `git checkout main` e `git pull`.
3. **Publicar:** `publicar-release.bat x.y.z publicar`. Confere que os arquivos de `dist\release-x.y.z\` não mudaram desde a preparação
   (`state.json`) e são deste repositório (`-Repo`), que a `origin/main` contém a preparação (ou o mesmo conteúdo, num squash) e tem o mesmo
   `release/manifest.json` desses arquivos, que nada mais mudou na `origin/main` depois da preparação (senão a tag apontaria para outro código:
   prepare uma versão nova ou use `-AllowMainChanges`) e, pela API, que a tag `vx.y.z` e a release não existem; cria a release em
   **rascunho** (tag `vx.y.z` no commit da `origin/main`, nome `ThinQQ x.y.z`, texto do `release/notes.txt` da `origin/main`), envia os 5
   assets com os nomes exatos (`ThinQQ-x.y.z.zip`, `ThinQQ.xlam`, `ThinQQ.core.xlam`, `manifest.json`, `tutorial.html`), confere no GitHub o
   tamanho e o SHA-256 de cada um pelo campo `digest` da API e só então publica a release como a mais recente. Sem o campo `digest` a
   conferência falha e nada é publicado (o download do asset pela API redireciona para outro endereço, que o script não segue). Qualquer
   divergência ou falha no meio: nada é publicado, a mensagem diz o que houve e rodar de novo é seguro (um rascunho `ThinQQ x.y.z` deixado
   por uma execução interrompida, com o texto das notas, é apagado e recriado; release publicada, tag existente, rascunho com outro nome ou
   com o texto editado no GitHub fazem o script recusar sem apagar nada).
4. **Verificar** (manual): baixe o ZIP pela página da release, extraia a pasta inteira, rode `Instalar-ThinQQ.cmd` e confira a aba ThinQQ no
   Excel. Os usuários recebem o aviso na próxima abertura do Excel (no máximo uma verificação por dia).

Opções (o `.bat` repassa ao PowerShell): `-DryRun` (ou `set DRYRUN=1`) confere as pré-condições locais e mostra o que faria, sem gravar,
construir, enviar, ler o token nem chamar a API; `-Draft` termina a etapa 3 com a release em rascunho (publique-a depois pela página dela; se
editar o texto do rascunho no GitHub, publique por lá, sem rodar a etapa de novo); `-NoBrowser` não abre o navegador; `-RunTests` roda a
suíte completa (`build/run-tests.ps1 -Target all -LoaderSwap -TimeoutSec 900`) na cópia antes de construir; `-AllowMainChanges` deixa a
etapa 3 publicar mesmo com a `origin/main` alterada depois da preparação (a tag fica na `origin/main` atual); `-Repo` (padrão
`israeldias-hub/ThinQQ`) e `-NotesFile` (padrão `release/notes.txt`). `set NOPAUSE=1` dispensa a pausa no fim. Códigos de saída: 0 ok,
1 falha (a mensagem diz o que ficou feito e como retomar), 2 pré-condição recusada (nada mudou). Exemplo de ensaio completo:
`publicar-release.bat 1.2.1 -DryRun` e `publicar-release.bat 1.2.1 publicar -DryRun`. Se a etapa 1 falhar antes do commit, ela desfaz o
que gravou, menos um arquivo que alguém tenha editado enquanto o build rodava (esse fica, e é dito); o `installer/LEIA-ME.txt` mantém o BOM.

**Token do GitHub.** A etapa 3 usa o login do Git desta máquina: o token vem de `git credential fill` (Git Credential Manager, sem janela:
sem um login guardado, o script para e pede um `git fetch` para fazer o login), fica só na memória do processo e vai apenas no cabeçalho
`Authorization` das chamadas a `api.github.com` e `uploads.github.com` (TLS 1.2; redirecionamentos para outros endereços são recusados). Ele
nunca é impresso, gravado (nem no `state.json`, nem em log ou arquivo temporário) nem passado em linha de comando, e as mensagens de erro da
API são filtradas. A conta precisa poder escrever no repositório e criar releases. Os trechos que lidam com o token ignoram `-Debug`,
`-Verbose` e os padrões de `$PSDefaultParameterValues` de quem chama, e a biblioteca (`tools/lib/Release-Publish.ps1`) desliga o
rastreamento do `Set-PSDebug` na sessão que a carrega, porque o rastreamento mostraria o token. Uma transcrição (`Start-Transcript`) no
Windows PowerShell 5.1 grava o texto bruto das respostas de erro da API (nunca o token, que só vai no cabeçalho `Authorization`); mesmo
assim, quem guarda transcrições não deve publicá-las.

### Alternativa manual

O caminho antigo, passo a passo, continua valendo (por exemplo para publicar numa pasta de rede):

1. Atualize `release/notes.txt` (vai para o campo `notes` do manifesto) e o `CHANGELOG.md`.
2. Rode o release (a partir da raiz do repositório, com o Excel fechado):

   ```powershell
   pwsh build/release.ps1 -Version x.y.z -BaseUrl https://github.com/israeldias-hub/ThinQQ/releases/download/vx.y.z
   # pasta de rede ou OneDrive sincronizado:
   pwsh build/release.ps1 -Version x.y.z -BaseUrl '\\servidor\ThinQQ' -PublishTo '\\servidor\ThinQQ'
   ```

   Parâmetros: `-Version` (x.y.z, vira a versão do core), `-LoaderVersion` (opcional, só quando o loader muda),
   `-BaseUrl` (base das URLs do manifesto), `-PublishTo` (copia os arquivos para uma pasta), `-NotesFile` e `-ManifestUrl`
   (opcional; veja abaixo); `-SignPfx`, `-SignPassword` e `-SignTimestampUrl` (assinatura opcional dos `.xlam`; veja "Ambientes corporativos (TI)").
   O script recusa o tutorial-esqueleto, grava `version.json`, monta a configuração Release e gera
   `dist/manifest.json` (com os SHA-256) e `dist/ThinQQ-x.y.z.zip`. O ZIP leva também a pasta `exemplos` (o Excel de demonstração, copiado de `exemplos/`, e um LEIA-ME curto, `installer/LEIA-ME-exemplos.txt`); o script recusa o release se o ZIP tiver outro `.xlsx` ou `.xlsm`, se o xlsx do ZIP diferir (SHA-256) do do repositório ou se tiver macros.
3. No GitHub, crie a release `vx.y.z` e anexe como assets o ZIP e os arquivos soltos de `dist/`
   (`ThinQQ.xlam`, `ThinQQ.core.xlam`, `manifest.json`, `tutorial.html`): o atualizador baixa esses arquivos individualmente.
4. Copie `dist/manifest.json` para `release/manifest.json` e faça commit junto com `version.json` na
   branch `main` (a fonte padrão de atualização é o `release/manifest.json` da `main`). Só commite o
   manifesto com URLs reais, nunca o de um teste local.
5. Os usuários recebem o aviso na próxima abertura do Excel (no máximo uma verificação por dia).

**O que uma instalação pelo ZIP consulta.** O instalador grava como fonte de atualização (`UpdateSource`) o campo `manifestUrl` do
`manifest.json` do ZIP (só se o usuário ainda não tiver uma fonte própria). Com `-BaseUrl` http(s), como a release do GitHub, esse campo
é o endereço estável `https://raw.githubusercontent.com/israeldias-hub/ThinQQ/main/release/manifest.json` (o mesmo padrão do
carregador), nunca o `manifest.json` anexado à release: quem instalou uma versão pelo ZIP (por exemplo a 1.2.0) passa a ver a seguinte (1.2.1) assim que o `release/manifest.json` da `main` mudar (no fluxo automático, ao mesclar o pull request da etapa 1; na alternativa manual, no passo 4).
Com `-BaseUrl` numa pasta de rede ou local, o campo é `<pasta>\manifest.json`, e cada versão nova deve
ser publicada na mesma pasta (o manifesto é substituído no lugar). `-ManifestUrl` grava outro endereço.

Para ensaiar sem publicar, use uma cópia do repositório (o Excel do usuário pode manter `dist\ThinQQ.core.test.xlam` travado):
`pwsh build/release.ps1 -Version x.y.z -BaseUrl C:\temp\ensaio -PublishTo C:\temp\ensaio`, confira o ZIP e o
`manifest.json` (versões e os dois SHA-256) e apague a pasta. O ensaio não grava `release/manifest.json`.

Detalhes em `release/README.md`.

### Tutorial online (GitHub Pages)

O botão ThinQQ > Tutorial abre o `tutorial.html` da versão instalada; se ele não existir, abre
<https://israeldias-hub.github.io/ThinQQ/tutorial/tutorial.html>. Para essa página existir, habilite uma vez
em Settings > Pages > "Deploy from a branch", branch `main`, pasta `/docs`. Em repositório privado, o GitHub Pages só existe em plano pago: sem ele, o tutorial que vai no ZIP é o que vale.

Atenção: o Pages por branch só serve a raiz (`/`) ou a pasta `/docs`, então mover o tutorial para outra subpasta de `docs/` (como `docs/tutorial-site`) **não** reduz o que é publicado: com `/docs`, tudo sob `docs/` vai ao ar, inclusive `docs/superpowers` (especificação e plano). Para publicar só o tutorial há duas saídas: (1) publicar por um workflow do GitHub Actions (Settings > Pages > Source: GitHub Actions) que envie apenas a pasta do tutorial, mantendo o caminho `tutorial/tutorial.html` para que `TUTORIAL_URL` em `src/loader/modules/modRibbon.bas` continue valendo; ou (2) tirar a especificação e os planos de `docs/` (por exemplo para `design/`) e manter o Pages em `/docs`. Em qualquer caso, confira antes o item de privacidade da checklist acima (o histórico do Git).

## Arquitetura

- **Loader** (`ThinQQ.xlam`): faixa de opções (`customUI14.xml`), inicialização, atualizador e configurações
  (registro em `HKCU\Software\VB and VBA Program Settings\ThinQQ\Updater`). Muda raramente.
- **Core** (`ThinQQ.core.xlam`): motor de gráficos e API pública `Core_*`. O loader abre o core mais novo de
  `%LOCALAPPDATA%\ThinQQ\versions\<x.y.z>\` e encaminha os comandos por `Application.Run`. Toda função `Core_*`
  trata os próprios erros.
- **Troca a quente:** uma atualização do core baixa o arquivo, confere o SHA-256 do manifesto, instala em uma
  pasta de versão nova e recarrega o core sem reiniciar o Excel. O loader é substituído pelo core (com backup
  `ThinQQ.xlam.bak`) em um segundo momento, para que nenhum código do loader esteja em execução.
- **Dados:** cada gráfico tem um bloco de fórmulas na planilha muito oculta `_ThinQQ`, ligado à tabela de origem
  (maior nos tipos novos: 5 linhas por categoria no Mekko e 40 linhas de escala no Gantt); séries auxiliares `tqq_*`
  (XY no grupo secundário nos gráficos horizontais) levam totais, rótulos e linhas de valor; as setas, o indicador 100% e
  os colchetes são formas dentro do gráfico, reposicionadas nos eventos de cálculo e de redimensionamento.
- **Capacidades:** `modChartKinds` guarda a matriz de recursos por tipo (`Kinds_Supports`); a faixa, os comandos de
  opção e a reconstrução ("Alterar dados") consultam a mesma matriz.
- Fonte em texto em `src/`; os `.xlam` são montados pelo build e não ficam no Git.

## Estrutura

```
build/       build, testes e release (PowerShell 7)
src/loader/  fonte do loader (módulos, formulário, faixa de opções em menus: customUI14.xml)
src/core/    fonte do core: um módulo por família de gráfico (modChartStacked: colunas, barras e linha;
             modChartWaterfall; modChartArea; modChartCombo; modChartMekko; modChartGantt; modChartPie; modChartButterfly;
             modChartScatter), modRounding (soma consistente), modAxisBreak (corte de eixo), modPowerPoint,
             a matriz de capacidades (modChartKinds), decorações (modDecorations), anotações, geometria,
             rótulos, estilo, exemplos (modSamples), classes e o formulário de setas
tests/       testes VBA (TAP), testes PowerShell, servidor Node de fixtures
installer/   instalador e desinstalador por usuário, LEIA-ME.txt
release/     notas e manifesto publicado
tools/       publicação da release (Publish-Release, chamado por publicar-release.bat na raiz), certificado e assinatura
             (New-ThinQQCert, Trust-ThinQQPublisher, Sign-ThinQQ) e gerador da pasta de demonstração e da galeria do tutorial
exemplos/    ThinQQ-Demonstracao-Materiais-Construcao.xlsx (Excel de exemplo; vai no ZIP da release)
docs/        tutorial (tutorial.html), especificações e planos
```

Especificação: `docs/superpowers/specs/2026-10-10-thinqq-design.md` (arquitetura, distribuição e atualização; Fase 1).
Plano da Fase 1: `docs/superpowers/plans/2026-10-10-thinqq-fase1.md`.

Fase 2 (1.1.0): adendo à especificação, com a pesquisa do think-cell, as técnicas nativas do Excel escolhidas para
cada tipo e a matriz de capacidades: `docs/superpowers/specs/2026-10-11-thinqq-fase2-design.md`. Plano de execução:
`docs/superpowers/plans/2026-10-11-thinqq-fase2.md`.

Fase 3 (1.2.0): pesquisa, sondas em Excel e PowerPoint reais, decisões e matriz 20 x 20 em
`docs/superpowers/specs/2026-10-11-thinqq-fase3-design.md`. Plano de execução: `docs/superpowers/plans/2026-10-11-thinqq-fase3.md`.

Cascata comparativa (1.3.0): decisões e medições em Excel real em
`docs/superpowers/specs/2026-10-13-thinqq-cascata-comparativa-design.md`. Plano de execução: `docs/superpowers/plans/2026-10-13-thinqq-cascata-comparativa.md`.

## Verificações manuais no Excel real

Os testes automáticos usam um Excel invisível; estas verificações precisam de uma sessão de Excel real (depois de `Instalar-ThinQQ.cmd` ou de `build.ps1 -Install`), com uma tabela 4x3 com cabeçalho para os itens 1 a 11 e as tabelas de exemplo de cada tipo (planilhas "Exemplos 2" e "Exemplos 3") para os demais:

1. A aba ThinQQ aparece na faixa de opções, ao lado de Inserir, com os cinco menus de inserção (Colunas, Barras, Cascata, Linha e área, Mais tipos). Esta é a primeira conferência: sem a aba, nenhum outro item vale.
2. Com células selecionadas, só os cinco menus, Mais opções (com Mesma escala, Escala automática e Arredondar intervalo...), Atualizar e o grupo ThinQQ ficam ativos; o resto do grupo "Gráfico selecionado" (inclusive os menus Setas e PowerPoint) fica cinza.
3. ThinQQ > Colunas > Colunas empilhadas insere o gráfico à direita, selecionado; Totais e Rótulos pressionados; Totais e Conectores ligam e desligam.
4. Formato, Paleta e Espaço (que fica cinza em gráfico de linha); Colunas 100% mostra participações nos segmentos e nasce com "100%" à esquerda.
5. Setas > Seta CAGR abre o formulário e desenha o texto; mudar um valor atualiza; Seta diferença; Remover setas; redimensionar reposiciona.
6. Cascata (colunas) com `e`: barras, rótulos e conectores; clicar numa célula deixa cinza o resto do grupo (Mais opções e Atualizar continuam ativos); Sobre, Tutorial, Configurações, Exemplos.
7. Copiar um gráfico e apagar o original: a cópia mantém os dados; inserir linha na tabela de origem mostra o aviso uma vez.
8. Colunas 100% com Totais desligado: copiar e colar o gráfico, apagar o original e clicar em Atualizar; a cópia mantém as participações e continua certa depois de inserir outro gráfico.
9. Verificar atualizações e um ciclo real de atualização (core e, depois, loader); conferir `Workbooks("ThinQQ.core.xlam")`.
10. Instalar-ThinQQ.cmd e uninstall.ps1 no Excel real; conferir que nenhuma chave extra permanece.
11. Os ícones da faixa não aparecem em branco e são coerentes (Mesma escala, Eixo secundário (linha), Indicador 100%, Escala automática, Colchete (Gantt)).
12. Barras: inserir os três tipos a partir de uma tabela 3x3; totais no fim das barras; mudar um valor e ver os totais acompanharem; redimensionar; Seta CAGR e Seta diferença; Formato 1.234,5; Barras 100% com participações e conectores.
13. Cascata com várias séries (tabela Etapa, Loja, Online com `e` nas linhas de total): segmentos coloridos por série, legenda com 2 entradas, rótulos por segmento, conectores no líquido; mudar -200 para -100 e ver tudo acompanhar; `e` na primeira linha com um número na última: a primeira barra é calculada.
14. Cascata (barras) com a tabela Receita a Lucro do exemplo 9: rótulos à direita ou à esquerda, conectores verticais; redimensionar e mudar valores.
15. Área empilhada e Área 100% (rótulos no meio das faixas, totais acima, bordas brancas); Combinação com uma coluna de percentual (linha no eixo secundário, totais só das colunas); Legenda > Nome no fim da linha num gráfico de linha; Legenda > Nenhuma.
16. Mais opções > Linha de valor (média) num empilhado: linha tracejada com "Média N" à direita; mudar um valor move a linha.
17. Dois gráficos com escalas diferentes selecionados com Ctrl+clique > Mais opções > Mesma escala; depois Escala automática; com um só gráfico selecionado ou nenhum, vale para todos os da planilha e a mensagem informa quantos foram afetados (Escala automática com um selecionado age só nele); com dados acima do máximo fixo, o aviso aparece uma vez.
18. Pizza e Rosca de uma tabela de 3 linhas: Mais opções > Rótulos fora da pizza (nomes nas fatias, legenda some; arraste um rótulo para ver as linhas de chamada; na rosca o item fica cinza); Formato 12% mostra participações; Paleta; as setas ficam cinza.
19. Mekko (%) de uma tabela 3x3: larguras proporcionais aos totais, separadores brancos, participações nos segmentos, larguras no topo, nomes embaixo, "100%" à esquerda; mudar um valor redistribui as larguras ao vivo.
20. Mekko (unidades) com a coluna Largura e Linha de valor (média): média ponderada pela largura.
21. Gantt com 3 atividades (uma sem fim, que vira marco): escala semanal no topo com separadores; mudar o fim para o ano seguinte e clicar Atualizar passa a escala para meses; Setas > Colchete (Gantt); Espaço muda o espaçamento das linhas; a 4.ª coluna ("Responsável") aparece à direita das barras.
22. Alterar dados em barras, Mekko e Gantt mantém o tipo e as opções; numa cascata de 1 para 2 colunas de valores, Legenda passa a ficar ativa sem clicar em outro gráfico.
23. Estado dos controles por tipo: pizza (Totais, Conectores, Espaço e as três setas cinza; Rótulos fora da pizza ativo), rosca (Rótulos fora da pizza cinza), Gantt (Colchete (Gantt) ativo; Seta CAGR, Seta diferença, Formato e Legenda cinza), cascata de 1 série (Legenda cinza), Colunas 100% (Seta diferença cinza; Indicador 100% pressionado), linha (Totais, Conectores e Espaço cinza), combinação (Eixo secundário (linha) ativo), Mekko (%) (setas cinza).
24. Trocar a seleção entre gráficos de tipos diferentes da mesma planilha, só clicando: a faixa muda sem outro gesto.
25. ThinQQ > Exemplos abre uma pasta com três planilhas; percorrer "Exemplos 2" e "Exemplos 3" e conferir cada gráfico contra o título.
26. ThinQQ > Tutorial abre o tutorial novo (cinco menus na maquete, uma seção por tipo e as seções Borboleta, Dispersão e bolhas, Arredondamento consistente, Corte de eixo, Copiar para o PowerPoint, Pasta de demonstração e Instalação em empresas (TI)) e ThinQQ > Sobre mostra Loader 1.2.0 e Core 1.2.0 depois da instalação pelo ZIP.
27. Borboleta (tabela Faixa, Mulheres, Homens com 4 linhas): Mais tipos > Borboleta mostra os dois lados espelhados, os nomes no centro e os totais depois das barras da direita; mudar um valor atualiza; um nome de 25 caracteres abre o vão sem esconder as barras e redimensionar o gráfico recalcula o vão; uma barra muito fina tem o rótulo fora dela; com 3 colunas de valores o aviso "A borboleta precisa de exatamente 2 colunas de valores (lado esquerdo e lado direito)." aparece; Mesma escala entre duas borboletas mantém as duas centralizadas.
28. Dispersão com a coluna Grupo (3 grupos) e Bolhas (com Tamanho): uma série por grupo, nomes ao lado dos pontos, títulos dos eixos iguais aos cabeçalhos, eixos no formato da coluna (percentual) e apertados em volta dos dados; Linha de tendência e Quadrantes (médias) ligam e desligam; um 7.º grupo dá "No máximo 6 grupos (coluna de grupo)."; um grupo novo exige Alterar dados.
29. Rótulos com soma consistente numa coluna 100% (34, 33, 33): os rótulos somam 100%; trocar o Formato muda as casas; desligar volta aos rótulos simples; o mesmo numa área empilhada; num tipo sem o recurso o aviso diz que não se aplica.
30. Arredondar intervalo...: com uma coluna de 5 números selecionada, as duas perguntas aparecem, o destino já preenchido pede "Substituir?", e o aviso final traz o endereço; uma faixa com texto, uma faixa de mais de 1.000 células e uma planilha protegida sem Formatar células são recusadas sem gravar nada; numa origem em percentual, 2 casas dão percentuais inteiros.
31. Corte de eixo... numa coluna empilhada com uma coluna muito maior (exemplo 23): a sugestão aparece; um valor abaixo do mínimo é recusado com a mensagem do quarto do segmento; com o corte aceito há a marca de corte, o total real no rótulo e o rótulo do segmento atravessado na base; ligar a Linha de valor (média) acima do corte avisa; Remover corte de eixo devolve o eixo; Mesma escala não inclui o gráfico com corte; Alterar dados com valores menores retira o corte e avisa.
32. PowerPoint, imagem, com o PowerPoint fechado: Copiar como imagem abre o PowerPoint, cria uma apresentação com um slide em branco e cola a imagem (até 80% da largura, limitado pela altura do slide, centralizada); o aviso diz o slide.
33. PowerPoint, com uma apresentação aberta: a imagem entra no slide em exibição; no Classificador de Slides, no slide selecionado; nada fecha nem troca a apresentação que já estava aberta (inclusive com duas apresentações abertas, a da janela ativa recebe o gráfico).
34. PowerPoint, gráfico vinculado (pasta salva): o gráfico colado é do PowerPoint, Arquivo > Informações > Editar Links para Arquivos (em versões mais antigas, Editar Vínculos com Arquivos) mostra a pasta, as setas de CAGR vêm junto; mudar um valor no Excel, salvar e, com o gráfico selecionado, Design do Gráfico > Atualizar Dados atualiza o gráfico; Quebrar Vínculo incorpora os dados; renomear o arquivo do Excel quebra o vínculo.
35. PowerPoint, avisos: com alterações não salvas o aviso acrescenta a nota da pasta; numa pasta nova o aviso pede para salvar e nada é colado; com a área de transferência ocupada por outro programa o aviso pede para tentar de novo; com o PowerPoint em edição de texto o aviso diz que ele está ocupado; sem o PowerPoint instalado o aviso diz que ele não está instalado; com uma célula selecionada o menu PowerPoint fica cinza.
36. PowerPoint, segurança: depois de qualquer cópia, o PowerPoint e as apresentações que já estavam abertas continuam abertas e intactas; o que o ThinQQ criou (a apresentação nova) fica aberto para você; rodar `pwsh build/run-tests.ps1 -Target core -Filter PowerPoint` com nenhum PowerPoint aberto (os testes reais criam o deles e o fecham; a colagem substitui a área de transferência).
37. Ícones próprios: os controles com ícone do ThinQQ (Colunas, Barras, Cascata, Cascata comparativa, Linha e área, Mais tipos, Borboleta, Dispersão, Bolhas, Mais opções, Rótulos com soma consistente, Linha de tendência, Quadrantes (médias), Corte de eixo..., Arredondar intervalo..., Setas (com Setas de variação e Variação em valor) e PowerPoint) desenham o próprio ícone, nenhum em branco (confira também com os itens cinza).
38. Pasta de demonstração: abrir `exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx` do ZIP, com e sem o ThinQQ instalado: a primeira aba é o índice, cada nome leva ao exemplo e "Voltar ao índice" (A1) volta; os 28 exemplos mostram o que a legenda diz; sem macros e sem aviso de vínculo externo.
39. Instalação pelo ZIP num perfil limpo: o ZIP traz a pasta `exemplos`, `Instalar-ThinQQ.cmd` conclui e o Excel abre com a aba ThinQQ; ThinQQ > Verificar atualizações informa a versão; `uninstall.ps1` deixa o perfil como estava.
40. Com um core 1.1.0 e o loader 1.2.0 (ou um core de teste antigo), Arredondar intervalo... e o menu PowerPoint (com um gráfico selecionado) respondem "Atualize o ThinQQ para usar este recurso." em vez de não fazer nada.
41. Assinatura (TI): com um certificado de teste de `tools/New-ThinQQCert.ps1` (e depois remover o certificado pelo `certmgr.msc`), assinar um `.xlam` pelo Editor do VBA (Ferramentas > Assinatura Digital) e abri-lo; em uma máquina com Windows SDK e Office SIPs, rodar `tools/Sign-ThinQQ.ps1` (ainda não foi executado ponta a ponta) e conferir o resultado no Excel.
42. Cascata comparativa com a tabela da imagem do think-cell (Bruto 27/29 até Líquido e/e): ThinQQ > Cascata > Cascata comparativa mostra as duas pontes, as elipses +7%, +75%, -67%, +67%, -9%, "11 (41%)" e "10 (34%)" e os nomes dentro das barras; ThinQQ > Setas > Variação em valor troca as elipses pela diferença; Legenda > Direita troca os nomes pela legenda; redimensionar mantém as setas acima dos rótulos; abrir a pasta sem o ThinQQ e mudar um número muda as elipses.

## Limitações conhecidas

Os detalhes de cada tipo estão na seção "Limitações" do tutorial.

- Adicionar ou remover linhas na tabela de origem exige o botão Alterar dados (o ThinQQ avisa uma vez por gráfico e por sessão quando detecta a mudança de tamanho).
- Linhas filtradas ou ocultas continuam plotadas; a linha de totais de uma Tabela do Excel entra como categoria; formato `.xls` não é suportado (use `.xlsx`, `.xlsm` ou `.xlsb`); colunas empilhadas com sinais misturados têm rótulo de total líquido no topo da parte positiva.
- Cópias de gráficos do ThinQQ e gráficos movidos para planilha de gráfico mantêm os dados, mas deixam de ser reconhecidos como gráficos do ThinQQ (sem botões nem setas automáticas).
- Copiar uma planilha inteira (Mover ou copiar > Criar uma cópia) copia o gráfico com o mesmo nome: a cópia divide o bloco e as opções com o original e as séries auxiliares (totais, rótulos) continuam na tabela original. O ThinQQ avisa uma vez ao abrir a planilha copiada; use Alterar dados no gráfico da cópia para separá-lo.
- Somente Excel para Windows, Office 16.0; sem Mac nem Excel na web. A compatibilidade com Excel 64 bits é por construção (sem `Declare`), ainda não verificada em teste.
- **Cascata comparativa:** exatamente 2 séries (mais de 2 séries lado a lado não existe); só na vertical; com o primeiro valor zero ou sinais opostos a seta mostra n/d (percentual (B - A) / A); o afastamento das setas e os rótulos finos só se recalculam com o ThinQQ; os conectores passam por cima da barra da outra série. Em gráfico muito estreito ou com cerca de 60 etapas as elipses das setas podem se sobrepor (desligue as setas de variação); com menos de cerca de 92 pt de altura útil e escala automática a reserva de espaço das setas não é recalibrada; com Mesma escala, um gráfico parceiro que estende o mínimo para baixo (ou um gráfico baixo) pode cortar as setas do topo (use Escala automática ou desligue as setas). O percentual dos totais some sozinho quando os dois rótulos de total da mesma etapa se tocariam (a largura média dos dois rótulos maior que a fatia e os totais a menos de cerca de 14 pt de altura); alargue o gráfico. Rótulos de total vizinhos podem se sobrepor em gráfico estreito ou com números de 5 ou mais dígitos (desligue Totais ou alargue o gráfico).
- Cascata com no máximo 6 séries. Pizza e rosca usam só a primeira coluna de valores; valores negativos, zero, texto e erros viram fatias vazias. As cores se repetem a partir da 11.ª fatia (a paleta tem 10 cores).
- A seta de diferença não existe em gráficos 100%, Mekko, pizza e Gantt; a seta CAGR não existe em Mekko, pizza e Gantt; o colchete só existe no Gantt.
- Mesma escala é uma escala fixa, não dinâmica: depois de mudar muito os dados, use de novo ou Escala automática; um ajuste manual do eixo é desfeito no próximo cálculo ou redimensionamento. Gráficos 100%, Mekko (%), pizza, rosca, dispersão, bolhas, Gantt e gráficos com corte de eixo não entram. Mesma escala entre uma borboleta e gráficos de outro tipo tira a borboleta do centro (use só entre borboletas). Com um só gráfico selecionado (ou nenhum), a Mesma escala vale para todos os gráficos ThinQQ da planilha.
- A linha de valor só existe na média dos valores. Nome no fim da linha só em linha e combinação.
- Gantt: sem sombreamento de fins de semana, colunas de responsável e observação (só uma 4ª coluna, que vira o rótulo da barra), âncoras, setas de processo nem escala de dias; os separadores da escala são desenhados sobre as barras; o intervalo do eixo só é recalculado pelo ThinQQ (Atualizar ou eventos de cálculo).
- Mekko: sem setas; colunas estreitas ou muitas categorias deixam os rótulos e os nomes sobrepostos.
- O espaço reservado para setas, totais e legenda só é refeito pelos comandos do ThinQQ; trocar a Legenda, desligar o Indicador 100% ou a Linha de valor e usar Escala automática descartam um layout manual da área de plotagem.
- Os eventos que reposicionam as setas só disparam com o gráfico visível; Atualizar refaz o layout.
- **Borboleta:** exatamente 2 colunas de valores; na borboleta que define a Mesma escala, um nome longo pode encostar nas barras; o vão central é recalculado pelo ThinQQ (Atualizar, redimensionar); rótulos de barras muito finas ficam fora da barra e só ficam ocultos se nem fora cabem.
- **Dispersão e bolhas:** inserir 300 linhas com 4 grupos leva cerca de 28 s; uma linha acrescentada depois da última linha da tabela não é detectada; o ThinQQ regrava a posição e a cor de um rótulo de bolha mudado à mão; um grupo novo exige Alterar dados; colunas extras depois da coluna de grupo são ignoradas; rótulos de pontos vizinhos podem se sobrepor.
- **Arredondamento:** o Arredondar intervalo... aceita até 1.000 células; texto digitado na origem dá `#VALOR!`; uma origem em % mantém o %; planilha protegida precisa permitir Formatar células; a sobrescrita do destino não pode ser desfeita. Nos rótulos com soma consistente, uma série só reordenada à mão é ligada por posição, e depois de apagar uma série à mão os segmentos voltam aos rótulos simples, sem aviso.
- **Corte de eixo:** só em colunas e barras empilhadas e agrupadas; não é compressão de escala; a linha de valor (média) que fica acima do corte some (com aviso ao ligar); um gráfico com corte não entra num grupo de Mesma escala (os dois são excludentes).
- **Combinação:** com Totais e eixo secundário a linha sobe para uma faixa acima dos totais (e o máximo do eixo das colunas, oculto, sobe junto); com Mesma escala o comportamento anterior volta e a linha pode cruzar os totais.
- **Área:** o rótulo da primeira ou da última categoria que encosta no do vizinho fica oculto e volta quando cabe.
- **Copiar para o PowerPoint:** a colagem usa a área de transferência do Windows e substitui o que estava nela; o destino é o slide da janela ativa do PowerPoint; uma pasta de trabalho aberta de um endereço da web (SharePoint ou OneDrive) pode não manter o vínculo do gráfico vinculado; renomear ou mover o arquivo do Excel quebra o vínculo.
- O texto das setas de CAGR e de diferença e dos colchetes é calculado pelo ThinQQ: sem ele, o texto não se recalcula quando os números mudam.
- **Assinatura:** os arquivos não vêm assinados; o ramo `signtool` mais Office SIPs de `tools/Sign-ThinQQ.ps1` nunca foi executado ponta a ponta na máquina de desenvolvimento (sem SDK nem SIPs); o caminho padrão do Office, que não depende do SDK, é Ferramentas > Assinatura Digital no Editor do VBA.
- Macros precisam estar habilitadas; políticas corporativas podem exigir local confiável ou assinatura digital (veja "Ambientes corporativos (TI)").

## Licença

Nenhuma licença de código aberto foi escolhida: todos os direitos reservados. O repositório não tem arquivo `LICENSE`; para permitir o uso por terceiros, o autor precisa escolher e adicionar uma licença.
