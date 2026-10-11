# ThinQQ

Suplemento do Excel (VBA, Windows) para criar gráficos de apresentação no estilo think-cell: colunas
empilhadas, agrupadas, 100%, linha e cascata (ponte), com totais, rótulos vivos, conectores e setas de
CAGR e de diferença. Os gráficos gerados são gráficos nativos do Excel, ligados aos dados por fórmulas:
quem recebe o arquivo não precisa ter o ThinQQ instalado.

Repositório: <https://github.com/israeldias-hub/ThinQQ>

## Recursos

- Aba própria na faixa de opções, com opções que mudam conforme o gráfico selecionado.
- Colunas empilhadas (com totais), agrupadas, 100% e linha; cascata com marcador `e` para totais e
  subtotais calculados e barras que cruzam o zero.
- Rótulos ligados às células, conectores entre colunas ou barras, paletas (Padrão ThinQQ, Tons de cinza,
  Cores do tema), formatos numéricos (inteiro, decimais, percentual, milhares, milhões) e espaço entre colunas.
- Setas de CAGR e de diferença que se reposicionam quando os dados ou o tamanho do gráfico mudam.
- Pasta de exemplos (ThinQQ > Exemplos) e tutorial (ThinQQ > Tutorial).
- Instalação por usuário, sem administrador. Atualização dentro do Excel, por manifesto (HTTP ou pasta),
  com conferência SHA-256 e troca do núcleo sem reiniciar o Excel.

## Instalação (usuários)

Requisitos: Excel do Office 16.0 (2016, 2019, 2021 ou Microsoft 365) para Windows, 32 ou 64 bits.

1. Baixe `ThinQQ-x.y.z.zip` da página de [Releases](https://github.com/israeldias-hub/ThinQQ/releases) e descompacte a pasta inteira.
2. Feche o Excel e dê dois cliques em `Instalar-ThinQQ.cmd`.
3. Abra o Excel: a aba ThinQQ aparece na faixa de opções.

A instalação manual (se o script for bloqueado), a desinstalação e a solução de problemas estão em
`installer/LEIA-ME.txt` (incluída no ZIP) e no tutorial (`docs/tutorial/tutorial.html`).

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
```

Os testes usam um Excel invisível próprio e pastas temporárias. Feche instâncias do Excel que tenham
`dist\ThinQQ.core.test.xlam` aberto (o arquivo fica travado). O build de teste (`*.test.xlam`) nunca é
publicado: o `release.ps1` recusa artefatos com módulos de teste.

## Publicar uma versão

1. Atualize `release/notes.txt` (vai para o campo `notes` do manifesto) e o `CHANGELOG.md`.
2. Rode o release (a partir da raiz do repositório, com o Excel fechado):

   ```powershell
   pwsh build/release.ps1 -Version x.y.z -BaseUrl https://github.com/israeldias-hub/ThinQQ/releases/download/vx.y.z
   # pasta de rede ou OneDrive sincronizado:
   pwsh build/release.ps1 -Version x.y.z -BaseUrl '\\servidor\ThinQQ' -PublishTo '\\servidor\ThinQQ'
   ```

   Parâmetros: `-Version` (x.y.z, vira a versão do core), `-LoaderVersion` (opcional, só quando o loader muda),
   `-BaseUrl` (base das URLs do manifesto), `-PublishTo` (copia os arquivos para uma pasta) e `-NotesFile`.
   O script recusa o tutorial-esqueleto, grava `version.json`, monta a configuração Release e gera
   `dist/manifest.json` (com os SHA-256) e `dist/ThinQQ-x.y.z.zip`.
3. No GitHub, crie a release `vx.y.z` e anexe como assets o ZIP e os arquivos soltos de `dist/`
   (`ThinQQ.xlam`, `ThinQQ.core.xlam`, `manifest.json`, `tutorial.html`): o atualizador baixa esses arquivos individualmente.
4. Copie `dist/manifest.json` para `release/manifest.json` e faça commit junto com `version.json` na
   branch `main` (a fonte padrão de atualização é o `release/manifest.json` da `main`). Só commite o
   manifesto com URLs reais, nunca o de um teste local.
5. Os usuários recebem o aviso na próxima abertura do Excel (no máximo uma verificação por dia).

Detalhes em `release/README.md`.

### Tutorial online (GitHub Pages)

O botão ThinQQ > Tutorial abre o `tutorial.html` da versão instalada; se ele não existir, abre
<https://israeldias-hub.github.io/ThinQQ/tutorial/tutorial.html>. Para essa página existir, habilite uma vez
em Settings > Pages > "Deploy from a branch", branch `main`, pasta `/docs`.

## Arquitetura

- **Loader** (`ThinQQ.xlam`): faixa de opções (`customUI14.xml`), inicialização, atualizador e configurações
  (registro em `HKCU\Software\VB and VBA Program Settings\ThinQQ\Updater`). Muda raramente.
- **Core** (`ThinQQ.core.xlam`): motor de gráficos e API pública `Core_*`. O loader abre o core mais novo de
  `%LOCALAPPDATA%\ThinQQ\versions\<x.y.z>\` e encaminha os comandos por `Application.Run`. Toda função `Core_*`
  trata os próprios erros.
- **Troca a quente:** uma atualização do core baixa o arquivo, confere o SHA-256 do manifesto, instala em uma
  pasta de versão nova e recarrega o core sem reiniciar o Excel. O loader é substituído pelo core (com backup
  `ThinQQ.xlam.bak`) em um segundo momento, para que nenhum código do loader esteja em execução.
- **Dados:** cada gráfico tem um bloco de fórmulas na planilha muito oculta `_ThinQQ`, ligado à tabela de origem;
  as setas são formas dentro do gráfico, reposicionadas nos eventos de cálculo e de redimensionamento.
- Fonte em texto em `src/`; os `.xlam` são montados pelo build e não ficam no Git.

## Estrutura

```
build/       build, testes e release (PowerShell 7)
src/loader/  fonte do loader (módulos, formulário, faixa de opções)
src/core/    fonte do core (módulos, classes, formulário)
tests/       testes VBA (TAP), testes PowerShell, servidor Node de fixtures
installer/   instalador e desinstalador por usuário, LEIA-ME.txt
release/     notas e manifesto publicado
docs/        tutorial (tutorial.html), especificação e plano
```

Especificação: `docs/superpowers/specs/2026-10-10-thinqq-design.md`. Plano da Fase 1:
`docs/superpowers/plans/2026-10-10-thinqq-fase1.md`.

## Limitações conhecidas

- Somente Excel para Windows, Office 16.0; sem Mac nem Excel na web.
- Cascata vertical com uma série. Barras horizontais, Mekko, Gantt e pizza ficam para a Fase 2.
- A seta de diferença não existe em gráficos 100%.
- Os eventos que reposicionam as setas só disparam com o gráfico visível; Atualizar refaz o layout.
- Macros precisam estar habilitadas; políticas corporativas podem exigir local confiável ou assinatura digital.

## Licença

Todos os direitos reservados.
