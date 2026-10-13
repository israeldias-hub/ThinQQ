# Publicação de um release

## Fluxo automático (publicar-release.bat)

Um comando local em duas etapas, `publicar-release.bat` na raiz do repositório (chama `tools/Publish-Release.ps1`; Windows PowerShell 5.1
ou pwsh 7). Detalhes, pré-condições e opções no README, seção "Publicar uma versão".

0. Na `main`: notas da versão nova em `release/notes.txt` e entradas do `CHANGELOG.md` em `[Não lançado]`, já commitadas.
1. `publicar-release.bat x.y.z`: confere as pré-condições, cria a branch `release/vx.y.z` com `version.json`, `CHANGELOG.md`, rodapé do
   tutorial, exemplo do LEIA-ME e `release/manifest.json` da versão nova, constrói numa cópia em `C:\temp`, guarda os 5 arquivos e o
   `state.json` em `dist\release-x.y.z\`, envia a branch e abre o pull request.
2. Mescle o pull request (manual, com "Create a merge commit"; recomendado desmarcar "Allow squash merging" e "Allow rebase merging" em
   Settings > General > Pull Requests) e rode o passo 3 IMEDIATAMENTE: a partir do merge a `main` já anuncia a versão nova. Num merge por
   squash ou rebase, o passo 3 só segue se os arquivos da versão na `main` forem idênticos aos da preparação. Depois, `git checkout main` e
   `git pull` (a etapa 1 deixa você na branch `release/vx.y.z`).
3. `publicar-release.bat x.y.z publicar`: cria a release em rascunho, envia os assets, confere tamanho e SHA-256 de cada um pelo campo
   `digest` da API (sem ele, nada é publicado) e publica. Com `-Draft` a release fica em rascunho; rodar de novo depois de uma falha é seguro
   (um rascunho editado no GitHub não é apagado). Se a `origin/main` mudou depois da preparação, a etapa recusa (ou segue com
   `-AllowMainChanges`).
4. Verifique a instalação em Excel real (manual): baixe o ZIP pela release, extraia a pasta inteira, rode `Instalar-ThinQQ.cmd` e confira
   a aba ThinQQ.

`-DryRun` (ou `set DRYRUN=1`) mostra o que cada etapa faria sem gravar, construir, enviar, ler o token nem chamar a API. O token do GitHub vem
de `git credential fill` (o login do Git desta máquina), fica só na memória, vai só no cabeçalho `Authorization` das chamadas a
`api.github.com` e `uploads.github.com` e nunca é impresso nem gravado. Testes: `tests/ps/Test-PublishRelease.ps1` (sem rede e sem Excel).
A biblioteca desliga o rastreamento do `Set-PSDebug` na sessão que a carrega (ele mostraria o token) e ignora `-Debug`/`-Verbose` padrão nos
trechos com o token; uma transcrição (`Start-Transcript`) no 5.1 grava o texto bruto dos erros da API (não o token): não publique transcrições.

## Alternativa manual

1. Rode `pwsh build/release.ps1 -Version x.y.z -BaseUrl <endereço>` (ex.: `https://github.com/israeldias-hub/ThinQQ/releases/download/vx.y.z`
   ou uma pasta de rede, usando também `-PublishTo` para copiar os arquivos).
2. Publique no `BaseUrl`: o ZIP `dist/ThinQQ-x.y.z.zip` e os arquivos soltos `ThinQQ.xlam`, `ThinQQ.core.xlam`,
   `manifest.json` e `tutorial.html` (o atualizador baixa esses arquivos individualmente). Em um GitHub Release, anexe todos como assets.
   O ZIP traz, além do pacote de instalação, a pasta `exemplos` (o Excel de demonstração `exemplos/ThinQQ-Demonstracao-Materiais-Construcao.xlsx` e
   o `LEIA-ME.txt` de `installer/LEIA-ME-exemplos.txt`); o `release.ps1` recusa o ZIP se tiver outro `.xlsx` ou `.xlsm`, se o xlsx diferir (SHA-256) do
   repositório ou se tiver macros. Regenere a demonstração com o core final (`tools/New-DemoWorkbook.ps1`) antes do release.
3. Copie `dist/manifest.json` para `release/manifest.json` e faça commit junto com `version.json`.
   O `release/manifest.json` só deve ser commitado com URLs reais (nunca de um release de teste local).
4. Fonte de atualização de quem instala pelo ZIP: o instalador grava o campo `manifestUrl` do `manifest.json`. Com `-BaseUrl` http(s)
   ele é o endereço estável `https://raw.githubusercontent.com/israeldias-hub/ThinQQ/main/release/manifest.json` (o `release/manifest.json`
   da `main`, o mesmo padrão do carregador), nunca o asset da release; com pasta de rede/local é `<pasta>\manifest.json` (republique cada
   versão na mesma pasta). `-ManifestUrl` troca o endereço.
5. Edite `release/notes.txt` antes do release: o texto vai para o campo `notes` do manifesto. Mantenha o campo `notes` por último no manifesto (o leitor JSON simples do atualizador procura as chaves por texto; `notes` pode conter qualquer texto, então deve vir depois dos campos de versão e de URL).

## Antes de distribuir (ações do responsável, fora do código)

- [ ] Publicar o conteúdo do repositório no GitHub.
- [ ] Gerar os arquivos com `release.ps1` e publicar a release `v1.2.0` (a primeira publicada: a 1.0.0 e a 1.1.0 nunca saíram) com os assets (ZIP, `ThinQQ.xlam`, `ThinQQ.core.xlam`, `manifest.json`, `tutorial.html`).
- [ ] Fazer commit de `release/manifest.json` e de `version.json` na `main`.
- [ ] Privacidade antes de publicar: os arquivos versionados já foram limpos e `tests/ps/Test-Privacy.ps1` falha se um caminho de perfil do Windows, um id de máquina ou um e-mail corporativo voltar; falta decidir o histórico do Git (autor e e-mail dos commits e o texto antigo dos arquivos): publicar um repositório novo a partir do estado atual ou reescrever o histórico (veja o README, "Antes de distribuir").
- [ ] Habilitar o GitHub Pages (Settings > Pages > `main` / `/docs`). Atenção: o Pages por branch só serve `/` ou `/docs` e publica tudo sob `docs/`, inclusive `docs/superpowers`; mover o tutorial para outra subpasta de `docs/` não reduz isso. Para publicar só o tutorial, use um workflow do GitHub Actions que envie só a pasta do tutorial (mantendo `tutorial/tutorial.html`) ou tire a especificação e os planos de `docs/`.
- [ ] Verificar a atualização a partir de um perfil limpo.
- [ ] Ativar autenticação em dois fatores na conta do GitHub: quem controla a `main` e os assets controla o código nas máquinas dos colegas (os `.xlam` não vêm assinados; a TI pode assiná-los: veja "Ambientes corporativos (TI)" no README).
- [ ] Repositório privado: quem não tem acesso não baixa a release nem alcança o manifesto, e o GitHub Pages do tutorial exige repositório público (ou plano pago). Decida entre convidar colaboradores, tornar o repositório público ou publicar o ZIP e o manifesto numa pasta compartilhada com `-BaseUrl` e `-PublishTo` (veja "O repositório é privado", na seção Instalação do README).
