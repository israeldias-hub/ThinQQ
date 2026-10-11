# Publicação de um release

1. Rode `pwsh build/release.ps1 -Version x.y.z -BaseUrl <endereço>` (ex.: `https://github.com/israeldias-hub/ThinQQ/releases/download/vx.y.z`
   ou uma pasta de rede, usando também `-PublishTo` para copiar os arquivos).
2. Publique no `BaseUrl`: o ZIP `dist/ThinQQ-x.y.z.zip` e os arquivos soltos `ThinQQ.xlam`, `ThinQQ.core.xlam`,
   `manifest.json` e `tutorial.html` (o atualizador baixa esses arquivos individualmente). Em um GitHub Release, anexe todos como assets.
3. Copie `dist/manifest.json` para `release/manifest.json` e faça commit junto com `version.json`.
   O `release/manifest.json` só deve ser commitado com URLs reais (nunca de um release de teste local).
4. Edite `release/notes.txt` antes do release: o texto vai para o campo `notes` do manifesto. Mantenha o campo `notes` por último no manifesto (o leitor JSON simples do atualizador procura as chaves por texto; `notes` pode conter qualquer texto, então deve vir depois dos campos de versão e de URL).

## Antes de distribuir (ações do responsável, fora do código)

- [ ] Publicar o conteúdo do repositório no GitHub.
- [ ] Gerar os arquivos com `release.ps1` e publicar a release `v1.1.0` (a primeira publicada: a 1.0.0 nunca saiu) com os assets (ZIP, `ThinQQ.xlam`, `ThinQQ.core.xlam`, `manifest.json`, `tutorial.html`).
- [ ] Fazer commit de `release/manifest.json` e de `version.json` na `main`.
- [ ] Privacidade antes de publicar: os arquivos versionados já foram limpos e `tests/ps/Test-Privacy.ps1` falha se um caminho de perfil do Windows, um id de máquina ou um e-mail corporativo voltar; falta decidir o histórico do Git (autor e e-mail dos commits e o texto antigo dos arquivos): publicar um repositório novo a partir do estado atual ou reescrever o histórico (veja o README, "Antes de distribuir").
- [ ] Habilitar o GitHub Pages (Settings > Pages > `main` / `/docs`). Atenção: o Pages por branch só serve `/` ou `/docs` e publica tudo sob `docs/`, inclusive `docs/superpowers`; mover o tutorial para outra subpasta de `docs/` não reduz isso. Para publicar só o tutorial, use um workflow do GitHub Actions que envie só a pasta do tutorial (mantendo `tutorial/tutorial.html`) ou tire a especificação e os planos de `docs/`.
- [ ] Verificar a atualização a partir de um perfil limpo.
- [ ] Ativar autenticação em dois fatores na conta do GitHub: quem controla a `main` e os assets controla o código nas máquinas dos colegas (a assinatura de código é da Fase 3).
