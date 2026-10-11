# Publicação de um release

1. Rode `pwsh build/release.ps1 -Version x.y.z -BaseUrl <endereço>` (ex.: `https://github.com/israeldias-hub/ThinQQ/releases/download/vx.y.z`
   ou uma pasta de rede, usando também `-PublishTo` para copiar os arquivos).
2. Publique no `BaseUrl`: o ZIP `dist/ThinQQ-x.y.z.zip` e os arquivos soltos `ThinQQ.xlam`, `ThinQQ.core.xlam`,
   `manifest.json` e `tutorial.html` (o atualizador baixa esses arquivos individualmente). Em um GitHub Release, anexe todos como assets.
3. Copie `dist/manifest.json` para `release/manifest.json` e faça commit junto com `version.json`.
   O `release/manifest.json` só deve ser commitado com URLs reais (nunca de um release de teste local).
4. Edite `release/notes.txt` antes do release: o texto vai para o campo `notes` do manifesto.
