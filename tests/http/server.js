// Servidor estático mínimo para fixtures de teste: node server.js <porta> <pasta>
// Rotas de redirecionamento (para o teste do transporte com reserva):
//   /redir-host/<caminho>  302 para o OUTRO host local (127.0.0.1 <-> localhost), mesma porta: redirecionamento entre domínios
//   /redir-same/<caminho>  302 para /<caminho> no mesmo host
const http = require('http'), fs = require('fs'), path = require('path');
const port = Number(process.argv[2] || 18080);
const root = process.argv[3] || '.';
const handler = (req, res) => {
  const urlPath = req.url.split('?')[0];
  const m = /^\/redir-(host|same)\/(.*)$/.exec(urlPath);
  if (m) {
    const host = String(req.headers.host || '127.0.0.1').replace(/:\d+$/, '');
    let loc = '/' + m[2];
    if (m[1] === 'host') loc = 'http://' + (host === 'localhost' ? '127.0.0.1' : 'localhost') + ':' + port + loc;
    res.writeHead(302, { 'Location': loc, 'Cache-Control': 'no-store' });
    res.end();
    return;
  }
  const file = path.join(root, decodeURIComponent(urlPath));
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); res.end('not found'); return; }
    res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Cache-Control': 'no-store' });
    res.end(data);
  });
};
http.createServer(handler).listen(port, '127.0.0.1');
// "localhost" pode resolver para ::1: escuta também no loopback IPv6 (melhor esforço; sem IPv6, segue só no IPv4)
http.createServer(handler).on('error', () => {}).listen(port, '::1');
