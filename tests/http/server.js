// Servidor estático mínimo para fixtures de teste: node server.js <porta> <pasta>
const http = require('http'), fs = require('fs'), path = require('path');
const port = Number(process.argv[2] || 18080);
const root = process.argv[3] || '.';
http.createServer((req, res) => {
  const file = path.join(root, decodeURIComponent(req.url.split('?')[0]));
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); res.end('not found'); return; }
    res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Cache-Control': 'no-store' });
    res.end(data);
  });
}).listen(port, '127.0.0.1');
