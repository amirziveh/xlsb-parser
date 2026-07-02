import http from 'http';
import fs from 'fs';
import path from 'path';

const MIME = {
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.html': 'text/html;charset=utf-8',
};

http.createServer((req, res) => {
  let p = path.join('public', req.url === '/' ? 'index.html' : req.url.split('?')[0]);
  fs.readFile(p, (err, d) => {
    if (err) { res.writeHead(404); res.end('Not found'); return; }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' });
    res.end(d);
  });
}).listen(8080, () => console.log('XLSB Parser at http://localhost:8080'));
