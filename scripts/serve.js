const http = require('http');
const fs = require('fs');
const path = require('path');

const base = path.join(__dirname, '..');
const port = 8842;

const types = { '.html': 'text/html; charset=utf-8', '.json': 'application/json; charset=utf-8' };

http.createServer((req, res) => {
    const filePath = path.join(base, decodeURIComponent(req.url.split('?')[0]) === '/' ? '/하이마트_호남지사_재고대시보드.html' : decodeURIComponent(req.url.split('?')[0]));
    fs.readFile(filePath, (err, data) => {
        if (err) { res.writeHead(404); res.end('not found'); return; }
        const ext = path.extname(filePath);
        res.writeHead(200, { 'Content-Type': types[ext] || 'application/octet-stream' });
        res.end(data);
    });
}).listen(port, () => console.log('serving on http://localhost:' + port));
