const http = require('http');
const fs = require('fs');
const path = require('path');

const port = 3000;
const dir = __dirname;

const types = {
  '.html': 'text/html',
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.png': 'image/png',
  '.svg': 'image/svg+xml'
};

http.createServer((req, res) => {
  let filePath = path.join(dir, req.url === '/' ? 'index.html' : req.url);
  const ext = path.extname(filePath);
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found');
      return;
    }
    res.writeHead(200, { 'Content-Type': types[ext] || 'text/plain', 'Access-Control-Allow-Origin': '*' });
    res.end(data);
  });
}).listen(port, () => {
  console.log(`Server running at http://127.0.0.1:${port}`);
});
