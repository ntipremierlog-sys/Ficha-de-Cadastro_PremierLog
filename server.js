const http = require('http');
const fs   = require('fs');
const path = require('path');

const PORT = 8000;

// ============================================================
// Tratamento global de erros — evita que o servidor caia
// sem aviso quando ocorre um erro não capturado.
// ============================================================
process.on('uncaughtException', function(err) {
  console.error('[' + new Date().toISOString() + '] UNCAUGHT EXCEPTION — servidor continuando:', err.message);
  console.error(err.stack);
});

process.on('unhandledRejection', function(reason, promise) {
  console.error('[' + new Date().toISOString() + '] UNHANDLED REJECTION:', reason);
});

const mimeTypes = {
  '.html': 'text/html; charset=utf-8',
  '.css':  'text/css; charset=utf-8',
  '.js':   'text/javascript; charset=utf-8',
  '.png':  'image/png',
  '.jpg':  'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg':  'image/svg+xml',
  '.pdf':  'application/pdf',
  '.txt':  'text/plain; charset=utf-8',
  '.json': 'application/json; charset=utf-8'
};

// ============================================================
// Headers CORS + Segurança — permite que o formulário seja
// acessado de qualquer origem (necessário para GitHub Pages
// fazer requisições ao servidor local durante desenvolvimento).
// ============================================================
function setCorsHeaders(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('X-Content-Type-Options', 'nosniff');
}

const server = http.createServer(function(req, res) {
  const startTime = Date.now();

  // Preflight OPTIONS para CORS
  if (req.method === 'OPTIONS') {
    setCorsHeaders(res);
    res.writeHead(204);
    res.end();
    return;
  }

  setCorsHeaders(res);

  let filePath = path.join(__dirname, decodeURIComponent(req.url.split('?')[0]));
  if (filePath === __dirname || req.url === '/' || req.url.startsWith('/?')) {
    filePath = path.join(__dirname, 'index.html');
  }

  // Proteção básica: impede path traversal (acesso a arquivos fora da pasta)
  if (!filePath.startsWith(__dirname)) {
    res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('403 Forbidden');
    console.warn('[' + new Date().toISOString() + '] PATH TRAVERSAL BLOQUEADO: ' + req.url);
    return;
  }

  const ext         = path.extname(filePath).toLowerCase();
  const contentType = mimeTypes[ext] || 'application/octet-stream';

  fs.readFile(filePath, function(err, content) {
    const elapsed = Date.now() - startTime;

    if (err) {
      if (err.code === 'ENOENT') {
        res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('404 Not Found');
        console.warn('[' + new Date().toISOString() + '] 404 ' + req.method + ' ' + req.url + ' (' + elapsed + 'ms)');
      } else {
        res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('500 Internal Server Error');
        console.error('[' + new Date().toISOString() + '] 500 ' + req.method + ' ' + req.url + ' — ' + err.message);
      }
    } else {
      res.writeHead(200, { 'Content-Type': contentType });
      res.end(content);
      console.log('[' + new Date().toISOString() + '] 200 ' + req.method + ' ' + req.url + ' (' + elapsed + 'ms)');
    }
  });
});

// Keep-alive e timeout configurados para evitar conexões penduradas
server.keepAliveTimeout = 65000;  // 65 segundos (maior que o timeout de proxies comuns)
server.headersTimeout   = 66000;  // sempre maior que keepAliveTimeout

server.listen(PORT, function() {
  console.log('============================================================');
  console.log('  Premier Logistics — Servidor RH rodando!');
  console.log('  URL: http://localhost:' + PORT + '/');
  console.log('  Iniciado em: ' + new Date().toISOString());
  console.log('============================================================');
});

// Log ao encerrar o servidor graciosamente (Ctrl+C)
process.on('SIGINT', function() {
  console.log('\n[' + new Date().toISOString() + '] Servidor encerrado graciosamente.');
  server.close(function() {
    process.exit(0);
  });
});
