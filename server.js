const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = parseInt(process.env.PORT || process.argv[2] || '3000', 10);
const BASE_DIR = __dirname;

const MIME_TYPES = {
    '.html': 'text/html; charset=UTF-8',
    '.js': 'text/javascript; charset=UTF-8',
    '.mjs': 'text/javascript; charset=UTF-8',
    '.css': 'text/css; charset=UTF-8',
    '.json': 'application/json; charset=UTF-8',
    '.xml': 'application/xml; charset=UTF-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon',
    '.txt': 'text/plain; charset=UTF-8'
};

const server = http.createServer((req, res) => {
    // Enable CORS for all requests (important for Office Add-ins and WOPI hosts)
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, HEAD, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', '*');
    res.setHeader('Cache-Control', 'no-cache');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    if (req.method !== 'GET' && req.method !== 'HEAD') {
        res.writeHead(405, { 'Content-Type': 'text/plain' });
        res.end('405 Method Not Allowed');
        return;
    }

    const parsedUrl = url.parse(req.url);
    let reqPath = decodeURIComponent(parsedUrl.pathname);

    // Normalize path to prevent directory traversal
    const safePath = path.normalize(reqPath).replace(/^(\.\.[\/\\])+/, '');
    let filePath = path.join(BASE_DIR, safePath);

    // Ensure requested path is within BASE_DIR
    if (!filePath.startsWith(BASE_DIR)) {
        res.writeHead(403, { 'Content-Type': 'text/plain' });
        res.end('403 Forbidden');
        return;
    }

    fs.stat(filePath, (err, stats) => {
        if (err) {
            console.log(`[404] ${req.method} ${req.url}`);
            res.writeHead(404, { 'Content-Type': 'text/plain' });
            res.end('404 Not Found');
            return;
        }

        // If directory, attempt to serve index.html or taskpane.html
        if (stats.isDirectory()) {
            const indexFile = path.join(filePath, 'index.html');
            const taskpaneFile = path.join(filePath, 'taskpane.html');

            if (fs.existsSync(indexFile)) {
                filePath = indexFile;
            } else if (fs.existsSync(taskpaneFile)) {
                filePath = taskpaneFile;
            } else {
                res.writeHead(403, { 'Content-Type': 'text/plain' });
                res.end('403 Directory Listing Denied');
                return;
            }
        }

        const ext = path.extname(filePath).toLowerCase();
        const contentType = MIME_TYPES[ext] || 'application/octet-stream';

        res.writeHead(200, { 'Content-Type': contentType });

        if (req.method === 'HEAD') {
            res.end();
            return;
        }

        const readStream = fs.createReadStream(filePath);
        readStream.on('error', (streamErr) => {
            console.error('Stream error:', streamErr);
            if (!res.headersSent) {
                res.writeHead(500, { 'Content-Type': 'text/plain' });
            }
            res.end('500 Internal Server Error');
        });

        readStream.pipe(res);
        console.log(`[200] ${req.method} ${req.url} (${contentType})`);
    });
});

server.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}/`);
    console.log(`Serving files from: ${BASE_DIR}`);
    console.log(`\nQuick links:`);
    console.log(`  - Taskpane: http://localhost:${PORT}/taskpane.html`);
    console.log(`  - Commands: http://localhost:${PORT}/commands.html`);
    console.log(`  - Manifest: http://localhost:${PORT}/manifest.xml`);
    console.log(`\nPress Ctrl+C to stop.\n`);
});
