/**
 * server.js - XuY_Sheet 后端服务
 *
 * 功能：
 *   1. 提供静态文件服务（前端页面）
 *   2. POST /api/export/excel — 接收前端表格数据，保存为 Excel 文件
 *   3. GET  /api/exports/list  — 列出已导出的文件
 *   4. GET  /api/exports/download?file=xxx — 下载已导出的文件
 *
 * 启动: node server.js
 * 默认端口: 8765
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const { saveTableAsExcel, EXPORT_DIR } = require('./excel-export');

const PORT = process.env.PORT || 8765;
const ROOT_DIR = __dirname;

/** MIME 类型映射 */
const MIME = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon',
    '.woff': 'font/woff',
    '.woff2': 'font/woff2',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

/**
 * 读取请求体（JSON）
 * @param {http.IncomingMessage} req
 * @returns {Promise<object>}
 */
function readJsonBody(req) {
    return new Promise((resolve, reject) => {
        const chunks = [];
        let size = 0;
        const MAX_SIZE = 200 * 1024 * 1024; // 50MB

        req.on('data', (chunk) => {
            size += chunk.length;
            if (size > MAX_SIZE) {
                reject(new Error('请求体超过 200MB 限制'));
                req.destroy();
                return;
            }
            chunks.push(chunk);
        });

        req.on('end', () => {
            try {
                const raw = Buffer.concat(chunks).toString('utf-8');
                resolve(JSON.parse(raw));
            } catch (e) {
                reject(new Error(`JSON 解析失败: ${e.message}`));
            }
        });

        req.on('error', reject);
    });
}

/**
 * 发送 JSON 响应
 */
function sendJson(res, statusCode, data) {
    const body = JSON.stringify(data);
    res.writeHead(statusCode, {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-store'
    });
    res.end(body);
}

/**
 * 静态文件服务（带 no-cache 头）
 */
function serveStatic(req, res) {
    let urlPath = decodeURIComponent(req.url.split('?')[0]);
    if (urlPath === '/') urlPath = '/index.html';

    // 安全检查：防止路径遍历
    const safePath = path.normalize(urlPath).replace(/^(\.\.[\/\\])+/, '');
    const filePath = path.join(ROOT_DIR, safePath);

    if (!filePath.startsWith(ROOT_DIR)) {
        sendJson(res, 403, { error: '禁止访问' });
        return;
    }

    fs.stat(filePath, (err, stats) => {
        if (err || !stats.isFile()) {
            sendJson(res, 404, { error: '文件未找到' });
            return;
        }

        const ext = path.extname(filePath).toLowerCase();
        const contentType = MIME[ext] || 'application/octet-stream';

        res.writeHead(200, {
            'Content-Type': contentType,
            'Cache-Control': 'no-store, no-cache, must-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0'
        });

        fs.createReadStream(filePath).pipe(res);
    });
}

/**
 * API 路由处理
 */
async function handleApi(req, res, urlPath) {
    // GET /api/import/file?path=xxx — 按文件路径读取 .xls/.xlsx/.csv 并返回 base64。
    // 支持相对路径（以服务根目录为基准，含路径穿越防护）与绝对路径（本地开发工具场景）。
    if (urlPath === '/api/import/file' && req.method === 'GET') {
        try {
            const u = new URL(req.url, 'http://localhost');
            const rawPath = (u.searchParams.get('path') || '').trim();
            if (!rawPath) {
                return sendJson(res, 400, { success: false, error: '缺少 path 参数（文件路径）' });
            }

            const ext = path.extname(rawPath).toLowerCase();
            if (!['.xls', '.xlsx', '.csv'].includes(ext)) {
                return sendJson(res, 400, {
                    success: false,
                    error: `不支持的文件格式：${ext || '无扩展名'}。仅支持 .xls / .xlsx / .csv`
                });
            }

            // 相对路径：以服务根目录为基准解析，并防止路径穿越到根目录之外
            let filePath;
            if (path.isAbsolute(rawPath)) {
                filePath = path.resolve(rawPath);
            } else {
                const resolved = path.resolve(ROOT_DIR, rawPath);
                if (!resolved.startsWith(ROOT_DIR)) {
                    return sendJson(res, 400, { success: false, error: '非法相对路径：不允许越出服务根目录' });
                }
                filePath = resolved;
            }

            if (!fs.existsSync(filePath)) {
                return sendJson(res, 404, { success: false, error: `文件不存在：${rawPath}` });
            }
            if (!fs.statSync(filePath).isFile()) {
                return sendJson(res, 400, { success: false, error: `路径不是文件：${rawPath}` });
            }

            const buf = fs.readFileSync(filePath);
            return sendJson(res, 200, {
                success: true,
                fileName: path.basename(filePath),
                filePath: filePath,
                size: buf.length,
                base64: buf.toString('base64')
            });
        } catch (err) {
            console.error('[Server] 导入读取失败:', err.message);
            return sendJson(res, 500, { success: false, error: `读取文件失败：${err.message || err}` });
        }
    }

    // POST /api/export/excel — 导出 Excel
    if (urlPath === '/api/export/excel' && req.method === 'POST') {
        try {
            const data = await readJsonBody(req);

            if (!data || typeof data !== 'object') {
                return sendJson(res, 400, { success: false, error: '请求数据格式错误: 期望 JSON 对象' });
            }

            const filename = data.filename || null;
            const tableData = data.sheets ? data : { sheets: [data] };

            const result = await saveTableAsExcel(tableData, filename);

            return sendJson(res, 200, result);
        } catch (err) {
            console.error('[API] 导出失败:', err.message);
            return sendJson(res, 500, {
                success: false,
                error: err.message || '导出失败',
                code: 'EXPORT_ERROR'
            });
        }
    }

    // GET /api/exports/list — 列出已导出文件
    if (urlPath === '/api/exports/list' && req.method === 'GET') {
        try {
            if (!fs.existsSync(EXPORT_DIR)) {
                return sendJson(res, 200, { success: true, files: [] });
            }

            const files = fs.readdirSync(EXPORT_DIR)
                .filter(f => f.endsWith('.xlsx'))
                .map(f => {
                    const stat = fs.statSync(path.join(EXPORT_DIR, f));
                    return {
                        filename: f,
                        size: stat.size,
                        createdAt: stat.mtime.toISOString()
                    };
                })
                .sort((a, b) => b.createdAt.localeCompare(a.createdAt));

            return sendJson(res, 200, { success: true, files });
        } catch (err) {
            console.error('[API] 列出文件失败:', err.message);
            return sendJson(res, 500, { success: false, error: err.message });
        }
    }

    // GET /api/exports/download?file=xxx — 下载文件
    if (urlPath === '/api/exports/download' && req.method === 'GET') {
        try {
            const params = new URLSearchParams(req.url.split('?')[1] || '');
            const filename = params.get('file');

            if (!filename) {
                return sendJson(res, 400, { success: false, error: '缺少 file 参数' });
            }

            // 安全检查：防止路径遍历
            const safeName = path.basename(filename);
            const filePath = path.join(EXPORT_DIR, safeName);

            if (!filePath.startsWith(EXPORT_DIR)) {
                return sendJson(res, 403, { success: false, error: '禁止访问' });
            }

            if (!fs.existsSync(filePath)) {
                return sendJson(res, 404, { success: false, error: '文件未找到' });
            }

            const stat = fs.statSync(filePath);
            res.writeHead(200, {
                'Content-Type': MIME['.xlsx'],
                'Content-Disposition': `attachment; filename="${encodeURIComponent(safeName)}"`,
                'Content-Length': stat.size
            });
            fs.createReadStream(filePath).pipe(res);
            return;
        } catch (err) {
            console.error('[API] 下载失败:', err.message);
            return sendJson(res, 500, { success: false, error: err.message });
        }
    }

    // 未知 API
    sendJson(res, 404, { error: `API 未找到: ${req.method} ${urlPath}` });
}

/**
 * 创建 HTTP 服务器
 */
const server = http.createServer(async (req, res) => {
    const urlPath = req.url.split('?')[0];

    // CORS 头
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    // API 路由
    if (urlPath.startsWith('/api/')) {
        try {
            await handleApi(req, res, urlPath);
        } catch (err) {
            console.error('[Server] 未捕获异常:', err);
            sendJson(res, 500, { success: false, error: '服务器内部错误' });
        }
        return;
    }

    // 静态文件
    serveStatic(req, res);
});

// 启动服务器
server.listen(PORT, () => {
    console.log(`========================================`);
    console.log(`  XuY_Sheet Server`);
    console.log(`  端口: ${PORT}`);
    console.log(`  静态文件: http://localhost:${PORT}/`);
    console.log(`  导出 API: POST http://localhost:${PORT}/api/export/excel`);
    console.log(`  文件列表: GET  http://localhost:${PORT}/api/exports/list`);
    console.log(`  文件下载: GET  http://localhost:${PORT}/api/exports/download?file=xxx.xlsx`);
    console.log(`  导出目录: ${EXPORT_DIR}`);
    console.log(`========================================`);
});

// 优雅关闭
process.on('SIGINT', () => {
    console.log('\n[Server] 正在关闭...');
    server.close(() => {
        console.log('[Server] 已关闭');
        process.exit(0);
    });
});

// 未捕获异常处理
process.on('uncaughtException', (err) => {
    console.error('[Server] 未捕获异常:', err);
});

process.on('unhandledRejection', (reason) => {
    console.error('[Server] 未处理的 Promise 拒绝:', reason);
});
