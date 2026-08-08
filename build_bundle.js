/**
 * XuY_Sheet 打包脚本
 *
 * 默认行为（源码模式）：
 *   - 生成 js/core/AppTemplate.js：从 index.html 提取 #app 内部 HTML，
 *     作为 window.ExcelApp.APP_HTML 模板，供 XuY_Sheet.setOption(obj) 在任意容器内构建控件。
 *
 * 可选行为（--bundle，按需生成发布包）：
 *   - XuY_Sheet.js  ：内置控件完整 HTML 骨架 + 全部应用模块 + 依赖库（xlsx / fflate / echarts）
 *   - XuY_Sheet.css ：全部样式
 *   （默认不生成这两个文件，避免与"删除打包文件、源码模式开发"的目录约定冲突）
 *
 * 用法：node build_bundle.js            # 仅刷新 js/core/AppTemplate.js
 *       node build_bundle.js --bundle   # 额外生成 XuY_Sheet.js / XuY_Sheet.css
 */
const fs = require('fs');
const path = require('path');

const ROOT = __dirname;

// index.html 中的脚本加载顺序（即依赖顺序）
const SCRIPT_ORDER = [
    'js/core/TableStore.js',
    'js/core/CanvasRender.js',
    'js/core/Chart.js',
    'js/core/Selection.js',
    'js/core/MergeCell.js',
    'js/core/FillHandle.js',
    'js/core/History.js',
    'js/core/Clipboard.js',
    'js/core/Formula.js',
    'js/core/FindReplace.js',
    'js/core/ControlAPI.js',
    'js/core/AppTemplate.js',
    'js/operate/RowColOperate.js',
    'js/operate/CellStyle.js',
    'js/operate/ClearOperate.js',
    'js/operate/FormatBrush.js',
    'js/operate/FreezePane.js',
    'js/menu/TopMenu.js',
    'js/menu/RightMenu.js',
    'js/event/MouseEvent.js',
    'js/event/KeyEvent.js',
    'js/event/ScrollEvent.js',
    'js/ui/SheetBar.js',
    'vendor/xlsx.full.min.js',
    'vendor/fflate.min.js',
    'vendor/echarts.min.js',
    'js/main.js'
];

const CSS_ORDER = [
    'assets/css/index.css',
    'assets/css/main.css'
];

/**
 * 从 index.html 提取 <div id="app"> 的内部 HTML（控件完整骨架）
 * @returns {string}
 */
function extractAppHtml() {
    const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
    const startMarker = '<div id="app">';
    const endMarker = '</div><!-- #app -->';
    const start = html.indexOf(startMarker);
    const end = html.indexOf(endMarker);
    if (start < 0 || end < 0) {
        throw new Error('无法在 index.html 中定位 <div id="app"> 区域');
    }
    let inner = html.slice(start + startMarker.length, end);
    // 剥离 <script> 标签：宿主页面自行加载 JS 模块；
    // 注入模板时 innerHTML 也不会执行它们，剥离后 DOM 更干净。
    inner = inner.replace(/<script[\s\S]*?<\/script>/gi, '');
    return inner.replace(/\r\n/g, '\n').replace(/\s+$/g, '') + '\n';
}

function buildJs() {
    const appHtml = extractAppHtml();
    const parts = [];
    parts.push('/*!');
    parts.push(' * XuY_Sheet - Web 电子表格控件（打包版）');
    parts.push(' * 生成时间: ' + new Date().toISOString());
    parts.push(' * 引入方式: <script src="XuY_Sheet.js"></script>');
    parts.push(' * 初始化:   XuY_Sheet.init({ container: document.getElementById("demo") })');
    parts.push(' */');
    parts.push('(function (global) {');
    parts.push('    // 控件完整 HTML 骨架（顶部菜单/公式栏/画布/工作表栏/菜单等）');
    parts.push('    global.__XUY_SHEET_APP_HTML__ = ' + JSON.stringify(appHtml) + ';');
    parts.push('})(window);');
    parts.push('');

    for (const rel of SCRIPT_ORDER) {
        const abs = path.join(ROOT, rel);
        if (!fs.existsSync(abs)) {
            console.warn('跳过缺失文件: ' + rel);
            continue;
        }
        let code = fs.readFileSync(abs, 'utf8');
        // 移除各文件的 UTF-8 BOM
        if (code.charCodeAt(0) === 0xFEFF) code = code.slice(1);
        parts.push('/* ==================== ' + rel + ' ==================== */');
        parts.push(code);
        parts.push('');
    }
    return parts.join('\n');
}

function buildCss() {
    const parts = [];
    parts.push('/*!');
    parts.push(' * XuY_Sheet - Web 电子表格控件样式（打包版）');
    parts.push(' * 生成时间: ' + new Date().toISOString());
    parts.push(' */');
    for (const rel of CSS_ORDER) {
        const abs = path.join(ROOT, rel);
        if (!fs.existsSync(abs)) {
            console.warn('跳过缺失样式: ' + rel);
            continue;
        }
        let code = fs.readFileSync(abs, 'utf8');
        if (code.charCodeAt(0) === 0xFEFF) code = code.slice(1);
        parts.push('/* ==================== ' + rel + ' ==================== */');
        parts.push(code);
        parts.push('');
    }
    return parts.join('\n');
}

/**
 * 生成源码模板模块 js/core/AppTemplate.js：
 * 将 index.html 的 #app 内部 HTML 提取为 window.ExcelApp.APP_HTML，
 * 供 XuY_Sheet.setOption(obj) / init(config) 在任意容器内快速构建控件。
 * 该文件属于源码（可重复生成），页面需在 main.js 之前引入。
 * @returns {string}
 */
function buildAppTemplate() {
    const appHtml = extractAppHtml();
    const parts = [];
    parts.push('/**');
    parts.push(' * 控件完整 HTML 骨架模板（自动生成，勿手改）');
    parts.push(' * 由 build_bundle.js 从 index.html 的 #app 区域提取，');
    parts.push(' * 供 XuY_Sheet.setOption(obj) / init(config) 在任意容器内构建控件。');
    parts.push(' * 重新生成：node build_bundle.js');
    parts.push(' */');
    parts.push('window.ExcelApp = window.ExcelApp || {};');
    parts.push('window.ExcelApp.APP_HTML = ' + JSON.stringify(appHtml) + ';');
    parts.push('');
    return parts.join('\n');
}

function main() {
    const tpl = buildAppTemplate();
    fs.writeFileSync(path.join(ROOT, 'js/core/AppTemplate.js'), tpl, 'utf8');
    console.log('AppTemplate.js: ' + (tpl.length / 1024).toFixed(1) + ' KB');

    // 仅在显式传入 --bundle 时生成打包发布文件（默认源码模式，不产生 XuY_Sheet.js/.css）
    if (process.argv.includes('--bundle')) {
        const js = buildJs();
        const css = buildCss();
        fs.writeFileSync(path.join(ROOT, 'XuY_Sheet.js'), js, 'utf8');
        fs.writeFileSync(path.join(ROOT, 'XuY_Sheet.css'), css, 'utf8');
        console.log('XuY_Sheet.js  : ' + (js.length / 1024).toFixed(1) + ' KB');
        console.log('XuY_Sheet.css : ' + (css.length / 1024).toFixed(1) + ' KB');
    }
}

main();
