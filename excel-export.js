/**
 * excel-export.js - 使用 exceljs 将前端表格数据保存为 Excel 文件
 *
 * 数据格式（前端 toJSON 输出）：
 * {
 *   sheets: [{
 *     name: "Sheet1",
 *     totalRows: 100,
 *     totalCols: 30,
 *     defaultColWidth: 80,
 *     defaultRowHeight: 24,
 *     cells: { "0_0": { value, formula, style, border, mergeRange, ... } },
 *     colWidths: { "0": 80, "1": 120 },
 *     rowHeights: { "0": 24 },
 *     mergeRanges: { "0_0": { startRow, startCol, endRow, endCol } },
 *     frozenRows: 0,
 *     frozenCols: 0
 *   }]
 * }
 */

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

/** 保存目录 */
const EXPORT_DIR = path.join(__dirname, 'exports');

/** jszip：用于导出后修补 styles.xml/theme1.xml 的默认字体（保证 Excel 计算列宽 MDW 一致） */
let JSZip = null;
function getJSZip() {
    if (!JSZip) JSZip = require('jszip');
    return JSZip;
}

/**
 * XML 转义
 * @param {*} s
 * @returns {string}
 */
function escapeXml(s) {
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/** 确保导出目录存在 */
function ensureExportDir() {
    if (!fs.existsSync(EXPORT_DIR)) {
        fs.mkdirSync(EXPORT_DIR, { recursive: true });
    }
}

/**
 * 将前端样式对象映射为 exceljs 样式
 * @param {object} style - 前端样式 { bgColor, font, fontSize, bold, italic, fontColor, hAlign, vAlign, ... }
 * @returns {object} exceljs 样式对象（仅包含非空属性，避免空对象覆盖 exceljs 默认值）
 */
function mapStyle(style) {
    if (!style) return {};
    const result = {};

    // 字体（仅在有任意字体属性时才设置，避免空 font 对象清空 exceljs 默认字体）
    const font = {};
    if (style.font) font.name = style.font;
    if (style.fontSize) font.size = style.fontSize;
    if (style.bold) font.bold = true;
    if (style.italic) font.italic = true;
    if (style.strikethrough) font.strike = true;
    if (style.vertAlign === 'superscript' || style.vertAlign === 'subscript') {
        font.vertAlign = style.vertAlign;
    }
    if (style.underline && style.underline !== 'none') {
        // exceljs 支持 true/'single'/'double'，保留双下划线
        font.underline = style.underline === 'double' ? 'double' : true;
    }
    if (style.fontColor) font.color = { argb: colorToARGB(style.fontColor) };
    if (Object.keys(font).length > 0) result.font = font;

    // 填充：solid 实色 + Excel 图案填充
    if (style.bgPattern && style.bgPattern !== 'solid') {
        result.fill = {
            type: 'pattern',
            pattern: style.bgPattern,
            fgColor: { argb: colorToARGB(style.bgColor || '#000000') },
            bgColor: { argb: colorToARGB(style.bgPatternBg || '#FFFFFF') }
        };
    } else if (style.bgColor) {
        const bgLower = String(style.bgColor).toLowerCase();
        if (bgLower !== '#ffffff' && bgLower !== '#fff') {
            result.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: colorToARGB(style.bgColor) }
            };
        }
    }

    // 对齐（仅在有任意对齐属性时才设置）
    const alignment = {};
    if (style.hAlign && style.hAlign !== 'general') {
        // 'general' 是 Excel 默认水平对齐（文本左、数字右、布尔居中），
        // 导出时不写显式 horizontal，保持 Excel General 语义
        const hMap = { left: 'left', center: 'center', right: 'right', justify: 'justify' };
        alignment.horizontal = hMap[style.hAlign] || 'left';
    }
    if (style.vAlign) {
        // exceljs 中垂直居中为 'middle'，部分旧版用 'center'，'middle' 是标准值
        const vMap = { top: 'top', middle: 'middle', bottom: 'bottom' };
        alignment.vertical = vMap[style.vAlign] || 'middle';
    }
    if (style.wrapText) alignment.wrapText = true;
    if (style.textRotate) {
        // OOXML 原始值 → exceljs 模型角度：0-90 原样；91-180 转负角（90-tr）；255 竖排
        const tr = Number(style.textRotate);
        if (tr === 255) alignment.textRotation = 'vertical';
        else if (tr >= 0 && tr <= 90) alignment.textRotation = tr;
        else if (tr > 90 && tr <= 180) alignment.textRotation = 90 - tr;
    }
    if (style.shrinkFit) alignment.shrinkToFit = true;
    if (Object.keys(alignment).length > 0) result.alignment = alignment;

    return result;
}

/**
 * 将前端边框类型 + 宽度映射为 exceljs 边框样式字符串
 * 前端 type 仅取值：solid / dashed / dotted / double
 * exceljs 有效样式：thin / medium / thick / dashed / dotted / double / hair /
 *                   mediumDashed / dashDot / mediumDashDot / dashDotDot / mediumDashDotDot / slantDashDot
 * @param {string} type - 前端边框类型
 * @param {number} width - 前端边框宽度
 * @returns {string} exceljs 边框样式
 */
function mapBorderStyle(type, width) {
    const w = width || 1;
    // 虚线/点线/双线：结合宽度区分细/中
    if (type === 'dashed') return w >= 2 ? 'mediumDashed' : 'dashed';
    if (type === 'dotted') return w >= 2 ? 'mediumDotted' : 'dotted';
    if (type === 'double') return 'double';
    // solid 或未指定：按宽度映射到 thin/medium/thick
    if (w >= 3) return 'thick';
    if (w >= 2) return 'medium';
    return 'thin';
}

/**
 * 将前端边框对象映射为 exceljs 边框
 * @param {object} border - { top, right, bottom, left } each { width, color, type }
 * @returns {object} exceljs 边框对象
 */
function mapBorder(border) {
    if (!border) return {};
    const result = {};

    for (const side of ['top', 'right', 'bottom', 'left']) {
        const b = border[side];
        if (!b || !b.color) continue;
        // 跳过默认网格线色 #d0d0d0（Excel 自带网格线，不应导出为单元格边框）
        if (String(b.color).toLowerCase() === '#d0d0d0') continue;

        result[side] = {
            style: mapBorderStyle(b.type, b.width),
            color: { argb: colorToARGB(b.color) }
        };
    }

    return result;
}

/**
 * 将前端 numFormat 语义标识符 + 格式参数映射为 Excel 数字格式码
 * 前端取值：general / number / currency / accounting / date / time / datetime /
 *           percentage / fraction / scientific / special / text / custom
 * 依据 style 中的 decimals / currencySymbol / dateFormat / timeFormat /
 * fractionType / specialType / numFmtCode 生成完整格式码，保证导出后
 * 再导入时参数（小数位、货币符号、日期样式等）不丢失。
 *
 * @param {string} numFormat - 前端数字格式标识
 * @param {object} [style] - 单元格样式（含格式参数）
 * @returns {string|null} Excel 格式码，null 表示无需设置（通用格式）
 */
function mapNumFormat(numFormat, style) {
    const s = style || {};
    const dec = (s.decimals != null) ? Math.max(0, Math.min(10, Math.round(s.decimals))) : 2;
    const zeros = dec > 0 ? '.' + '0'.repeat(dec) : '';
    const sym = s.currencySymbol || '¥';

    switch (numFormat) {
        case 'general':
            return null;
        case 'number':
            return '#,##0' + zeros;
        case 'currency':
            return '"' + sym + '"#,##0' + zeros;
        case 'accounting':
            return '_-"' + sym + '"* #,##0' + zeros + '_-;-"' + sym + '"* (#,##0' + zeros + ');_-"' +
                sym + '"* "-"??_-;_-@_-';
        case 'percentage':
            return '0' + zeros + '%';
        case 'date':
            return s.dateFormat || 'yyyy-mm-dd';
        case 'time':
            return s.timeFormat || 'hh:mm:ss';
        case 'datetime':
            return (s.dateFormat || 'yyyy-mm-dd') + ' ' + (s.timeFormat || 'hh:mm:ss');
        case 'fraction':
            return fractionFormatCode(s.fractionType);
        case 'scientific':
            return '0' + zeros + 'E+00';
        case 'special':
            return specialFormatCode(s.specialType);
        case 'text':
            return '@';
        case 'custom':
            return s.numFmtCode || '@';
        default:
            return numFormat || null;
    }
}

/**
 * 分数格式码
 * @param {string|number} type - '1'|'2'|'3' 表示分母位数；'8'|'16'|'100' 表示固定分母
 * @returns {string}
 */
function fractionFormatCode(type) {
    const t = String(type == null ? 2 : type);
    if (t === '1') return '# ?/?';
    if (t === '3') return '# ???/???';
    if (t === '8') return '# ?/8';
    if (t === '16') return '# ??/16';
    if (t === '100') return '# ??/100';
    return '# ??/??';
}

/**
 * 特殊格式码
 * @param {string} type - 'zip' | 'phone' | 'chineseUpper'
 * @returns {string}
 */
function specialFormatCode(type) {
    if (type === 'phone') return '000-00000000';
    if (type === 'chineseUpper') return '[DbNum2][$-804]G/通用格式';
    return '00000';
}

/**
 * 将 #RRGGBB 或 RRGGBB 转为 exceljs 的 AARRGGBB 格式
 * @param {string} color - 颜色值
 * @returns {string} ARGB 格式颜色
 */
function colorToARGB(color) {
    if (!color) return 'FF000000';
    let hex = color.replace('#', '').toUpperCase();
    if (hex.length === 6) hex = 'FF' + hex;
    else if (hex.length === 8) hex = hex.toUpperCase();
    else hex = 'FF000000'; // 默认黑色
    return hex;
}

/**
 * 将像素列宽转换为 exceljs 列宽（字符宽度）
 * Excel 存储的 width 与其像素宽满足：wch = px / mdw
 * （读取时 Excel 用 ECMA 公式 trunc(((wch*256+trunc(128/mdw))/256)*mdw) 还原像素，
 *   存储 wch = px/mdw 可精确往返，实测 Excel 保存 85px → width=10.625 即 85/8）
 * @param {number} px - 像素宽度
 * @param {number} [mdw] - 最大数字宽度（像素），缺省 8
 * @returns {number} 字符宽度（1/256 粒度）
 */
function pxToColWidth(px, mdw) {
    if (!px || px <= 0) return 10;
    const m = (mdw && mdw > 0) ? mdw : 8;
    return Math.round(px / m * 256) / 256;
}

/**
 * 将像素行高转换为 exceljs 行高（磅）
 * 1pt = 4/3px（96 DPI），即 1px = 0.75pt；保留 2 位小数避免 0.5pt 级别丢失
 * @param {number} px - 像素高度
 * @returns {number} 磅高度
 */
function pxToRowHeight(px) {
    if (!px || px <= 0) return 15;
    return Math.round(px * 0.75 * 100) / 100;
}

/**
 * 解析 dataURL，返回 exceljs addImage 所需的 { extension, base64 }
 * 仅支持 image/png、image/jpeg、image/gif
 * @param {string} dataURL - 形如 "data:image/png;base64,xxxx"
 * @returns {{extension: string, base64: string}|null}
 */
function parseDataURL(dataURL) {
    if (!dataURL || typeof dataURL !== 'string') return null;
    const match = /^data:image\/(\w+);base64,(.+)$/i.exec(dataURL);
    if (!match) return null;
    let ext = match[1].toLowerCase();
    // exceljs 接受 'png' / 'jpeg' / 'gif'，统一 jpeg 写法
    if (ext === 'jpg') ext = 'jpeg';
    if (ext !== 'png' && ext !== 'jpeg' && ext !== 'gif') return null;
    return { extension: ext, base64: match[2] };
}

/**
 * 获取指定列的像素宽度（基于前端 colWidths/defaultColWidth）
 * @param {object} sheetData - 单个工作表数据
 * @param {number} colIdx - 0-based 列索引
 * @returns {number} 像素宽度
 */
function getColWidthPx(sheetData, colIdx) {
    const cw = sheetData.colWidths && sheetData.colWidths[String(colIdx)];
    if (cw != null) return cw;
    return sheetData.defaultColWidth || 80;
}

/**
 * 获取指定行的像素高度（基于前端 rowHeights/defaultRowHeight）
 * @param {object} sheetData - 单个工作表数据
 * @param {number} rowIdx - 0-based 行索引
 * @returns {number} 像素高度
 */
function getRowHeightPx(sheetData, rowIdx) {
    const rh = sheetData.rowHeights && sheetData.rowHeights[String(rowIdx)];
    if (rh != null) return rh;
    return sheetData.defaultRowHeight || 24;
}

/**
 * 将表格数据写入 Excel 工作簿
 * @param {object} data - 前端表格数据
 * @param {string} filePath - 输出文件路径
 * @returns {Promise<{ success: boolean, filePath: string, sheetsCount: number, cellsCount: number }>}
 */
async function writeTableToExcel(data, filePath) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'XuY_Sheet';
    workbook.created = new Date();

    const sheets = data.sheets || [data]; // 兼容单 sheet 和多 sheet 格式
    let totalCells = 0;
    // 默认字体（工作簿级）：优先顶层 defaultFont，其次第一个 sheet 的 defaultFont
    const defaultFont = data.defaultFont
        || (sheets[0] && sheets[0].defaultFont) || null;

    for (const sheetData of sheets) {
        if (!sheetData || typeof sheetData !== 'object') {
            throw new Error(`工作表数据格式错误: ${JSON.stringify(sheetData)}`);
        }

        const sheetName = sheetData.name || `Sheet${workbook.worksheets.length + 1}`;
        const ws = workbook.addWorksheet(sheetName);

        const cells = sheetData.cells || {};
        const colWidths = sheetData.colWidths || {};
        const rowHeights = sheetData.rowHeights || {};
        // 导入时保留的原始字符宽/磅高（最高优先级，直接写回保证与源文件一致）
        const colWidthsChars = sheetData.colWidthsChars || {};
        const rowHeightsPt = sheetData.rowHeightsPt || {};
        const defaultColWidthChars = sheetData.defaultColWidthChars;
        const defaultRowHeightPt = sheetData.defaultRowHeightPt;
        const mdw = (sheetData.mdw && sheetData.mdw > 0) ? sheetData.mdw : 8;
        const mergeRanges = sheetData.mergeRanges || {};
        const defaultColWidth = sheetData.defaultColWidth || 80;
        const totalCols = sheetData.totalCols || 30;
        const totalRows = sheetData.totalRows || 100;

        // 工作表默认列宽/行高：优先原始字符宽/磅值，否则由像素换算
        const sheetDefaultColWidth = (defaultColWidthChars != null && defaultColWidthChars > 0)
            ? Math.round(defaultColWidthChars * 256) / 256
            : pxToColWidth(defaultColWidth, mdw);
        const sheetDefaultRowHeight = (defaultRowHeightPt != null && defaultRowHeightPt > 0)
            ? Math.round(defaultRowHeightPt * 100) / 100
            : pxToRowHeight(sheetData.defaultRowHeight || 24);
        ws.properties.defaultColWidth = sheetDefaultColWidth;
        ws.properties.defaultRowHeight = sheetDefaultRowHeight;

        // 仅写入显式自定义的列宽（不再为所有 totalCols 列写宽，保留"默认列宽"语义）
        const explicitCols = new Set([
            ...Object.keys(colWidths),
            ...Object.keys(colWidthsChars)
        ]);
        for (const cKey of explicitCols) {
            const colIdx = parseInt(cKey, 10);
            if (isNaN(colIdx) || colIdx < 0 || colIdx >= 16384) continue;
            let width = null;
            if (colWidthsChars[cKey] != null && Number(colWidthsChars[cKey]) > 0) {
                // 原始字符宽度直接写回（保留 Excel 的小数宽度）
                width = Math.round(Number(colWidthsChars[cKey]) * 256) / 256;
            } else if (colWidths[cKey] != null && Number(colWidths[cKey]) > 0) {
                width = pxToColWidth(Number(colWidths[cKey]), mdw);
            }
            if (width != null) ws.getColumn(colIdx + 1).width = width;
        }

        // 构建合并区域索引：将 "row_col" 映射到所属合并区域（便于跳过从属单元格的值写入）
        const mergeIndex = new Map(); // key -> range
        for (const [, range] of Object.entries(mergeRanges)) {
            if (!range || range.startRow == null) continue;
            for (let r = range.startRow; r <= range.endRow; r++) {
                for (let c = range.startCol; c <= range.endCol; c++) {
                    mergeIndex.set(`${r}_${c}`, range);
                }
            }
        }

        // 写入单元格数据与样式
        for (const [key, cell] of Object.entries(cells)) {
            const parts = key.split('_');
            const rowIdx0 = parseInt(parts[0], 10); // 0-based
            const colIdx0 = parseInt(parts[1], 10);
            const row = rowIdx0 + 1; // exceljs 1-based
            const col = colIdx0 + 1;

            if (isNaN(row) || isNaN(col) || row < 1 || col < 1) {
                continue;
            }

            const excelCell = ws.getCell(row, col);

            // 判断是否为合并区域的从属单元格（非左上角）
            const mergeRange = mergeIndex.get(`${rowIdx0}_${colIdx0}`);
            const isMergeSlave = mergeRange &&
                (mergeRange.startRow !== rowIdx0 || mergeRange.startCol !== colIdx0);

            // 值：仅合并区域左上角单元格或普通单元格才写值
            // 从属单元格清空值（exceljs 合并后仅显示左上角值）
            if (!isMergeSlave) {
                if (cell.formula) {
                    // exceljs 公式不含前导 '='，需去除
                    const formulaStr = String(cell.formula).replace(/^=/, '');
                    excelCell.value = { formula: formulaStr };
                } else if (cell.value != null && cell.value !== '') {
                    // 尝试解析数字（排除空字符串和纯空白）
                    const strVal = String(cell.value).trim();
                    if (strVal !== '') {
                        const num = Number(strVal);
                        excelCell.value = !isNaN(num) ? num : cell.value;
                    }
                }
            }

            // 样式：所有单元格（含合并从属）都应用样式，确保合并区域边框完整
            const style = mapStyle(cell.style);
            if (style.font) excelCell.font = style.font;
            if (style.fill) excelCell.fill = style.fill;
            if (style.alignment) excelCell.alignment = style.alignment;

            // 边框：所有单元格都应用（Excel 合并区域边框存储在各边缘单元格上）
            const border = mapBorder(cell.border);
            if (Object.keys(border).length > 0) {
                excelCell.border = border;
            }

            // 数字格式：将前端语义标识符 + 格式参数映射为 Excel 格式码
            if (cell.style && cell.style.numFormat) {
                const fmt = mapNumFormat(cell.style.numFormat, cell.style);
                if (fmt) excelCell.numFmt = fmt;
            }

            totalCells++;
        }

        // 设置行高：优先原始磅值
        const explicitRows = new Set([
            ...Object.keys(rowHeights),
            ...Object.keys(rowHeightsPt)
        ]);
        for (const rowKey of explicitRows) {
            const rowIdx = parseInt(rowKey, 10);
            if (isNaN(rowIdx) || rowIdx < 0 || rowIdx >= 1048576) continue;
            let height = null;
            if (rowHeightsPt[rowKey] != null && Number(rowHeightsPt[rowKey]) > 0) {
                height = Math.round(Number(rowHeightsPt[rowKey]) * 100) / 100;
            } else if (rowHeights[rowKey] != null && Number(rowHeights[rowKey]) > 0) {
                height = pxToRowHeight(Number(rowHeights[rowKey]));
            }
            if (height != null) ws.getRow(rowIdx + 1).height = height;
        }

        // 隐藏行列（Excel 隐藏后不显示，且滚动跳过）
        const hiddenRows = sheetData.hiddenRows || [];
        for (const r of hiddenRows) {
            if (r >= 0 && r < 1048576) ws.getRow(r + 1).hidden = true;
        }
        const hiddenCols = sheetData.hiddenCols || [];
        for (const c of hiddenCols) {
            if (c >= 0 && c < 16384) ws.getColumn(c + 1).hidden = true;
        }

        // 合并单元格（在单元格样式写入之后，exceljs 会保留各单元格已设置的边框样式）
        for (const [, range] of Object.entries(mergeRanges)) {
            if (!range || range.startRow == null) continue;
            ws.mergeCells(
                range.startRow + 1, range.startCol + 1,
                range.endRow + 1, range.endCol + 1
            );
        }

        // 冻结窗格
        const frozenRows = sheetData.frozenRows || 0;
        const frozenCols = sheetData.frozenCols || 0;
        if (frozenRows > 0 || frozenCols > 0) {
            ws.views = [{ state: 'frozen', xSplit: frozenCols, ySplit: frozenRows }];
        }

        // 插入图片（oneCellAnchor：tl 锚定到 row/col + 像素偏移，ext 指定像素尺寸）
        const images = Array.isArray(sheetData.images) ? sheetData.images : [];
        for (const img of images) {
            if (!img || !img.src) continue;
            const parsed = parseDataURL(img.src);
            if (!parsed) continue;
            let imageId;
            try {
                imageId = workbook.addImage({
                    base64: parsed.base64,
                    extension: parsed.extension
                });
            } catch (err) {
                console.warn(`[excel-export] addImage 失败: ${err.message || err}`);
                continue;
            }
            // 偏移量转 fractional col/row（exceljs tl.col 支持 0.5 表示列中点）
            const colW = getColWidthPx(sheetData, img.col) || 80;
            const rowH = getRowHeightPx(sheetData, img.row) || 24;
            const tlCol = (img.col || 0) + (img.offsetX || 0) / colW;
            const tlRow = (img.row || 0) + (img.offsetY || 0) / rowH;
            ws.addImage(imageId, {
                tl: { col: tlCol, row: tlRow },
                ext: { width: img.width || 100, height: img.height || 100 }
            });
        }
    }

    // 写入文件（先生成 buffer，便于修补默认字体后落盘）
    let buffer = await workbook.xlsx.writeBuffer();
    if (defaultFont && defaultFont.name) {
        buffer = await patchWorkbookDefaultFont(buffer, defaultFont);
    }
    fs.writeFileSync(filePath, buffer);

    return {
        success: true,
        filePath: filePath,
        sheetsCount: workbook.worksheets.length,
        cellsCount: totalCells
    };
}

/**
 * 修补导出工作簿的默认字体（Normal 样式字体 + 主题字体）
 * Excel 依据 styles.xml 中 Normal 样式（fonts[0]）的字号/字体计算列宽 MDW，
 * exceljs 固定写 Calibri 11，若源文件默认字体不同，会导致导出文件在 Excel 中
 * 的列宽像素与源文件不一致。这里把 fonts[0] 与 theme1.xml 的字体改为源文件默认字体。
 * @param {Buffer} buffer - exceljs 生成的 xlsx buffer
 * @param {{name: string, size: number|null, scheme?: string|null}} defaultFont - 源文件默认字体
 * @returns {Promise<Buffer>} 修补后的 buffer（失败时原样返回）
 */
async function patchWorkbookDefaultFont(buffer, defaultFont) {
    try {
        const JSZipLib = getJSZip();
        const zip = await JSZipLib.loadAsync(buffer);
        const name = String(defaultFont.name);
        const size = defaultFont.size != null ? Number(defaultFont.size) : null;
        // 还原 scheme：源字体是主题字体（scheme="minor"/"major"）则保留，
        // 否则移除 exceljs 自动添加的 scheme（避免影响 Excel 的 MDW 计算）
        const scheme = defaultFont.scheme ? String(defaultFont.scheme) : null;
        if (!name) return buffer;

        // 1. styles.xml：替换 <fonts> 中第一个 <font> 的 name/sz/scheme
        const stylesXml = await zip.file('xl/styles.xml').async('string');
        if (stylesXml) {
            const patchedStyles = stylesXml.replace(
                /<fonts[^>]*>([\s\S]*?)<\/fonts>/,
                (whole, inner) => {
                    const openTag = whole.slice(0, whole.indexOf('>') + 1);
                    const fonts = inner.replace(/<font>[\s\S]*?<\/font>/, (fontXml) => {
                        let out = fontXml;
                        out = out.replace(/<name[^>]*\/>/, `<name val="${escapeXml(name)}"/>`);
                        out = out.replace(/<sz[^>]*\/>/, size != null ? `<sz val="${size}"/>` : '$&');
                        if (scheme) {
                            out = out.replace(/<scheme[^>]*\/>/, `<scheme val="${escapeXml(scheme)}"/>`);
                        } else {
                            out = out.replace(/<scheme[^>]*\/>/, '');
                        }
                        return out;
                    });
                    return openTag + fonts + '</fonts>';
                }
            );
            zip.file('xl/styles.xml', patchedStyles);
        }

        // 2. theme1.xml：minorFont（对应 Normal 样式字体）的 latin/ea/cs 同步为默认字体
        const themeXml = await zip.file('xl/theme/theme1.xml').async('string');
        if (themeXml) {
            const patchedTheme = themeXml.replace(
                /<a:minorFont>([\s\S]*?)<\/a:minorFont>/i,
                (whole, inner) => {
                    const out = inner
                        .replace(/<a:latin[^>]*\/>/i, `<a:latin typeface="${escapeXml(name)}"/>`)
                        .replace(/<a:ea[^>]*\/>/i, `<a:ea typeface="${escapeXml(name)}"/>`)
                        .replace(/<a:cs[^>]*\/>/i, `<a:cs typeface="${escapeXml(name)}"/>`);
                    return `<a:minorFont>${out}</a:minorFont>`;
                }
            );
            zip.file('xl/theme/theme1.xml', patchedTheme);
        }

        return await zip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });
    } catch (err) {
        console.warn(`[excel-export] 默认字体修补失败（忽略）: ${err.message || err}`);
        return buffer;
    }
}

/**
 * 保存表格数据为 Excel 文件
 * @param {object} data - 前端表格数据
 * @param {string} [filename] - 可选文件名（不含路径）
 * @returns {Promise<object>} 保存结果
 */
async function saveTableAsExcel(data, filename) {
    // 参数校验
    if (!data || typeof data !== 'object') {
        throw new Error('表格数据格式错误: 期望对象');
    }

    const sheets = data.sheets || [data];
    if (!Array.isArray(sheets) || sheets.length === 0) {
        throw new Error('表格数据格式错误: 缺少工作表数据');
    }

    // 确保导出目录存在
    ensureExportDir();

    // 生成文件名
    const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const safeName = filename || `getsheet_${ts}`;
    const fileName = safeName.endsWith('.xlsx') ? safeName : `${safeName}.xlsx`;
    const filePath = path.join(EXPORT_DIR, fileName);

    console.log(`[excel-export] 开始导出: ${fileName}`);
    console.log(`[excel-export] 工作表数量: ${sheets.length}`);

    const result = await writeTableToExcel(data, filePath);

    console.log(`[excel-export] 导出成功: ${filePath}`);
    console.log(`[excel-export] 单元格数: ${result.cellsCount}`);

    return {
        success: true,
        filename: fileName,
        filePath: filePath,
        sheetsCount: result.sheetsCount,
        cellsCount: result.cellsCount,
        message: '文件保存成功'
    };
}

module.exports = {
    saveTableAsExcel,
    writeTableToExcel,
    patchWorkbookDefaultFont,
    mapStyle,
    mapBorder,
    mapBorderStyle,
    mapNumFormat,
    colorToARGB,
    parseDataURL,
    getColWidthPx,
    getRowHeightPx,
    EXPORT_DIR
};
