// docx-exporter.js — экспорт Markdown -> .docx с нативными формулами Word (OMML).
//
// Архитектура (аналог пайплайна pandoc, но без внешних бинарников):
//   markdown-it (тот же рендер, что и в ридере) -> HTML
//   -> разбор HTML в блок-структуру (document.json)
//   -> для каждой формулы: capture исходного LaTeX из KaTeX-вывода (formula-capture.js)
//   -> LaTeX -> OMML (latex2omml.js)  [здесь сила: pandoc часто «не видел» $...$ вообще]
//   -> сборка document.xml + zip -> .docx
//
// Используется и в main-процессе Electron (Node), и может быть вызван напрямую.

'use strict';

const zlib = require('zlib');
const { latexToOmml } = require('./latex2omml');
const { captureFromNode } = require('./formula-capture');

// ================================================================ 1. HTML parse
const VOID_TAGS = new Set(['br', 'img', 'hr', 'input', 'meta', 'link']);
// MathML-теги (KaTeX/TeXMath): не поддерживают самозакрытие в «HTML-режиме»,
// поэтому парсим их в XML-режиме (самозакрытие корректно закрывает узел).
const MATHML_TAGS = new Set([
    'math', 'semantics', 'annotation', 'annotation-xml', 'mrow', 'mi', 'mn', 'mo',
    'ms', 'mmultiscripts', 'mprescripts', 'none', 'mtext', 'mspace', 'msub', 'msup',
    'msubsup', 'mfrac', 'msqrt', 'mroot', 'mstyle', 'merror', 'mpadded', 'mphantom',
    'menclose', 'maction', 'munder', 'mover', 'munderover', 'mtable', 'mtr', 'mtd',
    'mlabeledtr', 'maligngroup', 'malignmark', 'mintoken', 'mglyph', 'mipartition',
]);

function decodeEntities(s) {
    return s
        .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
        .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
        .replace(/&nbsp;/g, ' ')
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&');
}

function parseHtml(html) {
    const root = { tag: 'root', attrs: {}, className: '', children: [] };
    const stack = [root];
    // Текст остаётся «сырым» (entity-ссылки декодируются точечно при извлечении):
    // KaTeX пишет LaTeX в <annotation> именно через entity-ссылки.
    const re = /<\/([a-zA-Z][\w:-]*)\s*>|<([a-zA-Z][\w:-]*)((?:"[^"]*"|'[^']*'|[^>"'])+?)(\/?)>/g;
    let m, last = 0;
    while ((m = re.exec(html)) !== null) {
        const textChunk = html.slice(last, m.index);
        if (textChunk) stack[stack.length - 1].children.push({ text: textChunk });
        last = re.lastIndex;
        if (m[1]) {
            const name = m[1].toLowerCase();
            for (let i = stack.length - 1; i > 0; i--) {
                if (stack[i].tag === name) { stack.length = i; break; }
            }
            continue;
        }
        const tag = m[2].toLowerCase();
        const attrs = parseAttrs(m[3] || '');
        const node = { tag, attrs, className: attrs.class || '', children: [] };
        stack[stack.length - 1].children.push(node);
        const selfClosing = !!m[4];
        if (!selfClosing && !VOID_TAGS.has(tag) && !(MATHML_TAGS.has(tag) && selfClosing)) {
            stack.push(node);
        }
    }
    const tail = html.slice(last);
    if (tail) stack[stack.length - 1].children.push({ text: tail });
    return root;
}

function parseAttrs(str) {
    const attrs = {};
    const re = /([\w:@.\-\[\]]+)\s*(?:=\s*("([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g;
    let m;
    while ((m = re.exec(str)) !== null) {
        const name = m[1].toLowerCase();
        const val = m[3] !== undefined ? m[3] : m[4] !== undefined ? m[4] : m[5] !== undefined ? m[5] : '';
        attrs[name] = decodeEntities(val);
    }
    return attrs;
}

// обход helpers
function walk(node, fn) {
    if (!node) return;
    fn(node);
    for (const c of node.children || []) walk(c, fn);
}

// ============================================================ 2. inline runs
// returns array of runs: {t:'text'|'omml'|'br', ...}
const FORMULA_TAGS = new Set(['annotation', 'semantics', 'math']);

function renderInline(node, ctx) {
    const runs = [];
    function visit(n, style) {
        if (n.text !== undefined) {
            const t = decodeEntities(n.text);
            if (t) runs.push({ t: 'text', v: t, ...style });
            return;
        }
        // формула?
        const cap = captureFromNode(n);
        if (cap) {
            const omml = latexToOmml(cap.latex);
            if (omml) { runs.push({ t: 'omml', xml: omml, display: cap.display }); return; }
            runs.push({ t: 'text', v: cap.latex, ...style });
            return;
        }
        // MathML-внутренности KaTeX — не выводим дважды (формула уже взята выше)
        if (FORMULA_TAGS.has(n.tag)) return;
        const tag = n.tag;
        const st = Object.assign({}, style);
        switch (tag) {
            case 'strong': case 'b': st.b = true; break;
            case 'em': case 'i': st.i = true; break;
            case 'del': case 's': case 'strike': st.del = true; break;
            case 'code': st.code = true; break;
            case 'mark': st.mark = true; break;
            case 'sup': st.sup = true; break;
            case 'sub': st.sub = true; break;
            case 'u': st.u = true; break;
            case 'br': runs.push({ t: 'br' }); return;
            case 'img': {
                const alt = (n.attrs && n.attrs.alt) || '';
                runs.push({ t: 'text', v: alt ? `[${alt}]` : '[изображение]', ...style });
                return;
            }
            case 'input': {
                const checked = n.attrs && (n.attrs.checked !== undefined);
                runs.push({ t: 'text', v: checked ? '☑ ' : '☐ ', ...style });
                return;
            }
            case 'label': { for (const c of n.children || []) visit(c, st); return; }
            case 'a': {
                st.link = (n.attrs && n.attrs.href) || '';
                break;
            }
            default: break;
        }
        for (const c of n.children || []) visit(c, st);
    }
    visit(node, {});
    return runs.filter(r => r.t !== 'text' || r.v !== '');
}

function runsToXml(runs) {
    return runs.map(r => {
        if (r.t === 'br') return '<w:r><w:br/></w:r>';
        if (r.t === 'omml') return r.xml;
        return textRun(r.v, r);
    }).join('');
}

function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function textRun(text, st = {}) {
    if (!text) return '';
    let rPr = '';
    if (st.b) rPr += '<w:b/>';
    if (st.i) rPr += '<w:i/>';
    if (st.u) rPr += '<w:u w:val="single"/>';
    if (st.del) rPr += '<w:strike/>';
    if (st.mark) rPr += '<w:highlight w:val="yellow"/>';
    if (st.code) rPr += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:shd w:val="clear" w:fill="F2F2F2"/>';
    if (st.sup) rPr += '<w:vertAlign w:val="superscript"/>';
    if (st.sub) rPr += '<w:vertAlign w:val="subscript"/>';
    if (st.link) rPr += '<w:color w:val="0563C1"/><w:u w:val="single"/>';
    const pr = rPr ? `<w:rPr>${rPr}</w:rPr>` : '';
    return `<w:r>${pr}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

// ============================================================== 3. blocks
function collectBlocks(root) {
    const blocks = [];
    const body = findFirst(root, 'body') || root;

    function handleList(node, ordered, level) {
        for (const li of directChildren(node, 'li')) {
            // содержимое li: инлайн + вложенные списки
            const inlineKids = [];
            const subLists = [];
            let taskState = null;
            for (const c of li.children || []) {
                if (c.tag === 'ul' || c.tag === 'ol') subLists.push(c);
                else if (c.tag === 'input') {
                    taskState = (c.attrs && c.attrs.checked !== undefined) ? '☑' : '☐';
                } else inlineKids.push(c);
            }
            const pseudo = { tag: 'span', attrs: {}, className: '', children: inlineKids };
            const runs = renderInline(pseudo);
            blocks.push({ type: 'list-item', runs, ordered, level, task: taskState });
            for (const sub of subLists) {
                handleList(sub, sub.tag === 'ol', level + 1);
            }
        }
    }

    function visitBlock(node) {
        if (node.text !== undefined) {
            const t = node.text.trim();
            if (t) blocks.push({ type: 'p', runs: [{ t: 'text', v: t }] });
            return;
        }
        const tag = node.tag;
        switch (tag) {
            case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'h6':
                blocks.push({ type: 'heading', level: +tag[1], runs: renderInline(node) });
                return;
            case 'p': {
                // Possible katex-display wrapper inside
                const runs = renderInline(node);
                if (runs.some(r => r.t === 'omml' && r.display)) {
                    // выносим дисплейные формулы отдельными блоками
                    const before = [], after = [], math = [];
                    for (const r of runs) {
                        if (r.t === 'omml' && r.display) math.push(r);
                        else (math.length ? after : before).push(r);
                    }
                    if (before.length) blocks.push({ type: 'p', runs: before });
                    for (const m of math) blocks.push({ type: 'p', runs: [m], displayMath: true });
                    if (after.length) blocks.push({ type: 'p', runs: after });
                } else {
                    blocks.push({ type: 'p', runs });
                }
                return;
            }
            case 'ul': case 'ol':
                handleList(node, tag === 'ol', 0);
                return;
            case 'blockquote': {
                const saved = blocks.length;
                for (const c of node.children || []) visitBlock(c);
                const newBlocks = blocks.splice(saved);
                for (const b of newBlocks) { b.quote = (b.quote || 0) + 1; blocks.push(b); }
                return;
            }
            case 'pre': {
                const codeEl = findFirst(node, 'code');
                const src = (codeEl || node);
                const lang = ((src.className || '').match(/language-([\w-]+)/) || [])[1] || '';
                let text = rawText(src);
                if (lang === 'mermaid') {
                    blocks.push({ type: 'code', text, lang: 'mermaid' });
                } else {
                    text = text.replace(/^\n+|\n+$/g, '');
                    for (const line of text.split('\n')) {
                        blocks.push({ type: 'code', text: line, lang });
                    }
                }
                return;
            }
            case 'table': {
                const rows = [];
                let isHeader = false;
                for (const tr of findAll(node, 'tr')) {
                    const cells = [];
                    let headerRow = false;
                    for (const cell of directChildren(tr, 'td').concat(directChildren(tr, 'th'))) {
                        if (cell.tag === 'th') headerRow = true;
                        cells.push(renderInline(cell));
                    }
                    if (headerRow && rows.length === 0) isHeader = true;
                    rows.push(cells);
                }
                if (rows.length) blocks.push({ type: 'table', rows, header: isHeader });
                return;
            }
            case 'hr':
                blocks.push({ type: 'hr' });
                return;
            case 'section': case 'article': case 'div': case 'main': case 'body': case 'root': case 'figure': {
                // katex-display div -> отдельный блок формулы
                const cap = captureFromNode(node);
                if (cap && (node.tag === 'div' || node.tag === 'span')) {
                    const omml = latexToOmml(cap.latex);
                    if (omml) { blocks.push({ type: 'p', runs: [{ t: 'omml', xml: omml, display: true }], displayMath: true }); return; }
                }
                const kids = node.children || [];
                const hasBlockKid = kids.some(c => c.tag && ['p','h1','h2','h3','h4','h5','h6','ul','ol','pre','table','blockquote','hr','div','section'].includes(c.tag));
                if (hasBlockKid) { for (const c of kids) visitBlock(c); }
                else if (kids.length) {
                    const runs = renderInline(node);
                    if (runs.length) blocks.push({ type: 'p', runs });
                }
                return;
            }
            default: {
                // span и прочая мелочь верхнего уровня
                const runs = renderInline(node);
                if (runs.length) blocks.push({ type: 'p', runs });
                return;
            }
        }
    }

    for (const c of body.children || []) visitBlock(c);
    return blocks;
}

function rawText(node) {
    if (node.text !== undefined) return decodeEntities(node.text);
    if (node.tag === 'annotation') return ''; // MathML-аннотации игнорируем в textual-выводе
    let s = '';
    if (node.tag === 'br') return '\n';
    for (const c of node.children || []) s += rawText(c);
    return s;
}

function directChildren(node, tag) {
    return (node.children || []).filter(c => c.tag === tag);
}
function findFirst(node, tag) {
    let found = null;
    walk(node, n => { if (!found && n.tag === tag) found = n; });
    return found;
}
function findAll(node, tag) {
    const out = [];
    walk(node, n => { if (n.tag === tag) out.push(n); });
    return out;
}

// ============================================================ 4. docx build
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

const CTRL_PR = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr>';

function attachCtrlPr(xml) {
    // pandoc вставляет ctrlPr в каждый *Pr OMML-элемента — Word это любит
    return xml
        .replace(/<m:ctrlPr\/>/g, `<m:ctrlPr>${CTRL_PR}</m:ctrlPr>`)
        .replace(/(<m:(?:f|rad|d|nary|acc|bar|sSub|sSup|sSubSup|m)f?Pr>)(?!<m:ctrlPr)/g, '$1')
        ;
}

function paraXml(block, opts = {}) {
    const pPrParts = [];
    if (opts.style) pPrParts.push(`<w:pStyle w:val="${opts.style}"/>`);
    if (block.quote) pPrParts.push('<w:pStyle w:val="Quote"/>');
    if (opts.listId !== undefined) {
        pPrParts.push(`<w:numPr><w:ilvl w:val="${opts.level || 0}"/><w:numId w:val="${opts.listId}"/></w:numPr>`);
    }
    if (block.displayMath) pPrParts.push('<w:jc w:val="center"/>');
    const pPr = pPrParts.length ? `<w:pPr>${pPrParts.join('')}</w:pPr>` : '';
    let content = runsToXml(block.runs || []).replace(/<m:ctrlPr\/>/g, `<m:ctrlPr>${CTRL_PR}</m:ctrlPr>`);
    if (block.task !== null && block.task !== undefined) {
        content = textRun(block.task === '☑' ? '☑ ' : '☐ ') + content;
    }
    return `<w:p>${pPr}${content}</w:p>`;
}

function tableXml(table) {
    const gridCols = Math.max(...table.rows.map(r => r.length));
    const width = Math.floor(9026 / gridCols);
    let xml = '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>' +
        '<w:tblBorders>' +
        ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
            .map(s => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>`).join('') +
        '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>';
    xml += '<w:tblGrid>' + Array(gridCols).fill(`<w:gridCol w:w="${width}"/>`).join('') + '</w:tblGrid>';
    table.rows.forEach((row, ri) => {
        const isHdr = table.header && ri === 0;
        xml += '<w:tr>' + (isHdr ? '<w:trPr><w:tblHeader/></w:trPr>' : '');
        for (let ci = 0; ci < gridCols; ci++) {
            const cellRuns = row[ci] || [];
            const shd = isHdr ? '<w:shd w:val="clear" w:fill="E7E6E6"/>' : '';
            const runs = cellRuns.map(r => {
                if (r.t === 'br') return '<w:r><w:br/></w:r>';
                if (r.t === 'omml') return r.xml.replace(/<m:ctrlPr\/>/g, `<m:ctrlPr>${CTRL_PR}</m:ctrlPr>`);
                return textRun(isHdr ? r.v : r.v, isHdr ? Object.assign({ b: true }, r) : r);
            }).join('');
            xml += `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${shd}</w:tcPr><w:p>${runs}</w:p></w:tc>`;
        }
        xml += '</w:tr>';
    });
    xml += '</w:tbl><w:p/>'; // пустой абзац после таблицы (требование Word при таблице в конце)
    return xml;
}

function buildDocumentXml(blocks) {
    let body = '';
    let numbering = null; // собранная нумерация
    const numDefs = [];
    let curList = null; // {ordered, id}

    function ensureNumbering(ordered) {
        if (curList && curList.ordered === ordered) return curList.id;
        const id = numDefs.length + 1;
        numDefs.push({ id, ordered });
        curList = { ordered, id };
        return id;
    }

    for (const b of blocks) {
        switch (b.type) {
            case 'heading':
                body += paraXml(b, { style: `Heading${b.level}` });
                break;
            case 'p':
                if (b.displayMath) {
                    // <m:oMathPara> — дисплейная формула отдельным абзацем
                    const om = b.runs.find(r => r.t === 'omml');
                    if (om) {
                        const inner = om.xml.replace(/<m:ctrlPr\/>/g, `<m:ctrlPr>${CTRL_PR}</m:ctrlPr>`);
                        body += `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>${inner}</m:oMathPara></w:p>`;
                    } else body += paraXml(b);
                } else {
                    curList = null;
                    body += paraXml(b, b.quote ? { style: 'Quote' } : {});
                }
                break;
            case 'list-item': {
                const listId = ensureNumbering(b.ordered);
                body += paraXml(b, { listId, level: b.level });
                break;
            }
            case 'code':
                curList = null;
                body += `<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr>` +
                    (b.text ? textRun(b.text, { code: true }) : '') + '</w:p>';
                break;
            case 'hr':
                curList = null;
                body += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="999999"/></w:pBdr></w:pPr></w:p>';
                break;
            case 'table':
                curList = null;
                body += tableXml(b);
                break;
            default: break;
        }
    }

    const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>';

    const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<w:document ${NS}><w:body>${body}${sectPr}</w:body></w:document>`;

    return { document, numDefs };
}

function buildNumberingXml(numDefs) {
    if (!numDefs.length) return '';
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering ${NS}>`;
    for (const d of numDefs) {
        const abstractId = d.id;
        xml += `<w:abstractNum w:abstractNumId="${abstractId}">` +
            `<w:multiLevelType w:val="hybridMultilevel"/>`;
        for (let lvl = 0; lvl < 9; lvl++) {
            const fmt = d.ordered ? 'decimal' : 'bullet';
            const txt = d.ordered ? `%${lvl + 1}.` : (lvl % 3 === 0 ? '•' : lvl % 3 === 1 ? '◦' : '▪');
            xml += `<w:lvl w:ilvl="${lvl}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>` +
                `<w:lvlText w:val="${txt}"/><w:lvlJc w:val="left"/>` +
                `<w:pPr><w:ind w:left="${720 * (lvl + 1)}" w:hanging="360"/></w:pPr>` +
                (d.ordered ? '' : '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>') +
                `</w:lvl>`;
        }
        xml += `</w:abstractNum>`;
    }
    for (const d of numDefs) {
        xml += `<w:num w:numId="${d.id}"><w:abstractNumId w:val="${d.id}"/></w:num>`;
    }
    xml += '</w:numbering>';
    return xml;
}

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="AAAAAA"/></w:pBdr><w:shd w:val="clear" w:fill="F7F7F7"/></w:pPr><w:rPr><w:i/><w:color w:val="444444"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:shd w:val="clear" w:fill="F2F2F2"/><w:spacing w:after="20" w:line="240" w:lineRule="auto"/><w:ind w:left="240"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/></w:rPr></w:style>
${[1, 2, 3, 4, 5, 6].map(i => `<w:style w:type="paragraph" w:styleId="Heading${i}"><w:name w:val="heading ${i}"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="${i - 1}"/><w:spacing w:before="${(7 - i) * 80}" w:after="120"/></w:pPr><w:rPr><w:b/><w:color w:val="${['1F4E79', '2E5B8F', '2E5B8F', '404040', '404040', '595959'][i - 1]}"/><w:sz w:val="${(8 - i) * 6 + 10}"/><w:szCs w:val="${(8 - i) * 6 + 10}"/></w:rPr></w:style>`).join('\n')}
</w:styles>`;

const SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings ${NS}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

const CONTENT_TYPES = (withNumbering) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
${withNumbering ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : ''}
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = (withNumbering) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
${withNumbering ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' : ''}
</Relationships>`;

// ---------------------------------------------------------------- zip writer
function crc32(buf) {
    let c, table = crc32.table;
    if (!table) {
        table = crc32.table = new Int32Array(256);
        for (let n = 0; n < 256; n++) {
            c = n;
            for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
            table[n] = c;
        }
    }
    c = -1;
    for (let i = 0; i < buf.length; i++) c = table[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ -1) >>> 0;
}

function makeZip(entries) {
    // entries: [{name, data(Buffer)}]
    const chunks = [];
    const central = [];
    let offset = 0;
    for (const e of entries) {
        const nameBuf = Buffer.from(e.name, 'utf8');
        const data = e.data;
        const crc = crc32(data);
        const local = Buffer.alloc(30);
        local.writeUInt32LE(0x04034b50, 0);
        local.writeUInt16LE(20, 4);          // version needed
        local.writeUInt16LE(0, 6);           // flags
        local.writeUInt16LE(0, 8);           // method: store (OOXML так можно)
        local.writeUInt16LE(0, 10);          // time
        local.writeUInt16LE(0x2821, 12);     // date (2026-01-01-ish)
        local.writeUInt32LE(crc, 14);
        local.writeUInt32LE(data.length, 18);
        local.writeUInt32LE(data.length, 22);
        local.writeUInt16LE(nameBuf.length, 26);
        local.writeUInt16LE(0, 28);
        chunks.push(local, nameBuf, data);

        const cen = Buffer.alloc(46);
        cen.writeUInt32LE(0x02014b50, 0);
        cen.writeUInt16LE(20, 4);
        cen.writeUInt16LE(20, 6);
        cen.writeUInt16LE(0, 8);
        cen.writeUInt16LE(0, 10);
        cen.writeUInt16LE(0, 12);
        cen.writeUInt16LE(0x2821, 14);
        cen.writeUInt32LE(crc, 16);
        cen.writeUInt32LE(data.length, 20);
        cen.writeUInt32LE(data.length, 24);
        cen.writeUInt16LE(nameBuf.length, 28);
        cen.writeUInt32LE(offset, 42);
        central.push(cen, nameBuf);
        offset += local.length + nameBuf.length + data.length;
    }
    const centralBuf = Buffer.concat(central);
    const end = Buffer.alloc(22);
    end.writeUInt32LE(0x06054b50, 0);
    end.writeUInt16LE(entries.length, 8);
    end.writeUInt16LE(entries.length, 10);
    end.writeUInt32LE(centralBuf.length, 12);
    end.writeUInt32LE(offset, 16);
    return Buffer.concat([...chunks, centralBuf, end]);
}

// ================================================================== API
// html — результат render() того же markdown-it, что и в ридере
function convertHtmlToDocxBuffer(html) {
    const dom = parseHtml(html);
    const blocks = collectBlocks(dom);
    const { document, numDefs } = buildDocumentXml(blocks);
    const entries = [
        { name: '[Content_Types].xml', data: Buffer.from(CONTENT_TYPES(numDefs.length > 0), 'utf8') },
        { name: '_rels/.rels', data: Buffer.from(ROOT_RELS, 'utf8') },
        { name: 'word/_rels/document.xml.rels', data: Buffer.from(DOC_RELS(numDefs.length > 0), 'utf8') },
        { name: 'word/document.xml', data: Buffer.from(document, 'utf8') },
        { name: 'word/styles.xml', data: Buffer.from(STYLES_XML, 'utf8') },
        { name: 'word/settings.xml', data: Buffer.from(SETTINGS_XML, 'utf8') },
    ];
    if (numDefs.length > 0) {
        entries.push({ name: 'word/numbering.xml', data: Buffer.from(buildNumberingXml(numDefs), 'utf8') });
    }
    return makeZip(entries);
}

// прямой путь: markdown + готовый html (рендер делает вызывающий — тот же markdown-it)
function exportMarkdownToDocx({ markdown, html }) {
    const srcHtml = html || (() => {
        // запасной вариант без markdown-it: не должен использоваться в app
        throw new Error('html обязателен: нужен render того же markdown-it, что и в ридере');
    })();
    return convertHtmlToDocxBuffer(srcHtml);
}

module.exports = {
    parseHtml,
    collectBlocks,
    convertHtmlToDocxBuffer,
    exportMarkdownToDocx,
    makeZip,
};
