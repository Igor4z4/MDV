/*
 * md_ast.js — этап 1 конвейера md2docx: Markdown → JSON-AST.
 *
 * Использует ТЕМ ЖЕ стеком, что и ридер (web/lib): markdown-it + texmath + KaTeX +
 * footnote/mark/sub/sup/task-lists/multimd-table. Поэтому формулы в $...$ разбираются
 * ровно так же, как на экране. Для каждого math-токена дополнительно компилируем
 * LaTeX → MathML (KaTeX output:'mathml'), который затем конвертируется в OMML
 * (нативные формулы Word). Это главное улучшение относительно pandoc: там $...$
 * часто теряются или превращаются в текст, здесь формула становится редактируемым
 * объектом Equation внутри .docx.
 *
 * Запуск: node md_ast.js input.md > ast.json
 */
'use strict';

const fs = require('fs');
const path = require('path');

function die(msg) {
    process.stderr.write('MD_AST_ERROR: ' + msg + '\n');
    process.exit(2);
}

// --- загрузка UMD-бандлов из web/lib (те же файлы, что использует ридер) ----
function loadUMD(file) {
    if (!fs.existsSync(file)) die('не найден модуль: ' + file);
    const module_ = { exports: {} };
    // sandbox без доступа к глобальному window/document
    const fn = new Function('window', 'self', 'globalThis', 'module', 'exports',
        fs.readFileSync(file, 'utf8'));
    const g = {};
    fn(g, g, g, module_, module_.exports);
    return module_.exports.default || module_.exports;
}

const LIB = path.join(__dirname, '..', 'web', 'lib');

const MarkdownIt = loadUMD(path.join(LIB, 'markdown-it.min.js'));
const texmath = loadUMD(path.join(LIB, 'texmath.js'));
const katex = loadUMD(path.join(LIB, 'katex.min.js'));

let md;
try {
    md = new MarkdownIt({ html: false, linkify: true, typographer: false });
} catch (e) { die('MarkdownIt init: ' + e.message); }

function tryUse(name, ...args) {
    try {
        md.use(loadUMD(path.join(LIB, name)), ...args);
        return true;
    } catch (e) {
        return false;
    }
}

// Формулы: texmath с dollars+brackets, вывод MathML (для последующего OMML)
const texmathOk = (() => {
    try {
        md.use(texmath, {
            engine: katex,
            delimiters: ['dollars', 'brackets'],
            katexOptions: { macros: {}, output: 'mathml', throwOnError: false },
        });
        return true;
    } catch (e) { return false; }
})();

if (!texmathOk) {
    // Резервный собственный dollar-парсер — всё равно лучше pandoc-дефолтов:
    md.inline.ruler.before('escape', 'math_inline', function (state, silent) {
        if (state.src[state.pos] !== '$') return false;
        let end = -1;
        for (let i = state.pos + 1; i < state.posMax; i++) {
            if (state.src[i] === '\\') { i++; continue; }
            if (state.src[i] === '$') { end = i; break; }
            if (state.src[i] === '\n') return false;
        }
        if (end < 0) return false;
        if (!silent) {
            const token = state.push('math_inline', 'math', 0);
            token.content = state.src.slice(state.pos + 1, end);
        }
        state.pos = end + 1;
        return true;
    });
    md.block.ruler.before('fence', 'math_block', function (state, startLine, endLine, silent) {
        const startPos = state.bMarks[startLine] + state.tShift[startLine];
        const lineTxt = state.src.slice(startPos, state.eMarks[startLine]).trim();
        if (!lineTxt.startsWith('$$')) return false;
        if (silent) return true;
        let content = lineTxt.slice(2), lastLine = startLine;
        if (content.trim().endsWith('$$') && content.trim().length >= 2) {
            content = content.replace(/\$\$\s*$/, '');
        } else {
            const parts = [content];
            let i = startLine + 1;
            for (; i < endLine; i++) {
                const s = state.src.slice(state.bMarks[i] + state.tShift[i], state.eMarks[i]);
                if (s.trim().endsWith('$$')) { parts.push(s.trim().replace(/\$\$\s*$/, '')); lastLine = i; break; }
                parts.push(s); lastLine = i;
            }
            content = parts.join('\n');
        }
        const token = state.push('math_block', 'math', 0);
        token.block = true; token.content = content; token.map = [startLine, lastLine + 1];
        state.line = lastLine + 1;
        return true;
    });
}

tryUse('markdown-it-footnote.min.js');
tryUse('markdown-it-mark.min.js');
tryUse('markdown-it-sub.min.js');
tryUse('markdown-it-sup.min.js');
tryUse('markdown-it-task-lists.min.js', { enabled: true });
tryUse('markdown-it-multimd-table.min.js', { multiline: true, rowspan: true, headerless: true });
tryUse('markdown-it-emoji.min.js');

// --- парсим ---------------------------------------------------------------
const inputFile = process.argv[2];
if (!inputFile) die('не указан входной .md файл');
const src = fs.readFileSync(inputFile, 'utf8');
const tokens = md.parse(src, {});

// --- токен-лист → дерево ----------------------------------------------------
function tokNode(t) {
    return {
        type: t.type, tag: t.tag, content: t.content, markup: t.markup,
        info: t.info, attrs: t.attrs || [], map: t.map || null, children: null,
    };
}

function buildTree(toks) {
    const root = [];
    const stack = [{ kids: root }];
    for (const t of toks) {
        if (t.type === 'inline') {
            // содержимое inline-токена — тоже дерево
            const node = tokNode(t);
            node.children = buildInline(t.children || []);
            stack[stack.length - 1].kids.push(node);
            continue;
        }
        if (t.nesting === 1) {
            const node = tokNode(t);
            node.children = [];
            stack[stack.length - 1].kids.push(node);
            stack.push({ kids: node.children });
        } else if (t.nesting === -1) {
            if (stack.length > 1) stack.pop();
        } else {
            stack[stack.length - 1].kids.push(tokNode(t));
        }
    }
    return root;
}

function buildInline(toks) {
    const out = [];
    const stack = [{ kids: out }];
    for (const t of toks) {
        if (t.nesting === 1) {
            const node = tokNode(t);
            node.children = [];
            stack[stack.length - 1].kids.push(node);
            stack.push({ kids: node.children });
        } else if (t.nesting === -1) {
            if (stack.length > 1) stack.pop();
        } else {
            stack[stack.length - 1].kids.push(tokNode(t));
        }
    }
    return out;
}

const tree = buildTree(tokens);

// --- для каждого math-узла компилируем LaTeX → MathML ----------------------
function renderMathML(tex, displayMode) {
    try {
        const htmlOut = katex.renderToString(tex, {
            output: 'mathml', throwOnError: false, displayMode: !!displayMode,
        });
        const m = htmlOut.match(/<math[\s\S]*?<\/math>/);
        return m ? m[0] : null;
    } catch (e) { return null; }
}

(function annotate(nodes) {
    for (const n of nodes) {
        if (n.type === 'math_inline' || n.type === 'math_block') {
            n.tex = n.content;
            n.mathml = renderMathML(n.content, n.type === 'math_block');
        }
        if (n.children) annotate(n.children);
    }
})(tree);

process.stdout.write(JSON.stringify(tree));
