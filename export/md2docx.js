/*
 * md2docx.js — этап 2 конвейера «Сохранить как .docx» внутри приложения:
 * JSON-AST (см. md_ast.js) → файл .docx. Реализован на пакете `docx`,
 * поэтому не требует Python/lxml — работает прямо в процессе Electron.
 *
 * Формулы: LaTeX из AST превращается в OMML (<m:oMath>) через KaTeX-parser?
 * Нет — здесь собственный лёгкий конвертер подмножества LaTeX → OMML для
 * инлайновых формул; блочные и сложные формулы используют MathML→OMML
 * преобразование того же подмножества. Если что-то не поддерживается,
 * формула сохраняется как редактируемое текстовое представление LaTeX.
 *
 * API: const { exportDocx } = require('./md2docx');
 *      await exportDocx(astTree, outputPath, { baseDir });
 */
'use strict';

const fs = require('fs');
const path = require('path');
const docx = require('docx');
const {
    Document, Packer, Paragraph, TextRun, HeadingLevel, ExternalHyperlink,
    Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
    AlignmentType, LevelFormat, convertMillimetersToTwip, ImageRun, Math: OMath,
    MathRun, UnderlineType, TabStopType, PageNumber, Footer,
} = docx;

// ---------------------------------------------------------------------------
// LaTeX → OMML (подмножество, покрывающее типовые формулы в Markdown)
// ---------------------------------------------------------------------------
const M = 'm:';
function el(tag, attrs, children) {
    return { tag: M + tag, attrs: attrs || {}, children: children || [] };
}
function txt(s) { return { text: s }; }

const SYMBOLS = {
    alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε',
    zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ',
    lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ',
    tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'ϕ', chi: 'χ', psi: 'ψ', omega: 'ω',
    Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π',
    Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
    infty: '∞', partial: '∂', nabla: '∇', forall: '∀', exists: '∃',
    times: '×', div: '÷', pm: '±', mp: '∓', cdot: '⋅', ast: '∗',
    leq: '≤', le: '≤', geq: '≥', ge: '≥', neq: '≠', ne: '≠', approx: '≈',
    equiv: '≡', sim: '∼', propto: '∝', ll: '≪', gg: '≫',
    subset: '⊂', supset: '⊃', subseteq: '⊆', supseteq: '⊇',
    in: '∈', notin: '∉', ni: '∋', cup: '∪', cap: '∩', emptyset: '∅',
    rightarrow: '→', to: '→', leftarrow: '←', Rightarrow: '⇒',
    Leftarrow: '⇐', leftrightarrow: '↔', mapsto: '↦',
    ldots: '…', dots: '…', cdots: '⋯', vdots: '⋮', ddots: '⋱',
    circ: '∘', bullet: '•', star: '★', dagger: '†',
    angle: '∠', perp: '⊥', parallel: '∥', triangle: '△',
    prime: '′', degree: '°', hbar: 'ℏ', ell: 'ℓ', Re: 'ℜ', Im: 'ℑ',
    aleph: 'ℵ', wp: '℘', therefore: '∴', because: '∵',
    langle: '⟨', rangle: '⟩', lceil: '⌈', rceil: '⌉', lfloor: '⌊', rfloor: '⌋',
    sum: '∑', prod: '∏', int: '∫', oint: '∮', bigcup: '⋃', bigcap: '⋂',
};

const FUNCS = new Set(['sin','cos','tan','cot','sec','csc','sinh','cosh','tanh',
    'log','ln','exp','lim','max','min','det','dim','gcd','arg','deg','Pr']);

function tokenizeTex(src) {
    const toks = [];
    let i = 0;
    while (i < src.length) {
        const c = src[i];
        if (/\s/.test(c)) { i++; continue; }
        if (c === '\\') {
            const m = /^\\([a-zA-Z]+|.)/.exec(src.slice(i));
            toks.push({ t: 'cmd', v: m[1] });
            i += m[0].length;
        } else if (/[{}&^_]/.test(c)) {
            toks.push({ t: c }); i++;
        } else if (c === '[') { toks.push({ t: 'open', v: '[' }); i++; }
        else if (c === ']') { toks.push({ t: 'close', v: ']' }); i++; }
        else if (c === '(') { toks.push({ t: 'open', v: '(' }); i++; }
        else if (c === ')') { toks.push({ t: 'close', v: ')' }); i++; }
        else { toks.push({ t: 'ch', v: c }); i++; }
    }
    return toks;
}

class TexParser {
    constructor(toks) { this.toks = toks; this.i = 0; }
    peek() { return this.toks[this.i]; }
    next() { return this.toks[this.i++]; }
    atEnd() { return this.i >= this.toks.length; }

    // группа: один токен или {...} / [..] аргумент
    group() {
        const tk = this.peek();
        if (!tk) return [];
        if (tk.t === '{') {
            this.next();
            const nodes = this.list('}');
            if (this.peek() && this.peek().t === '}') this.next();
            return nodes;
        }
        if (tk.t === 'ch' || tk.t === 'cmd') { this.next(); return [this.atom(tk)]; }
        return [];
    }

    list(stop) {
        const out = [];
        while (!this.atEnd()) {
            const tk = this.peek();
            if (stop && tk.t === stop) break;
            if (tk.t === '&') { this.next(); continue; } // разделитель колонок обрабатывается выше
            const node = this.item();
            if (node) out.push(node);
        }
        return out;
    }

    item() {
        const tk = this.next();
        if (!tk) return null;
        switch (tk.t) {
            case 'ch': return { kind: 'run', text: tk.v };
            case 'open': {
                const inner = this.list('close');
                if (this.peek() && this.peek().t === 'close') this.next();
                return { kind: 'delim', left: tk.v, right: this.matchClose(tk.v), inner };
            }
            case 'close': return null;
            case '^': {
                const sup = this.group();
                return { kind: 'script', which: 'sup', sup };
            }
            case '_': {
                const sub = this.group();
                return { kind: 'script', which: 'sub', sub };
            }
            case 'cmd': return this.command(tk.v);
            default: return null;
        }
    }

    matchClose(open) {
        return { '[': ']', '(': ')' }[open] || '';
    }

    atom(tk) {
        if (tk.t === 'ch') return { kind: 'run', text: tk.v };
        if (tk.t === 'cmd') return this.command(tk.v);
        return null;
    }

    command(name) {
        switch (name) {
            case 'frac': case 'dfrac': case 'tfrac': {
                const num = this.group(), den = this.group();
                return { kind: 'frac', num, den };
            }
            case 'sqrt': {
                let deg = null;
                if (this.peek() && this.peek().t === 'open' && this.peek().v === '[') {
                    this.next(); deg = this.list('close');
                    if (this.peek() && this.peek().t === 'close') this.next();
                }
                const rad = this.group();
                return { kind: 'sqrt', deg, rad };
            }
            case 'sum': case 'prod': case 'int': case 'oint': case 'bigcup': case 'bigcap':
                return { kind: 'nary', sym: SYMBOLS[name] || name };
            case 'text': case 'mathrm': case 'operatorname': case 'mbox': {
                const g = this.group();
                return { kind: 'run', text: g.map(nodesToPlain).join(''), plain: true };
            }
            case 'mathbf': case 'bm': case 'boldsymbol': {
                const g = this.group();
                return { kind: 'styled', style: 'bold', inner: g };
            }
            case 'mathbb': {
                const g = this.group();
                return { kind: 'run', text: g.map(nodesToPlain).join(''), doubleStruck: true };
            }
            case 'hat': case 'widehat': return { kind: 'accent', chr: '\u0302', inner: this.group() };
            case 'bar': case 'overline': return { kind: 'bar', inner: this.group() };
            case 'vec': case 'overrightarrow': return { kind: 'accent', chr: '\u20D7', inner: this.group() };
            case 'dot': return { kind: 'accent', chr: '\u0307', inner: this.group() };
            case 'tilde': case 'widetilde': return { kind: 'accent', chr: '\u0303', inner: this.group() };
            case 'left': {
                const lt = this.next(); const lc = lt ? (lt.v || '') : '';
                const inner = this.list('right');
                if (this.peek() && this.peek().t === 'cmd' && this.peek().v === 'right') this.next();
                const rt = this.next(); const rc = rt ? (rt.v || '') : '';
                return { kind: 'delim', left: cleanDelim(lc), right: cleanDelim(rc), inner };
            }
            case 'begin': {
                const envName = this.envName();
                return { kind: 'env', name: envName, rows: this.envRows(envName) };
            }
            case 'end': { this.envName(); return null; }
            case '\\': return { kind: 'br' };
            default:
                if (SYMBOLS[name]) return { kind: 'sym', text: SYMBOLS[name] };
                if (FUNCS.has(name)) return { kind: 'run', text: name, func: true };
                // неизвестная команда — просто текст без обратного слэша
                return { kind: 'run', text: name, plain: true };
        }
    }

    envName() {
        const tk = this.peek();
        if (tk && tk.t === '{') {
            this.next();
            let s = '';
            while (!this.atEnd() && this.peek().t !== '}') s += (this.next().v || '');
            if (this.peek() && this.peek().t === '}') this.next();
            return s;
        }
        return '';
    }

    // собирает строки матриц/систем до \end{...}
    envRows() {
        const rows = [[]];
        while (!this.atEnd()) {
            const tk = this.peek();
            if (tk.t === 'cmd' && tk.v === 'end') break;
            if (tk.t === '&') { this.next(); continue; }
            if (tk.t === '\\' ) { this.next(); rows.push([]); continue; }
            if (tk.t === 'cmd' && ['\\\\'].includes(tk.v)) { this.next(); rows.push([]); continue; }
            const node = this.item();
            if (node) rows[rows.length - 1].push(node);
        }
        return rows;
    }
}

function cleanDelim(d) {
    if (!d || d === '.') return '';
    return ({ '\\{': '{', '\\}': '}', '\\langle': '⟨', '\\rangle': '⟩',
              '\\lfloor': '⌊', '\\rfloor': '⌋', '\\lceil': '⌈', '\\rceil': '⌉' }[d]) || d;
}

function nodesToPlain(nodes) {
    return nodes.map(n => {
        if (!n) return '';
        if (n.kind === 'run' || n.kind === 'sym') return n.text;
        if (n.kind === 'frac') return nodesToPlain(n.num) + '/' + nodesToPlain(n.den);
        if (n.kind === 'sqrt') return '√(' + nodesToPlain(n.rad) + ')';
        if (n.kind === 'delim') return n.left + nodesToPlain(n.inner) + n.right;
        if (n.kind === 'script') return (n.sub ? '_' + nodesToPlain(n.sub) : '') + (n.sup ? '^' + nodesToPlain(n.sup) : '');
        if (n.kind === 'styled') return nodesToPlain(n.inner);
        if (n.kind === 'accent') return nodesToPlain(n.inner);
        if (n.kind === 'bar') return nodesToPlain(n.inner);
        if (n.kind === 'nary') return n.sym;
        if (n.kind === 'env') return n.rows.map(r => nodesToPlain(r)).join(' & ');
        return '';
    }).join('');
}

// --- OMML XML ---------------------------------------------------------------
function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function ommlRun(text, opts) {
    opts = opts || {};
    let st = '';
    const sty = opts.sty || (opts.plain ? 'p' : (/[a-zA-Z]/.test(text) && text.length === 1 && !opts.func ? 'i' : 'p'));
    st = `<m:r><m:rPr>${sty === 'p' ? '' : `<m:sty m:val="${opts.bold ? 'bi' : sty}"/>`}</m:rPr><w:t xml:space="preserve">${esc(text)}</w:t></m:r>`;
    if (sty === 'p' && !opts.bold) st = `<m:r><w:t xml:space="preserve">${esc(text)}</w:t></m:r>`;
    return st;
}

function seq(childrenXml) { return `<m:e>${childrenXml.join('')}</m:e>`; }

function nodeToOmml(n) {
    if (!n) return '';
    switch (n.kind) {
        case 'run':
            return ommlRun(n.text, { plain: n.plain, bold: n.doubleStruck ? false : undefined, func: n.func, sty: n.func || n.plain ? 'p' : undefined });
        case 'sym':
            return ommlRun(n.text, { plain: true });
        case 'styled':
            return n.inner.map(nodeToOmml).map(x => x.replace('<m:r>', '<m:r><m:rPr><m:sty m:val="b"/></m:rPr>').replace(/<m:rPr\/>/, '')) .join('') ||
                ommlRun(nodesToPlain(n.inner), { bold: true });
        case 'frac':
            return `<m:f>${seq([n.num.map(nodeToOmml)])}${seq([n.den.map(nodeToOmml)])}</m:f>`;
        case 'sqrt':
            if (n.deg) return `<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg/>${seq([n.deg.map(nodeToOmml)])}${seq([n.rad.map(nodeToOmml)])}</m:rad>`;
            return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${seq([n.rad.map(nodeToOmml)])}</m:rad>`;
        case 'delim': {
            const l = n.left || '', r = n.right || '';
            const pr = (l === '' && r === '') ? '' :
                `<m:dPr>${l !== '(' || r !== ')' ? `<m:begChr m:val="${esc(l)}"/><m:endChr m:val="${esc(r)}"/>` : ''}</m:dPr>`;
            return `<m:d>${pr}${seq([n.inner.map(nodeToOmml)])}</m:d>`;
        }
        case 'nary': {
            let sub = '', sup = '';
            // прицепившиеся ^/_ забираем из последовательности на уровне вызова — здесь просто символ
            return `<m:nary><m:naryPr><m:chr m:val="${esc(n.sym)}"/><m:limLoc m:val="undOvr"/></m:naryPr><m:sub/><m:sup/><m:e/></m:nary>`;
        }
        case 'accent':
            return `<m:acc><m:accPr><m:chr m:val="${esc(n.chr)}"/></m:accPr>${seq([n.inner.map(nodeToOmml)])}</m:acc>`;
        case 'bar':
            return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>${seq([n.inner.map(nodeToOmml)])}</m:bar>`;
        case 'env': {
            const trs = n.rows.map(cells =>
                `<m:tr>${cells.map(c => `<m:tc>${seq([c.map(nodeToOmml)])}</m:tc>`).join('')}</m:tr>`).join('');
            return `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="${Math.max(1, n.rows.reduce((a, r) => Math.max(a, r.length), 1))}"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>${trs}</m:m>`;
        }
        case 'br': return '';
        default: return '';
    }
}

// склейка последовательности узлов с учётом ^/_ скриптов после базового элемента
function nodesToOmml(nodes) {
    const out = [];
    for (let i = 0; i < nodes.length; i++) {
        const n = nodes[i];
        if (n && n.kind === 'script') {
            // присоединяем к предыдущему элементу как sSub/sSup/sSubSup
            const prev = out.pop() || '<m:e></m:e>';
            const base = `<m:e>${prev}</m:e>`;
            const nxt = nodes[i + 1];
            let sub = n.sub ? nodesToOmml(n.sub) : null;
            let sup = n.sup ? nodesToOmml(n.sup) : null;
            if (!sub && nxt && nxt.kind === 'script' && (nxt.sup && !n.sub)) {
                sup = nodesToOmml(nxt.sup); i++;
            } else if (!sup && nxt && nxt.kind === 'script' && (nxt.sub && !n.sup)) {
                sub = nodesToOmml(nxt.sub); i++;
            }
            if (sub && sup) out.push(`<m:sSubSup>${base}<m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`);
            else if (sub) out.push(`<m:sSub>${base}<m:sub>${sub}</m:sub></m:sSub>`);
            else out.push(`<m:sSup>${base}<m:sup>${nodesToOmml(n.sup || [])}</m:sup></m:sSup>`);
            continue;
        }
        if (n && n.kind === 'nary') {
            // у оператора лимиты: следующие _/^
            let sub = null, sup = null;
            while (nodes[i + 1] && nodes[i + 1].kind === 'script') {
                const s = nodes[++i];
                if (s.sub && sub === null) sub = nodesToOmml(s.sub);
                if (s.sup && sup === null) sup = nodesToOmml(s.sup);
            }
            out.push(`<m:nary><m:naryPr><m:chr m:val="${esc(n.sym)}"/><m:limLoc m:val="undOvr"/></m:naryPr>` +
                `<m:sub>${sub || ''}</m:sub><m:sup>${sup || ''}</m:sup><m:e/></m:nary>`);
            continue;
        }
        const xml = nodeToOmml(n);
        if (xml) out.push(xml);
    }
    return out.join('');
}

function texToOmml(tex, display) {
    try {
        const p = new TexParser(tokenizeTex(tex));
        const body = nodesToOmml(p.list(null));
        if (!body) throw new Error('empty');
        const wrap = display ? 'oMathPara' : 'oMath';
        if (display) {
            return `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr><m:oMath>${body}</m:oMath></m:oMathPara>`;
        }
        return `<m:oMath>${body}</m:oMath>`;
    } catch (e) {
        // запасной вариант: LaTeX как текст (всё ещё редактируется в Word)
        return `<m:oMath>${ommlRun(tex, { plain: true })}</m:oMath>`;
    }
}

// ---------------------------------------------------------------------------
// Построение документа docx (пакет `docx`) с внедрением сырого OMML
// ---------------------------------------------------------------------------
function runsFromInline(nodes, fmt) {
    // возвращает массив "частей": {type:'run'|'raw', ...} где raw — OMML-строка
    const parts = [];
    const walk = (list) => {
        for (const n of list) {
            const f = Object.assign({}, fmt, inlineFmt(n));
            switch (n.type) {
                case 'text': if (n.content) parts.push({ type: 'run', text: n.content, f }); break;
                case 'code_inline': parts.push({ type: 'run', text: n.content, f: Object.assign({}, f, { code: true }) }); break;
                case 'softbreak': parts.push({ type: 'run', text: ' ', f }); break;
                case 'hardbreak': parts.push({ type: 'run', text: '\n', f }); break;
                case 'math_inline': parts.push({ type: 'raw', xml: texToOmml(n.tex || n.content, false) }); break;
                case 'strong': case 'em': case 's': case 'mark': case 'sub': case 'sup':
                case 'link_open': case 'footnote_ref':
                    if (n.children) walk(n.children); break;
                case 'image': parts.push({ type: 'run', text: `[${n.content || (n.attrs && n.attrs[0] && n.attrs[0][1]) || 'изображение'}]`, f: Object.assign({}, f, { italic: true }) }); break;
                default:
                    if (n.children) walk(n.children);
                    else if (n.content) parts.push({ type: 'run', text: n.content, f });
            }
        }
    };
    walk(nodes || []);
    return parts;
}

function inlineFmt(n) {
    switch (n.type) {
        case 'strong': return { bold: true };
        case 'em': return { italics: true };
        case 's': return { strike: true };
        case 'mark': return { highlight: 'yellow' };
        case 'sub': return { subScript: true };
        case 'sup': return { superScript: true };
        default: return {};
    }
}

function makeTextRuns(parts) {
    const out = [];
    for (const p of parts) {
        if (p.type === 'run') {
            const lines = String(p.text).split('\n');
            lines.forEach((line, idx) => {
                if (idx > 0) out.push(new TextRun({ break: 1 }));
                out.push(new TextRun(Object.assign({ text: line }, p.f)));
            });
        }
    }
    return out;
}

// Собирает childElements параграфа: TextRun'ы + сырой OMML (как Element без валидации)
function buildParagraphChildren(parts) {
    const children = [];
    let runBuf = [];
    const flush = () => {
        if (runBuf.length) { children.push(...runBuf); runBuf = []; }
    };
    for (const p of parts) {
        if (p.type === 'raw') {
            flush();
            children.push(rawXmlFragment(p.xml));
        } else {
            runBuf.push(...makeTextRuns([p]));
        }
    }
    flush();
    return children;
}

// Обёртка над внутренним механизмом docx: вставляем произвольный OOXML-фрагмент.
// Пакет docx сериализует элементы через key/values; используем "фальшивый"
// элемент, который сам пишет XML (поддерживается XmlComponent-протоколом).
function rawXmlFragment(xml) {
    return {
        // docx v8 вызывает .prepForXml(context) у каждого child-элемента
        rootKey: '__raw__',
        prepForXml() { return { __raw__: xml }; },
    };
}

class Builder {
    constructor(opts) {
        this.baseDir = (opts && opts.baseDir) || '.';
        this.children = [];   // секции Document
        this.footnotesSeen = 0;
    }

    build(tree) {
        this.blocks(tree, (el) => this.children.push(el));
        const doc = new Document({
            creator: 'MDV Markdown Viewer',
            styles: {
                default: {
                    document: { run: { font: 'Calibri', size: 22 } },
                    paragraph: { run: { font: 'Calibri', size: 22 } },
                },
            },
            numbering: {
                config: [
                    { levels: [
                        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
                          style: { paragraph: { indent: { left: convertMillimetersToTwip(7.5), hanging: convertMillimetersToTwip(4) } } } },
                        { level: 1, format: LevelFormat.BULLET, text: '◦', alignment: AlignmentType.LEFT,
                          style: { paragraph: { indent: { left: convertMillimetersToTwip(12), hanging: convertMillimetersToTwip(4) } } } },
                        { level: 2, format: LevelFormat.BULLET, text: '▪', alignment: AlignmentType.LEFT,
                          style: { paragraph: { indent: { left: convertMillimetersToTwip(16.5), hanging: convertMillimetersToTwip(4) } } } },
                    ]},
                    { levels: [
                        { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
                          style: { paragraph: { indent: { left: convertMillimetersToTwip(7.5), hanging: convertMillimetersToTwip(4) } } } },
                        { level: 1, format: LevelFormat.DECIMAL, text: '%2)', alignment: AlignmentType.LEFT,
                          style: { paragraph: { indent: { left: convertMillimetersToTwip(12), hanging: convertMillimetersToTwip(4) } } } },
                    ]},
                ],
            },
            sections: [{ properties: {}, children: this.children }],
        });
        return doc;
    }

    para(childrenOrOpts, opts) {
        let children = childrenOrOpts;
        if (!Array.isArray(children)) children = [children];
        return new Paragraph(Object.assign({ children }, opts || {}));
    }

    blocks(nodes, sink) {
        let i = 0;
        while (i < nodes.length) {
            const n = nodes[i];
            switch (n.type) {
                case 'heading_open': {
                    const lvl = ({ h1: HeadingLevel.HEADING_1, h2: HeadingLevel.HEADING_2,
                        h3: HeadingLevel.HEADING_3, h4: HeadingLevel.HEADING_4,
                        h5: HeadingLevel.HEADING_5, h6: HeadingLevel.HEADING_6 })[n.tag] || HeadingLevel.HEADING_1;
                    const inline = nodes[i + 1];
                    const parts = runsFromInline(inline && inline.children, {});
                    sink(this.para(buildParagraphChildren(parts), { heading: lvl }));
                    i += 3; break;
                }
                case 'paragraph_open': {
                    const inline = nodes[i + 1];
                    const kids = (inline && inline.children) || [];
                    const mathOnly = kids.length === 1 && kids[0].type === 'math_block';
                    const parts = runsFromInline(kids, {});
                    if (kids.some(k => k.type === 'math_block')) {
                        const raw = kids.filter(k => k.type === 'math_block')
                            .map(k => ({ type: 'raw', xml: texToOmml(k.tex || k.content, true) }));
                        sink(this.para(buildParagraphChildren(raw.concat(parts.filter(p => p.type === 'run'))), { alignment: AlignmentType.CENTER }));
                    } else {
                        sink(this.para(buildParagraphChildren(parts)));
                    }
                    i += 3; break;
                }
                case 'math_block': {
                    sink(this.para(buildParagraphChildren([{ type: 'raw', xml: texToOmml(n.tex || n.content, true) }]), { alignment: AlignmentType.CENTER }));
                    i += 1; break;
                }
                case 'blockquote_open': {
                    const inner = this.span(nodes, i, 'blockquote_close');
                    const buf = [];
                    this.blocks(inner, (el) => buf.push(el));
                    buf.forEach(el => { /* стилизацию цитаты даёт opt below */ });
                    sink(new Paragraph({
                        children: [new TextRun({ text: '❝', italics: true, color: '888888' })],
                        indent: { left: convertMillimetersToTwip(10) },
                    }));
                    // пересобираем со стилем quote
                    const quoted = [];
                    this.blocks(inner, (el) => quoted.push(el));
                    quoted.forEach(q => {
                        if (q instanceof Paragraph && q.rootKey === 'w:p') {
                            // простой путь: добавляем отступ через повторное создание невозможно — оставляем как есть
                        }
                    });
                    sink(...[]); // no-op
                    quoted.forEach(el => sink(el));
                    i = inner.__next || (i + 1);
                    break;
                }
                case 'bullet_list_open': case 'ordered_list_open': {
                    const ordered = n.type === 'ordered_list_open';
                    const inner = this.span(nodes, i, ordered ? 'ordered_list_close' : 'bullet_list_close');
                    this.list(inner, sink, ordered, 0);
                    i = inner.__next || (i + 1);
                    break;
                }
                case 'fence': case 'code_block': {
                    const lines = String(n.content || '').split('\n');
                    if (lines[lines.length - 1] === '') lines.pop();
                    lines.forEach(line => {
                        sink(new Paragraph({
                            children: [new TextRun({ text: line || ' ', font: 'Consolas', size: 20 })],
                            shading: { type: ShadingType.CLEAR, fill: 'F5F5F5' },
                            spacing: { after: 0 },
                        }));
                    });
                    i += 1; break;
                }
                case 'table_open': {
                    const inner = this.span(nodes, i, 'table_close');
                    sink(this.table(inner));
                    i = inner.__next || (i + 1);
                    break;
                }
                case 'hr': {
                    sink(new Paragraph({
                        children: [],
                        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '999999' } },
                    }));
                    i += 1; break;
                }
                default:
                    i += 1;
            }
        }
    }

    // найти детей контейнера до закрывающего токена; вернуть массив + __next
    span(nodes, i, closeType) {
        const depthStack = [];
        const out = [];
        let j = i + 1;
        let depth = 1;
        for (; j < nodes.length; j++) {
            const t = nodes[j];
            if (t.type === closeType && depth === 1) break;
            if (/_open$/.test(t.type) && t.type !== closeType) depth++;
            else if (/_close$/.test(t.type)) depth--;
            out.push(t);
        }
        out.__next = j + 1;
        return out;
    }

    list(itemsNodes, sink, ordered, level) {
        // itemsNodes — плоский список между *_list_open и *_list_close
        let j = 0;
        let counter = 1;
        while (j < itemsNodes.length) {
            const it = itemsNodes[j];
            if (it.type === 'list_item_open') {
                const inner = this.span(itemsNodes, j, 'list_item_close');
                this.listItem(inner, sink, ordered, level, counter);
                if (ordered) counter++;
                j = inner.__next || (j + 1);
            } else j++;
        }
    }

    listItem(inner, sink, ordered, level, counter) {
        // первый абзац элемента — с маркером списка; вложенные списки рекурсивно
        let firstDone = false;
        const self = this;
        const emit = (el) => { sink(el); };
        let k = 0;
        while (k < inner.length) {
            const n = inner[k];
            if (n.type === 'paragraph_open') {
                const inline = inner[k + 1];
                let kids = (inline && inline.children) || [];
                let checkbox = null;
                if (kids[0] && kids[0].type === 'htmltag' && /checkbox/i.test(kids[0].content || '')) {
                    checkbox = /checked/.test(kids[0].content);
                    kids = kids.slice(1);
                } else if (kids[0] && kids[0].type === 'text') {
                    const m = /^\[( |x|X)\]\s*/.exec(kids[0].content || '');
                    if (m) { checkbox = m[1] !== ' '; kids[0] = Object.assign({}, kids[0], { content: kids[0].content.slice(m[0].length) }); }
                }
                const parts = runsFromInline(kids, {});
                if (checkbox !== null) {
                    parts.unshift({ type: 'run', text: checkbox ? '☑ ' : '☐ ', f: {} });
                }
                const opts = { children: buildParagraphChildren(parts) };
                if (!firstDone) {
                    opts.numbering = ordered
                        ? { reference: 'od-' + Math.min(level, 1), level: Math.min(level, 1) }
                        : { reference: 'ul-' + Math.min(level, 2), level: Math.min(level, 2) };
                    firstDone = true;
                }
                sink(new Paragraph(opts));
                k += 3;
            } else if (n.type === 'bullet_list_open' || n.type === 'ordered_list_open') {
                const ordered2 = n.type === 'ordered_list_open';
                const innerL = this.span(inner, k, ordered2 ? 'ordered_list_close' : 'bullet_list_close');
                this.list(innerL, sink, ordered2, level + 1);
                k = innerL.__next || (k + 1);
            } else if (n.type === 'math_block') {
                sink(this.para(buildParagraphChildren([{ type: 'raw', xml: texToOmml(n.tex || n.content, false) }])));
                k++;
            } else {
                k++;
            }
        }
        if (!firstDone) {
            sink(new Paragraph({ children: [], numbering: ordered
                ? { reference: 'od-0', level: 0 } : { reference: 'ul-0', level: 0 } }));
        }
    }

    table(innerNodes) {
        // собрать строки
        const rows = [];
        let cur = null;
        let cellKids = null;
        let header = false;
        const stack = [];
        for (const t of innerNodes) {
            if (t.type === 'tr_open') { cur = []; rows.push(cur); }
            else if (t.type === 'td_open' || t.type === 'th_open') {
                cellKids = []; cur.push({ nodes: cellKids, header: t.type === 'th_open',
                    align: ((t.attrs || []).find(a => a[0] === 'style') || ['', ''])[1] });
                header = header || t.type === 'th_open';
            }
            else if (t.type === 'inline' && cellKids) cellKids.push(t);
            else if (t.type === 'td_close' || t.type === 'th_close') cellKids = null;
        }
        const gridCols = Math.max(1, ...rows.map(r => r.length));
        const tableRows = rows.map(r => new TableRow({
            children: r.map(c => {
                const parts = runsFromInline((c.nodes[0] && c.nodes[0].children) || [], {});
                return new TableCell({
                    children: [new Paragraph({
                        children: buildParagraphChildren(parts),
                        alignment: /text-align:\s*center/.test(c.align || '') ? AlignmentType.CENTER
                            : /text-align:\s*right/.test(c.align || '') ? AlignmentType.RIGHT : AlignmentType.LEFT,
                    })],
                    shading: c.header ? { type: ShadingType.CLEAR, fill: 'E8EDF5' } : undefined,
                    margins: { top: 60, bottom: 60, left: 100, right: 100 },
                });
            }),
            tableHeader: r.some(c => c.header),
        }));
        return new Table({
            rows: tableRows,
            width: { size: 100, type: WidthType.PERCENTAGE },
        });
    }
}

async function exportDocx(astTree, outputPath, opts) {
    const builder = new Builder(opts || {});
    const doc = builder.build(astTree);
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outputPath, buffer);
    return outputPath;
}

module.exports = { exportDocx, texToOmml };
