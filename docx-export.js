'use strict';

/**
 * Базовый экспорт Markdown → DOCX.
 *
 * Поддерживается:
 *  - заголовки (#..######), абзацы, горизонтальные правила;
 *  - инлайн-форматирование: **жирный**, *курсив*, `код`, ~~зачёркнутый~~;
 *  - списки (нумерованные и маркированные);
 *  - блоки кода ``` (моноширинным шрифтом);
 *  - математика в LaTeX-синтаксисе: $...$ (инлайн) и $$...$$ / \[...\] (отдельные блоки)
 *    — конвертируется в нативные формулы Word (Office Math).
 */

const {
    Document, Packer, Paragraph, TextRun,
    Math: DocxMath, MathRun, MathFraction, MathRadical, MathSubScript, MathSuperScript,
    MathSum, MathIntegral, MathProduct,
} = require('docx');

// ──────────────────────────────────────────────
// Конвертер LaTeX → компоненты docx Math
// ──────────────────────────────────────────────

const GREEK = {
    alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε',
    zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ',
    nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'φ',
    chi: 'χ', psi: 'ψ', omega: 'ω',
    Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
};

const SYMBOLS = {
    times: '×', div: '÷', pm: '±', mp: '∓', cdot: '⋅', ast: '∗', star: '⋆', circ: '∘', bullet: '∙',
    leq: '≤', le: '≤', geq: '≥', ge: '≥', neq: '≠', ne: '≠', approx: '≈', equiv: '≡', sim: '∼', cong: '≅',
    ll: '≪', gg: '≫',
    infty: '∞', partial: '∂', nabla: '∇', forall: '∀', exists: '∃', neg: '¬',
    wedge: '∧', lando: '∧', vee: '∨', lor: '∨', cap: '∩', cup: '∪', subset: '⊂', supset: '⊃',
    subseteq: '⊆', supseteq: '⊇', emptyset: '∅', varnothing: '∅',
    to: '→', rightarrow: '→', leftarrow: '←', Rightarrow: '⇒', Leftarrow: '⇐', leftrightarrow: '↔',
    Leftrightarrow: '⇔', mapsto: '↦', iff: '⇔', therefore: '∴', because: '∵',
    ldots: '…', dots: '…', cdots: '⋯', prime: '′', degree: '°', angle: '∠', triangle: '△',
    in: '∈', ni: '∋', perp: '⊥', parallel: '∥', propto: '∝', oplus: '⊕', otimes: '⊗',
};

const BIG_OPERATORS = { sum: '∑', prod: '∏', int: '∫', oint: '∮', bigcup: '⋃', bigcap: '⋂', lim: 'lim' };

// Функции, выводимые прямым (не курсивным) шрифтом
const FUNC_NAMES = [
    'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh', 'coth', 'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
    'log', 'ln', 'exp', 'max', 'min', 'sup', 'inf', 'det', 'dim', 'ker', 'deg', 'gcd', 'arg',
];

class Parser {
    constructor(src) {
        this.s = src;
        this.i = 0;
    }

    eof() {
        // пропускаем пробелы lookahead'ом не меняя позицию
        let j = this.i;
        while (j < this.s.length && /\s/.test(this.s[j])) j++;
        return j >= this.s.length;
    }

    peekChar() {
        let j = this.i;
        while (j < this.s.length && /\s/.test(this.s[j])) j++;
        return j < this.s.length ? this.s[j] : null;
    }

    skipWs() {
        while (this.i < this.s.length && /\s/.test(this.s[this.i])) this.i++;
    }

    parseGroup(stopAtBrace) {
        const children = [];
        while (!this.eof()) {
            const c = this.peekChar();
            if (stopAtBrace && c === '}') break;
            if (!stopAtBrace && c === ')') break;
            const nodes = this.parseAtom();
            for (const n of nodes) if (n) children.push(n);
        }
        return children;
    }

    // Основной узел: атом + возможные постфиксы (_{}, ^{}, /)
    // Возвращает массив узлов (обычно из одного элемента), т.к. пробелы
    // передаются отдельными MathRun(' ') — Word показывает их как thin space.
    parseAtom() {
        // явные пробелы в LaTeX игнорируются, но один «тонкий» пробел сохраняем
        let spaces = 0;
        while (this.i < this.s.length && /\s/.test(this.s[this.i])) { spaces++; this.i++; }

        let node = this.parsePrefix();
        // Большие операторы (∑, ∫, ∏) с пределами → nary-конструкция Word
        if (node && node.__bigOp) {
            const big = node;
            let sub = null, sup = null;
            for (;;) {
                const c = this.peekChar();
                if (c === '_' || c === '^') {
                    this.skipWs();
                    this.i++;
                    const script = this.parseArgOrChar();
                    if (c === '_') sub = script; else sup = script;
                } else break;
            }
            const opts = { children: [new MathRun('')] };
            if (sub) opts.subScript = sub;
            if (sup) opts.superScript = sup;
            node = new big.Ctor(opts);
        }
        for (;;) {
            const c = this.peekChar();
            if (c === '_' || c === '^') {
                this.skipWs();
                this.i++; // consume _ or ^
                const script = this.parseArgOrChar();
                node = c === '_'
                    ? new MathSubScript({ children: [node], subScript: script })
                    : new MathSuperScript({ children: [node], superScript: script });
            } else if (c === '/') {
                this.skipWs();
                this.i++; // consume /
                this.skipWs();
                const denom = this.parseArgOrChar();
                node = new MathFraction({ numerator: [node], denominator: denom });
            } else {
                break;
            }
        }

        const out = [];
        if (spaces > 0 && node) out.push(new MathRun(' '));
        if (node) out.push(node);
        return out;
    }

    // Аргумент после _/^ или знаменатель после /: один символ или {...}
    parseArgOrChar() {
        this.skipWs();
        const c = this.s[this.i];
        if (c === '{') {
            this.i++;
            return this.parseGroup(true);
        }
        if (c === '\\') {
            const n = this.parsePrefix();
            return n ? [n] : [];
        }
        if (c !== undefined && c !== '') {
            this.i++;
            return [new MathRun(c)];
        }
        return [];
    }

    readToken() {
        this.skipWs();
        if (this.s[this.i] !== '\\') return null;
        this.i++;
        let t = '';
        while (this.i < this.s.length && /[a-zA-Z]/.test(this.s[this.i])) t += this.s[this.i++];
        if (t === '') {
            const ch = this.s[this.i];
            if (ch !== undefined) { this.i++; return { name: ch, raw: true }; }
            return { name: '', raw: true };
        }
        return { name: t, raw: false };
    }

    grabBracedContent() {
        this.skipWs();
        if (this.s[this.i] !== '{') return null;
        this.i++;
        let depth = 1, out = '';
        while (this.i < this.s.length) {
            const c = this.s[this.i];
            if (c === '{') depth++;
            else if (c === '}') { depth--; if (depth === 0) { this.i++; break; } }
            out += c;
            this.i++;
        }
        return out;
    }

    parsePrefix() {
        if (this.eof()) return null;
        const save = this.i;
        const c = this.peekChar();

        if (c === '{') {
            this.skipWs();
            this.i++;
            const children = this.parseGroup(true);
            this.skipWs();
            if (this.s[this.i] === '}') this.i++;
            // вложенная группа без смысла-обёртки — просто перечисляем детей через DocxMath
            return new DocxMath({ children });
        }

        if (c === '\\') {
            const tok = this.readToken();
            const name = tok ? tok.name : '';

            if (name === 'frac' || name === 'dfrac' || name === 'tfrac') {
                const num = this.grabBracedContent();
                const den = this.grabBracedContent();
                const p = new Parser(num == null ? '' : num);
                const q = new Parser(den == null ? '' : den);
                return new MathFraction({ numerator: p.parseGroup(false), denominator: q.parseGroup(false) });
            }
            if (name === 'sqrt') {
                const arg = this.grabBracedContent();
                const inner = arg != null ? new Parser(arg).parseGroup(false) : [];
                return new MathRadical({ children: inner });
            }
            if (name === 'text' || name === 'mathrm' || name === 'operatorname' || name === 'mbox') {
                const arg = this.grabBracedContent();
                return new MathRun(arg == null ? '' : arg.replace(/\\/g, ''));
            }
            if (name === 'left' || name === 'right') {
                // \left( ... \right) — скобки включаются как обычные символы
                const ch = this.s[this.i];
                if (ch && '()[]{}|.,'.includes(ch)) { this.i++; return new MathRun(ch === '.' ? '' : ch); }
                return null;
            }
            if (name === 'cdot') return new MathRun('·');
            if (BIG_OPERATORS[name]) {
                // маркер большого оператора — parseAtom соберёт из него nary с пределами
                const marker = new MathRun(BIG_OPERATORS[name]);
                marker.__bigOp = true;
                marker.Ctor = name === 'int' || name === 'oint' ? MathIntegral
                    : name === 'prod' ? MathProduct : MathSum;
                return marker;
            }
            if (GREEK[name]) return new MathRun(GREEK[name]);
            if (SYMBOLS[name]) return new MathRun(SYMBOLS[name]);
            if (tok && tok.raw) return new MathRun(name);
            // неизвестная команда — вывести имя как текст
            return new MathRun(name);
        }

        if (/[a-zA-Z]/.test(c)) {
            // возможное имя функции (sin, log...)
            let j = this.i;
            while (j < this.s.length && /\s/.test(this.s[j])) j++;
            let word = '';
            while (j < this.s.length && /[a-zA-Z]/.test(this.s[j])) word += this.s[j++];
            if (FUNC_NAMES.includes(word)) {
                this.i = j;
                return new MathRun(word);
            }
            this.i++;
            return new MathRun(c);
        }

        if (c !== undefined && c !== '') {
            this.i++;
            return new MathRun(c);
        }

        this.i = save;
        return null;
    }
}

function latexToDocxMath(latex) {
    try {
        const parser = new Parser(latex);
        const children = parser.parseGroup(false);
        if (!children.length) return new DocxMath({ children: [new MathRun(latex)] });
        return new DocxMath({ children });
    } catch (e) {
        return new DocxMath({ children: [new MathRun(latex)] });
    }
}

// ──────────────────────────────────────────────
// Инлайн-парсер Markdown (текст + математика + форматирование)
// ──────────────────────────────────────────────

function splitInline(text) {
    const parts = [];
    const re = /(\$\$[\s\S]+?\$\$|\$[^$\n]+?\$)/g;
    let last = 0, m;
    while ((m = re.exec(text)) !== null) {
        if (m.index > last) parts.push({ type: 'text', value: text.slice(last, m.index) });
        const t = m[0];
        if (t.startsWith('$$')) parts.push({ type: 'math', value: t.slice(2, -2).trim(), display: true });
        else parts.push({ type: 'math', value: t.slice(1, -1).trim(), display: false });
        last = re.lastIndex;
    }
    if (last < text.length) parts.push({ type: 'text', value: text.slice(last) });
    return parts;
}

// Инлайн-токены: **жирный**, __жирный__, *курсив*, _курсив_, `код`, ~~зачёркнутый~~.
const INLINE_TOKEN_RE = /(\*\*[^*\n]+\*\*|__[^_\n]+__|~~[^~\n]+~~|`[^`\n]+`|\*[^*\s][^*\n]*\*|_[^_\s][^_\n]*_)/g;

function parseRuns(text, fmt) {
    fmt = fmt || {};
    const runs = [];
    let last = 0, m;
    while ((m = INLINE_TOKEN_RE.exec(text)) !== null) {
        if (m.index > last) {
            const plain = text.slice(last, m.index);
            if (plain) runs.push(new TextRun({ text: plain, ...fmt }));
        }
        const s = m[0];
        if (s.startsWith('**') || s.startsWith('__')) {
            runs.push(...parseRuns(s.slice(2, -2), { ...fmt, bold: true }));
        } else if (s.startsWith('~~')) {
            runs.push(...parseRuns(s.slice(2, -2), { ...fmt, strike: true }));
        } else if (s.startsWith('`')) {
            runs.push(new TextRun({ text: s.slice(1, -1), font: 'Consolas', ...fmt }));
        } else {
            runs.push(...parseRuns(s.slice(1, -1), { ...fmt, italics: true }));
        }
        last = INLINE_TOKEN_RE.lastIndex;
    }
    if (last < text.length) {
        const plain = text.slice(last);
        if (plain) runs.push(new TextRun({ text: plain, ...fmt }));
    }
    return runs;
}

function inlineChildren(text) {
    const children = [];
    for (const part of splitInline(text)) {
        if (part.type === 'math') children.push(latexToDocxMath(part.value));
        else children.push(...parseRuns(part.value));
    }
    return children;
}

// ──────────────────────────────────────────────
// Поблочный парсер Markdown
// ──────────────────────────────────────────────

function mdToDocxElements(md) {
    const lines = md.replace(/\r\n/g, '\n').split('\n');
    const blocks = []; // {type, ...}
    let i = 0;

    const isBreakLine = (l) => /^\s*$/.test(l) ||
        /^#{1,6}\s+/.test(l) ||
        /^\s*([-*+]|\d+[.)])\s+/.test(l) ||
        /^\s*(---+|___+|\*\*\*+)\s*$/.test(l) ||
        /^\s*>/.test(l);

    while (i < lines.length) {
        const line = lines[i];

        if (/^\s*$/.test(line)) { i++; continue; }

        // Блок математики $$...$$ (открыт на этой строке)
        const ddOpen = line.match(/^\s*\$\$(.*)$/);
        if (ddOpen && !/\$\$/.test(ddOpen[1])) {
            let buf = ddOpen[1] + '\n';
            i++;
            while (i < lines.length && !/^\s*.*\$\$/.test(lines[i])) buf += lines[i++] + '\n';
            if (i < lines.length) {
                const closeIdx = lines[i].indexOf('$$');
                buf += lines[i].slice(0, closeIdx);
                i++;
            }
            blocks.push({ type: 'mathblock', latex: buf.trim() });
            continue;
        }
        // Однострочный $$...$$
        const ddOne = line.match(/^\s*\$\$(.+?)\$\$\s*$/);
        if (ddOne) { blocks.push({ type: 'mathblock', latex: ddOne[1].trim() }); i++; continue; }

        // \[ ... \]
        const brOpen = line.match(/^\s*\\\[(.*)$/);
        if (brOpen && !/\\\]/.test(brOpen[1])) {
            let buf = brOpen[1] + '\n';
            i++;
            while (i < lines.length && !/\\\]/.test(lines[i])) buf += lines[i++] + '\n';
            if (i < lines.length) buf += lines[i].replace(/\\\].*$/, '');
            i++;
            blocks.push({ type: 'mathblock', latex: buf.trim() });
            continue;
        }

        // Кодовый блок ```
        const cb = line.match(/^```(.*)$/);
        if (cb) {
            i++;
            const code = [];
            while (i < lines.length && !/^```/.test(lines[i])) code.push(lines[i++]);
            i++; // closing ```
            blocks.push({ type: 'code', lines: code });
            continue;
        }

        // Заголовок
        const h = line.match(/^(#{1,6})\s+(.*)$/);
        if (h) { blocks.push({ type: 'heading', level: h[1].length, text: h[2].trim() }); i++; continue; }

        // Горизонтальная черта
        if (/^\s*(-{3,}|\*{3,}|_{3,})\s*$/.test(line)) { blocks.push({ type: 'hr' }); i++; continue; }

        // Список
        const li = line.match(/^(\s*)([-*+]|\d+[.)])\s+(.*)$/);
        if (li) {
            const ordered = /\d/.test(li[2]);
            const items = [];
            while (i < lines.length) {
                const l2 = lines[i].match(/^(\s*)([-*+]|\d+[.)])\s+(.*)$/);
                if (l2 && (!!l2[3] === true)) {
                    items.push({ indent: Math.floor(l2[1].length / 2), text: l2[3] });
                    i++;
                    // continuation-строки пункта
                    while (i < lines.length && /^\s+\S/.test(lines[i]) && !isBreakLine(lines[i])) {
                        items[items.length - 1].text += ' ' + lines[i].trim();
                        i++;
                    }
                } else break;
            }
            blocks.push({ type: 'list', ordered, items });
            continue;
        }

        // Цитата
        if (/^\s*>/.test(line)) {
            const quote = [];
            while (i < lines.length && /^\s*>/.test(lines[i])) quote.push(lines[i++].replace(/^\s*>\s?/, ''));
            blocks.push({ type: 'quote', text: quote.join(' ').trim() });
            continue;
        }

        // Абзац (собираем соседние строки)
        let para = line.trim();
        i++;
        while (i < lines.length && !isBreakLine(lines[i]) && !/^\s*\$\$/.test(lines[i]) && !/^\\\[/.test(lines[i])) {
            para += ' ' + lines[i].trim();
            i++;
        }
        blocks.push({ type: 'paragraph', text: para });
    }

    return blocks.reduce((acc, b) => acc.concat(blockToElements(b)), []);
}

function blockToElements(b) {
    switch (b.type) {
        case 'heading':
            return [new Paragraph({ heading: `Heading${b.level}`, spacing: { before: 240, after: 120 }, children: inlineChildren(b.text) })];
        case 'paragraph':
            return [new Paragraph({ spacing: { after: 120 }, children: inlineChildren(b.text) })];
        case 'mathblock':
            return [new Paragraph({ alignment: 'center', spacing: { before: 120, after: 120 }, children: [latexToDocxMath(b.latex)] })];
        case 'code':
            return b.lines.map((l) => new Paragraph({
                shading: { fill: 'F5F5F5' },
                spacing: { after: 0 },
                children: [new TextRun({ text: l || ' ', font: 'Consolas', size: 20 })],
            }));
        case 'quote':
            return [new Paragraph({
                indent: { left: 720 },
                spacing: { after: 120 },
                children: inlineChildren(b.text).map((c) => c),
            })];
        case 'list':
            return b.items.map((it) => new Paragraph({
                bullet: b.ordered ? undefined : { level: Math.min(it.indent, 2) },
                numbering: b.ordered ? { reference: 'md-ordered', level: Math.min(it.indent, 1) } : undefined,
                children: inlineChildren(it.text),
            }));
        case 'hr':
            return [new Paragraph({ thematicBreak: true })];
        default:
            return [new Paragraph({ children: [] })];
    }
}

async function exportMarkdownToDocx(markdownText) {
    const elements = mdToDocxElements(markdownText);
    const doc = new Document({
        features: { math: { char: 'Cambria Math' } },
        numbering: {
            config: [
                {
                    reference: 'md-ordered',
                    levels: [
                        { level: 0, format: 'decimal', text: '%1.' },
                        { level: 1, format: 'lowerLetter', text: '%2.' },
                    ],
                },
            ],
        },
        sections: [{ children: elements }],
    });
    return Packer.toBuffer(doc);
}

module.exports = { exportMarkdownToDocx, mdToDocxElements, latexToDocxMath };
