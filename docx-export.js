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
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    Math: DocxMath, MathRun, MathFraction, MathRadical, MathSubScript, MathSuperScript,
    MathSum, MathIntegral, MathProduct, MathAccent, MathRoundBrackets, MathSquareBrackets,
    MathCurlyBrackets, MathAngledBrackets, MathBox,
    ExternalHyperlink, InternalHyperlink, CommentRangeStart, CommentRangeEnd,
    FootnoteReferenceRun, PageNumber, BorderStyle, WidthType, VerticalAlign, AlignmentType,
} = require('docx');

// Элементы, которые можно добавлять в абзац напрямую (дочерние узлы w:p).
const PARAGRAPH_LEVEL_TYPES = new Set([
    'document', 'paragraph', 'run', 'external-hyperlink', 'internal-hyperlink',
    'comment-range-start', 'comment-range-end', 'footnote-reference-run', 'page-number',
]);

function isParagraphLevel(el) {
    return el instanceof Paragraph || el instanceof TextRun || el instanceof Table ||
        el instanceof ExternalHyperlink || el instanceof InternalHyperlink ||
        el instanceof CommentRangeStart || el instanceof CommentRangeEnd ||
        el instanceof FootnoteReferenceRun || el === PageNumber;
}

// Внутри составных конструкций Word (m:d, m:f, m:sSub, m:m ...) все дети
// должны быть прямыми потомками соответствующего элемента. Если в children
// попался «контейнер» DocxMath (например, группа {...}), раскрываем его —
// иначе получается <m:oMath> внутри <m:e>/<m:num>, что Word не может прочитать.
function normalizeChildren(arr) {
    const out = [];
    for (const el of arr || []) {
        if (!el) continue;
        if (el instanceof DocxMath) {
            // контейнер-группа: поднимаем детей наверх (рекурсивно нормализуя)
            out.push(...normalizeChildren(el.children));
        } else {
            if (Array.isArray(el.children)) el.children = normalizeChildren(el.children);
            out.push(el);
        }
    }
    return out;
}
function wrapCompositeChildren(node) {
    try {
        if (node && Array.isArray(node.children)) node.children = normalizeChildren(node.children);
    } catch (e) { /* best effort */ }
}

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

const BIG_OPERATORS = { sum: '∑', prod: '∏', int: '∫', oint: '∮', iint: '∬', iiint: '∭', bigcup: '⋃', bigcap: '⋂', lim: 'lim' };

// Акценты (\hat{x}, \vec{a} ...) → символ акцента Word
const ACCENTS = {
    hat: '\u0302', widehat: '\u0302', tilde: '\u0303', widetilde: '\u0303',
    bar: '\u0304', overline: '\u0305', vec: '\u20D7', dot: '\u0307', ddot: '\u0308',
    check: '\u030C', breve: '\u0306', acute: '\u0301', grave: '\u0300',
    mathring: '\u030A', ring: '\u030A', grade: '\u030C',
};

// Конструкции «обёртки» с одним аргументом-узлом: m:d и т.п.
const NAME_TO_OMATH = {
    overbrace: MathRoundBrackets, underbrace: MathRoundBrackets,
    overview: MathRoundBrackets, underbar: MathSquareBrackets,
};

// Таблица без рамок (для матриц/случаев внутри формул)
const NO_BORDER = {
    top: { style: BorderStyle.NONE, size: 0, color: 'auto' },
    bottom: { style: BorderStyle.NONE, size: 0, color: 'auto' },
    left: { style: BorderStyle.NONE, size: 0, color: 'auto' },
    right: { style: BorderStyle.NONE, size: 0, color: 'auto' },
};

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

    // Подстрока, начинающаяся с первого непробельного символа от this.i
    // (для проверки многосимвольных команд вида \begin{bmatrix}).
    restTrimmed() {
        let j = this.i;
        while (j < this.s.length && /\s/.test(this.s[j])) j++;
        return this.s.slice(j);
    }

    skipWs() {
        while (this.i < this.s.length && /\s/.test(this.s[this.i])) this.i++;
    }

    parseGroup(stopAtBrace) {
        const children = [];
        let braceDepth = 0;
        while (this.i < this.s.length) {
            const c = this.peekChar();
            if (c === null) break; // остались только пробелы
            if (c === '}') {
                // незакрытая группа или конец нашей группы — выходим,
                // чтобы «}» не превратился в лишний символ в документе
                if (stopAtBrace || braceDepth <= 0) break;
            }
            if (!stopAtBrace && c === ')') break;
            if (c === '{') braceDepth++;
            const before = this.i;
            const nodes = this.parseAtom();
            for (const n of nodes) if (n) children.push(n);
            // Защита от зависания: если позиция не сдвинулась — принудительно
            // пропускаем один символ, чтобы цикл гарантированно завершался.
            if (this.i === before) this.i++;
            else if (this.s.slice(before, this.i).includes('}')) braceDepth--;
        }
        return children;
    }

    // Аргумент-фигурные скобки {...} → массив узлов (с возвратом курсора на место).
    parseArgNodes() {
        const save = this.i;
        try {
            this.skipWs();
            if (this.s[this.i] !== '{') { this.i = save; return null; }
            this.i++;
            const nodes = this.parseGroup(true);
            this.skipWs();
            if (this.s[this.i] === '}') this.i++;
            return nodes;
        } catch (e) {
            this.i = save;
            return null;
        }
    }

    // Одиночный ограничитель после \left / \right / \big и т.п.
    parseDelimiterRun() {
        this.skipWs();
        const ch = this.s[this.i];
        if (ch === undefined) return null;
        if ('()[]|.,/'.includes(ch)) { this.i++; return new MathRun(ch === '.' ? '' : ch); }
        if (ch === '{') { this.i++; return new MathRun('{'); }
        if (ch === '\\') {
            const save = this.i;
            const t2 = this.readToken();
            const n2 = t2 ? t2.name : '';
            if (n2 === 'lbrace') return new MathRun('{');
            if (n2 === 'rbrace') return new MathRun('}');
            if (n2 === 'langle') return new MathRun('⟨');
            if (n2 === 'rangle') return new MathRun('⟩');
            if (n2 === '.') return new MathRun('');
            if (n2 === '|') return new MathRun('|');
            if (GREEK[n2]) return new MathRun(GREEK[n2]);
            if (SYMBOLS[n2]) return new MathRun(SYMBOLS[n2]);
            this.i = save;
            return null;
        }
        return null;
    }

    // Линейное представление массива узлов для матриц/случаев:
    // строки разделяются \\ или &, ячейки склеиваются в текст.
    static mathRunText(n) {
        if (n instanceof MathRun) {
            try { return n.root.options.text; } catch (e) { return ''; }
        }
        return null;
    }

    static toMatrixRows(nodes) {
        let src = '';
        for (const n of nodes || []) {
            const t = Parser.mathRunText(n);
            if (t !== null) src += t;
            else if (n && n.constructor && /Math(Fraction|Radical|SubScript|SuperScript)$/.test(n.constructor.name)) src += '□';
            else src += ' ';
        }
        src = src.replace(/&/g, '\\\\');
        const rows = src.split('\\\\').map((r) => r.trim()).filter((r) => r !== '');
        return rows.map((r) => r.split(/\s+/));
    }

    // \begin{X}...\end{X}: матрицы/случаи → таблица Word, остальное — содержимое как есть.
    parseEnvironment(envName, beginRaw) {
        const startI = this.i;
        const endIdx = this.s.indexOf('\\end{' + envName, startI);
        const bodySrc = endIdx >= 0 ? this.s.slice(startI, endIdx) : this.s.slice(startI);
        this.i = endIdx >= 0 ? endIdx : this.s.length;
        const bodyNodes = new Parser(bodySrc).parseGroup(false);

        const matrixEnvs = ['matrix', 'pmatrix', 'bmatrix', 'vmatrix', 'Vmatrix', 'smallmatrix'];
        const caseEnvs = ['cases', 'rcases'];
        if (matrixEnvs.includes(envName) || caseEnvs.includes(envName)) {
            const rows = Parser.toMatrixRows(bodyNodes);
            if (rows.length) {
                const cols = Math.max(...rows.map((r) => r.length));
                const tblRows = rows.map((cells) => new TableRow({
                    children: Array.from({ length: cols }, (_, c) => new TableCell({
                        borders: NO_BORDER,
                        margins: { top: 40, bottom: 40, left: 80, right: 80 },
                        verticalAlign: VerticalAlign.CENTER,
                        children: [new Paragraph({
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 0 },
                            children: [new TextRun({ text: cells[c] != null ? cells[c] : '', size: 22 })],
                        })],
                    })),
                }));
                const table = new Table({
                    columnWidths: Array.from({ length: cols }, () => 900),
                    width: { size: 900 * cols, type: WidthType.DXA },
                    rows: tblRows,
                });
                const bracket = { pmatrix: ['(', ')'], bmatrix: ['[', ']'], vmatrix: ['|', '|'], Vmatrix: ['‖', '‖'] }[envName];
                const children = [];
                if (caseEnvs.includes(envName)) children.push(new MathRun('{'));
                else if (bracket) children.push(new MathRun(bracket[0]));
                children.push(new MathBox({ children: [table] }));
                if (bracket) children.push(new MathRun(bracket[1]));
                return new DocxMath({ children });
            }
        }
        // Неизвестное окружение — просто выводим содержимое
        return new DocxMath({ children: bodyNodes.length ? bodyNodes : [new MathRun(beginRaw)] });
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
            const opts = { children: [new MathRun(big.text)] };
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
        if (node) {
            // Все children должны быть прямыми потомками <m:oMath>: оборачиваем
            // составные конструкции (матрицы/скобки с содержимым-узлами).
            wrapCompositeChildren(node);
            out.push(node);
        }
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

    // Содержимое ближайших {...} (или одиночный токен) в виде массива узлов.
    grabArgNodes() {
        this.skipWs();
        if (this.s[this.i] === '{') {
            this.i++;
            const nodes = this.parseGroup(true);
            this.skipWs();
            if (this.s[this.i] === '}') this.i++;
            return nodes;
        }
        if (this.s[this.i] === '\\') {
            const n = this.parsePrefix();
            return n ? [n] : [];
        }
        const c = this.s[this.i];
        if (c !== undefined && c !== '') { this.i++; return [new MathRun(c)]; }
        return [];
    }

    // Строка до первого непробельного символа (для \tag{...})
    skipSpacesOnly() {
        while (this.i < this.s.length && /\s/.test(this.s[this.i])) this.i++;
    }

    // \left/ \right: собираем содержимое до парной команды как m:d с нужными скобками.
    parseBracketPair(openName) {
        const closer = openName === 'left' ? 'right' : 'left';
        const startI = this.i;
        let depth = 0;
        let j = startI;
        let delimIdx = -1;
        while (j < this.s.length) {
            if (this.s.startsWith('\\left', j)) { depth++; j += 5; continue; }
            if (this.s.startsWith('\\' + closer, j)) {
                if (depth === 0) { delimIdx = j; break; }
                depth--; j += closer.length + 1; continue;
            }
            j++;
        }
        const innerSrc = delimIdx >= 0 ? this.s.slice(startI, delimIdx) : this.s.slice(startI);
        this.i = delimIdx >= 0 ? delimIdx : this.s.length;
        const openDelim = this.parseDelimiterRun();
        const closeDelim = this.parseDelimiterRun();
        const children = new Parser(innerSrc).parseGroup(false);
        const openCh = openDelim ? Parser.mathRunText(openDelim) || '' : '';
        const closeCh = closeDelim ? Parser.mathRunText(closeDelim) || '' : '';
        const bracketKind = { '(' : 'round', '[' : 'square', '{': 'curly', '|': 'round', '‖': 'double', '⟨': 'angle' }[openCh] || 'none';
        const opts = { children };
        if (bracketKind !== 'none') {
            if (openCh === '|') opts.beginCharacter = '|';
            if (openCh === '‖') opts.beginCharacter = '‖';
            if (closeCh === '|') opts.endCharacter = '|';
            if (closeCh === '‖') opts.endCharacter = '‖';
        }
        switch (bracketKind) {
            case 'round': return new MathRoundBrackets(opts);
            case 'square': return new MathSquareBrackets(opts);
            case 'curly': return new MathCurlyBrackets(opts);
            case 'angle': return new MathAngledBrackets(opts);
            default: return new DocxMath({ children });
        }
    }

    // Постфиксы: ', " ^ _, акценты над одним символом
    applyPostfixes(node) {
        for (;;) {
            const c = this.s[this.i];
            if (c === "'" || c === '"') {
                this.i++;
                node = new MathSuperScript({ children: [node], superScript: [new MathRun(c === "'" ? '′' : '″')] });
            } else break;
        }
        return node;
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

            if (name === 'frac' || name === 'dfrac' || name === 'tfrac' || name === 'cfrac') {
                const num = this.grabArgNodes();
                const den = this.grabArgNodes();
                return new MathFraction({ numerator: num, denominator: den });
            }
            if (name === 'binom') {
                const top = this.grabArgNodes();
                const bot = this.grabArgNodes();
                return new DocxMath({
                    children: [
                        new MathRun('('),
                        new MathFraction({ numerator: top, denominator: bot }),
                        new MathRun(')'),
                    ],
                });
            }
            if (name === 'sqrt') {
                this.skipWs();
                // \sqrt[n]{x} — индекс корня
                let degree = null;
                if (this.s[this.i] === '[') {
                    const close = this.s.indexOf(']', this.i);
                    if (close > this.i) {
                        const dsrc = this.s.slice(this.i + 1, close);
                        this.i = close + 1;
                        degree = new Parser(dsrc).parseGroup(false);
                    }
                }
                const arg = this.grabArgNodes();
                return new MathRadical({ children: arg.length ? arg : [], root: degree || undefined });
            }
            if (name === 'text' || name === 'textrm' || name === 'mbox' || name === 'intertext') {
                const arg = this.grabBracedContent();
                return new MathRun(arg == null ? '' : arg.replace(/\\/g, ''));
            }
            if (name === 'mathrm' || name === 'operatorname') {
                const arg = this.grabBracedContent();
                return new MathRun(arg == null ? '' : arg.replace(/\\/g, ''));
            }
            if (name === 'mathbf' || name === 'boldsymbol' || name === 'bm') {
                const arg = this.grabBracedContent();
                return new MathRun(arg == null ? '' : arg.replace(/\\/g, ''));
            }
            if (name === 'mathit' || name === 'mathsf' || name === 'mathtt' || name === 'mathcal' || name === 'mathfrak') {
                const arg = this.grabBracedContent();
                return new MathRun(arg == null ? '' : arg.replace(/\\/g, ''));
            }
            if (name === 'begin') {
                const m = this.restTrimmed().match(/^\{([a-zA-Z*]+)\}/);
                if (m) {
                    this.skipWs();
                    this.i += m[0].length;
                    return this.parseEnvironment(m[1], '\\begin' + m[0]);
                }
                return new MathRun('begin');
            }
            if (name === 'end') {
                const m = this.restTrimmed().match(/^\{[^}]*\}/);
                if (m) { this.skipWs(); this.i += m[0].length; }
                return null;
            }
            if (name === 'left' || name === 'right') {
                return this.parseBracketPair(name);
            }
            if (name === 'big' || name === 'Big' || name === 'bigg' || name === 'Bigg' ||
                name === 'bigl' || name === 'bigr' || name === 'biggl' || name === 'biggr' ||
                name === 'Bigl' || name === 'Bigr' || name === 'middle') {
                const r = this.parseDelimiterRun();
                return r || null;
            }
            if (name === 'over') {
                // a \over b handled simply as fraction of preceding? rare — render as ⁄
                return new MathRun('/');
            }
            if (name === 'dots' || name === 'ldots' || name === 'cdots' || name === 'vdots' || name === 'ddots') {
                return new MathRun(name === 'vdots' ? '⋮' : name === 'ddots' ? '⋱' : '…');
            }
            if (name === 'quad' || name === 'qquad' || name === ',' || name === ';' || name === ':' || name === '!') {
                return new MathRun(name === 'qquad' ? '  ' : ' ');
            }
            if (name === 'tag') {
                // \tag{1} — номер формулы: выводим текстом после формулы
                const arg = this.grabBracedContent();
                return new MathRun(arg == null ? '' : '  (' + arg + ')');
            }
            if (BIG_OPERATORS[name]) {
                // маркер большого оператора — parseAtom соберёт из него nary с пределами
                const marker = new MathRun(BIG_OPERATORS[name]);
                marker.__bigOp = true;
                marker.text = BIG_OPERATORS[name];
                marker.Ctor = name === 'oint' || name === 'iint' || name === 'iiint' || name === 'int'
                    ? MathIntegral
                    : name === 'prod' ? MathProduct : MathSum;
                return marker;
            }
            if (ACCENTS[name]) {
                const arg = this.grabArgNodes();
                return new MathAccent({ accent: ACCENTS[name], children: arg });
            }
            if (NAME_TO_OMATH[name]) {
                const arg = this.grabArgNodes();
                return new NAME_TO_OMATH[name]({ children: arg });
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

// Разбивает формулу на «математическую» часть и текстовые куски (\text{...}, \tag{...}),
// которые в Word корректно отображаются только как обычный текст вокруг oMath.
// Возвращает массив элементов: { math } | { text, bold? }.
function splitTextChunks(latex) {
    const parts = [];
    let buf = '';
    let i = 0;
    const pushMath = () => { if (buf.trim()) parts.push({ math: buf }); buf = ''; };
    const grabBrace = (from) => {
        // from — индекс '{'; возвращает [содержимое, индекс после '}']
        let depth = 0, j = from;
        for (; j < latex.length; j++) {
            if (latex[j] === '{') depth++;
            else if (latex[j] === '}') { depth--; if (depth === 0) return [latex.slice(from + 1, j), j + 1]; }
        }
        return [latex.slice(from + 1), j];
    };
    while (i < latex.length) {
        if (latex.startsWith('\\text', i) && /\{/.test(latex.slice(i, i + 8))) {
            let k = i + 5;
            while (k < latex.length && /\s/.test(latex[k])) k++;
            if (latex[k] === '{') {
                const [content, end] = grabBrace(k);
                pushMath();
                parts.push({ text: content.replace(/\\/g, '') });
                i = end;
                continue;
            }
        }
        if (latex.startsWith('\\tag', i)) {
            let k = i + 4;
            while (k < latex.length && /\s/.test(latex[k])) k++;
            if (latex[k] === '{') {
                const [content, end] = grabBrace(k);
                pushMath();
                parts.push({ text: '  (' + content + ')' });
                i = end;
                continue;
            }
        }
        buf += latex[i];
        i++;
    }
    pushMath();
    return parts;
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

// Блочная формула: набор абзацев с нативными формулами и текстом между ними.
function mathBlockParagraphs(latex, opts = {}) {
    const spacing = { before: 120, after: 120 };
    const parts = splitTextChunks(latex);
    if (!parts.length) {
        return [new Paragraph({ alignment: 'center', spacing, children: [latexToDocxMath(latex)] })];
    }
    const runs = [];
    for (const p of parts) {
        if (p.math != null) runs.push(latexToDocxMath(p.math));
        else runs.push(new TextRun({ text: p.text, italics: true }));
    }
    return [new Paragraph({ alignment: 'center', spacing, children: runs })];
}

// ──────────────────────────────────────────────
// Инлайн-парсер Markdown (текст + математика + форматирование)
// ──────────────────────────────────────────────

// Инлайн-формулы \( ... \), \[ ... \] и $ ... $ внутри текста.
const INLINE_MATH_RE = /(\\\[[\s\S]+?\\\]|\\\([\s\S]+?\\\)|\$\$[\s\S]+?\$\$|\$[^$\n]+?\$)/;

function splitMath(text) {
    const parts = [];
    const re = new RegExp(INLINE_MATH_RE.source, 'g');
    let last = 0, m;
    while ((m = re.exec(text)) !== null) {
        if (m.index > last) parts.push({ type: 'text', value: text.slice(last, m.index) });
        const s = m[0];
        let v;
        if (s.charAt(0) === '\\') v = s.slice(2, -2);
        else if (s.startsWith('$$')) v = s.slice(2, -2);
        else v = s.slice(1, -1);
        parts.push({ type: 'math', value: v.trim() });
        last = re.lastIndex;
    }
    if (last < text.length) parts.push({ type: 'text', value: text.slice(last) });
    return parts;
}

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
    // Сначала вырезаем формулы: внутри \( \) могут встречаться звёздочки
    // и подчёркивания, не относящиеся к Markdown-форматированию.
    for (const seg of splitMath(text)) {
        if (seg.type === 'math') { runs.push(latexToDocxMath(seg.value)); continue; }
        // Локальная копия регулярки с собственной lastIndex — иначе рекурсивные
        // вызовы «съедают» позицию общего глобального объекта и цикл зависает.
        const re = new RegExp(INLINE_TOKEN_RE.source, 'g');
        let last = 0, m;
        while ((m = re.exec(seg.value)) !== null) {
            if (m.index > last) {
                const plain = seg.value.slice(last, m.index);
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
            last = re.lastIndex;
        }
        if (last < seg.value.length) {
            const plain = seg.value.slice(last);
            if (plain) runs.push(new TextRun({ text: plain, ...fmt }));
        }
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
