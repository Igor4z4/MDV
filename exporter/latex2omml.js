// latex2omml.js — конвертер LaTeX-математики в Office Math Markup Language (OMML).
// Идея вдохновлена pandoc (LaTeX -> TeXMath -> OMML), но реализация написана заново:
// без внешних бинарников, работает прямо внутри Electron.
//
// Поддерживается:
//  - дроби \frac/\dfrac/\tfrac/\binom (включая вложенные)
//  - индексы/верхние индексы, пределы \sum_{i=1}^{n}, \int_a^b, \lim_{x\to0}
//  - радикалы \sqrt[n]{...}
//  - скобки \left( ... \right), \langle/\rceil и т.п.
//  - матрицы и случаи: matrix/pmatrix/bmatrix/vmatrix/Bmatrix/Vmatrix/cases
//  - греческие буквы, операторы (\cdot, \times, \leq ...), функции (sin, log...)
//  - \text{...}, \mathbf/\mathrm/\mathbb и пр., акценты \hat \bar \vec \tilde
//  - штрихи (') -> ′, юникод-формулы (α, ∑, ≤) проходят как есть
//
// Выход: строка XML <m:oMath>...</m:oMath> для document.xml документа Word.

'use strict';

const esc = (s) => String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

// один «run» формулы; sty: i|p|b (italic/plain/bold)
function run(text, opts = {}) {
    let rPr = '';
    if (opts.sty) rPr += `<m:rPr><m:sty m:val="${opts.sty}"/></m:rPr>`;
    return `<m:r>${rPr}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;
}

// ------------------------------------------------------------------ tables ---
const SYMS = {
    alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ϵ', varepsilon: 'ε',
    zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ',
    lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', varpi: 'ϖ', rho: 'ρ',
    varrho: 'ϱ', sigma: 'σ', varsigma: 'ς', tau: 'τ', upsilon: 'υ', phi: 'φ',
    varphi: 'ϕ', chi: 'χ', psi: 'ψ', omega: 'ω',
    Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π',
    Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
    hbar: 'ℏ', ell: 'ℓ', partial: '∂', infty: '∞', nabla: '∇', forall: '∀',
    exists: '∃', neg: '¬', emptyset: '∅', varnothing: '∅', ldots: '…', cdots: '⋯',
    dots: '…', vdots: '⋮', ddots: '⋱', prime: '′', angle: '∠', triangle: '△',
    square: '□', surd: '√', degree: '°',
    leq: '≤', le: '≤', geq: '≥', ge: '≥', neq: '≠', ne: '≠', approx: '≈',
    equiv: '≡', sim: '∼', simeq: '≃', cong: '≅', propto: '∝', ll: '≪', gg: '≫',
    times: '×', div: '÷', pm: '±', mp: '∓', cdot: '⋅', ast: '∗', star: '⋆',
    circ: '∘', bullet: '∙', oplus: '⊕', ominus: '⊖', otimes: '⊗', oslash: '⊘',
    odot: '⊙', cap: '∩', cup: '∪', subset: '⊂', supset: '⊃',
    subseteq: '⊆', supseteq: '⊇', notin: '∉', ni: '∋', mid: '∣',
    parallel: '∥', perp: '⊥', wedge: '∧', land: '∧', vee: '∨', lor: '∨',
    top: '⊤', therefore: '∴', because: '∵',
    rightarrow: '→', to: '→', leftarrow: '←', leftrightarrow: '↔', Rightarrow: '⇒',
    Leftarrow: '⇐', Leftrightarrow: '⇔', mapsto: '↦', implies: '⟹', iff: '⟺',
    longrightarrow: '⟶', longleftarrow: '⟵',
    langle: '⟨', rangle: '⟩', lfloor: '⌊', rfloor: '⌋', lceil: '⌈', rceil: '⌉',
    setminus: '∖', dotplus: '∔',
    aleph: 'ℵ', Re: 'ℜ', Im: 'ℑ', wp: '℘',
    checkmark: '✓', dagger: '†',
    quad: ' ', qquad: '  ', space: ' ',
};

const NARY = {
    sum: '∑', prod: '∏', coprod: '∐', int: '∫', iint: '∬', iiint: '∭', oint: '∮',
    bigcup: '⋃', bigcap: '⋂', bigsqcup: '⨆', bigoplus: '⨁', bigotimes: '⨂',
    bigodot: '⨀', lim: 'lim', max: 'max', min: 'min', sup: 'sup', inf: 'inf',
};

const FUNCS = new Set(['sin','cos','tan','cot','sec','csc','arcsin','arccos','arctan',
    'sinh','cosh','tanh','coth','log','ln','lg','exp','det','dim','gcd','deg',
    'arg','ker','Pr']);

const DELIMS = {
    '(': '(', ')': ')', '[': '[', ']': ']', '{': '{', '}': '}',
    '.': '', '|': '|', '/': '/',
};
const DELIM_CMDS = {
    '\\{': '{', '\\}': '}', '\\langle': '⟨', '\\rangle': '⟩',
    '\\lfloor': '⌊', '\\rfloor': '⌋', '\\lceil': '⌈', '\\rceil': '⌉',
    '\\vert': '|', '\\Vert': '‖', '\\|': '‖', '\\backslash': '\\',
};

// --------------------------------------------------------------- tokenizer ---
function tokenize(src) {
    const toks = [];
    let i = 0;
    const n = src.length;
    while (i < n) {
        const c = src[i];
        if (c === '\\') {
            let j = i + 1;
            if (j < n && /[a-zA-Z]/.test(src[j])) {
                while (j < n && /[a-zA-Z]/.test(src[j])) j++;
                toks.push({ t: 'cmd', v: src.slice(i, j) });
                if (src[j] === ' ') j++; // съедает один пробел-разделитель
                i = j;
            } else if (j < n) {
                toks.push({ t: 'cmd', v: '\\' + src[j] });
                i = j + 1;
            } else { i++; }
        } else if (c === '{') { toks.push({ t: 'open' }); i++; }
        else if (c === '}') { toks.push({ t: 'close' }); i++; }
        else if (c === '^') { toks.push({ t: 'sup' }); i++; }
        else if (c === '_') { toks.push({ t: 'sub' }); i++; }
        else if (c === '&') { toks.push({ t: 'amp' }); i++; }
        else if (c === '\n') { toks.push({ t: 'nl' }); i++; }
        else if (c === "'") { toks.push({ t: 'prime' }); i++; }
        else { toks.push({ t: 'chr', v: c }); i++; }
    }
    return toks;
}

// ------------------------------------------------------------- parser (AST) --
class Parser {
    constructor(toks) { this.toks = toks; this.i = 0; }
    peek() { return this.toks[this.i]; }
    next() { return this.toks[this.i++]; }
    eof() { return this.i >= this.toks.length; }

    parseGroupRaw() {
        this.next(); // consume {
        const items = [];
        while (!this.eof() && this.peek().t !== 'close') {
            const save = this.i;
            const it = this.parseAtom();
            if (it) items.push(it);
            if (this.i === save) this.next(); // защита от зацикливания
        }
        if (!this.eof()) this.next(); // consume }
        return { k: 'row', items };
    }

    parseArg() {
        const tk = this.peek();
        if (!tk) return { k: 'row', items: [] };
        if (tk.t === 'open') return this.parseGroupRaw();
        if (tk.t === 'close') { this.next(); return { k: 'row', items: [] }; }
        this.next();
        if (tk.t === 'cmd') { const a = this.cmdToAtom(tk.v); return a ? { k: 'row', items: [a] } : { k: 'row', items: [] }; }
        if (tk.t === 'prime') return { k: 'row', items: [{ k: 'sym', v: '′' }] };
        if (tk.t === 'chr') return { k: 'row', items: [mkChar(tk.v)] };
        return { k: 'row', items: [] };
    }

    parseAtom() {
        const base = this.parseSimple();
        if (!base) return null;
        let sup = null, sub = null;
        while (true) {
            const tk = this.peek();
            if (tk && tk.t === 'sup') { this.next(); sup = appendScript(sup, this.parseArg()); }
            else if (tk && tk.t === 'sub') { this.next(); sub = appendScript(sub, this.parseArg()); }
            else break;
        }
        if (sup || sub) return { k: 'supsub', base, sup, sub };
        return base;
    }

    parseSimple() {
        const tk = this.peek();
        if (!tk) return null;
        if (tk.t === 'chr') {
            if (tk.v === ' ' || tk.v === '\t') { this.next(); return null; }
            this.next();
            return mkChar(tk.v);
        }
        if (tk.t === 'prime') { this.next(); return { k: 'sym', v: '′' }; }
        if (tk.t === 'open') return this.parseGroupRaw();
        if (tk.t === 'close') { this.next(); return null; }
        if (tk.t === 'nl') { this.next(); return null; }
        if (tk.t === 'sup' || tk.t === 'sub') { this.next(); return null; }
        if (tk.t === 'amp') { this.next(); return { k: 'sep' }; }
        if (tk.t === 'cmd') { this.next(); return this.cmdToAtom(tk.v); }
        return null;
    }

    cmdToAtom(name) {
        const p = this;
        const bare = name.replace(/^\\/, '');
        switch (name) {
            case '\\frac': case '\\dfrac': case '\\tfrac': {
                const a = p.parseArg(), b = p.parseArg();
                return { k: 'frac', num: a, den: b };
            }
            case '\\binom': case '\\dbinom': case '\\tbinom': {
                const a = p.parseArg(), b = p.parseArg();
                return { k: 'delim', open: '(', close: ')', body: { k: 'frac', num: a, den: b, noBar: true } };
            }
            case '\\sqrt': {
                let idx = null;
                if (p.peek() && p.peek().t === 'chr' && p.peek().v === '[') {
                    p.next();
                    const items = [];
                    while (!p.eof() && !(p.peek().t === 'chr' && p.peek().v === ']')) {
                        const save = p.i;
                        const a = p.parseAtom();
                        if (a) items.push(a);
                        if (p.i === save) p.next();
                    }
                    if (!p.eof()) p.next();
                    idx = { k: 'row', items };
                }
                const body = p.parseArg();
                return { k: 'sqrt', idx, body };
            }
            case '\\text': case '\\textrm': case '\\textnormal': case '\\mbox':
            case '\\textsf': case '\\texttt': case '\\textmd': case '\\textup': {
                const g = p.parseArg();
                return { k: 'text', v: rowToPlain(g) };
            }
            case '\\textbf': case '\\mathbf': case '\\bm': case '\\boldsymbol': {
                const g = p.parseArg(); return { k: 'styled', sty: 'b', body: g };
            }
            case '\\textit': case '\\mathit': {
                const g = p.parseArg(); return { k: 'styled', sty: 'i', body: g };
            }
            case '\\mathrm': case '\\operatorname': {
                const g = p.parseArg(); return { k: 'styled', sty: 'p', body: g };
            }
            case '\\mathbb': { const g = p.parseArg(); return { k: 'styled', sty: 'bb', body: g }; }
            case '\\mathcal': { const g = p.parseArg(); return { k: 'styled', sty: 'cal', body: g }; }
            case '\\mathfrak': { const g = p.parseArg(); return { k: 'styled', sty: 'frak', body: g }; }
            case '\\hat': case '\\widehat': { const g = p.parseArg(); return { k: 'accent', chr: '\u0302', body: g }; }
            case '\\check': { const g = p.parseArg(); return { k: 'accent', chr: '\u030C', body: g }; }
            case '\\grave': { const g = p.parseArg(); return { k: 'accent', chr: '\u0300', body: g }; }
            case '\\acute': { const g = p.parseArg(); return { k: 'accent', chr: '\u0301', body: g }; }
            case '\\dot': { const g = p.parseArg(); return { k: 'accent', chr: '\u0307', body: g }; }
            case '\\ddot': { const g = p.parseArg(); return { k: 'accent', chr: '\u0308', body: g }; }
            case '\\tilde': case '\\widetilde': { const g = p.parseArg(); return { k: 'accent', chr: '\u0303', body: g }; }
            case '\\vec': case '\\overrightarrow': { const g = p.parseArg(); return { k: 'accent', chr: '\u20D7', body: g }; }
            case '\\bar': { const g = p.parseArg(); return { k: 'accent', chr: '\u0304', body: g }; }
            case '\\overline': { const g = p.parseArg(); return { k: 'barline', over: true, body: g }; }
            case '\\underline': { const g = p.parseArg(); return { k: 'barline', over: false, body: g }; }
            case '\\left': {
                const d1 = p.readDelimSpec();
                const items = [];
                while (!p.eof() && !(p.peek().t === 'cmd' && p.peek().v === '\\right')) {
                    const save = p.i;
                    const a = p.parseAtom();
                    if (a) items.push(a);
                    if (p.i === save) p.next();
                }
                let d2 = '';
                if (!p.eof()) { p.next(); d2 = p.readDelimSpec(); }
                return { k: 'delim', open: d1, close: d2, body: { k: 'row', items } };
            }
            case '\\begin': {
                const envRow = p.parseArg();
                const env = rowToPlain(envRow).trim();
                return p.parseEnvironment(env);
            }
            case '\\end': { p.parseArg(); return null; }
            case '\\limits': case '\\nolimits': case '\\displaystyle': case '\\textstyle':
            case '\\scriptstyle': case '\\scriptscriptstyle': case '\\!': case '\\,':
            case '\\;': case '\\:': return null;
            case '\\\\': return { k: 'nlbr' };
            case '~': return { k: 'sym', v: ' ' };
            case '%': return { k: 'op', v: '%' };
            case '$': return { k: 'op', v: '$' };
            case '#': return { k: 'op', v: '#' };
            case '&': return { k: 'op', v: '&' };
            case '_': return { k: 'op', v: '_' };
            case '\\': return { k: 'nlbr' };
        }
        if (SYMS[bare] !== undefined) {
            const v = SYMS[bare];
            if (v.trim() === '') return null;
            return { k: isLetterish(v) ? 'idn' : 'op', v };
        }
        if (NARY[bare]) return { k: 'nary-marker', chr: NARY[bare], name: bare };
        if (FUNCS.has(bare)) return { k: 'func', name: bare };
        if (/^[a-zA-Z]$/.test(bare)) return { k: 'idn', v: bare };
        return { k: 'text', v: bare }; // неизвестная команда — имя текстом
    }

    readDelimSpec() {
        const tk = this.peek();
        if (!tk) return '';
        if (tk.t === 'cmd') {
            this.next();
            if (DELIM_CMDS[tk.v] !== undefined) return DELIM_CMDS[tk.v];
            const bare = tk.v.replace(/^\\/, '');
            if (SYMS[bare] !== undefined) return SYMS[bare];
            return bare;
        }
        if (tk.t === 'chr') { this.next(); return DELIMS[tk.v] !== undefined ? DELIMS[tk.v] : tk.v; }
        this.next();
        return '';
    }

    parseEnvironment(env) {
        const rows = [[]];
        let cur = [];
        const pushCell = () => { rows[rows.length - 1].push({ k: 'row', items: cur }); cur = []; };
        const pushRow = () => { pushCell(); rows.push([]); };
        while (!this.eof()) {
            if (this.peek().t === 'cmd' && this.peek().v === '\\end') {
                this.next();
                this.parseArg();
                break;
            }
            const save = this.i;
            const a = this.parseAtom();
            if (this.i === save) { this.next(); continue; }
            if (a === null) continue;
            if (a.k === 'sep') pushCell();
            else if (a.k === 'nlbr') pushRow();
            else cur.push(a);
        }
        pushCell();
        if (rows.length > 1) {
            const last = rows[rows.length - 1];
            if (last.length === 1 && last[0].items.length === 0) rows.pop();
        }
        let open = '', close = '';
        switch (env) {
            case 'pmatrix': open = '('; close = ')'; break;
            case 'bmatrix': open = '['; close = ']'; break;
            case 'Bmatrix': open = '{'; close = '}'; break;
            case 'vmatrix': case 'Vmatrix': open = '|'; close = '|'; break;
            case 'cases': open = '{'; close = ''; break;
        }
        return { k: 'matrix', env, rows, open, close };
    }
}

function appendScript(prev, add) {
    if (!prev) return add;
    const items = [];
    const flatten = (node) => {
        if (node && node.k === 'row') items.push(...node.items);
        else if (node) items.push(node);
    };
    flatten(prev); flatten(add);
    return { k: 'row', items };
}

function isLetterish(ch) {
    return /[A-Za-zα-ωΑ-ΩͰ-Ͽ]/.test(ch);
}

function mkChar(c) {
    if (/[A-Za-z0-9]/.test(c)) return { k: 'idn', v: c };
    if (/[+\-=<>*/!,;:?]/.test(c)) return { k: 'op', v: c };
    if (/[()[\]|/.{}]/.test(c)) return { k: 'delimchar', v: c };
    return { k: 'sym', v: c }; // в т.ч. юникод-символы
}

function rowToPlain(row) {
    if (!row) return '';
    if (row.k === 'row') return row.items.map(atomToPlain).join('');
    return atomToPlain(row);
}
function atomToPlain(a) {
    if (!a) return '';
    switch (a.k) {
        case 'idn': case 'sym': case 'op': case 'delimchar': return a.v;
        case 'text': return a.v;
        case 'row': return rowToPlain(a);
        default: return '';
    }
}

// пост-обработка: nary-marker + supsub -> nary с пределами
function fixNary(node) {
    if (!node) return node;
    if (node.k === 'row') {
        const out = [];
        for (const it0 of node.items) {
            let it = it0;
            if (it && it.k === 'supsub' && it.base && it.base.k === 'nary-marker') {
                it = { k: 'nary', chr: it.base.chr, name: it.base.name, sub: it.sub, sup: it.sup };
            } else if (it && it.k === 'nary-marker') {
                it = { k: 'nary', chr: it.chr, name: it.name, sub: null, sup: null };
            }
            out.push(fixNary(it));
        }
        return { k: 'row', items: out };
    }
    if (node.k === 'supsub') {
        return { ...node, base: fixNary(node.base), sup: fixNary(node.sup), sub: fixNary(node.sub) };
    }
    for (const key of ['num', 'den', 'body', 'idx']) if (node[key]) node[key] = fixNary(node[key]);
    if (node.rows) node.rows = node.rows.map(r => r.map(fixNary));
    return node;
}

// ---------------------------------------------------------------- renderer ---
function renderNode(node) {
    if (!node) return '';
    switch (node.k) {
        case 'row': return node.items.map(renderNode).join('');
        case 'idn': return run(node.v, { sty: 'i' });
        case 'sym': case 'op': case 'delimchar': return run(node.v);
        case 'text': return run(node.v, { sty: 'p' });
        case 'func': return run(node.name, { sty: 'p' });
        case 'supsub': {
            const base = renderNode(node.base);
            if (node.sup && node.sub) {
                return `<m:sSubSup><m:sSubSupPr><m:ctrlPr/></m:sSubSupPr><m:e>${base}</m:e><m:sub>${renderNode(node.sub)}</m:sub><m:sup>${renderNode(node.sup)}</m:sup></m:sSubSup>`;
            }
            if (node.sup) return `<m:sSup><m:sSupPr><m:ctrlPr/></m:sSupPr><m:e>${base}</m:e><m:sup>${renderNode(node.sup)}</m:sup></m:sSup>`;
            return `<m:sSub><m:sSubPr><m:ctrlPr/></m:sSubPr><m:e>${base}</m:e><m:sub>${renderNode(node.sub)}</m:sub></m:sSub>`;
        }
        case 'frac': {
            const bar = node.noBar ? '<m:type m:val="noBar"/>' : '';
            return `<m:f><m:fPr>${bar}<m:ctrlPr/></m:fPr><m:num>${renderNode(node.num)}</m:num><m:den>${renderNode(node.den)}</m:den></m:f>`;
        }
        case 'sqrt': {
            if (node.idx) {
                return `<m:rad><m:radPr><m:degHide m:val="off"/><m:ctrlPr/></m:radPr><m:deg>${renderNode(node.idx)}</m:deg><m:e>${renderNode(node.body)}</m:e></m:rad>`;
            }
            return `<m:rad><m:radPr><m:degHide m:val="on"/><m:ctrlPr/></m:radPr><m:deg/><m:e>${renderNode(node.body)}</m:e></m:rad>`;
        }
        case 'accent': {
            return `<m:acc><m:accPr><m:chr m:val="${esc(node.chr)}"/><m:ctrlPr/></m:accPr><m:e>${renderNode(node.body)}</m:e></m:acc>`;
        }
        case 'barline': {
            const pos = node.over ? 'top' : 'bot';
            return `<m:bar><m:barPr><m:pos m:val="${pos}"/><m:ctrlPr/></m:barPr><m:e>${renderNode(node.body)}</m:e></m:bar>`;
        }
        case 'delim': {
            return `<m:d><m:dPr><m:begChr m:val="${esc(node.open)}"/><m:endChr m:val="${esc(node.close)}"/><m:ctrlPr/></m:dPr><m:e>${renderNode(node.body)}</m:e></m:d>`;
        }
        case 'nary': {
            if (!node.sub && !node.sup) return run(node.chr);
            const sub = node.sub ? renderNode(node.sub) : '';
            const sup = node.sup ? renderNode(node.sup) : '';
            const hideSub = node.sub ? 'off' : 'on';
            const hideSup = node.sup ? 'off' : 'on';
            const isIntegral = /^(int|iint|iiint|oint)$/.test(node.name || '');
            const limLoc = isIntegral ? 'subSup' : 'undOvr';
            return `<m:nary><m:naryPr><m:chr m:val="${esc(node.chr)}"/><m:limLoc m:val="${limLoc}"/><m:subHide m:val="${hideSub}"/><m:supHide m:val="${hideSup}"/><m:ctrlPr/></m:naryPr><m:e>${run(node.chr)}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:nary>`;
        }
        case 'styled': return renderStyled(node.body, node.sty);
        case 'matrix': {
            const cols = Math.max(...node.rows.map(r => r.length));
            const tblRows = node.rows.map(r =>
                '<m:tr>' + r.map(cell => `<m:tc>${renderNode(cell)}</m:tc>`).join('') + '</m:tr>'
            ).join('');
            const mats = `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="${cols}"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>${tblRows}</m:m>`;
            if (node.open || node.close) {
                return `<m:d><m:dPr><m:begChr m:val="${esc(node.open)}"/><m:endChr m:val="${esc(node.close)}"/><m:ctrlPr/></m:dPr><m:e>${mats}</m:e></m:d>`;
            }
            return mats;
        }
        default: return '';
    }
}

function renderStyled(body, sty) {
    const xml = renderNode(body);
    let msty = 'p';
    if (sty === 'b' || sty === 'bb') msty = 'b';
    else if (sty === 'i' || sty === 'cal' || sty === 'frak') msty = 'i';
    let out = xml.replace(/<m:sty m:val="[^"]*"/g, `<m:sty m:val="${msty}"`);
    out = out.replace(/<m:r>/g, `<m:r><m:rPr><m:sty m:val="${msty}"/></m:rPr>`);
    return out;
}

// публичный API -----------------------------------------------------------
function latexToOmml(latex) {
    try {
        const toks = tokenize(latex);
        const p = new Parser(toks);
        const items = [];
        while (!p.eof()) {
            const save = p.i;
            const a = p.parseAtom();
            if (a) items.push(a);
            if (p.i === save) p.next();
        }
        const ast = fixNary({ k: 'row', items });
        const inner = renderNode(ast);
        if (!inner.trim()) return '';
        return `<m:oMath>${inner}</m:oMath>`;
    } catch (e) {
        // деградация: отдаём исходник текстом, чтобы экспорт не падал
        return `<m:oMath>${run(latex)}</m:oMath>`;
    }
}

module.exports = { latexToOmml };
