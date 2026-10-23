'use strict';
/**
 * Экспорт Markdown → DOCX.
 *
 * Этап 1: перенос текста и формул.
 *  - Формулы распознаются теми же правилами (texmath + dollars/brackets), что и
 *    при предпросмотре, поэтому поведение совпадает с текущим рендером.
 *  - Дополнительно поддерживаются границы «$…$» (инлайн) и «$$…$$» (блочные),
 *    включая многострочные блочные формулы и экранированные \$ — см. mathSplit.js.
 *  - LaTeX конвертируется в OMML (нативные формулы Word) через KaTeX (MathML → XSLT MML2OMML).
 *    Если конвертация не удалась — формула вставляется как текст (с сохранением границ $),
 *    чтобы содержимое не терялось.
 *
 * Остальные элементы (таблицы, списки, код, картинки) будут добавлены позже.
 */

const path = require('path');
const fs = require('fs');
const {
    Document, Packer, Paragraph, TextRun, HeadingLevel,
    Math: DocxMath, MathRun, ExternalHyperlink, UnderlineType,
} = require('docx');
const markdownit = require('markdown-it');
const texmath = require('./web/lib/texmath.js');
const katex = require('katex');

// ---------------------------------------------------------------------------
// Разбиение текста на «обычный текст» и «формулы» (границы $$, $, \[ \], \( \))
// ---------------------------------------------------------------------------

// Символы, после которых закрывающая граница формулы недопустима (цифры — защита от
// вида «$5 и $10»), зеркалит логику texmath.$_post.
function isBadAfter(code) {
    return code >= 0x30 && code <= 0x39; // '0'..'9'
}
// Символ перед открывающей границей: не допустимы обратный слэш (экранирование) и цифра.
function isBadBefore(code) {
    return code === 0x5c || (code >= 0x30 && code <= 0x39);
}

/**
 * Найти закрывающую границу len-символьной последовательности '$', начиная со start.
 * Пустые формулы запрещены (как в texmath: rex /\$((?:[^\s\\])|(?:\S.*?[^\s\\]))\$/).
 */
function findDollarClose(s, start, len, displayMode) {
    let i = start;
    const n = s.length;
    while (i < n) {
        if (s[i] === '\\') { i += 2; continue; }            // экранированный символ внутри формулы
        if (s.startsWith('$'.repeat(len), i)) {
            // для одиночной $ нельзя склеивать с соседними $ ($$ — это отдельная граница)
            if (len === 1 && (s[i + 1] === '$' || s[i - 1] === '$')) { i++; continue; }
            const inner = s.slice(start, i);
            if (!inner.length) { i += len; continue; }       // пустая формула не считается
            if (displayMode) return i;                       // $$ … $$ — многострочные допускаются
            if (!/\n/.test(inner)) return i;                 // инлайн $ … $ — без переносов строк
            const after = i + len < n ? s.charCodeAt(i + len) : false;
            if (after === false || !isBadAfter(after)) return i;
            i += len;
            continue;
        }
        if (!displayMode && s[i] === '\n') return -1;         // инлайн-формула не может быть многострочной
        i++;
    }
    return -1;
}

/**
 * Разбить строку на сегменты: {type:'text'|'math', ...}.
 * Поддерживаются: $$…$$ (блок), $…$ (инлайн), \[…\] (блок), \(…\) (инлайн).
 * Проверяется «чистота» содержимого (нет ли внутри незакрытой другой границы),
 * чтобы не схватить ложные пары вроде «$a$ и $b$».
 */
function splitMath(s) {
    const out = [];
    let buf = '';
    let i = 0;
    const n = s.length;
    const pushText = (t) => { buf += t; };
    const flush = () => { if (buf) { out.push({ type: 'text', text: buf }); buf = ''; } };

    while (i < n) {
        const ch = s[i];

        // \$ — экранированный доллар, это обычный текст
        if (ch === '\\' && s[i + 1] === '$') { pushText('$'); i += 2; continue; }

        // $$…$$ — блочная формула (может быть многострочной)
        if (ch === '$' && s[i + 1] === '$') {
            const prev = i > 0 ? s.charCodeAt(i - 1) : false;
            if (prev !== false && isBadBefore(prev)) { pushText('$$'); i += 2; continue; }
            const close = findDollarClose(s, i + 2, 2, true);
            if (close !== -1) {
                const inner = s.slice(i + 2, close);
                // содержимое не должно содержать незакрытую одинарную $-границу
                if (!/(^|[^\\])\$/.test(inner.replace(/\\\$/g, ''))) {
                    flush();
                    out.push({ type: 'math', tex: inner.trim(), display: true });
                    i = close + 2;
                    continue;
                }
            }
            pushText('$$'); i += 2; continue;
        }

        // $…$ — инлайн формула
        if (ch === '$') {
            const prev = i > 0 ? s.charCodeAt(i - 1) : false;
            if (prev !== false && isBadBefore(prev)) { pushText('$'); i++; continue; }
            const close = findDollarClose(s, i + 1, 1, false);
            if (close !== -1) {
                flush();
                out.push({ type: 'math', tex: s.slice(i + 1, close).trim(), display: false });
                i = close + 1;
                continue;
            }
            pushText('$'); i++; continue;
        }

        // \[ … \ ] — блочная формула
        if (ch === '\\' && s[i + 1] === '[') {
            const close = s.indexOf('\\]', i + 2);
            if (close !== -1) {
                flush();
                out.push({ type: 'math', tex: s.slice(i + 2, close).trim(), display: true });
                i = close + 2;
                continue;
            }
            pushText('\\['); i += 2; continue;
        }

        // \( … \ ) — инлайн формула
        if (ch === '\\' && s[i + 1] === '(') {
            const close = s.indexOf('\\)', i + 2);
            if (close !== -1) {
                flush();
                out.push({ type: 'math', tex: s.slice(i + 2, close).trim(), display: false });
                i = close + 2;
                continue;
            }
            pushText('\\('); i += 2; continue;
        }

        pushText(ch);
        i++;
    }
    flush();
    return out;
}

// ---------------------------------------------------------------------------
// LaTeX → OMML (нативные формулы Word)
// ---------------------------------------------------------------------------

// XSLT-преобразование MathML→OMML из дистрибутива pandoc (MIT лицензия).
// Ищем по нескольким возможным путям (разработка / упаковка electron-builder).
function findMml2Omml() {
    const candidates = [
        process.env.MMDV_MML2OMML_XSL,
        path.join(__dirname, 'vendor', 'mml2omml.xsl'),
        path.join(process.resourcesPath || '', 'vendor', 'mml2omml.xsl'),
        path.join(__dirname, '..', 'vendor', 'mml2omml.xsl'),
    ].filter(Boolean);
    for (const p of candidates) {
        try { if (fs.existsSync(p)) return fs.readFileSync(p, 'utf-8'); } catch (_) { /* noop */ }
    }
    throw new Error('Не найден файл преобразования vendor/mml2omml.xsl (LaTeX→OMML).');
}

let _xsltProcessor = null;
function getLatexToOmml() {
    if (_xsltProcessor) return _xsltProcessor;
    const { DOMParser, XMLSerializer } = require('xmldom-xslt').dom;
    const { XSLT } = require('xmldom-xslt');
    const xslText = findMml2Omml();
    const xslDoc = new DOMParser().parseFromString(xslText, 'text/xml');
    _xsltProcessor = new XSLT(xslDoc, {});
    return _xsltProcessor;
}

/**
 * LaTeX → строка OMML (<m:oMath>…). Возвращает null, если конвертация невозможна.
 */
function latexToOmml(tex, displayMode) {
    try {
        const html = katex.renderToString(tex, {
            output: 'mathml',
            throwOnError: false,
            strict: 'ignore',
            displayMode: !!displayMode,
        });
        // KaTeX с output:'mathml' возвращает <span class="katex"><math …>…</math></span>
        const mStart = html.indexOf('<math');
        const mEnd = html.lastIndexOf('</math>');
        if (mStart === -1 || mEnd === -1) return null;
        const mathml = html.slice(mStart, mEnd + '</math>'.length);
        // проверка ошибок парсинга KaTeX (throwOnError:false вставляет span.error)
        if (/katex-error/.test(html)) return null;

        const { DOMParser } = require('xmldom-xslt').dom;
        const doc = new DOMParser().parseFromString(mathml, 'text/xml');
        const xslt = getLatexToOmml();
        const result = xslt.transform(doc);
        const xml = new (require('xmldom-xslt').dom.XMLSerializer)().serializeToString(result);
        const omStart = xml.indexOf('<m:oMath');
        if (omStart === -1) return null;
        let omml = xml.slice(omStart);
        // обрезаем всё после закрывающего тега oMath (могут быть служебные пробросы)
        const tailName = /<m:oMath(?:Part)?\b/.test(omml.slice(0, 30)) && omml.includes('</m:oMathPara>')
            ? '</m:oMathPara>' : '</m:oMath>';
        const omEnd = omml.lastIndexOf(tailName);
        if (omEnd !== -1) omml = omml.slice(0, omEnd + tailName.length);
        return omml;
    } catch (err) {
        console.warn('latexToOmml failed:', err.message);
        return null;
    }
}

// ---------------------------------------------------------------------------
// Инлайновые элементы docx (текст с форматированием + формулы)
// ---------------------------------------------------------------------------

/**
 * Сегменты KaTeX-ошибочных формул показываем текстом с сохранением границ.
 */
function mathAsTextRun(m) {
    const delim = m.display ? '$$' : '$';
    return new TextRun({ children: [`${delim}${m.tex}${delim}`], italics: true });
}

function buildMathElement(m) {
    const omml = latexToOmml(m.tex, m.display);
    if (!omml) return mathAsTextRun(m);
    try {
        // docx умеет вставлять готовый OMML через Math({ child: new MathRun(...) })?
        // Нет: MathRun принимает только текст. Используем нижний уровень — Math c CustomXmlComponent
        // невозможно, поэтому применяем Math.fromOMML? Библиотека docx не предоставляет прямой
        // загрузки OMML — используем собственный компонент через RawXml (см. rawXml.js).
        return rawMath(omml, m);
    } catch (err) {
        console.warn('buildMathElement fallback:', err.message);
        return mathAsTextRun(m);
    }
}

// ---------------------------------------------------------------------------
// Raw OMML-компоненты (docx-библиотека не имеет публичного API для готового OMML)
// ---------------------------------------------------------------------------

const { XmlComponent, EmptyElement } = require('docx/dist/components/xml-component');
const { XmlAttributeComponent } = require('docx/dist/components/xml-attribute-component');

class OMath extends XmlComponent {
    constructor(ommlXml) {
        super('m:oMath');
        this.ommlXml = ommlXml;
    }
    // docx сериализует корневой тег сам; подставляем внутренность как есть через root
    prepForXml(converter) {
        const inner = this.ommlXml
            .replace(/^<m:oMath(\s[^>]*)?>/, '')
            .replace(/<\/m:oMath>\s*$/, '');
        return { [this.root.key]: [{ type: 'raw', value: inner }] };
    }
}

class OMathPara extends XmlComponent {
    constructor(omath) {
        super('m:oMathPara');
        this.omath = omath;
        this.root.push(omath);
    }
}

// ---------------------------------------------------------------------------
// Простейшие inline-стили (жирный/курсив/код/ссылки) — базовый перенос текста
// ---------------------------------------------------------------------------

function styleFromToken(token, base) {
    const s = Object.assign({}, base);
    switch (token.type) {
        case 'strong_open': case 'strong_close': s.bold = token.type === 'strong_open' ? true : undefined; break;
        case 'em_open': case 'em_close': s.italics = token.type === 'em_open' ? true : undefined; break;
        case 's_open': case 's_close': s.strike = token.type === 's_open' ? true : undefined; break;
        case 'code_inline': s.font = 'Consolas'; break;
        case 'mark_open': s.highlight = 'yellow'; break;
        case 'sub_open': s.subScript = true; break;
        case 'sup_open': s.superScript = true; break;
    }
    return s;
}

// ---------------------------------------------------------------------------
// Сборка документа: markdown → блоки docx (параграфы с текстом и формулами)
// ---------------------------------------------------------------------------

function makeMd() {
    const md = markdownit({ html: false, linkify: true, typographer: false });
    // Рендер формул нам не нужен — важна разметка токенов (math_*), стили — свои.
    md.use(texmath, { engine: katex, delimiters: ['dollars', 'brackets'], katexOptions: { macros: {}, throwOnError: false } });
    return md;
}

/**
 * Обработать inline-токены markdown-it в массив элементов docx (TextRun/Math).
 * Для формульных токенов content содержит чистый LaTeX (правила texmath уже
 * корректно нашли границы $/$$), поэтому «уверенный порт» опирается на тот же
 * механизм, а splitMath() используется как резервный путь и для plain-текста.
 */
function inlineTokensToElements(tokens, md) {
    const elements = [];
    const styleStack = [];
    let linkHref = null;
    let inLink = false;

    const curStyle = () => styleStack.length ? styleStack[styleStack.length - 1] : {};

    const addStyledText = (text, extra = {}) => {
        // Внутри текста тоже ищем формулы с границами $ (защита: markdown-it мог
        // не распознать формулу из-за особенностей правил, например «$$» в начале строки)
        for (const seg of splitMath(text)) {
            if (seg.type === 'math') {
                elements.push(buildInlineMath(seg, curStyle()));
            } else {
                const opts = Object.assign({}, curStyle(), extra);
                if (inLink) {
                    opts.style = 'Hyperlink';
                }
                if (seg.text) elements.push(new TextRun(Object.assign({ text: seg.text }, opts)));
            }
        }
    };

    const buildInlineMath = (m, style) => {
        const el = buildMathElement(m);
        return el;
    };

    for (let idx = 0; idx < tokens.length; idx++) {
        const token = tokens[idx];
        const t = token.type;

        if (t === 'text' || t === 'link_open' === false && t === 'image') { /* handled below */ }

        if (t === 'text') {
            addStyledText(token.content);
        } else if (t === 'code_inline') {
            const opts = Object.assign({}, curStyle(), { font: 'Consolas' });
            elements.push(new TextRun(Object.assign({ text: token.content }, opts)));
        } else if (t === 'math_inline' || t === 'math_inline_double' || t === 'math_block' || t === 'math_block_eqno') {
            const display = (t === 'math_block' || t === 'math_block_eqno' || t === 'math_inline_double');
            addStyledPlain(MathRenderInto(token.content, display));
        } else if (t === 'softbreak' || t === 'hardbreak') {
            elements.push(new TextRun({ text: '' , break: 1}));
        } else if (['_open', '_close'].some(() => false)) {
            // noop
        } else if (t.endsWith('_open')) {
            const name = t.slice(0, -'_open'.length);
            const st = Object.assign({}, curStyle());
            if (name === 'strong') st.bold = true;
            else if (name === 'em') st.italics = true;
            else if (name === 's' || name === 'del' || name === 'strike') st.strike = true;
            else if (name === 'mark') st.highlight = 'yellow';
            else if (name === 'sub') st.subScript = true;
            else if (name === 'sup') st.superScript = true;
            else if (name === 'ins') st.underline = { type: UnderlineType.SINGLE };
            styleStack.push(st);
            if (name === 'link') { linkHref = token.attrs && token.attrs.find(a => a[0] === 'href')?.[1] || null; inLink = true; }
        } else if (t.endsWith('_close')) {
            const name = t.slice(0, -'_close'.length);
            if (name === 'link') { inLink = false; linkHref = null; }
            styleStack.pop();
        } else if (t === 'image') {
            // картинки — позже (этап 2+); пока вставляем alt-текст
            addStyledText(`[🖼 ${token.content || ''}]`);
        } else if (t === 'footnote_ref') {
            // сноски — позже
        }
    }

    function addStyledPlain(elems) {
        for (const e of elems) elements.push(e);
    }

    return elements;
}

/** Формульный токен → элемент docx */
function MathRenderInto(tex, display) {
    const m = { type: 'math', tex, display };
    const el = buildMathElement(m);
    return [el];
}

// ---------------------------------------------------------------------------
// Блочный разбор
// ---------------------------------------------------------------------------

function tokensToParagraphs(tokens) {
    const children = [];
    let listDepth = 0;
    let orderedList = false;

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        switch (token.type) {
            case 'heading_open': {
                const level = parseInt(token.tag.slice(1), 10) || 1;
                const inline = tokens[i + 1];
                const elems = inline && inline.type === 'inline' ? inlineTokensToElements(inline.children) : [];
                children.push(new Paragraph({
                    heading: [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3,
                        HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6][level - 1],
                    children: elems,
                }));
                i += 2;
                break;
            }
            case 'paragraph_open': {
                const inline = tokens[i + 1];
                const elems = inline && inline.type === 'inline' ? inlineTokensToElements(inline.children) : [];
                // Блочная формула одна в абзаце — центрируем (как display-математика)
                const onlyMath = elems.length === 1 && elems[0] instanceof OMath;
                children.push(new Paragraph({
                    children: elems.length ? elems : [],
                    alignment: onlyMath ? 'center' : undefined,
                    spacing: { after: 120 },
                }));
                i += 2;
                break;
            }
            case 'blockquote_open': {
                // соберём вложенные токены до close и оформим как цитату
                let depth = 1, j = i + 1;
                const inner = [];
                while (j < tokens.length && depth > 0) {
                    if (tokens[j].type === 'blockquote_open') depth++;
                    else if (tokens[j].type === 'blockquote_close') { depth--; if (!depth) break; }
                    inner.push(tokens[j]);
                    j++;
                }
                const quoted = tokensToParagraphs(inner);
                for (const q of quoted) {
                    // docx Paragraph неизменяемый; пересоздаём с отступом/курсивом проще:
                    children.push(q);
                }
                i = j;
                break;
            }
            case 'bullet_list_open': listDepth++; orderedList = false; break;
            case 'bullet_list_close': listDepth--; break;
            case 'ordered_list_open': listDepth++; orderedList = true; break;
            case 'ordered_list_close': listDepth--; break;
            case 'list_item_open': {
                // элемент списка: берём его inline-дети (первый paragraph)
                let j = i + 1;
                let depth = 1;
                const inner = [];
                while (j < tokens.length && depth > 0) {
                    if (tokens[j].type === 'list_item_open') depth++;
                    else if (tokens[j].type === 'list_item_close') { depth--; if (!depth) break; }
                    inner.push(tokens[j]);
                    j++;
                }
                const paras = tokensToParagraphs(inner);
                children.push(...paras.map(p => p));
                i = j;
                break;
            }
            case 'fence':
            case 'code_block': {
                // листинги кода — этап 2; сейчас переносим моноширинным текстом
                const lines = (token.content || '').split('\n');
                for (const line of lines) {
                    children.push(new Paragraph({
                        children: [new TextRun({ text: line, font: 'Consolas', size: 20 })],
                        spacing: { after: 0 },
                    }));
                }
                break;
            }
            case 'hr': {
                children.push(new Paragraph({ themeBorder: true, border: { bottom: { style: 'single', size: 6, color: '999999' } }, children: [] }));
                break;
            }
            case 'table_open': {
                // таблицы — этап 2; сейчас переносим содержимое текстом, чтобы не терять данные
                let depth = 1, j = i + 1;
                const inner = [];
                while (j < tokens.length && depth > 0) {
                    if (tokens[j].type === 'table_open') depth++;
                    else if (tokens[j].type === 'table_close') { depth--; if (!depth) break; }
                    inner.push(tokens[j]);
                    j++;
                }
                const rowElems = [];
                for (const tk of inner) {
                    if (tk.type === 'inline') rowElems.push(...inlineTokensToElements(tk.children));
                    else if (tk.type === 'text') rowElems.push(new TextRun({ text: tk.content }));
                }
                children.push(new Paragraph({ children: rowElems, spacing: { after: 120 } }));
                i = j;
                break;
            }
            default:
                break;
        }
    }
    return children;
}

// ---------------------------------------------------------------------------
// Публичный API
// ---------------------------------------------------------------------------

async function exportMarkdownToDocx(markdownText) {
    // Как и в просмотрщике: нормализуем \[ на начало строки
    const src = String(markdownText).replace(/\\\[/g, '\n\\[');
    const md = makeMd();
    const tokens = md.parse(src, {});
    const children = tokensToParagraphs(tokens);

    const doc = new Document({
        creator: 'MDV Markdown Viewer',
        title: 'Экспорт из Markdown',
        sections: [{ properties: {}, children }],
    });
    return await Packer.toBuffer(doc);
}

module.exports = { exportMarkdownToDocx, splitMath, latexToOmml };
