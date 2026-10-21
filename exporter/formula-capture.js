// formula-capture.js — извлекает исходный LaTeX из HTML, который markdown-it +
// texmath + KaTeX уже отрендерили во «представлении» ридера.
//
// Почему это важно: именно этот слой делает экспорт «ридер-совместимым».
// Что умеет наш рендер (dollars/brackets/$$/\[\]/\(\\), escaped \$, "не математика"
// внутри $...$ с пробелами по правилам texmath) — то же самое попадает в .docx,
// потому что мы берём формулы из того же самого KaTeX-вывода:
//   <span class="katex"><annotation encoding="application/x-tex">...LaTeX...</annotation>
// Плюс поддерживаем <img class="mathml-image" alt="LaTeX"> (fallback texmath без KaTeX).

'use strict';

const TEX_ATTRS = ['data-tex', 'data-latex', 'data-formula'];

// Разбирает DOM-узел формулы и возвращает { latex, display } или null.
// Порядок поиска: аннотация x-tex -> явные data-атрибуты -> alt у mathml-image.
function captureFromNode(node) {
    if (!node) return null;
    const cls = node.className || '';

    // 1) любой узел может содержать KaTeX-аннотацию (<annotation encoding="application/x-tex">)
    //    — она есть у katex-html и у MathML-обёртки texmath (.mathjax .texmath).
    const latex = findAnnotation(node);
    if (latex !== null && latex.trim() !== '') {
        const display = /katex-display|display-mode/.test(cls)
            || !!(node.attrs && node.attrs.display === 'block');
        return { latex, display };
    }

    // 2) явные атрибуты data-tex
    if (node.attrs) {
        for (const a of TEX_ATTRS) {
            if (node.attrs[a]) {
                return { latex: node.attrs[a], display: /display/.test(cls) };
            }
        }
    }

    // 3) <img class="mathml-image" alt="...latex..."> (fallback texmath без KaTeX)
    if (node.tag === 'img' && /mathml-image/.test(cls)) {
        const alt = node.attrs && node.attrs.alt;
        if (alt) return { latex: alt, display: false };
    }

    // 4) голый контейнер .katex без annotation — берём текст как есть
    if (/(^|\s)katex(\s|$)/.test(cls) || /katex-display/.test(cls)) {
        const t = textOf(node).trim();
        if (t) return { latex: t, display: /katex-display/.test(cls) };
    }
    return null;
}

function findAnnotation(node) {
    // обход в ширину: annotation — всегда неглубоко
    const queue = [node];
    while (queue.length) {
        const cur = queue.shift();
        if (cur.tag === 'annotation') {
            const enc = cur.attrs && cur.attrs.encoding;
            if (!enc || /x-tex|tex/i.test(enc)) return textOf(cur);
        }
        if (cur.children) queue.push(...cur.children);
    }
    return null;
}

function textOf(node) {
    if (!node) return '';
    if (node.text !== undefined) return node.text;
    let s = '';
    for (const c of node.children || []) s += textOf(c);
    return s;
}

module.exports = { captureFromNode, textOf };
