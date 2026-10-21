// md2docx.js — полный пайплайн Markdown -> .docx «как в ридере».
//
// Использует ровно те же модули и настройки, что и web-ридер:
//   markdown-it + markdown-it-texmath + KaTeX (плагины добавляет вызывающий код).
// Поэтому всё, что понимает ридер в формулах $...$, понимает и экспорт.
//
// Улучшения относительно pandoc:
//   1) Формулы не теряются: LaTeX извлекается из KaTeX-вывода (annotation x-tex),
//      т.е. используются правила распознавания texmath (dollars/brackets/$$/\\[\\]),
//      а затем конвертируются в нативные формулы Word (OMML) — редактируемые!
//   2) Pandoc часто «не видел» инлайн-формулы в $...$ внутри текста с русскими
//      словами/пробелами — здесь той проблемы нет, т.к. парсер тот же, что в ридере.
//   3) Собственный OMML-рендерер: дроби, индексы, пределы сумм/интегралов,
//      радикалы, матрицы, случаи, скобки авторазмера, акценты, \text и т.д.

'use strict';

const { convertHtmlToDocxBuffer } = require('./docx-exporter');

let _md = null;

function getMarkdownIt() {
    if (_md) return _md;
    const markdownit = require('markdown-it');
    const texmath = require('markdown-it-texmath');
    const katex = require('katex');

    const md = markdownit({ html: true, linkify: true, typographer: false });

    // плагины, доступные в проекте (совпадают с набором ридера)
    const optional = [
        ['markdown-it-multimd-table', { multiline: true, rowspan: true, headerless: true }],
        ['markdown-it-task-lists', { enabled: true, label: true }],
        ['markdown-it-footnote', undefined],
        ['markdown-it-mark', undefined],
        ['markdown-it-sub', undefined],
        ['markdown-it-sup', undefined],
    ];
    for (const [name, opts] of optional) {
        try {
            let plugin = require(name);
            if (plugin && plugin.default) plugin = plugin.default;
            md.use(plugin, opts);
        } catch (e) { /* плагин не установлен — пропускаем */ }
    }

    md.use(texmath, {
        engine: katex,
        delimiters: ['dollars', 'brackets'],
        katexOptions: { macros: {}, throwOnError: false },
    });

    _md = md;
    return md;
}

// markdown (string) -> Buffer (.docx)
function convertMarkdownToDocx(markdown, options = {}) {
    const md = options.markdownIt || getMarkdownIt();
    const html = md.render(markdown);
    return convertHtmlToDocxBuffer(html);
}

module.exports = { convertMarkdownToDocx, getMarkdownIt };
