const { convertMarkdownToDocx } = require('../exporter/md2docx');
const fs = require('fs');

const md = `# Заголовок H1

Обычный абзац с **жирным**, *курсивом*, ~~зачёркнутым~~, ==маркером== и \`кодом\`.

## Математика

Инлайн: энергия $E = mc^2$, а также $\\alpha + \\beta \\leq \\gamma$ в тексте на русском $x_i^2$.

Дисплейная:

$$ \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a} $$

Сумма с пределами: $\\sum_{i=1}^{n} i^2 = \\frac{n(n+1)(2n+1)}{6}$

Интеграл: $\\int_0^\\infty e^{-x^2} dx = \\frac{\\sqrt{\\pi}}{2}$

Матрица:

$$ A = \\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix} $$

Случаи: $f(x) = \\begin{cases} x^2 & x \\geq 0 \\\\ -x^2 & x < 0 \\end{cases}$

Акценты и стрелки: $\\hat{x}, \\vec{v}, \\tilde{y}, \\bar{z}, a \\to b, A \\subseteq B$

\\[ \\lim_{x \\to 0} \\frac{\\sin x}{x} = 1 \\]

Таблица:

| Предмет | Балл | Формула |
|---------|------|---------|
| Матан   | 5    | $e^{i\\pi}+1=0$ |
| Физика  | 4    | $F = ma$ |

- Маркированный список
- Второй пункт
  - Вложенный
1. Нумерованный
2. Второй

> Цитата_block с формулой $\\sigma = \\sqrt{\\frac{1}{N}\\sum_i (x_i - \\mu)^2}$

- [ ] Задача не сделана
- [x] Задача сделана

\`\`\`python
def f(x):
    return x**2
\`\`\`

---

[^1]: Сноска.
`;

const buf = convertMarkdownToDocx(md);
fs.writeFileSync('test/out.docx', buf);
console.log('DOCX written:', buf.length, 'bytes');
