#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert.py — единая точка входа md → docx (улучшенный аналог `pandoc x.md -o x.docx`).

Конвейер:
    1) node export/md_ast.js file.md            → JSON-AST + MathML для формул
       (markdown-it + texmath + KaTeX — тот же стек, что у ридера)
    2) python export/md2docx.py ast.json out.docx
       (AST → Word; MathML → OMML — нативные редактируемые формулы Word)

Запуск:
    python3 convert.py input.md [output.docx]
"""

import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))


def main():
    if len(sys.argv) < 2:
        print("использование: python3 convert.py input.md [output.docx]", file=sys.stderr)
        sys.exit(1)
    md_path = os.path.abspath(sys.argv[1])
    if not os.path.isfile(md_path):
        print(f"файл не найден: {md_path}", file=sys.stderr)
        sys.exit(1)
    docx_path = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else \
        os.path.splitext(md_path)[0] + ".docx"

    node_bin = os.environ.get("NODE_BIN", "node")
    py_bin = sys.executable or "python3"

    with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as tf:
        ast_path = tf.name
    try:
        # этап 1: markdown → AST (+MathML)
        r1 = subprocess.run([node_bin, os.path.join(HERE, "md_ast.js"), md_path],
                            capture_output=True, text=True)
        if r1.returncode != 0:
            print("Этап 1 (разбор Markdown) завершился с ошибкой:\n" + r1.stderr,
                  file=sys.stderr)
            sys.exit(r1.returncode)
        with open(ast_path, "w", encoding="utf-8") as f:
            f.write(r1.stdout)

        # проверка валидности JSON раньше, чем запустим python-этап
        try:
            json.loads(r1.stdout)
        except Exception as e:
            print(f"Этап 1 вернул некорректный JSON: {e}", file=sys.stderr)
            sys.exit(2)

        # этап 2: AST → docx
        r2 = subprocess.run([py_bin, os.path.join(HERE, "md2docx.py"),
                             ast_path, docx_path,
                             "--base-dir", os.path.dirname(md_path)],
                            capture_output=True, text=True)
        if r2.returncode != 0:
            print("Этап 2 (сборка DOCX) завершился с ошибкой:\n" + r2.stderr,
                  file=sys.stderr)
            sys.exit(r2.returncode)
        if r2.stdout:
            print(r2.stdout.strip())
        n_math = r1.stdout.count('"mathml":')
        print(f"Формул обработано: {n_math}")
    finally:
        try:
            os.unlink(ast_path)
        except OSError:
            pass


if __name__ == "__main__":
    main()
