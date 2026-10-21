#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2docx.py — этап 2 конвейера: JSON-AST (от md_ast.js) → .docx.

Идея наследует архитектуру pandoc (markdown → AST → writer docx), но реализация
написана заново и с двумя ключевыми улучшениями:

1. ФОРМУЛЫ. pandoc при порте $...$ часто роняет формулы в обычный текст или
   требует MathJax/LaTeX-pipeline. Здесь KaTeX компилирует LaTeX → MathML,
   а собственный конвертер MathML → OMML (Office Math Markup Language) вставляет
   в документ нативные формулы Word — их можно редактировать двойным щелчком,
   они корректно выглядят в MS Word, LibreOffice, Google Docs.
   Поддерживаются: дроби, индексы, корни, суммы/интегралы (nary), пределы,
   акценты (hat/bar/vec), матрицы, скобки mrow, text-режим и др.

2. ТОЧНОСТЬ РАЗБОРА. Парсер — тот же markdown-it + texmath, что и в самом
   ридере, поэтому поведение $...$ идентично просмотру (включая граничные
   случаи: «$100 и $200» не считается формулой).

Покрытие AST: заголовки H1–H6, абзацы, жирный/курсив/зачёркнутый/подчёркнутый,
==маркер==, подстрочные/надстрочные, инлайн-код и блоки кода (с фоном),
цитаты (с левой границей), нумерованные/маркированные списки любой вложенности,
task-lists (☐/☑), таблицы со выравниванием по колонкам, горизонтальные линии,
изображения (локальные и http/https — скачиваются), ссылки (hyperlink + стиль),
сноски (real Word footnotes w:footnote), формулы inline и display ($$…$$).

Запуск: python3 md2docx.py ast.json output.docx --base-dir DIR
"""

import argparse
import json
import os
import re
import sys
import urllib.request
from lxml import etree as ET

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor, Inches, Emu
except ImportError:
    sys.stderr.write("Не найден python-docx. Установите: pip install python-docx\n")
    raise

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSM = {"m": M, "w": W}


OMML_NSMAP = {"m": M, "w": W}


def mel(tag):
    return ET.Element(f"{{{M}}}{tag}", nsmap=OMML_NSMAP)


def wel(tag):
    return ET.Element(f"{{{W}}}{tag}", nsmap=OMML_NSMAP)


def localname(tag):
    return tag.split("}")[-1] if "}" in tag else tag


# ---------------------------------------------------------------------------
# MathML → OMML
# ---------------------------------------------------------------------------

BIG_OPS = {
    "∑": "∑", "∏": "∏", "∫": "∫", "∬": "∬", "∭": "∭", "∮": "∮",
    "⋃": "⋃", "⋂": "⋂", "⨁": "⨁", "⨂": "⨂", "lim": "lim",
    "max": "max", "min": "min", "sup": "sup", "inf": "inf",
    "det": "det", "gcd": "gcd", "deg": "deg",
}

ACCENT_MAP = {
    "^": "hat", "ˆ": "hat", "‾": "bar", "¯": "bar", "~": "tilde",
    "→": "arrow", "´": "acute", "`": "grave", "ˇ": "caron", "˘": "breve",
    "⇀": "arrow", "⃗": "arrow", ".": "dot", "..": "ddot", "...": "ddd",
}


class MathMLToOMML:
    """Конвертирует дерево MathML (ElementTree) в OMML <m:oMath>."""

    def convert(self, math_root):
        omath = mel("oMath")
        # KaTeX кладёт всё в <semantics><mrow>…</mrow><annotation/>
        sem = self._find_first(math_root, "semantics")
        body = sem if sem is not None else math_root
        inner = self._find_first(body, "mrow")
        src = inner if inner is not None and len(list(inner)) else body
        for child in list(src):
            if isinstance(child.tag, str):
                self._node(child, omath)
        return omath

    @staticmethod
    def _find_first(node, name):
        for c in node.iter():
            if isinstance(c.tag, str) and localname(c.tag) == name:
                return c
        return None

    # -- runs ----------------------------------------------------------
    def _run(self, parent, text, sty=None):
        r = mel("r")
        if sty:
            rPr = mel("rPr")
            st = mel(sty)
            rPr.append(st)
            r.append(rPr)
        t = mel("t")
        t.text = text
        r.append(t)
        parent.append(r)

    def _kids_of(self, node):
        return [c for c in node if isinstance(c.tag, str)
                and localname(c.tag) != "annotation"]

    def _fill(self, container, node):
        """Наполняет OMML-контейнер содержимым MathML-узла."""
        kids = self._kids_of(node)
        if kids:
            for k in kids:
                self._node(k, container)
        else:
            txt = "".join(node.itertext())
            if txt:
                self._run(container, txt, self._style_for(node, txt))

    @staticmethod
    def _style_for(node, txt):
        ln = localname(node.tag)
        if ln == "mi" and len(txt.strip()) > 1 and txt.strip().isalpha():
            return "normal"      # «sin», «mod» — прямым, не курсивом
        if ln == "mtext":
            return "normal"
        return None

    # -- dispatcher ------------------------------------------------------
    def _node(self, node, parent):
        ln = localname(node.tag)
        method = getattr(self, "_m_" + ln, None)
        if method:
            method(node, parent)
        else:
            self._default(node, parent)

    def _default(self, node, parent):
        kids = self._kids_of(node)
        if kids:
            for k in kids:
                self._node(k, parent)
        else:
            txt = "".join(node.itertext())
            if txt:
                self._run(parent, txt, self._style_for(node, txt))

    def _m_math(self, node, parent):
        for k in self._kids_of(node):
            self._node(k, parent)

    def _m_semantics(self, node, parent):
        kids = self._kids_of(node)
        if kids:
            self._node(kids[0], parent)

    def _m_mrow(self, node, parent):
        kids = self._kids_of(node)
        # stretchy-скобки превращаем в m:d (нужно для \left( ... \right))
        if (len(kids) >= 2
                and not any(localname(k.tag) in ("mtable", "mtr") for k in kids)
                and localname(kids[0].tag) == "mo"
                and "".join(kids[0].itertext()).strip() in ("(", "[", "{", "|", "⟨", "⌊", "⌈")
                and localname(kids[-1].tag) == "mo"
                and "".join(kids[-1].itertext()).strip() in (")", "]", "}", "|", "⟩", "⌋", "⌉")):
            d = mel("d")
            sep_open = "".join(kids[0].itertext()).strip()
            sep_close = "".join(kids[-1].itertext()).strip()
            if sep_open != "(" or sep_close != ")":
                dPr = mel("dPr")
                begChr = mel("begChr"); begChr.set(f"{{{M}}}val", sep_open); dPr.append(begChr)
                endChr = mel("endChr"); endChr.set(f"{{{M}}}val", sep_close); dPr.append(endChr)
                d.append(dPr)
            e = mel("e")
            for k in kids[1:-1]:
                self._node(k, e)
            d.append(e)
            parent.append(d)
            return
        # special case: [ ( , mtable , ) ] — скобки относятся к матрице
        j = 0
        while j < len(kids):
            k = kids[j]
            if localname(k.tag) == "mtable":
                beg = end = None
                if j > 0 and localname(kids[j - 1].tag) == "mo":
                    t = "".join(kids[j - 1].itertext()).strip()
                    if t in "([{":
                        beg = t
                if j + 1 < len(kids) and localname(kids[j + 1].tag) == "mo":
                    t = "".join(kids[j + 1].itertext()).strip()
                    if t in ")]}":
                        end = t
                d = mel("d")
                if beg and end and (beg, end) != ("(", ")"):
                    dPr = mel("dPr")
                    bc = mel("begChr"); bc.set(f"{{{M}}}val", beg); dPr.append(bc)
                    ec = mel("endChr"); ec.set(f"{{{M}}}val", end); dPr.append(ec)
                    d.append(dPr)
                e = mel("e")
                self._node(k, e)
                d.append(e)
                parent.append(d)
                j += 3 if (beg or end) else 1
                continue
            self._node(k, parent)
            j += 1

    def _m_mi(self, node, parent):
        txt = "".join(node.itertext())
        self._run(parent, txt, self._style_for(node, txt))

    def _m_mn(self, node, parent):
        self._run(parent, "".join(node.itertext()))

    def _m_mo(self, node, parent):
        txt = "".join(node.itertext())
        # символы-операторы KaTeX иногда содержит &ApplyFunction; и т.п. уже декодированными
        if txt.strip() in ("", ):
            return
        self._run(parent, txt)

    def _m_mtext(self, node, parent):
        self._run(parent, "".join(node.itertext()), "normal")

    def _m_mspace(self, node, parent):
        pass

    def _m_msub(self, node, parent):
        kids = self._kids_of(node)
        if len(kids) < 2:
            return self._default(node, parent)
        s = mel("sSub")
        e = mel("e"); self._fill(e, kids[0]); s.append(e)
        sub = mel("sub"); self._fill(sub, kids[1]); s.append(sub)
        parent.append(s)

    def _m_msup(self, node, parent):
        kids = self._kids_of(node)
        if len(kids) < 2:
            return self._default(node, parent)
        s = mel("sSup")
        e = mel("e"); self._fill(e, kids[0]); s.append(e)
        sup = mel("sup"); self._fill(sup, kids[1]); s.append(sup)
        parent.append(s)

    def _m_msubsup(self, node, parent):
        kids = self._kids_of(node)
        if len(kids) < 3:
            return self._default(node, parent)
        base = kids[0]
        btxt = "".join(base.itertext()).strip()
        if localname(base.tag) == "mo" and btxt in BIG_OPS:
            self._nary(parent, btxt, kids[1], kids[2], kids[3:])
            return
        s = mel("sSubSup")
        e = mel("e"); self._fill(e, base); s.append(e)
        sub = mel("sub"); self._fill(sub, kids[1]); s.append(sub)
        sup = mel("sup"); self._fill(sup, kids[2]); s.append(sup)
        parent.append(s)
        for extra in kids[3:]:
            self._node(extra, parent)

    def _m_munderover(self, node, parent):
        kids = self._kids_of(node)
        if len(kids) == 2:
            return self._m_munder(node, parent)
        if len(kids) < 2:
            return self._default(node, parent)
        base = kids[0]
        btxt = "".join(base.itertext()).strip()
        if localname(base.tag) == "mo" and btxt in BIG_OPS:
            self._nary(parent, btxt, kids[1], kids[2], kids[3:])
            return
        low = mel("limLow")
        e = mel("e"); self._fill(e, base); low.append(e)
        lim = mel("lim"); self._fill(lim, kids[1]); low.append(lim)
        parent.append(low)
        upp = mel("limUpp")
        e2 = mel("e"); self._fill(e2, base); upp.append(e2)
        lim2 = mel("lim"); self._fill(lim2, kids[2]); upp.append(lim2)
        parent.append(upp)

    def _m_munder(self, node, parent):
        kids = self._kids_of(node)
        if len(kids) < 2:
            return self._default(node, parent)
        btxt = "".join(kids[0].itertext()).strip()
        low = mel("limLow")
        e = mel("e"); self._fill(e, kids[0]); low.append(e)
        lim = mel("lim"); self._fill(lim, kids[1]); low.append(lim)
        parent.append(low)

    def _m_mover(self, node, parent):
        kids = self._kids_of(node)
        if len(kids) < 2:
            return self._default(node, parent)
        accent_txt = "".join(kids[1].itertext()).strip()
        acc_type = ACCENT_MAP.get(accent_txt)
        if acc_type:
            acc = mel("acc")
            accPr = mel("accPr")
            chr_ = mel("chr"); chr_.set(f"{{{M}}}val", accent_txt); accPr.append(chr_)
            acc.append(accPr)
            e = mel("e"); self._fill(e, kids[0]); acc.append(e)
            parent.append(acc)
        else:
            ow = mel("groupChr")
            grpPr = mel("groupChrPr")
            chr_ = mel("chr"); chr_.set(f"{{{M}}}val", accent_txt or "‾"); grpPr.append(chr_)
            pos = mel("pos"); pos.set(f"{{{M}}}val", "top"); grpPr.append(pos)
            ow.append(grpPr)
            e = mel("e"); self._fill(e, kids[0]); ow.append(e)
            parent.append(ow)

    def _nary(self, parent, chr_val, sub_node, sup_node, rest):
        nary = mel("nary")
        naryPr = mel("naryPr")
        chr_ = mel("chr"); chr_.set(f"{{{M}}}val", chr_val); naryPr.append(chr_)
        limLoc = mel("limLoc"); limLoc.set(f"{{{M}}}val", "undOvr"); naryPr.append(limLoc)
        hideSub = mel("hideSub"); hideSup = mel("hideSup")
        sub_txt = "".join(sub_node.itertext()).strip() if sub_node is not None else ""
        sup_txt = "".join(sup_node.itertext()).strip() if sup_node is not None else ""
        if not sub_txt:
            hideSub.set(f"{{{M}}}val", "1"); naryPr.append(hideSub)
        if not sup_txt:
            hideSup.set(f"{{{M}}}val", "1"); naryPr.append(hideSup)
        nary.append(naryPr)
        sub = mel("sub")
        if sub_node is not None:
            self._fill(sub, sub_node)
        nary.append(sub)
        sup = mel("sup")
        if sup_node is not None:
            self._fill(sup, sup_node)
        nary.append(sup)
        e = mel("e")
        for k in rest:
            self._node(k, e)
        nary.append(e)
        parent.append(nary)

    def _m_mfrac(self, node, parent):
        kids = self._kids_of(node)
        if len(kids) < 2:
            return self._default(node, parent)
        f = mel("f")
        num = mel("num"); self._fill(num, kids[0]); f.append(num)
        den = mel("den"); self._fill(den, kids[1]); f.append(den)
        parent.append(f)

    def _m_msqrt(self, node, parent):
        rad = mel("rad")
        radPr = mel("radPr")
        degHide = mel("degHide"); degHide.set(f"{{{M}}}val", "1"); radPr.append(degHide)
        rad.append(radPr)
        deg = mel("deg"); rad.append(deg)
        e = mel("e"); self._fill(e, node); rad.append(e)
        parent.append(rad)

    def _m_mroot(self, node, parent):
        kids = self._kids_of(node)
        if len(kids) < 2:
            return self._default(node, parent)
        rad = mel("rad")
        deg = mel("deg"); self._fill(deg, kids[1]); rad.append(deg)
        e = mel("e"); self._fill(e, kids[0]); rad.append(e)
        parent.append(rad)

    def _m_mtable(self, node, parent):
        mt = mel("m")
        rows = [r for r in self._kids_of(node) if localname(r.tag) == "mtr"]
        ncols = max((len(self._kids_of(r)) for r in rows), default=0)
        for r in rows:
            mr = mel("mr")
            cells = [c for c in self._kids_of(r) if localname(c.tag) == "mtd"]
            for c in cells:
                e = mel("e"); self._fill(e, c); mr.append(e)
            for _ in range(ncols - len(cells)):
                mr.append(mel("e"))
            mt.append(mr)
        parent.append(mt)

    def _m_mmultiscripts(self, node, parent):
        kids = self._kids_of(node)
        if not kids:
            return
        i = 1
        while i + 1 < len(kids):
            a, b = kids[i], kids[i + 1]
            la, lb = localname(a.tag), localname(b.tag)
            if la == "mprescripts" or lb == "mprescripts":
                i += 1
                continue
            sss = mel("sSubSup")
            e = mel("e"); self._fill(e, kids[0]); sss.append(e)
            sub = mel("sub")
            if la != "none":
                self._fill(sub, a)
            sss.append(sub)
            sup = mel("sup")
            if lb != "none":
                self._fill(sup, b)
            sss.append(sup)
            parent.append(sss)
            i += 2

    def _m_mstyle(self, node, parent):
        self._default(node, parent)

    def _m_menclose(self, node, parent):
        d = mel("d")
        e = mel("e"); self._fill(e, node); d.append(e)
        parent.append(d)

    def _m_phantom(self, node, parent):
        pass


def mathml_str_to_omml(mathml_str):
    """Строка MathML → ElementTree OMML. Возвращает None при ошибке."""
    try:
        root = ET.fromstring(mathml_str)
    except ET.ParseError:
        return None
    try:
        conv = MathMLToOMML()
        return conv.convert(root)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Построение документа
# ---------------------------------------------------------------------------

CODE_FONT = "Consolas"
LINK_COLOR = RGBColor(0x11, 0x55, 0xAA)
LINK_UNDERLINE = True


class Builder:
    def __init__(self, base_dir="."):
        self.base_dir = base_dir
        self.doc = Document()
        self.footnotes = {}          # id -> индекс (1-based)
        self._fn_counter = 0
        self._setup_styles()
        self._setup_footnotes_part()

    # -- стили ------------------------------------------------------------
    def _setup_styles(self):
        normal = self.doc.styles["Normal"]
        normal.font.name = "Calibri"
        normal.font.size = Pt(11)
        normal.paragraph_format.space_after = Pt(6)
        normal.paragraph_format.line_spacing = 1.15

        code = self.doc.styles.add_style("CodeBlock", WD_STYLE_TYPE.PARAGRAPH)
        code.base_style = self.doc.styles["Normal"]
        code.font.name = CODE_FONT
        code.font.size = Pt(9.5)
        code.paragraph_format.left_indent = Pt(16)
        code.paragraph_format.right_indent = Pt(8)
        code.paragraph_format.space_after = Pt(0)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), "F4F4F4")
        code.element.get_or_add_pPr().append(shd)

    # -- настоящая часть сносок Word ---------------------------------------
    def _setup_footnotes_part(self):
        """python-docx не умеет сноски — создаём word/footnotes.xml вручную.

        Часть формируется один раз при __init__ (поэтому она гарантированно
        присутствует в rels и [Content_Types]); тела реальных сносок подставляются
        методом _rebuild_footnotes_blob() непосредственно перед сохранением.
        """
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.opc.part import Part
        from docx.opc.packuri import PackURI
        package = self.doc.part.package
        content = self._compose_footnotes_xml([])
        part = Part(
            PackURI("/word/footnotes.xml"),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            content, package,
        )
        self._footnotes_part = part
        self.doc.part.relate_to(part, RT.FOOTNOTES)

    @staticmethod
    def _compose_footnotes_xml(body_parts):
        sep = ('<w:footnote w:id="-1" w:type="separator">'
               '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
               '<w:r><w:footnoteRef/></w:r></w:p></w:footnote>'
               '<w:footnote w:id="-2" w:type="continuationSeparator">'
               '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
               '<w:r><w:footnoteContinuedReference/></w:r></w:p></w:footnote>')
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:footnotes xmlns:w="{W}" xmlns:m="{M}">{sep}' + "".join(body_parts) +
            "</w:footnotes>"
        ).encode("utf-8")

    def _next_fn_id(self):
        self._fn_counter += 1
        return self._fn_counter

    # -- публичный вход ------------------------------------------------------
    def build(self, tree):
        # сначала реестр сносок (порядок определений = порядок ссылок обычно)
        self._register_footnotes(tree)
        self._blocks(tree, self.doc)
        self._rebuild_footnotes_blob()

    def _register_footnotes(self, tree):
        defs = []
        self._collect_footnote_defs(tree, defs)
        for i, (_fid, _nodes) in enumerate(defs, 1):
            pass
        # id сопоставляются по порядку появления определений
        for order, (fid, nodes) in enumerate(defs, 1):
            self.footnotes[fid] = order
        self._fn_defs = defs

    def _collect_footnote_defs(self, nodes, out):
        i = 0
        while i < len(nodes):
            n = nodes[i]
            if n["type"] == "footnote_open":
                inner = n.get("children") or []
                fid = ""
                for a in n.get("attrs") or []:
                    if a[0] == "id":
                        fid = a[1]
                if not fid:
                    fid = n.get("markup", "") or n.get("content", "") or ""
                texts = []
                self._gather_inline_text(inner, texts)
                out.append((fid, inner))
                i += 1
                continue
            if n.get("children"):
                self._collect_footnote_defs(n["children"], out)
            i += 1

    def _gather_inline_text(self, nodes, out):
        for n in nodes:
            if n["type"] in ("text", "code_inline"):
                out.append(n["content"])
            elif n["type"] in ("math_inline", "math_block"):
                out.append("$" + (n.get("tex") or n.get("content") or "") + "$")
            elif n.get("children"):
                self._gather_inline_text(n["children"], out)

    def _rebuild_footnotes_blob(self):
        """Пересобираем байты word/footnotes.xml с реальными телами сносок.

        Тела сериализуются из временных Paragraph (python-docx/lxml), затем
        нормализуются префиксы пространств имён, чтобы валидный OOXML получился
        внутри <w:footnote>.
        """
        body_parts = []
        for order, (raw_id, inner_nodes) in enumerate(getattr(self, "_fn_defs", []), 1):
            tmp_doc = Document()
            para = tmp_doc.add_paragraph()
            self._inlines(para, inner_nodes, in_footnote=True)
            # стиль текста сноски
            pPr = para._p.get_or_add_pPr()
            xml = ET.tostring(para._p, encoding="unicode")
            # lxml уже использует корректные префиксы w:/m:, убираем лишние xmlns-объявления
            xml = re.sub(r'\sxmlns(?::\w+)?="[^"]*"', "", xml, count=0)
            body_parts.append(f'<w:footnote w:id="{order}">{xml}</w:footnote>')
        self._footnotes_part._blob = self._compose_footnotes_xml(body_parts)

    # -- блочный обход (дерево от md_ast.js: *_open содержат children) --------
    def _blocks(self, nodes, container):
        for n in nodes:
            t = n["type"]
            if t == "heading_open":
                lvl = int(re.sub(r"\D", "", n["tag"]) or 1)
                para = container.add_paragraph(style=f"Heading {min(lvl, 6)}")
                inline = self._first_child(n, "inline")
                if inline is not None:
                    self._inlines(para, inline.get("children") or [])
            elif t == "paragraph_open":
                para = container.add_paragraph()
                inline = self._first_child(n, "inline")
                if inline is not None:
                    self._inlines(para, inline.get("children") or [])
            elif t == "blockquote_open":
                self._blockquote(n.get("children") or [], container)
            elif t in ("bullet_list_open", "ordered_list_open"):
                self._list(n.get("children") or [], container,
                           t == "ordered_list_open", level=0)
            elif t in ("code_block", "fence"):
                self._code_block(n, container)
            elif t == "math_block":
                self._math_block(n, container)
            elif t == "hr":
                self._hr(container)
            elif t == "table_open":
                self._table(n.get("children") or [], container)
            elif t == "html_block":
                pass  # сырой HTML не экспортируем
            elif t == "footnote_block_open":
                pass  # определения уже собраны в реестр сносок
            # узлы "inline" верхнего уровня вне параграфов игнорируются

    @staticmethod
    def _first_child(node, type_):
        for c in node.get("children") or []:
            if c["type"] == type_:
                return c
        return None

    @staticmethod
    def _span(nodes, i, kind):
        """Возвращает детей контейнера kind, открытого в nodes[i], и индекс после close."""
        open_t, close_t = f"{kind}_open", f"{kind}_close"
        depth = 0
        start = i + 1
        j = i
        while j < len(nodes):
            if nodes[j]["type"] == open_t:
                depth += 1
                if depth == 1:
                    start = j + 1
            elif nodes[j]["type"] == close_t:
                depth -= 1
                if depth == 0:
                    return nodes[start:j], j + 1
            j += 1
        return nodes[start:], len(nodes)

    # -- цитаты -------------------------------------------------------------
    def _blockquote(self, inner, container):
        holder = []
        self._blocks(inner, _ParaCollector(holder))
        for p in holder:
            newp = container.add_paragraph()
            pf = newp._p.get_or_add_pPr()
            pbdr = OxmlElement("w:pBdr")
            left = OxmlElement("w:left")
            left.set(qn("w:val"), "single")
            left.set(qn("w:sz"), "18")
            left.set(qn("w:space"), "8")
            left.set(qn("w:color"), "9AA5B1")
            pbdr.append(left)
            pf.append(pbdr)
            ind = OxmlElement("w:ind")
            ind.set(qn("w:left"), "240")
            pf.append(ind)
            # копируем содержимое исходного параграфа целиком (runs, oMath...)
            for child in list(p._p):
                tag = localname(child.tag)
                if tag == "pPr":
                    continue
                newp._p.append(child)

    # -- списки --------------------------------------------------------------
    def _list(self, items, container, ordered, level):
        for n in items:
            if n["type"] == "list_item_open":
                self._list_item(n.get("children") or [], container, ordered, level)

    def _list_item(self, inner, container, ordered, level):
        first_para = True
        i = 0
        while i < len(inner):
            n = inner[i]
            t = n["type"]
            if t == "paragraph_open":
                para = container.add_paragraph()
                style = ("ListNumber" if ordered else "ListBullet")
                if level > 0:
                    style += f" {min(level + 1, 3) + 1}"
                try:
                    para.style = self.doc.styles[style]
                except KeyError:
                    pass
                kids = []
                inline = self._first_child(n, "inline")
                if inline is not None:
                    kids = list(inline.get("children") or [])
                task = self._pop_task_checkbox(kids)
                if task is not None:
                    box = para.add_run("\u2611 " if task else "\u2610 ")
                    box.font.name = "Segoe UI Symbol"
                self._inlines(para, kids)
                i += 1
                first_para = False
                continue
            if t in ("bullet_list_open", "ordered_list_open"):
                sub_ordered = t == "ordered_list_open"
                sub_inner, j = self._span(inner, i, "list")
                self._list(sub_inner, container, sub_ordered, level + 1)
                i = j
                continue
            if t == "math_block":
                self._math_block(n, container)
                i += 1
                continue
            if t == "blockquote_open":
                sub, j = self._span(inner, i, "blockquote")
                self._blockquote(sub, container)
                i = j
                continue
            i += 1

    @staticmethod
    def _pop_task_checkbox(kids):
        """task-lists вставляет <input type=checkbox>; либо текст '[ ] '/'[x] '."""
        if not kids:
            return None
        first = kids[0]
        if first["type"] == "html_inline" and "checkbox" in (first.get("content") or ""):
            checked = "checked" in first["content"]
            kids.pop(0)
            if kids and kids[0]["type"] == "text":
                kids[0]["content"] = kids[0]["content"].lstrip()
            return checked
        if first["type"] == "text":
            c = first["content"]
            for mark, val in (("[ ] ", False), ("[x] ", True), ("[X] ", True)):
                if c.startswith(mark):
                    kids[0] = dict(first)
                    kids[0]["content"] = c[len(mark):]
                    return val
        return None

    # -- код -------------------------------------------------------------------
    def _code_block(self, n, container):
        info = (n.get("info") or "").strip().lower()
        if info.startswith("mermaid"):
            para = container.add_paragraph()
            r = para.add_run("[диаграмма Mermaid: экспорт в .docx не поддерживается]")
            r.font.italic = True
            r.font.color.rgb = RGBColor(0xB0, 0x60, 0x00)
            return
        content = n.get("content") or ""
        lines = content.split("\n")
        if lines and lines[-1] == "":
            lines.pop()
        for line in lines:
            para = container.add_paragraph(style="CodeBlock")
            r = para.add_run(line if line else " ")
            r.font.name = CODE_FONT
            r.font.size = Pt(9.5)

    # -- формулы ---------------------------------------------------------------
    def _math_block(self, n, container):
        para = container.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if not self._insert_math(para, n, display=True):
            r = para.add_run("$$" + (n.get("tex") or n.get("content") or "") + "$$")
            r.font.italic = True

    def _insert_math(self, para, node, display=False):
        mathml = node.get("mathml")
        if not mathml:
            return False
        omml = mathml_str_to_omml(mathml)
        if omml is None:
            return False
        if display:
            oMathPara = mel("oMathPara")
            oMathPara.append(omml)
            para._p.append(oMathPara)
        else:
            para._p.append(omml)
        return True

    # -- hr --------------------------------------------------------------------
    def _hr(self, container):
        para = container.add_paragraph()
        pPr = para._p.get_or_add_pPr()
        pbdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "8")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "auto")
        pbdr.append(bottom)
        pPr.append(pbdr)

    # -- таблицы -----------------------------------------------------------------
    def _table(self, inner, container):
        rows = []
        self._collect_table_rows(inner, rows)
        if not rows:
            return
        ncols = max(len(r) for r in rows)
        table = container.add_table(rows=len(rows), cols=ncols)
        try:
            table.style = "Table Grid"
        except KeyError:
            pass
        header_flags = [any(c["header"] for c in r) for r in rows]
        for ri, row in enumerate(rows):
            tr = table.rows[ri]._tr
            if header_flags[ri]:
                trPr = tr.get_or_add_trPr()
                tblHeader = OxmlElement("w:tblHeader")
                tblHeader.set(qn("w:val"), "true")
                trPr.append(tblHeader)
            for ci in range(ncols):
                cell = table.cell(ri, ci)
                cell.paragraphs[0].text = ""
                para = cell.paragraphs[0]
                para.paragraph_format.space_after = Pt(2)
                if ci >= len(row):
                    continue
                cd = row[ci]
                self._inlines(para, cd["nodes"])
                align = cd["align"]
                if align == "center":
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif align == "right":
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                if cd["header"]:
                    for r in para.runs:
                        r.bold = True
                    tcPr = cell._tc.get_or_add_tcPr()
                    shd = OxmlElement("w:shd")
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:fill"), "EAF0F6")
                    tcPr.append(shd)

    def _collect_table_rows(self, nodes, rows):
        for n in nodes:
            if n["type"] == "tr_open":
                rows.append(self._extract_row(n))
            elif n.get("children"):
                self._collect_table_rows(n["children"], rows)

    def _extract_row(self, tr_node):
        row = []
        for cn in tr_node.get("children") or []:
            if cn["type"] in ("td_open", "th_open"):
                nodes = []
                for x in cn.get("children") or []:
                    if x["type"] == "inline":
                        nodes = x.get("children") or []
                    elif x["type"] == "math_block":
                        nodes = [x]
                attrs = dict(cn.get("attrs") or [])
                align = ""
                style = attrs.get("style", "")
                m = re.search(r"text-align:\s*(left|center|right)", style or "")
                if m:
                    align = m.group(1)
                row.append({
                    "nodes": nodes,
                    "header": cn["type"] == "th_open",
                    "align": align,
                    "colspan": int(attrs.get("colspan", "1") or 1),
                    "rowspan": int(attrs.get("rowspan", "1") or 1),
                })
        return row

    # -- изображения ---------------------------------------------------------------
    def _add_image(self, para, src, alt):
        path_ = None
        try:
            if re.match(r"^https?://", src, re.I):
                req = urllib.request.Request(src, headers={"User-Agent": "md2docx/1.0"})
                data = urllib.request.urlopen(req, timeout=20).read()
                suffix = os.path.splitext(src.split("?")[0])[1] or ".png"
                tmp = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   "_tmp_img" + suffix)
                with open(tmp, "wb") as f:
                    f.write(data)
                path_ = tmp
            else:
                cand = os.path.normpath(os.path.join(self.base_dir, src))
                if os.path.isfile(cand):
                    path_ = cand
            if path_:
                run = para.add_run()
                run.add_picture(path_, width=Inches(5.5))
                return
        except Exception:
            pass
        r = para.add_run(f"[изображение недоступно: {alt or src}]")
        r.font.color.rgb = RGBColor(0xB0, 0x00, 0x00)
        r.font.italic = True

    # -- инлайны ------------------------------------------------------------------------
    def _inlines(self, para, nodes, in_footnote=False):
        fmt = []           # стек форматирования (cur)
        link_buf = None    # не None => внутри ссылки: накапливаем (узел, копия стека)
        image_pending = None   # [src, alt_parts]

        for n in nodes:
            t = n["type"]
            cur = {}
            for f in fmt:
                cur.update(f)

            if t == "text":
                if image_pending is not None:
                    image_pending[1].append(n["content"])
                elif link_buf is not None:
                    link_buf.append(("text", n["content"], dict(cur)))
                else:
                    self._runs(para, n["content"], cur)
            elif t == "softbreak":
                if image_pending is None and link_buf is None:
                    para.add_run().add_break()
            elif t == "hardbreak":
                if image_pending is None and link_buf is None:
                    para.add_run().add_break(WD_BREAK.LINE)
            elif t == "code_inline":
                if image_pending is None:
                    if link_buf is not None:
                        link_buf.append(("code", n["content"], dict(cur)))
                    else:
                        r = para.add_run(n["content"])
                        r.font.name = CODE_FONT
                        r.font.size = Pt(10)
                        rPr = r._r.get_or_add_rPr()
                        shd = OxmlElement("w:shd")
                        shd.set(qn("w:val"), "clear")
                        shd.set(qn("w:fill"), "F0F0F0")
                        rPr.append(shd)
            elif t == "math_inline":
                if image_pending is None and link_buf is None:
                    if not self._insert_math(para, n):
                        r = para.add_run("$" + (n.get("tex") or n.get("content") or "") + "$")
                        r.font.italic = True
            elif t in ("strong_open", "em_open", "s_open", "del_open", "ins_open",
                       "u_open", "mark_open", "sub_open", "sup_open"):
                key = {
                    "strong_open": "bold", "em_open": "italic",
                    "s_open": "strike", "del_open": "strike", "ins_open": "underline",
                    "u_open": "underline", "mark_open": "highlight",
                    "sub_open": "sub", "sup_open": "sup",
                }[t]
                fmt.append({key: True})
            elif t in ("strong_close", "em_close", "s_close", "del_close", "ins_close",
                       "u_close", "mark_close", "sub_close", "sup_close"):
                if fmt:
                    fmt.pop()
            elif t == "link_open":
                attrs = dict(n.get("attrs") or [])
                if link_buf is None:
                    link_buf = []
                    link_href_val = attrs.get("href") or ""
            elif t == "link_close":
                if link_buf is not None:
                    self._emit_hyperlink(para, link_href_val, link_buf)
                    link_buf = None
            elif t == "image":
                attrs = dict(n.get("attrs") or [])
                alt_txt = attrs.get("alt", "") or ""
                if not alt_txt and n.get("children"):
                    buf = []
                    self._gather_inline_text(n["children"], buf)
                    alt_txt = "".join(buf)
                image_pending = [attrs.get("src", ""), [alt_txt]]
            elif t == "footnote_ref":
                attrs = dict(n.get("attrs") or [])
                fid = attrs.get("id", "") or n.get("markup", "") or n.get("content", "")
                idx = self.footnotes.get(fid)
                if idx is None:
                    # матчинг по порядковому номеру: fn1 -> id="fn1"
                    mnum = re.search(r"(\d+)", str(fid))
                    if mnum:
                        idx = self.footnotes.get("fn" + mnum.group(1))
                if idx is None:
                    idx = self._register_adhoc_footnote(str(order_hint := (len(self.footnotes) + 1)))
                self._add_real_footnote(para, idx)
            elif t == "html_inline":
                c = n.get("content") or ""
                if "<br" in c and image_pending is None and link_buf is None:
                    para.add_run().add_break(WD_BREAK.LINE)
            else:
                if n.get("children"):
                    self._inlines(para, n["children"], in_footnote)
        if image_pending is not None:
            self._add_image(para, image_pending[0], "".join(image_pending[1]).strip())

    def _emit_hyperlink(self, para, href, buf_items):
        """Настоящая Word-гиперссылка w:hyperlink с содержимым из буфера."""
        try:
            rid, _ = self.doc.part.relate_to(
                href,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                is_external=True)
        except Exception:
            rid = None
        hl = OxmlElement("w:hyperlink")
        if rid:
            hl.set(qn("r:id"), rid)
        for kind, payload, cur in buf_items:
            if kind == "text":
                parts = re.split(r"(\n)", payload)
                for part in parts:
                    if part == "\n":
                        continue
                    if not part:
                        continue
                    r_el = self._make_run_xml(part, cur, link_style=True)
                    hl.append(r_el)
            elif kind == "code":
                r_el = self._make_run_xml(payload, cur, code=True, link_style=True)
                hl.append(r_el)
        para._p.append(hl)

    def _make_run_xml(self, text, fmt, code=False, link_style=False):
        """Собирает <w:r> с нужным форматированием (для вставки в hyperlink)."""
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        if code:
            rf = OxmlElement("w:rFonts")
            rf.set(qn("w:ascii"), CODE_FONT)
            rf.set(qn("w:hAnsi"), CODE_FONT)
            rPr.append(rf)
            sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "19"); rPr.append(sz)
        if fmt.get("bold"):
            b = OxmlElement("w:b"); rPr.append(b)
        if fmt.get("italic"):
            i = OxmlElement("w:i"); rPr.append(i)
        if fmt.get("strike"):
            s = OxmlElement("w:strike"); rPr.append(s)
        if fmt.get("sub"):
            v = OxmlElement("w:vertAlign"); v.set(qn("w:val"), "subscript"); rPr.append(v)
        if fmt.get("sup"):
            v = OxmlElement("w:vertAlign"); v.set(qn("w:val"), "superscript"); rPr.append(v)
        if fmt.get("highlight"):
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear"); shd.set(qn("w:fill"), "FFF2A8")
            rPr.append(shd)
        if link_style:
            color = OxmlElement("w:color"); color.set(qn("w:val"), "1155AA")
            rPr.append(color)
            u = OxmlElement("w:u"); u.set(qn("w:val"), "single")
            rPr.append(u)
        elif fmt.get("underline"):
            u = OxmlElement("w:u"); u.set(qn("w:val"), "single")
            rPr.append(u)
        if len(rPr):
            r.append(rPr)
        t_el = OxmlElement("w:t")
        t_el.set(qn("xml:space"), "preserve")
        t_el.text = text
        r.append(t_el)
        return r

    def _register_adhoc_footnote(self, fid):
        idx = len(getattr(self, "_fn_defs", [])) + 1
        if not hasattr(self, "_fn_defs"):
            self._fn_defs = []
        self._fn_defs.append((fid, [{"type": "text", "content": "(сноска)", "children": None}]))
        self.footnotes[fid] = idx
        return idx

    def _add_real_footnote(self, para, idx):
        """Вставляет ссылку на сноску Word (w:footnoteReference с нужным id)."""
        ref = OxmlElement("w:r")
        rp = OxmlElement("w:rPr")
        v = OxmlElement("w:vertAlign")
        v.set(qn("w:val"), "superscript")
        rp.append(v)
        ref.append(rp)
        fr = OxmlElement("w:footnoteReference")
        fr.set(qn("w:id"), str(idx))
        ref.append(fr)
        para._p.append(ref)

    def _runs(self, para, text, fmt, href=None):
        parts = re.split(r"(\n)", text)
        for part in parts:
            if part == "\n":
                para.add_run().add_break()
                continue
            if not part:
                continue
            r = para.add_run(part)
            if fmt.get("bold"):
                r.bold = True
            if fmt.get("italic"):
                r.italic = True
            if fmt.get("strike"):
                r.font.strike = True
            if fmt.get("underline"):
                r.underline = True
            if fmt.get("sub"):
                r.font.subscript = True
            if fmt.get("sup"):
                r.font.superscript = True
            if fmt.get("highlight"):
                rPr = r._r.get_or_add_rPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:fill"), "FFF2A8")
                rPr.append(shd)
            if fmt.get("link"):
                r.font.color.rgb = LINK_COLOR
                r.underline = True


class _ParaCollector:
    """Псевдо-container: add_paragraph создаёт настоящий Paragraph из другого
    документа (для буферизации содержимого цитат)."""

    def __init__(self, sink):
        self.sink = sink
        self.tmp = Document()

    def add_paragraph(self, style=None):
        p = self.tmp.add_paragraph(style=style) if style else self.tmp.add_paragraph()
        self.sink.append(p)
        return p

    def add_table(self, **kw):
        return self.tmp.add_table(**kw)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="JSON-AST (md_ast.js) → DOCX")
    ap.add_argument("ast_json")
    ap.add_argument("output")
    ap.add_argument("--base-dir", default=".")
    args = ap.parse_args()

    with open(args.ast_json, encoding="utf-8") as f:
        tree = json.load(f)

    builder = Builder(base_dir=args.base_dir)
    builder.build(tree)
    builder.doc.save(args.output)
    print(f"OK: {args.output}")


if __name__ == "__main__":
    main()
