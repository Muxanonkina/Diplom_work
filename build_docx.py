#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Diplom ishini .docx (OOXML) formatida generatsiya qiluvchi skript.
Faqat Python standart kutubxonasidan foydalanadi (zipfile + xml).

Manba matnlar `manuscript/` papkasidagi *.md fayllarda joylashgan.
Belgilash (markup):
  # Sarlavha        -> Heading 1 (BOB / asosiy bo'lim)
  ## Sarlavha       -> Heading 2 (paragraf 1.1, 1.2 ...)
  ### Sarlavha      -> Heading 3
  CENTER: matn      -> markazlashtirilgan paragraf
  PAGEBREAK         -> sahifa uzilishi
  [^label]          -> matn ichidagi izoh (footnote) havolasi
  [^label]: matn    -> izoh ta'rifi (alohida qatorda)
  **qalin**         -> qalin (bold) matn
  oddiy paragraf    -> asoslangan (justified) asosiy matn
Bo'sh qator paragraflarni ajratadi.
"""

import os
import re
import glob
import zipfile

BASE = os.path.dirname(os.path.abspath(__file__))
MANUSCRIPT_DIR = os.path.join(BASE, "manuscript")
OUTPUT = os.path.join(BASE, "Diplom_AML_CFT_Uzbekiston.docx")

# ----------------------------------------------------------------------------
# Titul varaq ma'lumotlari (zarur bo'lsa tahrirlang)
# ----------------------------------------------------------------------------
TITLE_PAGE = {
    "ministry": "O‘ZBEKISTON RESPUBLIKASI OLIY TA’LIM, FAN VA INNOVATSIYALAR VAZIRLIGI",
    "university": "_________________________ UNIVERSITETI",
    "faculty": "Yuridik fakultet",
    "department": "Jinoyat huquqi, kriminologiya va korrupsiyaga qarshi kurashish kafedrasi",
    "worktype": "BITIRUV MALAKAVIY ISHI",
    "topic": ("Pul yuvishga va terrorizmni moliyalashtirishga qarshi kurashishning "
              "huquqiy mexanizmlari: muammolar va qonunchilikni takomillashtirish istiqbollari"),
    "student_label": "Bajardi:",
    "student": "_____ yo‘nalishi bitiruvchisi  ________________________",
    "supervisor_label": "Ilmiy rahbar:",
    "supervisor": "yuridik fanlar doktori (DSc), professor  ________________________",
    "city_year": "Toshkent — 2026",
}

# ============================================================================
#  XML yordamchi funksiyalar
# ============================================================================

def esc(text: str) -> str:
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;"))


# ----------------------------------------------------------------------------
#  Inline (matn ichidagi) elementlarni run'larga aylantirish
# ----------------------------------------------------------------------------
FOOTNOTE_REF_RE = re.compile(r"\[\^([^\]]+)\]")
BOLD_RE = re.compile(r"\*\*(.+?)\*\*")


def make_runs(text, footnote_ids, base_rpr=""):
    """Matnni run XML ro'yxatiga aylantiradi (footnote va bold qo'llab-quvvatlanadi)."""
    runs = []
    pos = 0
    # avval footnote havolalarini topamiz
    tokens = []
    last = 0
    for m in FOOTNOTE_REF_RE.finditer(text):
        if m.start() > last:
            tokens.append(("text", text[last:m.start()]))
        label = m.group(1)
        tokens.append(("fn", label))
        last = m.end()
    if last < len(text):
        tokens.append(("text", text[last:]))

    for kind, val in tokens:
        if kind == "fn":
            fid = footnote_ids.get(val)
            if fid is not None:
                runs.append(
                    '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                    f'<w:footnoteReference w:id="{fid}"/></w:r>'
                )
            continue
        # text bo'lagi -> bold bo'laklarga ajratamiz
        sub = val
        last2 = 0
        for bm in BOLD_RE.finditer(sub):
            if bm.start() > last2:
                runs.append(_text_run(sub[last2:bm.start()], base_rpr))
            runs.append(_text_run(bm.group(1), base_rpr + '<w:b/>'))
            last2 = bm.end()
        if last2 < len(sub):
            runs.append(_text_run(sub[last2:], base_rpr))
    return "".join(runs)


def _text_run(text, rpr_inner=""):
    if text == "":
        return ""
    rpr = f"<w:rPr>{rpr_inner}</w:rPr>" if rpr_inner else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


# ============================================================================
#  Manba fayllarni o'qish va footnote'larni yig'ish
# ============================================================================

def load_sources():
    files = sorted(glob.glob(os.path.join(MANUSCRIPT_DIR, "*.md")))
    raw_lines = []
    for f in files:
        with open(f, "r", encoding="utf-8") as fh:
            raw_lines.extend(fh.read().split("\n"))
        raw_lines.append("")  # fayllar orasida ajratuvchi
    return raw_lines


def collect_footnotes(lines):
    """Izoh ta'riflarini yig'adi va matn qatorlaridan ajratadi."""
    defs = {}
    body = []
    def_re = re.compile(r"^\[\^([^\]]+)\]:\s?(.*)$")
    for ln in lines:
        m = def_re.match(ln)
        if m:
            defs[m.group(1)] = m.group(2)
        else:
            body.append(ln)
    return body, defs


def assign_ids(body_lines, defs):
    """Footnote havolalariga ko'rinish tartibi bo'yicha ID beradi (1, 2, 3 ...)."""
    ids = {}
    counter = 1
    text = "\n".join(body_lines)
    for m in FOOTNOTE_REF_RE.finditer(text):
        label = m.group(1)
        if label in defs and label not in ids:
            ids[label] = counter
            counter += 1
    return ids


# ============================================================================
#  Body (document.xml) qurish
# ============================================================================

def build_body(body_lines, footnote_ids):
    out = []

    # ---- Titul varaq ----
    out.append(title_page_xml())
    out.append(page_break())

    # ---- Mundarija (TOC field) ----
    out.append(heading_para("MUNDARIJA", level=1, toc=False, center=True))
    out.append(toc_field())
    out.append(page_break())

    # ---- Asosiy matn ----
    paragraph_buf = []

    def flush():
        if not paragraph_buf:
            return
        joined = " ".join(s.strip() for s in paragraph_buf).strip()
        paragraph_buf.clear()
        if joined:
            out.append(body_para(joined, footnote_ids))

    for ln in body_lines:
        s = ln.rstrip()
        if s.strip() == "":
            flush()
            continue
        if s.strip() == "PAGEBREAK":
            flush()
            out.append(page_break())
            continue
        if s.startswith("### "):
            flush()
            out.append(heading_para(s[4:].strip(), level=3))
            continue
        if s.startswith("## "):
            flush()
            out.append(heading_para(s[3:].strip(), level=2))
            continue
        if s.startswith("# "):
            flush()
            out.append(page_break())
            out.append(heading_para(s[2:].strip(), level=1, center=True))
            continue
        if s.startswith("CENTER:"):
            flush()
            out.append(center_para(s[len("CENTER:"):].strip(), footnote_ids))
            continue
        paragraph_buf.append(s)
    flush()

    out.append(final_sectpr())
    return "".join(out)


def body_para(text, footnote_ids):
    runs = make_runs(text, footnote_ids)
    return (
        '<w:p><w:pPr>'
        '<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
        '<w:ind w:firstLine="709"/>'
        '<w:jc w:val="both"/>'
        '</w:pPr>'
        f'{runs}</w:p>'
    )


def center_para(text, footnote_ids):
    runs = make_runs(text, footnote_ids)
    return (
        '<w:p><w:pPr>'
        '<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
        '<w:jc w:val="center"/>'
        '</w:pPr>'
        f'{runs}</w:p>'
    )


def heading_para(text, level=1, toc=True, center=False):
    style = {1: "Heading1", 2: "Heading2", 3: "Heading3"}[level]
    jc = '<w:jc w:val="center"/>' if center else '<w:jc w:val="both"/>'
    return (
        '<w:p><w:pPr>'
        f'<w:pStyle w:val="{style}"/>'
        f'{jc}'
        '</w:pPr>'
        f'<w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
    )


def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def toc_field():
    return (
        '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">'
        '[Mundarijani yangilash uchun ushbu joyni sichqoncha o‘ng tugmasi bilan bosing -> Update Field (F9)]'
        '</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    )


def title_page_xml():
    p = []

    def line(text, bold=False, size=28, after=120, before=0, caps=False):
        rpr = ""
        if bold:
            rpr += "<w:b/>"
        rpr += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        if caps:
            rpr += "<w:caps/>"
        return (
            '<w:p><w:pPr>'
            f'<w:spacing w:before="{before}" w:after="{after}" w:line="240" w:lineRule="auto"/>'
            '<w:jc w:val="center"/>'
            f'</w:pPr><w:r><w:rPr>{rpr}</w:rPr>'
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
        )

    def empty(n=1):
        return ('<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:p>') * n

    p.append(line(TITLE_PAGE["ministry"], bold=True, size=24, after=80))
    p.append(line(TITLE_PAGE["university"], bold=True, size=28, after=80))
    p.append(line(TITLE_PAGE["faculty"], size=26, after=40))
    p.append(line(TITLE_PAGE["department"], size=26, after=40))
    p.append(empty(4))
    p.append(line(TITLE_PAGE["worktype"], bold=True, size=40, after=120))
    p.append(line("mavzusida", size=26, after=120))
    p.append(empty(1))
    p.append(line(TITLE_PAGE["topic"], bold=True, size=32, after=200))
    p.append(empty(5))

    # Student / rahbar (o'ngga moyil)
    def right_line(text, size=26):
        return (
            '<w:p><w:pPr>'
            '<w:spacing w:before="0" w:after="60" w:line="276" w:lineRule="auto"/>'
            '<w:ind w:left="4536"/><w:jc w:val="both"/>'
            f'</w:pPr><w:r><w:rPr><w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
        )

    p.append(right_line(TITLE_PAGE["student_label"] + " " + TITLE_PAGE["student"]))
    p.append(right_line(TITLE_PAGE["supervisor_label"] + " " + TITLE_PAGE["supervisor"]))
    p.append(empty(4))
    p.append(line(TITLE_PAGE["city_year"], bold=True, size=26, after=0))
    return "".join(p)


def final_sectpr():
    return (
        '<w:sectPr>'
        '<w:footerReference w:type="default" r:id="rIdFooter"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" '
        'w:header="708" w:footer="708" w:gutter="0"/>'
        '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )


# ============================================================================
#  footnotes.xml
# ============================================================================

def build_footnotes_xml(defs, footnote_ids):
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    # ID bo'yicha tartiblaymiz
    for label, fid in sorted(footnote_ids.items(), key=lambda kv: kv[1]):
        txt = defs.get(label, "")
        runs = make_runs(" " + txt, {})  # footnote ichida ichki footnote yo'q
        parts.append(
            f'<w:footnote w:id="{fid}">'
            '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/><w:jc w:val="both"/></w:pPr>'
            '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            f'{runs}</w:p></w:footnote>'
        )
    parts.append('</w:footnotes>')
    return "".join(parts)


# ============================================================================
#  Statik fayllar
# ============================================================================

def content_types_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '</Types>'
    )


def root_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>'
    )


def document_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
        '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'
        '<Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
        '</Relationships>'
    )


def settings_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:zoom w:percent="100"/>'
        '<w:defaultTabStop w:val="708"/>'
        '<w:footnotePr><w:numFmt w:val="decimal"/><w:numRestart w:val="continuous"/></w:footnotePr>'
        '<w:compat><w:compatSetting w:name="compatibilityMode" '
        'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
        '<w:updateFields w:val="true"/>'
        '</w:settings>'
    )


def footer1_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>1</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        '</w:ftr>'
    )


def fonttable_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:font w:name="Times New Roman"><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
        '</w:fonts>'
    )


def core_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>Bitiruv malakaviy ishi</dc:title>'
        '<dc:creator>Bitiruvchi</dc:creator>'
        '<cp:revision>1</cp:revision>'
        '</cp:coreProperties>'
    )


def app_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        '<Application>Kiro DOCX Builder</Application>'
        '</Properties>'
    )


def styles_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        '<w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="uz-Latn-UZ"/>'
        '</w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        '</w:docDefaults>'
        # Normal
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>'
        '<w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>'
        # Heading 1
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="200" w:line="360" w:lineRule="auto"/>'
        '<w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:b/><w:caps/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>'
        # Heading 2
        '<w:style w:type="paragraph" w:styleId="Heading2">'
        '<w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:keepNext/><w:spacing w:before="200" w:after="120" w:line="360" w:lineRule="auto"/>'
        '<w:outlineLvl w:val="1"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>'
        # Heading 3
        '<w:style w:type="paragraph" w:styleId="Heading3">'
        '<w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80" w:line="360" w:lineRule="auto"/>'
        '<w:outlineLvl w:val="2"/></w:pPr>'
        '<w:rPr><w:b/><w:i/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>'
        # Footnote text
        '<w:style w:type="paragraph" w:styleId="FootnoteText">'
        '<w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>'
        '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
        # Footnote reference
        '<w:style w:type="character" w:styleId="FootnoteReference">'
        '<w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
        '</w:styles>'
    )


def document_xml(body):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{body}</w:body></w:document>'
    )


# ============================================================================
#  Asosiy build
# ============================================================================

def main():
    lines = load_sources()
    body_lines, defs = collect_footnotes(lines)
    footnote_ids = assign_ids(body_lines, defs)

    body = build_body(body_lines, footnote_ids)
    doc = document_xml(body)
    fns = build_footnotes_xml(defs, footnote_ids)

    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types_xml())
        z.writestr("_rels/.rels", root_rels_xml())
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", document_rels_xml())
        z.writestr("word/styles.xml", styles_xml())
        z.writestr("word/settings.xml", settings_xml())
        z.writestr("word/footnotes.xml", fns)
        z.writestr("word/footer1.xml", footer1_xml())
        z.writestr("word/fontTable.xml", fonttable_xml())
        z.writestr("docProps/core.xml", core_xml())
        z.writestr("docProps/app.xml", app_xml())

    # statistika
    words = 0
    for ln in body_lines:
        words += len(ln.split())
    print(f"OK: {OUTPUT}")
    print(f"Footnotes: {len(footnote_ids)} | approx words: {words} | approx pages (~330 w/p): {words/330:.1f}")


if __name__ == "__main__":
    main()
