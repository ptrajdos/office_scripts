#!/usr/bin/env python3
"""Python port of prepare_protocols.vba (StrikeThroughPhrasesInDocs).

For every .docx in a folder:
  1. joins the last two non-empty paragraphs with a manual line break,
  2. deletes the last page,
  3. strikes through the given phrases (case insensitive).

python-docx has no layout engine, so "last page" means everything after the last
page boundary stored in the file: a page break, a page-starting section break,
"page break before" or the w:lastRenderedPageBreak marker that Word writes when
saving. Files saved by Word (or converted by pdf2docx) contain these.
"""
import os
import re
import copy
import argparse
from docx import Document
from docx.oxml.ns import qn
from docx.text.run import Run
from docx.text.paragraph import Paragraph

PHRASES_TO_STRIKE = ["większością głosów", "nie uzyskał tytułu zawodowego", "nie uzyskała tytułu zawodowego"]

# Run children that Run.text can represent losslessly, so the run can be split
TEXT_RUN_CHILDREN = {qn("w:rPr"), qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr"), qn("w:noBreakHyphen")}


def all_paragraphs(doc):
    return [Paragraph(p, doc._body) for p in doc.element.body.iter(qn("w:p"))]


def move_last_line_up(doc):
    non_empty = [p for p in all_paragraphs(doc) if p.text.strip()]
    if len(non_empty) < 2:
        return
    second_last, last = non_empty[-2]._p, non_empty[-1]._p

    # Manual line break (Shift+Enter), then move the last paragraph's content after it
    br_run = second_last._add_r()
    br_run.append(br_run.makeelement(qn("w:br"), {}))
    for child in list(last):
        if child.tag != qn("w:pPr"):
            second_last.append(child)

    # Keep the paragraph if it carries a section break, otherwise remove it
    if last.find(qn("w:pPr") + "/" + qn("w:sectPr")) is None:
        last.getparent().remove(last)


def last_page_start(body):
    """Return (node, remove_node) marking where the last page begins, or None."""
    start = None
    for el in body.iter(qn("w:br"), qn("w:lastRenderedPageBreak"), qn("w:pageBreakBefore"), qn("w:sectPr")):
        if el.tag == qn("w:br"):
            if el.get(qn("w:type")) == "page":
                start = (el, True)
        elif el.tag == qn("w:lastRenderedPageBreak"):
            start = (el, True)
        elif el.tag == qn("w:pageBreakBefore"):
            if el.get(qn("w:val"), "true") not in ("0", "false", "off"):
                start = (el.getparent().getparent(), True)
        elif el.getparent().tag == qn("w:pPr"):
            # Section break stored on a paragraph: the next page starts after that paragraph
            sect_type = el.find(qn("w:type"))
            if sect_type is None or sect_type.get(qn("w:val")) != "continuous":
                start = (el.getparent().getparent(), False)
    return start


def delete_last_page(doc):
    body = doc.element.body
    start = last_page_start(body)
    if start is None:
        return
    node, remove_node = start

    # Delete everything after the start point, climbing up to the body
    current = node
    while current is not body:
        parent = current.getparent()
        for sibling in list(current.itersiblings()):
            if not (parent is body and sibling.tag == qn("w:sectPr")):
                parent.remove(sibling)
        current = parent
    if remove_node:
        node.getparent().remove(node)
    else:
        # The section break would leave an empty trailing page: make it the final section instead
        sect_pr = node.find(qn("w:pPr") + "/" + qn("w:sectPr"))
        body.replace(body.find(qn("w:sectPr")), sect_pr)
        if not node.xpath("./w:r | ./w:hyperlink"):
            body.remove(node)


def runs_with_offsets(p):
    runs, pos = [], 0
    for r in p.xpath("./w:r | ./w:hyperlink/w:r | ./w:ins/w:r"):
        text = Run(r, None).text
        runs.append((r, pos, pos + len(text), text))
        pos += len(text)
    return runs


def strike_through_phrases(doc, phrases):
    for paragraph in all_paragraphs(doc):
        p = paragraph._p
        text = "".join(t for _, _, _, t in runs_with_offsets(p))
        spans = [m.span() for phrase in phrases
                 for m in re.finditer(re.escape(phrase), text, re.IGNORECASE)]
        if not spans:
            continue

        # Split runs so every span boundary falls between runs
        boundaries = sorted({b for span in spans for b in span})
        for r, start, end, run_text in runs_with_offsets(p):
            if any(child.tag not in TEXT_RUN_CHILDREN for child in r):
                continue
            cuts = [b - start for b in boundaries if start < b < end]
            pieces = [run_text[i:j] for i, j in zip([0] + cuts, cuts + [len(run_text)])]
            if len(pieces) < 2:
                continue
            Run(r, None).text = pieces[0]
            prev = r
            for piece in pieces[1:]:
                new_r = copy.deepcopy(r)
                Run(new_r, None).text = piece
                prev.addnext(new_r)
                prev = new_r

        for r, start, end, _ in runs_with_offsets(p):
            if end > start and any(s <= start and end <= e for s, e in spans):
                Run(r, None).font.strike = True


def process_protocols(directory_path, phrases):
    for file in sorted(os.listdir(directory_path)):
        if not file.lower().endswith(".docx") or file.startswith("~$"):
            continue
        file_path = os.path.join(directory_path, file)
        try:
            doc = Document(file_path)
        except Exception as e:
            print(f"Failed to open: {file_path}: {e}")
            continue

        move_last_line_up(doc)
        delete_last_page(doc)
        strike_through_phrases(doc, phrases)
        doc.save(file_path)

    print("Done striking through phrases in documents.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Prepare protocol DOCX files: merge last line, delete last page, strike phrases.")
    parser.add_argument("input_dir", help="Directory containing DOCX files (modified in place)")
    parser.add_argument("--phrase", action="append", help="Phrase to strike through (repeatable; default: built-in list)")
    args = parser.parse_args()
    process_protocols(args.input_dir, args.phrase or PHRASES_TO_STRIKE)
