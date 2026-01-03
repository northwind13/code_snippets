#!/usr/bin/env python3
"""
Add a citation to a Microsoft Word Sources.xml (bibliography) file.

Usage:
  python add_word_source.py "CODE_NAME" "Author, A. A. (YYYY). Title. Publisher." --in Sources.xml

Example:
  python add_word_source.py "Simon1996" "Simon, H. A. (1996). The Sciences of the Artificial. MIT Press." --in Sources.xml

Behavior:
- If the source (Tag=CODE_NAME) already exists, the file is copied to new_<inputname> unchanged.
- If it does not exist, the source is appended and written to new_<inputname>.
"""

from __future__ import annotations

import argparse
import re
import uuid
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Tuple

BIB_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
NS = {"b": BIB_NS}
ET.register_namespace("b", BIB_NS)

EXAMPLE = (
    'python add_word_source.py "Simon1996" '
    '"Simon, H. A. (1996). The Sciences of the Artificial. MIT Press." --in Sources.xml'
)


def die(msg: str):
    print("\nERROR:", msg, "\n", file=sys.stderr)
    print("Correct usage example:\n", EXAMPLE, "\n", file=sys.stderr)
    sys.exit(1)


def parse_book_apa_like(s: str) -> Tuple[List[Tuple[str, str]], str, str, str]:
    s = re.sub(r"\s+", " ", s.strip())

    m_year = re.search(r"\((\d{4})\)", s)
    if not m_year:
        die("Year not found. Expected '(YYYY)'.")

    year = m_year.group(1)
    before_year = s[:m_year.start()].strip().rstrip(".")
    after_year = s[m_year.end():].strip()

    authors_part = before_year.replace(", &", " &")
    author_chunks = [a.strip() for a in authors_part.split(" & ")]

    authors: List[Tuple[str, str]] = []
    for chunk in author_chunks:
        parts = re.split(r"\.\s+(?=[A-Z][^,]+,)", chunk)
        if len(parts) == 1:
            parts = [p.strip() for p in chunk.split(".,") if p.strip()]
            parts = [p + "." if not p.endswith(".") else p for p in parts]

        for p in parts:
            p = p.strip().rstrip(".")
            m = re.match(r"^(?P<last>[^,]+),\s*(?P<first>.+)$", p)
            if not m:
                die(f"Author format error near: '{p}'")
            authors.append((m.group("last").strip(), m.group("first").strip()))

    after_year = after_year.lstrip(". ").strip()
    pieces = [p.strip() for p in after_year.split(".") if p.strip()]
    if len(pieces) < 2:
        die("Expected format 'Title. Publisher.' after year.")

    title = re.sub(r"\s*\([^)]*ed\.\)\s*$", "", pieces[0], flags=re.I).strip()
    publisher = pieces[1]

    return authors, year, title, publisher


def source_tag_exists(root: ET.Element, tag: str) -> bool:
    for src in root.findall("b:Source", NS):
        if src.findtext("b:Tag", "", NS) == tag:
            return True
    return False


def next_reforder(root: ET.Element):
    vals = []
    for src in root.findall("b:Source", NS):
        ro = src.findtext("b:RefOrder", "", NS)
        if ro.isdigit():
            vals.append(int(ro))
    return max(vals) + 1 if vals else None


def build_book_source(tag: str, ref: str) -> ET.Element:
    authors, year, title, publisher = parse_book_apa_like(ref)

    src = ET.Element(f"{{{BIB_NS}}}Source")
    ET.SubElement(src, f"{{{BIB_NS}}}Tag").text = tag
    ET.SubElement(src, f"{{{BIB_NS}}}SourceType").text = "Book"

    ao = ET.SubElement(src, f"{{{BIB_NS}}}Author")
    ai = ET.SubElement(ao, f"{{{BIB_NS}}}Author")
    nl = ET.SubElement(ai, f"{{{BIB_NS}}}NameList")

    for last, first in authors:
        p = ET.SubElement(nl, f"{{{BIB_NS}}}Person")
        ET.SubElement(p, f"{{{BIB_NS}}}Last").text = last
        ET.SubElement(p, f"{{{BIB_NS}}}First").text = first

    ET.SubElement(src, f"{{{BIB_NS}}}Title").text = title
    ET.SubElement(src, f"{{{BIB_NS}}}Year").text = year
    ET.SubElement(src, f"{{{BIB_NS}}}Publisher").text = publisher
    ET.SubElement(src, f"{{{BIB_NS}}}Guid").text = "{" + str(uuid.uuid4()).upper() + "}"

    return src


def indent(elem, level=0):
    i = "\n" + level * "\t"
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "\t"
        for c in elem:
            indent(c, level + 1)
        if not elem[-1].tail or not elem[-1].tail.strip():
            elem[-1].tail = i
    if level and (not elem.tail or not elem.tail.strip()):
        elem.tail = i


def main():
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("code_name", nargs="?")
    ap.add_argument("reference", nargs="?")
    ap.add_argument("--in", dest="infile")
    args = ap.parse_args()

    if not args.code_name or not args.reference or not args.infile:
        die("Missing required arguments.")

    in_path = Path(args.infile)
    if not in_path.exists():
        die(f"Input file not found: {in_path}")

    out_path = in_path.with_name("new_" + in_path.name)

    tree = ET.parse(in_path)
    root = tree.getroot()

    if source_tag_exists(root, args.code_name):
        tree.write(out_path, encoding="utf-8", xml_declaration=True)
        print(f"Tag '{args.code_name}' already exists. Wrote: {out_path}")
        return

    src = build_book_source(args.code_name, args.reference)

    nro = next_reforder(root)
    if nro is not None:
        ET.SubElement(src, f"{{{BIB_NS}}}RefOrder").text = str(nro)

    root.append(src)
    indent(root)
    tree.write(out_path, encoding="utf-8", xml_declaration=True)
    print(f"Added '{args.code_name}'. Wrote: {out_path}")


if __name__ == "__main__":
    main()
