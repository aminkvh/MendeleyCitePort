#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mendeley_cite_port.py
=====================

Convert citations made with the *new* Mendeley Cite add-in (which stores each
citation as a Word content control with base64-encoded CSL JSON) into classic
Word field codes:

  * ``mendeley``  -> the format written by the old Mendeley Desktop plugin
                     (``ADDIN CSL_CITATION {...}``)
  * ``zotero``    -> the format written by the Zotero Word plugin
                     (``ADDIN ZOTERO_ITEM CSL_CITATION {...}``)

Paperpile ("Settings and Tools > Convert from > Mendeley / Zotero") and Zotero
can read those legacy field codes, but neither can read Mendeley Cite content
controls.  This script bridges that gap.  It edits the .docx package directly,
needs only the Python standard library, and never touches the visible text.

Usage
-----
    python mendeley_cite_port.py manuscript.docx
    python mendeley_cite_port.py manuscript.docx --format zotero
    python mendeley_cite_port.py manuscript.docx --format both --bib refs.bib --dois dois.txt

Run with ``-h`` for all options.

Author: Amin Akbari.  Released into the public domain (The Unlicense).
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape, unescape

__version__ = "1.0.0"

CSL_SCHEMA = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
TAG_PREFIX = "MENDELEY_CITATION_v3_"
BIB_TAG = "MENDELEY_BIBLIOGRAPHY"
MENDELEY_STORE_ID = "wa104382081"          # AppSource id of the Mendeley Cite add-in

PARTS_RE = re.compile(r"^word/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$")
SDT_TOK = re.compile(r"<w:sdt>|<w:sdt |</w:sdt>")
SDT_START = re.compile(r"<w:sdt>|<w:sdt ")

CUSTOM_PART = "docProps/custom.xml"
CUSTOM_CT = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CUSTOM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
EMPTY_CUSTOM = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>'
)


class ConversionError(Exception):
    """Raised when the input cannot be converted."""


# --------------------------------------------------------------------------- #
# Low-level OOXML helpers (string based on purpose: byte-for-byte fidelity of
# everything we do not touch, and no namespace-prefix rewriting)
# --------------------------------------------------------------------------- #
def sdt_end(xml: str, start: int) -> int:
    """Index just past the ``</w:sdt>`` that closes the sdt opening at *start*."""
    depth = 0
    for m in SDT_TOK.finditer(xml, start):
        depth += -1 if m.group(0) == "</w:sdt>" else 1
        if depth == 0:
            return m.end()
    raise ConversionError("unbalanced <w:sdt> element")


def split_sdt(sdt: str) -> tuple[str, str]:
    """Return (inner xml of <w:sdtPr>, inner xml of <w:sdtContent>)."""
    ps, pe = sdt.find("<w:sdtPr>"), sdt.find("</w:sdtPr>")
    sdtpr = sdt[ps + 9:pe] if ps != -1 and pe != -1 else ""
    cs = sdt.find("<w:sdtContent>")
    if cs == -1:
        return sdtpr, ""
    cs += len("<w:sdtContent>")
    ce = sdt.rfind("</w:sdtContent>")
    return sdtpr, sdt[cs:ce]


def rpr_of(sdtpr: str) -> str:
    m = re.search(r"<w:rPr>.*?</w:rPr>", sdtpr, re.S)
    return m.group(0) if m else ""


def text_of(xml: str) -> str:
    return unescape("".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", xml)))


def field_begin(rpr: str, code: str) -> str:
    def run(inner: str) -> str:
        return "<w:r>%s%s</w:r>" % (rpr, inner)
    return (run('<w:fldChar w:fldCharType="begin"/>')
            + run('<w:instrText xml:space="preserve"> %s </w:instrText>' % escape(code))
            + run('<w:fldChar w:fldCharType="separate"/>'))


def field_end(rpr: str) -> str:
    return '<w:r>%s<w:fldChar w:fldCharType="end"/></w:r>' % rpr


def strip_tags(html: str) -> str:
    return re.sub(r"<[^>]+>", "", html)


# --------------------------------------------------------------------------- #
# Mendeley Cite decoding
# --------------------------------------------------------------------------- #
def decode_tag(tag: str) -> dict:
    return json.loads(base64.b64decode(tag[len(TAG_PREFIX):]).decode("utf-8"))


def formatted_text(cite: dict, content_xml: str) -> str:
    """The citation as Mendeley rendered it (may contain <sup>/<i> tags)."""
    mo = cite.get("manualOverride") or {}
    if mo.get("isManuallyOverridden") and mo.get("manualOverrideText"):
        return mo["manualOverrideText"]
    return mo.get("citeprocText") or escape(text_of(content_xml))


def clean_item(item_data: dict, keep_abstracts: bool) -> dict:
    d = dict(item_data)
    if not keep_abstracts:
        d.pop("abstract", None)
    return d


# --------------------------------------------------------------------------- #
# Field-code builders
# --------------------------------------------------------------------------- #
class MendeleyDesktopFormat:
    name = "mendeley"
    suffix = "LEGACY_MENDELEY_fields"
    bib_code = "ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY"

    def __init__(self, keep_abstracts: bool = False, **_: object) -> None:
        self.keep_abstracts = keep_abstracts

    def cite_code(self, cite: dict, content_xml: str) -> str:
        items = []
        for n, ci in enumerate(cite.get("citationItems", []), 1):
            d = clean_item(ci["itemData"], self.keep_abstracts)
            d["id"] = "ITEM-%d" % n
            uris = list(ci.get("uris") or [])
            if not uris:  # Mendeley Desktop always wrote a document uuid; synthesize one
                uris = ["http://www.mendeley.com/documents/?uuid=%s"
                        % (ci.get("legacyDesktopId") or ci.get("id") or d.get("DOI", "unknown"))]
            it = {"id": "ITEM-%d" % n, "itemData": d, "uris": uris}
            for k in ("locator", "label", "prefix", "suffix", "suppress-author", "author-only"):
                if k in ci:
                    it[k] = ci[k]
            items.append(it)
        fmt = formatted_text(cite, content_xml)
        payload = {
            "citationItems": items,
            "mendeley": {"formattedCitation": fmt,
                         "plainTextFormattedCitation": strip_tags(fmt),
                         "previouslyFormattedCitation": fmt},
            "properties": {"noteIndex": (cite.get("properties") or {}).get("noteIndex", 0)},
            "schema": CSL_SCHEMA,
        }
        return "ADDIN CSL_CITATION " + json.dumps(payload, ensure_ascii=False,
                                                  separators=(",", ":"), sort_keys=True)

    def extra_package_edits(self, pkg: "Package") -> None:
        pass


class ZoteroFormat:
    name = "zotero"
    suffix = "ZOTERO_fields"
    bib_code = 'ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY'

    def __init__(self, keep_abstracts: bool = False,
                 zotero_style: str = "http://www.zotero.org/styles/apa", **_: object) -> None:
        self.keep_abstracts = keep_abstracts
        self.zotero_style = zotero_style
        self._keys: dict[str, tuple[str, int]] = {}

    def _key(self, item_id: str) -> tuple[str, int]:
        if item_id not in self._keys:
            h = hashlib.sha1(item_id.encode("utf-8")).hexdigest().upper()
            self._keys[item_id] = (h[:8], int(h[:6], 16) % 100000 + 1)
        return self._keys[item_id]

    def cite_code(self, cite: dict, content_xml: str) -> str:
        items = []
        for ci in cite.get("citationItems", []):
            d = clean_item(ci["itemData"], self.keep_abstracts)
            key, num = self._key(str(ci.get("id") or d.get("DOI") or d.get("title", "")))
            d["id"] = num
            it = {"id": num, "uris": ["http://zotero.org/users/local/a1b2c3d4/items/%s" % key],
                  "itemData": d}
            for k in ("locator", "label", "prefix", "suffix", "suppress-author", "author-only"):
                if k in ci:
                    it[k] = ci[k]
            items.append(it)
        fmt = formatted_text(cite, content_xml)
        payload = {
            "citationID": hashlib.md5(str(cite.get("citationID", fmt)).encode("utf-8")).hexdigest()[:8],
            "properties": {"formattedCitation": fmt, "plainCitation": strip_tags(fmt), "noteIndex": 0},
            "citationItems": items,
            "schema": CSL_SCHEMA,
        }
        return "ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload, ensure_ascii=False,
                                                              separators=(",", ":"))

    def extra_package_edits(self, pkg: "Package") -> None:
        pref = ('<data data-version="3" zotero-version="6.0.30"><session id="k9Xz2Qwe"/>'
                '<style id="%s" locale="en-US" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
                '<prefs><pref name="fieldType" value="Field"/></prefs></data>' % self.zotero_style)
        pkg.add_custom_property("ZOTERO_PREF_1", pref)


FORMATS = {"mendeley": MendeleyDesktopFormat, "zotero": ZoteroFormat}


# --------------------------------------------------------------------------- #
# Converting one XML part (document.xml, footnotes.xml, ...)
# --------------------------------------------------------------------------- #
def convert_part(xml: str, fmt) -> tuple[str, int, int, list[str]]:
    """Return (new_xml, n_citations, n_bibliographies, unknown_tags)."""
    out: list[str] = []
    pos = 0
    n_cit = n_bib = 0
    unknown: list[str] = []
    for m in SDT_START.finditer(xml):
        if m.start() < pos:          # nested inside an sdt we already consumed
            continue
        s = m.start()
        e = sdt_end(xml, s)
        sdt = xml[s:e]
        tm = re.search(r'<w:tag w:val="([^"]*)"', sdt)
        tag = unescape(tm.group(1)) if tm else ""
        out.append(xml[pos:s])
        sdtpr, content = split_sdt(sdt)
        rpr = rpr_of(sdtpr)
        if tag.startswith(TAG_PREFIX):
            cite = decode_tag(tag)
            out.append(field_begin(rpr, fmt.cite_code(cite, content)) + content + field_end(rpr))
            n_cit += 1
        elif tag == BIB_TAG:
            # One field spanning every bibliography paragraph: begin in the first, end in the last.
            first_p = re.search(r"<w:p(?: [^>]*)?>", content)
            if first_p is None:      # empty bibliography: keep nothing but the field itself
                content = "<w:p>%s%s</w:p>" % (field_begin(rpr, fmt.bib_code), field_end(rpr))
            else:
                ins = first_p.end()
                if content.startswith("<w:pPr>", ins):
                    ins = content.find("</w:pPr>", ins) + len("</w:pPr>")
                content = content[:ins] + field_begin(rpr, fmt.bib_code) + content[ins:]
                last = content.rfind("</w:p>")
                content = content[:last] + field_end(rpr) + content[last:]
            out.append(content)
            n_bib += 1
        else:
            unknown.append(tag)
            out.append(sdt)          # some other content control: leave untouched
        pos = e
    out.append(xml[pos:])
    return "".join(out), n_cit, n_bib, unknown


# --------------------------------------------------------------------------- #
# Package (zip) level editing
# --------------------------------------------------------------------------- #
class Package:
    def __init__(self, src: str) -> None:
        self.zin = zipfile.ZipFile(src)
        self.names = self.zin.namelist()
        self.replacements: dict[str, str] = {}
        self.skip: set[str] = set()

    def read(self, name: str) -> str:
        if name in self.replacements:
            return self.replacements[name]
        return self.zin.read(name).decode("utf-8")

    def write(self, name: str, data: str) -> None:
        self.replacements[name] = data

    # -- Mendeley Cite web-extension (task pane) removal ---------------------
    def strip_mendeley_addin(self) -> int:
        tp, tprels = "word/webextensions/taskpanes.xml", "word/webextensions/_rels/taskpanes.xml.rels"
        if tp not in self.names or tprels not in self.names:
            return 0
        rels = self.read(tprels)
        removed_ids: list[str] = []
        for rel in re.findall(r"<Relationship [^>]*/>", rels):
            rid = re.search(r'Id="([^"]+)"', rel)
            tgt = re.search(r'Target="([^"]+)"', rel)
            if not rid or not tgt:
                continue
            part = os.path.normpath(os.path.join("word/webextensions", tgt.group(1))).replace("\\", "/")
            if part in self.names:
                body = self.zin.read(part).decode("utf-8", "replace")
                if MENDELEY_STORE_ID in body or "MENDELEY_" in body:
                    self.skip.add(part)
                    removed_ids.append(rid.group(1))
        if not removed_ids:
            return 0
        for rid in removed_ids:
            rels = re.sub(r'<Relationship [^>]*Id="%s"[^>]*/>' % re.escape(rid), "", rels)
        taskpanes = self.read(tp)
        for rid in removed_ids:
            taskpanes = re.sub(r'<wetp:taskpane\b(?:(?!</wetp:taskpane>).)*?r:id="%s"(?:(?!</wetp:taskpane>).)*?</wetp:taskpane>'
                               % re.escape(rid), "", taskpanes, flags=re.S)
        if "<wetp:taskpane" in taskpanes:
            self.write(tp, taskpanes)
            self.write(tprels, rels)
        else:                                  # nothing left: drop the whole task-pane part
            self.skip.update({tp, tprels})
            self.write("_rels/.rels", re.sub(r"<Relationship [^>]*webextensiontaskpanes[^>]*/>", "",
                                             self.read("_rels/.rels")))
        ct = self.read("[Content_Types].xml")
        for part in self.skip:
            ct = re.sub(r'<Override PartName="/%s"[^>]*/>' % re.escape(part), "", ct)
        self.write("[Content_Types].xml", ct)
        return len(removed_ids)

    # -- custom document properties ------------------------------------------
    def add_custom_property(self, name: str, value: str) -> None:
        if CUSTOM_PART in self.names or CUSTOM_PART in self.replacements:
            custom = self.read(CUSTOM_PART)
        else:
            custom = EMPTY_CUSTOM
            ct = self.read("[Content_Types].xml")
            if "/%s" % CUSTOM_PART not in ct:
                ct = ct.replace("</Types>", '<Override PartName="/%s" ContentType="%s"/></Types>'
                                % (CUSTOM_PART, CUSTOM_CT))
                self.write("[Content_Types].xml", ct)
            rels = self.read("_rels/.rels")
            if CUSTOM_REL not in rels:
                rels = rels.replace("</Relationships>",
                                    '<Relationship Id="rIdCustomProps" Type="%s" Target="%s"/></Relationships>'
                                    % (CUSTOM_REL, CUSTOM_PART))
                self.write("_rels/.rels", rels)
        custom = re.sub(r'<property [^>]*name="%s">.*?</property>' % re.escape(name), "", custom, flags=re.S)
        pids = [int(x) for x in re.findall(r'pid="(\d+)"', custom)] or [1]
        prop = ('<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="%d" name="%s">'
                '<vt:lpwstr>%s</vt:lpwstr></property>' % (max(pids) + 1, escape(name), escape(value)))
        self.write(CUSTOM_PART, custom.replace("</Properties>", prop + "</Properties>"))

    # -- output ---------------------------------------------------------------
    def save(self, dst: str) -> None:
        written = set()
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in self.zin.infolist():
                if info.filename in self.skip:
                    continue
                data = self.replacements.get(info.filename)
                zout.writestr(info, data.encode("utf-8") if data is not None else self.zin.read(info.filename))
                written.add(info.filename)
            for name, data in self.replacements.items():      # brand-new parts
                if name not in written and name not in self.skip:
                    zout.writestr(name, data.encode("utf-8"))

    def close(self) -> None:
        self.zin.close()


# --------------------------------------------------------------------------- #
# Reference export (BibTeX / DOI list) from the embedded CSL data
# --------------------------------------------------------------------------- #
def collect_items(src: str) -> dict[str, dict]:
    """All unique cited items (keyed by Mendeley item id) across the whole document."""
    items: dict[str, dict] = {}
    with zipfile.ZipFile(src) as z:
        for name in z.namelist():
            if not PARTS_RE.match(name):
                continue
            xml = z.read(name).decode("utf-8")
            for tag in re.findall(r'<w:tag w:val="(%s[^"]+)"' % TAG_PREFIX, xml):
                for ci in decode_tag(unescape(tag)).get("citationItems", []):
                    items.setdefault(str(ci.get("id")), ci["itemData"])
    return items


def _bib_escape(s: str) -> str:
    return (s or "").replace("&", r"\&").replace("%", r"\%").replace("#", r"\#").replace("_", r"\_")


CSL_TO_BIBTYPE = {"article-journal": "article", "article": "article", "book": "book",
                  "chapter": "incollection", "paper-conference": "inproceedings",
                  "thesis": "phdthesis", "report": "techreport", "webpage": "misc"}


def _record_score(d: dict) -> tuple[bool, int]:
    """Rank duplicate library records so a symbol-preserving one wins over an
    ASCII-transliterated duplicate (e.g. title "...beta1" vs "...β1") regardless
    of citation order."""
    title = d.get("title") or ""
    has_non_ascii = any(ord(c) > 127 for c in title)
    n_fields = sum(1 for v in d.values() if v)
    return (has_non_ascii, n_fields)


def write_bibtex(items: dict[str, dict], path: str) -> int:
    best: dict[str, dict] = {}
    for d in items.values():
        dedupe = (d.get("DOI") or "").strip().lower() or (d.get("title") or "").strip().lower()
        if not dedupe:
            continue
        if dedupe not in best or _record_score(d) > _record_score(best[dedupe]):
            best[dedupe] = d

    keys: set[str] = set()
    entries: list[str] = []
    for d in best.values():
        authors = d.get("author") or []
        au = " and ".join(("%s, %s" % (a.get("family", ""), a.get("given", ""))).strip(", ")
                          if a.get("family") else a.get("literal", "") for a in authors)
        first = re.sub(r"[^A-Za-z]", "", authors[0].get("family", "") if authors else "") or "Anon"
        dp = (d.get("issued") or {}).get("date-parts") or [[None]]
        year = dp[0][0] if dp and dp[0] else None
        key = base = "%s%s" % (first, year or "nd")
        n = 0
        while key in keys:
            n += 1
            key = "%s%s" % (base, chr(ord("a") + n - 1))
        keys.add(key)
        fields = [("author", au), ("title", d.get("title", "")), ("journal", d.get("container-title", "")),
                  ("year", str(year) if year else ""), ("volume", d.get("volume", "")),
                  ("number", d.get("issue", "")), ("pages", d.get("page", "")), ("doi", d.get("DOI", "")),
                  ("url", d.get("URL", "")), ("publisher", d.get("publisher", "")),
                  ("issn", d.get("ISSN", "")), ("isbn", d.get("ISBN", "")), ("pmid", d.get("PMID", ""))]
        body = ",\n".join("  %s = {%s}" % (k, v if k in ("doi", "url") else _bib_escape(v))
                          for k, v in fields if v)
        entries.append("@%s{%s,\n%s\n}\n" % (CSL_TO_BIBTYPE.get(d.get("type", ""), "misc"), key, body))
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(entries))
    return len(entries)


def write_dois(items: dict[str, dict], path: str) -> int:
    dois = sorted({(d.get("DOI") or "").strip().lower() for d in items.values() if d.get("DOI")})
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(dois) + ("\n" if dois else ""))
    return len(dois)


# --------------------------------------------------------------------------- #
# Top-level conversion
# --------------------------------------------------------------------------- #
def convert_docx(src: str, dst: str, fmt_name: str = "mendeley", keep_abstracts: bool = False,
                 keep_addin: bool = False, zotero_style: str = "http://www.zotero.org/styles/apa") -> dict:
    fmt = FORMATS[fmt_name](keep_abstracts=keep_abstracts, zotero_style=zotero_style)
    pkg = Package(src)
    stats = {"parts": [], "citations": 0, "bibliographies": 0, "unknown_tags": [], "addins_removed": 0}
    try:
        for name in pkg.names:
            if not PARTS_RE.match(name):
                continue
            xml = pkg.read(name)
            if TAG_PREFIX not in xml and BIB_TAG not in xml:
                continue
            new_xml, n_cit, n_bib, unknown = convert_part(xml, fmt)
            pkg.write(name, new_xml)
            stats["parts"].append(name)
            stats["citations"] += n_cit
            stats["bibliographies"] += n_bib
            stats["unknown_tags"] += unknown
        if stats["citations"] == 0 and stats["bibliographies"] == 0:
            raise ConversionError("no Mendeley Cite citations found in %s" % src)
        if not keep_addin:
            stats["addins_removed"] = pkg.strip_mendeley_addin()
        fmt.extra_package_edits(pkg)
        pkg.save(dst)
    finally:
        pkg.close()
    return stats


def default_output(src: str, suffix: str) -> str:
    root, _ = os.path.splitext(src)
    return "%s_%s.docx" % (root, suffix)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="Convert Mendeley Cite content-control citations in a .docx into legacy "
                    "Mendeley Desktop or Zotero field codes (readable by Paperpile and Zotero).")
    p.add_argument("input", help="input .docx written with Mendeley Cite")
    p.add_argument("-o", "--output", help="output .docx (default: <input>_<format>_fields.docx; "
                                          "ignored with --format both)")
    p.add_argument("-f", "--format", choices=["mendeley", "zotero", "both"], default="mendeley",
                   help="field-code flavour to write (default: mendeley)")
    p.add_argument("--bib", metavar="FILE", help="also export the cited references as BibTeX")
    p.add_argument("--dois", metavar="FILE", help="also export a plain list of cited DOIs")
    p.add_argument("--keep-abstracts", action="store_true",
                   help="keep abstracts inside the field codes (bigger file)")
    p.add_argument("--keep-addin", action="store_true",
                   help="do not remove the embedded Mendeley Cite task pane from the document")
    p.add_argument("--zotero-style", default="http://www.zotero.org/styles/apa",
                   help="style id stored in the Zotero document preferences (zotero format only)")
    p.add_argument("--version", action="version", version="%(prog)s " + __version__)
    a = p.parse_args(argv)

    if not os.path.isfile(a.input):
        p.error("input file not found: %s" % a.input)
    formats = ["mendeley", "zotero"] if a.format == "both" else [a.format]
    try:
        for f in formats:
            dst = a.output if (a.output and a.format != "both") else default_output(a.input, FORMATS[f].suffix)
            if os.path.abspath(dst) == os.path.abspath(a.input):
                p.error("output would overwrite the input; choose another -o path")
            st = convert_docx(a.input, dst, f, a.keep_abstracts, a.keep_addin, a.zotero_style)
            print("[%s] %s" % (f, dst))
            print("    citations converted : %d" % st["citations"])
            print("    bibliographies      : %d" % st["bibliographies"])
            print("    parts touched       : %s" % ", ".join(st["parts"]))
            if st["addins_removed"]:
                print("    Mendeley task panes removed: %d" % st["addins_removed"])
            if st["unknown_tags"]:
                print("    other content controls left untouched: %d" % len(st["unknown_tags"]))
        if a.bib or a.dois:
            items = collect_items(a.input)
            missing = [d.get("title", "?")[:70] for d in items.values() if not d.get("DOI")]
            if a.bib:
                print("[bib] %s  (%d unique references)" % (a.bib, write_bibtex(items, a.bib)))
            if a.dois:
                print("[dois] %s  (%d unique DOIs)" % (a.dois, write_dois(items, a.dois)))
            if missing:
                print("    references without a DOI (match these by hand): %s" % "; ".join(missing))
    except ConversionError as exc:
        print("error: %s" % exc, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
