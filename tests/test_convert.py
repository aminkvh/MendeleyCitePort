# -*- coding: utf-8 -*-
"""Round-trip test on a tiny synthetic .docx containing one Mendeley Cite citation
and one Mendeley Cite bibliography.  Run with:  python -m unittest discover tests"""
import base64
import json
import os
import re
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import mendeley_cite_port as m  # noqa: E402

CITE = {
    "citationID": "MENDELEY_CITATION_00000000-0000-0000-0000-000000000001",
    "properties": {"noteIndex": 0},
    "isEdited": False,
    "manualOverride": {"isManuallyOverridden": False, "citeprocText": "<sup>1</sup>", "manualOverrideText": ""},
    "citationItems": [{
        "id": "c1e3e6e3-80c1-31d7-afae-3fbcef1fcf85",
        "itemData": {"type": "article-journal", "id": "c1e3e6e3-80c1-31d7-afae-3fbcef1fcf85",
                     "title": "A test article", "author": [{"family": "Doe", "given": "Jane"}],
                     "container-title": "Journal of Tests", "issued": {"date-parts": [[2024]]},
                     "DOI": "10.1000/test.1", "abstract": "Long abstract text"},
    }],
}
TAG = m.TAG_PREFIX + base64.b64encode(json.dumps(CITE).encode("utf-8")).decode("ascii")

DOCUMENT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
    '<w:p><w:r><w:t xml:space="preserve">Some text </w:t></w:r>'
    '<w:sdt><w:sdtPr><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:tag w:val="%s"/><w:id w:val="1"/></w:sdtPr>'
    '<w:sdtContent><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r></w:sdtContent></w:sdt>'
    '<w:r><w:t>.</w:t></w:r></w:p>'
    '<w:sdt><w:sdtPr><w:tag w:val="MENDELEY_BIBLIOGRAPHY"/><w:id w:val="2"/></w:sdtPr><w:sdtContent>'
    '<w:p><w:pPr><w:ind w:hanging="640"/></w:pPr><w:r><w:t>1. Doe, J. A test article.</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>2. Another entry.</w:t></w:r></w:p>'
    '</w:sdtContent></w:sdt>'
    '<w:sdt><w:sdtPr><w:tag w:val="SOMETHING_ELSE"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>keep me</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:sectPr/></w:body></w:document>' % TAG
)
CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>'
)
RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>'
)


def make_docx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", DOCUMENT)


class ConvertTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.src = os.path.join(self.tmp, "in.docx")
        make_docx(self.src)

    def _convert(self, fmt):
        dst = os.path.join(self.tmp, "out_%s.docx" % fmt)
        stats = m.convert_docx(self.src, dst, fmt)
        with zipfile.ZipFile(dst) as z:
            names = z.namelist()
            doc = z.read("word/document.xml").decode("utf-8")
        return stats, doc, names, dst

    def test_mendeley_format(self):
        stats, doc, names, dst = self._convert("mendeley")
        self.assertEqual(stats["citations"], 1)
        self.assertEqual(stats["bibliographies"], 1)
        self.assertEqual(stats["unknown_tags"], ["SOMETHING_ELSE"])
        self.assertEqual(len(m.SDT_START.findall(doc)), 1)              # only the foreign sdt survives
        self.assertIn("ADDIN CSL_CITATION {", doc)
        self.assertIn("ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY", doc)
        self.assertNotIn("Long abstract text", doc)                      # abstracts stripped by default
        # field code round-trips as JSON with the classic Mendeley Desktop shape
        code = re.search(r"<w:instrText[^>]*> ADDIN CSL_CITATION (.*?) </w:instrText>", doc).group(1)
        payload = json.loads(m.unescape(code))
        self.assertEqual(payload["citationItems"][0]["id"], "ITEM-1")
        self.assertEqual(payload["citationItems"][0]["itemData"]["DOI"], "10.1000/test.1")
        self.assertTrue(payload["citationItems"][0]["uris"][0].startswith("http://www.mendeley.com/documents/?uuid="))
        self.assertEqual(payload["mendeley"]["plainTextFormattedCitation"], "1")
        # visible text is unchanged and the bibliography field spans both paragraphs
        self.assertEqual(m.text_of(doc), "Some text 1.1. Doe, J. A test article.2. Another entry.keep me")
        self.assertEqual(doc.count('w:fldCharType="begin"'), 2)
        self.assertEqual(doc.count('w:fldCharType="end"'), 2)
        self.assertLess(doc.find("CSL_BIBLIOGRAPHY"), doc.find("2. Another entry."))
        self.assertLess(doc.find("2. Another entry."), doc.rfind('w:fldCharType="end"'))
        self.assertNotIn("docProps/custom.xml", names)

    def test_zotero_format(self):
        stats, doc, names, dst = self._convert("zotero")
        self.assertEqual(stats["citations"], 1)
        self.assertIn("ADDIN ZOTERO_ITEM CSL_CITATION {", doc)
        self.assertIn("ADDIN ZOTERO_BIBL", doc)
        self.assertIn("docProps/custom.xml", names)                       # prefs part created
        with zipfile.ZipFile(dst) as z:
            custom = z.read("docProps/custom.xml").decode("utf-8")
            ct = z.read("[Content_Types].xml").decode("utf-8")
            rels = z.read("_rels/.rels").decode("utf-8")
        self.assertIn('name="ZOTERO_PREF_1"', custom)
        self.assertIn("/docProps/custom.xml", ct)
        self.assertIn("custom-properties", rels)

    def test_no_citations_is_an_error(self):
        empty = os.path.join(self.tmp, "empty.docx")
        with zipfile.ZipFile(empty, "w") as z:
            z.writestr("[Content_Types].xml", CONTENT_TYPES)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/document.xml", DOCUMENT.replace(TAG, "X").replace("MENDELEY_BIBLIOGRAPHY", "Y"))
        with self.assertRaises(m.ConversionError):
            m.convert_docx(empty, os.path.join(self.tmp, "o.docx"), "mendeley")

    def test_bibtex_and_dois(self):
        items = m.collect_items(self.src)
        bib = os.path.join(self.tmp, "r.bib")
        dois = os.path.join(self.tmp, "d.txt")
        self.assertEqual(m.write_bibtex(items, bib), 1)
        self.assertEqual(m.write_dois(items, dois), 1)
        text = open(bib, encoding="utf-8").read()
        self.assertIn("@article{Doe2024,", text)
        self.assertIn("doi = {10.1000/test.1}", text)

    def test_cli(self):
        out = os.path.join(self.tmp, "cli.docx")
        self.assertEqual(m.main([self.src, "-o", out, "--format", "mendeley"]), 0)
        self.assertTrue(os.path.exists(out))


if __name__ == "__main__":
    unittest.main()
