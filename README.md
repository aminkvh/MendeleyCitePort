# Mendeley Cite Port

Convert citations made with the newer **Mendeley Cite** Word add-in into fields
that **Paperpile** or **Zotero** can recognize.

The script creates a new `.docx` file. It does not overwrite the original or
change the document's visible text.

## Requirements

- [Python 3.8 or newer](https://www.python.org/downloads/)
- A `.docx` file containing Mendeley Cite citations

No extra packages are required.

## Usage

Download this repository and put your Word document in the same folder as
`mendeley_cite_port.py`. Close the document in Word before continuing.

For **Paperpile**:

```powershell
py mendeley_cite_port.py "manuscript.docx"
```

For **Zotero**:

```powershell
py mendeley_cite_port.py "manuscript.docx" --format zotero
```

On macOS or Linux, use `python3` instead of `py`.

The new file appears beside the original:

- `manuscript_LEGACY_MENDELEY_fields.docx` for Paperpile
- `manuscript_ZOTERO_fields.docx` for Zotero

> Keep the original until you have checked the converted document.

## Finish the conversion

**Paperpile:** Open the new document in Word, then choose
**Paperpile → Settings and Tools → Convert from → Mendeley**.

**Zotero:** Import your Mendeley library into Zotero, open the Zotero-format
document in Word, then choose **Zotero → Refresh**.

References without a DOI may need to be matched manually.

## Other options

Create both formats and export the references:

```powershell
py mendeley_cite_port.py "manuscript.docx" --format both --bib references.bib --dois dois.txt
```

See all options:

```powershell
py mendeley_cite_port.py --help
```

The converter supports citations in the main document, footnotes, endnotes,
headers, and footers. Paperpile or Zotero may reformat citations when refreshed,
so compare the result with the original.

## Tests

```bash
python -m unittest discover tests
```

## License

Created by Amin Akbari and released under [The Unlicense](LICENSE).
