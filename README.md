# MendeleyCitePort

Move a Microsoft Word document from **Mendeley Cite** to **Paperpile** or
**Zotero** without replacing every citation by hand.

This tool creates a new copy of your `.docx` file with citations that Paperpile
or Zotero can recognize. Your document's visible text and formatting stay the
same, and the original file is not overwritten.

> [!IMPORTANT]
> Keep your original document until you have opened the converted copy,
> refreshed its citations, and checked the reference list.

## Who is this for?

Use this tool if all of the following are true:

- You wrote a Word document using the newer **Mendeley Cite** add-in.
- You want to continue managing its citations with **Paperpile** or **Zotero**.
- Paperpile or Zotero cannot detect the existing citations.

It works with Word `.docx` files. Old `.doc` files are not supported.

## What you need

- A Windows, macOS, or Linux computer
- [Python 3.8 or newer](https://www.python.org/downloads/)
- A copy of this project
- Your Word document, saved as a `.docx` file

No extra Python packages are required.

## Quick start

### 1. Download this tool

On this GitHub page, select **Code**, then **Download ZIP**. Extract the ZIP to
a folder on your computer.

### 2. Make a backup

Close the document in Word and make a separate backup copy. The converter
creates a new file, but keeping an untouched backup is still recommended.

### 3. Put your document beside the script

Move a copy of your `.docx` file into the folder containing
`mendeley_cite_port.py`.

For the examples below, the document is named `manuscript.docx`. Replace that
name with the exact name of your own file.

### 4. Open a command window in that folder

**Windows:** Open the folder in File Explorer, click the address bar, type
`powershell`, and press Enter.

**macOS:** Open Terminal, type `cd ` (including the space), drag the folder into
the Terminal window, and press Enter.

### 5. Create a converted copy

For **Paperpile**, run:

```powershell
py mendeley_cite_port.py "manuscript.docx"
```

For **Zotero**, run:

```powershell
py mendeley_cite_port.py "manuscript.docx" --format zotero
```

On macOS or Linux, use `python3` instead of `py`:

```bash
python3 mendeley_cite_port.py "manuscript.docx" --format zotero
```

Quotation marks are recommended, especially when the filename contains spaces.

### 6. Find the new document

The converted file appears in the same folder:

- Paperpile: `manuscript_LEGACY_MENDELEY_fields.docx`
- Zotero: `manuscript_ZOTERO_fields.docx`

The command window also reports how many citations and reference lists were
found. If it reports an error, see [Troubleshooting](#troubleshooting).

## Finish the move

### Paperpile

1. Make sure the cited papers are in your Paperpile library. You can import
   your Mendeley library into Paperpile first.
2. Open `manuscript_LEGACY_MENDELEY_fields.docx` in the desktop version of Word.
3. In the Paperpile tab, open **Settings and Tools → Convert from → Mendeley**.
4. Select **Convert**.
5. Choose your citation style, refresh the citations, and check the reference
   list against the original document.

If Paperpile says it found no citations, create both supported formats:

```powershell
py mendeley_cite_port.py "manuscript.docx" --format both
```

Then open `manuscript_ZOTERO_fields.docx` and choose
**Settings and Tools → Convert from → Zotero**.

### Zotero

1. Import your Mendeley library into Zotero using **File → Import → Mendeley**.
2. Open `manuscript_ZOTERO_fields.docx` in the desktop version of Word.
3. In the Zotero tab, select **Refresh**.
4. If prompted, link the citations to the matching items in your Zotero
   library.
5. Check several citations and the full reference list against the original.

References without a DOI may require manual matching in Paperpile or Zotero.

## Troubleshooting

### “Python was not found”

Install [Python 3](https://www.python.org/downloads/), then close and reopen the
command window. On Windows, make sure **Add Python to PATH** is selected during
installation. You can also try `python` instead of `py`.

### “Input file not found”

Check that the document is in the same folder as the script and that its name,
including `.docx`, exactly matches the command. Keep quotation marks around the
filename.

### “No Mendeley Cite citations found”

Confirm that the original document uses the newer **Mendeley Cite** Word add-in
and has been saved as a `.docx` file. This tool does not convert plain citation
text or documents already using old Mendeley Desktop fields.

### Word or a reference manager shows a warning

Do not discard the original. Close the converted document without saving and
try the other output format using `--format both`. Because these field formats
are not officially documented, some documents may need citations to be matched
manually after conversion.

### Citation numbers change after refreshing

Paperpile and Zotero regenerate numbering from their own libraries. Compare a
few citations and the reference list with the original after refreshing.

## Optional features

Create both Paperpile/Mendeley and Zotero copies:

```powershell
py mendeley_cite_port.py "manuscript.docx" --format both
```

Export the cited references as BibTeX and a list of DOIs at the same time:

```powershell
py mendeley_cite_port.py "manuscript.docx" --format both --bib references.bib --dois dois.txt
```

You can import `references.bib` into your new reference manager. The DOI list
can help you identify references that need manual matching.

<details>
<summary>All command options</summary>

```text
python mendeley_cite_port.py INPUT.docx [options]

  -o, --output FILE       Choose the output filename
  -f, --format FORMAT     mendeley (default), zotero, or both
      --bib FILE          Export cited references as BibTeX
      --dois FILE         Export cited DOIs as plain text
      --keep-abstracts    Keep abstracts in citation data (larger file)
      --keep-addin        Keep the embedded Mendeley Cite task pane
      --zotero-style URL  Set the Zotero citation-style identifier
  -h, --help              Show built-in help
```

The `--output` option is ignored when `--format both` is used because two files
are created.

</details>

## What the tool changes

Mendeley Cite stores citations in Word content controls. Paperpile and Zotero
cannot import those controls directly, but they can read older Word citation
fields. This tool changes only that hidden citation structure:

- Citations become classic Mendeley Desktop or Zotero fields.
- The reference list becomes a matching bibliography field.
- Citation text and formatting displayed in Word are preserved.
- The embedded Mendeley Cite task pane is removed by default so it does not
  reclaim the converted document.
- Footnotes, endnotes, headers, and footers are included.

The conversion uses the citation data stored inside each citation. Abstracts
are removed from the new fields by default to keep the document smaller.

## Known limitations

- Mendeley and Zotero do not officially document these internal Word formats.
  The converter has been tested with real Microsoft 365 documents, but results
  can vary.
- Text manually edited inside a Mendeley citation is preserved initially, but
  Paperpile or Zotero may regenerate it from the reference metadata.
- References without stable identifiers such as DOIs may need manual matching.
- Always review the converted document after your new reference manager
  refreshes it.

## For contributors

The project uses only the Python standard library. Run the test suite with:

```bash
python -m unittest discover tests
```

Bug reports and pull requests are welcome. For a useful bug report, include
your operating system, Python version, target reference manager, exact error
message, and the smallest sample document you can safely share. Remove personal
or confidential information before attaching a document.

## License

Created by Amin Akbari and released into the public domain under
[The Unlicense](LICENSE). You may use, modify, and share it without attribution
or conditions.
