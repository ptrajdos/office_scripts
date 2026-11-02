# Office Scripts

Each script exists in up to three forms:

| Task | VBA macro | Python (no Office needed) | AppleScript wrapper (runs the VBA macro) |
|---|---|---|---|
| Remove first word in column A | `column_with_spaces.vba` (Excel, `RemoveFirstWord`) | `column_with_spaces.py` | `osascript/column_with_spaces.applescript` |
| Merge last line, delete last page, strike phrases | `prepare_protocols.vba` (Word, `StrikeThroughPhrasesInDocs`) | `prepare_protocols.py` | `osascript/prepare_protocols.applescript` |
| Convert PDFs to DOCX | `protocol_saver.vba` (Word, `ConvertPDFsToDocx`) | `protocol_saver.py` | `osascript/protocol_saver.applescript` |
| Extract protocol info | `protocol_info_extractor.vba` (Word, `MakeDocsSummary`, stub) | `protocol_info.py` | `osascript/protocol_info_extractor.applescript` |

## VBA usage

* Open Word.
* Press **`Alt + F11`** → opens the VBA editor.
* In the Project window, expand **Normal** or your document’s name.
* Insert a new module →  **Insert > Module** .
* Paste your VBA code there.
* Run with  **F5** , or assign it to a button/shortcut.

## Python usage

```bash
python3 -m venv venv && venv/bin/pip install -r requirements.txt

venv/bin/python column_with_spaces.py workbook.xlsx [--sheet Sheet1] [--column A] [--output out.xlsx]
venv/bin/python prepare_protocols.py <dir_with_docx> [--phrase "..." ...]
venv/bin/python protocol_saver.py <dir_with_pdf>
venv/bin/python protocol_info.py <dir_with_docx> output.csv
```

Differences from the VBA versions:

* `protocol_saver.py` uses `pdf2docx` instead of Word's PDF import, so the DOCX layout may differ.
* `prepare_protocols.py` cannot compute page layout. "Last page" is everything after the last
  page boundary stored in the file (page break, section break, or the page marker Word writes on save).

## AppleScript usage

Install the macros once so Word/Excel always has them. Use separate modules, because
`MacSelectFolder` is defined in more than one script:

* Word: `Normal.dotm`, or a `.dotm` placed in
  `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/`
* Excel: Personal Macro Workbook

Then run from the terminal:

```bash
osascript osascript/prepare_protocols.applescript
osascript osascript/column_with_spaces.applescript /path/to/workbook.xlsx
```
