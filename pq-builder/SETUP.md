# PQ M-Code Builder — Setup

Windows Excel only (Office 2010+ / VBA7). Everything ships as plain-text `.bas`
files — nothing binary, nothing to hand-lay-out in a form designer.

## Why no `.frm` file

A real Office UserForm always serializes its control layout into a binary
`.frx` companion (unlike classic VB6 forms, which had a pure-text option).
That binary can't be hand-authored reliably outside the VBE, so instead
`modBuildForm.bas` builds the form itself, once, using VBA's own
project-editing API. You'll enable one Trust Center setting for that (step 2
below) — after that, everything else is a normal macro.

## Assembly

1. **Save first, in this exact folder.** Open Excel → new blank workbook →
   `File > Save As` → save it as `PQ Builder.xlsm` (Excel Macro-Enabled
   Workbook) **directly inside this `pq-builder` folder**, next to
   `templates\` and `pq.txt`. Setup and the golden test both resolve paths
   relative to the workbook's location, so this isn't optional.

2. **Enable VBA project access (one-time).** `File > Options > Trust Center >
   Trust Center Settings > Macro Settings` → check **"Trust access to the VBA
   project object model"** → OK, OK.

3. **Import the modules.** `Alt+F11` to open the VBE → `File > Import File…`
   → import, in any order:
   - `modPQBuilder.bas`
   - `modClipboard.bas`
   - `modBuildForm.bas`
   - `modSetup.bas`
   - `modTests.bas`

4. **Build the form.** In the Immediate window (`Ctrl+G`), run:
   ```
   Run_BuildForm
   ```
   This creates `frmPQBuilder` with all its controls and wires up its click
   handlers. You should see a confirmation message box. If you instead see a
   Trust Center error, go back to step 2.

5. **Seed the first template and run the tests.** Still in the Immediate
   window:
   ```
   Run_FirstTimeSetup
   Run_AllTests
   ```
   `Run_AllTests` should report `16 passed, 0 failed` (see Immediate window
   for the line-by-line breakdown). If the golden test fails, something in
   `templates\T_PreRunLag.txt` or `pq.txt` doesn't line up — see
   Troubleshooting.

6. **Save**, then run:
   ```
   ShowBuilder
   ```
   to open the form. Optionally add a Quick Access Toolbar button pointing at
   the `ShowBuilder` macro for one-click access later.

## Using it

1. `ShowBuilder`.
2. Pick a template from the dropdown.
3. Fill in the fields — each one's type is shown next to its name.
4. **Generate** — the M code appears in the preview box, or an error tells
   you which field is wrong.
5. **Copy to Clipboard**, then in your target workbook: `Data > Get Data >
   Launch Power Query Editor > New Source > Blank Query > Advanced Editor >`
   paste over the placeholder.

If the clipboard button ever fails silently (rare — it uses the raw Win32
clipboard API, no known failure mode short of another process holding the
clipboard open), the preview text is auto-selected — press `Ctrl+C` yourself.

## Token syntax

`{{Name}}`, `{{Name:type}}`, or `{{Name:type=default}}`. Untyped defaults to
`text`.

| Type | Validates | Emits |
|---|---|---|
| `text` | non-empty | `"value"`, embedded `"` doubled |
| `host` | IPv4 or hostname | `"10.0.0.1"` |
| `xlname` | valid Excel defined name, not a cell reference | `"ReportDate"` |
| `date` | real calendar date, `yyyy-mm-dd` | `#date(2026,9,4)` — bare |
| `datetime` | real date/time, `yyyy-mm-dd hh:mm:ss` | `#datetime(2026,9,4,6,0,0)` |
| `number` | numeric | `42` — bare |
| `raw` | non-empty only | verbatim, unescaped — use with care |

If a template genuinely needs the literal two-character sequence `{{` or
`}}` and *not* a token, split it with a space — `{ {` / `} }` — which is
equivalent M syntax and won't be matched.

## Adding a new template

Add a worksheet named `T_<YourName>`. Put the M code in column A, one line
per row (paste multi-line M straight in — Excel splits it across rows
automatically). Optional short description in `B1`, shown under the dropdown
once selected. That's it — the builder discovers `T_` sheets and their
tokens automatically; no code changes needed.

## Troubleshooting

**"Couldn't build the form" / Trust Center error on `Run_BuildForm`** — Step
2 above wasn't done, or was done in the wrong Excel profile. Re-check the
setting.

**`Run_FirstTimeSetup` can't find `T_PreRunLag.txt`** — the workbook isn't
saved in this folder (step 1), or the `templates\` subfolder got separated
from it. `templates\*.txt` is only needed once, at setup — the sheet is
authoritative afterward, so once setup succeeds you can move the `.xlsm`
anywhere.

**Golden test fails** — means the rendered `T_PreRunLag` sheet no longer
reproduces `pq.txt` byte-for-byte with the default values. Most likely cause:
`T_PreRunLag` was hand-edited after setup and a token or a line got
mangled. Delete the sheet and re-run `Run_FirstTimeSetup` to reseed it, or
inspect column A directly against `templates\T_PreRunLag.txt`.

**Pasted query errors on "Missing column" or won't find `ReportDate`** —
this is inherited from the original query's design, not something the
builder introduced: `Excel.CurrentWorkbook(){[Name="X"]}[Content]{0}[Column1]`
only resolves cleanly when `X` is a **single-cell named range with no
header row**. If it points at a table or a range with a header, the column
name won't be `Column1` and the query errors. Make sure the target workbook
has a plain single-cell named range matching whatever name you entered as
`DateRangeName`.

**A future template's `{{` collides with real M list-literal syntax
(`{{a, b}}`)** — in practice this doesn't happen: the token regex requires an
identifier immediately after `{{` with nothing else before the closing
`}}`, which M's `{{value, value}}` list-of-lists syntax never satisfies (it
either starts with a quote or has more than one bare identifier before the
close). If you hit an edge case anyway, use the `{ {` / `} }` spacing trick
above.
