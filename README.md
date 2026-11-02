# optimize-pdf-qa-corpus

A QA test corpus for the **Office → optimized-PDF** conversion feature. It is a
set of Microsoft Word documents you upload through the admin tool to check that
each one produces a correct optimized-PDF preview — judged entirely as a black
box, with no engine, server, or database access required.

Scope right now is **Word (`.docx`) → PDF only**. PowerPoint/Excel and native
PDF sources are intentionally out of scope (PowerPoint and Excel support is
planned — see below).

## What's here

| Path | What it is |
|------|------------|
| `docx/` | The test documents, one file per case (42 files). Each prints its own test id / what-to-expect header on page 1. |
| `manifest/manifest.md` | Human-readable table: every file, what it checks, and the expected PDF result. |
| `manifest/manifest.csv` | The same manifest as CSV, for tracking a test pass. |
| `RESOURCES.md` | The full **tester guide** — how to run a pass, how to judge pass/fail, and what each group of files checks. |
| `generate_corpus.py` | The generator that (re)produces `docx/` and the manifest. |

## Quick start (running a QA pass)

1. Read `RESOURCES.md` — it is the tester guide.
2. In a test site/client with the optimized-PDF feature enabled, upload a file
   from `docx/` as a Document/resource.
3. Open its optimized-PDF preview in the admin tool and compare against the
   "Expected in PDF" note printed in the document (and in `manifest/manifest.md`).
4. Record the outcome per file (PDF produced? text correct? layout correct?
   thumbnail matches? anything stuck/looping?).

## What the corpus covers

- **Languages / scripts** (`lang_*`) — Latin, Cyrillic, Greek, Arabic, Hebrew
  (incl. niqqud), Devanagari, Thai, Chinese (simp/trad), Japanese, Korean,
  Vietnamese, plus a bidirectional mixed-direction line and an all-in-one sheet.
- **Emoji** (`emoji_basic`) — known-weak; records how emoji render.
- **Encoding stressors** (`enc_*`) — normalization, zero-width/control
  characters, rare astral-plane symbols.
- **Fonts** (`font_*`) — common paid Office fonts that may be substituted.
- **Layout / structure** (`struct_*`) — large tables, embedded images,
  header/footer, long paragraphs, many pages, headings & styles.
- **Conversion stress** (`stress_*`) — a dense document that makes the converter
  work hard (watch for timeouts / never-finishing).
- **Tricky inputs** (`edge_*`, `malformed_*`, `mislabeled_*`, `name_*`) — a
  nearly empty doc, a fake non-zip, a corrupt zip, a valid-zip-but-not-Word file,
  a zero-byte file, a mislabeled extension, and unusual file names.

See `manifest/manifest.md` for the complete list.

## Regenerating the corpus

```bash
python3 -m venv /tmp/qa-docx-venv
/tmp/qa-docx-venv/bin/pip install python-docx
/tmp/qa-docx-venv/bin/python generate_corpus.py
```

This rewrites `docx/` and `manifest/`. Extend it by adding to the `UDHR` list
(languages), the fonts list, or by following the existing case pattern.

## Hand-sourced extras (not auto-generated)

Some cases can't be produced by the script and must be added by hand — legacy
`.doc` files, password-protected documents, documents embedding specific paid
fonts, and additional languages. `RESOURCES.md` documents how to source these
from reliable, virus-free locations.

## Coming soon

PowerPoint (`.pptx`) and Excel (`.xlsx`) source support is planned. When it
lands, add matching cases (and their legacy `.ppt` / `.xls` twins);
`python-pptx` and `openpyxl` are the libraries to generate the modern ones.

## Note

All files are synthetic test fixtures — no private or production data. The
language samples are short public-domain excerpts from the Universal Declaration
of Human Rights (the standard i18n reference text).
