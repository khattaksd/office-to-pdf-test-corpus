# Office→PDF QA corpus — tester guide

A set of Microsoft Word documents for testing the "optimized PDF" feature end to
end from the **admin tool**, entirely as a black box: you upload a document, the
system produces an optimized PDF preview, and you check that the preview looks
right. This corpus covers the **Word (.docx)→PDF conversion path only**.
PowerPoint/Excel and native PDF sources are out of scope here.

You don't need any engine, server, or database knowledge to run this. Everything
is judged from what you can see in the admin tool.

- The test documents are in `docx/` (one file per case).
- `manifest/manifest.md` lists every file, what it checks, and what the PDF
  should look like.
- Each document also **prints its own instructions inside it** — open the file
  (or its PDF preview) and the top of page 1 shows the test id, what it checks,
  and what you should expect to see.

> A few cases can't be produced automatically and must be sourced by hand —
> legacy `.doc` files, password-protected documents, documents using specific
> paid fonts, and extra languages. See "Things to add by hand" at the end.
>
> **Coming soon:** PowerPoint (`.pptx`) and Excel (`.xlsx`) source support is
> planned. When it lands, add matching cases here (and their legacy `.ppt`/`.xls`
> twins) — they are intentionally **not** in this corpus yet.

---

## How to run a pass (from the admin tool)

1. In a test site/client where the optimized-PDF feature is turned on, upload a
   document from the corpus as you normally would (as a Document / resource).
2. Wait for the optimized PDF to be generated, then open the document's **PDF
   preview / optimized view** in the admin tool.
3. Compare what you see against the "Expected in PDF" line printed at the top of
   that document (and summarized in `manifest/manifest.md`).
4. Also check the **thumbnail / first-page image** the admin tool shows for the
   document — it should match page 1 of the PDF.

### What to record per file
- Did an optimized PDF get produced at all? (yes / no)
- Does the text render correctly — right letters, right direction, nothing
  showing as empty boxes ("▯▯▯")?
- Does the layout look right — tables, columns, headers/footers, images?
- Does the first-page thumbnail match the PDF?
- Anything stuck, duplicated, or never finishing.

Capture a screenshot of anything that looks wrong, plus the file id (printed in
the document and in the manifest).

### Pass / fail in plain terms
- **Pass** — an optimized PDF is produced and matches the "Expected in PDF" note.
- **Known-weak (not a fail)** — some cases are expected to look imperfect (e.g.
  emoji, or a paid font being swapped for a lookalike). Those say so in their
  note; record *how* they rendered rather than marking them failed.
- **Fail** — no PDF is produced for a normal document, text shows as boxes where
  it shouldn't, layout is broken, the thumbnail is wrong/blank, or processing
  never completes / keeps repeating.

> **Advanced users only (optional):** testers with access to the run-history
> table can additionally confirm the recorded result for each upload (e.g. a
> success vs. a rejected/invalid file). This is an optional cross-check — it is
> **not** required to run the pass, and most testers should judge purely from the
> admin-tool preview.

---

## What each group of files checks

### Languages / scripts (`lang_*`)
Checks that each language renders with the correct letters and reading direction.
What to look for:
- **Arabic, Hebrew** — read right-to-left; letters should be joined, not broken
  apart. `lang_bidi_mixed` mixes English, Arabic, and numbers on one line: the
  numbers and English must stay in sensible order inside the Arabic.
- **Hindi (Devanagari), Thai** — combined letter clusters and marks should sit in
  the right place, not scattered.
- **Chinese, Japanese, Korean** — characters should appear as real characters,
  not empty boxes.
- `lang_all_in_one` is a single quick-scan sheet with every language on it.

If a whole language shows up as boxes, that's a real finding — note which
language.

### Emoji (`emoji_basic`) — known-weak
Emoji often come out as plain black-and-white outlines or boxes instead of color.
That is expected; **record how they actually look** (color / outline / boxes)
rather than failing the case. Family and flag emoji are the most likely to split
into separate pieces.

### Special characters / encoding (`enc_*`)
- `enc_normalization` — the same accented words written two different ways; both
  lines should look identical.
- `enc_zero_width_and_controls` — invisible formatting characters; should produce
  no stray symbols and not visibly chop words.
- `enc_astral_surrogates` — rare math/symbol characters; main thing is the PDF
  **still gets produced** (odd glyphs or boxes are acceptable; a failure to
  convert is not).

### Fonts (`font_*`)
Each uses a common paid Office font (Calibri, Cambria, Arial, Times New Roman,
Verdana, Segoe UI). The system may not have the exact font and will swap in a
lookalike. **Record which ones look visibly different** from a normal document
(wider/narrower text, different line breaks). Some swaps are clean; some shift
the layout. This is expected behavior to document, not a hard fail.

### Layout / structure (`struct_*`)
- `large_table` — a big table; should flow across pages without cut-off columns.
- `embedded_image` — contains a picture; it should appear in the PDF **and** in
  the first-page thumbnail.
- `header_footer` — a header and footer; should appear on every page.
- `long_paragraph` — one huge paragraph; should wrap, not run off the page.
- `many_pages` — a long (~130 page) document; should finish and show all pages.
  If a long document never finishes, note it.
- `headings_links` — headings and bold/italic/underline; styling should be kept.

### Conversion stress (`stress_*`) — known-heavy
- `huge_table` — a very large, dense table meant to make the conversion step
  work hard. It should still **finish** and produce a PDF; the thing to watch is
  whether it stalls, times out, or never completes. Record the outcome and
  roughly how long it took.

### Tricky inputs (`edge_*`, `malformed_*`, `mislabeled_*`, `name_*`)
- `edge_minimal` — a nearly empty document. It should **still produce an
  optimized PDF** (even though there's almost nothing to shrink).
- `malformed_not_a_zip` — a fake document that isn't really a Word file. The
  system should **reject it cleanly** (no PDF, no error loop) — it should not
  crash or get stuck.
- `malformed_corrupt_zip` — a damaged Word file. Expect it to **fail cleanly**
  and not keep retrying forever or repeatedly re-processing.
- `malformed_zip_not_ooxml` — a *valid* zip archive that is **not** a Word file
  inside (plain text files in a zip renamed to `.docx`). It looks like a Word
  file to a quick check but isn't one; it should **fail cleanly**, not convert
  to a bogus PDF or get stuck.
- `malformed_empty` — a **zero-byte** file with a `.docx` name. Expect a clean
  rejection: no PDF, no crash, no retry loop.
- `mislabeled_docx_as_doc.doc` — a modern Word file that was given an old-style
  `.doc` name. Record whether it still converts.
- `résumé_整理_ملخص.docx`, `report_📊_v2.docx`, and the very long file name —
  these test **unusual file names** (accents, emoji, long, spaces). They should
  upload and convert with the name handled gracefully.

---

## Things to add by hand (can't be produced automatically)

### 1. Old-style `.doc` files
The corpus is all modern Word `.docx`. To also test the older binary `.doc`
format, open a file in **Microsoft Word → Save As → "Word 97-2004 Document
(.doc)"**. The manifest marks which files are worth converting (the "`.doc` twin"
column) — prioritize the right-to-left languages (Arabic/Hebrew) and a couple of
Chinese/Japanese files, since the old format is most likely to cause problems
with those.

If you'd rather not create them in Word, use ready-made samples from a
**reliable, virus-free** source:
- [py-pdf/sample-files](https://github.com/py-pdf/sample-files) — curated test
  documents in a well-known open-source project.
- Microsoft's own [Office sample documents](https://www.microsoft.com/en-in/download/details.aspx?id=34197)
  on the Microsoft Download Center.
- [file-examples.com](https://file-examples.com/index.php/sample-documents-download/sample-doc-download/)
  and [freetestdata.com](https://freetestdata.com/document-files/doc/) state
  their files are virus-clean; prefer the two sources above for anything
  committed to the repo. (Content was rephrased for compliance with licensing restrictions.)

### 2. Password-protected / encrypted documents
A document saved with a Microsoft Office **open-password** can't be produced by
the generator. Create one in Office (**Save As → Tools → General Options → set a
password to open**, or **File → Info → Protect Document → Encrypt with
Password**). The converter can't open it without the password, so the expected
result is a **clean failure** (no PDF, no crash, no endless retry) — record what
the admin tool reports.

### 3. Documents using specific paid fonts
To test a real paid font (rather than a lookalike swap), create the document in
**Microsoft Word with that font** and turn on **File → Options → Save → "Embed
fonts in the file"** so the font travels with the document. Then upload it and
check the preview.

### 4. More languages
The language files use short lines from the Universal Declaration of Human Rights
(a standard, free, professionally translated reference text available in hundreds
of languages). To cover more languages, grab the same sentence in another
language from a public UDHR translation source and add it to the generator:
- [UDHR in Unicode / UDHR in XML](https://unicode.org/udhr/) — plain-text and XML
  for 500+ languages (the canonical i18n reference this corpus draws from).
- [OHCHR official translations](https://www.ohchr.org/en/human-rights/universal-declaration/universal-declaration-human-rights/)
  — authoritative per-language text from the UN.
- [thammegowda/014-udhr-dataset](https://github.com/thammegowda/014-udhr-dataset)
  — a clean mirror of the plain-text corpus, convenient for scripting.

> These are **text** sources — you paste strings into `generate_corpus.py`, so
> there's no binary download to virus-check. Avoid "malicious PDF" / "stressful
> PDF" corpora you may find while searching; they're built to trigger exploits
> and are the wrong tool (and wrong format) for this conversion-QA corpus.

---

## Regenerating / extending the corpus (for the person maintaining it)

```
python3 -m venv /tmp/qa-docx-venv
/tmp/qa-docx-venv/bin/pip install python-docx
/tmp/qa-docx-venv/bin/python generate_corpus.py
```

Add languages by extending the `UDHR` list; add fonts via the fonts list; add new
layout cases by following the existing pattern in `generate_corpus.py`.

> **When PowerPoint/Excel support lands:** add `.pptx`/`.xlsx` cases (and legacy
> `.ppt`/`.xls` twins). `python-pptx` and `openpyxl` are the libraries to use for
> generating the modern ones.
