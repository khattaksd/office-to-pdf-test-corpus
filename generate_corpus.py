#!/usr/bin/env python3
"""
Generate an MS-Office QA corpus for testing the optimized-PDF feature from the
admin tool. Covers the Office->PDF conversion path only (native PDFs are out of
scope for this corpus).

Each file prints a visible header (id, what it checks, what to expect) at the top
of page 1 so a tester can judge pass/fail from the admin-tool PDF preview alone. A
manifest.csv / manifest.md is written alongside. See RESOURCES.md for the tester
guide.

NOTE: this writes modern Word .docx only. PowerPoint/Excel (.pptx/.xlsx) and the
legacy binary formats (.doc/.ppt/.xls) are out of scope for now -- PowerPoint and
Excel support is planned, so leave room to add those cases later. Old-style .doc
files, password-protected documents, documents using specific paid fonts, and
additional languages must be sourced by hand -- see RESOURCES.md.
"""
import csv
import struct
import zipfile
import zlib
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor, Inches

OUT = Path(__file__).resolve().parent
DOCX = OUT / "docx"
DOCX.mkdir(parents=True, exist_ok=True)

# UDHR Article 1 excerpts -- the standard i18n reference text, public domain.
# Short, human-translated, covers the target scripts.
UDHR = {
    "latin_en": "All human beings are born free and equal in dignity and rights.",
    "latin_diacritics": "Tä schøn café naïve coöperate Ærøskøbing — Zürich, Œuvre, ﬁ ﬂ ligatures.",
    "cyrillic_ru": "Все люди рождаются свободными и равными в своём достоинстве и правах.",
    "greek_el": "Όλοι οι άνθρωποι γεννιούνται ελεύθεροι και ίσοι στην αξιοπρέπεια.",
    "arabic_ar": "يولد جميع الناس أحرارًا متساوين في الكرامة والحقوق.",
    "hebrew_he": "כל בני האדם נולדו בני חורין ושווים בערכם ובזכויותיהם.",
    "hebrew_niqqud": "כָּל בְּנֵי הָאָדָם נוֹלְדוּ בְּנֵי חוֹרִין וְשָׁוִים.",
    "devanagari_hi": "सभी मनुष्य जन्म से स्वतंत्र तथा मर्यादा और अधिकारों में समान होते हैं।",
    "thai_th": "มนุษย์ทั้งหลายเกิดมามีอิสระและเสมอภาคกันในศักดิ์ศรีและสิทธิ",
    "chinese_simp": "人人生而自由，在尊严和权利上一律平等。",
    "chinese_trad": "人人生而自由，在尊嚴和權利上一律平等。",
    "japanese": "すべての人間は、生まれながらにして自由であり、かつ、尊厳と権利について平等である。",
    "korean": "모든 인간은 태어날 때부터 자유로우며 그 존엄과 권리에 있어 동등하다.",
    "vietnamese": "Tất cả mọi người sinh ra đều được tự do và bình đẳng về nhân phẩm và quyền lợi.",
}

WINDOWS_FONTS = ["Calibri", "Cambria", "Arial", "Times New Roman", "Verdana", "Segoe UI"]

# A valid 10x10 red PNG, built with correct CRC32 chunks so python-docx accepts it.
def _make_png(width=10, height=10, rgb=(220, 40, 40)):
    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)  # 8-bit RGB
    row = b"\x00" + bytes(rgb) * width
    idat = zlib.compress(row * height, 9)
    return sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


PNG_1x1 = _make_png()


def new_doc(test_id, title, tests, expected):
    """A document pre-seeded with a visible test header."""
    doc = Document()
    h = doc.add_paragraph()
    r = h.add_run(f"[{test_id}] {title}")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0x11, 0x44, 0xAA)
    meta = doc.add_paragraph()
    meta.add_run("Tests: ").bold = True
    meta.add_run(tests)
    exp = doc.add_paragraph()
    exp.add_run("Expected in PDF: ").bold = True
    exp.add_run(expected)
    doc.add_paragraph("-" * 60)
    return doc


MANIFEST = []


def emit(doc, test_id, title, tests, expected, needs_doc_twin="no", notes=""):
    path = DOCX / f"{test_id}.docx"
    doc.save(path)
    MANIFEST.append(
        dict(id=test_id, file=path.name, title=title, tests=tests,
             expected=expected, doc_twin_needed=needs_doc_twin, notes=notes)
    )
    print(f"  wrote {path.name}")


# --- 1. Multilingual / scripts -------------------------------------------------
for key, text in UDHR.items():
    rtl = key.startswith(("arabic", "hebrew"))
    tid = f"lang_{key}"
    doc = new_doc(
        tid, f"Script: {key}",
        f"{'RTL ' if rtl else ''}script rendering, correct glyphs, no tofu",
        "text matches the source line below with correct shaping / direction",
    )
    doc.add_paragraph(text)
    emit(doc, tid, f"Script {key}", "language renders with correct letters and direction",
         "native-script text, readable, no empty boxes",
         needs_doc_twin=("yes" if rtl else "no"),
         notes="right-to-left" if rtl else "")

# Bidi: mixed LTR + RTL on one line
doc = new_doc("lang_bidi_mixed", "Bidirectional (LTR+RTL mixed)",
              "bidi ordering of Arabic + English + digits on one line",
              "visual order correct: Arabic runs RTL, English LTR, numbers intact")
doc.add_paragraph("Invoice رقم 12345 due الاثنين — thanks شكرا!")
emit(doc, "lang_bidi_mixed", "Bidi mixed", "mixed right-to-left and left-to-right on one line",
     "numbers and English stay sensibly placed inside the Arabic", needs_doc_twin="yes")

# All scripts stacked in one doc (quick smoke sheet)
doc = new_doc("lang_all_in_one", "All scripts (smoke sheet)",
              "every script in one PDF for a fast visual scan",
              "each line legible in its own script")
for key, text in UDHR.items():
    p = doc.add_paragraph()
    p.add_run(f"{key}: ").bold = True
    p.add_run(text)
emit(doc, "lang_all_in_one", "All scripts", "quick scan of every language", "every line readable")

# --- 2. Emoji ------------------------------------------------------------------
doc = new_doc("emoji_basic", "Emoji: basic + modifiers + family + flags",
              "how emoji render (often plain outlines or boxes, not color)",
              "KNOWN-WEAK: record how they look (color / outline / boxes), don't fail")
doc.add_paragraph("Basic:  😀 🎉 🔥 ✅ ⚠️ ☕")
doc.add_paragraph("Skin tone:  👍🏻 👍🏽 👍🏿")
doc.add_paragraph("ZWJ family:  👨‍👩‍👧‍👦   couple:  👩‍❤️‍👨")
doc.add_paragraph("Flags:  🇺🇸 🇯🇵 🇸🇦 🏳️‍🌈")
doc.add_paragraph("Keycap/seq:  1️⃣ #️⃣   profession:  👩‍🚀 🧑‍💻")
emit(doc, "emoji_basic", "Emoji", "emoji / family / flags rendering",
     "known-weak: record how they look (color, outline, or boxes)", notes="expected imperfect; document behavior")

# --- 3. Encoding / unicode stressors ------------------------------------------
doc = new_doc("enc_normalization", "Accents written two ways",
              "same accented words encoded two different ways",
              "both lines look identical")
doc.add_paragraph("NFC precomposed: \u00e9 \u00f1 \u00fc  (café mañana)")
doc.add_paragraph("NFD combining:   e\u0301 n\u0303 u\u0308  (cafe\u0301 man\u0303ana)")
emit(doc, "enc_normalization", "Accents two ways", "accent encoding", "both lines identical")

doc = new_doc("enc_zero_width_and_controls", "Invisible formatting characters",
              "hidden zero-width and direction characters",
              "no stray symbols; words not visibly chopped")
doc.add_paragraph("ZWSP between\u200bwords  |  ZWNJ: \u0634\u200c\u0634  |  soft\u00adhyphen")
doc.add_paragraph("Bidi marks: \u202aLRE forced ltr\u202c \u202bRLE forced rtl\u202c")
emit(doc, "enc_zero_width_and_controls", "Invisible chars", "hidden characters",
     "no stray symbols, no visible artifacts")

doc = new_doc("enc_astral_surrogates", "Rare symbol characters",
              "rare math / symbol characters outside the common range",
              "PDF is still produced — odd glyphs or boxes are OK, a failed convert is not")
doc.add_paragraph("Math bold: \U0001D54F \U0001D54A   Alchemy: \U0001F701 \U0001F702")
doc.add_paragraph("Rare CJK ext-B: \U00020BB7 \U0002A6B2")
emit(doc, "enc_astral_surrogates", "Rare symbols", "rare characters",
     "PDF produced; glyphs or boxes both acceptable")

# --- 4. Fonts (requested Windows fonts -> substitution on the Linux worker) ----
for font in WINDOWS_FONTS:
    tid = f"font_{font.lower().replace(' ', '_')}"
    doc = new_doc(tid, f"Font requested: {font}",
                  f"'{font}' is a common paid font that may not be installed",
                  f"may be swapped for a lookalike — record if text looks visibly different")
    p = doc.add_paragraph()
    run = p.add_run(f"The quick brown fox 0123456789 — {font} 12pt. ABCDEFG abcdefg")
    run.font.name = font
    run.font.size = Pt(12)
    emit(doc, tid, f"Font {font}", "common paid font (may be swapped)",
         "lookalike swap OK; note any visible difference / changed line breaks", notes="paid font may not be installed")

# --- 5. Structural / layout / pipeline edge cases ------------------------------
# Large table
doc = new_doc("struct_large_table", "Large table (50x8)",
              "table pagination and column fit in PDF",
              "table spans pages cleanly, no clipped columns")
t = doc.add_table(rows=50, cols=8)
t.style = "Table Grid"
for i, row in enumerate(t.rows):
    for j, cell in enumerate(row.cells):
        cell.text = f"r{i}c{j}"
emit(doc, "struct_large_table", "Large table", "table layout/pagination", "clean multi-page table")

# Embedded image
doc = new_doc("struct_embedded_image", "Embedded image (PNG)",
              "a document containing a picture",
              "image appears in the PDF and in the first-page thumbnail")
img_path = DOCX / "_tmp_1x1.png"
img_path.write_bytes(PNG_1x1)
doc.add_paragraph("Image below:")
doc.add_picture(str(img_path), width=Inches(2))
img_path.unlink(missing_ok=True)
emit(doc, "struct_embedded_image", "Embedded image", "image + thumbnail", "image present on page 1 and in the thumbnail")

# Header/footer + page numbers
doc = new_doc("struct_header_footer", "Header / footer with page text",
              "running header/footer rendered in PDF",
              "header and footer appear on every page")
sec = doc.sections[0]
sec.header.paragraphs[0].text = "HEADER — confidential QA sample"
sec.footer.paragraphs[0].text = "FOOTER — page of pages"
for _ in range(60):
    doc.add_paragraph("Body line to force multiple pages. " * 3)
emit(doc, "struct_header_footer", "Header/footer", "running header/footer", "present on all pages")

# Very long single paragraph
doc = new_doc("struct_long_paragraph", "One very long paragraph",
              "line-breaking of a huge unbroken paragraph",
              "wraps normally, no overflow off the page")
doc.add_paragraph(("lorem ipsum dolor sit amet " * 400).strip())
emit(doc, "struct_long_paragraph", "Long paragraph", "wrapping", "no overflow")

# Many pages (pagination + possible timeout)
doc = new_doc("struct_many_pages", "100+ pages",
              "a long document (many pages)",
              "finishes and shows all pages; note if it never completes")
for i in range(1, 130):
    doc.add_paragraph(f"Page-filler block {i}. " + ("content " * 40))
    if i % 2 == 0:
        doc.add_page_break()
emit(doc, "struct_many_pages", "Many pages", "long document",
     "completes and shows all pages", notes="watch for slow / never-finishing")

# Hyperlinks + footnote-like + TOC-ish headings
doc = new_doc("struct_headings_links", "Headings + hyperlink text + styles",
              "heading styles, bold/italic/underline, link text",
              "styles preserved; heading hierarchy visible")
doc.add_heading("Heading 1", level=1)
doc.add_heading("Heading 2", level=2)
p = doc.add_paragraph()
p.add_run("bold ").bold = True
p.add_run("italic ").italic = True
u = p.add_run("underline")
u.underline = True
doc.add_paragraph("Visit https://example.com/qa for details (link as plain text).")
emit(doc, "struct_headings_links", "Headings/styles", "styles/headings", "preserved hierarchy & styles")

# Minimal near-empty doc
doc = new_doc("edge_minimal", "Minimal doc (header only)",
              "a nearly empty document",
              "an optimized PDF is still produced (even with almost no content)")
emit(doc, "edge_minimal", "Minimal", "nearly empty document",
     "optimized PDF still produced", notes="")

# Filename edge cases (content trivial; the FILENAME is the test)
for tid, fname, desc in [
    ("name_unicode", "résumé_整理_ملخص.docx", "unicode filename"),
    ("name_emoji", "report_📊_v2.docx", "emoji in filename"),
    ("name_spaces_long", "a very long file name with spaces " * 4 + ".docx", "long name + spaces"),
]:
    doc = new_doc(tid, f"Filename edge: {desc}",
                  "an unusual file name (accents / emoji / long / spaces)",
                  "uploads and converts; the file name is handled gracefully")
    doc.add_paragraph(f"Filename under test: {fname}")
    path = DOCX / fname
    doc.save(path)
    MANIFEST.append(dict(id=tid, file=fname, title=f"Filename {desc}",
                         tests="unusual file name", expected="uploads & converts cleanly",
                         doc_twin_needed="no", notes="file name is the subject"))
    print(f"  wrote {fname}")

# Content-Disposition filename cases: the optimized PDF is stored on S3 with a
# Content-Disposition header built from the source filename. These names have
# characters that must be quoted/encoded in that header (spaces, '%', quotes,
# commas, accents) -- the test is the filename shown when downloading/opening
# the optimized PDF, not whether conversion succeeds.
for tid, fname, desc in [
    ("name_cd_spaces", "Q4 Sales Report.docx", "spaces in filename"),
    ("name_cd_percent", "50% complete draft.docx", "percent sign in filename"),
    ("name_cd_quotes_commas", 'a "quoted", comma.docx',
     "quotes and commas in filename"),
    ("name_cd_accents", "Rapport été 2024.docx", "accents + spaces in filename"),
]:
    doc = new_doc(tid, f"Content-Disposition filename: {desc}",
                  "the download filename of the optimized PDF "
                  "(Content-Disposition header)",
                  "the optimized PDF downloads/opens with the ORIGINAL name intact "
                  "(spaces and special characters preserved), not truncated at the "
                  "first space/quote and not shown as a garbled or '_'-only name")
    doc.add_paragraph(f"Filename under test: {fname}")
    path = DOCX / fname
    doc.save(path)
    MANIFEST.append(dict(id=tid, file=fname, title=f"Content-Disposition {desc}",
                         tests="download filename of the optimized PDF",
                         expected="downloads with the original name intact "
                                  "(special chars preserved)",
                         doc_twin_needed="no",
                         notes="targets the Content-Disposition filename fix"))
    print(f"  wrote {fname}")

# --- 6. Malformed inputs (should be rejected / fail cleanly) -------------------
# A text file masquerading as .docx
bad = DOCX / "malformed_not_a_zip.docx"
bad.write_text("This is plain text pretending to be a .docx. Not a real Word file.\n")
MANIFEST.append(dict(id="malformed_not_a_zip", file=bad.name,
                     title="Fake Word file (plain text)",
                     tests="not really a Word file",
                     expected="rejected cleanly: no PDF, no crash, no retry loop",
                     doc_twin_needed="no", notes="negative test"))
print(f"  wrote {bad.name}")

# A truncated/corrupt docx
corrupt = DOCX / "malformed_corrupt_zip.docx"
sample = (DOCX / "edge_minimal.docx").read_bytes()
corrupt.write_bytes(sample[: max(64, len(sample) // 3)])
MANIFEST.append(dict(id="malformed_corrupt_zip", file=corrupt.name,
                     title="Damaged Word file",
                     tests="damaged / unreadable Word file",
                     expected="fails cleanly; does NOT keep retrying or reprocessing",
                     doc_twin_needed="no", notes="negative test"))
print(f"  wrote {corrupt.name}")

# A valid docx renamed to .doc
mis = DOCX / "mislabeled_docx_as_doc.doc"
mis.write_bytes((DOCX / "lang_latin_en.docx").read_bytes())
MANIFEST.append(dict(id="mislabeled_docx_as_doc", file=mis.name,
                     title="Modern Word file with an old .doc name",
                     tests="file extension does not match the real format",
                     expected="record whether it still converts",
                     doc_twin_needed="no", notes="edge: name vs real format"))
print(f"  wrote {mis.name}")

# A real ZIP archive that is NOT an OOXML Word document. Passes the "PK\x03\x04"
# header sniff (it is a valid zip) and is not truncated, so it reaches the
# converter -- which must still reject it cleanly rather than treat it as Word.
notooxml = DOCX / "malformed_zip_not_ooxml.docx"
with zipfile.ZipFile(notooxml, "w", zipfile.ZIP_DEFLATED) as zf:
    zf.writestr("readme.txt", "A valid zip, but not a Word (OOXML) package.\n")
    zf.writestr("data/notes.txt", "No [Content_Types].xml, no word/document.xml.\n")
MANIFEST.append(dict(id="malformed_zip_not_ooxml", file=notooxml.name,
                     title="Valid zip that is not a Word document",
                     tests="valid archive, wrong internal format",
                     expected="rejected / fails cleanly: no PDF, no crash, no retry loop",
                     doc_twin_needed="no",
                     notes="negative test: passes zip header sniff but is not OOXML"))
print(f"  wrote {notooxml.name}")

# A zero-byte file with a .docx name: empty source, no readable header.
empty = DOCX / "malformed_empty.docx"
empty.write_bytes(b"")
MANIFEST.append(dict(id="malformed_empty", file=empty.name,
                     title="Zero-byte file",
                     tests="empty source file (0 bytes)",
                     expected="rejected / fails cleanly: no PDF, no crash, no retry loop",
                     doc_twin_needed="no", notes="negative test: empty file"))
print(f"  wrote {empty.name}")

# --- 7. Conversion stressors (probe the convert timeout / resource limits) -----
# A pathological document: a very large, dense table plus heavy text, built to
# make the headless-LibreOffice convert step work hard. Testers watch for a
# never-finishing / timed-out conversion rather than a visual defect.
doc = new_doc("stress_huge_table", "Conversion stress: huge dense table",
              "a very large, dense table that makes conversion work hard",
              "KNOWN-HEAVY: should finish within the convert timeout; "
              "note if it stalls, times out, or never completes")
t = doc.add_table(rows=400, cols=10)
t.style = "Table Grid"
for i, row in enumerate(t.rows):
    for j, cell in enumerate(row.cells):
        cell.text = f"r{i}c{j} " + ("lorem ipsum " * 3)
emit(doc, "stress_huge_table", "Stress huge table", "conversion time / timeout",
     "finishes within convert timeout; note stalls or timeouts",
     notes="watch for slow / timed-out / never-finishing conversion")

# --- manifest ------------------------------------------------------------------
cols = ["id", "file", "title", "tests", "expected", "doc_twin_needed", "notes"]
with (OUT / "manifest" / "manifest.csv").open("w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=cols)
    w.writeheader()
    w.writerows(MANIFEST)

with (OUT / "manifest" / "manifest.md").open("w") as f:
    f.write("# QA corpus manifest\n\n")
    f.write(f"{len(MANIFEST)} documents. See RESOURCES.md for the tester guide and "
            "the hand-sourced extras (legacy .doc files, password-protected "
            "documents, paid fonts, more languages).\n\n")
    f.write("| id | file | what it checks | expected in the PDF preview | .doc twin |\n")
    f.write("|----|------|-------|-----------------|-----------|\n")
    for m in MANIFEST:
        f.write(f"| {m['id']} | `{m['file']}` | {m['tests']} | {m['expected']} | {m['doc_twin_needed']} |\n")

print(f"\nTotal: {len(MANIFEST)} files -> {DOCX}")
print(f"Manifest: {OUT / 'manifest'}")
