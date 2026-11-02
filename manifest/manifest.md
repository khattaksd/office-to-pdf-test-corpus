# QA corpus manifest

42 documents. See RESOURCES.md for the tester guide and the hand-sourced extras (legacy .doc files, password-protected documents, paid fonts, more languages).

| id | file | what it checks | expected in the PDF preview | .doc twin |
|----|------|-------|-----------------|-----------|
| lang_latin_en | `lang_latin_en.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | no |
| lang_latin_diacritics | `lang_latin_diacritics.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | no |
| lang_cyrillic_ru | `lang_cyrillic_ru.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | no |
| lang_greek_el | `lang_greek_el.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | no |
| lang_arabic_ar | `lang_arabic_ar.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | yes |
| lang_hebrew_he | `lang_hebrew_he.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | yes |
| lang_hebrew_niqqud | `lang_hebrew_niqqud.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | yes |
| lang_devanagari_hi | `lang_devanagari_hi.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | no |
| lang_thai_th | `lang_thai_th.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | no |
| lang_chinese_simp | `lang_chinese_simp.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | no |
| lang_chinese_trad | `lang_chinese_trad.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | no |
| lang_japanese | `lang_japanese.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | no |
| lang_korean | `lang_korean.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | no |
| lang_vietnamese | `lang_vietnamese.docx` | language renders with correct letters and direction | native-script text, readable, no empty boxes | no |
| lang_bidi_mixed | `lang_bidi_mixed.docx` | mixed right-to-left and left-to-right on one line | numbers and English stay sensibly placed inside the Arabic | yes |
| lang_all_in_one | `lang_all_in_one.docx` | quick scan of every language | every line readable | no |
| emoji_basic | `emoji_basic.docx` | emoji / family / flags rendering | known-weak: record how they look (color, outline, or boxes) | no |
| enc_normalization | `enc_normalization.docx` | accent encoding | both lines identical | no |
| enc_zero_width_and_controls | `enc_zero_width_and_controls.docx` | hidden characters | no stray symbols, no visible artifacts | no |
| enc_astral_surrogates | `enc_astral_surrogates.docx` | rare characters | PDF produced; glyphs or boxes both acceptable | no |
| font_calibri | `font_calibri.docx` | common paid font (may be swapped) | lookalike swap OK; note any visible difference / changed line breaks | no |
| font_cambria | `font_cambria.docx` | common paid font (may be swapped) | lookalike swap OK; note any visible difference / changed line breaks | no |
| font_arial | `font_arial.docx` | common paid font (may be swapped) | lookalike swap OK; note any visible difference / changed line breaks | no |
| font_times_new_roman | `font_times_new_roman.docx` | common paid font (may be swapped) | lookalike swap OK; note any visible difference / changed line breaks | no |
| font_verdana | `font_verdana.docx` | common paid font (may be swapped) | lookalike swap OK; note any visible difference / changed line breaks | no |
| font_segoe_ui | `font_segoe_ui.docx` | common paid font (may be swapped) | lookalike swap OK; note any visible difference / changed line breaks | no |
| struct_large_table | `struct_large_table.docx` | table layout/pagination | clean multi-page table | no |
| struct_embedded_image | `struct_embedded_image.docx` | image + thumbnail | image present on page 1 and in the thumbnail | no |
| struct_header_footer | `struct_header_footer.docx` | running header/footer | present on all pages | no |
| struct_long_paragraph | `struct_long_paragraph.docx` | wrapping | no overflow | no |
| struct_many_pages | `struct_many_pages.docx` | long document | completes and shows all pages | no |
| struct_headings_links | `struct_headings_links.docx` | styles/headings | preserved hierarchy & styles | no |
| edge_minimal | `edge_minimal.docx` | nearly empty document | optimized PDF still produced | no |
| name_unicode | `résumé_整理_ملخص.docx` | unusual file name | uploads & converts cleanly | no |
| name_emoji | `report_📊_v2.docx` | unusual file name | uploads & converts cleanly | no |
| name_spaces_long | `a very long file name with spaces a very long file name with spaces a very long file name with spaces a very long file name with spaces .docx` | unusual file name | uploads & converts cleanly | no |
| malformed_not_a_zip | `malformed_not_a_zip.docx` | not really a Word file | rejected cleanly: no PDF, no crash, no retry loop | no |
| malformed_corrupt_zip | `malformed_corrupt_zip.docx` | damaged / unreadable Word file | fails cleanly; does NOT keep retrying or reprocessing | no |
| mislabeled_docx_as_doc | `mislabeled_docx_as_doc.doc` | file extension does not match the real format | record whether it still converts | no |
| malformed_zip_not_ooxml | `malformed_zip_not_ooxml.docx` | valid archive, wrong internal format | rejected / fails cleanly: no PDF, no crash, no retry loop | no |
| malformed_empty | `malformed_empty.docx` | empty source file (0 bytes) | rejected / fails cleanly: no PDF, no crash, no retry loop | no |
| stress_huge_table | `stress_huge_table.docx` | conversion time / timeout | finishes within convert timeout; note stalls or timeouts | no |
