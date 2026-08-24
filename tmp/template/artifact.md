# Artifact contract: trademark-search report

## Reference

- Retained DOCX: `/Users/poma/Desktop/Прога Шаблон ТЗ/assets/report_template.docx`
- SHA-256: `618359bf3f9b5e7485e7acdb2a19881467eaa2f3701d4471bf1598c578efe69a`
- Reference render: `tmp/template/reference-render/`
- Pages in Word render: 5; sections: 1.
- The retained reference is read-only. Generated reports use a byte copy and never overwrite it.

## Page system

- A4 portrait: 11906 x 16838 DXA.
- Margins: top 851 DXA, right 1134, bottom 1366, left 1134.
- Header distance 47 DXA; footer distance 709 DXA.
- One section; same header and page-number footer on all pages.
- Page breaks are deliberate between summary, conclusion, appendix, and fees. The internal probability guidance page is excluded from client output.

## Typography and palette

- Primary typeface: Calibri, 11 pt body, black.
- Report title: Calibri 11 pt bold, centered, two lines.
- Header contact block: embedded source image, left; logo image, right; thin gray rule between.
- Metadata/date: 11 pt gray, right aligned.
- Blue table label/header fill: `DBE5F1`.
- Green conclusion/probability fill: source light green (`C4D79B` family).
- Standard borders: black single 4 eighth-points.
- Body paragraphs in conclusion: justified; selective bold/italic emphasis follows the source examples.
- Hyperlinks: blue, single underline.

## Tables and components

- Page 1 key-value tables: 9748 DXA wide; columns 3043 / 6705 DXA; left cells blue, bold, vertically centered.
- Conclusion heading: one-cell green band, full content width, centered bold.
- Probability table: two columns; label cell green; value cell white.
- Appendix cards: three columns. Left image column; middle labels; right values. Section header row spans all columns with blue fill.
- Fees table: four columns, blue header, centered values, variable-height rows, bold total.
- Table rows must expand; no exact-height clipping. Cell padding remains visually comparable to source.

## Content flow and slot map

1. Summary page: report date, designation, search queries, optional business area, repeatable Nice classes, absolute grounds, relative grounds, two database dates.
2. Conclusion page(s): selected conclusion template, optional expert additions, fixed disclaimer, repeatable probability rows, selected performer signature block.
3. Appendix: zero or more international marks, Russian marks, and applications. Each record has its own image and fields; registration/application number is a hyperlink to its FIPS card.
4. Fees: class count, excess goods/services count, calculated filing/examination fee, registration fee, total, fixed notes and social links.

Editable runtime slots are all body content described above. Source header/footer, page geometry, embedded brand images, style/theme/font parts, and social URLs are preserve-derived. Source comments, comment balloons, internal probability guidance, placeholders, yellow drafting highlights, and unused performer blocks must not appear in generated output.

## Package preservation

- Preserve source theme, embedded fonts, header images, header/footer relationships, core styles, numbering and page geometry where python-docx does not need to update them.
- Body XML is intentionally rebuilt because the application requires conditional and repeatable content.
- Comments-related parts may be removed from generated client reports.
- New external hyperlink relationships and new mark-image parts are expected body changes.

## Fidelity gates

- Render every generated page through Microsoft Word to PDF and PNG because LibreOffice is unavailable in this environment.
- Inspect all pages at 100%: no clipping, overlap, broken borders, missing Cyrillic glyphs, detached headings, or blank template placeholders.
- Compare visually against both completed sample PDFs, not against the source's comment balloons.
- Confirm all linked numbers have external hyperlink relationships and all supplied mark images are embedded.
