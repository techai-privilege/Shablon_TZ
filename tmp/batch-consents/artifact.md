# Template execution contract

- Reference: `/Users/poma/Desktop/Прога Шаблон ТЗ/assets/software_consent_template.docx`
- SHA-256: `8083d8a8d89721a9e87db7c53c250a5befd4aedd483a76adfdda50d191eced49`
- Structure: one A4 portrait section, 26 body paragraphs, one table; two consent pages per author.
- Page geometry: 21 × 29.7 cm; margins approximately 1.0 cm top/bottom, 0.6 cm right, 1.0 cm left.
- Editable slots: program/application details, applicant details, author identity, passport, contribution, author-mention choice, signatures and date.
- Preserve: all typography, paragraph/table geometry, borders, wording, page furniture and the two-page layout.
- Batch change: clone and fill the retained template once per author; never combine authors in one DOCX.
- Output names: `Согласие <ФИО автора>.docx`; avoid overwriting an existing file by adding ` (2)`, ` (3)`, etc.
- Fidelity gate: every output contains exactly one author and renders as two clean pages in Microsoft Word.
