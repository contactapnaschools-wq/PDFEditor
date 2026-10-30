# Roadmap

Step 11 is done. The tag is `<pdf-editor>`. The next step is undo, print, and JSON import/export. Later steps stay in this file so the next session continues from here.

## Done

1. **Shell and viewing** — toolbar, page canvas, comments drawer, open a PDF, zoom, fit, pan, previous and next page.
2. **Thumbnails and search** — left page panel, and in-document find with hit highlights, including next and previous.
3. **Annotate** — highlight, underline, strikethrough, and sticky note. Marks stay in memory and in the comments list. They are not written into the PDF yet.
4. **More ink** — free text, pen, and callout. Same in-memory marks and comments list.
5. **Shapes** — line, arrow, rectangle, ellipse, and polygon. Same in-memory marks and comments list.
6. **Insert** — text, image, link, and stamp. Same in-memory marks and comments list. Nothing is written into the PDF yet.
7. **Measure** — distance, perimeter, and area. Lengths are in inches from PDF points. Same in-memory marks and comments list.
8. **Edit** — change a horizontal line of existing text. The new words are written into the working PDF with pdf-lib. Marks stay in memory. Nothing is downloaded yet.
9. **Fill and Sign** — signature, initials, and date. Same in-memory marks and comments list. Nothing is written into the PDF yet.
10. **Forms** — text fields and checkboxes. Read them, fill them, and add them. Values are written into the working PDF with pdf-lib. Nothing is downloaded yet.
11. **Save** — download a copy of the working PDF. The file includes text edits, form values, and the in-memory marks drawn on the pages. The open document stays as it is.

## Next

12. **Undo, print, and JSON** — undo and redo, print, and JSON import/export of annotations.

13. **Host apps** — small Angular and React examples, and CSS variables for theming.

The Annotate tab shows highlight, underline, strikethrough, note, free text, pen, and callout. The Shapes tab shows line, arrow, rectangle, ellipse, and polygon. The Insert tab shows text, image, link, and stamp. The Measure tab shows distance, perimeter, and area. The Edit tab shows edit text. The Fill and Sign tab shows signature, initials, and date. The Forms tab shows text field and checkbox. Save downloads the working PDF when a document is open.
