# SheetPDF OCR upgrade

This static website uses PDF.js for PDF text and page rendering, Tesseract.js for English OCR, and SheetJS to create Excel workbooks.

## Features
- Extracts selectable PDF text and groups fragments into approximate rows/cells.
- OCR fallback for pages with little/no selectable text.
- Option to force OCR on every page.
- One worksheet per PDF page.

## Limitations
- Table extraction is heuristic, not perfect reconstruction; review and correct the workbook.
- OCR is configured for English printed text; handwriting and poor scans may fail.
- It extracts recognized text; it does not embed original PDF pictures as images in Excel.
- Large PDFs may be slow or use substantial browser memory.
- Requires internet to load CDN libraries and OCR language data.
- Replace the contact email placeholder and review legal pages before launch.

## Upload to existing GitHub repository
1. Download and extract this ZIP.
2. Open https://github.com/padhiary1996-png/pdftoolsfree
3. Choose **Add file → Upload files**.
4. Drag the files themselves (index.html, styles.css, app.js, privacy.html, terms.html, contact.html, README.md) into the upload area, not the enclosing folder.
5. Scroll down, enter commit message `Upgrade PDF to Excel with OCR`, and click **Commit changes**.
6. Cloudflare Pages should automatically deploy from the connected GitHub repo. Wait for a successful deployment, then refresh https://pdftoolsfree.pages.dev with Ctrl+F5.
