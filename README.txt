SHEETPDF EDITABLE PDF-TO-EXCEL UPGRADE

Files included:
- index.html
- app.js

WHAT THIS VERSION DOES
- Extracts selectable PDF text into editable Excel cells.
- Uses OCR (English) for scanned pages when enabled.
- Adds a separate worksheet for each PDF page with a visual page image, so the original layout remains viewable.
- Attempts to extract embedded raster images from the PDF and place them on separate Excel worksheets as selectable/resizable pictures.

IMPORTANT LIMITATIONS
- PDF files do not store all pictures in a simple standard way. The image extractor is best-effort and cannot guarantee every logo, chart, vector illustration, or background will be extracted individually.
- A full-page visual copy is also included, so the original appearance remains available even when individual image extraction is not possible.
- PDF tables are inferred from text positions; complex tables may need manual cleanup.
- OCR is English-only in this build and accuracy depends on scan quality.
- Large PDFs and high-resolution images can produce large Excel files and take longer.
- This is a browser-side tool. PDF data is processed locally, but third-party libraries are loaded from CDNs.

INSTALL
1. Back up your current index.html and app.js from the GitHub repository.
2. Upload/replace only index.html and app.js with these files.
3. Keep your existing styles.css, privacy.html, terms.html, and contact.html.
4. IMPORTANT: The index.html in this package intentionally does NOT include an AdSense script. Keep your existing verified AdSense script/publisher ID and ad units; add them back if replacing the whole index.html. Do not use an example publisher ID.
5. Commit to the main branch and wait for Cloudflare Pages deployment.
6. Hard refresh the website (Ctrl+F5) and test with a small PDF containing selectable text, a scan, and pictures.

No converter can guarantee perfect reconstruction of every PDF into editable spreadsheet cells. Review the workbook before using it.
