# SheetPDF — free PDF & Excel tools starter

A responsive static website with a working first-draft browser converter.

## Included
- PDF → Excel: extracts selectable text from a PDF into an XLSX file. This is a basic text extraction, not perfect table reconstruction.
- Excel/CSV → PDF: opens a print-friendly preview; use your browser's Print → Save as PDF.
- CSV → Excel.
- Searchable tool cards, responsive design, FAQ, privacy/terms/contact starter pages, and an AdSense placeholder.

## Run locally
Open `index.html` in a modern browser with internet access. External libraries are loaded from CDNs.

## Publish free using Cloudflare Pages
1. Create/sign in to GitHub.
2. Create a **public** repository named `pdftoolsfree`.
3. Upload the files from this folder (`index.html`, `styles.css`, `app.js`, `privacy.html`, `terms.html`, `contact.html`, `README.md`).
4. Sign in to Cloudflare Dashboard → Workers & Pages → Create → Pages → Connect to Git.
5. Select your repository. Use no build command and set the output directory to `/` (or leave blank if the dashboard allows it).
6. Deploy and test the `pages.dev` URL on desktop and mobile.

Alternative: use Cloudflare Pages Direct Upload if offered in your dashboard; upload the site files (not the zip).

## Before publishing
- Replace SheetPDF branding with your chosen name and check that the name is available.
- Replace placeholder contact information in `contact.html` and `privacy.html`.
- Review and adapt the legal templates for your actual services and jurisdiction.
- Test with your own sample files. PDF → Excel works best on text-based PDFs; scanned PDFs need OCR, which this demo does not include.
- The Excel → PDF tool uses the browser print dialog and simplified table styling. It doesn't preserve all workbook formatting or multiple worksheets.
- Add AdSense code only after applying and receiving approval. The ad area is only a placeholder and does not show ads.
- The page loads third-party scripts (PDF.js, SheetJS, Google Fonts) from CDNs. Review privacy/security and dependency choices before handling sensitive files.
- This site does not include a server-side upload, OCR service, contact form backend, analytics, or payment processing.
