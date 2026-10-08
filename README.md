# SheetPDF Platform — First Working Version

A no-build static website with browser-based PDF/image utilities and one minimal Cloudflare Pages Function backend endpoint. Designed for free hosting on Cloudflare Pages.

## Features included
- Responsive professional interface and red primary conversion buttons
- PDF text extraction to `.xlsx` (selectable PDF text; table layout may need cleanup)
- Merge PDFs
- Export a page range from a PDF
- JPG/PNG images to PDF
- Image compression to JPEG
- Percentage and percentage-change calculator
- Word/character/sentence counter
- Contact page using `padhiary1996@gmail.com`
- Privacy and terms starter pages
- Backend health endpoint at `/api/health`

## Architecture
- `index.html`, `styles.css`, `app.js`: static frontend hosted on Cloudflare Pages.
- `functions/api/health.js`: Cloudflare Pages Function at `/api/health`, returning service status. It does not accept or store files.
- PDF processing is done in the visitor's browser using PDF.js and pdf-lib from CDN; spreadsheets use SheetJS from CDN. The first version does not require a paid server or database.

## Deploy on Cloudflare Pages
1. Back up your current repository files.
2. Upload/replace `index.html`, `styles.css`, `app.js`, `contact.html`, `privacy.html`, `terms.html`, and the `functions/` folder in the root of your GitHub repository.
3. Commit to the connected branch and wait for Cloudflare Pages to deploy.
4. Test `https://YOUR-SITE.pages.dev/api/health` — expected JSON includes `"ok":true`.
5. Test each tool with small non-sensitive sample files, on desktop and mobile.

## Important before deployment
- Preserve and re-add your actual Google AdSense publisher script and ad placements if they exist in your current `index.html`. This package intentionally does not include a publisher ID.
- External libraries/fonts load from CDNs; test content blockers and browser compatibility.
- PDF-to-Excel only extracts selectable text in this first version. It does not yet perform OCR on scanned PDFs and does not yet embed separately extracted pictures into Excel. Do not advertise those features until implemented and tested.
- Privacy and terms pages are starter text, not legal advice. Update them to match the services actually enabled (including ads/analytics).
- No database, user accounts, file uploads to a server, or persistent file storage are included. This keeps the first version low-cost and minimizes handling of private files.
- Do not add a backend file upload endpoint without adding strict size limits, validation, rate limits, and a deletion policy.
