# MiS Stub Tool
*Making it easy to find the stubs that are relevant to you.*

Pulls only **our** payment stubs out of the big department stub PDFs, by matching each stub's
disbursement number (and payment number) against the Payments Database. Each payment is saved as its own PDF named
`PAYEE - YYYY-MM-DD - DISBURSEMENT NO.pdf` (date = the stub's action/paid date), optionally in a folder per payee.

## Privacy
Everything runs inside your browser. No uploads, no server, no network calls (enforced by a Content-Security-Policy in `index.html`).
The repository contains only the program, never any PDFs or spreadsheets.

## Use
Open `index.html` (double-click, or via GitHub Pages), then:
1. Drop the PAYMENTS DATABASE `.xlsx`
2. Drop the STUBS `.pdf` files
3. **Find our payments**, tick/search payees (e.g. `MATHOLE`), then **Download as ZIP** (or save straight into a folder in Chrome/Edge)

## Notes
- Continuation pages (payments with many lines that run onto extra pages) are kept with their payment.
- Pages with no readable text layer (e.g. garbled print-to-PDF) are listed under *Checks*; they cannot be matched.
- `lib/` holds pinned copies of pdf.js, pdf-lib, SheetJS and JSZip so it works offline.
