# Universal File Converter

A Homeroom app that converts files in the browser:

- **PDF to** Word (DOCX), Excel (XLSX), PowerPoint (PPTX), JPG or PNG
  (one image per page, zipped when there are several pages).
- **Word, Excel, PowerPoint, JPG, PNG or TIFF to** PDF.
- **Scans and photos to** Word or Excel: JPG, PNG and TIFF images, and the
  pages of a PDF that have no text of their own, are read with OCR
  (Tesseract.js) in English, Spanish, French, German, Portuguese or Italian.
  In Excel, words that line up in columns go into separate cells.

Drop a file on the page (or choose one), see a preview, pick a format,
convert, preview the result and download it. "Try a sample PDF" makes a
small two-page PDF to try it with, and "a sample scan" a made-up picture of
a table to try OCR with.

## How it works

Every conversion runs in the visitor's browser, so files never leave their
device and the server stores nothing. `server.js` serves the page and the
browser builds of the conversion libraries from `node_modules` under
`/vendor/…`; the page loads each library the first time it is needed.

| Path | What |
|---|---|
| `public/index.html` | The screen |
| `public/app.js` | Picking, previews, progress, download |
| `public/convert.js` | The conversions |

Libraries: pdf.js (reading and rendering PDFs), docx (writing Word),
ExcelJS (reading and writing Excel), PptxGenJS (writing PowerPoint),
mammoth (reading Word), pdfmake (writing PDFs), JSZip (ZIPs and reading
PowerPoint), Tesseract.js (OCR; its reading data per language is served
from `/vendor/tessdata/<code>`), UTIF with pako (reading TIFFs).

## Limits

- Files up to 25 MB; PDFs up to 200 pages.
- PDF to Word and Excel use the PDF's text layer where a page has one, and
  OCR where it doesn't. Up to 50 scanned pages are read per conversion.
  Handwriting is not supported; a page where no text could be read stays a
  picture in Word and an empty sheet in Excel.
- PowerPoint to PDF draws text, pictures, simple shapes and tables. Charts,
  SmartArt, theme colours and effects are left out, and slides become
  pictures in the PDF.
- PDFs written here use the Roboto font, which has no CJK characters.
- Older `.doc`, `.xls` and `.ppt` files are not supported.
