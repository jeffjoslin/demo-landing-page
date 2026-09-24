# demo-landing-page

A responsive contact form landing page with form validation and modal confirmation.

## Features

- Responsive design (mobile-friendly)
- Form validation for all fields
- Email format validation
- Modal popup on successful submission
- Clean, modern UI

## How to View

**Option 1: Open directly**
```
Open index.html in your browser
```

**Option 2: Local server**
```bash
python3 -m http.server 8000
```
Then visit http://localhost:8000

## Document Review (`review.html`)

Review a customer's Word document against your source document and save your
comments into their file as real Word margin comments.

1. Open `review.html` in a browser (needs internet access to load JSZip and PDF.js).
2. Enter your name at the top.
3. Load your **source document** (PDF or .docx) on the left. It is for reference only.
4. Load the **customer document** (.docx) in the middle. This is the file comments are saved into.
   Comments already in the file are listed as "In document", with replies shown under
   the comment they answer. Resolved threads are marked "Resolved" and can be hidden
   with the "Show resolved" toggle.
5. Select text in the customer document, click **Comment**, choose who the comment is from
   (your name, a colleague, or "Review Team"), and add it. Comments can be edited or deleted.
6. Click **Save Word document** to download `<name> - reviewed.docx`. Each comment appears in
   Word's margin on the exact text selected, with the author's name, initials and date.
   The rest of the document is left unchanged.

Documents and comments are kept in the browser (IndexedDB), so reloading the page
restores your work. Nothing is uploaded to a server.
