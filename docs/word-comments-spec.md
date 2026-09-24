# Word Comment Review: Specification and Implementation Guide

This document describes a working prototype that lets a review team comment on a
customer's Word document in a browser and hand back that same `.docx` with real
Word margin comments, replies and resolved threads. It is written so another
engineer or coding agent can re-implement the feature in a different project.

- Working reference implementation: [`review.html`](../review.html) (one file, no build step)
- End-to-end tests: [`tests/review/`](../tests/review/) (run `tests/review/run.sh`)

---

## 1. Problem and goals

The team reviews documents it receives from customers. Usually there is:

1. **A source document**: the team's original, most often a PDF, sometimes Word.
2. **A customer document**: a Word `.docx` the customer's legal team sent. It may
   already contain the customer's comments.

The customer's legal team works only in Microsoft Word. Feedback must reach them as
**native Word comments** in their own file, with each comment showing **who wrote it**
so they know whom to answer. Posting as a team name (for example "Review Team") must
also be possible.

### Goals

| # | Goal | Status |
|---|------|--------|
| G1 | Show source and customer documents side by side | Done |
| G2 | Select any text in the customer document and attach a comment | Done |
| G3 | Every comment carries author name, initials and date | Done |
| G4 | Show comments already in the file, including reply threads and resolved state | Done |
| G5 | Reply to any comment (theirs or ours) | Done |
| G6 | Resolve / reopen any thread | Done |
| G7 | Edit / delete our own comments and replies before saving | Done |
| G8 | Save the customer's own `.docx` with everything written as native Word data, leaving the rest of the file untouched | Done |
| G9 | Work survives a page reload | Done (browser storage) |
| G10 | No server; documents never leave the machine | Done |

### Out of scope for the prototype (see section 12, Next phases)

- Multi-user, real-time collaboration and user accounts
- Importing comments and classifications (such as "internal control issue" vs "finding")
  from the team's existing report tool
- Comments in headers, footers, footnotes, text boxes
- Tracked changes (insertions/deletions are shown as final text, see 9.2)

---

## 2. User workflow

1. Enter **Your name** in the header. It is the default author for new comments and replies.
2. Load the **source document** (PDF or `.docx`) in the left panel. It is for reference only
   and is never modified.
3. Load the **customer document** (`.docx`) in the middle panel. This is the file that receives
   all comments. Existing comments appear in the right panel as threads.
4. Select text in the customer document. A floating **Comment** button appears. Click it,
   choose who the comment is from (free text with suggestions: your name, "Review Team",
   and every author already seen), write the comment, and click **Add comment** (or Ctrl/Cmd+Enter).
5. On any thread: **Reply**, **Resolve** / **Reopen**. On your own comments/replies: **Edit**, **Delete**.
6. Click **Save Word document**. The browser downloads `<original name> - reviewed.docx`.
7. Send that file to the customer. In Word they see every comment in the margin on the exact
   text, with author names, reply threads and resolved state. When they send a new version
   back, load it as the customer document and continue.

The customer never uses this tool; they use Word.

---

## 3. Architecture

```
┌──────────────────────── Browser (single page) ───────────────────────────┐
│                                                                          │
│  Source panel         Customer document panel        Comments panel      │
│  PDF.js → <canvas>    DOCX → HTML render             threads, replies,   │
│  or DOCX → HTML       (text selectable, highlights)  composer, actions   │
│                                                                          │
│  ── state ───────────────────────────────────────────────────────────    │
│  doc: parsed customer doc (texts[], paraOf[], existing comments)        │
│  comments[]: our new comments + replies (anchored by text index)        │
│  resolutions{}: thread id → resolved (only where changed)               │
│                                                                          │
│  ── storage: IndexedDB  (files + comments + settings) ─────────────     │
│                                                                          │
│  ── save: JSZip opens ORIGINAL bytes → edit XML parts → download ──     │
└──────────────────────────────────────────────────────────────────────────┘
```

Libraries (CDN, cdnjs):

| Library | Version | Used for |
|---------|---------|----------|
| JSZip | 3.10.1 | Read/write the `.docx` zip package |
| PDF.js | 3.11.174 | Render the source PDF pages to canvas |

Everything else is browser built-ins: `DOMParser` / `XMLSerializer` for XML, `Range` /
`Selection` for text selection, IndexedDB for storage.

**Key design decision:** the tool never regenerates the Word document. It always re-opens
the original bytes and makes surgical XML insertions. That is what keeps the customer's
formatting, numbering, styles, tracked changes, and other content intact. Libraries that
*create* documents (for example the npm `docx` package) cannot do this and were rejected.

---

## 4. Word file format background (what an implementer must know)

A `.docx` is a zip of XML "parts". The parts involved:

| Part | Role |
|------|------|
| `[Content_Types].xml` | Declares the content type of each part |
| `_rels/.rels` | Points to the main document part (usually `word/document.xml`) |
| `word/document.xml` | Body text; contains comment **anchors** |
| `word/_rels/document.xml.rels` | Relationships from the main part to other parts |
| `word/comments.xml` | Comment **bodies**: author, initials, date, text |
| `word/commentsExtended.xml` | (Word 2013+) **Reply threading and resolved** state |
| `word/commentsIds.xml`, `word/commentsExtensible.xml` | (Word 2016+) durable IDs; optional, not required |

Namespaces:

```
w   = http://schemas.openxmlformats.org/wordprocessingml/2006/main
w14 = http://schemas.microsoft.com/office/word/2010/wordml
w15 = http://schemas.microsoft.com/office/word/2012/wordml
mc  = http://schemas.openxmlformats.org/markup-compatibility/2006
```

### 4.1 A comment anchor in `document.xml`

```xml
<w:p>
  <w:r><w:t xml:space="preserve">The Supplier shall </w:t></w:r>
  <w:commentRangeStart w:id="1"/>
  <w:r><w:t xml:space="preserve">deliver the </w:t></w:r>
  <w:r><w:rPr><w:b/></w:rPr><w:t>Services</w:t></w:r>
  <w:commentRangeEnd w:id="1"/>
  <w:r><w:commentReference w:id="1"/></w:r>
  ...
</w:p>
```

- `commentRangeStart` / `commentRangeEnd` mark the highlighted range. They sit **between runs**
  (`w:r`), never inside one, and may be in different paragraphs.
- A run containing `commentReference` must follow the range end. That is where the
  comment "lives" in the text flow.

### 4.2 A comment body in `comments.xml`

```xml
<w:comment w:id="1" w:author="Jane Smith" w:initials="JS" w:date="2026-09-24T23:42:23Z">
  <w:p w14:paraId="211A2B4A">
    <w:r><w:annotationRef/></w:r>
    <w:r><w:t xml:space="preserve">Please define "Services" precisely.</w:t></w:r>
  </w:p>
</w:comment>
```

- `w:author` and `w:initials` are what Word shows as the commenter. This is the attribution.
- Multi-line comments = one `w:p` per line; `annotationRef` goes in the first paragraph.
- `w14:paraId` on the **last** paragraph is the key used by `commentsExtended.xml`.

### 4.3 Threads and resolved state in `commentsExtended.xml`

```xml
<w15:commentsEx xmlns:w15="..." xmlns:mc="..." mc:Ignorable="w15">
  <w15:commentEx w15:paraId="1A2B3C03" w15:done="1"/>                               <!-- thread root, resolved -->
  <w15:commentEx w15:paraId="1A2B3C04" w15:paraIdParent="1A2B3C03" w15:done="1"/>   <!-- reply -->
</w15:commentsEx>
```

- A reply is an ordinary `w:comment` whose `commentEx` has `paraIdParent` = the root's paraId.
- A reply's anchor in `document.xml` uses **the same range as its parent** (its own
  `commentRangeStart/End` ids placed right next to the parent's, plus its own reference run).
- Resolved = `w15:done="1"`. Word sets it on the root; the prototype sets it on the root **and**
  every reply in the thread (and clears all of them on reopen). Readers should treat a thread
  as resolved if the root or any reply is done.
- The part must be registered in `document.xml.rels`
  (Type `http://schemas.microsoft.com/office/2011/relationships/commentsExtended`) and in
  `[Content_Types].xml`
  (`application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml`).

---

## 5. Reading the customer document

Implemented in `openDocx`, `walkDocument`, `readExistingComments`.

### 5.1 Locate parts
1. Load zip. Find the main part via `_rels/.rels` (type `.../officeDocument`); default `word/document.xml`.
2. From the main part's rels, find the comments part (type `.../relationships/comments`)
   and the commentsExtended part (type `.../2011/relationships/commentsExtended`). Resolve
   relative targets against the main part's folder.

### 5.2 Walk the body and index every text node (the core idea)
Walk `w:body` in reading order and give every visible `w:t` element a sequential integer
index **`ti`**. Record for each `ti`: its text (`texts[ti]`), its paragraph number (`paraOf[ti]`).

Walk rules:
- Blocks: `w:p` (paragraph), `w:tbl` → `w:tr` → `w:tc` → blocks (recursive), `w:sdt` → `w:sdtContent`, `w:customXml`.
- Inside a paragraph, descend through: `w:r`, `w:hyperlink`, `w:ins`, `w:moveTo`, `w:smartTag`,
  `w:fldSimple`, `w:customXml`, `w:dir`, `w:bdo`, `w:sdt/w:sdtContent`.
- Skip entirely: `w:del`, `w:moveFrom` (deleted text uses `w:delText`, which is never indexed),
  drawings, text boxes (`mc:AlternateContent` would otherwise duplicate text), field
  instructions (`w:instrText`).
- In a run: `w:t` → indexed text; `w:tab` → tab; `w:br`/`w:cr` → line break. Record bold/italic/underline for display.
- Track existing comment ranges while walking: on `commentRangeStart id` add to an open set,
  on `commentRangeEnd id` remove. Each indexed `w:t` extends the anchor of every open id.

**Invariant:** the same walk is re-run at save time on a fresh parse of the original bytes,
so `ti` N always refers to the same `w:t` node. The prototype asserts the count matches
before saving.

### 5.3 Existing comments
For each `w:comment` in comments.xml: id, author, initials, date, text (paragraphs joined by `\n`),
paraId of last paragraph, anchor (from 5.2, whole-`w:t` granularity) and quoted text.
Join with commentsExtended by paraId → `parentParaId`, `done`. Group replies under their root
(follow parent links to the root); a thread is resolved if the root or any reply is `done`.

---

## 6. Rendering and selection

Implemented in `renderBlocks`, `paintHighlights`, `anchorFromSelection`.

- Each indexed `w:t` renders as `<span data-ti="N">text</span>` inside `<p>` (headings
  styled from `pStyle` = Heading1..3/Title; list paragraphs get a bullet; tables become `<table>`).
  All document text is inserted with `textContent` (never `innerHTML`) to avoid script injection.
- **Anchor model** for our comments:
  `{ start: { ti, off }, end: { ti, off } }`, where `off` is a character offset inside `texts[ti]`.
  Start inclusive, end exclusive.
- **Selection → anchor:** take the DOM `Range`; the first/last `span[data-ti]` that
  `range.intersectsNode(span)` gives start/end `ti`. Offsets inside a span are computed by
  measuring `Range(span start → selection point).toString().length`, which works even after
  highlights split the span into `<mark>` pieces. Normalize (a start at the end of a span moves to
  the next span's 0; an end at 0 moves back). Empty selections are rejected.
- **Highlights:** for each span, compute the intervals of every visible thread (plus the pending
  selection), cut the span's text at all interval boundaries, and wrap covered pieces in
  `<mark>` with classes: ours (yellow), existing (blue), resolved (dashed green underline),
  overlap (stronger), active (outline). Clicking a mark selects its thread card; clicking a card
  scrolls the document to its anchor.

---

## 7. Data model and persistence

```ts
// New comment or reply created in the tool
type NewComment = {
  id: string;                 // "c" + random
  parentId?: string;          // set for replies: id of thread root ("ext-<w:id>" or a new comment id)
  author: string; initials: string; text: string;
  created: string;            // ISO date
  start?: { ti: number; off: number }; end?: { ti: number; off: number }; // top-level only
  quote?: string;
};
// Existing comment read from the file
type ExistingComment = {
  id: `ext-${string}`; ext: true;
  author; initials; created; text; paraId; parentParaId; done;
  parentId?: string; replies: ExistingComment[]; resolved?: boolean;
  start?; end?; quote;
};
resolutions: Record<threadId, boolean>; // only entries that differ from the file
```

IndexedDB database `word-review`, object store `kv`:

| Key | Value |
|-----|-------|
| `doc` | `{ name, buffer }`, the customer file bytes |
| `source` | `{ name, buffer }`, the source file bytes |
| `comments:<name>:<size>` | `{ comments: NewComment[], resolutions }` |
| `settings` | `{ name }`, the reviewer's name |

Comments are keyed by file name + size, so re-loading the same file restores its work and a
different file starts clean. All storage calls are wrapped in try/catch; the tool works
without storage (private windows) but forgets on reload.

**Porting note:** in a multi-user product this state moves to the server (see section 12). The
anchor model `(ti, off)` is only valid for one exact version of the file; store the file
hash with the comments and refuse to apply anchors to a different version.

---

## 8. Saving: the write algorithm

Implemented in `exportDocx`, `boundaryAt`, `addPart`. Steps, in order:

1. **Re-open the original bytes** with JSZip and re-run the walk (5.2). Abort if the `w:t` count differs.
2. **Comments part:** parse `comments.xml`, or create an empty `<w:comments>` and register it
   (relationship type `.../relationships/comments`, content type
   `application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml`).
3. **Next comment id** = 1 + max of every `w:id` in `w:comment`, `commentRangeStart`,
   `commentRangeEnd`, `commentReference`.
4. **Decide if thread data is needed:** yes if the file already has `commentsExtended.xml`,
   or we have any reply, or any resolve/reopen. If needed and missing, create it and register it.
5. **paraIds:** collect every existing `w14:paraId` / `w15:paraId` in the document, comments and
   commentsExtended parts. When thread data is in play, make sure **every** comment (existing and new)
   has a unique paraId on its last paragraph (8 hex digits, value < 0x80000000) and a `commentEx`
   entry. Declare `xmlns:w14` on the comments root.
6. **Create top-level comments** (sorted by position): build `w:comment` (author, initials, ISO date
   without milliseconds, one `w:p` per line, `annotationRef` in the first) and queue two
   anchor points: start `(ti, off)` and end `(ti, off)`.
7. **Create replies** (sorted by date): build the `w:comment`, set its `commentEx.paraIdParent` to the
   parent's paraId, then anchor:
   - Parent is new → queue the same start/end points as the parent (queued after the parent, so
     markers land right after the parent's).
   - Parent is existing → insert `commentRangeStart` directly after the parent's
     `commentRangeStart`, and `commentRangeEnd` + reference run directly after the parent's
     reference run (chaining multiple replies in order). If the parent has no range, only the reference run.
8. **Resolved state:** for each changed thread, set `w15:done` = 1/0 on the root and all replies.
9. **Insert anchor points, from the end of the document backwards** (group points at the same
   `(ti, off)`; process groups in descending order so earlier positions stay valid):
   - `boundaryAt(w:t, off)` returns an insertion point *between runs*:
     - `off == 0` and the `w:t` is the run's first content → before the run.
     - `off == len` and the `w:t` is the run's last content → after the run.
     - Otherwise **split the run**: clone the run (keeping attributes) with a copy of its `w:rPr`,
       move the text after `off` and all following siblings into the clone, truncate the original.
       Both halves keep identical formatting, so nothing visibly changes.
   - At each point insert, in order: for every range ending here `commentRangeEnd` + `<w:r><w:commentReference/></w:r>`;
     then for every range starting here `commentRangeStart`.
   - Because the original `w:t` element always keeps the *first* part of its text, processing in
     descending order means a later (smaller) offset in the same `w:t` still indexes correctly.
10. **Serialize** each modified part with `XMLSerializer` (prepend the XML declaration if missing),
    write back into the zip, generate a blob with the Word MIME type, and download.

Nothing else in the package is touched: styles, numbering, headers, tracked changes, images,
custom XML and existing comments pass through byte-for-byte (apart from paraIds added to
comment paragraphs that lacked them).

---

## 9. Edge cases and behaviors

### 9.1 Handled
- Selection crossing formatting runs, hyperlinks, paragraphs and table cells.
- Selection starting/ending exactly on run boundaries (no needless splits).
- Several comments starting/ending at the same character.
- Overlapping comments, including overlap with the customer's comments.
- Files with no comments part, with comments but no thread data (for example files from
  python-docx or older Word), and with full Word 2013+ thread data.
- Comments on text that has existing paraIds; new paraIds never collide.
- Existing comment ids are never reused.
- Multi-line comment text.
- Author suggestions; generic "Review Team" author; initials derived from the name.

### 9.2 Known limitations
- Text boxes, headers, footers, footnotes/endnotes are not shown and cannot be commented.
- Tracked deletions are hidden and insertions shown as normal text; comments can still be
  placed, but the viewer does not visualize tracked changes.
- Rendering is simplified (no page layout, list numbers, fonts, or images); it is a reading
  view, not a Word replica.
- Existing customer comments are read-only (reply/resolve only, no edit/delete), by design.
- Anchors are tied to the exact file bytes; editing the document text in the tool is not supported.
- No `commentsIds.xml` / `commentsExtensible.xml` entries are written for new comments;
  Word tolerates this.
- Single user per browser; no sync.

---

## 10. Verification

`tests/review/run.sh` builds fixtures and runs everything. It requires Node with Playwright
(Chromium), Python 3 with `python-docx`, and LibreOffice Writer (`soffice`).

| Fixture | Built by | Contents |
|---------|----------|----------|
| `customer.docx` | `make_fixture.py` | Heading, mixed bold/italic runs, table, one existing comment (no thread data) |
| `source.pdf` | LibreOffice from the above | Source-panel PDF |
| `threaded.docx` | `make_threads.py` | Adds a resolved thread with a reply and an open thread with a reply (Word 2013+ thread data) |

| Test | Checks |
|------|--------|
| `test.js` | Existing comment shown; 5 new comments (cross-run, cross-paragraph, table, overlap); edit; delete; reload persistence; save; saved file re-opened shows all comments on the right text; phone-width layout has no horizontal scroll; no console errors |
| `test2.js` | Replies nested, resolved label, "Show resolved" toggle, save keeps thread data |
| `test3.js <in> <out>` | Reply to customer thread, resolve, reopen, new comment + reply + edit reply + resolve, reload persistence, save, re-open |
| `lo_check.py <file>` | Independent read-back through LibreOffice: every comment's author, reply parent and resolved state |

Lessons from verification:
- Parse the LibreOffice output with `<office:annotation(?=[\s>])`; a looser pattern also matches
  `<office:annotation-end>` and produces false "unlinked reply" results.
- Always verify with a reader other than your own tool (python-docx, LibreOffice, and finally
  real Microsoft Word) because your own reader shares your assumptions.
- **Not yet done:** open saved files in real Microsoft Word (desktop and Word Online). This is the
  final acceptance check before production.

---

## 11. Porting to another project: checklist

1. **Keep the core modules separable.** The logic splits cleanly into:
   - `docx-read`: `openDocx`, `walkDocument`, `readExistingComments` (pure, no DOM rendering)
   - `docx-write`: `exportDocx` core (`boundaryAt`, marker insertion, comments/commentsExtended writing, `addPart`)
   - `viewer`: `renderBlocks`, `paintHighlights`, `anchorFromSelection`
   - `ui/state`: comment list, composer, reply/resolve actions, persistence
   The read/write modules work in Node too (swap `DOMParser` for `@xmldom/xmldom`), which allows
   server-side generation.
2. **Reuse the anchor model** `{start:{ti,off}, end:{ti,off}}` and the walk rules exactly. Reader and
   writer must share one walk function.
3. **Store the file hash** (SHA-256 of the bytes) with every set of comments.
4. **Always write into the original bytes.** Never regenerate the document.
5. **Map authors to real users** in a multi-user app: `w:author` = display name, `w:initials` =
   initials; keep a "team" pseudo-author option.
6. **Port the tests**: the fixtures and the LibreOffice read-back transfer as-is; add a manual Word check.
7. **Security**: render document text with `textContent` only; treat file content as untrusted;
   limit upload size; if server-side, process files in memory and apply retention rules suitable
   for legal documents.

---

## 12. Next phases

**Phase 2: import comments from the existing report tool.** The team's current report records
comments with a classification (for example "internal control issue" or "finding"). Plan:
- Define an import format per comment: `{ quote or anchor, text, author, classification, date }`.
- Locate each comment's position in the customer document: exact text match on the quoted
  passage using the `texts[]` index (with whitespace normalization), falling back to a manual
  "place this comment" step when not found or ambiguous.
- Carry the classification into Word, most simply as a prefix in the comment text
  (`[Finding] ...`) so it is visible in Word; optionally keep it as structured data in the app.
- Allow filtering by classification in the comments panel.
- Requires access to the report tool's data (API, export file or database); not yet examined.

**Phase 3: multi-user and real time.** Server-side storage of files and comment state per
review, user identity from login, live updates (WebSocket or similar), locking or merge rules,
and an audit trail. Word generation can move server-side using the same write algorithm.

---

## 13. Acceptance criteria (for the re-implementation)

- [ ] Loading a Word file with existing comments, replies and resolved threads shows all of them
      correctly grouped and labeled.
- [ ] A comment can be placed on any text selection in the body, including across formatting
      and paragraphs, and is highlighted on exactly that text.
- [ ] The saved file opens in Microsoft Word without a repair prompt.
- [ ] In Word, every new comment appears on exactly the selected text with the chosen author
      name and initials, and replies appear in their threads.
- [ ] Resolved/reopened state set in the tool matches Word's display.
- [ ] All pre-existing comments, replies, resolved flags, formatting, numbering and tracked
      changes are unchanged.
- [ ] Reloading the page (or re-opening the review) restores in-progress work.
- [ ] Automated tests equivalent to `tests/review` pass.
