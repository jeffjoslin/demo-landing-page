#!/usr/bin/env bash
# End-to-end checks for review.html. Needs: node + playwright (Chromium), python3 + python-docx,
# and LibreOffice Writer (soffice) for the independent read-back of saved files.
set -euo pipefail
cd "$(dirname "$0")"
export NODE_PATH="${NODE_PATH:-$(npm root -g)}"

mkdir -p cdn
for u in jszip/3.10.1/jszip.min.js pdf.js/3.11.174/pdf.min.js pdf.js/3.11.174/pdf.worker.min.js; do
  f="cdn/$(basename "$u")"
  [ -f "$f" ] || curl -sSf -o "$f" "https://cdnjs.cloudflare.com/ajax/libs/$u"
done

python3 make_fixture.py
soffice --headless --convert-to pdf customer.docx --outdir . >/dev/null 2>&1
mv customer.pdf source.pdf
python3 make_threads.py

echo "== basic comments";            node test.js
echo "== existing threads";          node test2.js
echo "== reply/resolve (threaded)";  node test3.js threaded.docx t3-threaded.docx
python3 lo_check.py "$PWD/t3-threaded.docx"
echo "== reply/resolve (plain)";     node test3.js customer.docx t3-plain.docx
python3 lo_check.py "$PWD/t3-plain.docx"
