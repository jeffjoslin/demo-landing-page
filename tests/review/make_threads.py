import zipfile, re
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
src = zipfile.ZipFile('customer.docx')
files = {n: src.read(n) for n in src.namelist()}
doc = files['word/document.xml'].decode()
com = files['word/comments.xml'].decode()

def cmt(i, author, ini, date, text, pid):
    return (f'<w:comment w:id="{i}" w:author="{author}" w:initials="{ini}" w:date="{date}">'
            f'<w:p w14:paraId="{pid}"><w:r><w:annotationRef/></w:r><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>')

# tag comment 0's paragraph and add replies / a second thread
if "xmlns:w14" not in com: com = re.sub(r"(<w:comments\b)", rf"\1 xmlns:w14=\"{W14}\"", com, count=1)
com = re.sub(r'(<w:comment [^>]*w:id="0"[^>]*>.*?)<w:p\b', r'\1<w:p w14:paraId="1A2B3C01"', com, count=1, flags=re.S)
com = com.replace('</w:comments>',
    cmt(10, 'Bob Lawyer', 'BL', '2026-09-20T10:00:00Z', 'Agreed, we can accept net 30.', '1A2B3C02') +
    cmt(11, 'Alex Customer', 'AC', '2026-09-20T11:00:00Z', 'Should liability be capped?', '1A2B3C03') +
    cmt(12, 'Jane Smith', 'JS', '2026-09-21T09:00:00Z', 'Yes, cap at 12 months of fees.', '1A2B3C04') +
    '</w:comments>')

# reply 10 shares comment 0's range; 11 and its reply 12 cover the "Liability" heading
doc = doc.replace('<w:commentRangeStart w:id="0"/>', '<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="10"/>')
doc = re.sub(r'(<w:commentRangeEnd w:id="0"/>\s*<w:r>.*?<w:commentReference w:id="0"/>\s*</w:r>)',
             r'\1<w:commentRangeEnd w:id="10"/><w:r><w:commentReference w:id="10"/></w:r>', doc, count=1, flags=re.S)
doc = re.sub(r'(<w:r>(?:(?!<w:r>).)*?<w:t>Liability</w:t></w:r>)',
             r'<w:commentRangeStart w:id="11"/><w:commentRangeStart w:id="12"/>\1<w:commentRangeEnd w:id="11"/><w:r><w:commentReference w:id="11"/></w:r><w:commentRangeEnd w:id="12"/><w:r><w:commentReference w:id="12"/></w:r>',
             doc, count=1, flags=re.S)
assert doc.count('commentRangeStart') == 4, doc.count('commentRangeStart')

ext = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w15:commentsEx xmlns:w15="{W15}">'
       '<w15:commentEx w15:paraId="1A2B3C01" w15:done="1"/>'
       '<w15:commentEx w15:paraId="1A2B3C02" w15:paraIdParent="1A2B3C01" w15:done="0"/>'
       '<w15:commentEx w15:paraId="1A2B3C03" w15:done="0"/>'
       '<w15:commentEx w15:paraId="1A2B3C04" w15:paraIdParent="1A2B3C03" w15:done="0"/>'
       '</w15:commentsEx>')
rels = files['word/_rels/document.xml.rels'].decode().replace('</Relationships>',
    '<Relationship Id="rIdExt1" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>')
ct = files['[Content_Types].xml'].decode().replace('</Types>',
    '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/></Types>')
files.update({'word/document.xml': doc.encode(), 'word/comments.xml': com.encode(), 'word/commentsExtended.xml': ext.encode(),
              'word/_rels/document.xml.rels': rels.encode(), '[Content_Types].xml': ct.encode()})
with zipfile.ZipFile('threaded.docx', 'w', zipfile.ZIP_DEFLATED) as z:
    for n, b in files.items(): z.writestr(n, b)
print('ok')
