import docx

doc = docx.Document('/Users/dima/Desktop/My project/Rfid-панель/docs/google_doc_downloaded.docx')
for i, p in enumerate(doc.paragraphs):
    print(f"P{i:02d}: {p.text}")
