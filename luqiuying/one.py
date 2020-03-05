import os
from win32com import client as wc

doc_path = "demo.doc"
word = wc.Dispatch("Word.Application")
doc = word.Documents.Open(doc_path)
#txt=4, html=10, docx=16， pdf=17
doc.SaveAs("demo.docx",16)
doc.Close()
word.Quit()

