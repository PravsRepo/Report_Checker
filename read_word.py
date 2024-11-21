from docx import Document

# Load a .docx file
doc = Document("")

# Read the content of the document
for paragraph in doc.paragraphs:
    print(paragraph.text)
