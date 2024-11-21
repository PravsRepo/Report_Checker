import pdfplumber

with pdfplumber.open("D:/Report_checker/input/Team 56.pdf") as pdf:
    first_page = pdf.pages[0]
    print(first_page.extract_text())