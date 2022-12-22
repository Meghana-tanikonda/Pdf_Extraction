import re
from pdfminer.high_level import extract_pages, extract_text
def Readpdf(n):
    text = extract_text(n)
    return(text)

if __name__ == "__main__":
    result=  Readpdf()
