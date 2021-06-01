# Edit existing .docx files effortlessly and without changing the original format.


# Imports
import pandas as pd
from docx import Document


# Insert .docx file
file_name = ''

doc = Document(file_name)


def replace_string(old_text, new_text):
    """Replace and old string with a new string"""
    
    global doc

    for p in doc.paragraphs:
        if old_text in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if old_text in inline[i].text:
                    text = inline[i].text.replace(old_text, new_text)
                    inline[i].text = text

    return


def show_string(old_text):
    """Show the 'line' of text in the document where the string is found (without editing)"""

    global doc

    for p in doc.paragraphs:
        if old_text in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if old_text in inline[i].text:
                    text = inline[i].text
                    inline[i].text = text
            print(p.text)

    return

if __name__ == '__main__':
  show_string('title')
  
  
  
