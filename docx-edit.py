# Useful functions to edit Word Documents with Python-Docx module

# Import
from docx import Document


# Function definitions

def show_line(old_text):
    """Shows the 'line' of text in the doc where the string is found (without replacing anything)"""

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


def replace_string(old_text, new_text):
    """Replaces an old string (placeholder) with a new string without changing the formatting of the text"""

    global doc

    for p in doc.paragraphs:
        if old_text in p.text:
            inline = p.runs

            for i in range(len(inline)):
                if old_text in inline[i].text:
                    text = inline[i].text.replace(str(old_text), str(new_text))
                    inline[i].text = text

    return


def delete_paragraph(paragraph):
    """Delete a paragraph. Input must be a paragraph object"""

    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None

    return


def remove_lines(first_line, number_of_rows):
    """Remove a line including any keyword (first_line), and a certain number of rows after that"""

    list_of_paragraphs = []

    for i in doc.paragraphs:
        list_of_paragraphs.append(i)

    for index, i in enumerate(doc.paragraphs):
        if first_line in i.text:
            try:
                delete_paragraph(i)
            except AttributeError:
                print('Could not remove line ' + str(index) + ': ' + i.text)

            b = 0
            c = 0
            while b < number_of_rows:
                try:
                    print(list_of_paragraphs[a + 1 + b])
                    delete_paragraph(list_of_paragraphs[index + 1 + c])
                    b += 1
                except AttributeError:
                    print('Could not remove line ' + str(index + 1 + c))
                    c += 1
                    continue

    return


if __name__ == '__main__':

    # For example:
    
    document_path = 'your full document path here'
    doc = Document(document_path)

    replace_string('placeholder', 'new text')
    remove_lines('remove this line', 1)
    remove_lines('remove the next 5 lines', 5)