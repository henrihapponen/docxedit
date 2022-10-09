import docx


def show_line(doc, current_text: str):
    """
    Shows the 'line' of text in the doc where the string is found (without replacing anything).
    Args:
        doc: The docx document.
        current_text: The string to search for.
    Returns:
        The line of text where the string is found.
    """

    try:
        for paragraph in doc.paragraphs:
            if current_text in paragraph.text:
                inline = paragraph.runs

                for i in range(len(inline)):
                    if current_text in inline[i].text:
                        text = inline[i].text
                        inline[i].text = text
                print(f'Text found in line: {paragraph.text}')
            else:
                print(f'Error: {current_text} not found in document.')
    except Exception as e:
        print(f'Error: Could not find {current_text}: {e}')


def replace_string(doc, old_string: str, new_string: str):
    """
    Replaces an old string (placeholder) with a new string
    without changing the formatting of the text.
    Args:
        doc: The docx document.
        old_string: The old string to replace.
        new_string: The new string to replace the old one with.
    Returns:
        Success or Error.
    """

    for paragraph in doc.paragraphs:
        if old_string in paragraph.text:
            inline = paragraph.runs

            for i in range(len(inline)):
                if old_string in inline[i].text:
                    text = inline[i].text.replace(str(old_string), str(new_string))
                    inline[i].text = text
            print(f'Success: Replaced {old_string} with {new_string}')
        else:
            print(f'Error: Could not find {old_string}')


def replace_string_up_to_paragraph(doc, old_string, new_string,
                                   paragraph_number):
    """
    Replaces an old string (placeholder) with a new string
    without changing the format of the text but only up to
    a specific paragraph number.
    Args:
        doc: The docx document.
        old_string: The old string to replace.
        new_string: The new string to replace the old one with.
        paragraph_number: The paragraph number to stop at.
    """

    for index, paragraph in enumerate(doc.paragraphs):

        # Replace every instance before paragraph number 'paragraph_number'
        if index < paragraph_number:

            if old_string in paragraph.text:
                inline = paragraph.runs

                for i in range(len(inline)):
                    if old_string in inline[i].text:
                        text = inline[i].text.replace(str(old_string), str(new_string))
                        inline[i].text = text
                print(f'Success: Replaced {old_string} with {new_string} up to paragraph {index}')


def remove_paragraph(paragraph):
    """
    Remove a paragraph. Input must be a paragraph object.
    Args:
        paragraph: The paragraph object to remove.
    Returns:
        Success or Error.
    """

    try:
        paragraph_element = paragraph._element
        paragraph_element.getparent().remove(paragraph_element)
        paragraph._p = paragraph._element = None
        print(f'Success: Removed paragraph {paragraph}')
    except Exception as e:
        print(f'Error: Could not remove paragraph {paragraph}: {e}')


def remove_lines(doc, first_line: str, number_of_lines: int):
    """
    Remove a line including any keyword (first_line),
    and a certain number of rows after that.
    Args:
        doc: The docx document.
        first_line: The first line to remove.
        number_of_lines: The number of lines to remove.
    Returns:
        Success or Error.
    """

    list_of_paragraphs = []

    for i in doc.paragraphs:
        list_of_paragraphs.append(i)

    for index, i in enumerate(doc.paragraphs):
        if first_line in i.text:
            try:
                remove_paragraph(i)
            except AttributeError:
                print(f'Error: Could not remove line {index}: {i.text}')

            b_var = 0
            c_var = 0
            while b_var < number_of_lines:
                try:
                    print(list_of_paragraphs[index + 1 + b_var])
                    remove_paragraph(list_of_paragraphs[index + 1 + c_var])
                    b_var += 1
                    print(f'Success: Removed line {str(index + 1 + c_var)}')
                except Exception as e:
                    print(f'Error: Could not remove line {str(index + 1 + c_var)}: {e}')
                    c_var += 1
                    continue


def add_text_in_table(table, row_num: int, column_num: int, new_string: str):
    """
    Add text to a cell in a table.
    Args:
        table: The table object.
        row_num: The row number to add the text to.
        column_num: The column number to add the text to.
        new_string: The text to add.
    Returns:
        Success or Error.
    """

    try:
        table.cell(row_num, column_num).text = new_string
        print(f'Success: Added {new_string} to row {row_num} and column {column_num}')
    except Exception as e:
        print(f'Error: Could not add {new_string} to row {row_num} and column {column_num}: {e}')


def change_table_font_size(table, font_size: int):
    """
    Change the font size af the whole table.
    Args:
        table: The table object.
        font_size: The font size to change to.
    Returns:
        Success or Error.
    """

    try:
        for row in table.rows:
            for cell in row.cells:
                paragraphs = cell.paragraphs
                for paragraph in paragraphs:
                    for run in paragraph.runs:
                        font = run.font
                        font.size = docx.shared.Pt(font_size)
        print(f'Success: Changed font size to {font_size}')
    except Exception as e:
        print(f'Error: Could not change font size to {font_size}: {e}')
