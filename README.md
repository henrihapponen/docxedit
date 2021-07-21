# Useful functions to edit Word Documents with Python (python-docx)

Edit existing Word documents effortlessly and without changing the original formatting of the document. These tools are very useful if you want to automate document writing or editing and you need to adhere to strict formatting rules.

## Requirements

Need to have python-docx installed and imported

## Functions

All of these functions work mainly with 'runs', which are sequences of strings with the same formatting style.

Functions
- `show_line(current_text)`: Prints out the 'line' of text in the document where the string is found (without replacing anything). A 'line' is typically a paragraph but it can be shorter if the formatting changes before the end of the paragraph (i.e. if the 'run' ends).
- `replace_string(old_text, new_text)`: Replaces an old string (placeholder) with a new string without changing the formatting of the text in the document. Very useful when automating the writing/editing of documents.
- `replace_string_up_to_paragraph(old_text, new_text, paragraph_number)`: Replaces an old string (placeholder) with a new string (variable) without changing the format of the text
    but only up to a specific paragraph number.
- `delete_paragraph(paragraph)`: Delete a paragraph. Input must be a paragraph object. This function is used by the following function.
- `remove_lines(first_line, number_of_lines)`: Remove a line including any keyword, and a certain number of rows after that. This allows for removal of entire sections/paragraphs or simply a few lines of text, depending on your inputs.
- `add_text_in_table`: Add text to a cell in a table.
- `change_table_font_size`: Change the font size af a full table.

