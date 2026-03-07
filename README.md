[![PyPI Downloads](https://static.pepy.tech/personalized-badge/docxedit?period=total&units=INTERNATIONAL_SYSTEM&left_color=BLACK&right_color=GREEN&left_text=downloads)](https://pepy.tech/projects/docxedit)
[![PyPI version](https://badge.fury.io/py/docxedit.svg)](https://badge.fury.io/py/docxedit)

# docxedit

Edit Word documents effortlessly without changing the original formatting.

The original `docx` library is great but it's missing one important feature: *keeping the original formatting*.

This is useful if you want to automate document writing or editing and need to adhere to strict formatting rules — a common requirement in professional settings.

## Install

```
pip install docxedit
```

## Dependencies

Included as a dependency: `python-docx` (`docx`)

## Functionalities

Most of the functions in this module work primarily with **runs**, which are sequences of strings with the same formatting style. Breaking the document into runs allows us to edit the text without changing the original formatting.

- **Replace** all occurrences of a string with a new string (optionally limit up to a paragraph number, and include or exclude tables)
- **Remove** a paragraph or a line (and optionally a number of lines after it)
- **Add** text to a table cell
- **Change** the font size of a table
- **Show** the line where a specific string is found

## How to Use

```python
from docx import Document
import docxedit

document = Document('path/to/your/document.docx')

# Replace all instances of 'Hello' with 'Goodbye' (including tables)
docxedit.replace_string(document, old_string='Hello', new_string='Goodbye')

# Replace all instances of 'Hello' with 'Goodbye' but only up to paragraph 10
docxedit.replace_string_up_to_paragraph(document, old_string='Hello', new_string='Goodbye',
                                        paragraph_number=10)

# Show the line where 'Hello' is found
docxedit.show_line(document, current_text='Hello')

# Remove any line that contains 'Hello' along with the next 5 lines after that
docxedit.remove_lines(document, first_line='Hello', number_of_lines=5)

# Remove a specific paragraph
docxedit.remove_paragraph(document.paragraphs[0])

# Add text in a table cell (row 1, column 1) in the first table
docxedit.add_text_in_table(document.tables[0], row_num=1, column_num=1, new_string='Hello')

# Change font size of the first table to 12pt
docxedit.change_table_font_size(document.tables[0], font_size=12)

# Save the document
document.save('path/to/your/edited/document.docx')
```

## Logging

By default, `docxedit` does not produce any console output. To enable debug logging:

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```
