# Mass Edit Word Documents Without Changing the Original Formatting

Edit Word documents effortlessly and without changing the original formatting of the document. 

The original `docx` library is great but it's missing one important feature: *keeping the original formatting*.

However, this functionality is pretty useful if you want to automate document writing or editing and need to adhere to strict formatting rules. This is a common requirement in corporate environments for example.

This module adds that feature with some functions that allow editing documents without changing the original formatting.

[![PyPI version](https://badge.fury.io/py/docxedit.svg)](https://badge.fury.io/py/docxedit)

## Install

With pip: `pip install docxedit`

## Dependencies

Included as a dependency: `python-docx` (`docx`)

## Functionalities
Most of the functions in this module work primarily with **runs**, which are sequences of strings with the same formatting style. Breaking the document into runs allows us to edit the text without changing the original formatting.

Some of the functionalities that this module include:
- Replacing all occurrences of a string with a new string (optionally limit this up to a paragraph number, and include or exclude tables)
- Removing a line that includes a specific string
- Add text to a table

The beauty of this module is that you can use all of its functions to **mass edit** Word documents with consistency and precision. This is useful especially in corporate environments where a lot of document writing or editing can be automated. 


## How to Use

Usage of this module is really simple. Here are some examples:

```python
from docx import Document
import docxedit

document = Document('path/to/your/document.docx')

# Replace all instances of the word 'Hello' in the document with 'Goodbye' (including tables)
docxedit.replace_string(document, old_string='Hello', new_string='Goodbye')

# Replace all instances of the word 'Hello' in the document with 'Goodbye' but only
# up to paragraph 10
docxedit.replace_string_up_to_paragraph(document, old_string='Hello', new_string='Goodbye', 
                                        paragraph_number=10)

# Remove any line that contains the word 'Hello' along with the next 5 lines after that
docxedit.remove_lines(document, first_line='Hello', number_of_lines=5)

# Add text in a table cell (row 1, column 1) in the first table in the document
docxedit.add_text_in_table(document.tables[0], row_num=1, column_num=1, new_string='Hello')

# Save the document
document.save('path/to/your/edited/document.docx')
```
