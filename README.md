# Mass Edit Word Documents But Keep Original Formatting

Edit existing Word documents effortlessly and without changing the original formatting of the document. 

As great as the original `docx` library is, this "keep same format" feature is not natively supported. However, this functionality is pretty useful if you want to automate document writing or editing and need to adhere to strict formatting rules. Hence, I wrote this small add-on module.

[![PyPI version](https://badge.fury.io/py/docxedit.svg)](https://badge.fury.io/py/docxedit)

## Installing

With pip: `pip install docxedit`

## Dependencies

Included as a dependency: `python-docx`

## Functions

Most of the functions in this module work primarily with **runs**, which are sequences of strings with the same formatting style.
This is how we can edit the document but keep the original format.

# Example Usage

```python
from docx import Document
import docxedit

document = Document()

# Replace all instances of the word 'Hello' in the document by 'Goodbye'
docxedit.replace_string(document, old_string='Hello', new_string='Goodbye')

# Replace all instances of the word 'Hello' in the document by 'Goodbye' but only
# up to paragraph 10
docxedit.replace_string_up_to_paragraph(document, old_string='Hello', new_string='Goodbye', 
                                        paragraph_number=10)

# Remove any line that contains the word 'Hello' along with the next 5 lines
docxedit.remove_lines(document, first_line='Hello', number_of_lines=5)
```