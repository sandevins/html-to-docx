# html-to-docx

This script transforms a HTML document into a properly formatted Word document (.docx). It currently supports the following.

- Headers
- Paragraphs
- Embedded base64 images
- Tables

## Usage

To use the script, the dependencies must be installed.

```bash
pip install -r requirements.txt
```

Then you can use the script specifying the input file and the output file in that order.

```bash
python html_converter.py <PATH_TO_HTML_FILE> <PATH_TO_OUTPUT_FILE>
```