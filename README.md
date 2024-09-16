# xlsx2md

**xlsx2md** is a Python program designed to convert Excel files (`.xlsx`) to Markdown format by parsing the underlying XML structure of the Excel file. The program extracts cell values, handles merged cells, and applies styles such as superscript, subscript, and indentation while converting the data into a Markdown table.

## Features

- Converts Excel spreadsheets to Markdown tables.
- Extracts shared strings, cell styles, and merged cell information.
- Supports rich text elements (superscript, subscript) and indentation.
- Optionally includes drawing metadata extracted from the Excel file.

## Requirements

- Python 3.x
- No additional dependencies are required beyond Python's standard libraries (`zipfile`, `xml.etree.ElementTree`, `os`, and `sys`).

## Usage

To convert an Excel file to a Markdown table, run the following command in your terminal:

```bash
python excel2markdown.py <excel_file.xlsx>
```

- Replace `<excel_file.xlsx>` with the path to the Excel file you want to convert.
- The resulting Markdown table will be printed to the console.

## Example

If you have a file `sample.xlsx`, run the following:

```bash
python xlsx2md.py sample.xlsx
```

The output will be a Markdown table representation of the data in `sample.xlsx`.

## Limitations

- This script only processes `.xlsx` files.
- It does not support heavily formatted or complex Excel sheets (e.g., pivot tables).
- Only basic text formatting and merged cells are handled.
