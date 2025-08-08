# Examples Directory

This directory contains example DOCX files for testing the parser.

## Usage

Place your test DOCX files in this directory to test the parser functionality.

## File Types Supported

- Standard DOCX documents
- Documents with tables
- Documents with images
- Documents with embedded objects
- Documents with SmartArt graphics

## Testing

Use the files in this directory with:

```python
from src.parsers.document_parser import DocumentParser

parser = DocumentParser()
result = parser.parse_document("Files/examples/your_document.docx")
```
