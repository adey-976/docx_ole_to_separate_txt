# extract_docx_into_separate_txt.py

Extracts text from all embedded files in a `.docx` document and saves them into a single `.txt` file with contextual references — designed to be used alongside the original `.docx` as input to an LLM chatbot.

The output contains **only the embedded files' content** (not the full document text), with enough surrounding context for an LLM to understand which embedded file corresponds to which location in the original document.

## Installation

pip install -r requirements_extract_docx_into_separate_txt.txt

## Usage

python3 extract_docx_into_separate_txt.py <input.docx> [output.txt]

- If `output.txt` is not specified, the output is saved as `<input_name>_embedded_extractions.txt`

## Supported Embedded File Types

- `.docx` (Word documents)
- `.xlsx` / `.xlsm` (Excel spreadsheets)
- `.pdf` (PDF documents)
- `.eml` (email files)
- `.msg` (Outlook messages)
- `.csv` / `.txt` (plain text files)

## Output Format

The output text file has:

1. A header identifying the source document
2. A count of total embedded files found
3. For each embedded file:
   - **Filename** — name of the embedded file
   - **Type** — human-readable type (e.g., "Excel Spreadsheet (.xlsx)")
   - **Location** — where it sits in the document (e.g., "Table 4, Row 2, Cell 2" or "Body paragraph 6")
   - **Text before** — the text immediately preceding the embedded file in the document (~150 chars)
   - **Text after** — the text immediately following the embedded file in the document (~150 chars)
   - **Extracted content** — the full text extraction
