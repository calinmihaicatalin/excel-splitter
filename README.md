# FinOps Excel Splitter

A desktop application that splits a formatted Excel template into individual workbooks for each owner, organized in folders by department.

## Features

- Select an `.xlsx` or `.xlsm` workbook and specify the header row (default: 5).
- Choose the department column for output folders and the owner column for individual files.
- Include or exclude departments before generating files.
- Preserve the header/template rows, cell formatting, merged headers, column widths, row heights, freeze panes, and basic print settings.
- Create output folders and safe file names automatically.
- Include only sheets that contain matching rows for each department and owner.

## Requirements

- Python 3.10 or later
- `openpyxl`
- `pandas`

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

## Run

```bash
python finops_splitter.py
```

In the app:

1. Browse to the source Excel workbook.
2. Set the row containing column headers, then select **Load Workbook**.
3. Choose the department and owner columns.
4. Check the departments to include.
5. Choose an output folder and select **Run Split**.

Generated files are arranged like this:

```text
output/
├── Department Name/
│   ├── Owner Name.xlsx
│   └── Another Owner.xlsx
└── Another Department/
    └── Owner Name.xlsx
```

## Notes

- The selected header row is copied along with rows above it. Matching data begins on the next row.
- Data rows are written using the template's style from the first row below the header.
- The application creates `.xlsx` output files. Macro-enabled input files (`.xlsm`) can be read, but VBA macros are not copied into generated files.
- Avoid using the same owner name for distinct people within one department, because they would be combined into one output workbook.
