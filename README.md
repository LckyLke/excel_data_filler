# README

### Usage

```bash
python main.py <file_path> <sheet_name> <column_index> -f <full_output_file> -c <column_output_file>
```

- `<file_path>`: Path to the Excel file.
- `<sheet_name>`: Name of the sheet to read.
- `<column_index>`: Index of the column to process (0-based).
- `-f <full_output_file>`: (Optional) Output file for the entire modified sheet. Default is `redistributed_full_output.xlsx`.
- `-c <column_output_file>`: (Optional) Output file for the modified column. Default is `redistributed_column.csv`.

**Example:**

```bash
python main.py "data.xlsx" "8760 h Korrektur" 6 -f modified_data.xlsx -c modified_column.csv
```
