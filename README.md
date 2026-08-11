# Excel Cleaner

A small Python script that cleans a student address spreadsheet by splitting a
combined `Student Name` column into separate `First Name` and `Last Name`
columns, then writes the result to a new Excel file.

## What it does

Given an input workbook (`students_addresses_2.xlsx`) with a `Students` sheet
containing a single `Student Name` column:

| Student Name    | Street Address    | City | Province | Postal Code |
|------------------|-------------------|------|----------|--------------|
| Lerato Mahlangu  | 19 Rissik Street  | ...  | Gauteng  | 2001         |

The script splits each name on the first space and outputs:

| First Name | Last Name | Street Address   | City | Province | Postal Code |
|------------|-----------|-------------------|------|----------|--------------|
| Lerato     | Mahlangu  | 19 Rissik Street  | ...  | Gauteng  | 2001         |

The cleaned data is written to `students_addresses_cleaned.xlsx`.

## Requirements

- Python 3.8+
- [pandas](https://pypi.org/project/pandas/)
- [openpyxl](https://pypi.org/project/openpyxl/) (used by pandas to read/write `.xlsx` files)

Install dependencies:

```bash
pip install pandas openpyxl
```

## Usage

1. Place your input file in the same directory as the script and name it
   `students_addresses_2.xlsx`, with the data on a sheet called `Students`.
2. Run the script:

```bash
python Excel_Cleaner.py
```

3. The cleaned file will be saved as `students_addresses_cleaned.xlsx` in the
   same directory.

## Notes / limitations

- Names are split on the **first space only**, so a name like
  `Anna Marie van der Berg` becomes First Name: `Anna`, Last Name:
  `Marie van der Berg`.
- Names with no space will raise an error (`not enough values to unpack`),
  since the script expects at least a first and last name.
- The input file name, sheet name, and output file name are currently
  hardcoded in the script rather than passed as arguments.

## Possible improvements

- Accept input/output file paths as command-line arguments.
- Handle names with no space, extra spaces, or titles (e.g. "Dr.", "Jr.") more
  gracefully.
- Remove the loop that re-saves the file on every iteration (it currently
  writes to Excel once per row, which is unnecessary and slow for large
  files) — move `Students.to_excel(...)` outside the `for` loop.
  
  ## License
This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.


