# excel2csv

Yet another Excel (.xlsx) to CSV (.csv) file convertor.

By Gemini 2.5 Pro.

Initial Prommpt:

>Write a Go program to convert an Excel file (.xlsx) to .csv. By default, only the first sheet is converted. The first line is treated as header. Do the row reading & writing in a stream way. Handle the escaping and quoting.

## Usage

```
excel2csv <file.xlsx>
```

Use `-` as filename to read .xlsx file from stdin.

## Command-line flags

```
Usage of excel2csv:
  -o string
        Path to the output CSV file. Use '-' for stdout. Defaults to <input_filename>.csv
  -sheet-index int
        0-based index of the Excel sheet to convert
```
