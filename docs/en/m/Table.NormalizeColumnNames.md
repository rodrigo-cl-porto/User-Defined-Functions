# [`Table.NormalizeColumnNames`](/src/m/Table.NormalizeColumnNames.pq)

Cleans and standardizes column names in a table by removing unwanted characters, trimming spaces, and applying specified text formatting (Proper, Lower, Upper). It also removes columns with default names like `Column1`, `Column2`, etc.

## Syntax

```fs
Table.NormalizeColumnNames(
    tbl as table,
    optional textFormat as text
) as table
```

## Parameters

- `tbl`: The input table whose column names need to be fixed.
- `textFormat` (_optional_): The desired text format for the column names. Accepts `Proper`, `Lower`, or `Upper`. If not specified, no formatting is applied.

## Return Value

A table with cleaned and standardized column names.

## Remarks

- The function processes the column names of the provided table to ensure they are clean and standardized. It removes non-printable characters, trims leading and trailing spaces, replaces non-breaking spaces with regular spaces, eliminates duplicated spaces, and applies the specified text formatting (Proper, Lower, Upper). Additionally, it removes any columns that have default names such as 'Column1', 'Column2', etc., ensuring that only relevant columns remain in the Result table.
- If the `textFormat` parameter is not provided, the function will only clean the column names without applying any specific text formatting.

## Examples

```fs
Table.NormalizeColumnNames(SourceTable, "Proper") // Cleans and formats column names to Proper case.
Table.NormalizeColumnNames(SourceTable, "Lower") // Cleans and formats column names to Lower case.
Table.NormalizeColumnNames(SourceTable, "Upper") // Cleans and formats column names to Upper case.
Table.NormalizeColumnNames(SourceTable) // Cleans column names without applying any specific text formatting.
```