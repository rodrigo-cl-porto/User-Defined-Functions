# [`Table.PreprocessTextColumns`](/src/power_query/Table.PreprocessTextColumns.pq)

This function cleans and formats text columns in a table. It removes line breaks, non-standard spaces, duplicated spaces, and applies optional casing (Proper, Lower, or Upper). You can specify which columns to process or let the function automatically detect all text columns.

## Syntax

```fs
Table.PreprocessTextColumns(
    tbl as table,
    optional columnNames as list,
    optional textCasing as text
) as table
```

## Parameters

- `tbl`: The input table containing text columns to be cleaned and formatted.
- `columnNames`: (_optional_) A list of column names to be processed. If not provided or empty, all columns of type text or nullable text will be processed.
- `textCasing`: (_optional_) A string indicating the desired text casing format. Accepted values are:
    - "Proper": Capitalizes the first letter of each word.
    - "Lower": Converts all texts to lowercase.
    - "Upper": Converts all texts to uppercase.
    - If not specified, casing is not changed.

## Remarks

- The function replaces line feed characters (`#(lf)`) with spaces.
- It removes non-breaking spaces (`Character.FromNumber(160)`), trims leading/trailing spaces, and collapses multiple spaces into one.
- This function is useful for preparing text data for analysis, comparison, or display.

## Examples

**Example 1**: Clean all text columns

```fs
let
    Source = #table(
        {"Name", "Comment"}, {
        {"  JOHN DOE  ", "Hello#(lf)World"},
        {"  jane smith", "Nice to meet you"}
    }),
    Result = Table.PreprocessTextColumns(Source)
in
    Result
```

**Result**

|Name          |Comment         |
|:-------------|:---------------|
|JOHN DOE      |Hello World     |
|jane smith    |Nice to meet you|

**Example 2**: Clean and apply Proper case to selected columns

```fs
let
    Source = #table(
        {"Name", "Note"}, {
        {"  MARIA   clara", "great#(lf)job"},
        {"joão   SILVA", "excellent work"}
    }),
    Result = Table.PreprocessTextColumns(Source, {"Name", "Note"}, "Proper")
in
    Result
```

**Result**

|Name       |Note          |
|:----------|:-------------|
|Maria Clara|Great Job     |
|João Silva |Excellent Work|
