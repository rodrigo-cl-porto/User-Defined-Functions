# [`Table.RemoveBlankColumns`](/src/m/Table.RemoveBlankColumns.pq)

Removes columns from a table that contain only blank values.

## Syntax

```fs
Table.RemoveBlankColumns(
    tbl as table
) as table
```

## Parameters

- `tbl`: The table from which blank columns will be removed.

## Example

Transposing the table and changing the first column name

```fs
let
    Source = #table(
        {"A", "B"}, {
        {null, "value1"},
        {"", "value2"}
    }),
    Result = Table.RemoveBlankColumns(Source)
in
    Result
```

**Result**

|B     |
|:----:|
|value1|
|value2|

## Credits

- Author: [Excel Off The Grid](https://exceloffthegrid.com/)
- Source: [Power Query Trick: Instantly Remove All Null Columns! 💥](https://www.youtube.com/watch?v=Zkg9ICg9i40)
