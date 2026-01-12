# [`Table.TransposeCorrectly`](/src/m/Table.TransposeCorrectly.pq)

Transposes a table by converting selected columns (or all columns if none are specified) into rows, promotes headers, and adds a new column containing the original column names. This is useful for restructuring data while preserving column identity.

## Syntax

```fs
Table.TransposeCorrectly(
    tbl as table,
    optional columns as list,
    optional firstColumnName as text
) as table
```

## Parameters

- `tbl`: The input table whose columns will be transposed.
- `columnNames`: (_optional_) A list of column names to transpose. If not provided, all columns in the table will be transposed.
- `firstColumnName`: (_optional_) The name to assign to the first column of the transposed table. If not provided, the first name from the columns list will be used.

## Remarks

- The function promotes the first row of the transposed table as headers.
- A new column is added containing the original column names, inserted at the beginning of the table.
- This function is useful for reshaping data, especially when preparing it for pivoting or normalization.

## Examples

**Example 1**: Transposing all columns

```fs
let
    Source = #table(
        {"A", "B", "C"}, {
        {1, 2, 3},
        {4, 5, 6}
    }),
    Result = Table.TransposeCorrectly(Source)
in
    Result
```

**Result**

|A  |1  |4  |
|:-:|:-:|:-:|
|B  |2  |5  |
|C  |3  |6  |

**Example 2**: Transposing only the selected columns

```fs
let
    Source = #table(
        {"A", "B", "C"}, {
        {1, 2, 3},
        {4, 5, 6}
    }),
    Result = Table.TransposeCorrectly(Source, {"A", "B"})
in
    Result
```

**Result**

|A  |1  |4  |
|:-:|:-:|:-:|
|B  |2  |5  |

**Example 3**: Transposing the table and changing the first column name

```fs
let
    Source = #table(
        {"A", "B", "C"}, {
        {1, 2, 3},
        {4, 5, 6}
    }),
    Result = Table.TransposeCorrectly(Source, null, "D")
in
    Result
```

**Result**

|D  |1  |4  |
|:-:|:-:|:-:|
|B  |2  |5  |
|C  |3  |6  |
