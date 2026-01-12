# [`Table.AddColumnFromList`](/src/m/Table.AddColumnFromList.pq)

Adds a new column to a table using values from a provided list. The new column can be inserted at a specified position and can have a defined data type.

## Syntax

```fs
Table.AddColumnFromList(
    tbl as table,
    columnName as text,
    columnValues as list,
    optional position as number, 
    optional columnType as type
) as table
```

## Parameters

- `tbl`: The input table to which the new column will be added.
- `columnName`: The name of the new column to be added.
- `columnValues`: A list of values to populate the new column.
- `position` (_optional_): The position (0-based index) where the new column should be inserted. If not specified, the column is added at the end.
- `columnType` (_optional_): The data type of the new column. If not specified, the column will have type `any`.

## Return Value

Returns a new table with the added column populated with values from the provided list. If the list has fewer items than the number of rows in the table, nulls are added for the remaining rows. If the list has more items than the number of rows, extra items are ignored.

## Remarks

- If the `position` parameter is provided, the new column will be inserted at the specified index. If the index is out of bounds, an error will occur.
- If the `columnType` parameter is provided, the new column will be created with the specified data type. If not provided, the column will have type `any`.

## Examples

**Example 1**: Add a list as a new column at the end of the table.",

```fs
let
    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4]}),
    NewColumnValues = {10, 20},
    Result = Table.AddColumnFromList(Source, "C", NewColumnValues)
in
    Result
```

**Result**

|A  |B  |C     |
|:-:|:-:|:----:|
|1  |2  |10    |
|3  |4  |20    |

**Example 2**: Add a list as a new column at a specific position with a defined data type.

```fs
let
    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4]}),
    NewColumnValues = {10, 20},
    Result = Table.AddColumnFromList(Source, "C", NewColumnValues, 2, Int64.Type)
in
    Result
```

**Result**

|A  |C     |B  |
|:-:|:----:|:-:|
|1  |10    |2  |
|3  |20    |4  |

**Example 3**: If list has fewer items than rows, nulls are added for remaining rows.

```fs
let
    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4], [A=5, B=6]}),
    NewColumnValues = {10, 20},
    Result = Table.AddColumnFromList(Source, "C", NewColumnValues)
in
    Result
```

**Result**

|A  |B  |C     |
|:-:|:-:|:----:|
|1  |2  |10    |
|3  |4  |20    |
|5  |6  |_null_|

**Example 4**: If list has more items than rows, extra items are ignored.

```fs
let
    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4]}),
    NewColumnValues = {10, 20, 30, 40},
    Result = Table.AddColumnFromList(Source, "C", NewColumnValues)
in
    Result
```

**Result**

|A  |B  |C  |
|:-:|:-:|:-:|
|1  |2  |10 |
|3  |4  |20 |
