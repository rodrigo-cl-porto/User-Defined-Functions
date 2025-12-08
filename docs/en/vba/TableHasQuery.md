# [`TableHasQuery`](/src/vba/TableHasQuery.vba)

Checks whether a ListObject (Excel table) has an associated QueryTable.

## Syntax

```vb
TableHasQuery( _
    tbl As ListObject _
) As Boolean
```

## Parameters

- `tbl`: The ListObject (table) to check.

## Return Value

Returns `True` if the table has an associated `QueryTable`; otherwise returns `False`. If `tbl` is `Nothing`, the function returns `False`.

## Example

```vb
Dim tbl As ListObject
Dim hasQuery As Boolean

Set tbl = ThisWorkbook.Worksheets("Sheet1").ListObjects("Table1")
hasQuery = TableHasQuery(tbl)

If hasQuery Then
    Debug.Print "Table has a QueryTable"
Else
    Debug.Print "Table does not have a QueryTable"
End If
```
