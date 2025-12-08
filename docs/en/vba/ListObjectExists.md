# [`ListObjectExists`](/src/vba/ListObjectExists.vba)

Checks whether a ListObject (Excel table) with a given name exists in a workbook.

## Syntax

```vb
ListObjectExists( _
    ByRef wb As Workbook, _
    loName As String _
) As Boolean
```

## Parameters

- `wb`: Workbook to search.
- `loName`: Name of the table (`ListObject`) to find.

## Return Value

Returns `True` if a ListObject with the specified name is found in any worksheet of the workbook; otherwise returns `False`.

## Remarks

- Performs a direct name comparison (behavior may be affected by the project's Option Compare setting).

## Example

```vb
Dim exists As Boolean
exists = ListObjectExists(ThisWorkbook, "Table1")

If exists Then
    Debug.Print "Table exists"
Else
    Debug.Print "Table not found"
End If
```