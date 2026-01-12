# [`WorksheetHasListObject`](/src/vba/WorksheetHasListObject.vba)

Checks if a worksheet contains at least one ListObject (Excel table).

## Syntax

```vb
WorksheetHasListObject( _
    ws As Worksheet _
) As Boolean
```

## Parameters

- `ws`: Worksheet to check for ListObjects.

## Return Value

Returns `True` if the worksheet contains one or more `ListObjects`; otherwise returns `False`.

## Example

```vb
Dim hasTable As Boolean
hasTable = WorksheetHasListObject(ThisWorkbook.Worksheets("Sheet1"))

If hasTable Then
    Debug.Print "Sheet1 contains at least one table."
Else
    Debug.Print "Sheet1 contains no tables."
End If
```
