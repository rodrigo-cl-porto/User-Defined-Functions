# [`GetTableColumnNames`](/src/vba/GetTableColumnNames.vba)

Returns the header names of an Excel ListObject (table) as a zero-based string array.

## Syntax

```vb
GetTableColumnNames( _
    lo As ListObject _
) As String()
```

## Parameters

- `lo`: The ListObject (Excel table) to read column headers from

## Return Value

Returns a zero-based array of strings containing the table column header values in left-to-right order.

## Remarks

- Includes hidden columns and preserves the table column order.

## Example

```vb
Dim colNames() As String
Dim i          As Long

Set tbl = ThisWorkbook.Worksheets("Sheet1").ListObjects("Table1")
colNames = GetTableColumnNames(tbl)

For i = 0 To UBound(colNames)
    Debug.Print colNames(i)
Next i
```
