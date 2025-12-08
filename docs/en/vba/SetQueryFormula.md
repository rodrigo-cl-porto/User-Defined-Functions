# [`SetQueryFormula`](/src/vba/SetQueryFormula.vba)

Modifies a Power Query formula in the current workbook based on a given value, handling different data types appropriately.

## Syntax

```vb
SetQueryFormula( _
    queryName As String, _
    value As Variant _
)
```

## Parameters

- `queryName`: Name of the Power Query to modify
- `value`: Value to set in the query formula (supports `String`, `Date`, and `Byte Array`)

## Dependencies

- Requires Excel version that supports Power Query

## Example

```vb
' Set a string value
SetQueryFormula "MyQuery", "Hello ""World"""  ' Results in: "Hello ""World"""

' Set a date value
SetQueryFormula "MyQuery", DateSerial(2023, 10, 17)  ' Results in: #date(2023,10,17)

' Set a byte array
Dim byteArr() As Byte
byteArr = Array(1, 2, 3)
SetQueryFormula "MyQuery", byteArr  ' Results in: {1,2,3}
```
