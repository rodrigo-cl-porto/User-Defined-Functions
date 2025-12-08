# [`RangeHasConstantValues`](/src/vba/RangeHasConstantValues.vba)

Checks whether a given range contains any constant (non-formula) cells.

## Syntax

```vb
RangeHasConstantValues( _
    rng As Range _
) As Boolean
```

## Parameters

- `rng`: Range to check for constant values.

## Return Value

Returns `True` if the range contains at least one constant cell; otherwise returns False. If `rng` is `Nothing` the function returns `False`.

## Example

```vb
Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Sheet1").Range("A1:A10")

If RangeHasConstantValues(rng) Then
    Debug.Print "Range contains constants"
Else
    Debug.Print "Range contains no constants or is invalid"
End If
```
