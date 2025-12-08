# [`RangeIsHidden`](/src/vba/RangeIsHidden.vba)

Determines whether a given range is entirely hidden (no visible cells).

## Syntax

```vb
RangeIsHidden( _
    rng As Range _
) As Boolean
```

## Parameters

- `rng`: The Range to check for visibility.

## Return Value

Returns `True` if the range contains no visible cells (i.e., is hidden). Returns `False` if at least one cell in the range is visible or if `rng` is `Nothing`.

## Example

```vb
Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Sheet1").Range("A1:A10")

If RangeIsHidden(rng) Then
    Debug.Print "Range is hidden (no visible cells)."
Else
    Debug.Print "Range has visible cells."
End If
```
