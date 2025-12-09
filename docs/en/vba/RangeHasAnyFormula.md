# [`RangeHasAnyFormula`](/src/vba/RangeHasAnyFormula.vba)

Checks if a given range contains any cell with formulas.

## Syntax

```vb
RangeHasAnyFormula( _
    ByVal rng As Range _
) As Boolean
```

## Parameters

- `rng`: The range to be checked for formulas

## Return Value

Returns `True` if the range contains at least one formula, `False` otherwise.

## Remarks

- Returns `False` if the range is Nothing
- Uses error handling to detect the presence of formulas
- Shows an error message if any unexpected error occurs during execution
- Uses Excel's `SpecialCells` method with `xlCellTypeFormulas` to perform the check

## Example

```vb
Dim rng As Range
Set rng = Range("A1:D10")

If RangeHasAnyFormula(rng) Then
    Debug.Print "Range contains at least one formula"
Else
    Debug.Print "Range contains no formulas"
End If
```

## **Error Handling**

- Displays a message box with error details if an unexpected error occurs
- Properly handles the "No cells were found" error which indicates no formulas are present
