# [`AutoFillFormulas`](/src/vba/AutoFillFormulas.vba)

Automatically fills formulas across a range using a reference cell's formula. The reference cell can be either the first or last cell containing a formula in the range.

## Syntax

```vb
AutoFillFormulas( _
    rng As Range, _
    Optional UseLastCellAsRef As Boolean = False _
)
```

## Parameters

- `rng`: The range where formulas will be filled
- `UseLastCellAsRef`: (_optional_) Boolean flag to determine which cell to use as reference
    - `False` (Default): Uses the first cell with formula as reference
    - `True`: Uses the last cell with formula as reference

## Remarks

- Does nothing if the range is empty (Nothing) or contains only one cell
- Only works if the range contains at least one formula
- Uses R1C1 formula notation to ensure proper relative references when filling
- Only fills formulas in cells that are part of the specified range
- Requires the helper function [`RangeHasAnyFormula`](#rangehasanyformula) to check for formulas in the range

## Example

```vb
Dim rng As Range
Set rng = Range("A1:A10")
AutoFillFormulas rng 'Uses first formula cell as reference

'Or using the last cell as reference:
AutoFillFormulas rng, True
```
