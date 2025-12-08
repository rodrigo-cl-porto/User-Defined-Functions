# [`AreArraysEquals`](/src/vba/AreArraysEqual.vba)

Compares two arrays to check if they are equal, meaning they have the same size and identical elements in the same order.

## Syntax

```vb
AreArraysEqual( _
    Array1 As Variant, _
    Array2 As Variant _
) As Boolean
```

## Parameters
- `Array1`: First array to compare
- `Array2`: Second array to compare

## Return Value

Returns `True` if both arrays are equal, `False` otherwise.

## Remarks

- Arrays must have the same upper and lower bounds
- Arrays must have identical elements in the same positions
- The function performs an element-by-element comparison
- Returns `False` if arrays have different sizes
- Can compare arrays of any type since parameters are declared as Variant

## Example

```vb
Dim arr1 As Variant
Dim arr2 As Variant
arr1 = Array(1, 2, 3)
arr2 = Array(1, 2, 3)

If AreArraysEqual(arr1, arr2) Then
    Debug.Print "Arrays are equal"
Else
    Debug.Print "Arrays are different"
End If
```
